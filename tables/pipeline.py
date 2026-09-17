"""一覧表の取り込みの実行関数（ジョブ本体）と、画面から呼ぶ補助。

- run_read: 保存した範囲と設定で全行を読む → 正規化 → チェック → rows.jsonl.gz / issues.json / issues.csv
- run_render: その取り込みの記録（＋照合に通った AI 整形の結果）から全 md を作り、取り込みのフォルダに保存
- build_download: md と管理用CSV をまとめた zip
期間の置き換え・投入済みとの差分・取り消しはしない（取り込みごとに、その取り込みの内容だけで md を作る）。
置き場所: TABLES_DIR/imports/<import_id>/（rows.jsonl.gz, issues.*, md/, preview_md/）
"""
from __future__ import annotations

import gzip
import io
import json
import os
import shutil
from pathlib import Path

from flask import current_app

from core import jobs
from core.files import UploadError, upload_path
from models import database
from tables import outputs, store
from tables.checks import count_levels, run_checks
from tables.detect import guess_layout
from tables.markdown import MdFile, render_all
from tables.normalize import read_records
from tables.source import open_source

ROWS_FILE = "rows.jsonl.gz"
ISSUES_CSV = "issues.csv"
ISSUES_JSON = "issues.json"
MD_DIR = "md"
PREVIEW_DIR = "preview_md"


class PipelineError(Exception):
    """利用者に見せる日本語メッセージの処理エラー。"""


# ---- ファイル -----------------------------------------------------------------------------

def import_dir(import_id: int) -> Path:
    return Path(current_app.config["TABLES_DIR"]) / "imports" / str(int(import_id))


def import_files(import_id: int) -> dict[str, Path]:
    base = import_dir(import_id)
    return {"dir": base, "rows": base / ROWS_FILE, "issues_csv": base / ISSUES_CSV, "issues_json": base / ISSUES_JSON,
            "md": base / MD_DIR, "preview": base / PREVIEW_DIR}


def write_atomic(path: Path, data: bytes) -> None:
    """一時ファイルに書いてから置き換える（途中で止まっても前のファイルが残る）。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.tmp")
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)


def write_rows(path: Path, records) -> None:
    """記録を gzip の JSON Lines で保存（mtime=0・キー順固定で決定的）。"""
    raw = io.BytesIO()
    with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0) as gz:
        for rec in records:
            d = rec.to_dict() if hasattr(rec, "to_dict") else dict(rec)
            gz.write(json.dumps(d, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8") + b"\n")
    write_atomic(path, raw.getvalue())


def load_rows(import_id: int, offset: int = 0, limit: int | None = None) -> list[dict]:
    path = import_files(import_id)["rows"]
    if not path.exists():
        return []
    with gzip.open(path, "rt", encoding="utf-8") as f:
        rows = [json.loads(line) for line in f if line.strip()]
    return rows[offset: offset + limit] if limit is not None else rows[offset:]


def load_issues(import_id: int) -> list[dict]:
    path = import_files(import_id)["issues_json"]
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))


def md_text(directory: Path, name: str) -> str | None:
    """フォルダ内の md の中身（名前はそのフォルダ直下の .md だけ許す）。"""
    path = directory / name
    if path.parent != directory or path.suffix != ".md" or not path.exists():
        return None
    return path.read_text(encoding="utf-8")


def md_paths(import_id: int) -> list[Path]:
    directory = import_files(import_id)["md"]
    return sorted(directory.glob("*.md"), key=lambda p: p.name) if directory.exists() else []


def _write_md_dir(target: Path, files: list[MdFile]) -> None:
    tmp = target.with_name(target.name + ".new")
    shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True)
    for f in files:
        (tmp / f.name).write_bytes(f.data)
    shutil.rmtree(target, ignore_errors=True)
    tmp.rename(target)


# ---- 設定と表の範囲 ---------------------------------------------------------------------------

def spec_for_import(imp: dict):
    if not imp.get("template_version_id"):
        return None
    version = store.get_version(imp["template_version_id"])
    return version["spec"] if version else None


def open_import_source(imp: dict):
    src = imp.get("source") or {}
    options = {key: src[key] for key in ("encoding", "delimiter", "errors") if src.get(key)}
    options["max_cells"] = current_app.config.get("EXCEL_MAX_CELLS", 500000)
    return open_source(upload_path(imp["stored_path"]), imp["file_name"], options)


def layout_for_import(source, imp: dict, spec):
    """保存した範囲（見出し行・データ終了行）で表の形を決め直す。未指定なら自動判定。"""
    src = imp.get("source") or {}
    sheet = src.get("sheet")
    if not sheet:
        sheets = [s for s in source.sheets() if not s.hidden] or source.sheets()
        sheet = sheets[0].name
    header_rows = [int(r) for r in (src.get("header_rows") or []) if r] or None
    anchors = list((spec.header or {}).get("anchors") or []) if spec is not None else None
    return guess_layout(source, sheet, anchors=anchors or None, header_row=src.get("header_row") or None,
                        data_end=src.get("data_end_row") or None, header_rows=header_rows)


# ---- 読み込み（ジョブ） ----------------------------------------------------------------------------

def run_read(ctx, import_id: int) -> dict:
    """ジョブ: 保存した範囲と設定で全行を読み、正規化・チェックして行データと問題一覧を書く。"""
    imp = store.get_import(import_id)
    version = store.get_version(imp["template_version_id"]) if imp and imp.get("template_version_id") else None
    if version is None:
        store.update_import(import_id, status="failed", stats={"error": "取り込み設定が見つかりません"})
        raise PipelineError("取り込み設定が見つかりません")
    spec = version["spec"]
    try:
        ctx.progress(phase="読み込み", done=0, total=0)
        source = open_import_source(imp)
        layout = layout_for_import(source, imp, spec)
        ctx.check_cancel()

        def on_progress(done, total):
            ctx.progress(phase="読み込み", done=done, total=total)
            ctx.check_cancel()

        records, row_issues, stats = read_records(
            source, {"sheet": layout.sheet, "file_name": imp["file_name"]}, layout, spec, on_progress=on_progress)
        issues = row_issues + run_checks(records, spec, stats)
        files = import_files(import_id)
        ctx.progress(phase="保存", done=len(records), total=len(records))
        write_rows(files["rows"], records)
        write_atomic(files["issues_json"], json.dumps([i.to_dict() for i in issues], ensure_ascii=False).encode("utf-8"))
        write_atomic(files["issues_csv"], outputs.issues_csv(issues))
        for directory in (files["md"], files["preview"]):
            shutil.rmtree(directory, ignore_errors=True)
        st = stats.to_dict()
        st.update({
            "spec_hash": version["spec_hash"], "template_version_id": version["id"],
            "layout": {"sheet": layout.sheet, "table_kind": layout.table_kind, "header_rows": layout.header_rows,
                       "data_start": layout.data_start, "data_end": layout.data_end, "headers": layout.headers},
            "issue_counts": count_levels(issues),
        })
        store.update_import(import_id, status="preview", stats=st, rows_path=str(files["rows"]),
                            issues_path=str(files["issues_csv"]))
        return {"records": len(records), **st["issue_counts"]}
    except jobs.JobCancelled:
        store.update_import(import_id, status="uploaded")
        raise
    except Exception as exc:
        message = str(exc) if isinstance(exc, (PipelineError, UploadError)) else f"{exc.__class__.__name__}: {exc}"
        store.update_import(import_id, status="failed", stats={**(imp.get("stats") or {}), "error": message})
        raise


def start_read_job(import_id: int) -> int:
    job_id = jobs.start_job("table_read", "table_import", import_id, lambda ctx: run_read(ctx, import_id),
                            {"import_id": import_id})
    store.update_import(import_id, status="reading", job_id=job_id)
    return job_id


# ---- Markdown の作成 ---------------------------------------------------------------------------

def usable_ai_results(import_id: int, imp: dict, spec) -> dict:
    """照合に通った AI の結果のうち、今の行の内容と合うものだけ（md 用の形 {key: {"status","result"}}）。"""
    if spec.log_stage is None or not imp.get("template_id"):
        return {}
    from aiproc import items as ai_items
    from aiproc import runner

    accepted = ai_items.results_for_render(imp["template_id"], "log")
    if not accepted:
        return {}
    data = runner.load_rows_for_ai(import_id)
    by_key = ai_items.items_by_key(imp["template_id"], "log")
    out = {}
    for w in runner.prepare_works(data, ["log"], list(accepted)):
        item = by_key.get(w.row_key)
        if item is None or ai_items.is_outdated(item, item.get("template_version_id"), w.source_hash, w.context_hash,
                                                w.segments_hash):
            continue
        out[w.row_key] = {"status": "ok", "result": accepted[w.row_key]}
    return out


def render_files(import_id: int, imp: dict, spec, records: list[dict] | None = None) -> list[MdFile]:
    if records is None:
        records = load_rows(import_id)
    return render_all(spec, records, usable_ai_results(import_id, imp, spec), {"coverage": {}})


def preview_files(import_id: int, imp: dict, spec) -> list[dict]:
    """プレビュー用に全 md を作る（行・設定・AI結果が変わらなければ前回の結果を使う）。"""
    from tables.spec import spec_hash

    base = import_files(import_id)
    rows_path = base["rows"]
    row = database.get_db().execute(
        "SELECT COUNT(*), COALESCE(MAX(updated_at), '') FROM ai_items WHERE template_id = ?",
        (imp.get("template_id") or 0,)).fetchone()
    signature = json.dumps([spec_hash(spec), rows_path.stat().st_mtime_ns if rows_path.exists() else 0, list(row)])
    index_path = base["preview"] / "index.json"
    if index_path.exists():
        try:
            index = json.loads(index_path.read_text(encoding="utf-8"))
            if index.get("signature") == signature:
                return index["files"]
        except (ValueError, OSError):
            pass
    files = render_files(import_id, imp, spec)
    _write_md_dir(base["preview"], files)
    listing = [{"name": f.name, "kind": f.kind, "size": len(f.data),
                "records": sum(1 for line in f.text.split("\n") if line.startswith("## ")) if f.kind == "records" else None}
               for f in files]
    index_path.write_text(json.dumps({"signature": signature, "files": listing}, ensure_ascii=False), encoding="utf-8")
    return listing


def run_render(ctx, import_id: int) -> dict:
    """ジョブ: この取り込みの記録から全 Markdown を作り、取り込みのフォルダに保存する。"""
    imp = store.get_import(import_id)
    try:
        spec = spec_for_import(imp)
        if spec is None:
            raise PipelineError("取り込み設定が見つかりません")
        ctx.progress(phase="記録の読み込み", done=0, total=3)
        records = load_rows(import_id)
        ctx.check_cancel()
        ctx.progress(phase="Markdownの作成", done=1, total=3)
        files = render_files(import_id, imp, spec, records)
        ctx.check_cancel()
        ctx.progress(phase="保存", done=2, total=3)
        _write_md_dir(import_files(import_id)["md"], files)
        stats = imp.get("stats") or {}
        stats["output"] = {"files": len(files), "records": len(records)}
        store.mark_version_used(imp["template_version_id"])
        store.update_import(import_id, status="confirmed", confirmed_at=database.now(), stats=stats)
        ctx.progress(phase="完了", done=3, total=3)
        return {"files": len(files), "records": len(records)}
    except Exception:
        store.update_import(import_id, status="preview")
        raise


def start_render_job(import_id: int) -> int:
    job_id = jobs.start_job("table_render", "table_import", import_id, lambda ctx: run_render(ctx, import_id),
                            {"import_id": import_id})
    store.update_import(import_id, status="confirming", job_id=job_id)
    return job_id


# ---- ダウンロード ----------------------------------------------------------------------------

def build_download(import_id: int, imp: dict, spec) -> bytes:
    """RAG投入用/*.md と 管理用_RAGには入れない/（正規化データ・問題一覧・取込レポート）の zip。"""
    paths = md_paths(import_id)
    extras = {
        "正規化データ.csv": outputs.normalized_csv(spec, load_rows(import_id)),
        "問題一覧.csv": outputs.issues_csv(load_issues(import_id)),
        "取込レポート.csv": outputs.report_csv(outputs.import_report_items(imp, spec, len(paths))),
    }
    return outputs.build_zip([(p.name, p.read_bytes()) for p in paths], extras)
