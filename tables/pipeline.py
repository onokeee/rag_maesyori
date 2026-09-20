"""一覧表の取り込みの実行関数（ジョブ本体）と、画面から呼ぶ補助。

- run_read: 保存した範囲と設定で全行を読む → 正規化 → チェック → rows.jsonl.gz / issues.json / issues.csv
- run_render: その取り込みの記録（＋照合に通った AI 整形の結果）から全 md を作り、取り込みのフォルダに保存
- build_download: md と管理用CSV をまとめた zip
期間の置き換え・投入済みとの差分・取り消しはしない（取り込みごとに、その取り込みの内容だけで md を作る）。
置き場所: TABLES_DIR/imports/<import_id>/（rows.jsonl.gz, issues.*, md/, preview_md/, source_cache.json）
画面とジョブは import_source() の控え（tables.source_cache）を通して表を読み、同じ計算を繰り返さない。
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
from tables.markdown import MdFile, render_all
from tables.normalize import read_records
from tables.source import EXCEL_EXTENSIONS, open_source
from tables.source_cache import ImportSource

ROWS_FILE = "rows.jsonl.gz"
ISSUES_CSV = "issues.csv"
ISSUES_JSON = "issues.json"
MD_DIR = "md"
PREVIEW_DIR = "preview_md"


class PipelineError(jobs.JobError):
    """利用者に見せる日本語メッセージの処理エラー（そのまま画面に出る）。"""


# ---- ファイル -----------------------------------------------------------------------------

def import_dir(import_id: int) -> Path:
    return Path(current_app.config["TABLES_DIR"]) / "imports" / str(int(import_id))


def import_files(import_id: int) -> dict[str, Path]:
    base = import_dir(import_id)
    return {"dir": base, "rows": base / ROWS_FILE, "issues_csv": base / ISSUES_CSV, "issues_json": base / ISSUES_JSON,
            "md": base / MD_DIR, "preview": base / PREVIEW_DIR}


# 取り込み1件を消すのは core/purge.py の purge_table_import（アップロードしたファイル・このフォルダ・
# DB の行・AI整形の控えをまとめて消す。design.md 3.3）


def write_atomic(path: Path, data: bytes) -> None:
    """一時ファイルに書いてから置き換える（途中で止まっても前のファイルが残る）。

    置き場所（imports/<id>/）は作らない。消された取り込みのフォルダを書き戻さないため（design.md 3.3）。
    """
    path = Path(path)
    tmp = path.with_name(f".{path.name}.tmp")
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)


def write_rows(path: Path, records) -> None:
    """記録を gzip の JSON Lines で保存（mtime=0・キー順固定で決定的）。一時ファイルに直接書いてから置き換える。

    置き場所は作らない（write_atomic と同じ）。
    """
    path = Path(path)
    tmp = path.with_name(f".{path.name}.tmp")
    with open(tmp, "wb") as raw, gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0, compresslevel=6) as gz:
        buf = io.BufferedWriter(gz, 1024 * 1024)
        for rec in records:
            d = rec.to_dict() if hasattr(rec, "to_dict") else dict(rec)
            buf.write(json.dumps(d, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8") + b"\n")
        buf.flush()
        buf.detach()
    os.replace(tmp, path)


def load_rows(import_id: int, offset: int = 0, limit: int | None = None) -> list[dict]:
    path = import_files(import_id)["rows"]
    if not path.exists():
        return []
    with gzip.open(path, "rt", encoding="utf-8") as f:
        rows = [json.loads(line) for line in f if line.strip()]
    return rows[offset: offset + limit] if limit is not None else rows[offset:]


def load_rows_page(import_id: int, offset: int, limit: int) -> tuple[list[dict], int]:
    """offset から limit 件の記録と全件数。範囲外の行は JSON を読まずに数えるだけ。"""
    path = import_files(import_id)["rows"]
    if not path.exists():
        return [], 0
    rows: list[dict] = []
    total = 0
    with gzip.open(path, "rb") as f:
        for line in f:
            if not line.strip():
                continue
            if offset <= total < offset + limit:
                rows.append(json.loads(line))
            total += 1
    return rows, total


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
    from core.files import path_limit

    tmp = target.with_name(target.name + ".new")
    limit = path_limit()
    if limit is not None and files:
        longest = max(files, key=lambda f: len(f.name)).name
        if len(os.path.abspath(tmp)) + 1 + len(longest) > limit:
            # Windows のパスの長さの上限（260文字）を超えると、書けずに分かりにくいエラーで止まる
            raise PipelineError(
                f"Markdownのファイル名が長すぎて、サーバーのデータの置き場所に書けません（最長 {len(longest)}文字）。"
                "取り込み設定の設定名・ファイル名の先頭を短くするか、アプリを浅いフォルダに置いてください")
    shutil.rmtree(tmp, ignore_errors=True)
    # 親（imports/<id>/）は作り直さない。消された取り込みのフォルダを復活させないため（design.md 3.3）
    tmp.mkdir()
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


def _source_options(imp: dict) -> dict:
    src = imp.get("source") or {}
    options = {key: src[key] for key in ("encoding", "delimiter", "errors") if src.get(key)}
    options["max_cells"] = current_app.config.get("EXCEL_MAX_CELLS", 500000)
    return options


def open_import_source(imp: dict, sheet_stats: dict | None = None):
    options = _source_options(imp)
    if sheet_stats:
        options["sheet_stats"] = sheet_stats
    return open_source(upload_path(imp["stored_path"]), imp["file_name"], options)


def import_source(imp: dict, real=None) -> ImportSource:
    """控え付きの表ソース（元のファイルは必要になったときだけ開く）。ファイルや読み込み設定が変われば控えは作り直す。"""
    key = json.dumps([imp.get("file_hash") or "", imp["stored_path"], imp["file_name"], _source_options(imp)],
                     ensure_ascii=False, sort_keys=True)
    kind = "excel" if Path(imp["file_name"]).suffix.lower() in EXCEL_EXTENSIONS else "csv"
    return ImportSource(import_dir(imp["id"]), key, kind, imp["file_name"],
                        lambda stats: open_import_source(imp, stats), real=real)


def layout_for_import(source, imp: dict, spec):
    """保存した範囲（見出し行・データ終了行）で表の形を決め直す。未指定なら自動判定。同じ条件の結果は控えから。"""
    cached = source if isinstance(source, ImportSource) else import_source(imp, real=source)
    src = imp.get("source") or {}
    sheet = src.get("sheet")
    if not sheet:
        sheets = [s for s in cached.sheets() if not s.hidden] or cached.sheets()
        sheet = sheets[0].name
    header_rows = [int(r) for r in (src.get("header_rows") or []) if r] or None
    anchors = list((spec.header or {}).get("anchors") or []) if spec is not None else None
    return cached.layout(sheet, anchors=anchors or None, header_row=src.get("header_row") or None,
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
        cached = import_source(imp)
        source = cached.real
        layout = layout_for_import(cached, imp, spec)
        ctx.check_cancel()

        def on_progress(done, total):
            ctx.progress(phase="読み込み", done=done, total=total)
            ctx.check_cancel()

        records, row_issues, stats = read_records(
            source, {"sheet": layout.sheet, "file_name": imp["file_name"]}, layout, spec, on_progress=on_progress)
        issues = row_issues + run_checks(records, spec, stats)
        files = import_files(import_id)
        ctx.check_cancel()
        ctx.progress(phase="保存", done=len(records), total=len(records))
        if store.get_import(import_id) is None:
            # 読み込みの間に渡し終えて（または削除されて）消えた。行データ・問題一覧を書き戻さない（design.md 3.3）
            raise PipelineError("取り込みが削除されたため、読み込んだ内容は保存しませんでした")
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
        # 例外の型と本文は core.jobs のログに残す。画面には Python の例外名を出さない
        message = str(exc) if isinstance(exc, (PipelineError, UploadError)) else (
            "表の読み込み中に予期しないエラーが起きました。もう一度読み込んでも直らない場合は、"
            "表の範囲・見出し行や列の対応づけを見直してください")
        store.update_import(import_id, status="failed", stats={**(imp.get("stats") or {}), "error": message})
        raise


def start_read_job(import_id: int) -> int:
    return _start_import_job(import_id, "table_read", "reading", run_read)


def _start_import_job(import_id: int, kind: str, status: str, fn) -> int:
    """取り込みの状態を先に書いてからジョブを始める。

    ジョブを先に始めると、ジョブがすぐ終わって書いた「preview」「confirmed」を、あとの「reading」「confirming」で
    上書きしてしまい、待ち画面が「途中で止まりました」と誤って出す。job_id だけはジョブを作ってから書く。
    """
    before = (store.get_import(import_id) or {}).get("status")
    store.update_import(import_id, status=status)
    try:
        job_id = jobs.start_job(kind, "table_import", import_id, lambda ctx: fn(ctx, import_id), {"import_id": import_id})
    except Exception:
        if before is not None:
            store.update_import(import_id, status=before)
        raise
    store.update_import(import_id, job_id=job_id)
    return job_id


# ---- Markdown の作成 ---------------------------------------------------------------------------

def usable_ai_results(import_id: int, imp: dict, spec) -> dict:
    """照合に通った AI の結果のうち、今の行の内容と合うものだけ（md 用の形 {key: {"status","result"}}）。"""
    if spec.log_stage is None or not imp.get("template_id"):
        return {}
    from aiproc import items as ai_items
    from aiproc import runner

    accepted = ai_items.results_for_render(imp["template_id"], "log", import_id=import_id)
    if not accepted:
        return {}
    data = runner.load_rows_for_ai(import_id)
    by_key = ai_items.items_by_key(imp["template_id"], "log", import_id=import_id)
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


def _preview_signature(import_id: int, imp: dict, spec) -> str:
    """プレビューの md を作った入力（設定・行データ・AI の結果）の目印。"""
    from tables.spec import spec_hash

    rows_path = import_files(import_id)["rows"]
    row = database.get_db().execute(
        # この取り込みの AI の結果だけ（usable_ai_results が読む範囲と同じ）。同じ設定の別の取り込みでは変わらない
        "SELECT COUNT(*), COALESCE(MAX(updated_at), '') FROM ai_items WHERE template_id = ? AND import_id = ?",
        (imp.get("template_id") or 0, import_id)).fetchone()
    return json.dumps([spec_hash(spec), rows_path.stat().st_mtime_ns if rows_path.exists() else 0, list(row)])


def _preview_index(import_id: int, signature: str) -> dict | None:
    index_path = import_files(import_id)["preview"] / "index.json"
    try:
        index = json.loads(index_path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return None
    return index if isinstance(index, dict) and index.get("signature") == signature else None


def _md_dir_from_preview(import_id: int, signature: str) -> int | None:
    """同じ入力で作ったプレビューの md を md/ に置く（作り直さない）。置いたファイル数。そろっていなければ None。

    作成は決定的なので、作り直しても同じバイト列になる。中身は読まずにハードリンク（できなければコピー）で置く
    （Windows では書いたばかりの多数の小さなファイルを読み直すと遅いため）。ファイルは書き換えずに作り直すので共有してよい。
    """
    index = _preview_index(import_id, signature)
    if index is None:
        return None
    files = import_files(import_id)
    preview, target = files["preview"], files["md"]
    names = [item["name"] for item in index["files"]]
    try:
        if any((preview / item["name"]).stat().st_size != item["size"] for item in index["files"]):
            return None
    except (OSError, KeyError, TypeError):
        return None
    tmp = target.with_name(target.name + ".new")
    shutil.rmtree(tmp, ignore_errors=True)
    try:
        tmp.mkdir()  # 親は作り直さない（消された取り込みのフォルダを復活させない）
    except OSError:
        return None
    try:
        for name in names:
            try:
                os.link(preview / name, tmp / name)
            except OSError:
                shutil.copyfile(preview / name, tmp / name)
    except OSError:
        shutil.rmtree(tmp, ignore_errors=True)
        return None
    shutil.rmtree(target, ignore_errors=True)
    tmp.rename(target)
    return len(names)


def preview_signature(import_id: int, imp: dict, spec) -> str:
    """プレビューの md を作った入力の目印（画面から、作り直しが要るかを見るために使う）。"""
    return _preview_signature(import_id, imp, spec)


def _record_count(f) -> int | None:
    """記録ファイル1つに入っている記録の件数（画面の「記録」の列）。

    大きい記録は「（続きn/m）」の見出しに分かれるが、それは1件の記録の続きなので数えない
    （数えると、上に出る「記録件数 60件」と食い違う。tables.markdown._split_record）。
    """
    if f.kind != "records":
        return None
    return sum(1 for line in f.text.split("\n") if line.startswith("## ") and "（続き" not in line)


def ready_preview_files(import_id: int, imp: dict, spec) -> list[dict] | None:
    """すでに作ってあるプレビューの一覧。作っていなければ None（作るのはジョブ run_preview）。"""
    index = _preview_index(import_id, _preview_signature(import_id, imp, spec))
    return index["files"] if index is not None else None


def preview_files(import_id: int, imp: dict, spec, ctx=None) -> list[dict]:
    """プレビュー用に全 md を作る（行・設定・AI結果が変わらなければ前回の結果を使う）。"""
    base = import_files(import_id)
    signature = _preview_signature(import_id, imp, spec)
    index_path = base["preview"] / "index.json"
    index = _preview_index(import_id, signature)
    if index is not None:
        return index["files"]
    if ctx is not None:
        ctx.progress(phase="Markdownの作成", done=1, total=3)
    files = render_files(import_id, imp, spec)
    if ctx is not None:
        ctx.check_cancel()
        ctx.progress(phase="保存", done=2, total=3)
    if store.get_import(import_id) is None:
        # 作っている間に取り込みが削除された。消したフォルダに記録の md を作り直さない（design.md 3.3）
        raise PipelineError("取り込みが削除されたため、Markdownの下書きは作りませんでした")
    _write_md_dir(base["preview"], files)
    listing = [{"name": f.name, "kind": f.kind, "size": len(f.data), "records": _record_count(f)}
               for f in files]
    index_path.write_text(json.dumps({"signature": signature, "files": listing}, ensure_ascii=False), encoding="utf-8")
    return listing


def run_preview(ctx, import_id: int) -> dict:
    """ジョブ: 「内容とファイルの確認」に出す md をまとめて作る（確定はしない）。

    件数が多いと十数秒かかるので、画面（GET）の中では作らず、待ち画面で進み具合を出せるようにする。
    """
    imp = store.get_import(import_id)
    spec = spec_for_import(imp)
    if spec is None:
        raise PipelineError("取り込み設定が見つかりません")
    ctx.progress(phase="記録の読み込み", done=0, total=3)
    ctx.check_cancel()
    files = preview_files(import_id, imp, spec, ctx=ctx)
    ctx.progress(phase="完了", done=3, total=3)
    return {"files": len(files)}


def start_preview_job(import_id: int, signature: str) -> int:
    return jobs.start_job("table_preview", "table_import", import_id, lambda ctx: run_preview(ctx, import_id),
                          {"import_id": import_id, "signature": signature})


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
        if store.get_import(import_id) is None:
            # 読み込みの間に渡し終えて（または削除されて）消えた。md を作り直さない（design.md 3.3）
            raise PipelineError("取り込みが削除されたため、Markdownは作りませんでした")
        # 確認画面で同じ入力から作ったプレビューがあれば、それを置く
        file_count = _md_dir_from_preview(import_id, _preview_signature(import_id, imp, spec))
        if file_count is None:
            files = render_files(import_id, imp, spec, records)
            ctx.check_cancel()
            ctx.progress(phase="保存", done=2, total=3)
            if store.get_import(import_id) is None:
                raise PipelineError("取り込みが削除されたため、Markdownは作りませんでした")
            _write_md_dir(import_files(import_id)["md"], files)
            file_count = len(files)
        stats = imp.get("stats") or {}
        stats["output"] = {"files": file_count, "records": len(records)}
        store.mark_version_used(imp["template_version_id"])
        store.update_import(import_id, status="confirmed", confirmed_at=database.now(), stats=stats)
        ctx.progress(phase="完了", done=3, total=3)
        return {"files": file_count, "records": len(records)}
    except Exception:
        store.update_import(import_id, status="preview")
        raise


def start_render_job(import_id: int) -> int:
    return _start_import_job(import_id, "table_render", "confirming", run_render)


# ---- ダウンロード ----------------------------------------------------------------------------

def build_download_files(import_id: int, imp: dict, spec) -> tuple[list[tuple[str, bytes]], dict[str, bytes]]:
    """渡すファイルの中身: (RAG投入用の md の [(名前, 中身)], 管理用のファイル {名前: 中身})。

    zip（build_download）と保存先フォルダへの保存（views.tables.save_to_folder）で同じ中身にするため、ここで1回だけ作る。
    """
    paths = md_paths(import_id)
    extras = {
        "正規化データ.csv": outputs.normalized_csv(spec, load_rows(import_id)),
        "問題一覧.csv": outputs.issues_csv(load_issues(import_id)),
        "取込レポート.csv": outputs.report_csv(outputs.import_report_items(imp, spec, len(paths))),
    }
    return [(p.name, p.read_bytes()) for p in paths], extras


def build_download(import_id: int, imp: dict, spec) -> bytes:
    """RAG投入用/*.md と 管理用_RAGには入れない/（正規化データ・問題一覧・取込レポート）の zip。"""
    md_files, extras = build_download_files(import_id, imp, spec)
    return outputs.build_zip(md_files, extras)
