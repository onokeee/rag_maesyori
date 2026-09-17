"""一覧表を取り込む（design.md 2.4）。

流れ: ファイルを選ぶ → 取り込み設定を選ぶ → 表の範囲と見出しを確認 → 列の対応づけ（設定を作る/変える時）
      → AI整形（任意。追記ログ列がある時だけ） → 内容とファイルの確認 → 完了（ダウンロード）
範囲の決まり（利用者の判断）: 期間の置き換え・投入済みとの差分・取り消しはしない。取り込みごとに
その取り込みの記録だけから全 Markdown を作り、全ファイルを zip で渡す。クロス集計と名寄せ辞書は扱わない。
読み込み・Markdown 作成・AI整形は core.jobs のジョブ（tables.pipeline / aiproc.runner）で動かし、待ち画面で進み具合を出す。
"""
from __future__ import annotations

import io
import sqlite3
from datetime import datetime
from pathlib import Path

from flask import Blueprint, abort, current_app, flash, jsonify, redirect, render_template, request, send_file, url_for

from core import jobs
from core.files import UploadError, precheck_excel, remove_upload, save_upload, upload_path
from core.naming import safe_filename_part
from models import database
from tables import outputs, pipeline, store
from tables.checks import count_levels, has_blocking
from tables.detect import guess_layout, looks_like_list, sample_data_rows
from tables.mapping import match_templates, suggest_columns
from tables.markdown import ai_point_lines, people_index_for, record_block
from tables.source import open_source
from tables.spec import (
    COLUMN_ROLES, COLUMN_TYPES, MD_MODES, resolve_columns, spec_from_suggestions, validate_spec,
)

bp = Blueprint("tables", __name__, url_prefix="/tables")

STEPS = ["ファイルを選ぶ", "取り込み設定を選ぶ", "表の範囲と見出しを確認", "列の対応づけ", "AI整形（任意）",
         "内容とファイルの確認", "完了（ダウンロード）"]
EXCEL_EXT = {".xlsx", ".xlsm"}
GRID_ROWS = 60
GRID_COLS = 30
GRID_CELL_CHARS = 40
DATA_PAGE = 100
ISSUES_SHOWN = 200

ROLE_LABELS = {"key": "記録番号", "date": "日付", "entity": "設備など（番号）", "entity_label": "設備名など",
               "category": "区分", "measure": "数値（集計する）", "text": "文章", "log": "追記ログ", "person": "人名",
               "attribute": "その他"}
TYPE_LABELS = {"code": "コード", "string": "文字", "text": "長文", "date": "日付", "datetime": "日時", "time": "時刻",
               "number": "数値", "enum": "選択肢", "status": "状態"}
MD_LABELS = {"body": "本文", "attribute": "項目として出す", "omit": "出さない"}
KIND_LABELS = {"header": "見出し", "data": "データ", "subtotal": "小計・合計", "note": "注記", "continuation": "継続行",
               "excluded": "除外", "title": "表題", "blank": "空行"}
TABLE_KIND_LABELS = {"list": "一覧表", "crosstab": "クロス集計", "form_like": "帳票らしい", "unknown": "不明"}
SCOPE_LABELS = {"pending": "まだ整形していない行と、内容が変わった行", "errors": "エラーになった行だけ",
                "flagged": "要確認の行だけ", "all": "すべての行をやり直す"}
MATCHED_LABELS = {"template": "設定", "dictionary": "辞書", "similar": "似た見出し", "none": "", "": ""}


# ---- 共通 ---------------------------------------------------------------------------------

def _steps_ctx(current: int) -> dict:
    return {"steps": STEPS, "step": current}


def _load_import(import_id: int) -> dict:
    imp = store.get_import(import_id)
    if imp is None:
        abort(404)
    return imp


def _spec_for(imp: dict):
    return pipeline.spec_for_import(imp)


def _is_excel(imp: dict) -> bool:
    return Path(imp["file_name"]).suffix.lower() in EXCEL_EXT


def _open(imp: dict):
    return pipeline.open_import_source(imp)


def _sheet(imp: dict, source) -> str:
    src = imp.get("source") or {}
    names = [s.name for s in source.sheets()]
    if src.get("sheet") in names:
        return src["sheet"]
    visible = [s.name for s in source.sheets() if not s.hidden]
    return (visible or names)[0]


def _layout(imp: dict, source, sheet: str, header_rows=None, data_end=None, use_saved=True):
    src = imp.get("source") or {}
    if use_saved:
        header_rows = header_rows or [int(r) for r in src.get("header_rows") or [] if r] or None
        data_end = data_end or src.get("data_end_row") or None
    return guess_layout(source, sheet, header_rows=header_rows or None, data_end=data_end or None)


def _after_read_url(import_id: int, spec) -> str:
    if spec is not None and spec.log_stage is not None:
        return url_for("tables.ai", import_id=import_id)
    return url_for("tables.preview", import_id=import_id)


def _json_error(message: str, status: int = 400, **extra):
    return jsonify({"error": message, **extra}), status


def _payload() -> dict:
    data = request.get_json(silent=True)
    return data if isinstance(data, dict) else {}


def _int(value, default=None):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def _int_list(value) -> list[int]:
    if isinstance(value, (list, tuple)):
        items = value
    else:
        items = str(value or "").replace("、", ",").replace(" ", ",").split(",")
    out = sorted({n for n in (_int(x) for x in items) if n and n > 0})
    return out


# ---- 待ち画面 ------------------------------------------------------------------------------

def _busy_page(imp: dict):
    """読み込み中・作成中なら待ち画面を返す。ジョブが終わっているのに状態が残っていれば直す。"""
    if imp.get("status") not in ("reading", "confirming"):
        return None
    job = jobs.get_job(imp["job_id"]) if imp.get("job_id") else None
    if job is None or job.get("finished"):
        fresh = store.get_import(imp["id"])
        if fresh and fresh.get("status") in ("reading", "confirming"):
            if fresh["status"] == "reading":
                stats = {**(fresh.get("stats") or {}), "error": "読み込みが途中で止まりました。もう一度読み込んでください"}
                store.update_import(imp["id"], status="failed", stats=stats)
            else:
                store.update_import(imp["id"], status="preview")
                flash("Markdownの作成が途中で止まりました。もう一度 [確定してMarkdownを作成] を押してください", "error")
        return redirect(request.path)
    title = "表を読み込んでいます" if imp["status"] == "reading" else "Markdownを作成しています"
    step = 6 if imp["status"] == "reading" else 7
    return render_template("tables/wait.html", imp=imp, job=job, title=title,
                           job_url=url_for("tables.api_job", job_id=job["id"]), **_steps_ctx(step))


# ---- 1 ファイルを選ぶ ------------------------------------------------------------------------

@bp.get("/")
def index():
    return redirect(url_for("tables.new"))


@bp.get("/new")
def new():
    recent = store.list_imports(limit=10)
    return render_template("tables/new.html", recent=recent, template_count=len(store.list_templates()),
                           **_steps_ctx(1))


@bp.post("/upload")
def upload():
    cfg = current_app.config
    try:
        stored = save_upload(request.files.get("file"), "tables", set(cfg["TABLE_ALLOWED_EXTENSIONS"]),
                             int(cfg["TABLE_MAX_UPLOAD_BYTES"]))
    except UploadError as exc:
        flash(str(exc), "error")
        return redirect(url_for("tables.new"))
    path = upload_path(stored.stored_path)
    source_info: dict = {}
    try:
        if Path(stored.file_name).suffix.lower() in EXCEL_EXT:
            precheck_excel(path)
            source = open_source(path, stored.file_name, {"max_cells": cfg.get("EXCEL_MAX_CELLS", 500000)})
            sheets = [s for s in source.sheets() if not s.hidden] or source.sheets()
            source_info = {"kind": "excel", "sheet": sheets[0].name if sheets else None}
        else:
            source = open_source(path, stored.file_name)
            sniff = source.sniff
            source_info = {"kind": "csv", "encoding": source.encoding, "delimiter": source.delimiter,
                           "errors": "strict"}
            if sniff is not None:
                source_info.update(bom=sniff.bom, preamble_rows=sniff.preamble_rows, sniff_warnings=sniff.warnings,
                                   confidence=sniff.confidence)
    except UploadError as exc:
        remove_upload(stored.stored_path)
        flash(str(exc), "error")
        return redirect(url_for("tables.new"))
    import_id = store.create_import(stored.file_name, stored.file_hash, stored.stored_path, source=source_info)
    return redirect(url_for("tables.source", import_id=import_id))


# ---- 2 取り込み設定を選ぶ ----------------------------------------------------------------------

def _list_templates() -> list[dict]:
    return [t for t in store.list_templates() if t.get("spec") is not None]


@bp.get("/imports/<int:import_id>/source")
def source(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    src = imp.get("source") or {}
    error = None
    sheets, sheet, headers, list_like = [], None, [], {}
    candidates = []
    templates = _list_templates()
    try:
        source_obj = _open(imp)
        sheets = source_obj.sheets() if _is_excel(imp) else []
        sheet = _sheet(imp, source_obj)
        if _is_excel(imp) and len([s for s in sheets if not s.hidden]) <= 8:
            list_like = {s.name: looks_like_list(source_obj, s.name) for s in sheets if not s.hidden}
        layout = guess_layout(source_obj, sheet, max_scan_rows=200)
        headers = layout.headers
    except UploadError as exc:
        error = str(exc)
    if templates and headers:
        by_id = {id(t["spec"]): t for t in templates}
        for spec, matched, total in match_templates(headers, sheet or imp["file_name"], [t["spec"] for t in templates]):
            candidates.append({"template": by_id[id(spec)], "matched": matched, "total": total})
    selected = imp.get("template_id")
    if not selected and candidates and candidates[0]["total"] and candidates[0]["matched"] == candidates[0]["total"]:
        selected = candidates[0]["template"]["id"]
    return render_template(
        "tables/source.html", imp=imp, src=src, error=error, sheets=sheets, sheet=sheet, list_like=list_like,
        headers=headers, candidates=candidates, selected=selected or "new",
        new_name=src.get("new_template_name") or Path(imp["file_name"]).stem,
        encodings=[("utf-8-sig", "UTF-8（BOM付き）"), ("utf-8", "UTF-8"), ("cp932", "CP932（Shift_JIS）"),
                   ("shift_jis_2004", "Shift_JIS 2004"), ("utf-16", "UTF-16")],
        delimiters=[(",", "カンマ"), ("\t", "タブ"), (";", "セミコロン"), ("|", "縦棒")],
        **_steps_ctx(2))


@bp.post("/imports/<int:import_id>/source")
def save_source(import_id: int):
    imp = _load_import(import_id)
    src = dict(imp.get("source") or {})
    before = (src.get("sheet"), src.get("encoding"), src.get("delimiter"), src.get("errors"))
    if _is_excel(imp):
        if request.form.get("sheet"):
            src["sheet"] = request.form["sheet"]
    else:
        for name in ("encoding", "delimiter"):
            value = request.form.get(name)
            if value:
                src[name] = value
        src["errors"] = "replace" if request.form.get("replace_errors") == "on" else "strict"
    if (src.get("sheet"), src.get("encoding"), src.get("delimiter"), src.get("errors")) != before:
        src.pop("header_rows", None)
        src.pop("data_end_row", None)
    choice = request.form.get("template", "new")
    columns: dict = {"source": src}
    if choice == "new":
        name = request.form.get("new_template_name", "").strip()
        if not name:
            flash("新しい取り込み設定の名前を入力してください", "error")
            return redirect(url_for("tables.source", import_id=import_id))
        src["new_template_name"] = name
        columns.update(template_id=None, template_version_id=None)
    else:
        template = store.get_template(_int(choice, 0))
        if template is None:
            flash("取り込み設定が見つかりません", "error")
            return redirect(url_for("tables.source", import_id=import_id))
        columns.update(template_id=template["id"], template_version_id=template["current_version_id"])
    if imp["status"] not in ("reading", "confirming"):
        columns["status"] = "uploaded"  # 選び直したら読み込みからやり直す
    store.update_import(import_id, **columns)
    return redirect(url_for("tables.layout", import_id=import_id))


# ---- 3 表の範囲と見出しを確認 --------------------------------------------------------------------

def _grid(source, sheet: str, layout) -> tuple[list[dict], int]:
    classes = {rc.index: rc for rc in layout.row_classes}
    rows, width = [], 0
    for row in source.rows(sheet, 1, GRID_ROWS):
        cells = [(c.text or "")[:GRID_CELL_CHARS] for c in row.cells[:GRID_COLS]]
        while cells and not cells[-1]:
            cells.pop()
        width = max(width, len(cells))
        rc = classes.get(row.index)
        rows.append({"index": row.index, "cells": cells, "kind": rc.kind if rc else ("blank" if row.is_blank else "title"),
                     "reason": rc.reason if rc else "", "hidden": bool(row.hidden),
                     "strike": any(getattr(c, "strike", False) for c in row.cells if c.text)})
    return rows, max(width, 1)


def _layout_json(layout) -> dict:
    return {
        "sheet": layout.sheet, "table_kind": layout.table_kind,
        "table_kind_label": TABLE_KIND_LABELS.get(layout.table_kind, layout.table_kind),
        "header_rows": layout.header_rows, "data_start": layout.data_start, "data_end": layout.data_end,
        "headers": layout.headers, "warnings": layout.warnings,
        "counts": {KIND_LABELS.get(k, k): v for k, v in (layout.counts or {}).items()},
        "rows": {str(rc.index): {"kind": rc.kind, "reason": rc.reason} for rc in layout.row_classes},
    }


@bp.get("/imports/<int:import_id>/layout")
def layout(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet)
        rows, width = _grid(source_obj, sheet, guess)
    except UploadError as exc:
        flash(str(exc), "error")
        return redirect(url_for("tables.source", import_id=import_id))
    src = imp.get("source") or {}
    return render_template("tables/layout.html", imp=imp, layout=guess, info=_layout_json(guess), rows=rows,
                           width=width, sheet=sheet, data_end_saved=src.get("data_end_row") or "",
                           kind_labels=TABLE_KIND_LABELS, **_steps_ctx(3))


@bp.post("/imports/<int:import_id>/layout/detect")
def layout_detect(import_id: int):
    imp = _load_import(import_id)
    data = _payload()
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet, header_rows=_int_list(data.get("header_rows")) or None,
                        data_end=_int(data.get("data_end")), use_saved=False)
    except UploadError as exc:
        return _json_error(str(exc))
    return jsonify(_layout_json(guess))


@bp.post("/imports/<int:import_id>/layout")
def save_layout(import_id: int):
    imp = _load_import(import_id)
    header_rows = _int_list(request.form.get("header_rows"))
    data_end = _int(request.form.get("data_end_row"))
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet, header_rows=header_rows or None, data_end=data_end, use_saved=False)
    except UploadError as exc:
        flash(str(exc), "error")
        return redirect(url_for("tables.layout", import_id=import_id))
    if guess.table_kind == "crosstab":
        flash("月別集計のようなクロス集計の表には対応していません。1行＝1件の一覧表を選んでください", "error")
        return redirect(url_for("tables.layout", import_id=import_id))
    if not guess.headers or guess.data_end < guess.data_start:
        flash("見出し行とデータの範囲が見つかりません。見出し行の番号を指定してください", "error")
        return redirect(url_for("tables.layout", import_id=import_id))
    src = dict(imp.get("source") or {})
    src.update(sheet=sheet, header_rows=guess.header_rows, header_row=None, data_end_row=data_end or None)
    src.pop("headers_changed", None)
    spec = _spec_for(imp)
    if spec is not None:
        res = resolve_columns(spec, guess.headers)
        ignored = set((spec.header or {}).get("ignored") or [])
        added = [h for h in res.unused_headers if h not in ignored]
        if res.missing or added:
            src["headers_changed"] = {"missing": res.missing, "added": added}
    store.update_import(import_id, source=src)
    if spec is None or src.get("headers_changed"):
        return redirect(url_for("tables.columns", import_id=import_id))
    pipeline.start_read_job(import_id)
    return redirect(_after_read_url(import_id, spec))


# ---- 4 列の対応づけ ------------------------------------------------------------------------------

def _options_ctx() -> dict:
    return {"types": [(t, TYPE_LABELS.get(t, t)) for t in COLUMN_TYPES],
            "roles": [(r, ROLE_LABELS.get(r, r)) for r in COLUMN_ROLES],
            "md_modes": [(m, MD_LABELS.get(m, m)) for m in MD_MODES]}


def _settings_of(spec, default_name: str) -> dict:
    md = (spec.markdown or {}) if spec is not None else {}
    return {"name": spec.name if spec is not None else default_name,
            "description": spec.description if spec is not None else "",
            "file_prefix": md.get("file_prefix") or "",
            "group_by": md.get("group_by") or "month",
            "max_records_per_file": md.get("max_records_per_file") or 300,
            "omit_person": md.get("omit_person", True),
            "lightrag_hint": bool(md.get("lightrag_hint"))}


def _column_row(index: int, header: str, col=None, sugg=None, use: bool = True) -> dict:
    src = col if col is not None else sugg
    role = getattr(src, "role", "attribute") or "attribute"
    return {
        "index": index, "header": header, "use": use,
        "key": getattr(src, "key", "") or "", "display": getattr(src, "display", "") or header,
        "description": getattr(col, "description", "") if col is not None else "",
        "type": getattr(src, "type", "string") or "string", "unit": getattr(src, "unit", "") or "",
        "role": role, "fill_down_blank": bool(getattr(src, "fill_down_blank", False)),
        "md": getattr(src, "md", "attribute") or "attribute", "ai": role == "log",
        "examples": list(getattr(sugg, "examples", []) or [])[:3] if sugg is not None else [],
        "type_error_rate": getattr(sugg, "type_error_rate", 0.0) if sugg is not None else 0.0,
        "blank_rate": getattr(sugg, "blank_rate", 0.0) if sugg is not None else 0.0,
        "inferred_type": getattr(sugg, "inferred_type", "") if sugg is not None else "",
        "matched_by": MATCHED_LABELS.get(getattr(sugg, "matched_by", "") if sugg is not None else "", ""),
    }


def _build_spec(payload: dict, base_spec=None):
    """列の対応づけ表（JSON）から取り込み設定を作る。戻り値: (spec, errors)"""
    rows = [r for r in payload.get("columns") or [] if isinstance(r, dict)]
    used = [r for r in rows if r.get("use")]
    name = str(payload.get("name") or "").strip()
    suggestions = []
    for i, r in enumerate(used):
        role = str(r.get("role") or "attribute")
        type_ = str(r.get("type") or "string")
        if r.get("ai"):
            role, type_ = "log", "text"
        elif role == "log":
            role = "text"
        suggestions.append({
            "index": _int(r.get("index"), i), "header": str(r.get("header") or ""),
            "key": str(r.get("key") or "").strip() or None, "display": str(r.get("display") or "").strip(),
            "type": type_, "role": role, "unit": str(r.get("unit") or "").strip(), "md": str(r.get("md") or "attribute"),
            "fill_down_blank": bool(r.get("fill_down_blank")),
        })
    header_count = int(((base_spec.header or {}).get("rows") if base_spec is not None else None)
                       or _int(payload.get("header_rows_count"), 1) or 1)
    group_by = payload.get("group_by") if payload.get("group_by") in ("month", "entity_month") else "month"
    options = {"description": str(payload.get("description") or ""), "group_by": group_by, "period_grain": "all",
               "max_records_per_file": max(1, _int(payload.get("max_records_per_file"), 300) or 300)}
    spec = spec_from_suggestions(name, {"table_kind": "list", "header_rows": list(range(1, header_count + 1))},
                                 suggestions, options)
    for col, r in zip(spec.columns, used):
        col.description = str(r.get("description") or "").strip()
    spec.header["ignored"] = [str(r.get("header") or "") for r in rows if not r.get("use")]
    spec.markdown["file_prefix"] = str(payload.get("file_prefix") or "").strip() or name
    spec.markdown["omit_person"] = bool(payload.get("omit_person", True))
    spec.markdown["lightrag_hint"] = bool(payload.get("lightrag_hint"))
    if base_spec is not None:
        # 画面で扱わない細かい設定は前の設定から引き継ぐ
        for attr in ("name_patterns", "file_types", "na_tokens", "fiscal_year_start_month", "data_end", "exclude",
                     "continuation_rows", "checks", "custom_stages"):
            setattr(spec, attr, getattr(base_spec, attr))
        spec.header["anchors"] = (base_spec.header or {}).get("anchors") or []
        keys = {c.key for c in spec.columns}
        spec.markdown["title_columns"] = [k for k in (base_spec.markdown or {}).get("title_columns") or []
                                          if str(k).split(":")[0] in keys]
        spec.custom_stages = [s for s in spec.custom_stages if all(k in keys for k in s.inputs)]
        old = base_spec.log_stage
        if old is not None and spec.log_stage is not None and old.column == spec.log_stage.column:
            old.context_columns = [k for k in old.context_columns if k in keys]
            spec.log_stage = old
    return spec, validate_spec(spec)


@bp.get("/imports/<int:import_id>/columns")
def columns(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet)
        samples = sample_data_rows(source_obj, sheet, guess, 200)
    except UploadError as exc:
        flash(str(exc), "error")
        return redirect(url_for("tables.source", import_id=import_id))
    spec = _spec_for(imp)
    suggestions = suggest_columns(guess.headers, samples, spec)
    pos_to_col, ignored = {}, set()
    if spec is not None:
        res = resolve_columns(spec, guess.headers)
        pos_to_col = {pos: spec.column(key) for key, pos in res.positions.items()}
        ignored = set((spec.header or {}).get("ignored") or [])
    rows = []
    for s in suggestions:
        col = pos_to_col.get(s.index)
        if spec is None:
            use = s.blank_rate < 1.0
        else:
            use = col is not None or s.header not in ignored
        rows.append(_column_row(s.index, s.header, col, s, use))
    src = imp.get("source") or {}
    return render_template(
        "tables/columns.html", imp=imp, rows=rows, settings=_settings_of(spec, src.get("new_template_name") or
                                                                         Path(imp["file_name"]).stem),
        changed=src.get("headers_changed"), is_new=spec is None, header_rows_count=len(guess.header_rows) or 1,
        save_url=url_for("tables.save_columns", import_id=import_id), **_options_ctx(), **_steps_ctx(4))


@bp.post("/imports/<int:import_id>/columns")
def save_columns(import_id: int):
    imp = _load_import(import_id)
    payload = _payload()
    template = store.get_template(imp["template_id"]) if imp.get("template_id") else None
    base_spec = template["spec"] if template else None
    spec, errors = _build_spec(payload, base_spec)
    if errors:
        return _json_error(errors[0], errors=errors)
    try:
        if template is not None:
            version_id = store.save_template_version(template["id"], spec)
            template_id = template["id"]
        else:
            template_id, version_id = store.create_template(spec.name, spec)
    except sqlite3.IntegrityError:
        database.get_db().rollback()
        return _json_error(f"「{spec.name}」という名前の取り込み設定がすでにあります。別の名前にしてください")
    src = dict(imp.get("source") or {})
    src.pop("headers_changed", None)
    store.update_import(import_id, template_id=template_id, template_version_id=version_id, source=src)
    pipeline.start_read_job(import_id)
    return jsonify({"ok": True, "redirect": _after_read_url(import_id, spec)})


@bp.post("/imports/<int:import_id>/read")
def reread(import_id: int):
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None:
        return redirect(url_for("tables.columns", import_id=import_id))
    if imp["status"] in ("reading", "confirming"):
        return redirect(url_for("tables.preview", import_id=import_id))
    pipeline.start_read_job(import_id)
    return redirect(_after_read_url(import_id, spec))


def _not_ready(imp: dict, spec):
    """読み込みが終わっていない・失敗したときの行き先。"""
    if spec is None:
        return redirect(url_for("tables.source", import_id=imp["id"]))
    if imp["status"] == "failed":
        return render_template("tables/failed.html", imp=imp, error=(imp.get("stats") or {}).get("error") or "",
                               **_steps_ctx(6))
    if imp["status"] == "uploaded":
        return redirect(url_for("tables.layout", import_id=imp["id"]))
    return None


# ---- 5 AI整形（任意） ------------------------------------------------------------------------------

def _ai_job(import_id: int) -> dict | None:
    return jobs.latest_job("table_import", import_id, kind="ai_format")


@bp.get("/imports/<int:import_id>/ai")
def ai(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    spec = _spec_for(imp)
    not_ready = _not_ready(imp, spec)
    if not_ready:
        return not_ready
    if spec.log_stage is None:
        return redirect(url_for("tables.preview", import_id=import_id))
    from aiproc import items as ai_items
    from services import llm

    log_key = spec.log_stage.column
    col = spec.column(log_key)
    row_choices = []
    for rec in pipeline.load_rows(import_id):
        text = str((rec.get("values") or {}).get(log_key) or "").strip()
        if text:
            label = rec["key"] + "｜" + " ".join(text.split())[:40]
            row_choices.append((rec["key"], label))
            if len(row_choices) >= 200:
                break
    ready = llm.is_configured()
    external = False
    model = ""
    if ready:
        try:
            settings = llm.job_client_settings()
            external = not llm.is_local_endpoint(settings.get("chat_url") or "")
            model = settings.get("model") or ""
        except Exception:
            ready = False
    job = _ai_job(import_id)
    return render_template(
        "tables/ai.html", imp=imp, spec=spec, log_display=col.display if col else log_key, row_choices=row_choices,
        trial_keys=[k for k, _ in row_choices[:10]], ai_ready=ready, external=external, model=model, job=job,
        job_active=bool(job and not job.get("finished")), counts=ai_items.counts(imp["template_id"], "log"),
        scopes=SCOPE_LABELS, item_labels=ai_items.STATUS_LABELS, **_steps_ctx(5))


@bp.post("/imports/<int:import_id>/ai/split-preview")
def ai_split_preview(import_id: int):
    from aiproc import prompts, runner
    from logproc import format_author, format_when, review_notes
    from logproc import render_timeline

    _load_import(import_id)
    row_key = str(_payload().get("row_key") or "")
    try:
        data = runner.load_rows_for_ai(import_id)
        works = runner.prepare_works(data, ["log"], [row_key])
    except runner.AIJobError as exc:
        return _json_error(str(exc))
    if not works:
        return _json_error("行が見つかりません")
    w = works[0]
    parse = w.parse
    segments = [{
        "id": s.id, "body": s.body, "when": format_when(s.when), "when_estimated": bool(s.when and s.when.estimated),
        "author": format_author(s.author), "author_estimated": bool(s.author and s.author.estimated),
        "marks": list(s.marks or []), "identifiers": list(s.identifiers or []), "quantities": list(s.quantities or []),
        "plans": list(s.plans or []),
    } for s in parse.segments]
    return jsonify({
        "row_key": row_key, "kind": parse.kind, "order": parse.order, "text": parse.text, "segments": segments,
        "route": w.route, "reason": w.reason, "notes": review_notes(parse),
        "timeline": render_timeline(parse, w.entity_label, glossary=(w.stage.glossary or None) if w.stage else None),
        "sent_text": prompts.messages_text(w.messages) if w.messages else "",
    })


@bp.post("/imports/<int:import_id>/ai/trial")
def ai_trial(import_id: int):
    from aiproc import runner
    from services import llm

    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or spec.log_stage is None:
        return _json_error("AI整形の対象の列がありません")
    row_key = str(_payload().get("row_key") or "")
    try:
        data = runner.load_rows_for_ai(import_id)
        result = runner.trial_row(import_id, row_key, stage_ids=["log"], data=data)
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません。設定の「AI接続」で設定してください")
    except llm.LLMCallError as exc:
        return _json_error(str(exc))
    except runner.AIJobError as exc:
        return _json_error(str(exc))
    stage = next((s for s in result["stages"] if s.get("stage_id") == "log"), None) or {}
    accepted = stage.get("result") if stage.get("status") in ("ok", "flagged") else None
    record = next((r for r in data.rows if r["key"] == row_key), None)
    md = ""
    if record is not None:
        ai_results = {row_key: {"status": "ok", "result": accepted}} if accepted else {}
        md = "\n".join(record_block(record, spec, ai_results, people_index_for(spec, data.rows)))
    checks = stage.get("checks") or {}
    issues = [i.get("message", "") for i in (checks.get("issues") or []) if isinstance(i, dict)]
    return jsonify({
        "row_key": row_key, "status": stage.get("status"), "route": stage.get("route"), "reason": stage.get("reason"),
        "error": stage.get("error") or "", "points": ai_point_lines(accepted, spec.log_stage.glossary) if accepted else [],
        "timeline": stage.get("timeline") or [], "issues": issues, "markdown": md, "cached": stage.get("cached"),
        "stats": runner.trial_stats([result]),
    })


@bp.post("/imports/<int:import_id>/ai/estimate")
def ai_estimate(import_id: int):
    from aiproc import estimate as ai_estimate_mod
    from aiproc import runner
    from services import llm

    _load_import(import_id)
    data = _payload()
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    try:
        settings = llm.job_client_settings()
        result = ai_estimate_mod.estimate(import_id, [s for s in data.get("trials") or [] if isinstance(s, dict)],
                                          scope=scope, concurrency=_int(data.get("concurrency")) or None,
                                          settings=settings, stage_ids=["log"])
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません")
    except runner.AIJobError as exc:
        return _json_error(str(exc))
    return jsonify(result)


@bp.post("/imports/<int:import_id>/ai/run")
def ai_run(import_id: int):
    from aiproc import runner
    from services import llm

    imp = _load_import(import_id)
    data = _payload()
    if imp["status"] != "preview":
        return _json_error("表の読み込みが終わってから実行してください")
    job = _ai_job(import_id)
    if job and not job.get("finished"):
        return _json_error("AI整形はすでに実行中です")
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    concurrency = min(16, max(1, _int(data.get("concurrency"), 1) or 1))
    try:
        settings = llm.job_client_settings()
        if not llm.is_local_endpoint(settings.get("chat_url") or "") and not data.get("confirm_external"):
            return _json_error("対応内容が外部のAIサービスに送信されます。確認のチェックを入れてください")
        job_id = runner.start_ai_job(import_id, scope=scope, concurrency=concurrency, stage_ids=["log"])
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません")
    return jsonify({"ok": True, "job_id": job_id, "job_url": url_for("tables.api_job", job_id=job_id)})


@bp.post("/imports/<int:import_id>/ai/<action>")
def ai_control(import_id: int, action: str):
    _load_import(import_id)
    handlers = {"pause": jobs.request_pause, "resume": jobs.request_resume, "cancel": jobs.request_cancel}
    if action not in handlers:
        abort(404)
    job = _ai_job(import_id)
    ok = bool(job) and handlers[action](job["id"])
    if request.is_json:
        return jsonify({"ok": ok})
    if not ok:
        flash("AI整形のジョブを操作できませんでした（すでに終わっている可能性があります）", "error")
    return redirect(url_for("tables.ai", import_id=import_id))


# ---- 6 内容とファイルの確認 ---------------------------------------------------------------------------

@bp.get("/imports/<int:import_id>/preview")
def preview(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    spec = _spec_for(imp)
    not_ready = _not_ready(imp, spec)
    if not_ready:
        return not_ready
    stats = imp.get("stats") or {}
    issues = pipeline.load_issues(import_id)
    files = pipeline.preview_files(import_id, imp, spec)
    page = max(1, _int(request.args.get("page"), 1) or 1)
    all_rows = pipeline.load_rows(import_id)
    total_pages = max(1, (len(all_rows) + DATA_PAGE - 1) // DATA_PAGE)
    page = min(page, total_pages)
    data_rows = all_rows[(page - 1) * DATA_PAGE: page * DATA_PAGE]
    del all_rows
    return render_template(
        "tables/preview.html", imp=imp, spec=spec, stats=stats, issues=issues[:ISSUES_SHOWN], issue_total=len(issues),
        counts=count_levels(issues), blocking=has_blocking(issues), files=files, data_rows=data_rows, page=page,
        total_pages=total_pages, columns=[(c.key, c.display) for c in spec.columns],
        confirmed=imp["status"] == "confirmed", **_steps_ctx(6))


@bp.get("/imports/<int:import_id>/preview/file")
def preview_file(import_id: int):
    _load_import(import_id)
    text = pipeline.md_text(pipeline.import_files(import_id)["preview"], request.args.get("name", ""))
    if text is None:
        return _json_error("ファイルが見つかりません", 404)
    return jsonify({"name": request.args.get("name"), "text": text})


@bp.get("/imports/<int:import_id>/issues.csv")
def issues_csv(import_id: int):
    imp = _load_import(import_id)
    data = outputs.issues_csv(pipeline.load_issues(import_id))
    return send_file(io.BytesIO(data), mimetype="text/csv", as_attachment=True,
                     download_name=f"{Path(imp['file_name']).stem}_問題一覧.csv")


@bp.post("/imports/<int:import_id>/confirm")
def confirm(import_id: int):
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        flash("この取り込みはまだ確定できる状態ではありません", "error")
        return redirect(url_for("tables.preview", import_id=import_id))
    if has_blocking(pipeline.load_issues(import_id)):
        flash("エラーが残っているため確定できません。問題一覧を確認して、範囲や列の対応づけを直してください", "error")
        return redirect(url_for("tables.preview", import_id=import_id))
    pipeline.start_render_job(import_id)
    return redirect(url_for("tables.done", import_id=import_id))


# ---- 7 完了（ダウンロード） --------------------------------------------------------------------------

def _download_base(imp: dict, spec) -> str:
    stem = safe_filename_part(spec.file_prefix if spec is not None else Path(imp["file_name"]).stem)
    return f"{stem}_{datetime.now().strftime('%Y-%m-%d')}"


@bp.get("/imports/<int:import_id>/done")
def done(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_page(imp)
    if busy:
        return busy
    if imp["status"] != "confirmed":
        return redirect(url_for("tables.preview", import_id=import_id))
    spec = _spec_for(imp)
    files = [{"name": p.name, "size": p.stat().st_size} for p in pipeline.md_paths(import_id)]
    return render_template("tables/done.html", imp=imp, spec=spec, files=files, stats=imp.get("stats") or {},
                           **_steps_ctx(7))


@bp.get("/imports/<int:import_id>/download.zip")
def download_zip(import_id: int):
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if imp["status"] != "confirmed" or spec is None or not pipeline.md_paths(import_id):
        flash("先に [確定してMarkdownを作成] を押してください", "error")
        return redirect(url_for("tables.preview", import_id=import_id))
    data = pipeline.build_download(import_id, imp, spec)
    return send_file(io.BytesIO(data), mimetype="application/zip", as_attachment=True, download_name=f"{_download_base(imp, spec)}.zip")


@bp.get("/imports/<int:import_id>/normalized.csv")
def download_csv(import_id: int):
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        abort(404)
    data = outputs.normalized_csv(spec, pipeline.load_rows(import_id))
    return send_file(io.BytesIO(data), mimetype="text/csv", as_attachment=True,
                     download_name=f"{_download_base(imp, spec)}_正規化データ.csv")


# ---- 取り込み設定の編集（一覧は設定画面） ------------------------------------------------------------------------

@bp.get("/templates/<int:template_id>")
def template_edit(template_id: int):
    template = store.get_template(template_id)
    if template is None or template.get("spec") is None:
        abort(404)
    spec = template["spec"]
    rows = [_column_row(i, (col.headers or [col.display])[0], col, None, True) for i, col in enumerate(spec.columns)]
    rows += [_column_row(len(rows) + i, h, None, None, False) for i, h in enumerate((spec.header or {}).get("ignored") or [])]
    return render_template("tables/template_edit.html", template=template, spec=spec, rows=rows,
                           settings=_settings_of(spec, spec.name),
                           save_url=url_for("tables.save_template", template_id=template_id), **_options_ctx())


@bp.post("/templates/<int:template_id>")
def save_template(template_id: int):
    template = store.get_template(template_id)
    if template is None or template.get("spec") is None:
        return _json_error("取り込み設定が見つかりません", 404)
    spec, errors = _build_spec(_payload(), template["spec"])
    if errors:
        return _json_error(errors[0], errors=errors)
    try:
        store.save_template_version(template_id, spec)
    except sqlite3.IntegrityError:
        database.get_db().rollback()
        return _json_error(f"「{spec.name}」という名前の取り込み設定がすでにあります")
    flash("取り込み設定を保存しました。次の取り込みから使われます", "success")
    return jsonify({"ok": True, "redirect": url_for("settings.table_templates")})


@bp.post("/templates/<int:template_id>/delete")
def delete_template(template_id: int):
    if store.get_template(template_id) is None:
        abort(404)
    store.delete_template(template_id)
    flash("取り込み設定を削除しました。作成済みの取り込み履歴は残ります", "success")
    return redirect(url_for("settings.table_templates"))


# ---- ジョブの進捗（JSON） ------------------------------------------------------------------------------

def api_job(job_id: int):
    job = jobs.get_job(job_id)
    if job is None:
        return _json_error("ジョブが見つかりません", 404)
    return jsonify({k: job.get(k) for k in ("id", "kind", "status", "status_label", "progress", "message", "result",
                                            "finished")})


@bp.record_once
def _register_api(setup_state) -> None:
    # /api/jobs は /tables の外に置く（design.md 2.4）
    setup_state.app.add_url_rule("/api/jobs/<int:job_id>", "tables.api_job", api_job)

