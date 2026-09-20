"""帳票取り込み（1ファイル＝1件）。画面は /forms の1枚だけ。

上から順に「ファイルを置く」「帳票の種類とシート」「読み取り結果」「確定してダウンロード」の
欄が現れる。画面の移動はなく、どの操作も fetch でこのファイルのルートを呼び、返ってきた
HTML の断片をその場に入れ替える（views/forms.py のルートは JSON か HTML の断片を返す）。

Markdown は常にデータから作る（読み取り結果のプレビュー＝作業中の値、ダウンロード＝確定済みの値）。
複数ファイルをまとめて置くと「取り込みのまとまり（batch）」になり、同じ画面で1件ずつ読み取って
最後に zip でまとめて渡す。
ダウンロードしたデータはその場で消す（design.md 3.3）。消すのは Markdown を作り終え、本文を送り終えたあとだけ
（途中で切れたときは消さない: core.purge.purge_after_send）。
"""
from __future__ import annotations

import hashlib
import io
import json
import zipfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from flask import (Blueprint, abort, current_app, jsonify, redirect, render_template, request, send_file,
                   url_for)

from core import purge
from core.files import FORM_MAX_MERGED_CELLS, UploadError, original_name, precheck_excel, remove_upload, save_upload, upload_path
from excel.extractor import apply_manual_values, extract_document, is_blank_value, refresh_summary
from excel.tables import clean_table_value, is_table_value, parse_table_text, table_text_lines
from excel.text import EXCEL_ERROR_WARNING
from excel.workbook import WorkbookInfo, load_workbook_info
from export.formats import build_markdown, markdown_filename
from models import database as db
from pattern.matcher import rank_patterns, table_like_sheets
from views import set_download_name

bp = Blueprint("forms", __name__, url_prefix="/forms")

# 左のシートプレビューに出す範囲の上限（大きいシートで画面が重くならないように）
GRID_MAX_ROWS = 300
GRID_MAX_COLS = 60
CONFIRMED_STATES = ("confirmed", "modified")
MAX_BATCH_FILES = 50
# ダウンロードでデータが消えることの案内（画面の文言・確認ダイアログで使う）
DELETE_ON_DOWNLOAD_NOTE = ("ダウンロードすると、この帳票の元のファイルと読み取り結果はこのPCから消えます。"
                           "同じものをもう一度ダウンロードすることはできません。")
DELETE_ON_DOWNLOAD_CONFIRM = "ダウンロードすると、この帳票のデータはこのPCから消えます。もう一度ダウンロードすることはできません。"
BATCH_DELETE_CONFIRM = ("ダウンロードすると、このまとまりの帳票のデータはこのPCからすべて消えます。"
                        "もう一度ダウンロードすることはできません。")
LOST_WORK_MESSAGE = "読み取り直すと、手で修正した値は失われます"
# 別のタブ・古い画面から保存・確定されたときの案内
STALE_MESSAGE = "別の画面で内容が変わりました。画面を読み込み直してください"


# ---- 共通 -------------------------------------------------------------------------

def _get_document(doc_id: int) -> dict:
    doc = db.get_document(doc_id)
    if doc is None:
        abort(404)
    return doc


def _load_info(doc: dict) -> WorkbookInfo | None:
    """元のファイルを読む。消えている・壊れている場合は None。"""
    try:
        return load_workbook_info(upload_path(doc["stored_path"]))
    except Exception:
        return None


def _data(doc: dict, column: str = "data_json") -> dict | None:
    raw = doc.get(column)
    if not raw:
        return None
    try:
        return json.loads(raw)
    except ValueError:
        return None


def _dumps(extraction: dict) -> str:
    # 状態（修正中かどうか）は文字列の比較で決まるので、保存は常にこの形にそろえる
    return json.dumps(extraction, ensure_ascii=False)


def _search_title(doc: dict, extraction: dict) -> str:
    """一覧に出す見出し（Markdown の1行目）。"""
    first = build_markdown(doc, extraction).split("\n", 1)[0]
    return first[2:].strip() if first.startswith("# ") else first.strip()


def _display_text(value) -> str:
    if is_table_value(value):
        return table_json(value)
    return "" if value is None else str(value)


def table_json(value) -> str:
    """明細表の値を画面の入力欄（hidden）に入れる JSON。"""
    return json.dumps(value, ensure_ascii=False) if is_table_value(value) else ""


@bp.app_template_filter("field_text")
def field_text(value) -> str:
    """一覧・読み取りテスト用の表示。明細表は1行1明細の「品番: X／品名: Y」。"""
    if is_table_value(value):
        return "\n".join(table_text_lines(value))
    return "" if value is None else str(value)


@bp.app_template_filter("table_json")
def _table_json_filter(value) -> str:
    return table_json(value)


def _same_text(a: str, b: str) -> bool:
    norm = lambda s: s.replace("\r\n", "\n").replace("\r", "\n").strip()  # noqa: E731
    return norm(a) == norm(b)


def _apply_values(extraction: dict, values: dict, confirmed: dict | None = None) -> bool:
    """画面の入力値を反映する。表示中の値と同じ文字列の項目は触らない（勝手に「手で修正」にしない）。

    confirmed（確定済みの版）を渡すと、確定済みと同じ値に戻した項目は確定済みの項目そのものに戻す。
    戻り値: 変わった項目があれば True。
    """
    before = _dumps(extraction)
    fields = {f["field_name"]: f for f in extraction["fields"]}
    changed = {}
    for name, text in values.items():
        f = fields.get(name)
        if f is None or text is None:
            continue
        text = text if isinstance(text, str) else json.dumps(text, ensure_ascii=False)
        if f["data_type"] == "table":
            parsed, warning = parse_table_text(text)
            # 行の無い表（列見出しだけ）と空の値は同じもの
            if warning is None and (parsed == clean_table_value(f["value"])
                                    or (is_blank_value(parsed) and is_blank_value(f["value"]))):
                continue
        elif _same_text(text, _display_text(f["value"])):
            continue
        changed[f"value-{name}"] = text
    if changed:
        apply_manual_values(extraction, changed)
        _restore_confirmed_fields(extraction, confirmed)
    else:
        refresh_summary(extraction)
    return _dumps(extraction) != before


# 手で直すと変わる項目のキー。これ以外（表示名・Markdownへの出し方など）が違う項目は戻さない
_EDIT_KEYS = {"value", "unit", "warning", "edited", "ai_filled"}
_LOCATION_KEYS = {"sheet", "value_cell"}


def _restore_confirmed_fields(extraction: dict, confirmed: dict | None) -> None:
    """確定済みと同じ値（数値は単位も）になった項目を、確定済みの項目の中身に戻す。"""
    new_pattern, old_pattern = extraction.get("pattern", {}), (confirmed or {}).get("pattern", {})
    if (not confirmed or old_pattern.get("id") != new_pattern.get("id")
            or old_pattern.get("version_no") != new_pattern.get("version_no")
            or confirmed.get("sheets") != extraction.get("sheets")):
        return  # 別の種類・版・シートで読み直した版は比べない
    before = {f["field_name"]: f for f in confirmed.get("fields", [])}
    for i, f in enumerate(extraction["fields"]):
        old = before.get(f["field_name"])
        same_value = old is not None and (old.get("value") == f.get("value") or (
            f.get("data_type") == "table" and is_blank_value(old.get("value")) and is_blank_value(f.get("value"))))
        if (old is not None and old != f and same_value
                and (old.get("unit") or "") == (f.get("unit") or "")
                and {k: v for k, v in old.items() if k not in _EDIT_KEYS | _LOCATION_KEYS}
                == {k: v for k, v in f.items() if k not in _EDIT_KEYS | _LOCATION_KEYS}):
            extraction["fields"][i] = json.loads(json.dumps(old))
    refresh_summary(extraction)


def _is_blank(value) -> bool:
    return is_blank_value(value)


_UNIT_WARNING_PREFIXES = ("この項目の単位は", "単位が書かれていません")


def _field_status(f: dict) -> dict:
    """項目の状態タグ: 値の出どころ（auto/manual/blank）と要確認かどうか。"""
    blank = _is_blank(f["value"])
    if f.get("ai_filled"):
        source = "ai"
    elif f.get("edited"):
        source = "manual"
    elif blank:
        source = "blank"
    else:
        source = "auto"
    # 単位の食い違い・単位なしの警告は、手で直した値でも Markdown に誤った単位で出るので要確認のまま
    unit_issue = f["data_type"] == "number" and str(f.get("warning") or "").startswith(_UNIT_WARNING_PREFIXES)
    date_issue = f["data_type"] == "date" and bool(f.get("warning"))
    value = f["value"]
    number_issue = f["data_type"] == "number" and (
        "数値の部分だけ" in str(f.get("warning") or "")
        or not isinstance(value, (int, float)) or isinstance(value, bool))
    error_value = source == "blank" and str(f.get("warning") or "").startswith(EXCEL_ERROR_WARNING)
    issue = (f["required"] and blank) or error_value or (
        not blank and bool(f.get("warning")) and (source == "auto" or unit_issue or date_issue or number_issue))
    return {"source": source, "issue": bool(issue), "blank": blank,
            "missing_required": bool(f["required"] and blank)}


def _summary(doc: dict, extraction: dict) -> dict:
    """読み取り結果の欄の数値（チップ・Markdownプレビュー・項目ごとの状態）。"""
    statuses = {f["field_name"]: _field_status(f) for f in extraction["fields"]}
    return {
        "markdown": build_markdown(doc, extraction),
        "file_name": markdown_filename(doc, extraction),
        "state": doc["state"],
        "counts": {
            "issue": sum(1 for s in statuses.values() if s["issue"]),
            "manual": sum(1 for s in statuses.values() if s["source"] == "manual"),
            "blank": sum(1 for s in statuses.values() if s["blank"]),
        },
        "missing_required": list(extraction.get("missing_required") or []),
        "fields": {name: {**s, "warning": f.get("warning") or ""}
                   for name, s, f in ((f["field_name"], statuses[f["field_name"]], f) for f in extraction["fields"])},
    }


def _sheet_grids(info: WorkbookInfo | None, sheet_names: list[str]) -> list[dict]:
    """元のシートを HTML の表にするための行データ（結合セルは rowspan/colspan）。"""
    if info is None:
        return []
    from openpyxl.utils import get_column_letter

    sheets = []
    for name in sheet_names:
        grid = info.grids.get(name)
        if grid is None:
            continue
        max_row, max_col = min(grid.max_row, GRID_MAX_ROWS), min(grid.max_col, GRID_MAX_COLS)
        rows = []
        for r in range(1, max_row + 1):
            cells = []
            for c in range(1, max_col + 1):
                top, left, bottom, right = grid.bounds(r, c)
                if (top, left) != (r, c):
                    continue  # 結合範囲の左上以外は描かない
                cell = grid.cells.get((r, c))
                cells.append({
                    "coord": f"{get_column_letter(c)}{r}",
                    "text": cell.text if cell else "",
                    "rowspan": min(bottom, max_row) - r + 1,
                    "colspan": min(right, max_col) - c + 1,
                    "label": bool(cell and (cell.bold or cell.filled)),
                })
            rows.append({"index": r, "cells": cells})
        sheets.append({
            "name": name,
            "rows": rows,
            "letters": [get_column_letter(c) for c in range(1, max_col + 1)],
            "truncated": grid.max_row > GRID_MAX_ROWS or grid.max_col > GRID_MAX_COLS,
            "images": len(info.images_in([name])),
        })
    return sheets


def _version(doc: dict) -> str:
    """読み取り結果を出したときの作業データの版（古い画面からの保存・確定を見分ける）。"""
    return hashlib.sha1(str(doc.get("data_json") or "").encode("utf-8")).hexdigest()[:16]


def _payload() -> dict:
    payload = request.get_json(force=True, silent=True)
    return payload if isinstance(payload, dict) else {}


def _is_stale(doc: dict) -> bool:
    """画面が送ってきた版が今の作業データと違うか。版を送らない呼び出しは比べない。"""
    sent = _payload().get("version") or request.form.get("version")
    return bool(sent) and str(sent) != _version(doc)


# ---- 1枚の画面 ----------------------------------------------------------------------

@bp.get("/", endpoint="index")
@bp.get("/new", endpoint="new")
def page():
    """帳票取り込みの1枚の画面。ここから先はすべて fetch で欄が増えていく。"""
    return render_template("forms/page.html",
                           active_pattern_count=db.count_active_patterns(),
                           pattern_count=len(db.list_patterns()),
                           max_files=MAX_BATCH_FILES,
                           max_mb=(current_app.config.get("MAX_CONTENT_LENGTH") or 0) // (1024 * 1024))


# ---- ファイルを置く ------------------------------------------------------------------

def upload_error_text(storage, exc: Exception, position: int | None = None) -> str:
    """取り込めなかったファイルのメッセージ。まとめて置いたときは「3件目のファイル」と位置で示す。"""
    message = str(exc)
    name = original_name(storage)
    if name and message.startswith(f"{name}: "):
        message = message[len(name) + 2:]
    return f"{position}件目のファイル: {message}" if position else message


def _store_document(storage, batch_id: str = "", order: int = 0) -> int:
    """1ファイルを保存して帳票を作る。読めないファイルは保存先から消して UploadError。"""
    cfg = current_app.config
    stored = save_upload(storage, "documents", cfg["ALLOWED_EXTENSIONS"], cfg["MAX_CONTENT_LENGTH"])
    try:
        path = upload_path(stored.stored_path)
        precheck_excel(path, cfg.get("EXCEL_MAX_CELLS"), max_merged=FORM_MAX_MERGED_CELLS)
        try:
            info = load_workbook_info(path)
        except Exception as exc:
            current_app.logger.warning("帳票を読み込めませんでした: %s", exc.__class__.__name__)
            raise UploadError("Excelファイルとして読み込めませんでした") from exc
        if not info.grids:
            raise UploadError("シートがないブックです。シートのあるブックを選んでください")
    except Exception:
        remove_upload(stored.stored_path)   # 思わぬエラーでもアップロードしたファイルを残さない（design.md 3.3）
        raise
    return db.create_document(stored.file_name, stored.file_hash, stored.stored_path,
                              batch_id=batch_id, batch_order=order)


@bp.post("/upload")
def upload():
    """置かれたファイルを取り込む（fetch）。複数なら1つのまとまり（batch）にして、同じ画面で1件ずつ読む。"""
    storages = [s for s in request.files.getlist("file") if s is not None and s.filename]
    if not storages:
        return jsonify(error="ファイルを置いてください"), 400
    if len(storages) > MAX_BATCH_FILES:
        return jsonify(error=f"一度に置けるのは{MAX_BATCH_FILES}ファイルまでです（置かれたのは{len(storages)}ファイル）"), 400
    batch_id = uuid4().hex if len(storages) > 1 else ""
    docs, errors = [], []
    for order, storage in enumerate(storages):
        try:
            doc_id = _store_document(storage, batch_id=batch_id, order=order)
        except UploadError as exc:
            errors.append(upload_error_text(storage, exc, order + 1 if len(storages) > 1 else None))
            continue
        doc = db.get_document(doc_id)
        docs.append({"id": doc_id, "file_name": doc["file_name"] if doc else ""})
    if not docs:
        return jsonify(error=errors[0] if errors else "取り込めるファイルがありませんでした", errors=errors), 400
    return jsonify(docs=docs, batch_id=batch_id, errors=errors)


# ---- 2 帳票の種類とシート ---------------------------------------------------------------

@bp.get("/<int:doc_id>/type")
def type_fragment(doc_id: int):
    """帳票の種類とシートの欄（HTML の断片）。"""
    doc = _get_document(doc_id)
    info = _load_info(doc)
    if info is None:
        return jsonify(error="元のファイルを読み込めませんでした。ファイルを置き直してください"), 409
    matches = rank_patterns(info, db.load_active_patterns())
    suggested = {str(m.pattern.id): m.sheet_names for m in matches}
    current = _data(doc)
    selected_id = doc["pattern_id"] if any(m.pattern.id == doc["pattern_id"] for m in matches) else None
    selected_id = selected_id or (matches[0].pattern.id if matches else None)
    selected_sheets = (current or {}).get("sheets") if current and selected_id == doc["pattern_id"] else None
    if not selected_sheets:
        selected_sheets = suggested.get(str(selected_id)) or []
    sheets = [
        {"name": name, "cells": len(grid.cells), "images": len(info.images_in([name])), "hidden": grid.hidden}
        for name, grid in info.grids.items()
    ]
    html = render_template(
        "forms/_type.html",
        doc=doc,
        matches=matches,
        sheets=sheets,
        suggested=suggested,
        selected_id=selected_id,
        selected_sheets=selected_sheets,
        table_sheets=table_like_sheets(info),
        duplicate=db.find_confirmed_by_hash(doc["file_hash"], exclude_id=doc["id"]),
        lost_work=doc["state"] in CONFIRMED_STATES,
        form_types_url=url_for("form_types.index"),
    )
    return jsonify(html=html, file_name=doc["file_name"], has_types=bool(matches))


@bp.post("/<int:doc_id>/read")
def read(doc_id: int):
    """選ばれた種類とシートで読み取り、読み取り結果の欄（HTML の断片）を返す。"""
    doc = _get_document(doc_id)
    pattern = db.load_pattern(request.form.get("pattern_id", type=int))
    sheets = request.form.getlist("sheets")
    if pattern is None or not sheets:
        return jsonify(error="帳票の種類と読み取るシートを選んでください"), 400
    if doc["state"] in CONFIRMED_STATES and request.form.get("acknowledge") != "on":
        return jsonify(error=f"{LOST_WORK_MESSAGE}。確認のチェックを入れてから読み取り直してください"), 400
    info = _load_info(doc)
    if info is None:
        return jsonify(error="元のファイルを読み込めませんでした"), 409
    sheets = [s for s in sheets if s in info.grids]
    if not sheets:
        return jsonify(error="選んだシートがファイルにありません"), 400

    extraction = extract_document(info, pattern, sheets)
    db.reset_document(doc_id, pattern.id, _dumps(extraction))
    if doc["state"] not in CONFIRMED_STATES:
        db.update_document(doc_id, title=_search_title(doc, extraction))
    return _review_response(doc_id)


def _review_response(doc_id: int):
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    info = _load_info(doc)
    summary = _summary(doc, extraction)
    html = render_template(
        "forms/_review.html",
        doc=doc,
        extraction=extraction,
        summary=summary,
        grids=_sheet_grids(info, extraction["sheets"]),
        info_missing=info is None,
        version=_version(doc),
    )
    return jsonify(html=html, version=_version(doc), summary=summary, doc_id=doc_id)


@bp.get("/<int:doc_id>/review")
def review_fragment(doc_id: int):
    """読み取り済みの帳票の読み取り結果の欄（画面を開き直したとき用）。"""
    return _review_response(doc_id)


# ---- 3 読み取り結果（その場で直す・途中保存） -----------------------------------------------

# 途中保存をした画面の目印（doc_id → (保存後の版, 画面の目印)）。
_DRAFT_TOKENS: dict[int, tuple[str, str]] = {}


@purge.on_documents_purged
def _forget_draft_tokens(doc_ids) -> None:
    for doc_id in doc_ids:
        _DRAFT_TOKENS.pop(int(doc_id), None)


def _page_token() -> str:
    token = _payload().get("page_token")
    return str(token)[:64] if token else ""


def _saved_by_same_page(doc_id: int, doc: dict) -> bool:
    token = _page_token()
    saved = _DRAFT_TOKENS.get(doc_id)
    return bool(token) and saved is not None and saved == (_version(doc), token)


def _json_values() -> dict:
    payload = _payload()
    if payload:
        values = payload.get("values", payload)
        return {str(k): v for k, v in values.items()} if isinstance(values, dict) else {}
    return {key[len("value-"):]: request.form[key] for key in request.form.keys() if key.startswith("value-")}


@bp.post("/<int:doc_id>/draft")
def draft(doc_id: int):
    """入力内容の途中保存（fetch）。変わった項目だけ保存し、204 を返す。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    if _is_stale(doc) and not _saved_by_same_page(doc_id, doc):
        return jsonify(error=STALE_MESSAGE), 409
    version = _version(doc)
    if _apply_values(extraction, _json_values(), _data(doc, "confirmed_json")):
        title = None if doc["state"] in CONFIRMED_STATES else _search_title(doc, extraction)
        new_json = _dumps(extraction)
        db.save_draft(doc_id, new_json, title=title)
        version = _version({"data_json": new_json})
        token = _page_token()
        if token:
            _DRAFT_TOKENS[doc_id] = (version, token)
    return "", 204, {"X-Doc-Version": version}


@bp.post("/<int:doc_id>/preview")
def preview(doc_id: int):
    """入力中の値で作った Markdown と状態（保存はしない）。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    if _is_stale(doc):
        return jsonify(error=STALE_MESSAGE), 409
    _apply_values(extraction, _json_values(), _data(doc, "confirmed_json"))
    return jsonify(_summary(doc, extraction))


# ---- 4 確定してダウンロード ---------------------------------------------------------------

@bp.post("/<int:doc_id>/confirm")
def confirm(doc_id: int):
    """読み取り結果を確定する（fetch）。値は途中保存で入っているので、ここでは版だけ確かめる。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    if _is_stale(doc):
        return jsonify(error=STALE_MESSAGE), 409
    payload = _payload()
    values = _json_values()
    if values and _apply_values(extraction, values, _data(doc, "confirmed_json")):
        db.save_draft(doc_id, _dumps(extraction))
    missing = extraction.get("missing_required") or []
    if missing and not payload.get("allow_missing"):
        return jsonify(error=f"必須の項目が空欄です（{'、'.join(missing)}）。"
                             "値を入れるか、「空欄のまま確定する」にチェックを入れてください",
                       missing_required=missing), 409
    db.confirm_document(doc_id, title=_search_title(doc, extraction))
    return jsonify(ok=True, doc_id=doc_id)


def _docs_of(ids: list[int]) -> list[dict]:
    docs = [d for d in (db.get_document(i) for i in ids) if d is not None]
    return docs


@bp.get("/finish", endpoint="finish")
def finish_fragment():
    """確定してダウンロードの欄（HTML の断片）。?ids=1,2,3 は同じ画面で扱っている帳票。"""
    raw = request.args.get("ids", "")
    ids = [int(part) for part in raw.split(",") if part.strip().isdigit()]
    docs = _docs_of(ids)
    if not docs:
        return jsonify(html="", ready=False, confirmed=0, total=0)
    current_id = request.args.get("current", type=int)
    current = next((d for d in docs if d["id"] == current_id), None) or docs[0]
    confirmed = [d for d in docs if d["state"] in CONFIRMED_STATES]
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    batch_id = current.get("batch_id") or ""
    working = _data(current)
    extraction = _data(current, "confirmed_json") or working
    missing = list((working or {}).get("missing_required") or [])
    html = render_template(
        "forms/_finish.html",
        docs=docs,
        current=current,
        confirmed=confirmed,
        pending=pending,
        batch_id=batch_id if len(docs) > 1 else "",
        missing=missing,
        read_yet=working is not None,
        confirmed_states=CONFIRMED_STATES,
        file_name=markdown_filename(current, extraction) if extraction else "",
        markdown=build_markdown(current, _data(current, "confirmed_json")) if current["state"] in CONFIRMED_STATES else "",
        delete_note=DELETE_ON_DOWNLOAD_NOTE,
        delete_confirm=DELETE_ON_DOWNLOAD_CONFIRM,
        batch_confirm=_batch_zip_confirm(docs),
    )
    return jsonify(html=html, ready=bool(confirmed), confirmed=len(confirmed), total=len(docs),
                   read_yet=working is not None,
                   next_id=(pending[0]["id"] if pending else None))


def _batch_zip_confirm(docs: list[dict]) -> str:
    """まとまりの zip ダウンロードの確認文。"""
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    confirmed = len(docs) - len(pending)
    if pending:
        return (f"確定済みの{confirmed}件だけを zip でダウンロードします。その{confirmed}件のデータはこのPCから消えます"
                f"（未確定の{len(pending)}件は残ります）。")
    return BATCH_DELETE_CONFIRM


# ---- ダウンロード -------------------------------------------------------------------

@bp.get("/<int:doc_id>/download.md")
def download_md(doc_id: int):
    """Markdown を渡し、渡し終えた帳票のデータを消す（design.md 3.3）。"""
    name, body = _single_markdown(doc_id)
    response = send_file(io.BytesIO(body), mimetype="text/markdown", as_attachment=True, download_name=name,
                         conditional=False)   # Range でも全体を返す
    set_download_name(response, name, "form")
    return purge.purge_after_send(response, purge.purge_documents, [doc_id])


def _single_markdown(doc_id: int) -> tuple[str, bytes]:
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        abort(404)
    return markdown_filename(doc, confirmed), build_markdown(doc, confirmed).encode("utf-8")


def _unique_name(name: str, used: set[str]) -> str:
    """まとまりの中で重ならない名前（大文字・小文字だけの違いも重なりとみなす）。"""
    stem, suffix = Path(name).stem, Path(name).suffix or ".md"
    candidate, n = name, 2
    while candidate.casefold() in used:
        candidate = f"{stem}_{n}{suffix}"
        n += 1
    used.add(candidate.casefold())
    return candidate


def _batch_markdown_files(confirmed_ids: list[int]) -> list[tuple[str, bytes]]:
    files, used = [], set()
    for doc in db.list_confirmed_documents(confirmed_ids):
        try:
            extraction = json.loads(doc["confirmed_json"])
        except (TypeError, ValueError):
            continue
        files.append((_unique_name(markdown_filename(doc, extraction), used),
                      build_markdown(doc, extraction).encode("utf-8")))
    return files


@bp.get("/batches/<batch_id>/download.zip")
def download_batch(batch_id: str):
    """まとめ取り込み1回分の Markdown を zip で渡し、渡し終えた分のデータを消す。

    確定済みの帳票だけを zip にして、その分だけ消す（未確定の帳票はこのPCに残る）。
    """
    docs = db.list_batch_documents(batch_id)
    if not docs:
        abort(404)
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    confirmed_ids = [d["id"] for d in docs if d["state"] in CONFIRMED_STATES]
    if not confirmed_ids:
        abort(404)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, body in _batch_markdown_files(confirmed_ids):
            zf.writestr(name, body)
    buffer.seek(0)
    zip_name = f"帳票Markdown_{datetime.now():%Y%m%d_%H%M%S}.zip"
    response = send_file(buffer, mimetype="application/zip", as_attachment=True, download_name=zip_name,
                         conditional=False)   # Range でも全体を返す
    set_download_name(response, zip_name, "forms_markdown")
    if pending:
        return purge.purge_after_send(response, purge.purge_documents, confirmed_ids)
    return purge.purge_after_send(response, purge.purge_batch, batch_id)


@bp.get("/<int:doc_id>/original")
def original(doc_id: int):
    doc = _get_document(doc_id)
    try:
        path = upload_path(doc["stored_path"])
    except UploadError:
        abort(404)
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=doc["file_name"])


@bp.post("/<int:doc_id>/delete")
def delete(doc_id: int):
    """この帳票の取り込みをやめる（元のファイルと読み取り結果を消す）。"""
    _get_document(doc_id)
    purge.purge_documents([doc_id])
    incomplete = purge.purge_incomplete()
    return jsonify(ok=True, incomplete=incomplete)


# 画面を作り直す前の URL（お気に入り・古いリンク）は1枚の画面へ送る
@bp.get("/<int:doc_id>")
@bp.get("/<int:doc_id>/done")
def legacy(doc_id: int):
    return redirect(url_for(".index"))
