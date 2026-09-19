"""帳票を取り込む（1ファイル＝1件）: ファイル選択 → 種類とシート → 読み取り結果の確認・修正 → 完了。

Markdown は常にデータから作る（確認画面のプレビュー＝作業中の値、ダウンロード＝確定済みの値）。
複数ファイルをまとめて選ぶと「取り込みのまとまり（batch）」になり、1件ずつ確認したあと zip でまとめて渡す。
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

from flask import (Blueprint, abort, current_app, flash, jsonify, redirect, render_template, request, send_file,
                   url_for)

from core import purge
from core.files import UploadError, original_name, precheck_excel, remove_upload, save_upload, upload_path
from excel.extractor import apply_manual_values, extract_document, is_blank_value, refresh_summary
from excel.tables import clean_table_value, is_table_value, parse_table_text, table_text_lines
from excel.workbook import WorkbookInfo, load_workbook_info
from export.formats import build_json, build_markdown, markdown_filename
from models import database as db
from pattern.matcher import rank_patterns, table_like_sheets
from pattern.model import DATA_TYPES
from services import ai_assist, llm
from views import form_link, safe_next, set_download_name

bp = Blueprint("forms", __name__, url_prefix="/forms")

STEPS = ["ファイルを選ぶ", "帳票の種類とシートを確認", "読み取り結果の確認・修正", "完了（Markdownをダウンロード）"]
# 左のシートプレビューに出す範囲の上限（大きいシートで画面が重くならないように）
GRID_MAX_ROWS = 300
GRID_MAX_COLS = 60
CONFIRMED_STATES = ("confirmed", "modified")
LOST_WORK_MESSAGE = "読み取り直すと、手で修正した値とAIが入力した値は失われます"
# ダウンロードでデータが消えることの案内（画面の文言・確認ダイアログで使う）
DELETE_ON_DOWNLOAD_NOTE = ("ダウンロードすると、この帳票の元のファイルと読み取り結果はこのPCから消えます。"
                           "同じものをもう一度ダウンロードすることはできません。")
DELETE_ON_DOWNLOAD_CONFIRM = "ダウンロードすると、この帳票のデータはこのPCから消えます。もう一度ダウンロードすることはできません。"
BATCH_DELETE_CONFIRM = ("ダウンロードすると、このまとまりの帳票のデータはこのPCからすべて消えます。"
                        "もう一度ダウンロードすることはできません。")
MAX_BATCH_FILES = 50
# 別のタブ・戻るボタンで開いた古い確認画面から保存・確定されたときの案内
STALE_MESSAGE = "別の画面で内容が変わりました。読み込み直してください"


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
    """明細表の値を確認画面の入力欄（hidden）に入れる JSON。"""
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

    confirmed（確定済みの版）を渡すと、確定済みと同じ値に戻した項目は確定済みの項目そのものに戻す
    （「手で修正」の印や警告の違いだけで、いつまでも「修正中」のままにならないように）。
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
            if warning is None and parsed == clean_table_value(f["value"]):
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


# 手で直すと変わる項目のキー。これ以外（表示名・Markdownへの出し方など、帳票の種類の設定）が違う項目は戻さない
_EDIT_KEYS = {"value", "unit", "warning", "edited", "ai_filled"}


def _restore_confirmed_fields(extraction: dict, confirmed: dict | None) -> None:
    """確定済みと同じ値（数値は単位も）になった項目を、確定済みの項目の中身に戻す。"""
    new_pattern, old_pattern = extraction.get("pattern", {}), (confirmed or {}).get("pattern", {})
    if (not confirmed or old_pattern.get("id") != new_pattern.get("id")
            or old_pattern.get("version_no") != new_pattern.get("version_no")
            or confirmed.get("sheets") != extraction.get("sheets")):
        return  # 別の種類・版・シートで読み直した版は比べない（直した種類の設定を古い設定に戻さない）
    before = {f["field_name"]: f for f in confirmed.get("fields", [])}
    for i, f in enumerate(extraction["fields"]):
        old = before.get(f["field_name"])
        if (old is not None and old != f and old.get("value") == f.get("value")
                and (old.get("unit") or "") == (f.get("unit") or "")
                and {k: v for k, v in old.items() if k not in _EDIT_KEYS}
                == {k: v for k, v in f.items() if k not in _EDIT_KEYS}):
            extraction["fields"][i] = json.loads(json.dumps(old))
    refresh_summary(extraction)


def _form_values(form) -> dict:
    return {key[len("value-"):]: form[key] for key in form.keys() if key.startswith("value-")}


def _is_blank(value) -> bool:
    return is_blank_value(value)


_UNIT_WARNING_PREFIXES = ("この項目の単位は", "単位が書かれていません")


def _field_status(f: dict) -> dict:
    """項目の状態タグ: 値の出どころ（auto/ai/manual/blank）と要確認かどうか。"""
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
    # 日付の警告（年が無い・日付として読めない）は、手で直した値でも入力した文字から出し直しているので要確認のまま
    date_issue = f["data_type"] == "date" and bool(f.get("warning"))
    issue = (f["required"] and blank) or (not blank and bool(f.get("warning"))
                                          and (source == "auto" or unit_issue or date_issue))
    return {"source": source, "issue": bool(issue), "blank": blank,
            "missing_required": bool(f["required"] and blank)}


def _summary(doc: dict, extraction: dict) -> dict:
    """確認画面の数値（チップ・Markdownプレビュー・項目ごとの状態）。"""
    statuses = {f["field_name"]: _field_status(f) for f in extraction["fields"]}
    return {
        "markdown": build_markdown(doc, extraction),
        "file_name": markdown_filename(doc, extraction),
        # 途中保存で「確定済み」→「修正中」に変わるので、画面の状態タグも書き替えられるように返す
        "state": doc["state"],
        "counts": {
            "issue": sum(1 for s in statuses.values() if s["issue"]),
            "ai": sum(1 for s in statuses.values() if s["source"] == "ai"),
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


# ---- 1 ファイルを選ぶ ------------------------------------------------------------------

@bp.get("/")
def index():
    return redirect(url_for(".new"))


@bp.get("/new")
def new():
    return render_template("forms/new.html", steps=STEPS, active_pattern_count=db.count_active_patterns(),
                           pattern_count=len(db.list_patterns()), max_files=MAX_BATCH_FILES,
                           max_mb=(current_app.config.get("MAX_CONTENT_LENGTH") or 0) // (1024 * 1024))


def upload_error_text(storage, exc: Exception, position: int | None = None) -> str:
    """取り込めなかったファイルのメッセージ。ファイル名は出さず、まとめて選んだときは「3件目のファイル」と位置で示す。

    flash はセッションの Cookie に入るので、ファイル名を入れない（design.md 3.3）。
    """
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
        precheck_excel(path, cfg.get("EXCEL_MAX_CELLS"))
        try:
            info = load_workbook_info(path)
        except Exception as exc:
            current_app.logger.warning("帳票を読み込めませんでした: %s", exc.__class__.__name__)
            raise UploadError("Excelファイルとして読み込めませんでした") from exc
        if not info.grids:
            # シートの一覧（<sheets>）が無いブックは precheck_excel では見分けられない。読み取り先が無いので断る
            raise UploadError("シートがないブックです。シートのあるブックを選んでください")
    except Exception:
        remove_upload(stored.stored_path)   # 思わぬエラーでもアップロードしたファイルを残さない（design.md 3.3）
        raise
    return db.create_document(stored.file_name, stored.file_hash, stored.stored_path,
                              batch_id=batch_id, batch_order=order)


@bp.post("/upload")
def upload():
    """1ファイルならそのまま確認へ。複数ファイルなら1つのまとまり（batch）にして先頭から確認する。"""
    storages = [s for s in request.files.getlist("file") if s is not None and s.filename]
    if not storages:
        flash("ファイルを選んでください", "error")
        return redirect(url_for(".new"))
    if len(storages) == 1:
        try:
            doc_id = _store_document(storages[0])
        except UploadError as exc:
            flash(upload_error_text(storages[0], exc), "error")
            return redirect(url_for(".new"))
        return redirect(url_for(".type_select", doc_id=doc_id))

    if len(storages) > MAX_BATCH_FILES:
        flash(f"一度に選べるのは{MAX_BATCH_FILES}ファイルまでです（選んだのは{len(storages)}ファイル）", "error")
        return redirect(url_for(".new"))
    batch_id = uuid4().hex
    doc_ids, errors = [], []
    for order, storage in enumerate(storages):
        try:
            doc_ids.append(_store_document(storage, batch_id=batch_id, order=order))
        except UploadError as exc:
            # どのファイルを取り込めなかったかは選んだ順の位置で示す（ファイル名は flash に入れない）
            errors.append(upload_error_text(storage, exc, order + 1))
    for message in errors:
        flash(message, "error")
    if not doc_ids:
        return redirect(url_for(".new"))
    flash(f"{len(doc_ids)}件の帳票を取り込みます。1件ずつ確認したあと、まとめて zip でダウンロードします", "info")
    return redirect(url_for(".type_select", doc_id=doc_ids[0]))


# ---- 取り込みのまとまり（複数ファイル） -------------------------------------------------------

def _batch_info(doc: dict) -> dict | None:
    """まとめ取り込みの進み具合（3/12）と次の帳票。まとまりに属さない帳票なら None。"""
    batch_id = doc.get("batch_id") or ""
    if not batch_id:
        return None
    docs = db.list_batch_documents(batch_id)
    if not docs:
        return None
    for d in docs:
        d["href"] = form_link(d)
    ids = [d["id"] for d in docs]
    position = ids.index(doc["id"]) + 1 if doc["id"] in ids else 0
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    after = [d for d in docs if d["state"] not in CONFIRMED_STATES and ids.index(d["id"]) > position - 1]
    modified = [d for d in docs if d["state"] == "modified"]
    confirmed = len(docs) - len(pending)
    zip_confirm = batch_zip_confirm(docs)
    return {
        "id": batch_id,
        "docs": docs,
        "position": position,
        "total": len(docs),
        "confirmed": confirmed,
        "pending": pending,
        "modified": modified,
        "zip_confirm": zip_confirm,
        "all_confirmed": not pending,
        "next": (after or pending or [None])[0],
    }


def batch_zip_confirm(docs: list[dict]) -> str:
    """まとまりの zip ダウンロードの確認文（確認画面・完了画面・ホームで同じ文言にする）。

    修正中の帳票は zip に確定済みの版で入り、確定し直していない変更はダウンロードで消える。確認の文言で名前を出す。
    """
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    modified = [d for d in docs if d["state"] == "modified"]
    confirmed = len(docs) - len(pending)
    if pending:
        text = (f"確定済みの{confirmed}件だけを zip でダウンロードします。その{confirmed}件のデータはこのPCから消えます"
                f"（未確定の{len(pending)}件は残ります）。")
    else:
        text = BATCH_DELETE_CONFIRM
    return _modified_warning(modified) + text if modified else text


def _modified_warning(docs: list[dict]) -> str:
    names = "、".join(d.get("title") or d["file_name"] for d in docs)
    return (f"修正中の帳票が{len(docs)}件あります（{names}）。確定し直していない変更は zip に入らず、消えます。"
            "変更を残すときは、先に確定し直してください。")


def delete_confirm(doc: dict) -> str:
    """1件ダウンロードの確認文。修正中なら、確定し直していない変更が入らずに消えることを先に書く。"""
    if doc["state"] == "modified":
        return ("この帳票は修正中です。確定し直していない変更は Markdown に入らず、消えます。"
                + DELETE_ON_DOWNLOAD_CONFIRM)
    return DELETE_ON_DOWNLOAD_CONFIRM


# ---- 2 帳票の種類とシートを確認 -------------------------------------------------------------

@bp.get("/<int:doc_id>/type")
def type_select(doc_id: int):
    return _render_type(_get_document(doc_id))


@bp.post("/<int:doc_id>/ai-classify")
def ai_classify(doc_id: int):
    """ルールでの候補をAIに見直させる。結果は選択の初期値にするだけで、決めるのは人。"""
    doc = _get_document(doc_id)
    info = _load_info(doc)
    if info is None:
        flash("元のファイルを読み込めませんでした", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    matches = rank_patterns(info, db.load_active_patterns())
    if not matches:
        flash("使用中の帳票の種類がありません", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    try:
        result = ai_assist.classify_pattern(info, matches)
    except Exception as exc:
        flash(f"AIで種類を推定できませんでした: {llm.friendly_error(exc)}", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    return _render_type(doc, info=info, matches=matches, ai_result=result)


def _render_type(doc: dict, info: WorkbookInfo | None = None, matches=None, ai_result: dict | None = None):
    info = info or _load_info(doc)
    if info is None:
        flash("元のファイルを読み込めませんでした。ファイルを選び直してください", "error")
        return render_template("forms/type.html", steps=STEPS, doc=doc, info_missing=True, matches=[], sheets=[],
                               suggested={}, selected_id=None, selected_sheets=[], table_sheets=[],
                               duplicate=None, ai_result=None, pattern_names={}, llm_ready=llm.is_configured(),
                               lost_work=doc["state"] in CONFIRMED_STATES, batch=_batch_info(doc))
    matches = matches if matches is not None else rank_patterns(info, db.load_active_patterns())
    suggested = {str(m.pattern.id): m.sheet_names for m in matches}
    current = _data(doc)
    selected_id = doc["pattern_id"] if any(m.pattern.id == doc["pattern_id"] for m in matches) else None
    selected_id = selected_id or (matches[0].pattern.id if matches else None)
    selected_sheets = (current or {}).get("sheets") if current and selected_id == doc["pattern_id"] else None
    if ai_result and ai_result.get("pattern_id"):
        selected_id = ai_result["pattern_id"]
        if ai_result.get("sheets"):
            suggested[str(selected_id)] = ai_result["sheets"]
        selected_sheets = None
    if not selected_sheets:
        selected_sheets = suggested.get(str(selected_id)) or []
    sheets = [
        {"name": name, "cells": len(grid.cells), "images": len(info.images_in([name])), "hidden": grid.hidden}
        for name, grid in info.grids.items()
    ]
    return render_template(
        "forms/type.html",
        steps=STEPS,
        doc=doc,
        info_missing=False,
        matches=matches,
        sheets=sheets,
        suggested=suggested,
        selected_id=selected_id,
        selected_sheets=selected_sheets,
        table_sheets=table_like_sheets(info),
        duplicate=db.find_confirmed_by_hash(doc["file_hash"], exclude_id=doc["id"]),
        ai_result=ai_result,
        pattern_names={m.pattern.id: m.pattern.name for m in matches},
        llm_ready=llm.is_configured(),
        lost_work=doc["state"] in CONFIRMED_STATES,
        batch=_batch_info(doc),
    )


@bp.post("/<int:doc_id>/read")
def read(doc_id: int):
    return _read(_get_document(doc_id))


@bp.post("/<int:doc_id>/reread")
def reread(doc_id: int):
    return _read(_get_document(doc_id))


def _read(doc: dict):
    doc_id = doc["id"]
    pattern = db.load_pattern(request.form.get("pattern_id", type=int))
    sheets = request.form.getlist("sheets")
    if pattern is None or not sheets:
        flash("帳票の種類と読み取るシートを選んでください", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    if doc["state"] in CONFIRMED_STATES and request.form.get("acknowledge") != "on":
        flash(f"{LOST_WORK_MESSAGE}。確認のチェックを入れてから読み取り直してください", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    info = _load_info(doc)
    if info is None:
        flash("元のファイルを読み込めませんでした", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))
    sheets = [s for s in sheets if s in info.grids]
    if not sheets:
        flash("選んだシートがファイルにありません", "error")
        return redirect(url_for(".type_select", doc_id=doc_id))

    extraction = extract_document(info, pattern, sheets)
    db.reset_document(doc_id, pattern.id, _dumps(extraction))
    if doc["state"] not in CONFIRMED_STATES:
        db.update_document(doc_id, title=_search_title(doc, extraction))
    return redirect(url_for(".review", doc_id=doc_id))


# ---- 3 読み取り結果の確認・修正 ------------------------------------------------------------

@bp.get("/<int:doc_id>/review")
def review(doc_id: int):
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return redirect(url_for(".type_select", doc_id=doc_id))
    info = _load_info(doc)
    summary = _summary(doc, extraction)
    return render_template(
        "forms/review.html",
        steps=STEPS,
        doc=doc,
        extraction=extraction,
        summary=summary,
        grids=_sheet_grids(info, extraction["sheets"]),
        info_missing=info is None,
        data_types=DATA_TYPES,
        llm_ready=llm.is_configured(),
        batch=_batch_info(doc),
        version=_version(doc),
    )


def _version(doc: dict) -> str:
    """確認画面を開いたときの作業データの版（古い画面からの保存・確定を見分ける）。"""
    return hashlib.sha1(str(doc.get("data_json") or "").encode("utf-8")).hexdigest()[:16]


def _is_stale(doc: dict) -> bool:
    """画面が送ってきた版が今の作業データと違うか。版を送らない呼び出し（古い画面など）は比べない。"""
    payload = request.get_json(force=True, silent=True)
    sent = payload.get("version") if isinstance(payload, dict) else request.form.get("version")
    return bool(sent) and str(sent) != _version(doc)


def _json_values() -> dict:
    payload = request.get_json(force=True, silent=True)
    if isinstance(payload, dict):
        values = payload.get("values", payload)
        return {str(k): v for k, v in values.items()} if isinstance(values, dict) else {}
    return _form_values(request.form)


@bp.post("/<int:doc_id>/draft")
def draft(doc_id: int):
    """入力内容の途中保存（fetch）。変わった項目だけ保存し、204 を返す。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    if _is_stale(doc):
        return jsonify(error=STALE_MESSAGE), 409
    version = _version(doc)
    if _apply_values(extraction, _json_values(), _data(doc, "confirmed_json")):
        title = None if doc["state"] in CONFIRMED_STATES else _search_title(doc, extraction)
        new_json = _dumps(extraction)
        db.save_draft(doc_id, new_json, title=title)
        version = _version({"data_json": new_json})
    # 画面は次の保存・確定でこの版を送る
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


@bp.post("/<int:doc_id>/ai-fill")
def ai_fill(doc_id: int):
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return redirect(url_for(".type_select", doc_id=doc_id))
    if _is_stale(doc):
        flash(STALE_MESSAGE, "error")
        return redirect(url_for(".review", doc_id=doc_id))
    _apply_values(extraction, _form_values(request.form), _data(doc, "confirmed_json"))  # 画面で入力中の値を先に反映
    info = _load_info(doc)
    filled = None
    if info is None:
        flash("元のファイルを読み込めませんでした", "error")
    else:
        try:
            filled = ai_assist.fill_missing(info, extraction)
        except Exception as exc:
            flash(f"AIで空欄を探せませんでした: {llm.friendly_error(exc)}", "error")
    refresh_summary(extraction)
    latest = db.get_document(doc_id)
    if latest is None or _version(latest) != _version(doc) or latest["state"] != doc["state"]:
        # AIの応答を待つ間に別のタブで途中保存・確定された。読み込んだときの版で上書きしない
        flash(STALE_MESSAGE, "error")
        return redirect(url_for(".review", doc_id=doc_id))
    new_json = _dumps(extraction)
    if new_json != doc["data_json"]:
        db.save_draft(doc_id, new_json)
    if filled:
        flash(f"AIが{len(filled)}項目を入力しました（{'、'.join(filled)}）。元のファイルと照合してから確定してください", "success")
    elif filled is not None:
        flash("AIでも空欄の値は見つかりませんでした", "info")
    return redirect(url_for(".review", doc_id=doc_id))


@bp.post("/<int:doc_id>/confirm")
def confirm(doc_id: int):
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return redirect(url_for(".type_select", doc_id=doc_id))
    if _is_stale(doc):
        # 別のタブで読み直した・直した内容を、見ていない画面から確定しない
        flash(STALE_MESSAGE, "error")
        return redirect(url_for(".review", doc_id=doc_id))
    changed = _apply_values(extraction, _form_values(request.form), _data(doc, "confirmed_json"))
    new_json = _dumps(extraction) if changed else doc["data_json"]
    if changed:
        db.save_draft(doc_id, new_json)
    missing = extraction.get("missing_required") or []
    if missing and request.form.get("allow_missing") != "on":
        flash(f"必須の項目が空欄です（{'、'.join(missing)}）。値を入れるか、「空欄のまま確定する」にチェックを入れてください",
              "error")
        return redirect(url_for(".review", doc_id=doc_id) + "#confirm")
    db.confirm_document(doc_id, title=_search_title(doc, extraction))
    return redirect(url_for(".done", doc_id=doc_id))


# ---- 4 完了 ----------------------------------------------------------------------

@bp.get("/<int:doc_id>/done")
def done(doc_id: int):
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        return redirect(url_for(".review", doc_id=doc_id))
    file_name = markdown_filename(doc, confirmed)
    hash8 = str(doc.get("file_hash") or "")[:8]
    return render_template("forms/done.html", steps=STEPS, doc=doc, file_name=file_name,
                           # 識別番号が無くてファイル名の末尾に元ファイルの符号が付いたときだけ、その説明を出す
                           hash_suffix=bool(hash8) and Path(file_name).stem.endswith(hash8),
                           markdown=build_markdown(doc, confirmed), batch=_batch_info(doc),
                           delete_note=DELETE_ON_DOWNLOAD_NOTE, delete_confirm=delete_confirm(doc),
                           batch_confirm=BATCH_DELETE_CONFIRM)


# ---- 詳細・修正・削除 --------------------------------------------------------------------

@bp.get("/<int:doc_id>")
def detail(doc_id: int):
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    working = _data(doc)
    shown = confirmed or working
    return render_template(
        "forms/detail.html",
        doc=doc,
        confirmed=confirmed,
        extraction=shown,
        markdown=build_markdown(doc, confirmed) if confirmed else None,
        file_name=markdown_filename(doc, confirmed) if confirmed else None,
        changed_fields=_changed_fields(confirmed, working) if confirmed and doc["state"] == "modified" else [],
        data_types=DATA_TYPES,
        batch=_batch_info(doc),
        delete_note=DELETE_ON_DOWNLOAD_NOTE,
        delete_confirm=delete_confirm(doc),
    )


def _changed_fields(confirmed: dict, working: dict | None) -> list[str]:
    if not working:
        return []
    before = {f["field_name"]: f.get("value") for f in confirmed.get("fields", [])}
    changed = [f["display_name"] for f in working.get("fields", []) if before.get(f["field_name"]) != f.get("value")]
    if confirmed.get("pattern", {}).get("id") != working.get("pattern", {}).get("id") or \
            confirmed.get("sheets") != working.get("sheets"):
        changed.insert(0, "帳票の種類・シート")
    return changed


@bp.post("/<int:doc_id>/discard-changes")
def discard_changes(doc_id: int):
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        flash("確定済みの版がないため、戻せません", "error")
        return redirect(url_for(".review", doc_id=doc_id))
    db.discard_changes(doc_id, title=_search_title(doc, confirmed))
    pattern_id = confirmed.get("pattern", {}).get("id")
    if pattern_id and pattern_id != doc["pattern_id"] and db.load_pattern(pattern_id) is not None:
        db.update_document(doc_id, pattern_id=pattern_id)
    flash("修正中の変更を破棄し、確定済みの版に戻しました", "info")
    return redirect(url_for(".detail", doc_id=doc_id))


@bp.get("/<int:doc_id>/download.md")
def download_md(doc_id: int):
    """Markdown を渡し、渡し終えた帳票のデータを消す（design.md 3.3）。

    Markdown は全文をメモリに作ってから消すので、消す処理で中身が欠けることはない。作れなかったときは
    何も消さない。消すのは応答を送り終えたあと（purge.purge_after_send）。
    """
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        abort(404)
    body = build_markdown(doc, confirmed).encode("utf-8")
    name = markdown_filename(doc, confirmed)
    # charset は Flask が text/* に付ける。ここで付けると「charset=utf-8」が二重になる
    response = send_file(io.BytesIO(body), mimetype="text/markdown", as_attachment=True, download_name=name,
                         conditional=False)   # Range でも全体を返す（一部だけ渡して消すことが無いように）
    set_download_name(response, name, "form")
    return purge.purge_after_send(response, purge.purge_documents, [doc_id])


def _unique_name(name: str, used: set[str]) -> str:
    stem, suffix = Path(name).stem, Path(name).suffix or ".md"
    candidate, n = name, 2
    while candidate in used:
        candidate = f"{stem}_{n}{suffix}"
        n += 1
    used.add(candidate)
    return candidate


@bp.get("/batches/<batch_id>/download.zip")
def download_batch(batch_id: str):
    """まとめ取り込み1回分の Markdown を zip で渡し、渡し終えた分のデータを消す。

    ?confirmed_only=1 のときは、確定済みの帳票だけを zip にして、その分だけ消す（未確定の帳票は残す）。
    読めない帳票が1件混ざっただけで、確定済みの帳票を取り出せなくならないようにするため。
    """
    docs = db.list_batch_documents(batch_id)
    if not docs:
        abort(404)
    pending = [d for d in docs if d["state"] not in CONFIRMED_STATES]
    confirmed_ids = [d["id"] for d in docs if d["state"] in CONFIRMED_STATES]
    confirmed_only = request.args.get("confirmed_only") == "1"
    if pending and not confirmed_only:
        flash(f"まだ確定していない帳票が{len(pending)}件あります。すべて確定するか、"
              f"確定済みの{len(confirmed_ids)}件だけをダウンロードしてください", "error")
        return redirect(url_for(".type_select", doc_id=pending[0]["id"]))
    if not confirmed_ids:
        flash("確定済みの帳票がありません", "error")
        return redirect(url_for("home.index"))
    buffer, used = io.BytesIO(), set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for doc in db.list_confirmed_documents(confirmed_ids):
            # 修正中の帳票も、確定済みの版で作る（Markdown は確定済みデータから毎回生成）
            try:
                extraction = json.loads(doc["confirmed_json"])
            except (TypeError, ValueError):
                continue
            zf.writestr(_unique_name(markdown_filename(doc, extraction), used),
                        build_markdown(doc, extraction).encode("utf-8"))
    buffer.seek(0)
    zip_name = f"帳票Markdown_{datetime.now():%Y%m%d_%H%M%S}.zip"
    response = send_file(buffer, mimetype="application/zip", as_attachment=True, download_name=zip_name,
                         conditional=False)   # Range でも全体を返す
    set_download_name(response, zip_name, "forms_markdown")
    if pending:
        return purge.purge_after_send(response, purge.purge_documents, confirmed_ids)
    return purge.purge_after_send(response, purge.purge_batch, batch_id)


@bp.get("/<int:doc_id>/download.json")
def download_json(doc_id: int):
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        abort(404)
    body = json.dumps(build_json(doc, confirmed), ensure_ascii=False, indent=2).encode("utf-8")
    name = str(Path(markdown_filename(doc, confirmed)).with_suffix(".json"))
    return set_download_name(
        send_file(io.BytesIO(body), mimetype="application/json", as_attachment=True, download_name=name), name, "form")


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
    _get_document(doc_id)
    purge.purge_documents([doc_id])
    # ファイル名は出さない: flash はブラウザのセッションクッキーに載るので、取引先名や「社外秘」を
    # 含みうる名前をサーバー側から消したあともブラウザに残ってしまう（design.md 3.3）
    flash("帳票を削除しました（元のファイルと読み取り結果を消しました）", "info")
    return redirect(safe_next(url_for("home.index")))
