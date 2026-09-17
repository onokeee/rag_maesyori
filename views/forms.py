"""帳票を取り込む（1ファイル＝1件）: ファイル選択 → 種類とシート → 読み取り結果の確認・修正 → 完了。

Markdown は常にデータから作る（確認画面のプレビュー＝作業中の値、ダウンロード＝確定済みの値）。
"""
from __future__ import annotations

import io
import json
from pathlib import Path

from flask import (Blueprint, abort, current_app, flash, jsonify, redirect, render_template, request, send_file,
                   url_for)

from core.files import UploadError, precheck_excel, remove_upload, save_upload, upload_path
from excel.extractor import apply_manual_values, extract_document, refresh_summary
from excel.workbook import WorkbookInfo, load_workbook_info
from export.formats import build_json, build_markdown, markdown_filename
from models import database as db
from pattern.matcher import rank_patterns, table_like_sheets
from pattern.model import DATA_TYPES
from services import ai_assist, llm
from views import safe_next

bp = Blueprint("forms", __name__, url_prefix="/forms")

STEPS = ["ファイルを選ぶ", "帳票の種類とシートを確認", "読み取り結果の確認・修正", "完了（Markdownをダウンロード）"]
# 左のシートプレビューに出す範囲の上限（大きいシートで画面が重くならないように）
GRID_MAX_ROWS = 300
GRID_MAX_COLS = 60
CONFIRMED_STATES = ("confirmed", "modified")
LOST_WORK_MESSAGE = "読み取り直すと、手で修正した値とAIが入力した値は失われます"


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
    """取り込み履歴の検索用タイトル（Markdown の1行目の見出し）。"""
    first = build_markdown(doc, extraction).split("\n", 1)[0]
    return first[2:].strip() if first.startswith("# ") else first.strip()


def _display_text(value) -> str:
    return "" if value is None else str(value)


def _same_text(a: str, b: str) -> bool:
    norm = lambda s: s.replace("\r\n", "\n").replace("\r", "\n").strip()  # noqa: E731
    return norm(a) == norm(b)


def _apply_values(extraction: dict, values: dict) -> bool:
    """画面の入力値を反映する。表示中の値と同じ文字列の項目は触らない（勝手に「手で修正」にしない）。

    戻り値: 変わった項目があれば True。
    """
    before = _dumps(extraction)
    fields = {f["field_name"]: f for f in extraction["fields"]}
    changed = {}
    for name, text in values.items():
        f = fields.get(name)
        if f is None or text is None:
            continue
        text = str(text)
        if _same_text(text, _display_text(f["value"])):
            continue
        changed[f"value-{name}"] = text
    if changed:
        apply_manual_values(extraction, changed)
    else:
        refresh_summary(extraction)
    return _dumps(extraction) != before


def _form_values(form) -> dict:
    return {key[len("value-"):]: form[key] for key in form.keys() if key.startswith("value-")}


def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


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
    issue = (f["required"] and blank) or (not blank and bool(f.get("warning")) and source == "auto")
    return {"source": source, "issue": bool(issue), "blank": blank,
            "missing_required": bool(f["required"] and blank)}


def _summary(doc: dict, extraction: dict) -> dict:
    """確認画面の数値（チップ・Markdownプレビュー・項目ごとの状態）。"""
    statuses = {f["field_name"]: _field_status(f) for f in extraction["fields"]}
    return {
        "markdown": build_markdown(doc, extraction),
        "file_name": markdown_filename(doc, extraction),
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
                           pattern_count=len(db.list_patterns()))


@bp.post("/upload")
def upload():
    storage = request.files.get("file")
    if storage is None or not storage.filename:
        flash("ファイルを選んでください", "error")
        return redirect(url_for(".new"))
    cfg = current_app.config
    try:
        stored = save_upload(storage, "documents", cfg["ALLOWED_EXTENSIONS"], cfg["MAX_CONTENT_LENGTH"])
    except UploadError as exc:
        flash(str(exc), "error")
        return redirect(url_for(".new"))
    try:
        path = upload_path(stored.stored_path)
        precheck_excel(path)
        try:
            load_workbook_info(path)
        except Exception as exc:
            raise UploadError(f"{stored.file_name}: Excelファイルとして読み込めませんでした（{exc.__class__.__name__}）") from exc
    except UploadError as exc:
        remove_upload(stored.stored_path)
        flash(str(exc), "error")
        return redirect(url_for(".new"))
    doc_id = db.create_document(stored.file_name, stored.file_hash, stored.stored_path)
    return redirect(url_for(".type_select", doc_id=doc_id))


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
                               lost_work=doc["state"] in CONFIRMED_STATES)
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
    )


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
    if _apply_values(extraction, _json_values()):
        title = None if doc["state"] in CONFIRMED_STATES else _search_title(doc, extraction)
        db.save_draft(doc_id, _dumps(extraction), title=title)
    return "", 204


@bp.post("/<int:doc_id>/preview")
def preview(doc_id: int):
    """入力中の値で作った Markdown と状態（保存はしない）。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    _apply_values(extraction, _json_values())
    return jsonify(_summary(doc, extraction))


@bp.post("/<int:doc_id>/ai-fill")
def ai_fill(doc_id: int):
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return redirect(url_for(".type_select", doc_id=doc_id))
    _apply_values(extraction, _form_values(request.form))  # 画面で入力中の値を先に反映
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
    changed = _apply_values(extraction, _form_values(request.form))
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
    return render_template("forms/done.html", steps=STEPS, doc=doc, file_name=markdown_filename(doc, confirmed),
                           markdown=build_markdown(doc, confirmed))


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
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        abort(404)
    body = build_markdown(doc, confirmed).encode("utf-8")
    return send_file(io.BytesIO(body), mimetype="text/markdown; charset=utf-8", as_attachment=True,
                     download_name=markdown_filename(doc, confirmed))


@bp.get("/<int:doc_id>/download.json")
def download_json(doc_id: int):
    doc = _get_document(doc_id)
    confirmed = _data(doc, "confirmed_json")
    if confirmed is None:
        abort(404)
    body = json.dumps(build_json(doc, confirmed), ensure_ascii=False, indent=2).encode("utf-8")
    name = str(Path(markdown_filename(doc, confirmed)).with_suffix(".json"))
    return send_file(io.BytesIO(body), mimetype="application/json", as_attachment=True, download_name=name)


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
    doc = _get_document(doc_id)
    try:
        remove_upload(doc["stored_path"])
    except UploadError:
        pass
    db.delete_document(doc_id)
    flash(f"{doc['file_name']} を削除しました", "info")
    return redirect(safe_next(url_for("history.index")))
