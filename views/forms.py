"""帳票取り込み（1ファイル＝1件）。画面は /forms の1枚だけ。

上から順に「ファイルを置く」「帳票の種類とシート」「読み取り結果」「確定してダウンロード」の
欄が現れる。画面の移動はなく、どの操作も fetch でこのファイルのルートを呼び、返ってきた
HTML の断片をその場に入れ替える（views/forms.py のルートは JSON か HTML の断片を返す）。

Markdown は常にデータから作る（読み取り結果のプレビュー＝作業中の値、ダウンロード＝確定済みの値）。
複数ファイルをまとめて置くと「取り込みのまとまり（batch）」になり、同じフォームの帳票として
まとめて1回だけ種類とシートを決め、読み取り結果は全部の帳票を縦に並べて見ていく（タブで1件ずつ
切り替えない。利用者の指示 2026-09-20）。最後に zip でまとめて渡す。
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
from types import SimpleNamespace
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
from views import current_session_id, owns, set_download_name

bp = Blueprint("forms", __name__, url_prefix="/forms")

# 左のシートプレビューに出す範囲の上限（大きいシートで画面が重くならないように）
GRID_MAX_ROWS = 300
GRID_MAX_COLS = 60
CONFIRMED_STATES = ("confirmed", "modified")
MAX_BATCH_FILES = 50
# ダウンロードでデータが消えることの案内（画面の文言・確認ダイアログで使う）
DELETE_ON_DOWNLOAD_NOTE = ("ダウンロードすると、この帳票の元のファイルと読み取り結果はサーバーから消えます。"
                           "同じものをもう一度ダウンロードすることはできません。")
DELETE_ON_DOWNLOAD_CONFIRM = "ダウンロードすると、この帳票のデータはサーバーから消えます。もう一度ダウンロードすることはできません。"
# 確定したあとに直した（修正中の）帳票は、直した値が .md に入るよう確定し直してから渡す
MODIFIED_DOWNLOAD_CONFIRM = "確定し直してから、直した値で Markdown を作ります。" + DELETE_ON_DOWNLOAD_CONFIRM
BATCH_DELETE_CONFIRM = ("ダウンロードすると、このまとまりの帳票のデータはサーバーからすべて消えます。"
                        "もう一度ダウンロードすることはできません。")
LOST_WORK_MESSAGE = "読み取り直すと、手で修正した値は失われます"
# 別のタブ・古い画面から保存・確定されたときの案内
STALE_MESSAGE = "別の画面で内容が変わりました。画面を読み込み直してください"


# ---- 共通 -------------------------------------------------------------------------

def _get_document(doc_id: int) -> dict:
    """この画面（このブラウザ）の帳票を返す。ほかの人の帳票は「無い」として扱う（404）。

    403 にすると「その番号の帳票はある」ことが分かってしまうので、404 にそろえる。
    """
    doc = db.get_document(doc_id)
    if doc is None or not owns(doc):
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


@bp.app_template_filter("table_json")
def table_json(value) -> str:
    """明細表の値を画面の入力欄（hidden）に入れる JSON。"""
    return json.dumps(value, ensure_ascii=False) if is_table_value(value) else ""


@bp.app_template_filter("field_text")
def field_text(value) -> str:
    """一覧・読み取りテスト用の表示。明細表は1行1明細の「品番: X／品名: Y」。"""
    if is_table_value(value):
        return "\n".join(table_text_lines(value))
    return "" if value is None else str(value)


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
_EDIT_KEYS = {"value", "unit", "warning", "edited"}
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
    if f.get("edited"):
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
    issue = error_value or (
        not blank and bool(f.get("warning")) and (source == "auto" or unit_issue or date_issue or number_issue))
    return {"source": source, "issue": bool(issue), "blank": blank}


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
    current_session_id()   # 画面を開いた時点で作業場所（クッキー）を決めておく（同時に置かれても取り違えない）
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
                              batch_id=batch_id, batch_order=order, session_id=current_session_id())


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


# ---- 画面を離れたので捨てる ------------------------------------------------------------
# 「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）。
# 画面を閉じた・隠したときに static/app.js の ragDiscard がここへ「捨てて」と送ってくる。
# 送り主は navigator.sendBeacon なので、
#   - 中身の型は text/plain（get_json(force=True) で読む）
#   - 応答は読めず、やり直しもできない（いつでも 204 を返し、4xx にしない）
# もう無い番号・ほかの人の番号・処理中のものは core.purge 側で黙って外れる。

@bp.post("/discard")
def discard():
    """この画面（このブラウザ）の、まだダウンロードしていない帳票を捨てる。"""
    payload = request.get_json(force=True, silent=True) or {}
    ids = payload.get("doc_ids") or []
    sid = current_session_id()
    try:
        if ids:
            purge.discard_documents(ids, sid)
        else:
            purge.purge_session(sid, tables=False)   # 表の取り込み（別のタブ）は巻き込まない
    except Exception as exc:   # 捨て損ねてもブラウザには伝えられない。時間切れの片付けに任せる
        current_app.logger.warning("帳票の片付けに失敗しました: %s", exc.__class__.__name__)
    return "", 204


# ---- 2 帳票の種類とシート ---------------------------------------------------------------

def _id_list(raw: str) -> list[int]:
    return [int(part) for part in raw.split(",") if part.strip().isdigit()]


@bp.get("/type", endpoint="type_all")
def type_all_fragment():
    """帳票の種類とシートの欄（まとめて置いた分すべてに同じ設定を使う）。?ids=1,2,3"""
    docs = _docs_of(_id_list(request.args.get("ids", "")))
    if not docs:
        return jsonify(error="取り込んだ帳票がありません。ファイルを置き直してください"), 404
    return _type_response(docs)


@bp.get("/<int:doc_id>/type")
def type_fragment(doc_id: int):
    """帳票1件の「帳票の種類とシート」（まとまりでも同じ欄を使う）。"""
    return _type_response([_get_document(doc_id)])


def _batch_matches(ranked: list[dict]) -> list[SimpleNamespace]:
    """置かれたファイル全部をまとめた帳票の種類の候補（見つかった項目が多い順）。

    同じフォームの帳票をまとめて置く前提なので、種類は1つだけ選ぶ。件数で違う分は
    「3項目中2〜3項目」のように幅で出し、ファイルごとの数は下のファイルの行に出す。
    """
    out = []
    for pattern_id in (ranked[0] if ranked else {}):
        ms = [r[pattern_id] for r in ranked if pattern_id in r]
        if not ms:
            continue
        sheets: list[str] = []
        for m in ms:
            sheets.extend(n for n in m.sheet_names if n not in sheets)
        lo, hi = min(m.found_fields for m in ms), max(m.found_fields for m in ms)
        total = ms[0].total_fields
        out.append(SimpleNamespace(
            pattern=ms[0].pattern, total_fields=total, found_fields=lo, sheet_names=sheets,
            confidence=sum(m.confidence for m in ms) / len(ms),
            found_label=(f"{total}項目中{lo}項目が見つかりました" if lo == hi
                         else f"{total}項目中{lo}〜{hi}項目が見つかりました")))
    out.sort(key=lambda m: m.confidence, reverse=True)
    return out


def _type_response(docs: list[dict]):
    patterns = db.load_active_patterns()
    ranked: list[dict] = []
    sheets: list[dict] = []
    by_name: dict[str, dict] = {}
    table_sheets: list[str] = []
    # ブックは1件ずつ開いて、必要な数だけ取り出したらすぐ手放す。50件を同時にメモリへ広げると
    # サーバーが落ちるので、置かれたファイル全部の WorkbookInfo を持ち続けない
    for doc in docs:
        info = _load_info(doc)
        if info is None:
            return jsonify(error="元のファイルを読み込めませんでした。ファイルを置き直してください"), 409
        ranked.append({m.pattern.id: m for m in rank_patterns(info, patterns)})
        for name, grid in info.grids.items():
            sheet = by_name.get(name)
            if sheet is None:
                sheet = {"name": name, "cells": len(grid.cells), "images": len(info.images_in([name])),
                         "hidden": grid.hidden, "files": 0}
                by_name[name] = sheet
                sheets.append(sheet)
            sheet["files"] += 1
        table_sheets.extend(n for n in table_like_sheets(info) if n not in table_sheets)
        info = grid = None   # 次の1件を開く前にブックを手放す

    matches = _batch_matches(ranked)
    suggested = {str(m.pattern.id): m.sheet_names for m in matches}
    first = docs[0]
    selected_id = first["pattern_id"] if any(m.pattern.id == first["pattern_id"] for m in matches) else None
    selected_id = selected_id or (matches[0].pattern.id if matches else None)
    # 読み取り済みの帳票が読んだシートを全部合わせる。1件分の "sheets" には、そのファイルに
    # 在ったシートしか残らないので、先頭ファイルだけを見ると選んだチェックが消えてしまう
    selected_sheets: list[str] = []
    for d in docs:
        data = _data(d) if d["pattern_id"] == selected_id else None
        for name in (data or {}).get("sheets") or []:
            if name not in selected_sheets:
                selected_sheets.append(name)
    if not selected_sheets:
        selected_sheets = suggested.get(str(selected_id)) or []

    files = [{"id": d["id"], "file_name": d["file_name"],
              "found": ranked[i].get(selected_id).found_fields if selected_id in ranked[i] else 0}
             for i, d in enumerate(docs)]
    # 種類を選び直したときにファイルごとの項目数を出し直すための表（画面の JS が使う）
    file_counts = {str(pid): {str(d["id"]): ranked[i][pid].found_fields for i, d in enumerate(docs)}
                   for pid in (ranked[0] if ranked else {})}
    html = render_template(
        "forms/_type.html",
        doc=first,
        docs=docs,
        files=files,
        file_counts=file_counts,
        matches=matches,
        sheets=sheets,
        suggested=suggested,
        selected_id=selected_id,
        selected_sheets=selected_sheets,
        table_sheets=table_sheets,
        duplicate=any(db.find_confirmed_by_hash(d["file_hash"], exclude_id=d["id"],
                                                session_id=current_session_id()) for d in docs),
        lost_work=any(d["state"] in CONFIRMED_STATES for d in docs),
        form_types_url=url_for("form_types.index"),
    )
    return jsonify(html=html, file_name=first["file_name"], has_types=bool(matches),
                   docs=[{"id": d["id"], "file_name": d["file_name"]} for d in docs])


@bp.post("/read", endpoint="read_all")
def read_all():
    """置かれた帳票を1つの種類・シートでまとめて読み取り、全部の読み取り結果を返す。"""
    return _read_documents(_docs_of(_id_list(request.form.get("ids", ""))))


@bp.post("/<int:doc_id>/read")
def read(doc_id: int):
    """帳票1件を読み取る（まとまりでも同じ道すじを通る）。"""
    return _read_documents([_get_document(doc_id)])


def _read_documents(docs: list[dict]):
    if not docs:
        return jsonify(error="取り込んだ帳票がありません。ファイルを置き直してください"), 404
    pattern = db.load_pattern(request.form.get("pattern_id", type=int))
    sheets = request.form.getlist("sheets")
    if pattern is None or not sheets:
        return jsonify(error="帳票の種類と読み取るシートを選んでください"), 400
    if any(d["state"] in CONFIRMED_STATES for d in docs) and request.form.get("acknowledge") != "on":
        return jsonify(error=f"{LOST_WORK_MESSAGE}。確認のチェックを入れてから読み取り直してください"), 400

    read_ids, errors = [], []

    def failed(doc: dict, reason: str) -> None:
        """読み取れなかった帳票。③に並ばないので、前の読み取り結果も残さない
        （残すと、画面に出ていない帳票を④が数えて zip の件数と合わなくなる）。"""
        errors.append(f"{doc['file_name']}: {reason}")
        if doc["state"] not in CONFIRMED_STATES and doc.get("data_json"):
            db.reset_document(doc["id"], doc["pattern_id"], None)

    for doc in docs:
        info = _load_info(doc)
        if info is None:
            failed(doc, "元のファイルを読み込めませんでした")
            continue
        use = [s for s in sheets if s in info.grids]
        if not use:
            failed(doc, "選んだシートがファイルにありません")
            continue
        extraction = extract_document(info, pattern, use)
        db.reset_document(doc["id"], pattern.id, _dumps(extraction))
        if doc["state"] not in CONFIRMED_STATES:
            db.update_document(doc["id"], title=_search_title(doc, extraction))
        read_ids.append(doc["id"])
    if not read_ids:
        return jsonify(error=errors[0] if errors else "読み取れる帳票がありませんでした", errors=errors), 400
    return _review_response(read_ids, errors=errors)


def _state_label(doc: dict, summary: dict) -> str:
    """読み取り結果の見出しに出す、その帳票の今の状態。"""
    if doc["state"] == "confirmed":
        return "確定済み"
    if doc["state"] == "modified":
        return "修正中"
    issue = summary["counts"]["issue"]
    return f"要確認 {issue}件" if issue else "未確定"


STATE_COLORS = {"確定済み": "green", "修正中": "violet", "未確定": "gray"}


def _review_response(doc_ids, errors=None):
    """読み取り結果の欄（帳票を縦に並べた HTML の断片）。

    重くならないよう、元のシートの表を作るのは先頭の帳票だけにする。残りは画面に入った時点で
    GET /forms/<id>/grid を呼んで読み込む（static/review.js）。
    """
    ids = [doc_ids] if isinstance(doc_ids, int) else list(doc_ids)
    items, docs = [], []
    for no, doc_id in enumerate(ids, start=1):
        doc = _get_document(doc_id)
        extraction = _data(doc)
        if extraction is None:
            if len(ids) == 1:
                return jsonify(error="まだ読み取りをしていません"), 409
            continue
        summary = _summary(doc, extraction)
        info = _load_info(doc) if no == 1 else None
        label = _state_label(doc, summary)
        items.append({
            "no": no, "doc": doc, "extraction": extraction, "summary": summary,
            "grids": _sheet_grids(info, extraction["sheets"]) if no == 1 else None,
            "info_missing": no == 1 and info is None,
            "version": _version(doc), "state_label": label,
            "state_color": STATE_COLORS.get(label, "amber"),
        })
        docs.append({"id": doc["id"], "file_name": doc["file_name"], "state": doc["state"],
                     "version": _version(doc), "summary": summary, "state_label": label})
    if not items:
        return jsonify(error="まだ読み取りをしていません"), 409
    html = render_template(
        "forms/_review.html",
        items=items,
        confirmed=sum(1 for d in docs if d["state"] in CONFIRMED_STATES),
    )
    first = items[0]
    return jsonify(html=html, version=first["version"], summary=first["summary"], doc_id=first["doc"]["id"],
                   docs=docs, errors=errors or [])


@bp.get("/<int:doc_id>/grid")
def grid_fragment(doc_id: int):
    """1件の帳票の「元のシート」（読み取り結果の欄が画面に入ったときに読み込む）。"""
    doc = _get_document(doc_id)
    extraction = _data(doc)
    if extraction is None:
        return jsonify(error="まだ読み取りをしていません"), 409
    info = _load_info(doc)
    html = render_template("forms/_grid.html", grids=_sheet_grids(info, extraction["sheets"]),
                           info_missing=info is None)
    return jsonify(html=html, doc_id=doc_id)


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
    values = _json_values()
    if values and _apply_values(extraction, values, _data(doc, "confirmed_json")):
        db.save_draft(doc_id, _dumps(extraction))
    db.confirm_document(doc_id, title=_search_title(doc, extraction))
    return jsonify(ok=True, doc_id=doc_id)


def _docs_of(ids: list[int]) -> list[dict]:
    """この画面の帳票だけ（ほかのブラウザの帳票は、番号を送られても無いものとして外す）。"""
    return [d for d in (db.get_document(i) for i in ids) if d is not None and owns(d)]


@bp.get("/finish", endpoint="finish")
def finish_fragment():
    """確定してダウンロードの欄（HTML の断片）。?ids=1,2,3 は同じ画面で扱っている帳票。"""
    docs = _docs_of(_id_list(request.args.get("ids", "")))
    if not docs:
        return jsonify(html="", ready=False, confirmed=0, total=0)
    # 読み取れていない帳票は確定できない＝zip に入らないので、件数には数えない
    # （数えると「残り12件も確定して…」と書いてあるのに .md が4件しか入らない zip になる）
    ready = [d for d in docs if d.get("data_json")]
    unread = [d for d in docs if not d.get("data_json")]
    current_id = request.args.get("current", type=int)
    current = next((d for d in ready if d["id"] == current_id), None) or (ready[0] if ready else docs[0])
    confirmed = [d for d in ready if d["state"] in CONFIRMED_STATES]
    pending = [d for d in ready if d["state"] not in CONFIRMED_STATES]
    batch_id = current.get("batch_id") or ""
    working = _data(current)
    read_yet = working is not None or any(_data(d) is not None for d in docs)
    extraction = _data(current, "confirmed_json") or working
    html = render_template(
        "forms/_finish.html",
        docs=docs,
        ready=ready,
        unread=unread,
        current=current,
        confirmed=confirmed,
        pending=pending,
        batch_id=batch_id if len(docs) > 1 else "",
        read_yet=read_yet,
        to_confirm=_to_confirm(ready),
        confirmed_states=CONFIRMED_STATES,
        file_name=markdown_filename(current, extraction) if extraction else "",
        markdown=build_markdown(current, _data(current, "confirmed_json")) if current["state"] in CONFIRMED_STATES else "",
        delete_note=DELETE_ON_DOWNLOAD_NOTE,
        delete_confirm=(MODIFIED_DOWNLOAD_CONFIRM if current["state"] == "modified"
                        else DELETE_ON_DOWNLOAD_CONFIRM),
        batch_confirm=_batch_zip_confirm(ready, unread),
    )
    return jsonify(html=html, ready=bool(confirmed), confirmed=len(confirmed), total=len(ready),
                   read_yet=read_yet,
                   next_id=(pending[0]["id"] if pending else None))


def _to_confirm(docs: list[dict]) -> list[dict]:
    """まとめてのダウンロードの前に確定し直す帳票（未確定と、直したまま確定していない修正中）。"""
    return [d for d in docs if d["state"] != "confirmed"]


def _batch_zip_confirm(docs: list[dict], unread: list[dict] | None = None) -> str:
    """まとまりの zip ダウンロードの確認文（確定していない分はこのボタンで確定してから渡す）。

    docs は zip に入る帳票（読み取り済み）だけ。読み取れていない帳票はサーバーに残るので、
    「すべて消えます」とは言わない。
    """
    rest = _to_confirm(docs)
    head = (f"まだ確定していない{len(rest)}件も確定してから、{len(docs)}件をまとめて zip でダウンロードします。"
            if rest else f"{len(docs)}件をまとめて zip でダウンロードします。")
    if unread:
        return (head + "ダウンロードすると、渡したこの"
                f"{len(docs)}件のデータはサーバーから消えます。"
                f"まだ読み取れていない{len(unread)}件はサーバーに残ります（②に戻ってもう一度読み取ってください）。")
    if rest:
        return head + "ダウンロードすると、このまとまりの帳票のデータはサーバーからすべて消えます。"
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
    if doc["confirmed_json"] is None:
        abort(404)   # まだ確定していない帳票は渡さない
    # 直したまま確定していない（修正中）帳票は、直した値で作る（design.md 3.3）。画面の JS を
    # 通さずにリンクを開いたとき（中クリック・「名前を付けてリンク先を保存」）でも同じにする
    extraction = _data(doc) if doc["state"] == "modified" else _data(doc, "confirmed_json")
    if extraction is None:
        abort(404)
    return markdown_filename(doc, extraction), build_markdown(doc, extraction).encode("utf-8")


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
    for doc in db.list_confirmed_documents(confirmed_ids, session_id=current_session_id()):
        # 修正中（確定したあとに直した）帳票は、直した値で作る（1件の .md と同じ）
        column = "data_json" if doc["state"] == "modified" else "confirmed_json"
        try:
            extraction = json.loads(doc[column])
        except (TypeError, ValueError):
            continue
        files.append((_unique_name(markdown_filename(doc, extraction), used),
                      build_markdown(doc, extraction).encode("utf-8")))
    return files


@bp.get("/batches/<batch_id>/download.zip")
def download_batch(batch_id: str):
    """まとめ取り込み1回分の Markdown を zip で渡し、渡し終えた分のデータを消す。

    確定済みの帳票だけを zip にして、その分だけ消す（未確定の帳票はサーバーに残る）。
    """
    docs = db.list_batch_documents(batch_id, session_id=current_session_id())
    if not docs:
        abort(404)   # ほかのブラウザのまとまりも「無い」として扱う
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
