"""帳票登録。画面は /form-types の1枚だけ。

上から「登録済みの帳票の種類」、その下に「新しく登録する」。Excel を1つ置くと、同じ画面に
名前（ファイル名から入れる）・置いた Excel のシート・読み取る項目・読み取りテストの結果・
［使用開始］が現れる。画面の移動はなく、どの操作も fetch でこのファイルのルートを呼び、
HTML の断片を入れ替える。

項目は「見出しのセル → 値のセル」をクリックするだけで作る。キー名・型・単位・探す見出しは
pattern.clicks が見本の値から決めるので、画面には出さない。
保存しても使用中にはしない。使用中になるのは［使用開始］を押したときだけ。

置いた Excel はサーバーに残さない（利用者の指示 2026-09-21「見本のExcelは置かずに、設定だけ
保持するようにしてほしい」）。受け取った要求の中で読み取り、中身はそのまま捨てる。ブラウザは
選んだファイルを持ったままなので、セルのクリック・項目の作り直し・読み取りテストのたびに同じ
Excel を送り直してくる。開き直すのが遅いので、読み取った結果だけを core.workbook_cache が
短い間メモリに覚えておく（ディスクには書かない）。
残すのは設定だけ: シート名・見出しのセル・値のセル・読み取る向き・項目名。
"""
from __future__ import annotations

import io
from pathlib import Path

from flask import Blueprint, abort, current_app, jsonify, redirect, render_template, request, url_for

from core import workbook_cache
from core.files import FORM_MAX_MERGED_CELLS, UploadError, precheck_excel, read_upload
from core.workbook_cache import Book
from excel.extractor import extract_document
from excel.workbook import load_workbook_info
from export.formats import build_markdown, markdown_filename
from models import database as db
from pattern.builder import suggest_title_fields
from pattern.clicks import (click_field, merge_labels, merge_target, same_sheet_field, separate_names, split_rows,
                            table_cells)
from pattern.forms import pattern_to_meta, pattern_to_rows, rows_to_pattern
from pattern.matcher import match_pattern
from pattern.model import PatternDef
from views import current_session_id
from views.forms import _sheet_grids, upload_error_text

bp = Blueprint("form_types", __name__, url_prefix="/form-types")

# excel/extractor.number_unit の単位不明の警告の書き出し
_NO_UNIT_WARNING = "単位が書かれていません"
TEST_NO_UNIT_WARNING = ("帳票に単位が書かれていません。取り込んだあとの「読み取り結果」で、"
                        "値に単位（分・時間など）を付けて入力できます")


def _get_pattern(pattern_id: int) -> PatternDef:
    pattern = db.load_pattern(pattern_id)
    if pattern is None:
        abort(404)
    return pattern


def _confirmed_count(pattern_id: int) -> int:
    row = db.get_db().execute(
        "SELECT COUNT(*) FROM documents WHERE pattern_id = ? AND confirmed_json IS NOT NULL", (pattern_id,)
    ).fetchone()
    return row[0] if row else 0


# ---- ブラウザが置いた Excel（保存しない） -----------------------------------------------
# 置かれた Excel は、この要求の中で読み取って中身を捨てる。次の操作のときはブラウザが同じ
# ファイルを送り直してくるので、2回目からは読み取った結果（core.workbook_cache）を使い回す。

BOOK_FIELD = "book"            # ブラウザが送ってくる Excel（<input type=file name=book>）
BOOK_HASH_FIELD = "book_hash"  # 送り直さずに、さっき読んだブックを指すとき（sha256）
NO_BOOK_ERROR = "この帳票のExcelをもう一度置いてください（サーバーには残していません）"
# ここで受け取る Excel の大きさの上限。保存せずにメモリで読む（数人が同時に置く）ので、
# 取り込みの上限（MAX_CONTENT_LENGTH＝まとめ置きの合計）より小さくしておく。
# 帳票は1枚の紙なので、写真付きでもこの大きさに収まる（見本のいちばん大きいもので約0.1MB）。
BOOK_MAX_BYTES = 50 * 1024 * 1024


def _read_book(storage) -> Book:
    """置かれた Excel をメモリで読み、読み取った結果だけを覚える。読めなければ UploadError。

    受け取ったファイルは werkzeug が 500KB まではメモリに、それを超える分だけ OS の一時ファイルに
    置く（要求が終わると消える、名前の無いファイル）。こちらからディスクに書くことはしない。
    """
    cfg = current_app.config
    limit = min(cfg["MAX_CONTENT_LENGTH"] or BOOK_MAX_BYTES, BOOK_MAX_BYTES)
    memory = read_upload(storage, cfg["ALLOWED_EXTENSIONS"], limit)
    known = workbook_cache.get(current_session_id(), memory.file_hash)
    if known is not None:
        known.file_name = memory.file_name   # 同じ中身を別の名前で置き直したとき
        return known
    precheck_excel(memory.data, cfg.get("EXCEL_MAX_CELLS"), max_merged=FORM_MAX_MERGED_CELLS)
    try:
        # BytesIO で渡すので、どこにもファイルを作らずに読める
        info = load_workbook_info(io.BytesIO(memory.data))
    except Exception as exc:
        current_app.logger.warning("置かれたExcelを読み込めませんでした: %s", exc.__class__.__name__)
        raise UploadError("Excelファイルとして読み込めませんでした") from exc
    if not info.grids:
        raise UploadError("シートがないブックです。シートのあるブックを選んでください")
    return workbook_cache.put(current_session_id(),
                              Book(file_name=memory.file_name, file_hash=memory.file_hash,
                                   size=memory.size, info=info))


def _request_book() -> tuple[Book | None, str]:
    """この操作で見ている Excel。戻り値: (ブック, エラー文)。

    ブラウザが送ってきたファイルを読む。ファイルが無いときは、さっき読んだブック（sha256）を探す。
    どちらも無ければ (None, "")＝Excel を置いていない画面（項目の一覧と見出しの手直しはできる）。
    """
    storage = request.files.get(BOOK_FIELD)
    if storage is not None and storage.filename:
        try:
            return _read_book(storage), ""
        except UploadError as exc:
            return None, upload_error_text(storage, exc)
    file_hash = (request.form.get(BOOK_HASH_FIELD) or request.args.get(BOOK_HASH_FIELD) or "").strip()
    if file_hash:
        return workbook_cache.get(current_session_id(), file_hash), ""
    return None, ""


# ---- 1枚の画面 --------------------------------------------------------------------

@bp.get("/", endpoint="index")
def page():
    # 帳票の種類は「設定」なのでみんなで使う（ブラウザごとに分けない）。
    # 作業場所のクッキーだけは、ここで開いたときにも決めておく（ほかの画面での取り違えを防ぐ）
    current_session_id()
    return render_template("form_types/page.html", list_html=_list_html(),
                           max_mb=BOOK_MAX_BYTES // (1024 * 1024))


def _list_html() -> str:
    return render_template("form_types/_list.html", patterns=db.list_patterns())


# 画面を作り直す前の URL（お気に入り・古いリンク）は1枚の画面へ送る
@bp.get("/new")
@bp.get("/<int:pattern_id>/build")
@bp.get("/<int:pattern_id>/edit")
@bp.get("/<int:pattern_id>/review")
@bp.get("/<int:pattern_id>/test")
def legacy(pattern_id: int | None = None):
    return redirect(url_for(".index"))


# ---- 新しく登録する ---------------------------------------------------------------

@bp.post("/new")
def create():
    """Excel を1つ置いて帳票の種類を作る（fetch）。名前は画面でファイル名から入れてある。

    作るのは設定だけ。置かれた Excel は読み取るだけで、サーバーには残さない。
    """
    storage = request.files.get(BOOK_FIELD)
    if storage is None or not storage.filename:
        return jsonify(error="帳票のExcelファイルを置いてください"), 400
    name = request.form.get("name", "").strip() or Path(storage.filename).stem.strip()
    if not name:
        return jsonify(error="帳票の種類の名前を入れてください"), 400
    try:
        book = _read_book(storage)
    except UploadError as exc:
        return jsonify(error=upload_error_text(storage, exc)), 400
    pattern_id = db.create_pattern(name)
    return jsonify(pattern_id=pattern_id, html=_build_html(pattern_id, book), list_html=_list_html(),
                   message=f"「{name}」を作りました。読み取りたい欄の見出しと値をクリックしてください")


# ---- 読み取る欄をクリックして決める ＋ 読み取りテスト ------------------------------------

def _build_html(pattern_id: int, book: Book | None = None, notes: list[str] | None = None) -> str:
    """登録中の帳票の種類の欄（HTML の断片）。book が無ければシートの無い（設定だけの）画面。"""
    pattern = _get_pattern(pattern_id)
    info = book.info if book is not None else None
    grids = _sheet_grids(info, list(info.grids)) if info is not None else []
    for g in grids:
        g["click_cells"] = table_cells(info.grids[g["name"]])
    return render_template(
        "form_types/_build.html",
        pattern=pattern,
        book=book,
        grids=grids,
        rows=_field_view_rows(pattern, info),
        test=_test_result(pattern, book),
        confirmed_count=_confirmed_count(pattern_id),
        notes=notes or [],
    )


@bp.get("/<int:pattern_id>/panel")
@bp.post("/<int:pattern_id>/panel")
def build_fragment(pattern_id: int):
    """項目の一覧を出す（GET）。Excel を一緒に置くと（POST）、そのシートを見ながら直せる。

    保存済みの種類を開き直したときは Excel が無いので、項目の一覧と見出しの手直しだけができる。
    同じ帳票の Excel をもう一度置くと、シートが出てセルをクリックできるようになる（種類は増えない）。
    """
    _get_pattern(pattern_id)
    book, error = _request_book()
    if error:
        return jsonify(error=error), 400
    message = ""
    if request.method == "POST" and book is not None:
        message = (f"「{book.file_name}」を読み込みました。読み取りたい欄の見出しと値をクリックしてください"
                   "（このExcelはサーバーに残しません）")
    return jsonify(html=_build_html(pattern_id, book), list_html=_list_html(), message=message)


def _soften_unit_warning(f: dict) -> None:
    """この欄では値を直せないので、単位なしの警告はどこで直せるかを書く。"""
    if f.get("data_type") == "number" and str(f.get("warning") or "").startswith(_NO_UNIT_WARNING):
        f["warning"] = TEST_NO_UNIT_WARNING


def _field_view_rows(pattern: PatternDef, info) -> list[dict]:
    """項目の一覧（見出し・置いた Excel で見つかった値・セル）。値はいま登録されている設定で読み直す。

    Excel を置いていないとき（info が None）は値の欄を空にする。見出しの手直しはそれでもできる。
    """
    found: dict[str, dict] = {}
    if info is not None and pattern.fields:
        sheets = [s.sheet_name for s in pattern.sheets if s.sheet_name in info.grids] or list(info.grids)[:1]
        found = {f["field_name"]: f for f in extract_document(info, pattern, sheets)["fields"]}
    for f in found.values():
        _soften_unit_warning(f)
    rows = []
    for fd in pattern.fields:
        f = found.get(fd.field_name) or {}
        rows.append({
            "field": fd,
            "label": (fd.candidates[0] if fd.candidates else "") or "（見出しなし）",
            "value": f.get("value"),
            "warning": f.get("warning") or "",
            "sheet": f.get("sheet") or fd.sheet_name,
            "label_cell": f.get("label_cell") or fd.label_cell,
            "value_cell": f.get("value_cell") or fd.cell,
        })
    return rows


def _test_result(pattern: PatternDef, book: Book | None) -> dict | None:
    """いま置いている Excel を、この設定で読み取った結果（Markdown・見つかった件数）。"""
    if book is None or not pattern.fields:
        return None
    info = book.info
    match = match_pattern(info, pattern)
    sheets = match.sheet_names or info.sheet_names[:1]
    extraction = extract_document(info, pattern, sheets)
    doc = {"id": 0, "file_name": book.file_name, "file_hash": book.file_hash}
    return {
        "sheets": sheets,
        "found": sum(1 for f in extraction["fields"] if f["value"] not in (None, "")),
        "total": len(extraction["fields"]),
        "markdown": build_markdown(doc, extraction),
        "file_name": markdown_filename(doc, extraction),
    }


@bp.post("/<int:pattern_id>/fields")
def add_field(pattern_id: int):
    """クリックした見出しセル（と値セル）から項目を1つ作る（fetch）。

    どのセルを指しているかは、ブラウザが一緒に送ってくる Excel を読み直して確かめる
    （サーバーには置いていないため）。
    """
    pattern = _get_pattern(pattern_id)
    book, error = _request_book()
    if error:
        return jsonify(error=error), 400
    if book is None:
        return jsonify(error=NO_BOOK_ERROR), 400
    info = book.info
    sheet = request.form.get("sheet", "")
    grid = info.grids.get(sheet)
    if grid is None:
        return jsonify(error="シートが見つかりません。同じ帳票のExcelを置き直してください"), 400

    label_cell = request.form.get("label_cell", "")
    value_cell = request.form.get("value_cell", "")
    row, error = click_field(grid, label_cell, value_cell, {f.field_name for f in pattern.fields})
    if row is None:
        return jsonify(error=error), 400
    if any(f.sheet_name == sheet and f.label_cell == row["label_cell"] and f.cell == row["cell"]
           for f in pattern.fields):
        return jsonify(html=_build_html(pattern_id, book), list_html=_list_html(),
                       message="そのセルはもう項目になっています")

    sheet_rows, field_rows = pattern_to_rows(pattern)
    # 番号と名前を1つのセルにまとめた「使用設備」欄は、設備番号・設備名の2項目になる
    added, separated, merged = [], [], []
    for part in split_rows(row, {f.field_name for f in pattern.fields}):
        # 別の見本で書き方の違う同じ欄（「設備No」と「設備番号」）をクリックしたときは、新しい項目にせず
        # その項目の探す見出しに足す
        same = merge_target(field_rows, part, grid)
        if same is not None:
            merge_labels(same, part)
            merged.append(same["display_name"])
            continue
        # 同じ見本の別のセル（「担当者」と「報告者」）なら、辞書の名前が同じでも別の項目にする
        twin = same_sheet_field(field_rows, part, grid)
        if twin is not None:
            separate_names(twin, part)
            separated.append((part["display_name"], twin["display_name"]))
        else:
            added.append(part["display_name"])
        field_rows.append(part)
    message = "。".join(_add_messages(row, added, separated, merged))
    if not any(r["sheet_name"] == sheet for r in sheet_rows):
        sheet_rows.append({"use": True, "sheet_name": sheet})
    _save_rows(pattern, sheet_rows, field_rows)
    return jsonify(html=_build_html(pattern_id, book), list_html=_list_html(), message=message)


def _add_messages(row: dict, added: list[str], separated: list[tuple[str, str]], merged: list[str]) -> list[str]:
    """クリックの結果の知らせ（項目にした／別の項目にした／見出しに足した、のどれをしたか）。"""
    out = []
    if added:
        out.append(f"「{'」「'.join(added)}」を項目にしました")
    for name, twin in separated:
        out.append(f"「{name}」を「{twin}」とは別の項目にしました（同じ帳票の別のセルなので、両方を読み取ります）")
    if merged:
        label = (row["candidates"].splitlines() or [""])[0]
        out.append(f"「{'」「'.join(merged)}」の見出しに「{label}」を足しました"
                   "（書き方の違う同じ欄なので、1つの項目として読み取ります）")
    return out


@bp.post("/<int:pattern_id>/fields/<field_name>/delete")
def delete_field(pattern_id: int, field_name: str):
    pattern = _get_pattern(pattern_id)
    sheet_rows, field_rows = pattern_to_rows(pattern)
    rest = [r for r in field_rows if r["field_name"] != field_name]
    if len(rest) == len(field_rows):
        abort(404)
    kept = {r["sheet_name"] for r in rest if r["sheet_name"]}
    sheet_rows = [s for s in sheet_rows if not kept or s["sheet_name"] in kept]
    # 読み取る項目が無くなった使用中の種類は、使用を停止する。そのままだと帳票取り込みの候補に出て、
    # 中身の無い Markdown ができてしまう（［使用開始］も同じ決まりで断っている）
    stopped = not rest and pattern.status == "active"
    _save_rows(pattern, sheet_rows, rest, status="inactive" if stopped else None)
    message = "項目を削除しました"
    if stopped:
        message += "。読み取る項目が無くなったので、この種類の使用を停止しました（帳票取り込みの候補に出なくなります）"
    # ブラウザが Excel を一緒に送ってきていれば、シートを出したままにする
    return jsonify(html=_build_html(pattern_id, _request_book()[0]), list_html=_list_html(), message=message)


@bp.post("/<int:pattern_id>/fields/<field_name>/label")
def rename_field(pattern_id: int, field_name: str):
    """読み取る項目の見出しを手で直す（fetch）。

    直すのは Markdown に書き出す名前だけ。探す見出し（クリックしたときの見出しの言葉）と
    読み取るセルはそのままにする。書き出す名前は項目どうしで重ならないようにする。
    """
    pattern = _get_pattern(pattern_id)
    payload = request.get_json(force=True, silent=True) or {}
    name = " ".join(str(payload.get("name") or request.form.get("name", "")).split())
    if not name:
        return jsonify(error="見出しを入れてください"), 400
    sheet_rows, field_rows = pattern_to_rows(pattern)
    target = next((r for r in field_rows if r["field_name"] == field_name), None)
    if target is None:
        abort(404)
    if any(r["display_name"] == name for r in field_rows if r is not target):
        return jsonify(error=f"「{name}」はほかの項目が使っています。別の見出しにしてください"), 400
    # 探す見出しを持たない項目（見出しのない表など）は、書き出す名前をそのまま探していた。
    # 書き替えで探す先が変わらないよう、いまの見出しを探す見出しとして控えてから名前を変える
    if not target["candidates"] and not target["cell"]:
        target["candidates"] = target["display_name"]
    target["display_name"] = name
    target["renamed"] = True   # このあと別の欄をクリックしても、手で付けた見出しに戻さない
    _save_rows(pattern, sheet_rows, field_rows)
    return jsonify(html=_build_html(pattern_id, _request_book()[0]), list_html=_list_html(),
                   message=f"見出しを「{name}」にしました")


@bp.post("/<int:pattern_id>/name")
def rename(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    payload = request.get_json(force=True, silent=True) or {}
    name = str(payload.get("name") or request.form.get("name", "")).strip()
    if not name:
        return jsonify(error="帳票の種類の名前を入れてください"), 400
    sheet_rows, field_rows = pattern_to_rows(pattern)
    _save_rows(pattern, sheet_rows, field_rows, name=name)
    return jsonify(ok=True, list_html=_list_html(), message="名前を変えました")


def _save_rows(pattern: PatternDef, sheet_rows: list[dict], field_rows: list[dict], name: str | None = None,
               status: str | None = None) -> None:
    """状態は変えずに保存する（使用開始は［使用開始］を押したときだけ）。タイトル項目は自動で決める。

    status を渡したときだけ状態も変える（読み取る項目が無くなったら使用を停止する）。
    """
    meta = pattern_to_meta(pattern)
    if name:
        meta["name"] = name
    meta["title_fields"] = suggest_title_fields(field_rows)
    db.save_pattern(rows_to_pattern(pattern.id, meta, sheet_rows, field_rows), status or pattern.status)


# ---- 使用開始・停止・削除 --------------------------------------------------------------

@bp.post("/<int:pattern_id>/status")
def change_status(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    payload = request.get_json(force=True, silent=True) or {}
    status = payload.get("status") or request.form.get("status")
    if status not in ("active", "inactive"):
        abort(400)
    if status == "active" and not pattern.fields:
        return jsonify(error="読み取る項目がありません。シートで見出しのセルと値のセルをクリックしてください"), 400
    db.set_pattern_status(pattern_id, status)
    if status == "active":
        # 残すのは設定だけ（Excel はもともと置いていない）。画面はそのまま続けて使える
        message = f"「{pattern.name}」の使用を開始しました。帳票取り込みの候補に出ます"
    else:
        message = f"「{pattern.name}」の使用を停止しました。帳票取り込みの候補に出なくなります"
    return jsonify(ok=True, status=status, html=_build_html(pattern_id, _request_book()[0]),
                   list_html=_list_html(), message=message)


@bp.post("/<int:pattern_id>/delete")
def delete(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    db.delete_pattern(pattern_id)
    return jsonify(ok=True, list_html=_list_html(), message=f"帳票の種類「{pattern.name}」を削除しました")
