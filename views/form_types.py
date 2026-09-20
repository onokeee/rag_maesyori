"""帳票登録。画面は /form-types の1枚だけ。

上から「登録済みの帳票の種類」、その下に「新しく登録する」。Excel を1つ置くと、同じ画面に
名前（ファイル名から入れる）・見本のシート・読み取る項目・読み取りテストの結果・［使用開始］が現れる。
画面の移動はなく、どの操作も fetch でこのファイルのルートを呼び、HTML の断片を入れ替える。

項目は「見出しのセル → 値のセル」をクリックするだけで作る。キー名・型・単位・探す見出しは
pattern.clicks が見本の値から決めるので、画面には出さない。
保存しても使用中にはしない。使用中になるのは［使用開始］を押したときだけ。
"""
from __future__ import annotations

from pathlib import Path

from flask import Blueprint, abort, current_app, jsonify, redirect, render_template, request, url_for

from core.files import FORM_MAX_MERGED_CELLS, UploadError, precheck_excel, remove_upload, save_upload, upload_path
from excel.extractor import extract_document
from excel.workbook import load_workbook_info
from export.formats import build_markdown, markdown_filename
from models import database as db
from pattern.builder import suggest_title_fields
from pattern.clicks import click_field, merge_labels, merge_target, split_rows, table_cells
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


# ---- 見本ファイル ------------------------------------------------------------------

def _save_samples(pattern_id: int, files) -> tuple[int, list[str]]:
    cfg = current_app.config
    saved, errors = 0, []
    for position, storage in enumerate(files, 1):
        if not storage or not storage.filename:
            continue
        try:
            stored = save_upload(storage, "samples", cfg["ALLOWED_EXTENSIONS"], cfg["MAX_CONTENT_LENGTH"])
        except UploadError as exc:
            errors.append(upload_error_text(storage, exc, position))
            continue
        try:
            path = upload_path(stored.stored_path)
            precheck_excel(path, cfg.get("EXCEL_MAX_CELLS"), max_merged=FORM_MAX_MERGED_CELLS)
            try:
                load_workbook_info(path)
            except Exception as exc:
                current_app.logger.warning("見本を読み込めませんでした: %s", exc.__class__.__name__)
                raise UploadError("Excelファイルとして読み込めませんでした") from exc
        except UploadError as exc:
            remove_upload(stored.stored_path)
            errors.append(upload_error_text(storage, exc, position))
            continue
        db.add_sample(pattern_id, stored.file_name, stored.file_hash, stored.stored_path)
        saved += 1
    return saved, errors


def _sample_infos(pattern_id: int) -> tuple[list[dict], list, list[str]]:
    samples, infos, errors = [], [], []
    for sample in db.list_samples(pattern_id):
        try:
            infos.append(load_workbook_info(upload_path(sample["stored_path"])))
            samples.append(sample)
        except Exception:
            errors.append("見本ファイルを読み込めませんでした")
    return samples, infos, errors


# ---- 1枚の画面 --------------------------------------------------------------------

@bp.get("/", endpoint="index")
def page():
    # 帳票の種類は「設定」なのでみんなで使う（ブラウザごとに分けない）。
    # 作業場所のクッキーだけは、ここで開いたときにも決めておく（ほかの画面での取り違えを防ぐ）
    current_session_id()
    return render_template("form_types/page.html", list_html=_list_html(),
                           max_mb=(current_app.config.get("MAX_CONTENT_LENGTH") or 0) // (1024 * 1024))


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
    """Excel を1つ置いて帳票の種類を作る（fetch）。名前は画面でファイル名から入れてある。"""
    files = [f for f in request.files.getlist("samples") if f and f.filename]
    name = request.form.get("name", "").strip() or (Path(files[0].filename).stem.strip() if files else "")
    if not files:
        return jsonify(error="帳票のExcelファイルを置いてください"), 400
    if not name:
        return jsonify(error="帳票の種類の名前を入れてください"), 400
    pattern_id = db.create_pattern(name)
    saved, errors = _save_samples(pattern_id, files)
    if not saved:
        db.delete_pattern(pattern_id)
        return jsonify(error=errors[0] if errors else "この Excel は読み込めませんでした", errors=errors), 400
    return jsonify(pattern_id=pattern_id, html=_build_html(pattern_id), list_html=_list_html(),
                   message=f"「{name}」を作りました。読み取りたい欄の見出しと値をクリックしてください", errors=errors)


# ---- 読み取る欄をクリックして決める ＋ 読み取りテスト ------------------------------------

def _sample_index(samples: list[dict], sample_id: int | None) -> int:
    return next((i for i, s in enumerate(samples) if s["id"] == sample_id), 0)


def _build_html(pattern_id: int, sample_id: int | None = None, notes: list[str] | None = None) -> str:
    """登録中の帳票の種類の欄（HTML の断片）。"""
    pattern = _get_pattern(pattern_id)
    samples, infos, errors = _sample_infos(pattern_id)
    index = _sample_index(samples, sample_id)
    info = infos[index] if infos else None
    grids = _sheet_grids(info, list(info.grids)) if info is not None else []
    for g in grids:
        g["click_cells"] = table_cells(info.grids[g["name"]])
    return render_template(
        "form_types/_build.html",
        pattern=pattern,
        samples=samples,
        sample=samples[index] if samples else None,
        grids=grids,
        rows=_field_view_rows(pattern, info),
        test=_test_result(pattern, samples[index] if samples else None, info),
        confirmed_count=_confirmed_count(pattern_id),
        notes=(notes or []) + errors,
    )


@bp.get("/<int:pattern_id>/panel")
def build_fragment(pattern_id: int):
    return jsonify(html=_build_html(pattern_id, request.args.get("sample", type=int)),
                   list_html=_list_html())


def _soften_unit_warning(f: dict) -> None:
    """この欄では値を直せないので、単位なしの警告はどこで直せるかを書く。"""
    if f.get("data_type") == "number" and str(f.get("warning") or "").startswith(_NO_UNIT_WARNING):
        f["warning"] = TEST_NO_UNIT_WARNING


def _field_view_rows(pattern: PatternDef, info) -> list[dict]:
    """項目の一覧（見出し・見本で見つかった値・セル）。値はいま登録されている設定で読み直す。"""
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


def _test_result(pattern: PatternDef, sample: dict | None, info) -> dict | None:
    """いま見ている見本を、この設定で読み取った結果（Markdown・見つかった件数）。"""
    if info is None or sample is None or not pattern.fields:
        return None
    match = match_pattern(info, pattern)
    sheets = match.sheet_names or info.sheet_names[:1]
    extraction = extract_document(info, pattern, sheets)
    doc = {"id": 0, "file_name": sample["file_name"], "file_hash": sample["file_hash"]}
    return {
        "sheets": sheets,
        "found": sum(1 for f in extraction["fields"] if f["value"] not in (None, "")),
        "total": len(extraction["fields"]),
        "missing_required": extraction["missing_required"],
        "markdown": build_markdown(doc, extraction),
        "file_name": markdown_filename(doc, extraction),
    }


@bp.post("/<int:pattern_id>/fields")
def add_field(pattern_id: int):
    """クリックした見出しセル（と値セル）から項目を1つ作る（fetch）。"""
    pattern = _get_pattern(pattern_id)
    samples, infos, _ = _sample_infos(pattern_id)
    index = _sample_index(samples, request.form.get("sample", type=int))
    info = infos[index] if infos else None
    sample_id = samples[index]["id"] if samples else None
    sheet = request.form.get("sheet", "")
    grid = info.grids.get(sheet) if info is not None else None
    if grid is None:
        return jsonify(error="シートが見つかりません。見本ファイルを置き直してください"), 400

    label_cell = request.form.get("label_cell", "")
    value_cell = request.form.get("value_cell", "")
    row, error = click_field(grid, label_cell, value_cell, {f.field_name for f in pattern.fields})
    if row is None:
        return jsonify(error=error), 400
    if any(f.sheet_name == sheet and f.label_cell == row["label_cell"] and f.cell == row["cell"]
           for f in pattern.fields):
        return jsonify(html=_build_html(pattern_id, sample_id), list_html=_list_html(),
                       message="そのセルはもう項目になっています")

    sheet_rows, field_rows = pattern_to_rows(pattern)
    # 番号と名前を1つのセルにまとめた「使用設備」欄は、設備番号・設備名の2項目になる
    added, merged = [], []
    for part in split_rows(row, {f.field_name for f in pattern.fields}):
        # 別の見本で書き方の違う同じ欄（「設備No」と「設備番号」）をクリックしたときは、新しい項目にせず
        # その項目の探す見出しに足す
        same = merge_target(field_rows, part)
        if same is not None:
            merge_labels(same, part)
            merged.append(same["display_name"])
        else:
            field_rows.append(part)
            added.append(part["display_name"])
    if added:
        message = f"「{'」「'.join(added)}」を項目にしました"
    else:
        message = f"「{'」「'.join(merged)}」の見出しに「{(row['candidates'].splitlines() or [''])[0]}」を足しました"
    if not any(r["sheet_name"] == sheet for r in sheet_rows):
        sheet_rows.append({"use": True, "sheet_name": sheet, "required": False})
    _save_rows(pattern, sheet_rows, field_rows)
    return jsonify(html=_build_html(pattern_id, sample_id), list_html=_list_html(), message=message)


@bp.post("/<int:pattern_id>/fields/<field_name>/delete")
def delete_field(pattern_id: int, field_name: str):
    pattern = _get_pattern(pattern_id)
    sheet_rows, field_rows = pattern_to_rows(pattern)
    rest = [r for r in field_rows if r["field_name"] != field_name]
    if len(rest) == len(field_rows):
        abort(404)
    kept = {r["sheet_name"] for r in rest if r["sheet_name"]}
    sheet_rows = [s for s in sheet_rows if not kept or s["sheet_name"] in kept]
    _save_rows(pattern, sheet_rows, rest)
    return jsonify(html=_build_html(pattern_id, request.form.get("sample", type=int)),
                   list_html=_list_html(), message="項目を削除しました")


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


def _save_rows(pattern: PatternDef, sheet_rows: list[dict], field_rows: list[dict], name: str | None = None) -> None:
    """状態は変えずに保存する（使用開始は［使用開始］を押したときだけ）。タイトル項目は自動で決める。"""
    meta = pattern_to_meta(pattern)
    if name:
        meta["name"] = name
    meta["title_fields"] = suggest_title_fields(field_rows)
    db.save_pattern(rows_to_pattern(pattern.id, meta, sheet_rows, field_rows), pattern.status)


# ---- 見本ファイルの追加・削除 -----------------------------------------------------------

@bp.post("/<int:pattern_id>/samples")
def add_samples(pattern_id: int):
    _get_pattern(pattern_id)
    saved, errors = _save_samples(pattern_id, request.files.getlist("samples"))
    if not saved:
        return jsonify(error=errors[0] if errors else "見本ファイルを追加できませんでした"), 400
    return jsonify(html=_build_html(pattern_id), list_html=_list_html(), errors=errors,
                   message=f"見本ファイルを{saved}件追加しました。この見本でも読めるか、下の読み取りテストで確かめてください")


@bp.post("/<int:pattern_id>/samples/<int:sample_id>/delete")
def delete_sample(pattern_id: int, sample_id: int):
    sample = db.get_sample(sample_id)
    if sample is None or sample["pattern_id"] != pattern_id:
        abort(404)
    try:
        remove_upload(sample["stored_path"])
    except UploadError:
        pass
    db.delete_sample(sample_id)
    return jsonify(html=_build_html(pattern_id), list_html=_list_html(), message="見本ファイルを削除しました")


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
        # 残すのは設定だけ。使用開始の時点で見本の Excel は消す（design.md 3.3）
        removed = _remove_samples(pattern_id)
        message = f"「{pattern.name}」の使用を開始しました。帳票取り込みの候補に出ます"
        if removed:
            message += "。見本のExcelはサーバーから消しました（設定だけ残ります）"
    else:
        message = f"「{pattern.name}」の使用を停止しました。帳票取り込みの候補に出なくなります"
    return jsonify(ok=True, status=status, html=_build_html(pattern_id), list_html=_list_html(), message=message)


def _remove_samples(pattern_id: int) -> int:
    """見本の Excel を消す（設定は残る）。戻り値は消した件数。"""
    removed = 0
    for sample in db.list_samples(pattern_id):
        try:
            remove_upload(sample["stored_path"])
        except UploadError:
            pass
        db.delete_sample(sample["id"])
        removed += 1
    return removed


@bp.post("/<int:pattern_id>/delete")
def delete(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    _remove_samples(pattern_id)
    db.delete_pattern(pattern_id)
    return jsonify(ok=True, list_html=_list_html(), message=f"帳票の種類「{pattern.name}」を削除しました")
