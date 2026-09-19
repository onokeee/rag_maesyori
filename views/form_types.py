"""帳票の種類（設定）: 見本ファイル → 候補の確認 → 読み取りテスト → 使用開始。

保存しても使用中にはしない。作成中の種類は、読み取りテストの画面で [使用開始] を押したときだけ使用中になる。
"""
from __future__ import annotations

from flask import Blueprint, abort, current_app, flash, redirect, render_template, request, url_for
from werkzeug.datastructures import MultiDict

from core.files import UploadError, precheck_excel, remove_upload, save_upload, upload_path
from excel.extractor import extract_document
from excel.workbook import load_workbook_info
from export.formats import build_markdown, markdown_filename
from models import database as db
from pattern.builder import merge_with_existing, suggest_rows, suggest_title_fields
from pattern.forms import parse_pattern_form, pattern_to_meta, pattern_to_rows, rows_to_pattern
from pattern.matcher import match_pattern
from pattern.model import DATA_TYPES, DIRECTIONS, IMAGE_PROCESSING, RAG_OUTPUTS, PatternDef
from views import safe_next
from views.forms import upload_error_text

bp = Blueprint("form_types", __name__, url_prefix="/settings/form-types")

STATUS_LABELS = {"draft": "作成中", "active": "使用中", "inactive": "停止中"}
CREATE_STEPS = ["見本ファイルを選ぶ", "読み取る項目の確認", "読み取りテスト", "使用開始"]


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
            precheck_excel(path, cfg.get("EXCEL_MAX_CELLS"))
            try:
                load_workbook_info(path)
            except Exception as exc:
                current_app.logger.warning("見本を読み込めませんでした: %s", exc.__class__.__name__)
                raise UploadError("Excelファイルとして読み込めませんでした") from exc
        except UploadError as exc:
            remove_upload(stored.stored_path)
            # flash に入るのでファイル名は出さず、選んだ順の位置で示す（design.md 3.3）
            errors.append(upload_error_text(storage, exc, position))
            continue
        db.add_sample(pattern_id, stored.file_name, stored.file_hash, stored.stored_path)
        saved += 1
    return saved, errors


def _sample_infos(pattern_id: int) -> tuple[list[dict], list]:
    samples, infos = [], []
    for sample in db.list_samples(pattern_id):
        try:
            infos.append(load_workbook_info(upload_path(sample["stored_path"])))
            samples.append(sample)
        except Exception:
            flash(f"見本ファイル {sample['file_name']} を読み込めませんでした", "error")
    return samples, infos


# ---- 一覧 ------------------------------------------------------------------------

@bp.get("/")
def index():
    return render_template("form_types/list.html", patterns=db.list_patterns(), status_labels=STATUS_LABELS)


# ---- 作成 ------------------------------------------------------------------------

@bp.get("/new")
def new():
    return render_template("form_types/new.html", steps=CREATE_STEPS, form={})


@bp.post("/new")
def create():
    name = request.form.get("name", "").strip()
    files = [f for f in request.files.getlist("samples") if f and f.filename]
    if not name or not files:
        flash("帳票の種類の名前と見本ファイルを指定してください", "error")
        return render_template("form_types/new.html", steps=CREATE_STEPS, form=request.form), 400
    pattern_id = db.create_pattern(name, request.form.get("version", "").strip() or "v1",
                                   request.form.get("description", "").strip())
    saved, errors = _save_samples(pattern_id, files)
    for message in errors:
        flash(message, "error")
    if not saved:
        db.delete_pattern(pattern_id)
        return render_template("form_types/new.html", steps=CREATE_STEPS, form=request.form), 400
    return redirect(url_for(".review", pattern_id=pattern_id))


@bp.get("/<int:pattern_id>/review")
def review(pattern_id: int):
    """見本ファイルから読み取る項目の候補を作る（登録済みの項目があればそこへ新しいラベルを足す）。"""
    pattern = _get_pattern(pattern_id)
    _, infos = _sample_infos(pattern_id)
    sheet_rows, field_rows = suggest_rows(infos)
    meta = pattern_to_meta(pattern)
    if pattern.fields:
        sheet_rows, field_rows = merge_with_existing(pattern, sheet_rows, field_rows)
    else:
        meta["title_fields"] = suggest_title_fields(field_rows)
    return _render_edit(pattern, meta, sheet_rows, field_rows, mode="review")


@bp.get("/<int:pattern_id>/edit")
def edit(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    sheet_rows, field_rows = pattern_to_rows(pattern)
    return _render_edit(pattern, pattern_to_meta(pattern), sheet_rows, field_rows, mode="edit")


def _with_title_order(form) -> MultiDict:
    """各項目行の「タイトルの順」（fields-N-title_order）から title_fields を作る。"""
    data = MultiDict(form)
    if not any(key.endswith("-title_order") for key in form.keys()):
        return data
    ordered = []
    for key in form.keys():
        parts = key.split("-")
        if len(parts) != 3 or parts[0] != "fields" or parts[2] != "title_order":
            continue
        raw = form.get(key, "").strip()
        name = form.get(f"fields-{parts[1]}-field_name", "").strip()
        if not raw or not name:
            continue
        try:
            order = float(raw)
        except ValueError:
            continue
        ordered.append((order, int(parts[1]) if parts[1].isdigit() else 0, name))
    data.setlist("title_fields", [name for _, _, name in sorted(ordered)])
    return data


@bp.post("/<int:pattern_id>/save")
def save(pattern_id: int):
    current = _get_pattern(pattern_id)
    form = _with_title_order(request.form)
    meta, sheet_rows, field_rows, errors = parse_pattern_form(form)
    mode = request.form.get("mode", "edit")
    if errors:
        for message in errors:
            flash(message, "error")
        draft = PatternDef(id=pattern_id, status=current.status, name=meta["name"], version=meta["version"],
                           description=meta["description"], image_processing=meta["image_processing"])
        return _render_edit(draft, meta, sheet_rows, field_rows, mode=mode), 400

    pattern = rows_to_pattern(pattern_id, meta, sheet_rows, field_rows)
    db.save_pattern(pattern, current.status)  # 状態は変えない（使用開始は読み取りテストの画面で）
    if current.status == "draft":
        flash("保存しました。まだ使用中ではありません。見本ファイルでの読み取り結果を確認して [使用開始] を押してください",
              "success")
    else:
        flash("保存しました。見本ファイルでの読み取り結果を確認してください", "success")
    return redirect(url_for(".test", pattern_id=pattern_id))


def _render_edit(pattern: PatternDef, meta: dict, sheet_rows: list[dict], field_rows: list[dict], mode: str):
    title_order = {name: i + 1 for i, name in enumerate(meta.get("title_fields") or [])}
    return render_template(
        "form_types/edit.html",
        steps=CREATE_STEPS,
        pattern=pattern,
        meta=meta,
        sheet_rows=sheet_rows,
        field_rows=field_rows,
        title_order=title_order,
        samples=db.list_samples(pattern.id) if pattern.id else [],
        mode=mode,
        data_types=DATA_TYPES,
        directions=DIRECTIONS,
        image_processing=IMAGE_PROCESSING,
        rag_outputs=RAG_OUTPUTS,
        status_labels=STATUS_LABELS,
        confirmed_count=_confirmed_count(pattern.id) if pattern.id else 0,
    )


# ---- 見本ファイル -------------------------------------------------------------------

@bp.post("/<int:pattern_id>/samples")
def add_samples(pattern_id: int):
    _get_pattern(pattern_id)
    saved, errors = _save_samples(pattern_id, request.files.getlist("samples"))
    for message in errors:
        flash(message, "error")
    if saved:
        flash(f"見本ファイルを{saved}件追加しました。「見本ファイルから候補を作り直す」で新しいラベルを取り込めます", "success")
    return redirect(safe_next(url_for(".edit", pattern_id=pattern_id)))


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
    # ファイル名は出さない（flash はブラウザのセッションクッキーに載る。design.md 3.3）
    flash("見本ファイルを削除しました", "info")
    return redirect(safe_next(url_for(".edit", pattern_id=pattern_id)))


# ---- 読み取りテスト ------------------------------------------------------------------

# excel/extractor.number_unit の単位不明の警告の書き出し
_NO_UNIT_WARNING = "単位が書かれていません"
TEST_NO_UNIT_WARNING = "単位が決まっていません。［項目を直す］でこの項目の単位（分・時間など）を入れてください"

@bp.get("/<int:pattern_id>/test")
def test(pattern_id: int):
    """登録した見本ファイル全件で読み取り、項目ごとの値と Markdown を見せる。"""
    pattern = _get_pattern(pattern_id)
    samples, infos = _sample_infos(pattern_id)
    results = []
    for sample, info in zip(samples, infos):
        match = match_pattern(info, pattern)
        sheets = match.sheet_names or info.sheet_names[:1]
        extraction = extract_document(info, pattern, sheets)
        doc = {"id": 0, "file_name": sample["file_name"], "file_hash": sample["file_hash"]}
        fields = {f["field_name"]: f for f in extraction["fields"]}
        for f in fields.values():
            if f["data_type"] == "number" and str(f.get("warning") or "").startswith(_NO_UNIT_WARNING):
                # 確認画面向けの「値に単位を付けて入力」はこの画面ではできないので、種類での直し方にする
                f["warning"] = TEST_NO_UNIT_WARNING
        results.append({
            "sample": sample,
            "match": match,
            "sheets": sheets,
            "fields": fields,
            "found": sum(1 for f in extraction["fields"] if f["value"] not in (None, "")),
            "missing_required": extraction["missing_required"],
            "markdown": build_markdown(doc, extraction) if pattern.fields else "",
            "file_name": markdown_filename(doc, extraction) if pattern.fields else "",
        })
    return render_template("form_types/test.html", steps=CREATE_STEPS, pattern=pattern, results=results,
                           status_labels=STATUS_LABELS, confirmed_count=_confirmed_count(pattern_id))


# ---- 使用開始・停止・削除 --------------------------------------------------------------

@bp.post("/<int:pattern_id>/status")
def change_status(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    status = request.form.get("status")
    if status not in ("active", "inactive"):
        abort(400)
    if status == "active" and not pattern.fields:
        flash("読み取る項目がないため使用開始できません。項目を設定してください", "error")
        return redirect(url_for(".edit", pattern_id=pattern_id))
    db.set_pattern_status(pattern_id, status)
    if status == "active":
        flash(f"「{pattern.name}」の使用を開始しました。帳票を取り込むときの候補に出ます", "success")
    else:
        flash(f"「{pattern.name}」の使用を停止しました。帳票を取り込むときの候補に出なくなります（確定済みの帳票はそのまま）",
              "info")
    return redirect(safe_next(url_for(".index")))


@bp.post("/<int:pattern_id>/delete")
def delete(pattern_id: int):
    pattern = _get_pattern(pattern_id)
    count = _confirmed_count(pattern_id)
    for sample in db.list_samples(pattern_id):
        try:
            remove_upload(sample["stored_path"])
        except UploadError:
            pass
    db.delete_pattern(pattern_id)
    flash(f"帳票の種類「{pattern.name}」を削除しました（確定済みの帳票{count}件は残っています）", "info")
    return redirect(url_for(".index"))
