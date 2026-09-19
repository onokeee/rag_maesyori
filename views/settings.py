"""設定: AI接続（＋接続テスト）、保存先フォルダ、LightRAGへの入れ方、一覧表の取り込み設定の一覧、ヘッダーのモデル切り替えAPI。"""
from __future__ import annotations

import io
import json
import sqlite3
import time
import unicodedata
from pathlib import Path

from flask import (Blueprint, Response, abort, current_app, flash, jsonify, redirect, render_template, request,
                   send_file, url_for)

from core.naming import LIGHTRAG_HINT_RECORDS
from models import database
from services import llm, output_folder
from tables import store
from tables.spec import spec_from_dict, spec_to_dict, validate_spec

bp = Blueprint("settings", __name__)

ENTITY_TYPES_FILE = "entity_types_setsubi.yml"
ENTITY_TYPES_DOWNLOAD_NAME = "entity_types_setsubi.yml"

# 取り込み設定のJSONとして読めないときの文言（読めない理由は利用者の対処が同じなので1つにまとめる）
NOT_A_TEMPLATE_JSON = "このファイルは一覧表の取り込み設定として読み込めません（このアプリの［JSONで書き出す］で作ったファイルを選んでください）"
MODEL_PREF_WRITE_ERROR = "モデルの選択を保存できませんでした。設定ファイルを開いているプログラムを閉じてから、もう一度選んでください"


@bp.get("/settings/")
def index():
    return redirect(url_for(".ai"))


# ---- AI接続 ----------------------------------------------------------------------

@bp.get("/settings/ai")
def ai():
    catalog = []
    if request.args.get("refresh"):
        try:
            catalog = llm.model_catalog(refresh=True)
            flash(f"APIからモデル一覧を取得しました（{len(catalog)}件）", "success")
        except Exception as exc:
            flash(f"モデル一覧を取得できませんでした: {llm.friendly_error(exc)}", "error")
    return render_template("settings/ai.html", status=llm.admin_status(), catalog=catalog)


@bp.post("/settings/ai")
def save_ai():
    models = request.form.getlist("models") + request.form.get("add_models", "").splitlines()
    models = list(dict.fromkeys(m.strip() for m in models if m.strip()))
    try:
        llm.save_admin({
            "models": models,
            "default": request.form.get("default", ""),
            "chat_url": request.form.get("chat_url", ""),
            "models_url": request.form.get("models_url", ""),
            "api_key": request.form.get("api_key", ""),
            "api_key_clear": request.form.get("api_key_clear") == "on",
        })
    except ValueError as exc:
        flash(str(exc), "error")
    except OSError as exc:
        flash(f"設定ファイルに書き込めませんでした（{exc.strerror or exc.__class__.__name__}）。"
              "少し待ってから、もう一度保存してください", "error")
    else:
        flash("AI接続の設定を保存しました", "success")
    return redirect(url_for(".ai"))


@bp.post("/settings/ai/test")
def test_ai():
    """接続テスト: モデル一覧の取得と、1回の短いチャット。"""
    steps = []
    if not llm.is_configured():
        return jsonify({"ok": False, "steps": [{"name": "設定", "ok": False,
                                                "detail": "APIキーまたは接続先が未設定です。"}]})
    started = time.monotonic()
    try:
        names = llm.fetch_api_models(refresh=True)
        steps.append({"name": "モデル一覧の取得", "ok": True,
                      "detail": f"{len(names)}件のモデルが見つかりました（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": "モデル一覧の取得", "ok": False, "detail": llm.friendly_error(exc)})

    model = llm.current_model()
    started = time.monotonic()
    try:
        # AI整形のジョブと同じ呼び出し口（再試行なし・明示のタイムアウト）で確かめる
        result = llm.chat_raw(llm.job_client_settings(),
                              [{"role": "user", "content": "接続テストです。「OK」とだけ返してください。"}], max_tokens=20)
        reply = (result.text or "").strip()
        steps.append({"name": f"チャット（{model}）", "ok": True,
                      "detail": f"応答あり: {reply[:40] or '（本文なし）'}（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": f"チャット（{model}）", "ok": False, "detail": llm.friendly_error(exc)})
    # チャットが通れば AI 整形は使える（モデル一覧APIが無い互換サーバもある）
    return jsonify({"ok": steps[-1]["ok"], "steps": steps})


def _ms(started: float) -> int:
    return int((time.monotonic() - started) * 1000)


@bp.get("/api/models")
def api_models():
    return jsonify({"current": llm.current_model(), "models": llm.available(), "llm_ready": llm.is_configured()})


@bp.post("/api/models")
def api_choose_model():
    body = request.get_json(silent=True) or {}
    try:
        llm.choose_model(body.get("model"))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except OSError:
        # 設定ファイルが他のプログラムに開かれたままなどで書けない（settings_store が再試行したあと）
        current_app.logger.exception("モデルの選択を保存できませんでした")
        return jsonify({"error": MODEL_PREF_WRITE_ERROR}), 500
    return jsonify({"current": llm.current_model(), "models": llm.available(), "llm_ready": llm.is_configured()})


# ---- 保存先フォルダ -----------------------------------------------------------------------
# 作った Markdown をダウンロードせず、選んだフォルダ（LightRAG の INPUT_DIR など）に直接書く（services/output_folder.py）

def _output_page(settings: dict, errors: list[str] | None = None, status: int = 200):
    """settings: 入力欄に出す値。「いまの設定」には、断られたときも保存済みの設定（saved）を出す。"""
    saved = output_folder.load()
    return render_template("settings/output.html", settings=settings, saved=saved, errors=errors or [],
                           policies=output_folder.CONFLICT_POLICIES, admin_subdir=output_folder.ADMIN_SUBDIR,
                           settings_file=str(Path(current_app.config["DATA_DIR"]) / output_folder.SETTINGS_FILE),
                           path_warning=output_folder.path_warning(saved.get("folder"))), status


@bp.get("/settings/output")
def output():
    return _output_page(output_folder.load())


@bp.post("/settings/output")
def save_output():
    form = {"folder": output_folder.clean_folder_text(request.form.get("folder")),
            "save_admin": request.form.get("save_admin") == "on",
            "on_conflict": request.form.get("on_conflict", output_folder.DEFAULT_POLICY)}
    try:
        output_folder.save_settings(form["folder"], form["save_admin"], form["on_conflict"])
    except ValueError as exc:
        # 入力した値を残したまま、理由を出す（リダイレクトすると打った値が消える）
        return _output_page(form, [str(exc)], 400)
    except OSError as exc:
        # 設定ファイルをほかのプログラムが掴んでいた・権限が無い（500 にせず、打った値を残して知らせる）
        return _output_page(form, [f"設定ファイルに書き込めませんでした（{exc.strerror or exc.__class__.__name__}）。"
                                   "少し待ってから、もう一度［設定を保存］を押してください"], 500)
    flash("保存先フォルダの設定を保存しました" if form["folder"] else
          "保存先フォルダを空にしました（［保存先フォルダに保存］のボタンは出なくなります）", "success")
    return redirect(url_for(".output"))


@bp.post("/settings/output/check")
def check_output():
    """[フォルダを確かめる]: 保存済みの（または入力中の）フォルダを、保存するときと同じ手順で確かめ直す。"""
    body = request.get_json(silent=True) or {}
    text = output_folder.clean_folder_text(body.get("folder") or output_folder.configured_folder())
    if not text:
        return jsonify({"ok": False, "folder": "", "errors": ["保存先フォルダが入力されていません"]})
    resolved, errors = output_folder.check_folder(text)
    result = {"ok": not errors, "folder": str(resolved) if resolved else text, "errors": errors}
    if not errors:
        result.update(output_folder.folder_report(resolved))
        result["warning"] = output_folder.path_warning(resolved)
    return jsonify(result)


# ---- LightRAGへの入れ方 -----------------------------------------------------------------

def _entity_types_path() -> Path:
    return Path(current_app.root_path) / "templates" / "settings" / ENTITY_TYPES_FILE


def _hint_params(hint: str) -> dict[str, str]:
    """'legacy-R(chunk_ts=1500,chunk_ol=0)' → {'chunk_ts': '1500', 'chunk_ol': '0'}（説明の表示用）。"""
    inner = hint.partition("(")[2].rstrip(")")
    return dict(kv.split("=", 1) for kv in inner.split(",") if "=" in kv)


@bp.get("/settings/lightrag")
def lightrag():
    params = _hint_params(LIGHTRAG_HINT_RECORDS)
    return render_template("settings/lightrag.html",
                           entity_yaml=_entity_types_path().read_text(encoding="utf-8"),
                           download_name=ENTITY_TYPES_DOWNLOAD_NAME,
                           hint_records=LIGHTRAG_HINT_RECORDS,
                           hint_chunk_ts=params.get("chunk_ts", ""),
                           hint_chunk_ol=params.get("chunk_ol", "0"))


@bp.get("/settings/lightrag/entity-types.yml")
def entity_types_yaml():
    body = _entity_types_path().read_bytes()
    return Response(body, mimetype="application/x-yaml",
                    headers={"Content-Disposition": f"attachment; filename={ENTITY_TYPES_DOWNLOAD_NAME}"})


# ---- 一覧表の取り込み設定（一覧・JSON書き出し/読み込み。編集と削除は tables blueprint） ----------------------------

def _clean(text, limit: int = 200) -> str:
    return " ".join(unicodedata.normalize("NFKC", str(text or "")).split())[:limit]


@bp.get("/settings/table-templates")
def table_templates():
    return render_template("settings/table_templates.html", templates=store.list_templates())


@bp.get("/settings/table-templates/<int:template_id>/export.json")
def export_table_template(template_id: int):
    template = store.get_template(template_id)
    if template is None or template.get("spec") is None:
        abort(404)
    data = {"format": "rag_maesyori.table_template", "name": template["name"],
            "description": template["description"] or "", "version": template["version"],
            "spec": spec_to_dict(template["spec"])}
    body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    # 他のダウンロードと同じく send_file に任せる（filename* だけでなく ASCII の filename= も付く）
    return send_file(io.BytesIO(body), mimetype="application/json", as_attachment=True,
                     download_name=f"{template['name']}.json")


@bp.post("/settings/table-templates/import")
def import_table_template():
    back = url_for(".table_templates")
    storage = request.files.get("file")
    if storage is None or not storage.filename:
        flash("読み込むJSONファイルを選んでください", "error")
        return redirect(back)
    try:
        data = json.loads(storage.read(5 * 1024 * 1024).decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError, RecursionError):  # 深く入れ子のJSONは RecursionError になる
        flash(NOT_A_TEMPLATE_JSON, "error")
        return redirect(back)
    spec_dict = data.get("spec") if isinstance(data, dict) else None
    if not isinstance(spec_dict, dict):
        flash(NOT_A_TEMPLATE_JSON, "error")
        return redirect(back)
    try:
        spec = spec_from_dict(spec_dict)
    except Exception:
        # 例外の本文には内部の変数名が出るので画面には出さず、ログにだけ残す
        current_app.logger.exception("取り込み設定のJSONを読み込めませんでした（%s）", storage.filename)
        flash(NOT_A_TEMPLATE_JSON, "error")
        return redirect(back)
    name = _clean(request.form.get("name") or data.get("name") or spec.name, 100)
    spec.name = name
    try:
        errors = validate_spec(spec)
    except Exception:
        # 手で書き換えた JSON の想定外の値で落ちても 500 にしない（中身はログにだけ残す）
        current_app.logger.exception("取り込み設定のJSONを確かめられませんでした（%s）", storage.filename)
        flash(NOT_A_TEMPLATE_JSON, "error")
        return redirect(back)
    if errors:
        flash("設定の内容に問題があります: " + " / ".join(errors[:5]), "error")
        return redirect(back)
    try:
        store.create_template(name, spec, _clean(data.get("description"), 500), note=f"読み込み: {storage.filename}")
    except sqlite3.IntegrityError:
        database.get_db().rollback()
        flash(f"同じ名前の取り込み設定「{name}」があります。名前を変えて読み込んでください", "error")
        return redirect(back)
    flash(f"取り込み設定「{name}」を読み込みました", "success")
    return redirect(back)
