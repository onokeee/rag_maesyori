"""設定: AI接続（＋接続テスト）、LightRAGへの入れ方、一覧表の取り込み設定の一覧、ヘッダーのモデル切り替えAPI。"""
from __future__ import annotations

import json
import sqlite3
import time
import unicodedata
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, Response, abort, current_app, flash, jsonify, redirect, render_template, request, url_for

from models import database
from services import llm
from tables import store
from tables.spec import spec_from_dict, spec_to_dict, validate_spec

bp = Blueprint("settings", __name__)

ENTITY_TYPES_FILE = "entity_types_setsubi.yml"
ENTITY_TYPES_DOWNLOAD_NAME = "entity_types_setsubi.yml"


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
    return jsonify({"current": llm.current_model(), "models": llm.available(), "llm_ready": llm.is_configured()})


# ---- LightRAGへの入れ方 -----------------------------------------------------------------

def _entity_types_path() -> Path:
    return Path(current_app.root_path) / "templates" / "settings" / ENTITY_TYPES_FILE


@bp.get("/settings/lightrag")
def lightrag():
    return render_template("settings/lightrag.html",
                           entity_yaml=_entity_types_path().read_text(encoding="utf-8"),
                           download_name=ENTITY_TYPES_DOWNLOAD_NAME)


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
    return Response(body, mimetype="application/json",
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(template['name'])}.json"})


@bp.post("/settings/table-templates/import")
def import_table_template():
    back = url_for(".table_templates")
    storage = request.files.get("file")
    if storage is None or not storage.filename:
        flash("読み込むJSONファイルを選んでください", "error")
        return redirect(back)
    try:
        data = json.loads(storage.read(5 * 1024 * 1024).decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError):
        flash("JSONとして読み込めませんでした（このアプリで書き出したファイルを選んでください）", "error")
        return redirect(back)
    spec_dict = data.get("spec") if isinstance(data, dict) else None
    if not isinstance(spec_dict, dict):
        flash("一覧表の取り込み設定のJSONではありません（spec がありません）", "error")
        return redirect(back)
    try:
        spec = spec_from_dict(spec_dict)
    except Exception as exc:
        flash(f"設定の内容が正しくありません（{exc}）", "error")
        return redirect(back)
    name = _clean(request.form.get("name") or data.get("name") or spec.name, 100)
    spec.name = name
    errors = validate_spec(spec)
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
