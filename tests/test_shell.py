"""画面の骨組み: ホーム、ナビ、設定（AI接続・名寄せ辞書・LightRAG案内・一覧表の取り込み設定）、取り込み履歴。"""
import io
import json
import zipfile

import pytest
import yaml
from flask import render_template_string

from app import create_app
from models import database as db
from tests.conftest import make_config
from tests.fake_servers import OPENAI_KEY, FakeServer


# ---- 共通 -------------------------------------------------------------------------

@pytest.fixture
def fake():
    with FakeServer() as server:
        yield server


@pytest.fixture
def ai_app(tmp_path, fake):
    return create_app(make_config(tmp_path, OPENAI_BASE_URL=f"{fake.url}/v1", OPENAI_API_KEY="",
                                  OPENAI_MODELS=[], OPENAI_MODEL="gpt-test"))


@pytest.fixture
def ai_client(ai_app):
    return ai_app.test_client()


def _save_ai_settings(client, **overrides):
    form = {"models": ["gpt-test", "picky-model"], "default": "gpt-test", "api_key": OPENAI_KEY,
            "chat_url": "", "models_url": "", "add_models": ""}
    form.update(overrides)
    return client.post("/settings/ai", data=form, follow_redirects=True)


def _add_document(app, name, *, data=None, confirmed=None, title="", pattern_id=None, created_at=None):
    with app.app_context():
        doc_id = db.create_document(name, "0" * 64, f"documents/{name}", pattern_id)
        updates = {"data_json": data, "confirmed_json": confirmed, "title": title}
        if created_at:
            updates["created_at"] = created_at
        db.update_document(doc_id, **updates)
    return doc_id


def _extraction(value="EQ-001"):
    return json.dumps({
        "pattern": {"id": 1, "name": "設備修理報告書", "version": "v1"},
        "values": {"equipment_id": value},
        "fields": [{"field_name": "equipment_id", "display_name": "設備番号", "data_type": "string", "value": value,
                    "sheet": "修理報告書", "label_cell": "A4", "value_cell": "B4", "edited": False, "ai_filled": False}],
        "missing_required": [], "attachments": [], "sheets": ["修理報告書"],
    }, ensure_ascii=False)


# ---- ホーム・ナビ ----------------------------------------------------------------------

def test_home_renders_with_empty_db(client):
    page = client.get("/").get_data(as_text=True)
    assert "RAG用Markdown作成" in page
    assert "帳票を取り込む" in page and "一覧表を取り込む" in page
    assert "はじめに" in page  # 帳票の種類も一覧表の取り込み設定も0件
    assert "作業中の帳票はありません" in page
    for word in ("Excel帳票AI前処理", "テンプレート", "登録文書"):
        assert word not in page


def test_nav_links(client):
    page = client.get("/").get_data(as_text=True)
    for href in ('href="/"', 'href="/forms/new"', 'href="/tables/new"', 'href="/history/"',
                 'href="/settings/table-templates"',
                 'href="/settings/ai"', 'href="/settings/lightrag"'):
        assert href in page, href
    assert 'href="/settings/form-types' in page
    for label in ("取り込み履歴", "設定", "AI接続", "LightRAGへの入れ方", "一覧表の取り込み設定", "帳票の種類"):
        assert label in page
    # ナビの行き先がすべて開ける
    for url in ("/forms/new", "/tables/new", "/history/", "/settings/form-types/", "/settings/table-templates",
                "/settings/ai", "/settings/lightrag"):
        assert client.get(url).status_code == 200, url
    # 旧ルートは無い
    assert "名寄せ辞書" not in page  # 後回しにした機能は出さない
    for url in ("/documents", "/patterns", "/settings/models", "/settings/aliases"):
        assert client.get(url).status_code == 404, url


def test_home_lists_work_in_progress(app, client):
    _add_document(app, "作業中.xlsx", data=_extraction(), title="R-001 作業中")
    _add_document(app, "確定.xlsx", data=_extraction(), confirmed=_extraction(), title="R-002 確定")
    page = client.get("/").get_data(as_text=True)
    assert "続きを開く" in page and "確認中" in page
    assert "確定済み" in page


def test_ui_macros_render(app):
    source = """{% import "components/_ui.html" as ui %}
    {{ ui.steps(["ファイルを選ぶ", "確認", "完了"], 2) }}
    {% call ui.card("見出し") %}本文{% endcall %}
    {{ ui.badge("modified") }}{{ ui.badge("active") }}
    {{ ui.chips([{"label": "要確認", "count": 2, "kind": "warn"}, ("AIが入力", 0, "ai")]) }}
    {{ ui.empty_state("何もありません") }}
    {{ ui.confirm_delete_form("/x/delete", "削除します。元に戻せません") }}
    {{ ui.progress({"status": "running", "progress": {"done": 3, "total": 10}, "message": "処理中"}, url="/api/jobs/1") }}
    {{ ui.data_grid([{"index": 1, "cells": ["管理No", "設備"], "kind": "header"},
                     {"index": 2, "cells": ["TR-1", "CMP-101"], "kind": "data", "strike": True}] + [{"index": 3, "cells": ["x"] * 28}]) }}
    {{ ui.grid_legend() }}
    {{ ui.file_drop("file", ".xlsx,.xlsm") }}"""
    with app.test_request_context("/"):
        html = render_template_string(source)
    assert 'aria-current="step"' in html and "修正中" in html and "使用中" in html
    assert 'data-confirm="削除します。元に戻せません"' in html
    assert 'data-job-url="/api/jobs/1"' in html and 'aria-valuenow="30"' in html
    assert "row-header" in html and "row-strike" in html and ">AB<" in html
    assert 'accept=".xlsx,.xlsm"' in html


# ---- AI接続（tests/test_ai.py から移動） --------------------------------------------------------

def test_model_settings_saved_from_screen(ai_client, ai_app):
    page = ai_client.get("/settings/ai").get_data(as_text=True)
    assert "APIキーが設定されていません" in page and "AI未設定" in page

    page = _save_ai_settings(ai_client).get_data(as_text=True)
    assert "AI接続の設定を保存しました" in page
    assert OPENAI_KEY not in page  # キーの値は画面に出さない
    assert "この画面で保存したキーを使用中です" in page

    saved = yaml.safe_load((ai_app.config["DATA_DIR"] / "model_settings.yaml").read_text(encoding="utf-8"))
    assert saved["models"] == ["gpt-test", "picky-model"] and saved["api_key"] == OPENAI_KEY
    assert "chat_url" not in saved  # env と同じURLは上書きとして持たない

    page = ai_client.get("/settings/ai?refresh=1").get_data(as_text=True)
    assert "catalog-only" in page  # APIの models.list() から取得

    # ヘッダーのモデル切り替え
    assert ai_client.get("/api/models").get_json()["llm_ready"] is True
    assert ai_client.post("/api/models", json={"model": "picky-model"}).get_json()["current"] == "picky-model"
    assert ai_client.post("/api/models", json={"model": "unknown"}).status_code == 400


def test_model_settings_validation(ai_client):
    assert "APIキーの長さが不自然です" in _save_ai_settings(ai_client, api_key="short").get_data(as_text=True)
    page = _save_ai_settings(ai_client, chat_url="https://example.com/v1").get_data(as_text=True)
    assert "/chat/completions で終わるフルパス" in page
    page = _save_ai_settings(ai_client, default="not-in-list").get_data(as_text=True)
    assert "候補に入っていません" in page


def test_ai_connection_test(ai_client, fake):
    result = ai_client.post("/settings/ai/test").get_json()
    assert result["ok"] is False and "未設定" in result["steps"][0]["detail"]

    _save_ai_settings(ai_client)
    fake.chat_replies = ["OK"]
    result = ai_client.post("/settings/ai/test").get_json()
    assert result["ok"] is True
    assert [s["ok"] for s in result["steps"]] == [True, True]
    assert "3件のモデル" in result["steps"][0]["detail"] and "OK" in result["steps"][1]["detail"]
    assert fake.requests[-1]["path"] == "/v1/chat/completions"


# ---- LightRAG 案内 ---------------------------------------------------------------------

def test_lightrag_guide_and_yaml(client):
    page = client.get("/settings/lightrag").get_data(as_text=True)
    for text in ("GET /health", "SUMMARY_LANGUAGE=Japanese", "ENTITY_TYPE_PROMPT_FILE", "LIGHTRAG_PARSER",
                 "集計ファイル", "RAG投入用/", "1.4.x"):
        assert text in page, text

    res = client.get("/settings/lightrag/entity-types.yml")
    assert res.status_code == 200 and "attachment" in res.headers["Content-Disposition"]
    data = yaml.safe_load(res.get_data(as_text=True))
    guidance = data["entity_types_guidance"]
    for kind in ("設備", "部品", "不具合現象", "原因", "処置", "工程/ライン", "部署", "アラーム"):
        assert f"- {kind}:" in guidance, kind
    assert "{tuple_delimiter}" in data["entity_extraction_examples"][0]
    example = json.loads(data["entity_extraction_json_examples"][0])
    assert {e["type"] for e in example["entities"]} <= {"設備", "部品", "不具合現象", "原因", "処置", "工程/ライン", "部署", "アラーム"}
    assert example["relationships"]


# ---- 取り込み履歴 -----------------------------------------------------------------------

def test_history_renders_and_filters(app, client):
    page = client.get("/history/").get_data(as_text=True)
    assert "取り込んだ帳票はまだありません" in page
    assert "取り込んだ一覧表はまだありません" in client.get("/history/?tab=tables").get_data(as_text=True)

    _add_document(app, "読み取り前.xlsx", created_at="2026-08-01T09:00:00")
    _add_document(app, "作業中.xlsx", data=_extraction(), title="R-001 作業中", created_at="2026-09-01T09:00:00")
    _add_document(app, "確定.xlsx", data=_extraction(), confirmed=_extraction(), title="R-002 確定",
                  created_at="2026-09-10T09:00:00")

    page = client.get("/history/").get_data(as_text=True)
    assert "読み取り前.xlsx" in page and "作業中.xlsx" in page and "確定.xlsx" in page
    assert "続きを確認" in page and "Markdownをダウンロード" in page
    assert "選択した" in page and "件をダウンロード（zip）" in page

    page = client.get("/history/?state=confirmed").get_data(as_text=True)
    assert "確定.xlsx" in page and "作業中.xlsx" not in page

    page = client.get("/history/?q=作業").get_data(as_text=True)
    assert "作業中.xlsx" in page and "確定.xlsx" not in page

    page = client.get("/history/?date_from=2026-09-01&date_to=2026-09-05").get_data(as_text=True)
    assert "作業中.xlsx" in page and "確定.xlsx" not in page and "読み取り前.xlsx" not in page

    page = client.get("/history/?q=存在しない").get_data(as_text=True)
    assert "条件に合う帳票はありません" in page
    # 不正な値は無視
    assert client.get("/history/?state=bogus&page=x&pattern_id=y&date_from=2026-13-01").status_code == 200


def test_history_bulk_zip_only_confirmed(app, client):
    working = _add_document(app, "作業中.xlsx", data=_extraction(), title="作業中")
    first = _add_document(app, "確定1.xlsx", data=_extraction("EQ-001"), confirmed=_extraction("EQ-001"))
    # 修正中は確定済みの版（EQ-002）で出す
    second = _add_document(app, "確定2.xlsx", data=_extraction("EQ-999"), confirmed=_extraction("EQ-002"))

    res = client.post("/history/forms/download", data={"ids": [str(first), str(second), str(working)]})
    assert res.status_code == 200 and res.mimetype == "application/zip"
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        names = zf.namelist()
        texts = [zf.read(n).decode("utf-8") for n in names]
    assert len(names) == 2 and all(n.endswith(".md") for n in names)
    assert any("EQ-001" in t for t in texts) and any("EQ-002" in t for t in texts)
    assert not any("EQ-999" in t for t in texts)

    page = client.post("/history/forms/download", data={"ids": [str(working)]}, follow_redirects=True)
    assert "確定済みのものがありません" in page.get_data(as_text=True)
    page = client.post("/history/forms/download", data={}, follow_redirects=True)
    assert "ダウンロードする帳票を選んでください" in page.get_data(as_text=True)


def test_table_templates_page_without_tables(client):
    page = client.get("/settings/table-templates").get_data(as_text=True)
    assert "一覧表の取り込み設定はまだありません" in page
