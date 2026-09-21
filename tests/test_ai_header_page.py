"""ヘッダー右上の「AI接続」のパネルが画面に埋める URL（data-*-url）が、本物のルートを指すこと。

base.html は views.ai_header_ctx の `ai_urls`（辞書）を <dialog id="aiDialog"> の data 属性に埋め、
static/app.js の window.ragAiHeader がそれを fetch する。Jinja で `ai_urls.clear` と書くと、辞書の項目ではなく
dict.clear メソッドに解決されるので、［キーを消す］が `/forms/<built-in method clear of dict object …>` を叩いて
404 になっていた（2026-09-21、ブラウザでの確認で見つけた）。テストクライアントでルートを直接叩く検査では
気づけないので、ここでは画面に埋まった URL をそのまま使って確かめる。
"""
from __future__ import annotations

import re

import pytest
from flask import url_for

from tests.conftest import OPENAI_KEY, FakeServer, keep_scenario

URL_ATTRS = ("save", "test", "models", "clear")


@pytest.fixture
def fake():
    with FakeServer() as server:
        server.responder = keep_scenario
        yield server


def _panel_urls(html: str) -> dict[str, str]:
    """<dialog id="aiDialog" …> の data-save-url / data-test-url / data-models-url / data-clear-url。"""
    tag = re.search(r'<dialog id="aiDialog"[^>]*>', html)
    assert tag, "AI接続のパネルが無い"
    out = {}
    for name in URL_ATTRS:
        m = re.search(rf'data-{name}-url="([^"]*)"', tag.group(0))
        assert m, f"data-{name}-url が無い"
        out[name] = m.group(1)
    return out


def test_the_urls_embedded_in_every_screen_are_the_real_routes(app, client):
    """3つの画面のどれでも、パネルの4つの URL は url_for と同じ（辞書のメソッド名に化けていない）。"""
    with app.test_request_context("/"):
        expected = {"save": url_for("tables.save_ai_connection"), "test": url_for("tables.test_ai_connection"),
                    "models": url_for("tables.refresh_ai_models"), "clear": url_for("tables.clear_ai_key")}
    for page in ("/forms/new", "/tables/new", "/form-types/"):
        urls = _panel_urls(client.get(page).get_data(as_text=True))
        assert urls == expected, page
        for name, url in urls.items():
            assert "built-in" not in url and url.startswith("/tables/ai-connection"), (page, name, url)


def test_every_embedded_url_answers_with_json_not_404(client):
    """画面に埋まった URL をそのまま POST しても 404 にならず、JSON で答える（未設定なら未設定と言う）。"""
    urls = _panel_urls(client.get("/forms/new").get_data(as_text=True))
    for name, url in urls.items():
        res = client.post(url, json={})
        assert res.status_code in (200, 400), (name, url, res.status_code)
        body = res.get_json()
        assert isinstance(body, dict) and ("ok" in body or "error" in body), (name, body)


def test_clear_key_through_the_url_in_the_page_turns_the_header_off(client, fake):
    """［キーを消す］が押す URL（画面に埋まったもの）で、保存したキーが消えて「未接続」に戻る。"""
    urls = _panel_urls(client.get("/forms/new").get_data(as_text=True))
    res = client.post(urls["save"], json={"api_key": OPENAI_KEY, "chat_url": f"{fake.url}/v1/chat/completions",
                                          "models_url": f"{fake.url}/v1/models", "model": "gpt-test"})
    assert res.status_code == 200 and res.get_json()["status"]["state"] == "ok"

    res = client.post(urls["clear"], json={})
    assert res.status_code == 200, res.get_data(as_text=True)
    status = res.get_json()["status"]
    assert status["state"] == "off" and status["state_label"] == "未接続" and status["api_key_saved"] is False
    assert status["chat_url"] == f"{fake.url}/v1/chat/completions"   # 接続先は残る

    html = client.get("/tables/new").get_data(as_text=True)
    assert re.search(r'data-ai-header data-state="off"', html)
    assert "sk-… を貼り付け" in html and "保存済み（変えるときだけ入力）" not in html
