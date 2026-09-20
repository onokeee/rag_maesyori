"""画面の骨組み: 3つの画面へのナビ、エラー画面、Host の確認、キャッシュ、他サイトからの書き込み。

画面は「帳票取り込み・表の取り込み・帳票登録」の3つだけになった（2026-09-20 の作り直し）。
ホーム画面・設定の画面・取り込み履歴・作業中の一覧は無い。AI接続は「表の取り込み」画面の
AI整形の段の中（/tables/ai-connection）へ移した。
"""
import io
import os
import time

import pytest
import yaml
from flask import render_template_string

from app import create_app
from models import database as db
from tests.conftest import add_confirmed_document, make_config
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
    """AI接続の保存（「表の取り込み」画面の AI整形の段が送る fetch と同じ）。"""
    body = {"models": ["gpt-test", "picky-model"], "default": "gpt-test", "api_key": OPENAI_KEY,
            "chat_url": "", "models_url": "", "add_models": ""}
    body.update(overrides)
    return client.post("/tables/ai-connection", json=body)


# ---- ナビ（3つの画面しかない） ---------------------------------------------------------------

def test_root_goes_to_the_form_import_screen(client):
    """ホーム画面は無い。「/」（古いブックマーク）は帳票取り込みへ送る。"""
    res = client.get("/")
    assert res.status_code == 302 and res.headers["Location"] == "/forms/new"


def test_nav_has_only_the_three_screens(client):
    page = client.get("/forms/new").get_data(as_text=True)
    for href, label in (('href="/forms/new"', "帳票取り込み"), ('href="/tables/new"', "表の取り込み"),
                        ('href="/form-types/"', "帳票登録")):
        assert href in page and label in page, href
    for url in ("/forms/new", "/tables/new", "/form-types/"):
        assert client.get(url).status_code == 200, url
    # 無くした画面の言葉はヘッダーに出さない
    for word in ("設定", "取り込み履歴", "LightRAGへの入れ方", "保存先フォルダ", "名寄せ辞書"):
        assert word not in page, word


def test_the_removed_screens_are_gone(client):
    """設定・履歴・ホームの一覧・LightRAG案内・保存先フォルダの URL は残っていない。"""
    for url in ("/home", "/history/", "/documents", "/patterns",
                "/settings/ai", "/settings/output", "/settings/lightrag", "/settings/lightrag/entity-types.yml",
                "/settings/table-templates", "/settings/form-types/", "/settings/models", "/settings/aliases",
                "/api/models", "/forms/1/ai-classify", "/forms/1/ai-fill"):
        assert client.get(url).status_code == 404, url


def test_old_deep_links_come_back_to_the_one_page_screens(client):
    """作り直す前の URL（お気に入り・開いたままの画面）は、その画面の1枚ページへ送る。"""
    for url, target in (("/forms/1", "/forms/"), ("/forms/1/done", "/forms/"),
                        ("/form-types/new", "/form-types/"), ("/form-types/1/edit", "/form-types/")):
        res = client.get(url)
        assert res.status_code == 302 and res.headers["Location"] == target, url


def test_ui_macros_render(app):
    source = """{% import "components/_ui.html" as ui %}
    {% call ui.step("sheet", 2, "帳票の種類とシート", done=True, summary="設備修理報告書") %}本文{% endcall %}
    {% call ui.card("見出し") %}本文{% endcall %}
    {{ ui.badge("modified") }}{{ ui.badge("active") }}
    {{ ui.chips([{"label": "要確認", "count": 2, "kind": "warn"}, ("AIが入力", 0, "ai")]) }}
    {{ ui.empty_state("何もありません") }}
    {{ ui.progress({"status": "running", "progress": {"done": 3, "total": 10}, "message": "処理中"}, url="/api/jobs/1") }}
    {{ ui.data_grid([{"index": 1, "cells": ["管理No", "設備"], "kind": "header"},
                     {"index": 2, "cells": ["TR-1", "CMP-101"], "kind": "data", "strike": True}] + [{"index": 3, "cells": ["x"] * 28}]) }}
    {{ ui.grid_legend() }}
    {{ ui.file_drop("file", ".xlsx,.xlsm") }}"""
    with app.test_request_context("/forms/new"):
        html = render_template_string(source)
    assert 'data-step="sheet"' in html and "is-done" in html and "設備修理報告書" in html
    assert "修正中" in html and "使用中" in html
    assert 'data-job-url="/api/jobs/1"' in html and 'aria-valuenow="30"' in html
    assert "row-header" in html and "row-strike" in html and ">AB<" in html
    assert 'accept=".xlsx,.xlsm"' in html


# ---- AI接続（設定画面は無い。「表の取り込み」画面の AI整形の段の中） -----------------------------------

def test_ai_connection_is_saved_from_the_table_screen(ai_client, ai_app):
    res = _save_ai_settings(ai_client)
    assert res.status_code == 200
    body = res.get_json()
    assert body["ok"] is True and "AI接続の設定を保存しました" in body["message"]
    assert OPENAI_KEY not in res.get_data(as_text=True)   # キーの値は返さない

    saved = yaml.safe_load((ai_app.config["DATA_DIR"] / "model_settings.yaml").read_text(encoding="utf-8"))
    assert saved["models"] == ["gpt-test", "picky-model"] and saved["api_key"] == OPENAI_KEY
    assert "chat_url" not in saved   # env と同じURLは上書きとして持たない

    catalog = ai_client.post("/tables/ai-connection/models").get_json()
    assert catalog["ok"] is True and "catalog-only" in catalog["models"]   # APIの models.list() から取得


def test_ai_connection_validation(ai_client):
    assert "APIキーの長さが不自然です" in _save_ai_settings(ai_client, api_key="short").get_json()["error"]
    assert "/chat/completions で終わるフルパス" in \
        _save_ai_settings(ai_client, chat_url="https://example.com/v1").get_json()["error"]
    assert "候補に入っていません" in _save_ai_settings(ai_client, default="not-in-list").get_json()["error"]


def test_ai_connection_test(ai_client, fake):
    result = ai_client.post("/tables/ai-connection/test").get_json()
    assert result["ok"] is False and "未設定" in result["steps"][0]["detail"]

    _save_ai_settings(ai_client)
    fake.chat_replies = ["OK"]
    result = ai_client.post("/tables/ai-connection/test").get_json()
    assert result["ok"] is True
    assert [s["ok"] for s in result["steps"]] == [True, True]
    assert "3件のモデル" in result["steps"][0]["detail"] and "OK" in result["steps"][1]["detail"]
    assert fake.requests[-1]["path"] == "/v1/chat/completions"


# ---- エラー画面・Host の確認 ----------------------------------------------------------

def test_not_found_page_is_japanese(client):
    """消した帳票の URL を開いても、英語の既定ページではなく日本語の案内と戻り道を出す。"""
    for url in ("/forms/999/original", "/tables/imports/999/panel/preview", "/form-types/999/panel"):
        res = client.get(url)
        assert res.status_code == 404, url
        page = res.get_data(as_text=True)
        assert "見つかりません" in page, url
        assert 'href="/forms/new"' in page, url


def test_not_found_answers_json_when_the_screen_asked_for_json(client):
    """開いたままの画面が、消えた帳票へ途中保存・プレビューを送ったとき（app.js の postJson）。"""
    res = client.post("/forms/999/draft", json={"values": {}})
    assert res.status_code == 404 and "このPCに残っていません" in res.get_json()["error"]


def test_method_not_allowed_page_is_japanese(client):
    """送信専用の URL をアドレス欄から開いても、Werkzeug の英語の画面を出さない（design.md 2）。"""
    res = client.get("/forms/1/delete")
    assert res.status_code == 405
    page = res.get_data(as_text=True)
    assert "Method Not Allowed" not in page and "取り込みの画面から開き直してください" in page
    assert 'href="/forms/new"' in page
    res = client.get("/forms/1/delete", headers={"Accept": "application/json"})
    assert res.status_code == 405 and "取り込みの画面から" in res.get_json()["error"]


def test_bad_request_page_is_japanese(client):
    res = client.get("/forms/new", headers={"Host": "evil.example"})
    assert res.status_code == 400 and "Bad Request" not in res.get_data(as_text=True)


def test_other_host_is_refused(client):
    """このPC以外の名前で届いたリクエストは断る（DNSリバインディング対策）。"""
    assert client.get("/forms/new", headers={"Host": "127.0.0.1:5000"}).status_code == 200
    assert client.get("/forms/new", headers={"Host": "localhost:5000"}).status_code == 200
    assert client.get("/forms/new", headers={"Host": "evil.example"}).status_code == 400
    assert client.get("/tables/new", headers={"Host": "evil.example:5000"}).status_code == 400


def test_a_refused_host_is_told_which_address_works(client):
    """別の名前（PC名・hosts の別名）で開かれたときは、同じアドレスへ戻さず、開けるアドレスへ案内する。"""
    res = client.get("/forms/new", headers={"Host": "mypc:5123"})
    page = res.get_data(as_text=True)
    assert res.status_code == 400
    assert "http://127.0.0.1:5123/" in page and 'href="http://127.0.0.1:5123/"' in page
    assert "このページは再読み込みや古いアドレスからは開けません" not in page   # 断られた理由に合う文だけ出す
    res = client.post("/forms/1/delete", headers={"Host": "mypc", "Accept": "application/json"})
    assert res.status_code == 400 and "http://127.0.0.1:5000/" in res.get_json()["error"]


def test_pages_and_json_are_not_kept_in_the_browser_cache(app, client):
    """取り込んだ値の載る画面・JSON はブラウザに保存させない（ダウンロードで消したあと、戻るで出さない）。"""
    doc_id, _path = add_confirmed_document(app, "キャッシュ.xlsx")
    for url, code in (("/forms/new", 200), ("/tables/new", 200), (f"/forms/{doc_id}/review", 200),
                      ("/forms/999/original", 404), ("/api/jobs/999", 404)):
        res = client.get(url)
        assert res.status_code == code, url
        assert res.headers.get("Cache-Control") == "no-store", url
    assert client.get("/forms/new", headers={"Host": "evil.example"}).headers.get("Cache-Control") == "no-store"
    static = client.get("/static/app.js")
    assert static.status_code == 200 and "no-store" not in (static.headers.get("Cache-Control") or "")


def test_no_response_can_be_shown_inside_another_sites_frame(client):
    """ほかのサイトの iframe に入れて削除ボタンなどを押させない（クリックジャッキング）。静的ファイル・エラーも同じ。"""
    responses = {url: client.get(url) for url in ("/forms/new", "/tables/new", "/static/app.js", "/forms/999/original")}
    responses["他のホスト名"] = client.get("/forms/new", headers={"Host": "evil.example"})
    responses["他のサイトからの書き込み"] = client.post("/forms/1/delete", headers={"Origin": "http://evil.example"})
    for name, res in responses.items():
        assert res.headers.get("X-Frame-Options") == "DENY", name
        assert res.headers.get("Content-Security-Policy") == "frame-ancestors 'none'", name
    assert responses["/forms/new"].status_code == 200 and responses["/tables/new"].status_code == 200
    assert responses["他のサイトからの書き込み"].status_code == 403


def test_server_error_page_is_japanese(tmp_path):
    """想定外の例外でも、英語の既定ページではなく日本語の案内と戻り道を出す。"""
    app = create_app(make_config(tmp_path, PROPAGATE_EXCEPTIONS=False))

    @app.route("/_raise_for_test")
    def _raise_for_test():
        raise RuntimeError("テスト用の例外")

    res = app.test_client().get("/_raise_for_test")
    assert res.status_code == 500
    page = res.get_data(as_text=True)
    assert "エラーが発生しました" in page and 'href="/forms/new"' in page


def test_too_large_upload_page_is_japanese(tmp_path):
    """MAX_CONTENT_LENGTH を超えた送信は Werkzeug の英語の画面ではなく、日本語の案内と戻り道。"""
    small = create_app(make_config(tmp_path, MAX_CONTENT_LENGTH=1024))
    res = small.test_client().post("/forms/upload", data={"file": (io.BytesIO(b"x" * 5000), "大きい.xlsx")},
                                   content_type="multipart/form-data")
    body = res.get_data(as_text=True)
    assert res.status_code == 413
    assert "ファイルが大きすぎます" in body and "合計 1KB まで" in body and "分けて" in body
    assert "Request Entity Too Large" not in body and 'href="/forms/new"' in body
    res = small.test_client().post("/forms/upload", data=b"x" * 5000, headers={"Accept": "application/json"},
                                   content_type="application/json")
    assert res.status_code == 413 and "大きすぎます" in res.get_json()["error"]


# ---- 起動時の片付け ---------------------------------------------------------------------

def test_startup_removes_orphan_uploads(tmp_path):
    """起動時に、DB から参照されていない取り込み途中の残骸だけを片付ける。"""
    config = make_config(tmp_path)
    app = create_app(config)
    documents = app.config["UPLOAD_DIR"] / "documents"
    documents.mkdir(parents=True, exist_ok=True)
    orphan = documents / f"{'a' * 32}.xlsx"
    orphan.write_bytes(b"broken")
    old = time.time() - 3600
    os.utime(orphan, (old, old))   # 取り込みの途中ではない（できてから時間がたった）残骸
    fresh = documents / f"{'c' * 32}.xlsx"
    fresh.write_bytes(b"uploading")   # 別に起動しているアプリが保存したばかり（DB の行を作る前）
    used = documents / f"{'b' * 32}.xlsx"
    used.write_bytes(b"used")
    with app.app_context():
        db.create_document("点検表.xlsx", "0" * 64, f"documents/{used.name}")

    create_app(config)   # 起動し直す

    assert not orphan.exists()
    assert used.exists()
    assert fresh.exists()   # できたばかりのファイルは消さない


# ---- 他サイトからの書き込みを断る ------------------------------------------------------------

CROSS_SITE = {"Origin": "https://evil.example", "Sec-Fetch-Site": "cross-site"}


def test_cross_site_post_is_refused(ai_client):
    res = ai_client.post("/tables/ai-connection", json={"chat_url": "http://evil.example/v1/chat/completions"},
                         headers=CROSS_SITE)
    assert res.status_code == 403
    assert ai_client.post("/tables/ai-connection/test", headers=CROSS_SITE).status_code == 403
    # Origin だけ・Sec-Fetch-Site だけでも断る
    assert ai_client.post("/tables/ai-connection/test", headers={"Origin": "https://evil.example"}).status_code == 403
    assert ai_client.post("/tables/ai-connection/test", headers={"Sec-Fetch-Site": "cross-site"}).status_code == 403


def test_same_origin_post_still_works(ai_client):
    """この画面から送られた（同じサイトの）書き込みは通す。"""
    res = ai_client.post("/tables/ai-connection",
                         json={"models": ["gpt-test"], "default": "gpt-test", "api_key": ""},
                         headers={"Origin": "http://localhost", "Sec-Fetch-Site": "same-origin"})
    assert res.status_code == 200 and res.get_json()["ok"] is True


def test_cross_site_get_is_allowed(client):
    assert client.get("/forms/new", headers=CROSS_SITE).status_code == 200


def test_refused_write_shows_a_japanese_page(ai_client):
    """断ったことが日本語で分かり、取り込みの画面に戻れる（Flask の英語の403ページを出さない）。"""
    res = ai_client.post("/tables/ai-connection/test", headers=CROSS_SITE)
    assert res.status_code == 403
    page = res.get_data(as_text=True)
    assert "ほかのサイトのページから送られてきた操作" in page
    assert "データは変わっていません" in page and 'href="/forms/new"' in page
    assert "Forbidden" not in page


def test_refused_write_answers_json_when_the_screen_asked_for_json(ai_client):
    """画面の JSON 送信（app.js の postJson）には JSON で返す。HTML だと理由が出ない。"""
    res = ai_client.post("/tables/ai-connection/test", headers={**CROSS_SITE, "Accept": "application/json"})
    assert res.status_code == 403
    assert "ほかのサイト" in res.get_json()["error"]
