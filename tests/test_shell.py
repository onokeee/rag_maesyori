"""画面の骨組み: ホーム、ナビ、設定（AI接続・LightRAG案内・一覧表の取り込み設定）。"""
import io
import json
import os
import time

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
    for href in ('href="/"', 'href="/forms/new"', 'href="/tables/new"',
                 'href="/settings/table-templates"',
                 'href="/settings/ai"', 'href="/settings/lightrag"'):
        assert href in page, href
    assert 'href="/settings/form-types' in page
    for label in ("設定", "AI接続", "LightRAGへの入れ方", "一覧表の取り込み設定", "帳票の種類"):
        assert label in page
    # ナビの行き先がすべて開ける
    for url in ("/forms/new", "/tables/new", "/settings/form-types/", "/settings/table-templates",
                "/settings/ai", "/settings/lightrag"):
        assert client.get(url).status_code == 200, url
    # 旧ルートは無い（取り込み履歴はデータを残さないので画面ごと無くした）
    assert "名寄せ辞書" not in page  # 後回しにした機能は出さない
    assert "取り込み履歴" not in page
    for url in ("/documents", "/patterns", "/settings/models", "/settings/aliases", "/history/"):
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
    # doc_id の説明が2か所で食い違わない（ヒントを外してから文書IDを決めるので同じ文書）
    assert "文書ID（doc_id）は同じ" in page and "同じ文書" in page
    assert "別の文書</strong>として扱われます" not in page

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


# ---- ホームの一覧（取り込み履歴の画面は無い。データを残さないので、まだ残っているものだけを出す） ----

def test_home_lists_only_what_is_left(app, client):
    page = client.get("/").get_data(as_text=True)
    assert "作業中の帳票はありません" in page and "ダウンロード待ちの帳票はありません" in page
    assert "作業中の一覧表はありません" in page and "ダウンロード待ちの一覧表はありません" in page

    _add_document(app, "読み取り前.xlsx", created_at="2026-08-01T09:00:00")
    _add_document(app, "作業中.xlsx", data=_extraction(), title="R-001 作業中", created_at="2026-09-01T09:00:00")
    _add_document(app, "確定.xlsx", data=_extraction(), confirmed=_extraction(), title="R-002 確定",
                  created_at="2026-09-10T09:00:00")

    page = client.get("/").get_data(as_text=True)
    assert "読み取り前.xlsx" in page and "R-001 作業中" in page and "R-002 確定" in page
    assert "続きを開く" in page and ".mdをダウンロード" in page
    # 作業中のものは、この画面から消せる
    assert "/delete" in page and "元に戻せません" in page
    # ダウンロードするとデータが消えることを画面で知らせる
    assert "ダウンロードすると" in page and "もう一度ダウンロードすることはできません" in page


def test_table_templates_page_without_tables(client):
    page = client.get("/settings/table-templates").get_data(as_text=True)
    assert "一覧表の取り込み設定はまだありません" in page


# ---- エラー画面・Host の確認 ----------------------------------------------------------

def test_not_found_page_is_japanese(client):
    """消した帳票の URL を開いても、英語の既定ページではなく日本語の案内と戻り道を出す。"""
    for url in ("/forms/999", "/settings/form-types/999/edit", "/tables/imports/999/preview"):
        res = client.get(url)
        assert res.status_code == 404, url
        page = res.get_data(as_text=True)
        assert "見つかりませんでした" in page, url
        assert 'href="/"' in page, url


def test_method_not_allowed_page_is_japanese(client):
    """送信専用の URL をアドレス欄から開いても、Werkzeug の英語の画面を出さない（design.md 2）。"""
    res = client.get("/forms/1/ai-classify")
    assert res.status_code == 405
    page = res.get_data(as_text=True)
    assert "Method Not Allowed" not in page and "ホームから開き直してください" in page and 'href="/"' in page
    res = client.get("/forms/1/ai-classify", headers={"Accept": "application/json"})
    assert res.status_code == 405 and "ホームから" in res.get_json()["error"]


def test_bad_request_page_is_japanese(client):
    res = client.get("/", headers={"Host": "evil.example"})
    assert res.status_code == 400 and "Bad Request" not in res.get_data(as_text=True)


def test_other_host_is_refused(client):
    """このPC以外の名前で届いたリクエストは断る（DNSリバインディング対策）。"""
    assert client.get("/", headers={"Host": "127.0.0.1:5000"}).status_code == 200
    assert client.get("/", headers={"Host": "localhost:5000"}).status_code == 200
    assert client.get("/", headers={"Host": "evil.example"}).status_code == 400
    assert client.get("/forms/new", headers={"Host": "evil.example:5000"}).status_code == 400


def test_a_refused_host_is_told_which_address_works(client):
    """別の名前（PC名・hosts の別名）で開かれたときは、同じアドレスのホームへではなく、開けるアドレスへ案内する。"""
    res = client.get("/", headers={"Host": "mypc:5123"})
    page = res.get_data(as_text=True)
    assert res.status_code == 400
    assert "http://127.0.0.1:5123/" in page and 'href="http://127.0.0.1:5123/"' in page
    assert ">ホームへ</a>" not in page   # 同じアドレスの / へ戻す（また断られる）ボタンは出さない
    res = client.post("/forms/1/delete", headers={"Host": "mypc", "Accept": "application/json"})
    assert res.status_code == 400 and "http://127.0.0.1:5000/" in res.get_json()["error"]


def test_pages_and_json_are_not_kept_in_the_browser_cache(app, client):
    """取り込んだ値の載る画面・JSON はブラウザに保存させない（ダウンロードで消したあと、戻るで出さない）。"""
    from tests.test_retention import _add_confirmed_document

    doc_id, _path = _add_confirmed_document(app, "キャッシュ.xlsx")
    for url in ("/", f"/forms/{doc_id}/done", "/forms/999", "/api/jobs/999"):
        res = client.get(url)
        assert res.status_code == (404 if "999" in url else 200), url
        assert res.headers.get("Cache-Control") == "no-store", url
    assert client.get("/", headers={"Host": "evil.example"}).headers.get("Cache-Control") == "no-store"
    static = client.get("/static/app.js")
    assert static.status_code == 200 and "no-store" not in (static.headers.get("Cache-Control") or "")


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


def test_home_keeps_modified_forms_in_recent_confirmed(app, client):
    # 確定したあとに1文字直すと 修正中 になるが、ダウンロードされるのは確定済みの版
    _add_document(app, "直した.xlsx", data=_extraction("EQ-002"), confirmed=_extraction("EQ-001"),
                  title="8D-2025-007 直した")
    page = client.get("/").get_data(as_text=True)
    assert "8D-2025-007 直した" in page
    assert "確定した帳票はまだありません" not in page
    assert "修正中（確定済みの版あり）" in page


def test_home_shows_the_modified_badge_inside_a_batch_row(app, client):
    """まとめ取り込みの中の修正中の帳票も、確認文を開く前に一覧で分かる。"""
    with app.app_context():
        for order, (name, data, title) in enumerate((("1.xlsx", _extraction("EQ-002"), "B-001 直した"),
                                                     ("2.xlsx", _extraction("EQ-001"), "B-002 そのまま"))):
            doc_id = db.create_document(name, "0" * 64, f"documents/{name}", batch_id="B", batch_order=order)
            db.update_document(doc_id, data_json=data, confirmed_json=_extraction("EQ-001"), title=title)
    page = client.get("/").get_data(as_text=True)
    assert "まとめ取り込み（2ファイル）" in page
    row = page[page.index("B-001 直した"):page.index("B-002 そのまま")]
    assert "修正中（確定済みの版あり）" in row
    rest = page[page.index("B-002 そのまま"):page.index("zipをダウンロード")]
    assert "修正中（確定済みの版あり）" not in rest


def test_home_ready_list_excludes_forms_without_confirmed_version(app, client):
    _add_document(app, "確認中.xlsx", data=_extraction(), title="R-003 確認中")
    page = client.get("/").get_data(as_text=True)
    assert "ダウンロード待ちの帳票はありません" in page
    assert "R-003 確認中" in page  # 作業中には出る


def test_home_says_how_many_are_not_shown_and_can_show_all(app, client):
    """ホームはこのPCに残っているデータの唯一の一覧。50件を超えた分を黙って切らない。"""
    for i in range(51):
        _add_document(app, f"未読{i:02d}.xlsx")
    _add_document(app, "確定.xlsx", data=_extraction(), confirmed=_extraction(), title="R-900 確定")
    page = client.get("/").get_data(as_text=True)
    assert "未読50.xlsx" in page and "未読00.xlsx" not in page   # 新しい順に50件
    assert "ほかに1件あります" in page and "/?all=1" in page
    assert page.count("ほかに") == 1                               # 超えていない一覧には出さない
    page = client.get("/?all=1").get_data(as_text=True)
    assert "未読00.xlsx" in page and "ほかに1件あります" not in page


def test_server_error_page_is_japanese(tmp_path):
    """想定外の例外でも、英語の既定ページではなく日本語の案内と戻り道を出す。"""
    app = create_app(make_config(tmp_path, PROPAGATE_EXCEPTIONS=False))

    @app.route("/_raise_for_test")
    def _raise_for_test():
        raise RuntimeError("テスト用の例外")

    res = app.test_client().get("/_raise_for_test")
    assert res.status_code == 500
    page = res.get_data(as_text=True)
    assert "エラーが発生しました" in page and 'href="/"' in page


# ---- 一覧表の取り込み設定（JSONの書き出し・読み込み） ------------------------------------------

def _make_template(app, name="トラブル対応一覧"):
    from tables import store
    from tables.spec import spec_from_dict

    spec = spec_from_dict({
        "name": name,
        "columns": [{"key": "record_no", "display": "管理No", "headers": ["管理No"], "type": "code",
                     "role": "key", "required": True}],
        "record": {"key": ["record_no"]},
    })
    with app.app_context():
        template_id, _version_id = store.create_template(name, spec)
    return template_id


def test_table_template_export_has_ascii_filename(app, client):
    template_id = _make_template(app)
    res = client.get(f"/settings/table-templates/{template_id}/export.json")
    assert res.status_code == 200 and json.loads(res.data)["name"] == "トラブル対応一覧"
    disposition = res.headers["Content-Disposition"]
    # filename* を解さない古いクライアント向けに ASCII の filename= も付ける（他のダウンロードと同じ）
    assert "filename=" in disposition.replace("filename*=", "") and "filename*=UTF-8''" in disposition


@pytest.mark.parametrize("body", ['{"spec": {}}', '{"spec": {"name": "x", "columns": [{"key": "a"}]}}',
                                  '{"foo": 1}', "これはJSONではありません"])
def test_table_template_import_rejects_other_json(client, body):
    res = client.post("/settings/table-templates/import",
                      data={"file": (io.BytesIO(body.encode("utf-8")), "設定.json")},
                      content_type="multipart/form-data", follow_redirects=True)
    page = res.get_data(as_text=True)
    assert "このファイルは一覧表の取り込み設定として読み込めません" in page
    # Python の例外文や内部のキー名は画面に出さない
    for word in ("__init__", "TypeError", "positional argument", "spec がありません"):
        assert word not in page


# ---- 他サイトからの書き込みを断る ------------------------------------------------------------

CROSS_SITE = {"Origin": "https://evil.example", "Sec-Fetch-Site": "cross-site"}


def test_cross_site_post_is_refused(ai_client):
    res = ai_client.post("/settings/ai", data={"chat_url": "http://evil.example/v1/chat/completions"},
                         headers=CROSS_SITE)
    assert res.status_code == 403
    assert ai_client.post("/settings/ai/test", headers=CROSS_SITE).status_code == 403
    # Origin だけ・Sec-Fetch-Site だけでも断る
    assert ai_client.post("/settings/ai/test", headers={"Origin": "https://evil.example"}).status_code == 403
    assert ai_client.post("/settings/ai/test", headers={"Sec-Fetch-Site": "cross-site"}).status_code == 403


def test_same_origin_post_still_works(ai_client):
    res = ai_client.post("/settings/ai", data={"models": ["gpt-test"], "default": "gpt-test", "api_key": "",
                                               "chat_url": "", "models_url": "", "add_models": ""},
                         headers={"Origin": "http://localhost", "Sec-Fetch-Site": "same-origin"},
                         follow_redirects=True)
    assert res.status_code == 200 and "AI接続の設定を保存しました" in res.get_data(as_text=True)


def test_cross_site_get_is_allowed(client):
    assert client.get("/", headers=CROSS_SITE).status_code == 200


def test_refused_write_shows_a_japanese_page(ai_client):
    """断ったことが日本語で分かり、ホームに戻れる（Flask の英語の403ページを出さない）。"""
    res = ai_client.post("/settings/ai/test", headers=CROSS_SITE)
    assert res.status_code == 403
    page = res.get_data(as_text=True)
    assert "ほかのサイトのページから送られてきた操作" in page
    assert "データは変わっていません" in page and 'href="/"' in page
    assert "Forbidden" not in page


def test_refused_write_answers_json_when_the_screen_asked_for_json(ai_client):
    """画面の JSON 送信（app.js の postJson）には JSON で返す。HTML だと理由が出ない。"""
    res = ai_client.post("/settings/ai/test", headers={**CROSS_SITE, "Accept": "application/json"})
    assert res.status_code == 403
    assert "ほかのサイト" in res.get_json()["error"]


def test_too_large_upload_page_is_japanese(tmp_path):
    """MAX_CONTENT_LENGTH を超えた送信は Werkzeug の英語の画面ではなく、日本語の案内とホームへのボタン。"""
    small = create_app(make_config(tmp_path, MAX_CONTENT_LENGTH=1024))
    res = small.test_client().post("/forms/upload", data={"files": (io.BytesIO(b"x" * 5000), "大きい.xlsx")},
                                   content_type="multipart/form-data")
    body = res.get_data(as_text=True)
    assert res.status_code == 413
    assert "ファイルが大きすぎます" in body and "合計 1KB まで" in body and "分けて" in body
    assert "Request Entity Too Large" not in body and 'href="/"' in body
    res = small.test_client().post("/forms/upload", data=b"x" * 5000, headers={"Accept": "application/json"},
                                   content_type="application/json")
    assert res.status_code == 413 and "大きすぎます" in res.get_json()["error"]
