"""ヘッダー右上の「AI接続」（ブラウザごとの保存・状態の表示）と、分割プレビューの「順番」「原文と並べて」。

利用者の指示（2026-09-21）:
  「AI接続の設定は、もともとの位置ヘッダーの画面右上『AI接続』に移動させる。全部空欄にしておいて
   cookieでユーザー毎に登録内容をずっと保持させるようにしてほしい。」
  「AI接続はヘッダー上で、接続中 か 未接続 一目で分かるように」
  「（S1，S2…は）順番1、順番2の表示にしてほしい。また、原文と並べて見れるようにしたい」

設定は database.ai_connections にブラウザ（views.current_session_id）ごとに入り、取り込みを捨てる片付けでは消えない。
接続を確かめに行くのは「保存」「パネルを開く」「AI整形を始める」の3つだけで、画面を開くだけでは行かない。
"""
from __future__ import annotations

import re
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from app import core
from app import create_app
from app import database as db
from app import llm
from tests.conftest import (
    OPENAI_KEY,
    BufferedClient,
    FakeServer,
    add_confirmed_document,
    imported,
    keep_scenario,
    make_config,
    upload_csv,
)

STATIC = Path(__file__).resolve().parents[1] / "app" / "static"

DEAD_URL = "http://127.0.0.1:9/v1/chat/completions"   # 答えない接続先（discard ポート）


@pytest.fixture
def fake():
    with FakeServer() as server:
        server.responder = keep_scenario
        yield server


@pytest.fixture
def app_env_blank(tmp_path):
    """env に何も無いアプリ（キー・接続先は画面から入れる）。"""
    llm.reset_llm_client()
    llm.forget_structured_modes()
    app = create_app(make_config(tmp_path, OPENAI_API_KEY="", OPENAI_MODELS=[], OPENAI_MODEL="gpt-test"))
    app.test_client_class = BufferedClient
    yield app
    llm.reset_llm_client()
    llm.forget_structured_modes()


@pytest.fixture
def client(app_env_blank):
    return app_env_blank.test_client()


@pytest.fixture
def other_client(app_env_blank):
    """もう1台のPC（別のブラウザ）。クッキーが別なので設定も別になる。"""
    return app_env_blank.test_client()


def _sid(client) -> str:
    with client.session_transaction() as sess:
        return sess["sid"]


def _save(client, fake=None, **fields):
    body = {"api_key": OPENAI_KEY, "chat_url": f"{fake.url}/v1/chat/completions" if fake else "",
            "models_url": f"{fake.url}/v1/models" if fake else "", "model": "gpt-test"}
    body.update(fields)
    res = client.post("/tables/ai-connection", json=body)
    assert res.status_code == 200, res.get_json()
    return res.get_json()


def _header(html: str) -> dict:
    """ヘッダーの「AI接続」の状態（data-state と表示語）を HTML から拾う。"""
    m = re.search(r'data-ai-header data-state="(\w+)".*?data-ai-state>([^<]*)<.*?data-ai-sub>([^<]*)<', html, re.DOTALL)
    assert m, "ヘッダーに AI接続 が無い"
    return {"state": m.group(1), "label": m.group(2), "sub": m.group(3)}


def _input_value(html: str, field: str) -> str:
    m = re.search(rf'data-ai-field="{field}"[^>]*\svalue="([^"]*)"', html)
    return m.group(1) if m else ""


# ---- ヘッダーとパネル --------------------------------------------------------------------------

def test_every_screen_has_the_ai_connection_control_in_the_header(client):
    """3つの画面すべてのヘッダー右上に「AI接続」があり、押すと開くパネル（dialog）も同じページにある。"""
    for url in ("/forms/new", "/tables/new", "/form-types/"):
        html = client.get(url).get_data(as_text=True)
        assert 'data-ai-header' in html and 'id="aiDialog"' in html, url
        head = _header(html)
        assert head["label"] == "未接続" and head["state"] == "off", url
        # ナビは3つのまま（AI接続は画面ではない）
        assert html.count('class="mainnav"') == 1 and "/settings" not in html
    # パネルの注意書き（利用者に伝えること）
    html = client.get("/tables/new").get_data(as_text=True)
    assert "APIキーはこのサーバーに、あなたのブラウザのIDと結びつけて保存します（暗号化はしていません）" in html
    assert "共有PCでは使い終わりに［キーを消す］を押してください" in html
    assert "キーを消す" in html and "接続を確かめる" in html and "保存する" in html


def test_fields_start_blank_and_server_values_only_appear_as_placeholders(tmp_path):
    """env にキーと接続先があっても、入力欄は空。ふつうの値は placeholder に出す。env のキーは画面に出ない。"""
    app = create_app(make_config(tmp_path, OPENAI_BASE_URL="http://ai.example.local:8000/v1",
                                 OPENAI_API_KEY="sk-env-secret-000", OPENAI_MODEL="gpt-env"))
    html = app.test_client().get("/forms/new").get_data(as_text=True)
    assert "sk-env-secret-000" not in html
    assert _input_value(html, "chat_url") == "" and _input_value(html, "models_url") == "" and _input_value(html, "model") == ""
    assert 'placeholder="https://api.openai.com/v1/chat/completions"' in html
    assert 'placeholder="https://api.openai.com/v1/models"' in html
    # サーバー共通のキーがあることは知らせる（キーの値は出さない）
    assert "env ファイルに共通のAPIキーがあります" in html
    head = _header(html)
    assert head["state"] == "unchecked" and head["label"] == "未確認"   # 使えるが、このブラウザではまだ確かめていない


def test_step_five_no_longer_holds_the_connection_form(app_env_blank, client):
    """⑤AI整形の段には接続の入力欄が無く、ヘッダー右上を指す1行だけ。"""
    import_id = imported(app_env_blank, client, "a.csv", "T", ai_role="log")
    html = client.get(f"/tables/imports/{import_id}/panel/ai").get_json()["html"]
    assert "画面右上の［AI接続］" in html
    assert 'data-ai-field="api_key"' not in html and "AI接続を保存" not in html
    assert "分割プレビュー" in html   # AI整形の段そのものは残る


# ---- ブラウザごとの保存 ------------------------------------------------------------------------

def test_settings_are_saved_per_browser_and_the_key_is_never_echoed(app_env_blank, client, other_client, fake):
    res = _save(client, fake)
    assert res["connected"] is True and res["status"]["state"] == "ok" and res["status"]["state_label"] == "接続中"
    assert OPENAI_KEY not in str(res)

    html = client.get("/forms/new").get_data(as_text=True)
    head = _header(html)
    assert head["state"] == "ok" and head["label"] == "接続中" and head["sub"].startswith("最終確認 ")
    assert OPENAI_KEY not in html
    assert _input_value(html, "chat_url") == f"{fake.url}/v1/chat/completions"   # 自分が保存した値は欄に戻る
    assert "保存済み（変えるときだけ入力）" in html

    # 別のブラウザには何も無い（キーも接続先も、状態も）
    other = other_client.get("/forms/new").get_data(as_text=True)
    assert _header(other)["state"] == "off" and _input_value(other, "chat_url") == ""
    with app_env_blank.app_context():
        assert db.get_ai_connection(_sid(other_client)) is None
        assert db.get_ai_connection(_sid(client))["api_key"] == OPENAI_KEY


def test_a_url_that_does_not_answer_turns_the_light_red_with_the_reason(app_env_blank, client, fake):
    """答えない接続先を保存すると「つながりません」。パネルを開くと理由が出る。直すと「接続中」に戻る。"""
    res = _save(client, fake, chat_url=DEAD_URL)
    assert res["connected"] is False
    status = res["status"]
    assert status["state"] == "ng" and status["state_label"] == "つながりません"
    assert "接続できませんでした" in status["check_detail"]
    assert [s["ok"] for s in res["steps"]] == [True, False]   # モデル一覧は取れたがチャットが届かない

    html = client.get("/tables/new").get_data(as_text=True)
    head = _header(html)
    assert head["state"] == "ng" and head["label"] == "つながりません" and head["sub"].startswith("最終確認 ")
    assert "接続できませんでした" in html   # パネルに理由

    res = _save(client, fake)   # 接続先を直す
    assert res["status"]["state"] == "ok"


def test_page_loads_never_check_but_opening_the_panel_does(app_env_blank, client, fake):
    """画面を開くだけでは AI に問い合わせない（費用がかかる・待たされる）。パネルを開いたときの /test で確かめる。"""
    _save(client, fake)
    before = len(fake.requests)
    for url in ("/forms/new", "/tables/new", "/form-types/"):
        client.get(url)
    assert len(fake.requests) == before   # 何も送っていない（GET /v1/models も記録されないが POST は増えない）
    chat_before = len(fake.chat_requests())

    res = client.post("/tables/ai-connection/test").get_json()
    assert res["ok"] is True and res["status"]["state"] == "ok"
    assert len(fake.chat_requests()) == chat_before + 1   # 短いチャットを1回
    headers = {k.lower(): v for k, v in fake.chat_requests()[-1]["headers"].items()}
    assert headers.get("authorization") == f"Bearer {OPENAI_KEY}"   # このブラウザのキーで


def test_clear_key_button_forgets_only_this_browsers_key(app_env_blank, client, other_client, fake):
    _save(client, fake)
    _save(other_client, fake)
    res = client.post("/tables/ai-connection/clear").get_json()
    assert res["ok"] is True and res["status"]["state"] == "off" and res["status"]["api_key_saved"] is False
    assert res["status"]["chat_url"] == f"{fake.url}/v1/chat/completions"   # 接続先は残る
    with app_env_blank.app_context():
        assert db.get_ai_connection(_sid(client))["api_key"] == ""
        assert db.get_ai_connection(_sid(other_client))["api_key"] == OPENAI_KEY   # 隣の人の分は消えない
    # 消したあとは「未設定」なので、確かめに行かない
    chat_before = len(fake.chat_requests())
    assert client.post("/tables/ai-connection/test").get_json()["status"]["state"] == "off"
    assert len(fake.chat_requests()) == chat_before


def test_saving_without_a_key_keeps_the_saved_key(app_env_blank, client, fake):
    """キーの欄は空のまま保存しても（入力欄にはキーを出さないので）、保存済みのキーは消えない。"""
    _save(client, fake)
    res = _save(client, fake, api_key="", model="gpt-test")
    assert res["status"]["api_key_saved"] is True and res["status"]["state"] == "ok"


def test_validation_messages(client):
    bad = client.post("/tables/ai-connection", json={"api_key": "short"}).get_json()
    assert "APIキーの長さが不自然です" in bad["error"]
    bad = client.post("/tables/ai-connection", json={"chat_url": "https://example.com/v1"}).get_json()
    assert "/chat/completions で終わるフルパス" in bad["error"]
    bad = client.post("/tables/ai-connection", json={"models_url": "ftp://x/models"}).get_json()
    assert "http:// か https://" in bad["error"]


# ---- 片付けで消えない・クッキーは1年 ------------------------------------------------------------

def test_ai_settings_survive_the_purges_that_throw_away_uploads(app_env_blank, client, fake, tmp_path):
    """その人の帳票・一覧表を捨てても、起動時の片付け・見回りを通しても、AI接続の設定は残る。"""
    _save(client, fake)
    sid = _sid(client)
    import_id = upload_csv(client, "捨てられる一覧.csv")
    doc_id, _ = add_confirmed_document(app_env_blank, "捨てられる帳票.xlsx")
    with app_env_blank.app_context():
        db.update_document(doc_id, session_id=sid)
        conn = db.get_db()
        conn.execute("UPDATE table_imports SET session_id = ? WHERE id = ?", (sid, import_id))
        conn.commit()

        forms, tables_removed = core.purge_session(sid)          # 画面を離れたときの片付け
        assert (forms, tables_removed) == (1, 1)
        assert db.get_ai_connection(sid)["api_key"] == OPENAI_KEY

        core.sweep_stale(0)                                       # 見回り（0時間 = 全部が古い）
        core.purge_all_pending()                                  # 起動時の片付け
        assert db.get_ai_connection(sid)["api_key"] == OPENAI_KEY
        assert db.get_document(doc_id) is None

    # 起動し直しても残る（_purge_pending / _cleanup_leftovers を通る）
    config = make_config(tmp_path, OPENAI_API_KEY="", OPENAI_MODELS=[], OPENAI_MODEL="gpt-test")
    config["TESTING"] = False
    again = create_app(config)
    with again.app_context():
        assert db.get_ai_connection(sid)["api_key"] == OPENAI_KEY
    # 画面にも「接続中」のまま出る（同じクッキーで開く）
    html = client.get("/forms/new").get_data(as_text=True)
    assert _header(html)["state"] == "ok"


def test_purge_never_touches_settings_tables_even_by_column_name(app_env_blank):
    """取り込みを指す列を探して消す仕組みが、設定の表（ai_connections・patterns）には及ばない。"""
    with app_env_blank.app_context():
        conn = db.get_db()
        assert "ai_connections" not in core._tables_with_column(conn, "session_id")
        for table in core.SETTINGS_TABLES:
            assert table not in core._tables_with_column(conn, "import_id")


def test_only_connections_older_than_the_cookie_are_swept(app_env_blank, client, fake):
    _save(client, fake)
    sid = _sid(client)
    with app_env_blank.app_context():
        conn = db.get_db()
        old = (datetime.now() - timedelta(days=core.AI_CONNECTION_KEEP_DAYS + 1)).isoformat(timespec="seconds")
        conn.execute("INSERT INTO ai_connections (session_id, api_key, created_at, updated_at) VALUES ('x' * 32, 'sk-old', ?, ?)",
                     (old, old))
        conn.commit()
        assert core.sweep_stale_ai_connections() == 1
        assert db.get_ai_connection("x" * 32) is None
        assert db.get_ai_connection(sid)["api_key"] == OPENAI_KEY


def test_the_session_cookie_lasts_about_a_year_and_browsers_stay_apart(app_env_blank, client, other_client):
    res = client.get("/tables/new")
    cookie = res.headers.get("Set-Cookie", "")
    assert "HttpOnly" in cookie and "SameSite=Lax" in cookie
    m = re.search(r"Expires=([^;]+)", cookie)
    assert m, cookie
    expires = datetime.strptime(m.group(1).strip(), "%a, %d %b %Y %H:%M:%S GMT").replace(tzinfo=timezone.utc)
    days = (expires - datetime.now(timezone.utc)).days
    assert 360 <= days <= 366
    # 長くしても分かれ方は変わらない
    import_id = upload_csv(client, "自分の一覧.csv")
    assert other_client.get(f"/tables/imports/{import_id}/panel/source").status_code == 404
    assert _sid(client) != _sid(other_client)


# ---- 前の版の設定ファイル・env は「サーバー共通の値」として読むだけ ------------------------------------

def test_legacy_yaml_is_read_as_a_server_wide_fallback_and_announced(tmp_path, fake):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "model_settings.yaml").write_text(
        f"api_key: {OPENAI_KEY}\nchat_url: {fake.url}/v1/chat/completions\nmodels_url: {fake.url}/v1/models\n"
        "default: gpt-test\nmodels: [gpt-test]\n", encoding="utf-8")
    app = create_app(make_config(tmp_path, OPENAI_API_KEY="", OPENAI_MODELS=[], OPENAI_MODEL="gpt-env"))
    c = app.test_client()
    html = c.get("/tables/new").get_data(as_text=True)
    assert "前の版の設定ファイル" in html and "model_settings.yaml" in html
    assert OPENAI_KEY not in html and _input_value(html, "chat_url") == ""   # 欄には入れない
    assert _header(html)["state"] == "unchecked"
    with app.test_request_context("/"):
        assert llm.is_configured() and llm.llm_api_key_source() == "server" and llm.current_model() == "gpt-test"
    # 確かめれば「接続中」になる（yaml の値でつながる）
    res = c.post("/tables/ai-connection/test").get_json()
    assert res["ok"] is True and res["status"]["state"] == "ok"
    # 自分の値を入れればそちらが勝つ（キーは yaml のまま）
    assert c.post("/tables/ai-connection", json={"model": "gpt-mine"}).status_code == 200
    status = c.post("/tables/ai-connection/test").get_json()["status"]
    assert status["effective"]["model"] == "gpt-mine" and status["effective"]["model_source"] == "browser"
    assert status["api_key_source"] == "server" and status["api_key_saved"] is False


# ---- AI整形はブラウザの設定で動く（ジョブのスレッドにも渡る）・始める前に確かめる ------------------------

def test_trial_and_run_use_the_browser_settings(app_env_blank, client, fake):
    """env にキーが無くても、ブラウザが保存したキー・接続先で試し実行と全件実行が動く。"""
    _save(client, fake)
    import_id = imported(app_env_blank, client, "a.csv", "T", ai_role="log")
    res = client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 200, res.get_json()
    assert res.get_json()["status"] == "ok"
    headers = {k.lower(): v for k, v in fake.chat_requests()[-1]["headers"].items()}
    assert headers["authorization"] == f"Bearer {OPENAI_KEY}"

    res = client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "all", "concurrency": 2})
    assert res.status_code == 200, res.get_json()
    body = res.get_json()
    assert body["ai_status"]["state"] == "ok"   # 始める前に確かめた結果
    with app_env_blank.app_context():
        job = core.wait_job(body["job_id"], timeout=60)
    assert job["status"] == "done", job
    assert job["result"]["ok"] >= 1


def test_run_refuses_and_turns_the_light_red_when_the_url_does_not_answer(app_env_blank, client, fake):
    _save(client, fake)
    import_id = imported(app_env_blank, client, "a.csv", "T", ai_role="log")
    _save(client, fake, chat_url=DEAD_URL)
    res = client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "all", "concurrency": 1})
    assert res.status_code == 400
    body = res.get_json()
    assert "AIにつながりません" in body["error"] and "画面右上の「AI接続」" in body["error"]
    assert body["ai_status"]["state"] == "ng"
    with app_env_blank.app_context():
        assert core.latest_job("table_import", import_id, kind="ai_format") is None   # ジョブは作らない
    assert _header(client.get("/tables/new").get_data(as_text=True))["state"] == "ng"


def test_other_browser_cannot_run_ai_with_my_key(app_env_blank, client, other_client, fake):
    _save(client, fake)
    import_id = imported(app_env_blank, other_client, "b.csv", "T2", ai_role="log")
    res = other_client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 400 and "AI接続が設定されていません" in res.get_json()["error"]


# ---- 分割プレビュー: 順番と原文 ------------------------------------------------------------------

def test_split_preview_gives_order_numbers_and_positions_in_the_original(app_env_blank, client):
    import_id = imported(app_env_blank, client, "a.csv", "T", ai_role="log")
    split = client.post(f"/tables/imports/{import_id}/ai/split-preview", json={"row_key": "TR-001"}).get_json()
    segs = split["segments"]
    assert [s["no"] for s in segs] == [1, 2]
    assert [s["id"] for s in segs] == ["s1", "s2"]     # AI との受け渡し用の id はそのまま
    text = split["text"]
    assert "\n" in text                                 # 原文の改行はそのまま
    for s in segs:
        assert 0 <= s["start"] < s["end"] <= len(text)
        assert s["body"] in text[s["start"]:s["end"]]   # 位置は原文（text）の中を指す
    assert segs[0]["end"] <= segs[1]["start"]           # 重ならない


def test_the_screen_shows_order_numbers_and_the_original_side_by_side():
    """画面側（static/app.js）は s1, s2 を出さず「順番1」を出し、原文を色分けして左に並べる。"""
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    tables = js[js.index("// ==== tables: "):]
    assert "順番${" in tables and "text: s.id" not in tables
    assert "split-compare" in tables and "原文（セルのまま）" in tables and "分けた結果" in tables
    assert "つに分かれました（日付や記入者がオレンジのものは、書き方から推定したものです）" in tables
    assert "区切り ${" not in tables and "オレンジは推定）" not in tables
    css = (STATIC / "style.css").read_text(encoding="utf-8")
    assert "@media (max-width: 900px) { .split-compare { grid-template-columns: 1fr; } }" in css
