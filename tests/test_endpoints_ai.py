"""一覧表の AI整形の段の経路（試し・見積もり・実行・一時停止・再開・中止）を偽の AI サーバーで通す。"""
import pytest

from app import create_app
from core import jobs
from services import llm
from tests.conftest import BufferedClient, make_config
from tests.fake_servers import OPENAI_KEY, FakeServer, Reply, keep_scenario
from tests.tables_helpers import imported, panel


@pytest.fixture
def fake():
    with FakeServer() as server:
        server.responder = keep_scenario
        yield server


@pytest.fixture
def ai_app(tmp_path, fake):
    llm.reset_llm_client()
    llm.forget_structured_modes()
    app = create_app(make_config(tmp_path, OPENAI_BASE_URL=f"{fake.url}/v1", OPENAI_API_KEY=OPENAI_KEY,
                                 OPENAI_MODELS=["gpt-test"], OPENAI_MODEL="gpt-test"))
    yield app
    llm.reset_llm_client()
    llm.forget_structured_modes()


@pytest.fixture
def ai_client(ai_app):
    ai_app.test_client_class = BufferedClient
    return ai_app.test_client()


def _read_import(app, client) -> int:
    """CSV を取り込み、列の対応づけ（追記ログ列あり）まで済ませて読み込みを終える。"""
    return imported(app, client, "a.csv", "T", ai_role="log")


def test_trial_returns_markdown_for_a_row(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)
    res = ai_client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 200, res.get_json()
    body = res.get_json()
    assert body["row_key"] == "TR-001" and body["status"] == "ok" and body["route"] == "ai"
    assert "TR-001" in body["markdown"] and "対応の時系列" in body["markdown"]
    assert body["stats"] and body["stats"][0]["tokens_in"] > 0
    assert fake.chat_requests()


def test_trial_with_an_unknown_row_key_is_an_error(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)
    res = ai_client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "NOPE"})
    assert res.status_code == 400
    assert "NOPE" in res.get_json()["error"]
    assert not fake.chat_requests()


def test_estimate_uses_the_trial(ai_app, ai_client):
    import_id = _read_import(ai_app, ai_client)
    trial = ai_client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"}).get_json()
    res = ai_client.post(f"/tables/imports/{import_id}/ai/estimate",
                         json={"scope": "pending", "concurrency": 2, "trials": [trial]})
    assert res.status_code == 200, res.get_json()
    body = res.get_json()
    assert body["rows"] == 12 and body["concurrency"] == 2
    assert body["already_rows"] == 1 and body["ai_rows"] == 11    # 試した行は結果がもう保存されている


def _slow(fake):
    def slow(body, srv):
        reply = keep_scenario(body, srv)
        return Reply(content=reply.content, delay=0.3) if isinstance(reply, Reply) else reply

    fake.responder = slow


def test_run_pause_resume_and_cancel(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)
    _slow(fake)
    res = ai_client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "all", "concurrency": 1})
    assert res.status_code == 200, res.get_json()
    job_id = res.get_json()["job_id"]
    assert res.get_json()["job_url"] == f"/api/jobs/{job_id}"

    # 実行中にもう一度押しても2つ目は始めない
    again = ai_client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "all", "concurrency": 1})
    assert again.status_code == 400 and "実行中" in again.get_json()["error"]

    assert ai_client.post(f"/tables/imports/{import_id}/ai/pause", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        assert jobs.wait_job(job_id, timeout=30, statuses=("paused",))["status"] == "paused"
    assert ai_client.get(f"/api/jobs/{job_id}").get_json()["status"] == "paused"
    # 一時停止中は「内容の確認」の段を開かせず、理由を1行出す
    preview = panel(ai_client, import_id, "preview")
    assert "AI整形が動いています" in preview["locked"]

    assert ai_client.post(f"/tables/imports/{import_id}/ai/resume", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        assert jobs.wait_job(job_id, timeout=30, statuses=("running",) + tuple(jobs.TERMINAL_STATUSES))["status"] \
            != "paused"

    assert ai_client.post(f"/tables/imports/{import_id}/ai/cancel", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        job = jobs.wait_job(job_id, timeout=30)
    assert job["status"] in ("cancelled", "done")

    # 終わったジョブへの操作は、理由を返して断る
    for action in ("pause", "resume"):
        res = ai_client.post(f"/tables/imports/{import_id}/ai/{action}", json={})
        assert res.status_code == 400 and "AI整形のジョブを操作できませんでした" in res.get_json()["error"]
    assert ai_client.post(f"/tables/imports/{import_id}/ai/explode", json={}).status_code == 404

    # 中止のあとは残りをもう一度実行できる
    fake.responder = keep_scenario
    res = ai_client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "pending", "concurrency": 3})
    assert res.status_code == 200, res.get_json()
    with ai_app.app_context():
        assert jobs.wait_job(res.get_json()["job_id"], timeout=60)["status"] == "done"


def test_cancel_stops_the_job(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)
    _slow(fake)
    job_id = ai_client.post(f"/tables/imports/{import_id}/ai/run",
                            json={"scope": "all", "concurrency": 1}).get_json()["job_id"]
    assert ai_client.post(f"/tables/imports/{import_id}/ai/cancel", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        job = jobs.wait_job(job_id, timeout=30)
    assert job["status"] == "cancelled"


def test_control_without_any_job(ai_app, ai_client):
    import_id = _read_import(ai_app, ai_client)
    for action in ("pause", "resume", "cancel"):
        res = ai_client.post(f"/tables/imports/{import_id}/ai/{action}", json={})
        assert res.status_code == 400 and "操作できませんでした" in res.get_json()["error"]
    assert ai_client.post("/tables/imports/9999/ai/pause", json={}).status_code == 404
