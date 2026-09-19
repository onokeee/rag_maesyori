"""一覧表の AI整形の画面の経路（試し・見積もり・実行・一時停止・再開・中止）を偽の AI サーバーで通す。"""
import io

import pytest

from app import create_app
from core import jobs
from services import llm
from tables import store
from tests.conftest import BufferedClient, make_config
from tests.fake_servers import OPENAI_KEY, FakeServer, Reply, keep_scenario
from tests.test_tables_flow import COLUMNS, CSV_TEXT


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
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "a.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "T"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "T", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "lightrag_hint": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                            "fill_down_blank": False, "ai": r == "log", "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    res = client.post(f"/tables/imports/{import_id}/columns", json=payload)
    assert res.status_code == 200, res.get_json()
    with app.app_context():
        job = jobs.wait_job(store.get_import(import_id)["job_id"], timeout=60)
    assert job["status"] == "done", job
    return import_id


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


def test_run_pause_resume_and_cancel(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)

    def slow(body, srv):
        reply = keep_scenario(body, srv)
        return Reply(content=reply.content, delay=0.3) if isinstance(reply, Reply) else reply

    fake.responder = slow
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
    # 一時停止中は確認画面に進まず、AI整形の画面へ戻す
    res = ai_client.get(f"/tables/imports/{import_id}/preview")
    assert res.status_code == 302 and res.headers["Location"].endswith("/ai")

    assert ai_client.post(f"/tables/imports/{import_id}/ai/resume", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        assert jobs.wait_job(job_id, timeout=30, statuses=("running",) + tuple(jobs.TERMINAL_STATUSES))["status"] \
            != "paused"

    # 画面のボタン（フォーム送信）で中止すると AI整形の画面へ戻る
    res = ai_client.post(f"/tables/imports/{import_id}/ai/cancel")
    assert res.status_code == 302 and res.headers["Location"].endswith(f"/tables/imports/{import_id}/ai")
    with ai_app.app_context():
        job = jobs.wait_job(job_id, timeout=30)
    assert job["status"] in ("cancelled", "done")

    # 終わったジョブへの操作は ok: false（JSON）・エラーの表示（フォーム）
    assert ai_client.post(f"/tables/imports/{import_id}/ai/pause", json={}).get_json() == {"ok": False}
    res = ai_client.post(f"/tables/imports/{import_id}/ai/resume", follow_redirects=True)
    assert "AI整形のジョブを操作できませんでした" in res.get_data(as_text=True)
    assert ai_client.post(f"/tables/imports/{import_id}/ai/explode", json={}).status_code == 404

    # 中止のあとは残りをもう一度実行できる
    fake.responder = keep_scenario
    res = ai_client.post(f"/tables/imports/{import_id}/ai/run", json={"scope": "pending", "concurrency": 3})
    assert res.status_code == 200, res.get_json()
    with ai_app.app_context():
        assert jobs.wait_job(res.get_json()["job_id"], timeout=60)["status"] == "done"


def test_cancel_with_json_stops_the_job(ai_app, ai_client, fake):
    import_id = _read_import(ai_app, ai_client)

    def slow(body, srv):
        reply = keep_scenario(body, srv)
        return Reply(content=reply.content, delay=0.3) if isinstance(reply, Reply) else reply

    fake.responder = slow
    job_id = ai_client.post(f"/tables/imports/{import_id}/ai/run",
                            json={"scope": "all", "concurrency": 1}).get_json()["job_id"]
    assert ai_client.post(f"/tables/imports/{import_id}/ai/cancel", json={}).get_json() == {"ok": True}
    with ai_app.app_context():
        job = jobs.wait_job(job_id, timeout=30)
    assert job["status"] == "cancelled"


def test_control_without_any_job(ai_app, ai_client):
    import_id = _read_import(ai_app, ai_client)
    for action in ("pause", "resume", "cancel"):
        assert ai_client.post(f"/tables/imports/{import_id}/ai/{action}", json={}).get_json() == {"ok": False}
    assert ai_client.post("/tables/imports/9999/ai/pause", json={}).status_code == 404
