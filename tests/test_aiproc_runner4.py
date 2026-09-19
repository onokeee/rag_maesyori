"""AI整形の実行（4巡目）：再依頼だけの失敗、試し実行中の削除、AI呼び出しの失敗の日本語。"""
import threading
import time

import pytest

from aiproc import runner
from core import purge
from models import database
from services import llm
from tests.fake_servers import Reply, is_repair, keep_scenario, segments_of
from tests.test_aiproc import ROWS, _make_import, ai_app, fake  # noqa: F401  (fixture)


def _execute_r2(app, iid):
    with app.app_context():
        settings = llm.job_client_settings()
        work = runner.prepare_works(runner.load_rows_for_ai(iid), ["log"])[0]
        runner.assign_keys([work], settings, "json_schema")
        return runner.execute_work(work, settings, "json_schema")


def test_repair_400_keeps_first_flagged_result(ai_app, fake):
    """再依頼が 400（文脈長超過）でも、1回目の要確認の結果（通った項目）を残す。"""
    def responder(body, srv):
        if segments_of(body) and is_repair(body):
            return Reply(status=400, body={"error": {"message": "This model's maximum context length is 4096 tokens",
                                                     "type": "invalid_request_error"}})
        return keep_scenario(body, srv)

    fake.responder = responder
    iid = _make_import(ai_app, rows={"R2": ROWS["R2"]})
    out = _execute_r2(ai_app, iid)
    assert out.status == "flagged", out.error
    assert out.result["incident"]["permanent_actions"][0]["v"] == "コネクタの増し締め"
    assert out.checks["repaired"] is False and "context length" in out.checks["repair_error"]
    assert out.attempts == 2 and out.error is None


def test_repair_row_deadline_keeps_first_flagged_result(ai_app, fake, monkeypatch):
    """再依頼が行の時間の上限を過ぎても、1回目の要確認の結果を残す。"""
    monkeypatch.setattr(llm, "LOCAL_TIMEOUT", 0.5)       # 行の上限は1秒
    done = threading.Event()

    def responder(body, srv):
        if segments_of(body) and is_repair(body):
            done.set()
            return Reply(content="{}", delay=3)
        return keep_scenario(body, srv)

    fake.responder = responder
    iid = _make_import(ai_app, rows={"R2": ROWS["R2"]})
    out = _execute_r2(ai_app, iid)
    assert out.status == "flagged", out.error
    assert out.result["incident"]["permanent_actions"]
    assert "打ち切りました" in out.checks["repair_error"]
    assert done.wait(1)
    time.sleep(3)   # 見捨てた要求のハンドラが終わるまで待つ


def test_trial_does_not_write_after_import_deleted(ai_app, fake):
    """試し実行の応答待ちの間に取り込みを消したら、結果も生の応答も書かない（design.md 3.3）。"""
    started = threading.Event()

    def responder(body, srv):
        if segments_of(body):
            started.set()
            reply = keep_scenario(body, srv)
            reply.delay = 1.5
            return reply
        return keep_scenario(body, srv)

    fake.responder = responder
    iid = _make_import(ai_app, rows={"R1": ROWS["R1"]})
    errors = []

    def trial():
        with ai_app.app_context():
            try:
                runner.trial_row(iid, "R1", stage_ids=["log"])
            except runner.AIJobError as e:
                errors.append(str(e))

    th = threading.Thread(target=trial)
    th.start()
    assert started.wait(10)
    with ai_app.app_context():
        purge.purge_table_import(iid)
    th.join(20)
    assert errors == ["この取り込みは削除されました。"]
    with ai_app.app_context():
        conn = database.connect()
        try:
            assert conn.execute("SELECT COUNT(*) FROM ai_items").fetchone()[0] == 0
            assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0
        finally:
            conn.close()


@pytest.mark.parametrize("make,expected", [
    (lambda req: __import__("openai").APIConnectionError(request=req), "接続できませんでした"),
    (lambda req: __import__("openai").APITimeoutError(request=req), "時間内に返りませんでした"),
])
def test_friendly_error_is_japanese_for_connection_and_timeout(ai_app, make, expected):
    import httpx2 as httpx

    req = httpx.Request("POST", "http://127.0.0.1:9/v1/chat/completions")
    with ai_app.test_request_context():
        text = llm.friendly_error(make(req))
    assert expected in text
    assert "Connection error" not in text and "timed out" not in text.lower()


def test_llm_messages_name_the_ai_connection_screen(ai_app):
    with ai_app.test_request_context():
        class Denied(Exception):
            status_code = 401
        assert "「AI接続」" in llm.friendly_error(Denied("x"))
    import inspect
    assert "「AI設定」" not in inspect.getsource(llm)
