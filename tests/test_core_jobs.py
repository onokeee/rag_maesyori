"""core/jobs.py: 進捗・中止・一時停止・中断からの回復。"""
import threading
import time
from datetime import datetime, timedelta

import pytest
from flask import Flask, current_app

from core import jobs
from models import database as db


@pytest.fixture
def core_app(tmp_path):
    app = Flask(__name__)
    app.config.update(TESTING=True, DATABASE=tmp_path / "app.db", MARK="この取り込み")
    db.init_app(app)
    return app


def _wait_until(predicate, timeout=10.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        value = predicate()
        if value:
            return value
        time.sleep(0.02)
    raise AssertionError("時間内に条件を満たしませんでした")


def test_progress_and_result(core_app):
    def work(ctx):
        conn = db.connect()  # バックグラウンドスレッドでも DB と app 設定が使える
        try:
            count = conn.execute("SELECT COUNT(*) FROM jobs").fetchone()[0]
        finally:
            conn.close()
        for i in range(1, 4):
            ctx.progress(done=i, total=3, phase="読み込み")
        return {"rows": 3, "mark": current_app.config["MARK"], "jobs": count, "params": ctx.params}

    with core_app.app_context():
        job_id = jobs.start_job("table_read", "table_import", 7, work, params={"scope": "all"})
        job = jobs.wait_job(job_id, timeout=10)
        assert job["status"] == "done" and job["finished"] and job["status_label"] == "完了"
        assert job["progress"]["done"] == 3 and job["progress"]["phase"] == "読み込み"
        assert job["result"] == {"rows": 3, "mark": "この取り込み", "jobs": 1, "params": {"scope": "all"}}
        assert job["kind"] == "table_read" and job["ref_type"] == "table_import" and job["ref_id"] == 7
        assert job["heartbeat_at"]
        assert jobs.latest_job("table_import", 7)["id"] == job_id
        assert jobs.latest_job("table_import", 7, kind="other") is None


def test_failure_is_recorded(core_app):
    """利用者に見せるエラー（JobError）のメッセージは、そのまま画面に出る。"""
    def work(ctx):
        ctx.progress(done=1)
        raise jobs.JobError("列が見つかりません")

    with core_app.app_context():
        job = jobs.wait_job(jobs.start_job("table_read", "table_import", 1, work))
        assert job["status"] == "failed"
        assert job["message"] == "列が見つかりません" and job["progress"]["done"] == 1


def test_unexpected_failure_does_not_show_the_python_exception_name(core_app):
    """想定外の例外は Python の例外名（KeyError など）を画面に出さず、日本語の案内にする。"""
    def work(ctx):
        raise KeyError("equipment_no")

    with core_app.app_context():
        job = jobs.wait_job(jobs.start_job("table_read", "table_import", 1, work))
        assert job["status"] == "failed"
        assert "KeyError" not in job["message"] and "equipment_no" not in job["message"]
        assert "処理中にエラーが発生しました" in job["message"] and "ログ" in job["message"]


def test_ai_and_table_errors_are_shown_to_the_user(core_app):
    """AI整形・一覧表の処理エラーは日本語のメッセージを持つので、そのまま出す（JobError を継承している）。"""
    from aiproc.runner import AIJobError
    from tables.pipeline import PipelineError

    assert issubclass(AIJobError, jobs.JobError) and issubclass(PipelineError, jobs.JobError)
    assert jobs.error_message(AIJobError("AI接続が設定されていません")) == "AI接続が設定されていません"
    assert jobs.error_message(PipelineError("見出しの行が見つかりません")) == "見出しの行が見つかりません"


def test_cancel_running_job(core_app):
    started = threading.Event()

    def work(ctx):
        i = 0
        while True:
            i += 1
            ctx.progress(done=i)
            started.set()
            ctx.check_cancel()
            time.sleep(0.01)

    with core_app.app_context():
        job_id = jobs.start_job("ai", "table_import", 1, work)
        assert started.wait(10)
        assert jobs.request_cancel(job_id)
        job = jobs.wait_job(job_id)
        assert job["status"] == "cancelled" and job["message"] == "中止しました"
        assert job["cancel_requested"] == 1
        assert jobs.request_cancel(job_id) is False  # 終わったジョブには効かない


def test_pause_and_resume(core_app):
    finish = threading.Event()
    observed = {}

    def work(ctx):
        i = 0
        while not finish.is_set():
            i += 1
            ctx.progress(done=i)
            if ctx.should_stop():
                observed["should_stop"] = True
            if not ctx.wait_if_paused():
                return {"stopped": True}
            time.sleep(0.01)
        return {"done": i}

    with core_app.app_context():
        job_id = jobs.start_job("ai", "table_import", 1, work)
        _wait_until(lambda: (jobs.get_job(job_id) or {}).get("status") == "running")
        assert jobs.request_pause(job_id)
        paused = _wait_until(lambda: (j := jobs.get_job(job_id))["status"] == "paused" and j)
        assert paused["status_label"] == "一時停止中"
        time.sleep(0.5)
        assert jobs.get_job(job_id)["progress"]["done"] == paused["progress"]["done"]  # 止まっている
        assert observed.get("should_stop")

        assert jobs.request_resume(job_id)
        _wait_until(lambda: jobs.get_job(job_id)["progress"]["done"] > paused["progress"]["done"])
        assert jobs.get_job(job_id)["status"] == "running"
        finish.set()
        job = jobs.wait_job(job_id)
        assert job["status"] == "done" and job["result"]["done"] > paused["progress"]["done"]
        assert job["pause_requested"] == 0


def test_cancel_while_paused_and_queued(core_app):
    release = threading.Event()
    ran = []

    def blocker(ctx):
        while not release.is_set():
            if not ctx.wait_if_paused():
                return None
            time.sleep(0.01)
        return None

    def never(ctx):
        ran.append(ctx.job_id)

    with core_app.app_context():
        first = jobs.start_job("ai", "table_import", 1, blocker)
        second = jobs.start_job("ai", "table_import", 2, never)
        _wait_until(lambda: jobs.get_job(first)["status"] == "running")
        assert jobs.get_job(second)["status"] == "queued"
        assert jobs.request_cancel(second)
        assert jobs.get_job(second)["status"] == "cancelled"  # 待機中はその場で中止

        jobs.request_pause(first)
        _wait_until(lambda: jobs.get_job(first)["status"] == "paused")
        jobs.request_cancel(first)
        assert jobs.wait_job(first)["status"] == "cancelled"  # fn が正常終了しても中止扱い
        release.set()
        # 後ろのジョブが実行されないことを、次のジョブの完了で確かめる
        third = jobs.start_job("ai", "table_import", 3, lambda ctx: {"ok": True})
        assert jobs.wait_job(third)["status"] == "done"
        assert ran == []


def test_recover_interrupted(core_app):
    old = (datetime.now() - timedelta(minutes=5)).isoformat(timespec="seconds")
    fresh = datetime.now().isoformat(timespec="seconds")
    rows = [
        ("running", old, old),     # 止まったまま → 中断
        ("paused", old, old),      # 一時停止中のままアプリが終わった → 中断
        ("queued", None, old),     # 待機中のまま → 中断
        ("running", fresh, old),   # heartbeat が新しい → そのまま
        ("done", old, old),        # 終わったジョブ → そのまま
    ]
    with core_app.app_context():
        conn = db.connect()
        ids = []
        for status, beat, updated in rows:
            ids.append(conn.execute(
                "INSERT INTO jobs (kind, ref_type, ref_id, status, heartbeat_at, created_at, updated_at) "
                "VALUES ('table_confirm', 'table_import', 1, ?, ?, ?, ?)", (status, beat, updated, updated)).lastrowid)
        conn.commit()
        conn.close()

        assert jobs.recover_interrupted() == 3
        statuses = [jobs.get_job(i)["status"] for i in ids]
        assert statuses == ["interrupted", "interrupted", "interrupted", "running", "done"]
        job = jobs.get_job(ids[0])
        assert job["status_label"] == "中断" and "中断" in job["message"] and job["finished"]
        assert jobs.request_resume(ids[0]) is False
        assert jobs.recover_interrupted() == 0
