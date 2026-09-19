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


def _insert_job(status, beat, kind="table_render", ref_id=1):
    conn = db.connect()
    job_id = conn.execute(
        "INSERT INTO jobs (kind, ref_type, ref_id, status, heartbeat_at, created_at, updated_at) "
        "VALUES (?, 'table_import', ?, ?, ?, ?, ?)", (kind, ref_id, status, beat, beat, beat)).lastrowid
    conn.commit()
    conn.close()
    return job_id


def test_a_job_left_by_a_previous_process_is_interrupted_once_its_heartbeat_is_stale(core_app):
    """閉じてすぐ起動し直すと起動時の回復では拾えない（2分未満）。参照したときに持ち主がいなければ中断にする。"""
    recent = (datetime.now() - timedelta(seconds=40)).isoformat(timespec="seconds")
    old = (datetime.now() - timedelta(minutes=3)).isoformat(timespec="seconds")
    with core_app.app_context():
        left = _insert_job("running", recent)
        assert jobs.recover_interrupted() == 0          # 起動時: まだ新しいので残す
        assert jobs.get_job(left)["status"] == "running"
        conn = db.connect()
        conn.execute("UPDATE jobs SET heartbeat_at = ? WHERE id = ?", (old, left))   # 2分以上たった
        conn.commit()
        conn.close()
        job = jobs.latest_job("table_import", 1)
        assert job["status"] == "interrupted" and job["finished"] and "中断" in job["message"]


def test_a_queued_job_of_this_process_is_not_interrupted_while_it_waits(core_app, monkeypatch):
    release = threading.Event()
    with core_app.app_context():
        first = jobs.start_job("table_read", "table_import", 1, lambda ctx: release.wait(10) and None)
        second = jobs.start_job("table_render", "table_import", 2, lambda ctx: {"ok": True})
        _wait_until(lambda: jobs.get_job(first)["status"] == "running")
        monkeypatch.setattr(jobs, "STALE_AFTER", timedelta(seconds=-60))   # どれも「古い」とみなす
        assert jobs.get_job(second)["status"] == "queued"   # このプロセスのジョブなので中断にしない
        release.set()
        assert jobs.wait_job(second)["status"] == "done"


def test_a_long_queued_job_is_not_interrupted_by_a_second_launch(core_app, monkeypatch):
    """誤って2つ目を起動しても、その起動時の回復（別プロセス）が、1つ目で2分以上待っているジョブを中断にしない。"""
    monkeypatch.setattr(jobs, "HEARTBEAT_INTERVAL", 0.05)
    release = threading.Event()
    old = (datetime.now() - timedelta(minutes=5)).isoformat(timespec="seconds")
    with core_app.app_context():
        first = jobs.start_job("table_read", "table_import", 1, lambda ctx: release.wait(10) and None)
        second = jobs.start_job("table_render", "table_import", 2, lambda ctx: {"ok": True})
        _wait_until(lambda: jobs.get_job(first)["status"] == "running")
        conn = db.connect()
        conn.execute("UPDATE jobs SET heartbeat_at = NULL, created_at = ?, updated_at = ? WHERE id = ?",
                     (old, old, second))   # 5分前から待っている
        conn.commit()
        conn.close()
        # 動いているジョブの Ticker が、待機中のジョブの heartbeat も更新する
        _wait_until(lambda: (jobs.get_job(second)["heartbeat_at"] or "") > old)
        assert jobs.recover_interrupted() == 0   # 2つ目の起動時の回復（このプロセスのことは知らない）
        assert jobs.get_job(second)["status"] == "queued"
        release.set()
        assert jobs.wait_job(second)["status"] == "done"


def test_a_paused_ai_job_does_not_hold_up_other_imports(core_app):
    """AI整形を一時停止しても、ほかの取り込みの読み込みは進む。同じ取り込みの分は AI整形が終わるまで待つ。"""
    def ai(ctx):
        while True:
            if not ctx.wait_if_paused():
                return None
            time.sleep(0.01)

    with core_app.app_context():
        ai_job = jobs.start_job("ai_format", "table_import", 1, ai)
        _wait_until(lambda: jobs.get_job(ai_job)["status"] == "running")
        jobs.request_pause(ai_job)
        _wait_until(lambda: jobs.get_job(ai_job)["status"] == "paused")

        other = jobs.start_job("table_read", "table_import", 2, lambda ctx: {"rows": 3})
        assert jobs.wait_job(other, timeout=5)["status"] == "done"

        same = jobs.start_job("table_render", "table_import", 1, lambda ctx: {"files": 1})
        time.sleep(0.5)
        assert jobs.get_job(same)["status"] == "queued"    # 同じ取り込みの AI整形が終わるまで動かさない
        later = jobs.start_job("table_read", "table_import", 3, lambda ctx: {"rows": 1})
        assert jobs.wait_job(later, timeout=5)["status"] == "done"   # 後回しの分がほかを止めない

        jobs.request_cancel(ai_job)
        assert jobs.wait_job(ai_job)["status"] == "cancelled"
        assert jobs.wait_job(same, timeout=5)["status"] == "done"


def test_a_job_waiting_behind_a_paused_ai_job_says_what_it_waits_for(core_app):
    """一時停止中の AI整形の後ろで待つジョブは、理由（何を待っているか・どうすれば動くか）を出す。"""
    def ai(ctx):
        while True:
            if not ctx.wait_if_paused():
                return None
            time.sleep(0.01)

    with core_app.app_context():
        ai_job = jobs.start_job("ai_format", "table_import", 11, ai)
        _wait_until(lambda: jobs.get_job(ai_job)["status"] == "running")
        jobs.request_pause(ai_job)
        _wait_until(lambda: jobs.get_job(ai_job)["status"] == "paused")

        same = jobs.start_job("table_preview", "table_import", 11, lambda ctx: {"files": 1})
        other_ai = jobs.start_job("ai_format", "table_import", 12, lambda ctx: {"rows": 1})
        time.sleep(0.3)
        waiting = jobs.get_job(same)
        assert waiting["status"] == "queued"
        assert "この取り込みのAI整形が一時停止中" in waiting["message"] and "再開するか中止" in waiting["message"]
        assert waiting["waiting_for"]["job_id"] == ai_job and waiting["waiting_for"]["same_ref"]
        behind = jobs.latest_job("table_import", 12, kind="ai_format")
        assert behind["status"] == "queued"
        assert "別の取り込みのAI整形が一時停止中" in behind["message"]
        assert behind["waiting_for"]["ref_id"] == 11 and not behind["waiting_for"]["same_ref"]
        assert "message" not in jobs.get_job(ai_job) or "待っています" not in jobs.get_job(ai_job)["message"]

        jobs.request_cancel(ai_job)
        assert jobs.wait_job(same, timeout=5)["status"] == "done"
        assert jobs.wait_job(other_ai, timeout=5)["status"] == "done"
        assert jobs.get_job(same).get("waiting_for") is None


def test_a_job_whose_last_status_write_fails_ends_as_failed(core_app, monkeypatch):
    import sqlite3

    real = jobs.JobContext.progress

    def locked(self, **kw):
        if "result" in kw:
            raise sqlite3.OperationalError("database is locked")
        return real(self, **kw)

    monkeypatch.setattr(jobs.JobContext, "progress", locked)
    with core_app.app_context():
        job_id = jobs.start_job("table_render", "table_import", 1, lambda ctx: {"files": 1})
        job = jobs.wait_job(job_id)
        assert job["status"] == "failed" and "エラー" in job["message"] and job["finished"]
