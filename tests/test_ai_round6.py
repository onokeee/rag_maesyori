"""AI まわり（6巡目）：429 の送り直しが行の残り時間で打ち切られてもレート制限として扱う、
キャッシュの読み書きが「database is locked」でもジョブを止めない。"""
import sqlite3
import time

import pytest

from aiproc import cache, items, runner
from core import jobs
from services import llm
from tests.test_aiproc import ROWS, _make_import, ai_app, fake  # noqa: F401  (fixture)


# ---- R6-AI-2 429 の後の送り直しが時間切れになっても、状態コード 429 のまま打ち切る ------------------------

def _fake_429_chat(latency=0.3, retry_after=0.6):
    def chat_raw(settings, messages, response_format=None, max_tokens=None, timeout=None):
        if timeout is not None and timeout < latency:
            time.sleep(timeout)
            raise llm.LLMCallError("retry", "AIの応答がタイムアウトしました。", None, None)
        time.sleep(latency)
        raise llm.LLMCallError("retry", "混み合っています（429）。", 429, retry_after)
    return chat_raw


def test_resend_cut_by_row_deadline_keeps_rate_limit_status(monkeypatch):
    monkeypatch.setattr(llm, "chat_raw", _fake_429_chat())
    with pytest.raises(llm.LLMCallError) as ei:
        runner.call_with_retry({"timeout": 10, "model": "m"}, [], None, 10, deadline=time.monotonic() + 1.0)
    assert ei.value.kind == "row"
    assert ei.value.status == 429                       # タイムアウト扱いにしない（一時停止の判定に使う）
    assert "混み合っています" in str(ei.value)


def test_execute_work_marks_deadline_after_429_as_rate_limited(monkeypatch):
    monkeypatch.setattr(llm, "chat_raw", _fake_429_chat())
    monkeypatch.setattr(runner, "row_deadline_seconds", lambda settings: 1.0)
    monkeypatch.setattr(runner.cache, "get", lambda key: None)
    work = runner.StageWork.__new__(runner.StageWork)
    work.__dict__.update(kind="log", schema=None, messages=[], key="k", max_tokens=10)
    out = runner.execute_work(work, {"timeout": 10, "model": "m"}, "json_object")
    assert out.status == "error" and out.rate_limited


def test_plain_timeout_without_429_is_not_rate_limited(monkeypatch):
    def chat_raw(settings, messages, response_format=None, max_tokens=None, timeout=None):
        time.sleep(min(timeout or 0.3, 0.3))
        raise llm.LLMCallError("retry", "AIの応答がタイムアウトしました。", None, None)

    monkeypatch.setattr(llm, "chat_raw", chat_raw)
    with pytest.raises(llm.LLMCallError) as ei:
        runner.call_with_retry({"timeout": 10, "model": "m"}, [], None, 10, deadline=time.monotonic() + 1.0)
    assert ei.value.kind == "row" and ei.value.status is None


# ---- R6C-3 キャッシュの読み書きが「database is locked」でもジョブを失敗にしない ----------------------------

def test_cache_lock_retries_then_succeeds(monkeypatch):
    monkeypatch.setattr(runner, "CACHE_DB_BACKOFF", 0)
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 2:
            raise sqlite3.OperationalError("database is locked")
        return "ok"

    assert runner._cache_db(flaky) == "ok" and len(calls) == 2


def test_cache_other_operational_error_is_not_swallowed(monkeypatch):
    def broken():
        raise sqlite3.OperationalError("no such table: llm_calls")

    with pytest.raises(sqlite3.OperationalError):
        runner._cache_db(broken)


def test_ai_job_finishes_when_cache_stays_locked(ai_app, fake, monkeypatch):
    monkeypatch.setattr(runner, "CACHE_DB_BACKOFF", 0)

    def locked(*a, **k):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(cache, "put_result", locked)
    rows = {k: ROWS[k] for k in ("R1", "R2")}
    iid = _make_import(ai_app, rows=rows)
    with ai_app.app_context():
        job = jobs.wait_job(runner.start_ai_job(iid, concurrency=2, stage_ids=["log"]), timeout=60)
        assert job["status"] == "done", job["message"]
        got = items.items_by_key(iid, "log", import_id=iid)
        assert len(got) == len(rows)
        assert not [v for v in got.values() if v["status"] == "error"]
