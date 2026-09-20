"""6巡目の修正の、担当をまたぐ残り。

- R6-2 の続き: 消し切れなかったときは「消しました」と言い切らない（一覧表）／
  消し切れなかったことを画面に伝える（帳票。画面は fetch の答えで文面を決める）
"""
from core import purge
from tests.conftest import add_confirmed_document, confirmed_import


def test_table_delete_says_files_were_emptied_when_not_fully_removed(app, client, monkeypatch):
    import_id = confirmed_import(app, client)
    monkeypatch.setattr(purge.shutil, "rmtree", lambda *a, **k: None)
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 200
    message = res.get_json()["message"]
    assert "消し切れず" in message and "作った Markdown を消しました" not in message


def test_table_delete_normal_wording_is_kept(app, client):
    import_id = confirmed_import(app, client)
    message = client.post(f"/tables/imports/{import_id}/delete").get_json()["message"]
    assert "作った Markdown を消しました" in message and "消し切れず" not in message


def test_form_delete_reports_that_files_were_not_fully_removed(app, client, monkeypatch):
    """帳票の削除でも、消し切れなかったことを画面に伝える（core.purge.purge_incomplete の目印）。"""
    doc_id, _path = add_confirmed_document(app, "報告書.xlsx")
    monkeypatch.setattr(purge, "remove_upload", lambda _p: False)
    res = client.post(f"/forms/{doc_id}/delete")
    assert res.status_code == 200 and res.get_json() == {"ok": True, "incomplete": True}


def test_form_delete_reports_a_clean_removal(app, client):
    doc_id, _path = add_confirmed_document(app, "報告書.xlsx")
    res = client.post(f"/forms/{doc_id}/delete")
    assert res.status_code == 200 and res.get_json() == {"ok": True, "incomplete": False}


# ---- R6C-3 の続き（core/jobs.py）: 進捗の書き込みがロック中でもジョブを失敗にしない -----------------------

def _ctx(app, monkeypatch):
    import sqlite3

    from core import jobs
    monkeypatch.setattr(jobs, "SOFT_UPDATE_BACKOFF", 0)
    with app.app_context():
        from models import database
        conn = database.connect()
        conn.execute("INSERT INTO jobs (kind, status, params_json, progress_json, created_at, updated_at) "
                     "VALUES ('test', 'running', '{}', '{}', '2026-01-01', '2026-01-01')")
        conn.commit()
        job_id = conn.execute("SELECT max(id) FROM jobs").fetchone()[0]
        conn.close()
        return jobs.JobContext(job_id), sqlite3


def test_progress_retries_then_writes_when_the_lock_clears(app, monkeypatch):
    ctx, sqlite3 = _ctx(app, monkeypatch)
    real, calls = ctx._update, []

    def flaky(*a):
        calls.append(1)
        if len(calls) <= 2:
            raise sqlite3.OperationalError("database is locked")
        return real(*a)
    monkeypatch.setattr(ctx, "_update", flaky)
    ctx.progress(done=5)
    assert len(calls) == 3
    assert '"done": 5' in ctx._row()["progress_json"]
    ctx.close()


def test_progress_is_skipped_not_raised_while_the_db_stays_locked(app, monkeypatch):
    ctx, sqlite3 = _ctx(app, monkeypatch)

    def locked(*a):
        raise sqlite3.OperationalError("database is locked")
    monkeypatch.setattr(ctx, "_update", locked)
    ctx.progress(done=1)
    ctx.heartbeat()
    ctx.message("途中")
    ctx.close()


def test_other_db_errors_still_raise(app, monkeypatch):
    import pytest
    ctx, sqlite3 = _ctx(app, monkeypatch)

    def broken(*a):
        raise sqlite3.OperationalError("no such column: x")
    monkeypatch.setattr(ctx, "_update", broken)
    with pytest.raises(sqlite3.OperationalError):
        ctx.progress(done=1)
    ctx.close()
