"""6巡目の修正の、担当をまたぐ残り。

- R6-2 の続き: 消し切れなかったときは、削除の案内でも「消しました」と言い切らない（一覧表・帳票）
- UX6-3 の続き: 帳票の取り込み画面にも、保存先フォルダに保存したときも消えることを出す
- ホームのまとまりの各帳票の［フォルダに保存］は、forms の確認文（BATCH_MEMBER_SAVE_CONFIRM）を使う
"""
import html

from core import purge
from tests.test_output_folder import _set_folder, out_dir  # noqa: F401
from tests.test_retention import _add_confirmed_document, _confirmed_import
from views.forms import BATCH_MEMBER_SAVE_CONFIRM


def _flashes(client) -> str:
    return client.get("/").get_data(as_text=True)


def test_table_delete_says_files_were_emptied_when_not_fully_removed(app, client, monkeypatch):
    import_id = _confirmed_import(app, client)
    monkeypatch.setattr(purge.shutil, "rmtree", lambda *a, **k: None)
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 302
    page = _flashes(client)
    assert "消し切れず" in page and "作った Markdown を消しました" not in page


def test_table_delete_normal_wording_is_kept(app, client):
    import_id = _confirmed_import(app, client)
    client.post(f"/tables/imports/{import_id}/delete")
    page = _flashes(client)
    assert "作った Markdown を消しました" in page and "消し切れず" not in page


def test_form_delete_says_files_were_emptied_when_not_fully_removed(app, client, monkeypatch):
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    monkeypatch.setattr(purge, "remove_upload", lambda _p: False)
    client.post(f"/forms/{doc_id}/delete")
    page = _flashes(client)
    assert "消し切れず" in page and "元のファイルと読み取り結果を消しました" not in page


def test_form_upload_page_mentions_the_save_folder(client, out_dir):  # noqa: F811
    assert "保存先フォルダに保存したときも同じです" not in client.get("/forms/new").get_data(as_text=True)
    _set_folder(client, out_dir)
    assert "保存先フォルダに保存したときも同じです" in client.get("/forms/new").get_data(as_text=True)


def test_home_batch_member_save_uses_the_forms_wording(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    _add_confirmed_document(app, "a.xlsx", batch_id="B1", order=0)
    _add_confirmed_document(app, "b.xlsx", value="EQ-002", batch_id="B1", order=1)
    page = html.unescape(client.get("/").get_data(as_text=True))
    assert BATCH_MEMBER_SAVE_CONFIRM in page


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
