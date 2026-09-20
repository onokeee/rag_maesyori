"""社内LANで数人が同時に使うときの分かれ方（2026-09-20 の利用者の指示）。

- 取り込んだ帳票・一覧表は、置いたブラウザ（セッションのクッキー）のものだけが見える・触れる。
  ほかのブラウザからは「無い」として扱う（403 ではなく 404。あることも知らせない）。
- ジョブの進み具合も、その取り込みを持っているブラウザにしか見せない。
- 「設定」（帳票の種類・一覧表の取り込み設定）はみんなで使うので分けない。
- 読み込みは同時に動く（誰かの大きい表で、ほかの人が待たされない）。
"""
import threading

import pytest

from core import jobs
from tables import store
from tests.conftest import BufferedClient
from tests.tables_helpers import imported, panel_html, upload_csv
from tests.test_forms_flow import upload_forms


@pytest.fixture
def other_client(app):
    """もう1台のPC（別のブラウザ）。クッキーが別なので作業場所も別になる。"""
    app.test_client_class = BufferedClient
    return app.test_client()


# ---- 一覧表の取り込み --------------------------------------------------------------

def test_another_browser_cannot_see_or_delete_an_import(app, client, other_client):
    import_id = upload_csv(client, "他人の一覧.csv")

    for path in (f"/tables/imports/{import_id}/panel/source", f"/tables/imports/{import_id}/panel/preview",
                 f"/tables/imports/{import_id}/download.zip", f"/tables/imports/{import_id}/issues.csv"):
        assert other_client.get(path).status_code == 404, path
    for path in (f"/tables/imports/{import_id}/source", f"/tables/imports/{import_id}/layout",
                 f"/tables/imports/{import_id}/columns", f"/tables/imports/{import_id}/read",
                 f"/tables/imports/{import_id}/confirm", f"/tables/imports/{import_id}/preview",
                 f"/tables/imports/{import_id}/cancel", f"/tables/imports/{import_id}/ai/run",
                 f"/tables/imports/{import_id}/ai/trial", f"/tables/imports/{import_id}/delete"):
        assert other_client.post(path, json={}).status_code == 404, path

    with app.app_context():   # 消されていない（持ち主はそのまま使える）
        assert store.get_import(import_id) is not None
    assert client.get(f"/tables/imports/{import_id}/panel/source").status_code == 200


def test_another_browsers_job_is_invisible(app, client, other_client):
    """ほかのブラウザの取り込みのジョブは、進み具合も見せない。"""
    import_id = imported(app, client, "進み具合.csv", "同時実行テスト")
    with app.app_context():
        job_id = store.get_import(import_id)["job_id"]

    assert client.get(f"/api/jobs/{job_id}").status_code == 200
    assert other_client.get(f"/api/jobs/{job_id}").status_code == 404


# ---- 帳票の取り込み ----------------------------------------------------------------

def test_another_browser_cannot_see_or_delete_a_form(app, client, other_client, sample_dir):
    doc_id = upload_forms(client, sample_dir / "standard.xlsx")[0]

    assert other_client.get(f"/forms/{doc_id}/type").status_code == 404
    assert other_client.get(f"/forms/{doc_id}/type").status_code == 404
    assert other_client.get(f"/forms/{doc_id}/download.md").status_code == 404
    assert other_client.get(f"/forms/{doc_id}/original").status_code == 404
    assert other_client.post(f"/forms/{doc_id}/draft", json={}).status_code == 404
    assert other_client.post(f"/forms/{doc_id}/delete").status_code == 404
    # 番号を並べて送っても、ほかのブラウザの帳票は「確定してダウンロード」の欄に出てこない
    assert other_client.get(f"/forms/finish?ids={doc_id}").get_json()["total"] == 0

    from models import database as db
    with app.app_context():
        assert db.get_document(doc_id) is not None
    assert client.get(f"/forms/{doc_id}/type").status_code == 200


# ---- 設定は分けない ----------------------------------------------------------------

def test_settings_are_shared_between_browsers(app, client, other_client, sample_dir):
    """帳票の種類はみんなで使う設定なので、別のブラウザからも見える。表の取り込み設定は保存しない。"""
    from tests.test_forms_flow import add_field, create_type

    import_id = imported(app, client, "みんなの一覧.csv", "この取り込みだけの名前")
    pattern_id = create_type(client, sample_dir / "standard.xlsx", name="共有の帳票の種類")
    add_field(client, pattern_id, "修理報告書", "A4")

    other_import = upload_csv(other_client, "別の人の一覧.csv")
    # 表の設定は取り込みの中にしかないので、ほかのブラウザには名前も出ない
    assert "この取り込みだけの名前" not in panel_html(other_client, other_import, "source")
    assert other_client.get(f"/tables/imports/{import_id}/panel/columns").status_code == 404
    panel = other_client.get(f"/form-types/{pattern_id}/panel")
    assert panel.status_code == 200 and "共有の帳票の種類" in panel.get_json()["html"]


# ---- 同時に動く --------------------------------------------------------------------

def test_two_reads_run_at_the_same_time(app):
    """別々の取り込みの読み込みは同時に動く（1本ずつなら待ち合わせが成立せず時間切れになる）。"""
    both_started = threading.Barrier(2, timeout=10)

    def work(ctx):
        both_started.wait()      # もう1つが動き出すまで待つ
        return {"ok": True}

    with app.app_context():
        first = jobs.start_job("table_read", "table_import", 101, work)
        second = jobs.start_job("table_read", "table_import", 102, work)
        assert jobs.wait_job(first, timeout=20)["status"] == "done"
        assert jobs.wait_job(second, timeout=20)["status"] == "done"


def test_two_jobs_never_touch_the_same_import(app):
    """同じ取り込みのジョブは、ワーカーが増えても同時には動かさない。"""
    inside, overlapped = threading.Semaphore(0), []
    release = threading.Event()

    def work(ctx):
        inside.release()
        if not release.wait(5):
            return {"ok": False}
        return {"ok": True}

    def peek(ctx):
        overlapped.append(jobs.get_job(first)["status"])
        return {"ok": True}

    with app.app_context():
        first = jobs.start_job("table_read", "table_import", 201, work)
        assert inside.acquire(timeout=10)
        second = jobs.start_job("table_render", "table_import", 201, peek)
        assert jobs.get_job(second)["status"] == "queued"
        release.set()
        assert jobs.wait_job(second, timeout=20)["status"] == "done"
        assert overlapped and overlapped[0] in ("done", "cancelled")   # 前のジョブが終わってから動いた
