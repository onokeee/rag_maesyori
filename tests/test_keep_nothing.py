"""「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）の確認。

社内LANに置いて数人が同時に使うので、捨てるのは「その人の分」だけで、ほかの人の作業には触らない。
捨てる機会は4つ（core/purge.py）:
  - 画面を離れたとき（ブラウザが /forms/discard・/tables/discard に伝える）… purge_session / discard_*
  - 新しいファイルを置いたとき（同じ人の前の分）                          … discard_documents / discard_table_imports
  - しばらくさわられていないとき（IDLE_HOURS）                            … sweep_stale
  - 起動時                                                                … purge_all_pending
あとの2つは、動いているジョブが付いているものには手を出さない。
"""
from datetime import datetime, timedelta
from pathlib import Path

import pytest

from core import purge
from models import database as db
from views import SESSION_ID_KEY
from tests.conftest import add_confirmed_document

# views.current_session_id() が配る id と同じ形（32桁）にする（そうでないとブラウザ側で作り直される）
SESSION_A = "a" * 32
SESSION_B = "b" * 32


# ---- 下ごしらえ ---------------------------------------------------------------------------

def _ensure_session_column(app) -> None:
    """持ち主の列（views.current_session_id() が入れる）がまだ無い DB でも試せるようにする。"""
    with app.app_context():
        conn = db.get_db()
        for table in ("documents", "table_imports"):
            names = {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")')}
            if "session_id" not in names:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN session_id TEXT")
        conn.commit()


def _own(app, table: str, row_id: int, session_id: str) -> None:
    with app.app_context():
        conn = db.get_db()
        conn.execute(f"UPDATE {table} SET session_id = ? WHERE id = ?", (session_id, row_id))
        conn.commit()


def _document(app, name: str, session_id: str) -> tuple[int, Path]:
    doc_id, path = add_confirmed_document(app, name)
    _own(app, "documents", doc_id, session_id)
    return doc_id, path


def _import(app, session_id: str, *, file_name: str = "一覧.csv") -> int:
    """一覧表の取り込みを1件作る（読み込みまでは進めない。捨てられるかどうかだけを見る）。"""
    with app.app_context():
        stored = f"tables/{session_id}-{file_name}"
        path = Path(app.config["UPLOAD_DIR"]) / stored
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"a,b\r\n1,2\r\n")
        conn = db.get_db()
        now = datetime.now().isoformat(timespec="seconds")
        import_id = conn.execute(
            "INSERT INTO table_imports (file_name, file_hash, stored_path, created_at, updated_at, session_id) "
            "VALUES (?, ?, ?, ?, ?, ?)", (file_name, "0" * 64, stored, now, now, session_id)).lastrowid
        conn.commit()
    folder = _import_folder(app, import_id)
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "rows.jsonl").write_text('{"a": 1}\n', encoding="utf-8")
    return import_id


def _import_folder(app, import_id: int) -> Path:
    with app.app_context():
        return purge.import_dir(import_id)


def _running_job(app, ref_type: str, ref_id: int, kind: str = "table_read") -> int:
    with app.app_context():
        conn = db.get_db()
        now = datetime.now().isoformat(timespec="seconds")
        job_id = conn.execute(
            "INSERT INTO jobs (kind, ref_type, ref_id, status, created_at, updated_at) "
            "VALUES (?, ?, ?, 'running', ?, ?)", (kind, ref_type, ref_id, now, now)).lastrowid
        conn.commit()
        return job_id


def _age(app, table: str, row_id: int, hours: float) -> None:
    """その行を hours 時間前にさわられたことにする。"""
    old = (datetime.now() - timedelta(hours=hours)).isoformat(timespec="seconds")
    with app.app_context():
        conn = db.get_db()
        names = {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")')}
        sets = ", ".join(f"{c} = ?" for c in ("created_at", "updated_at", "confirmed_at") if c in names)
        args = [old] * sets.count("?") + [row_id]
        conn.execute(f"UPDATE {table} SET {sets} WHERE id = ?", args)
        conn.commit()


def _alive(app, table: str, row_id: int) -> bool:
    with app.app_context():
        return db.get_db().execute(f"SELECT 1 FROM {table} WHERE id = ?", (row_id,)).fetchone() is not None


@pytest.fixture
def sessions(app):
    _ensure_session_column(app)
    return app


# ---- 1 画面を離れたら捨てる ------------------------------------------------------------------

def test_leaving_the_page_discards_that_persons_work_only(sessions):
    """画面を離れた人の分だけを捨て、同時に使っているほかの人の分は残す。"""
    app = sessions
    mine, my_file = _document(app, "mine.xlsx", SESSION_A)
    yours, your_file = _document(app, "yours.xlsx", SESSION_B)
    my_import = _import(app, SESSION_A)
    your_import = _import(app, SESSION_B)
    my_folder = _import_folder(app, my_import)

    with app.app_context():
        assert purge.purge_session(SESSION_A) == (1, 1)

    assert not _alive(app, "documents", mine)
    assert not _alive(app, "table_imports", my_import)
    assert not my_file.exists()
    assert not my_folder.exists()          # imports/<id>/（読み込んだ行・控え・作った md）ごと消える
    assert _alive(app, "documents", yours)
    assert _alive(app, "table_imports", your_import)
    assert your_file.exists()


def test_discarding_twice_does_nothing_the_second_time(sessions):
    """画面を閉じる合図は2回届くことがある（pagehide と visibilitychange）。2回目は何もしない。"""
    app = sessions
    doc_id, _ = _document(app, "mine.xlsx", SESSION_A)
    import_id = _import(app, SESSION_A)
    with app.app_context():
        assert purge.discard_documents([doc_id], SESSION_A) == 1
        assert purge.discard_documents([doc_id], SESSION_A) == 0
        assert purge.discard_table_imports([import_id], SESSION_A) == 1
        assert purge.discard_table_imports([import_id], SESSION_A) == 0
        assert purge.purge_session(SESSION_A) == (0, 0)


def test_one_person_cannot_discard_another_persons_work(sessions):
    """番号を知っていても、ほかの人の取り込みは捨てられない。"""
    app = sessions
    yours, _ = _document(app, "yours.xlsx", SESSION_B)
    your_import = _import(app, SESSION_B)
    with app.app_context():
        assert purge.discard_documents([yours], SESSION_A) == 0
        assert purge.discard_table_imports([your_import], SESSION_A) == 0
    assert _alive(app, "documents", yours)
    assert _alive(app, "table_imports", your_import)


def test_rows_without_an_owner_are_never_swept_by_session(app):
    """持ち主の分からない行（この仕組みを入れる前のDB・直接作った行）は、まとめては捨てない。

    誰のものか分からないものを「自分の分」として巻き込むと、同時に使っている人の作業が消える。
    こういう行は起動時の片付け（purge_all_pending）と時間切れ（sweep_stale）で消える。
    """
    doc_id, _ = add_confirmed_document(app, "mine.xlsx")
    with app.app_context():
        assert purge.purge_session(SESSION_A) == (0, 0)
    assert _alive(app, "documents", doc_id)


# ---- 2 新しいファイルを置いたら前の分を捨てる ----------------------------------------------------

def test_a_new_upload_discards_the_previous_unfinished_one(sessions):
    """同じ人が新しいファイルを置いたら、前の（ダウンロードしていない）分はその場で消える。"""
    app = sessions
    first, first_file = _document(app, "first.xlsx", SESSION_A)
    first_import = _import(app, SESSION_A, file_name="first.csv")

    # 画面（static/review.js・static/tables.js）は新しいファイルを送る前にこれを呼ぶ
    with app.app_context():
        assert purge.discard_documents([first], SESSION_A) == 1
        assert purge.discard_table_imports([first_import], SESSION_A) == 1

    second, second_file = _document(app, "second.xlsx", SESSION_A)
    assert not _alive(app, "documents", first)
    assert not first_file.exists()
    assert not _alive(app, "table_imports", first_import)
    assert _alive(app, "documents", second)
    assert second_file.exists()


# ---- 3 処理中のものには手を出さない ----------------------------------------------------------

def test_a_running_job_is_not_discarded(sessions):
    """AI整形や読み込みが動いている取り込みは、画面を離れても・時間切れでも捨てない。"""
    app = sessions
    doc_id, _ = _document(app, "busy.xlsx", SESSION_A)
    import_id = _import(app, SESSION_A)
    _running_job(app, "document", doc_id, kind="form_read")
    _running_job(app, "table_import", import_id, kind="ai_format")

    with app.app_context():
        assert purge.purge_session(SESSION_A) == (0, 0)
        assert purge.discard_documents([doc_id], SESSION_A) == 0
        assert purge.discard_table_imports([import_id], SESSION_A) == 0
    assert _alive(app, "documents", doc_id)
    assert _alive(app, "table_imports", import_id)

    # 時間切れの掃除（IDLE_HOURS）も起動時の片付けも、動いているジョブには触らない
    _age(app, "documents", doc_id, purge.IDLE_HOURS + 1)
    _age(app, "table_imports", import_id, purge.IDLE_HOURS + 1)
    with app.app_context():
        assert purge.sweep_stale(purge.IDLE_HOURS) == (0, 0)
        assert purge.purge_all_pending() == (0, 0)
    assert _alive(app, "documents", doc_id)
    assert _alive(app, "table_imports", import_id)


def test_the_idle_sweep_discards_what_is_left(sessions):
    """合図が届かなかった分（ブラウザの強制終了・LANの切断）は、IDLE_HOURS で片付く。"""
    app = sessions
    old_doc, old_file = _document(app, "old.xlsx", SESSION_A)
    old_import = _import(app, SESSION_A, file_name="old.csv")
    fresh_doc, _ = _document(app, "fresh.xlsx", SESSION_B)
    _age(app, "documents", old_doc, purge.IDLE_HOURS + 1)
    _age(app, "table_imports", old_import, purge.IDLE_HOURS + 1)

    with app.app_context():
        assert purge.sweep_stale(purge.IDLE_HOURS) == (1, 1)
    assert not _alive(app, "documents", old_doc)
    assert not old_file.exists()
    assert not _alive(app, "table_imports", old_import)
    assert _alive(app, "documents", fresh_doc)   # いま使っている人の分は残る


def test_the_idle_timeout_is_short_and_in_one_place():
    """放っておかれた分を捨てるまでの時間は定数1つ（変えやすくする）。"""
    assert purge.IDLE_HOURS <= 4
    assert purge.STALE_HOURS == purge.IDLE_HOURS   # 旧名（app.py が参照している）


# ---- 4 ダウンロードした分はもう無い -----------------------------------------------------------

def test_a_downloaded_document_is_already_gone(sessions, client):
    """ダウンロードが終わった時点で消えている（purge_after_send）ので、捨てるものは残らない。"""
    app = sessions
    doc_id, path = _document(app, "download.xlsx", SESSION_A)
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A          # この帳票を取り込んだブラウザとして開く
    res = client.get(f"/forms/{doc_id}/download.md")
    assert res.status_code == 200

    assert not _alive(app, "documents", doc_id)
    assert not path.exists()
    with app.app_context():
        assert purge.purge_session(SESSION_A) == (0, 0)
        assert purge.discard_documents([doc_id], SESSION_A) == 0


# ---- 5 画面に1行で書いてある ----------------------------------------------------------------

@pytest.mark.parametrize("path", ["/forms", "/tables", "/form-types"])
def test_every_page_says_what_happens_when_you_leave(client, path):
    res = client.get(path, follow_redirects=True)
    assert res.status_code == 200
    html = res.get_data(as_text=True)
    assert "画面を閉じると" in html
    assert "捨てられます" in html


@pytest.mark.parametrize("path,url", [("/forms", "/forms/discard"), ("/tables", "/tables/discard")])
def test_the_pages_know_where_to_send_the_discard(client, path, url):
    """画面を離れたときの送り先（static/app.js の ragDiscard が使う）。"""
    assert url in client.get(path, follow_redirects=True).get_data(as_text=True)


# ---- 6 画面からの合図（/forms/discard・/tables/discard） ------------------------------------
# ブラウザは navigator.sendBeacon で送る。中身の型は text/plain で、応答は読めず、やり直しもできない。
# だからサーバーはいつでも 204 を返し、捨てられないもの（もう無い・ほかの人の・処理中）は黙って外す。

def _as_beacon(client, url: str, body: str):
    """sendBeacon と同じ送り方（text/plain・同一サイトからの POST）。"""
    return client.post(url, data=body, content_type="text/plain;charset=UTF-8",
                       headers={"Sec-Fetch-Site": "same-origin"})


def test_the_leave_signal_discards_only_the_senders_work(sessions, client):
    """画面を離れた合図で、その番号の帳票だけを捨てる。ほかの人の分は番号を指しても残る。"""
    app = sessions
    mine, my_file = _document(app, "mine.xlsx", SESSION_A)
    yours, your_file = _document(app, "yours.xlsx", SESSION_B)
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A

    res = _as_beacon(client, "/forms/discard", f'{{"doc_ids": [{mine}, {yours}]}}')
    assert res.status_code == 204
    assert res.get_data() == b""

    assert not _alive(app, "documents", mine)
    assert not my_file.exists()
    assert _alive(app, "documents", yours)      # ほかの人の分は、番号を書かれても捨てない
    assert your_file.exists()


def test_the_leave_signal_discards_the_senders_table_import(sessions, client):
    app = sessions
    mine = _import(app, SESSION_A)
    yours = _import(app, SESSION_B, file_name="ほかの人.csv")
    folder = _import_folder(app, mine)
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A

    assert _as_beacon(client, "/tables/discard", f'{{"import_ids": [{mine}, {yours}]}}').status_code == 204
    assert not _alive(app, "table_imports", mine)
    assert not folder.exists()
    assert _alive(app, "table_imports", yours)


def test_the_leave_signal_never_fails(sessions, client):
    """もう無い番号・中身の無い body・壊れた body でも 204（ブラウザは答えを読めず、やり直せない）。"""
    app = sessions
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A
    for url, body in [("/forms/discard", '{"doc_ids": [999999]}'), ("/tables/discard", '{"import_ids": [999999]}'),
                      ("/forms/discard", ""), ("/tables/discard", "これはJSONではない")]:
        assert _as_beacon(client, url, body).status_code == 204, (url, body)


def test_the_forms_signal_does_not_touch_the_tables_tab(sessions, client):
    """同じブラウザで表の取り込みも開いていることがある。帳票の合図で表の分まで捨てない（逆も同じ）。"""
    app = sessions
    doc_id, _ = _document(app, "mine.xlsx", SESSION_A)
    import_id = _import(app, SESSION_A)
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A

    assert _as_beacon(client, "/forms/discard", "{}").status_code == 204   # 番号なし＝「自分の分ぜんぶ」
    assert not _alive(app, "documents", doc_id)
    assert _alive(app, "table_imports", import_id)

    assert _as_beacon(client, "/tables/discard", "{}").status_code == 204
    assert not _alive(app, "table_imports", import_id)


def test_the_leave_signal_does_not_pull_a_running_job_out(sessions, client):
    """読み込み中に画面を閉じても、動いているジョブの取り込みは捨てない（時間切れの片付けに任せる）。"""
    app = sessions
    import_id = _import(app, SESSION_A)
    _running_job(app, "table_import", import_id)
    with client.session_transaction() as cookie:
        cookie[SESSION_ID_KEY] = SESSION_A

    assert _as_beacon(client, "/tables/discard", f'{{"import_ids": [{import_id}]}}').status_code == 204
    assert _alive(app, "table_imports", import_id)
