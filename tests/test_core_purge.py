"""ダウンロードで消す処理（core/purge.py）の、本番のサーバー（waitress）での動きと番号の続き（design.md 3.3）。"""
import io
import os
import socket
import struct
import threading
import time

from flask import send_file
from waitress.server import create_server

import app as app_module
from core import purge
from models import database as db
from tests.test_retention import _add_confirmed_document

BIG = 1024 * 1024   # 溜められる上限（OUTBUF_HIGH_WATERMARK）より十分大きい本文


def _serve(app):
    """本番と同じ waitress の設定で、このテストの間だけ空いているポートで動かす。"""
    server = create_server(app, host="127.0.0.1", port=0, **app_module.WAITRESS_OPTIONS)
    threading.Thread(target=server.run, daemon=True).start()
    return server


def _add_big_download(app, purged: list, finished: threading.Event):
    body = os.urandom(BIG)

    @app.get("/_test/big.zip")
    def _big():
        res = send_file(io.BytesIO(body), mimetype="application/zip", as_attachment=True, download_name="big.zip",
                        conditional=False)
        res = purge.purge_after_send(res, purged.append, "消した")
        res.call_on_close(finished.set)   # 応答を閉じた（送り終えた・切れた）ことをテストに知らせる
        return res


def _request(port: int) -> socket.socket:
    sock = socket.create_connection(("127.0.0.1", port))
    sock.sendall(f"GET /_test/big.zip HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nConnection: close\r\n\r\n".encode())
    return sock


def test_the_server_does_not_buffer_a_whole_download_ahead():
    """waitress の既定（16MB）だと、16MB 未満の zip は受け取られる前にアプリからは送り終えたように見える。"""
    assert app_module.WAITRESS_OPTIONS["outbuf_high_watermark"] <= 64 * 1024


def test_a_large_download_cut_off_under_waitress_keeps_the_data(app):
    purged, finished = [], threading.Event()
    _add_big_download(app, purged, finished)
    server = _serve(app)
    try:
        sock = _request(server.effective_port)
        assert sock.recv(1)   # 本文の最初だけ受け取って、ブラウザを閉じたように切る（RST）
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_LINGER, struct.pack("ii", 1, 0))
        sock.close()
        assert finished.wait(10)
        time.sleep(0.2)
        assert purged == []
    finally:
        server.close()


def test_a_large_download_received_to_the_end_under_waitress_removes_the_item(app):
    purged, finished = [], threading.Event()
    _add_big_download(app, purged, finished)
    server = _serve(app)
    try:
        sock = _request(server.effective_port)
        received = b""
        while chunk := sock.recv(65536):
            received += chunk
        sock.close()
        assert len(received.split(b"\r\n\r\n", 1)[1]) == BIG
        assert finished.wait(10)
        time.sleep(0.2)
        assert purged == ["消した"]
    finally:
        server.close()


# ---- 番号の続き（sqlite_sequence）を残さない ---------------------------------------------

def _sequence(app, table: str):
    with app.app_context():
        row = db.get_db().execute("SELECT seq FROM sqlite_sequence WHERE name = ?", (table,)).fetchone()
        return None if row is None else row[0]


def test_the_id_counter_does_not_record_how_many_forms_were_taken_in(app, client):
    ids = [_add_confirmed_document(app, f"報告書{i}.xlsx")[0] for i in range(3)]
    assert _sequence(app, "documents") == max(ids)
    for doc_id in ids[:-1]:
        assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
    assert _sequence(app, "documents") == max(ids)   # まだ作業中の帳票があるうちはそのまま

    assert client.get(f"/forms/{ids[-1]}/download.md").status_code == 200
    seq = _sequence(app, "documents")
    assert purge._ID_BASE_MIN <= seq < purge._ID_BASE_MAX   # 空になったら乱数に置き換わる

    # 次の帳票は消した番号を使い回さないので、開いたままの古い画面は新しい帳票を指さない
    new_id, _ = _add_confirmed_document(app, "次の報告書.xlsx")
    assert new_id == seq + 1 and new_id not in ids
    for doc_id in ids:
        assert client.get(f"/forms/{doc_id}/review").status_code == 404
        assert client.get(f"/forms/{doc_id}/download.md").status_code == 404
    assert new_id < 2 ** 53   # 画面の JavaScript でも正確に扱える


def test_the_id_counter_of_table_imports_and_jobs_is_forgotten_too(app, client):
    from tests.test_retention import _confirmed_import

    import_id = _confirmed_import(app, client)
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    for table in ("table_imports", "jobs"):
        seq = _sequence(app, table)
        assert purge._ID_BASE_MIN <= seq < purge._ID_BASE_MAX, table
    assert client.get(f"/tables/imports/{import_id}").status_code == 404


def test_the_counter_is_randomised_each_time_the_table_becomes_empty(app, client):
    seen = set()
    for i in range(3):
        doc_id, _ = _add_confirmed_document(app, f"報告書{i}.xlsx")
        assert doc_id not in seen
        seen.add(doc_id)
        assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
        seen.add(_sequence(app, "documents"))
    assert len(seen) == 6


def test_a_new_database_still_starts_at_one(app):
    with app.app_context():
        purge.forget_id_counters(db.get_db())   # 一度も帳票を入れていない表には番号の続きが無く、そのまま
    doc_id, _ = _add_confirmed_document(app, "最初.xlsx")
    assert doc_id == 1


def test_startup_forgets_a_counter_left_by_an_older_version(tmp_path):
    from tests.conftest import make_config

    first = app_module.create_app(make_config(tmp_path))
    with first.app_context():
        conn = db.get_db()
        conn.execute("INSERT INTO documents (file_name, file_hash, stored_path, created_at) "
                     "VALUES ('a.xlsx', '0', 'documents/a.xlsx', '2026-09-19')")
        conn.execute("DELETE FROM documents")   # 前の版は消しても番号の続き（1）を残していた
        conn.commit()
        assert conn.execute("SELECT seq FROM sqlite_sequence WHERE name = 'documents'").fetchone()[0] == 1

    again = app_module.create_app(make_config(tmp_path))
    assert purge._ID_BASE_MIN <= _sequence(again, "documents") < purge._ID_BASE_MAX


def test_the_counter_is_kept_while_the_table_still_has_rows(app):
    first, _ = _add_confirmed_document(app, "作業中.xlsx")
    with app.app_context():
        purge.forget_id_counters(db.get_db())
    assert _sequence(app, "documents") == first
