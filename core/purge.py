"""取り込んだデータを消す（design.md 3.3「データを残さない」）。

このアプリはダウンロードが終わった時点で、その取り込みに属するものをすべて消す。
消すもの: アップロードした元のファイル、imports/<id>/（読み込んだ行・控え・作った md）、
          DB の行（documents / table_imports / jobs / ai_items / llm_calls とそれらの子）。
残すもの: 帳票の種類（patterns）・一覧表の取り込み設定（table_templates）・AI接続の設定。これらは「設定」で
          データではない。

DB の行は表の名前を決め打ちせず、その取り込みを指す列（document_id / import_id）を持つ表を
sqlite_master から探して消す（後から表が増えても消し残さないため）。
"""
from __future__ import annotations

import logging
import shutil
import sqlite3
from pathlib import Path

from flask import current_app

from core.files import remove_upload
from models import database

log = logging.getLogger(__name__)

# 取り込み1件を指す外部キーの列名（この列を持つ表は、その取り込みと一緒に消す）
DOCUMENT_REF_COLUMNS = ("document_id",)
IMPORT_REF_COLUMNS = ("import_id", "table_import_id")


def _tables_with_column(db, column: str) -> list[str]:
    names = []
    for (table,) in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
        if table.startswith("sqlite_"):
            continue
        if column in {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}:
            names.append(table)
    return names


def _delete_by_columns(db, columns, value) -> int:
    removed = 0
    for column in columns:
        for table in _tables_with_column(db, column):
            removed += db.execute(f'DELETE FROM "{table}" WHERE "{column}" = ?', (value,)).rowcount
    return removed


def _delete_jobs(db, ref_type: str, ref_id: int) -> int:
    return db.execute("DELETE FROM jobs WHERE ref_type = ? AND ref_id = ?", (ref_type, ref_id)).rowcount


def _delete_orphan_llm_calls(db) -> int:
    """どの ai_items からも参照されていない AI の生の応答を消す（試し実行の分もここで消える）。"""
    return db.execute(
        "DELETE FROM llm_calls WHERE cache_key NOT IN "
        "(SELECT cache_key FROM ai_items WHERE cache_key IS NOT NULL)").rowcount


def purge_after_send(response, fn, *args):
    """本文を最後まで送り終えてから消す（design.md 3.3）。

    送る前に消すと、ブラウザを閉じた・通信が切れた・保存先の空きが足りない、のどれでも
    「手元にファイルが無いのにサーバー側は消えている」ことになり、取り戻せない（再ダウンロードもできない）。
    本文を最後まで渡しきったときだけ消し、途中で切れたときは残す（もう一度ダウンロードできる）。

    send_file の応答は direct_passthrough が立っていて、そのままだと WSGI が close のときの
    呼び出し（call_on_close）を拾わないので、ここで解除して本文を包み直す。
    """
    app = current_app._get_current_object()
    body = response.response
    sent: list[bool] = []

    def stream():
        try:
            for chunk in body:
                yield chunk
            sent.append(True)   # 最後まで渡した
        finally:
            close = getattr(body, "close", None)
            if close is not None:
                close()

    response.direct_passthrough = False
    response.response = stream()

    @response.call_on_close
    def _purge() -> None:
        if not sent:
            return   # 送信が途中で終わった: 消さずに残す
        try:
            with app.app_context():
                fn(*args)
        except Exception:  # 応答は送り終えているので、ここで落としても利用者に伝える先が無い
            log.exception("ダウンロードしたデータの削除に失敗しました")

    return response


# ---- 帳票 --------------------------------------------------------------------------

def purge_documents(doc_ids) -> int:
    """帳票を消す（元のファイルと DB の行）。戻り値: 消した帳票の件数。"""
    ids = [int(i) for i in doc_ids]
    if not ids:
        return 0
    db = database.get_db()
    marks = ", ".join("?" for _ in ids)
    for row in db.execute(f"SELECT stored_path FROM documents WHERE id IN ({marks})", ids).fetchall():
        remove_upload(row[0])
    removed = 0
    for doc_id in ids:
        _delete_by_columns(db, DOCUMENT_REF_COLUMNS, doc_id)
        _delete_jobs(db, "document", doc_id)
        removed += db.execute("DELETE FROM documents WHERE id = ?", (doc_id,)).rowcount
    db.commit()
    _shrink(db)
    return removed


def purge_batch(batch_id: str) -> int:
    """まとめ取り込み1回分（同じ batch_id の帳票）をすべて消す。"""
    if not batch_id:
        return 0
    ids = [row[0] for row in database.get_db().execute("SELECT id FROM documents WHERE batch_id = ?", (batch_id,))]
    return purge_documents(ids)


# ---- 一覧表 ------------------------------------------------------------------------

def import_dir(import_id: int) -> Path:
    return Path(current_app.config["TABLES_DIR"]) / "imports" / str(int(import_id))


def purge_table_import(import_id: int) -> int:
    """一覧表の取り込み1件を消す（元のファイル・imports/<id>/・DB の行・AI整形の控え）。取り込み設定は残す。"""
    import_id = int(import_id)
    db = database.get_db()
    row = db.execute("SELECT stored_path, template_id FROM table_imports WHERE id = ?", (import_id,)).fetchone()
    if row is None:
        # DB の行が無くてもフォルダが残っていることがあるので、そこだけ片付ける
        shutil.rmtree(import_dir(import_id), ignore_errors=True)
        return 0
    remove_upload(row["stored_path"])
    shutil.rmtree(import_dir(import_id), ignore_errors=True)
    _delete_by_columns(db, IMPORT_REF_COLUMNS, import_id)
    _delete_jobs(db, "table_import", import_id)
    removed = db.execute("DELETE FROM table_imports WHERE id = ?", (import_id,)).rowcount
    # AI整形の結果の控え（行ごとの整形結果）は取り込んだデータそのものなので一緒に消す。この取り込みの分だけ
    # （ai_items.import_id）を消すので、同じ設定で作業中の別の取り込みの結果と応答キャッシュは残る。
    # import_id が無い行は古いDBの分（持ち主が分からない）なので、その設定の分をまとめて消す。
    if row["template_id"] is not None:
        db.execute("DELETE FROM ai_items WHERE template_id = ? AND import_id IS NULL", (row["template_id"],))
    _delete_orphan_llm_calls(db)
    db.commit()
    _shrink(db)
    return removed


def _shrink(db) -> None:
    """消したあとの後始末（design.md 3.3）。

    - wal_checkpoint(TRUNCATE): 消す前後の内容が残る app.db-wal を切り詰める（強制終了時に拾われないように）
    - VACUUM: 解放したページを手放してファイルの大きさも戻す（中身は PRAGMA secure_delete で消えている）
    どちらも他の接続が使っていると実行できない。そのときは次に消したときに縮む。
    """
    for statement in ("PRAGMA wal_checkpoint(TRUNCATE)", "VACUUM"):
        try:
            db.execute(statement)
        except sqlite3.Error:
            pass
