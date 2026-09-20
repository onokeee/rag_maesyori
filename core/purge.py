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
import secrets
import shutil
import sqlite3
from pathlib import Path

from flask import current_app, g, has_app_context

from core.files import remove_upload
from models import database

log = logging.getLogger(__name__)

# 取り込み1件を指す外部キーの列名（この列を持つ表は、その取り込みと一緒に消す）
DOCUMENT_REF_COLUMNS = ("document_id",)
IMPORT_REF_COLUMNS = ("import_id", "table_import_id")
# 取り込みを指す列があっても、まとめては消さない表。llm_calls は AI の生の応答で、同じ文面の行なら
# 別の取り込みの ai_items が同じ応答を使っていることがある（消すと再実行で再課金になる）。
# 持ち主が消えた分は _delete_orphan_llm_calls が「参照されていなければ消す」で扱う。
SHARED_TABLES = ("llm_calls",)


def _tables_with_column(db, column: str) -> list[str]:
    names = []
    for (table,) in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
        if table.startswith("sqlite_") or table in SHARED_TABLES:
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


_UNREFERENCED_LLM_CALLS = ("cache_key NOT IN (SELECT cache_key FROM ai_items WHERE cache_key IS NOT NULL)")
# 持ち主の取り込みがもう無い応答（持ち主の分からない古いDBの分＝NULL も含む）
NO_LIVE_OWNER = "(import_id IS NULL OR import_id NOT IN (SELECT id FROM table_imports))"


def _ai_work_alive(db) -> bool:
    """まだある取り込みに、AI整形の結果かAI整形のジョブ（一時停止・中断を含む）が残っているか。

    ai_items が覚えているのは最後に使った応答のキーだけで、再依頼で直した行の1回目の応答や、
    一時停止の直前に受け取った応答（ai_items の行がまだ無い）は、どこからも参照されない。
    それでも再実行・再開ではキャッシュとして引く（引けないと再課金になる）ので、
    AI整形が残っている取り込みがある間は、持ち主の分からない（import_id が NULL の古いDBの）
    参照されない応答を消さない。持ち主の分かる応答は、その取り込みを消すときに一緒に消える。
    """
    return db.execute(
        "SELECT 1 FROM ai_items WHERE import_id IN (SELECT id FROM table_imports) "
        "UNION ALL SELECT 1 FROM jobs WHERE kind = 'ai_format' AND ref_type = 'table_import' "
        "AND ref_id IN (SELECT id FROM table_imports) LIMIT 1").fetchone() is not None


def _delete_orphan_llm_calls(db, keys=()) -> int:
    """どの ai_items からも参照されていない AI の生の応答を消す（試し実行の分もここで消える）。

    keys: 消した取り込みが使っていた応答のキー。ほかの行が使っておらず、まだある取り込みが払った分でも
    なければ、いつでもすぐ消す。消す取り込みは、別の取り込みが払った応答をキャッシュとして引いている
    ことがある（同じ文面なら同じキーになる）ので、払った取り込みがまだあるうちは消さない
    （消すと、その取り込みの再開・再実行で同じ応答をもう一度買うことになる）。その応答は、払った
    取り込みを消すときに下の「持ち主の取り込みがもう無い分」で消えるので、残り続けることはない。
    もう無い取り込みが払った応答（llm_calls.import_id）も、ほかから使われていなければすぐ消す。
    ai_items は最後に使った応答のキーしか覚えていないので、再依頼で直した行の1回目の応答や、
    一時停止の直前に受け取った応答は keys に入らない。持ち主の列があれば、それも一緒に消せる。
    持ち主の分からない応答（古いDBの分）は、AI整形が残っている取り込みが1件も無いときだけ消す（_ai_work_alive）。
    """
    removed = 0
    keys = sorted({k for k in keys if k})
    for start in range(0, len(keys), 500):
        chunk = keys[start:start + 500]
        marks = ", ".join("?" for _ in chunk)
        removed += db.execute(f"DELETE FROM llm_calls WHERE cache_key IN ({marks}) "
                              f"AND {NO_LIVE_OWNER} AND {_UNREFERENCED_LLM_CALLS}", chunk).rowcount
    # 持ち主の取り込みがもう無い分（design.md 3.3「取り込んだデータはダウンロードが終わった時点で消す」）
    removed += db.execute(f"DELETE FROM llm_calls WHERE import_id IS NOT NULL "
                          f"AND import_id NOT IN (SELECT id FROM table_imports) "
                          f"AND {_UNREFERENCED_LLM_CALLS}").rowcount
    # まだ使われていて残した分は、持ち主が居なくなったので「持ち主なし」に戻す
    # （使っている取り込みを消すときに、その keys で消える）
    db.execute("UPDATE llm_calls SET import_id = NULL WHERE import_id IS NOT NULL "
               "AND import_id NOT IN (SELECT id FROM table_imports)")
    if not _ai_work_alive(db):
        removed += db.execute(f"DELETE FROM llm_calls WHERE {_UNREFERENCED_LLM_CALLS}").rowcount
    return removed


def delete_orphan_ai_items(db) -> int:
    """もう無い取り込みを指す AI整形の結果を消す（取り込みを消したあとに、動いていた AI の書き込みが残った分など）。

    import_id の無い行（古いDBの分。持ち主の取り込みが分からない）は、その取り込み設定の取り込みが
    1件も残っていなければ消す（残っていれば、その取り込みを消すときに purge_table_import が消す）。
    """
    removed = db.execute("DELETE FROM ai_items WHERE import_id IS NOT NULL "
                         "AND import_id NOT IN (SELECT id FROM table_imports)").rowcount
    removed += db.execute("DELETE FROM ai_items WHERE import_id IS NULL AND template_id NOT IN "
                          "(SELECT template_id FROM table_imports WHERE template_id IS NOT NULL)").rowcount
    return removed


def sweep_orphan_ai(db) -> int:
    """持ち主の無い AI整形の結果と、どこからも使われない AI の生の応答を消して縮める（戻り値: 消した件数）。

    取り込み設定を消したとき（ai_items は消えるが llm_calls が残る）や起動時の片付けで使う。
    """
    removed = delete_orphan_ai_items(db) + _delete_orphan_llm_calls(db)
    db.commit()
    if removed:
        _shrink(db)
    return removed


def _mark_incomplete() -> None:
    """この要求の中で、消し切れなかったもの（掴まれていて消せなかったファイル）があったと印を付ける。"""
    if has_app_context():
        g.purge_incomplete = True


def purge_incomplete() -> bool:
    """この要求の中で消したときに、消し切れなかったファイルがあったか（保存先フォルダの結果の画面で知らせるため）。

    消せなかったファイルも中身は 0 バイトに切り詰めてあり、次の起動時に片付く。例外にはしない
    （消すこと自体は DB の行まで終わっていて、画面を 500 にするより「消し切れなかった」と知らせる方がよい）。
    """
    return has_app_context() and bool(g.get("purge_incomplete"))


def _truncate_leftovers(folder: Path) -> None:
    """消し切れなかったフォルダに残ったファイルを 0 バイトに切り詰める（切り詰められないものは諦める）。"""
    for path in folder.rglob("*"):
        try:
            if path.is_file():
                with open(path, "r+b") as f:
                    f.truncate(0)
        except OSError:
            pass


def purge_after_send(response, fn, *args):
    """本文を最後まで送り終えてから消す（design.md 3.3）。

    送る前に消すと、ブラウザを閉じた・通信が切れた・保存先の空きが足りない、のどれでも
    「手元にファイルが無いのにサーバー側は消えている」ことになり、取り戻せない（再ダウンロードもできない）。
    本文を最後まで渡しきったときだけ消し、途中で切れたときは残す（もう一度ダウンロードできる）。

    「渡しきった」はサーバー（waitress）の送信の溜めに入れ終えたことで、相手が受け取ったことではない。
    溜めの上限は app.OUTBUF_HIGH_WATERMARK（16KB）で、OS の溜めと合わせて最後の数十KBは受け取られる前に
    消すことになる（数十KBより小さい md は、途中で切れても消える。design.md 3.3）。

    send_file の応答は direct_passthrough が立っていて、そのままだと WSGI が close のときの
    呼び出し（call_on_close）を拾わないので、ここで解除して本文を包み直す。
    """
    if response.status_code != 200:
        # 206（Range で一部だけ）・304 などは本文の全部を渡していない。ダウンロードマネージャーや
        # 途中からの再開、ウイルス対策のプロキシが送ってくることがある。消すと残りを取り戻せないので消さない。
        return response
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

# 帳票を消したときに呼ぶ関数（画面側がメモリに持っている帳票ごとの目印を一緒に捨てるため）。
# 引数は消した帳票の id のリスト。
_DOCUMENT_PURGE_HOOKS: list = []


def on_documents_purged(fn):
    """帳票を消したあとに fn(doc_ids) を呼ぶよう登録する（同じ関数は1回だけ）。デコレーターとしても使える。"""
    if fn not in _DOCUMENT_PURGE_HOOKS:
        _DOCUMENT_PURGE_HOOKS.append(fn)
    return fn


def purge_documents(doc_ids) -> int:
    """帳票を消す（元のファイルと DB の行）。戻り値: 消した帳票の件数。"""
    ids = [int(i) for i in doc_ids]
    if not ids:
        return 0
    db = database.get_db()
    marks = ", ".join("?" for _ in ids)
    stored = [row[0] for row in db.execute(f"SELECT stored_path FROM documents WHERE id IN ({marks})", ids)]
    # DB の行を先に消す。ファイルを先に消すと、DB の削除が失敗したとき（ロック・強制終了）に
    # 「ファイルの無い帳票」が作業中として残ってしまう。逆なら残るのは行の無いファイルで、起動時に片付く。
    removed = 0
    for doc_id in ids:
        _delete_by_columns(db, DOCUMENT_REF_COLUMNS, doc_id)
        _delete_jobs(db, "document", doc_id)
        removed += db.execute("DELETE FROM documents WHERE id = ?", (doc_id,)).rowcount
    db.commit()
    for path in stored:
        if not remove_upload(path):
            _mark_incomplete()
    for hook in _DOCUMENT_PURGE_HOOKS:
        try:
            hook(ids)
        except Exception:  # 目印の片付けに失敗しても、消すこと自体は終わっている
            log.exception("帳票を消したあとの片付けに失敗しました")
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
    # DB の行を先に消し、ファイルはそのあと（purge_documents と同じ理由。残ったフォルダ・ファイルは起動時に片付く）
    # この取り込みが使っていた AI の応答のキー（行を消す前に控える。消したあと、ほかで使われていなければ消す）
    keys = [r[0] for r in db.execute("SELECT cache_key FROM ai_items WHERE import_id = ? AND cache_key IS NOT NULL",
                                     (import_id,))]
    if row["template_id"] is not None:
        keys += [r[0] for r in db.execute("SELECT cache_key FROM ai_items WHERE template_id = ? AND import_id IS NULL "
                                          "AND cache_key IS NOT NULL", (row["template_id"],))]
    _delete_by_columns(db, IMPORT_REF_COLUMNS, import_id)
    _delete_jobs(db, "table_import", import_id)
    removed = db.execute("DELETE FROM table_imports WHERE id = ?", (import_id,)).rowcount
    # AI整形の結果の控え（行ごとの整形結果）は取り込んだデータそのものなので一緒に消す。この取り込みの分だけ
    # （ai_items.import_id）を消すので、同じ設定で作業中の別の取り込みの結果と応答キャッシュは残る。
    # import_id が無い行は古いDBの分（持ち主が分からない）なので、その設定の分をまとめて消す。
    if row["template_id"] is not None:
        db.execute("DELETE FROM ai_items WHERE template_id = ? AND import_id IS NULL", (row["template_id"],))
    delete_orphan_ai_items(db)
    _delete_orphan_llm_calls(db, keys)
    db.commit()
    if not remove_upload(row["stored_path"]):
        _mark_incomplete()
    folder = import_dir(import_id)
    shutil.rmtree(folder, ignore_errors=True)
    if folder.exists():
        # Windows で別のスレッド・プロセスがファイルを掴んでいると消し残る。黙って成功扱いにせず記録し、
        # remove_upload と同じく残ったファイルの中身を 0 バイトに切り詰める（名前は残っても読み込んだ行・md は残さない）。
        # DB の行はもう無いので、残ったフォルダは次の起動時に remove_orphan_import_dirs が片付ける。
        _truncate_leftovers(folder)
        _mark_incomplete()
        log.warning("取り込みのフォルダを消し切れませんでした（中身は空にし、次の起動時に片付けます）: imports/%s",
                    import_id)
    _shrink(db)
    return removed


def _shrink(db) -> None:
    """消したあとの後始末（design.md 3.3）。

    - wal_checkpoint(TRUNCATE): 消す前後の内容が残る app.db-wal を切り詰める（強制終了時に拾われないように）
    - VACUUM: 解放したページを手放してファイルの大きさも戻す（中身は PRAGMA secure_delete で消えている）
    どちらも他の接続が使っていると実行できない。そのときは次に消したときに縮む。

    VACUUM は DB 全体を書き直すあいだ書き込みを止める。大きな app.db だと busy_timeout より長くなり、
    動いている AI整形などのジョブの書き込みが「database is locked」で失敗する。なので、待機中・実行中・
    一時停止中のジョブがあるときは VACUUM をしない（消した中身は secure_delete で上書き済み。
    ファイルの大きさは、ジョブが無いときの次の削除か起動時の片付けで戻る）。
    """
    forget_id_counters(db)
    statements = ["PRAGMA wal_checkpoint(TRUNCATE)"]
    if not _jobs_active(db):
        statements.append("VACUUM")
    for statement in statements:
        try:
            db.execute(statement)
        except sqlite3.Error:
            pass


def _jobs_active(db) -> bool:
    try:
        row = db.execute("SELECT 1 FROM jobs WHERE status IN ('queued', 'running', 'paused') LIMIT 1").fetchone()
        return row is not None
    except sqlite3.Error:
        return True   # 確かめられないときは、動いているかもしれないジョブを止めない側に倒す


# ---- 番号の続きを忘れる ------------------------------------------------------------
# 取り込むたびに行を作り、ダウンロードで消す表。AUTOINCREMENT の表は、消したあとも sqlite_sequence に
# 「これまでに使った一番大きい番号」が残り、何件取り込んだかの記録になってしまう（design.md 3.3「履歴は持たない」）。
WORK_TABLES = ("documents", "table_imports", "jobs", "ai_items")
# 表が空になったら、続きの番号をこの範囲の乱数にする。最初の番号（1〜）とも前回の番号とも重ならないので、
# 開いたままの古い画面やブラウザに残った画面が、消した番号で別の取り込みを開いたり書き換えたりしない
# （重なる確率は 1回あたり 取り込み件数 / 約4.5×10^15。JavaScript で正確に扱える 2^53 未満に収める）。
_ID_BASE_MIN = 2 ** 31
_ID_BASE_MAX = 2 ** 52


def forget_id_counters(db) -> int:
    """空になった取り込みの表の、番号の続き（sqlite_sequence）を乱数に置き換える。戻り値: 置き換えた表の数。

    AUTOINCREMENT は外さない（外すと、消した番号がすぐ使い回され、古い画面が別の取り込みを指す）。
    行が残っている表は置き換えない（作業中の取り込みの番号が見えている間は、続きの番号を隠しても意味が無い）。
    """
    changed = 0
    try:
        for table in WORK_TABLES:
            base = _ID_BASE_MIN + secrets.randbelow(_ID_BASE_MAX - _ID_BASE_MIN)
            # 空かどうかの確認と置き換えを1文で行う（あいだに別のスレッドが行を足しても番号が重ならない）
            changed += db.execute(f'UPDATE sqlite_sequence SET seq = ? WHERE name = ? '
                                  f'AND NOT EXISTS (SELECT 1 FROM "{table}")', (base, table)).rowcount
        db.commit()
    except sqlite3.Error:
        # sqlite_sequence がまだ無い（AUTOINCREMENT の表に1行も入れていない新しいDB）か、ロック中。
        # 次に消したとき・次の起動時にもう一度行う
        db.rollback()
        return 0
    return changed
