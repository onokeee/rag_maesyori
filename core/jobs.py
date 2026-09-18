"""バックグラウンドジョブ（単一ワーカースレッド＋キュー）。

- 状態・進捗・一時停止/中止の要求は jobs テーブルに持つ（画面はポーリングで読む）。
- fn(ctx) は start_job を呼んだアプリの app_context 内で実行する。fn 内で DB を使うときは database.connect()。
- 実行中は一定間隔で heartbeat_at を更新する。起動時に recover_interrupted() で止まったジョブを「中断」にする。
"""
from __future__ import annotations

import json
import queue
import threading
import time
from datetime import datetime, timedelta
from typing import Callable

from flask import current_app

from models import database

HEARTBEAT_INTERVAL = 30          # 秒。fn が呼ばなくても実行中はこの間隔で更新する
STALE_AFTER = timedelta(minutes=2)
PAUSE_POLL = 0.3                 # 一時停止中に再開・中止を確かめる間隔（秒）

ACTIVE_STATUSES = ("queued", "running", "paused")
TERMINAL_STATUSES = ("done", "failed", "cancelled", "interrupted")

STATUS_LABELS = {
    "queued": "待機中",
    "running": "処理中",
    "paused": "一時停止中",
    "done": "完了",
    "failed": "エラー",
    "cancelled": "中止",
    "interrupted": "中断",
}

_queue: queue.Queue = queue.Queue()
_worker: threading.Thread | None = None
_worker_lock = threading.Lock()


class JobCancelled(Exception):
    """中止の要求を受けて fn の処理を打ち切るときに投げる。"""


class JobError(Exception):
    """利用者にそのまま見せてよい日本語メッセージを持つエラー。

    これを継承した例外（aiproc.runner.AIJobError、tables.pipeline.PipelineError）は、そのメッセージを
    画面にそのまま出す。それ以外の例外は Python の例外名が画面に出ないようにし、内容はログにだけ残す。
    """


def error_message(exc: Exception) -> str:
    """ジョブの失敗を画面に出す日本語の文にする。詳しい内容（例外名）はログに残してあるので出さない。"""
    text = " ".join(str(exc).split())
    if isinstance(exc, JobError) and text:
        return text
    return "処理中にエラーが発生しました。もう一度実行してください（詳しい内容はアプリのログに記録しました）"


def _now() -> str:
    return database.now()


class JobContext:
    """fn に渡す操作口。スレッドをまたいで呼んでもよい（内部でロック）。"""

    def __init__(self, job_id: int, conn=None):
        self.job_id = job_id
        self._conn = conn or database.connect()
        self._lock = threading.Lock()
        row = self._row()
        self._progress: dict = json.loads(row["progress_json"] or "{}") if row else {}
        self.params: dict = json.loads(row["params_json"] or "{}") if row else {}

    # ---- 内部 ----
    def _row(self):
        with self._lock:
            return self._conn.execute("SELECT * FROM jobs WHERE id = ?", (self.job_id,)).fetchone()

    def _update(self, sql_sets: str, args=()) -> None:
        with self._lock:
            self._conn.execute(f"UPDATE jobs SET {sql_sets} WHERE id = ?", (*args, self.job_id))
            self._conn.commit()

    def _flags(self) -> tuple[bool, bool]:
        row = self._row()
        if row is None:
            return True, False  # ジョブが消された → 中止扱い
        return bool(row["cancel_requested"]), bool(row["pause_requested"])

    def close(self) -> None:
        with self._lock:
            self._conn.close()

    # ---- 公開 ----
    def progress(self, **kw) -> None:
        """進捗を上書きマージして保存する（例: done=10, total=100, phase="読み込み"）。heartbeat も更新。"""
        self._progress.update(kw)
        ts = _now()
        self._update("progress_json = ?, heartbeat_at = ?, updated_at = ?",
                     (json.dumps(self._progress, ensure_ascii=False, default=str), ts, ts))

    def message(self, text: str) -> None:
        self._update("message = ?, updated_at = ?", (text, _now()))

    def heartbeat(self) -> None:
        self._update("heartbeat_at = ?", (_now(),))

    def is_cancelled(self) -> bool:
        return self._flags()[0]

    def should_stop(self) -> bool:
        """中止または一時停止が要求されていれば True（新しい処理を出さないための判定）。"""
        cancel, pause = self._flags()
        return cancel or pause

    def wait_if_paused(self) -> bool:
        """一時停止の要求があれば再開まで待つ。続けてよければ True、中止なら False。"""
        cancel, pause = self._flags()
        if cancel:
            return False
        if not pause:
            return True
        self._update("status = 'paused', heartbeat_at = ?, updated_at = ?", (_now(), _now()))
        last_beat = time.monotonic()
        while True:
            time.sleep(PAUSE_POLL)
            cancel, pause = self._flags()
            if cancel:
                return False
            if not pause:
                self._update("status = 'running', heartbeat_at = ?, updated_at = ?", (_now(), _now()))
                return True
            if time.monotonic() - last_beat >= HEARTBEAT_INTERVAL:
                self.heartbeat()
                last_beat = time.monotonic()

    def check_cancel(self) -> None:
        """一時停止中なら待ち、中止が要求されていれば JobCancelled を投げる。"""
        if not self.wait_if_paused():
            raise JobCancelled()


# ---- 実行 -----------------------------------------------------------------------

def start_job(kind: str, ref_type: str, ref_id: int | None, fn: Callable[[JobContext], dict | None],
              params: dict | None = None) -> int:
    """ジョブを登録してキューに入れ、job_id を返す。app_context 内で呼ぶ。"""
    app = current_app._get_current_object()
    ts = _now()
    conn = database.connect()
    try:
        cur = conn.execute(
            """INSERT INTO jobs (kind, ref_type, ref_id, status, params_json, progress_json, created_at, updated_at)
               VALUES (?, ?, ?, 'queued', ?, '{}', ?, ?)""",
            (kind, ref_type or "", ref_id, json.dumps(params or {}, ensure_ascii=False, default=str), ts, ts),
        )
        conn.commit()
        job_id = cur.lastrowid
    finally:
        conn.close()
    _queue.put((app, job_id, fn))
    _ensure_worker()
    return job_id


def _ensure_worker() -> None:
    global _worker
    with _worker_lock:
        if _worker is None or not _worker.is_alive():
            _worker = threading.Thread(target=_worker_loop, name="job-worker", daemon=True)
            _worker.start()


def _worker_loop() -> None:
    while True:
        app, job_id, fn = _queue.get()
        try:
            _run(app, job_id, fn)
        except Exception:  # ワーカーは止めない
            try:
                app.logger.exception("ジョブ %s の実行管理でエラー", job_id)
            except Exception:
                pass
        finally:
            _queue.task_done()


class _Ticker:
    """fn が長い呼び出しで止まっていても heartbeat を更新し続ける。"""

    def __init__(self, ctx: JobContext):
        self._ctx = ctx
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, name=f"job-{ctx.job_id}-heartbeat", daemon=True)

    def _loop(self) -> None:
        while not self._stop.wait(HEARTBEAT_INTERVAL):
            try:
                self._ctx.heartbeat()
            except Exception:
                return

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(timeout=5)


def _run(app, job_id: int, fn: Callable[[JobContext], dict | None]) -> None:
    with app.app_context():
        ctx = JobContext(job_id)
        try:
            row = ctx._row()
            if row is None or row["status"] != "queued":
                return  # 待機中に中止・中断された
            if row["cancel_requested"]:
                ctx._update("status = 'cancelled', message = ?, updated_at = ?", ("中止しました", _now()))
                return
            ts = _now()
            with ctx._lock:  # 確認と開始の間に中止されていたら実行しない
                cur = ctx._conn.execute("UPDATE jobs SET status = 'running', heartbeat_at = ?, updated_at = ? "
                                        "WHERE id = ? AND status = 'queued'", (ts, ts, job_id))
                ctx._conn.commit()
            if cur.rowcount == 0:
                return
            try:
                with _Ticker(ctx):
                    result = fn(ctx)
            except JobCancelled:
                ctx._update("status = 'cancelled', pause_requested = 0, message = ?, updated_at = ?",
                            ("中止しました", _now()))
                return
            except Exception as exc:
                app.logger.exception("ジョブ %s（%s）でエラー", job_id, row["kind"])
                ctx._update("status = 'failed', pause_requested = 0, message = ?, updated_at = ?",
                            (error_message(exc), _now()))
                return
            if result is not None:
                ctx.progress(result=result)
            if ctx.is_cancelled():
                ctx._update("status = 'cancelled', pause_requested = 0, message = CASE WHEN message = '' "
                            "THEN '中止しました' ELSE message END, updated_at = ?", (_now(),))
            else:
                ctx._update("status = 'done', pause_requested = 0, updated_at = ?", (_now(),))
        finally:
            ctx.close()


# ---- 参照・操作（画面から） --------------------------------------------------------

def _decode(row) -> dict | None:
    if row is None:
        return None
    job = dict(row)
    job["params"] = json.loads(job.get("params_json") or "{}")
    job["progress"] = json.loads(job.get("progress_json") or "{}")
    job["result"] = job["progress"].get("result")
    job["status_label"] = STATUS_LABELS.get(job["status"], job["status"])
    job["finished"] = job["status"] in TERMINAL_STATUSES
    return job


def _with_conn(fn):
    conn = database.connect()
    try:
        return fn(conn)
    finally:
        conn.close()


def get_job(job_id: int) -> dict | None:
    """ジョブ1件（params / progress / result / status_label / finished を付けて返す）。"""
    return _with_conn(lambda c: _decode(c.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()))


def latest_job(ref_type: str, ref_id: int, kind: str | None = None) -> dict | None:
    sql, args = "SELECT * FROM jobs WHERE ref_type = ? AND ref_id = ?", [ref_type, ref_id]
    if kind:
        sql += " AND kind = ?"
        args.append(kind)
    return _with_conn(lambda c: _decode(c.execute(sql + " ORDER BY id DESC LIMIT 1", args).fetchone()))


def _request(job_id: int, sets: str, statuses=ACTIVE_STATUSES) -> bool:
    def run(conn):
        marks = ", ".join("?" for _ in statuses)
        cur = conn.execute(f"UPDATE jobs SET {sets}, updated_at = ? WHERE id = ? AND status IN ({marks})",
                           (_now(), job_id, *statuses))
        conn.commit()
        return cur.rowcount > 0
    return _with_conn(run)


def request_pause(job_id: int) -> bool:
    return _request(job_id, "pause_requested = 1", ("queued", "running"))


def request_resume(job_id: int) -> bool:
    return _request(job_id, "pause_requested = 0")


def request_cancel(job_id: int) -> bool:
    """中止を要求する。待機中のジョブはその場で中止にする。"""
    ok = _request(job_id, "cancel_requested = 1")
    _request(job_id, "status = 'cancelled', message = '中止しました'", ("queued",))
    return ok


def recover_interrupted() -> int:
    """起動時：実行中・待機中・一時停止中のまま heartbeat が2分以上古いジョブを「中断」にする。件数を返す。"""
    cutoff = (datetime.now() - STALE_AFTER).isoformat(timespec="seconds")

    def run(conn):
        marks = ", ".join("?" for _ in ACTIVE_STATUSES)
        cur = conn.execute(
            f"""UPDATE jobs SET status = 'interrupted', pause_requested = 0,
                    message = 'アプリの終了などで処理が中断されました。もう一度実行してください', updated_at = ?
                WHERE status IN ({marks}) AND COALESCE(heartbeat_at, updated_at, created_at) < ?""",
            (_now(), *ACTIVE_STATUSES, cutoff),
        )
        conn.commit()
        return cur.rowcount
    return _with_conn(run)


def wait_job(job_id: int, timeout: float = 30.0, statuses=TERMINAL_STATUSES) -> dict | None:
    """指定の状態になるまで待つ（テスト・同期実行用）。時間切れなら最後の状態を返す。"""
    deadline = time.monotonic() + timeout
    while True:
        job = get_job(job_id)
        if job is None or job["status"] in statuses or time.monotonic() >= deadline:
            return job
        time.sleep(0.05)
