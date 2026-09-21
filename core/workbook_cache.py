"""置かれた Excel を「読み取った形」だけ、しばらくメモリに覚えておく置き場。

帳票登録は見本の Excel をサーバーに置かない（利用者の指示 2026-09-21「見本のExcelは置かずに、
設定だけ保持するようにしてほしい」）。ブラウザはファイルを選んだまま持っているので、セルを
クリックするたびに同じ Excel を送り直してくる。そのたびにブックを開き直すと（結合セル・図形を
1つずつ作るので）数秒かかることがあるため、開いた結果（WorkbookInfo）だけをここに置く。

決まりごと:
  - ディスクには何も書かない。Excel の中身（bytes）も持たない（開いた結果だけ）。
  - アプリを終えれば消える（残るのは読み取りの設定だけ、という約束を崩さない）。
  - 鍵は「ブラウザの作業場所（session_id）＋ファイルの sha256」。ほかの人が置いたブックは見えない。
  - 古いもの・多すぎるもの・大きすぎるものは、置いた順に落とす。
  - waitress は1つのプロセスを複数のスレッドで回すので、出し入れは錠（Lock）の中で行う。
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field

# 覚えておく時間。帳票登録は「置く → セルをクリックする → 使用開始」を続けて行う作業なので、
# 手が止まっても1回分の作業（30分）で足りる。切れてもブラウザが送り直すので読み直すだけ。
TTL_SECONDS = 30 * 60
# 覚えておく数と、その元になった Excel の大きさの合計。数人が同時に使っても RAM を食い尽くさない値。
MAX_ENTRIES = 8
MAX_BYTES = 64 * 1024 * 1024


@dataclass
class Book:
    """ブラウザが置いた Excel を読み取った結果（中身そのものは持たない）。"""

    file_name: str
    file_hash: str
    size: int                      # 元の Excel の大きさ（置ける量を数えるため）
    info: object                   # excel.workbook.WorkbookInfo
    used_at: float = field(default_factory=time.monotonic)


_books: dict[tuple[str, str], Book] = {}
_lock = threading.Lock()


def _expired(book: Book, now: float) -> bool:
    return now - book.used_at > TTL_SECONDS


def _evict(now: float) -> None:
    """古いもの → 入れた順、の順に落とす（錠の中から呼ぶ）。"""
    for key in [k for k, b in _books.items() if _expired(b, now)]:
        del _books[key]
    total = sum(b.size for b in _books.values())
    while _books and (len(_books) > MAX_ENTRIES or total > MAX_BYTES):
        key, dropped = next(iter(_books.items()))   # dict は入れた順に並ぶ＝いちばん古いもの
        del _books[key]
        total -= dropped.size


def put(session_id: str, book: Book) -> Book:
    """読み取った結果を覚える（同じブックを置き直したら入れ替える）。"""
    key = (session_id or "", book.file_hash)
    with _lock:
        book.used_at = time.monotonic()
        _books.pop(key, None)       # 入れ直して「いちばん新しい」にする
        _books[key] = book
        _evict(book.used_at)
    return book


def get(session_id: str, file_hash: str) -> Book | None:
    """覚えているブックを返す（無ければ None。呼び出し側はブラウザに置き直してもらう）。"""
    key = (session_id or "", file_hash or "")
    now = time.monotonic()
    with _lock:
        book = _books.get(key)
        if book is None:
            return None
        if _expired(book, now):
            del _books[key]
            return None
        book.used_at = now
        _books.pop(key)
        _books[key] = book          # 使ったものを「いちばん新しい」にする
        return book


def clear() -> None:
    """全部忘れる（テスト・片付け用）。"""
    with _lock:
        _books.clear()


def count() -> int:
    with _lock:
        return len(_books)
