"""取り込みごとの読み取り結果の控え（シート一覧・シートの大きさ・先頭の行・表の形の推定・列の見本の行）。

画面を開くたびに表ファイル全体を開き直して読み直さないように、取り込みのフォルダに小さな JSON で持つ。
ファイル（ハッシュ）と読み込み設定（文字コード・区切り・読めない文字の扱い・セル数の上限）を鍵にし、
どれかが変わったら中身を捨てて作り直す（読み取り・判定のコードを直したときも作り直す）。
中身は元のファイルから同じ計算で作れるものだけ（消してもよい）。

ImportSource は TableSource として振る舞う: 控えにある先頭の行は控えから返し、それ以外を読むときだけ元のファイルを開く。
"""
from __future__ import annotations

import functools
import hashlib
import json
import os
import threading
from dataclasses import asdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Callable, Iterator

from tables.detect import HEAD_ROWS, LayoutGuess, RowClass, guess_layout, sample_data_rows
from tables.source import CellInfo, SheetInfo, SourceRow

CACHE_FILE = "source_cache.json"
CACHE_VERSION = 1
MAX_LAYOUTS = 24  # 見出し行を何度も指定し直したときに増えすぎないように
MAX_SAMPLES = 4


class _NotCacheable(Exception):
    """JSON にできない値（控えに入れずに毎回読む）。"""


# ---- 値・行・表の形の JSON 化 ------------------------------------------------------------

def _enc_value(v):
    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    if isinstance(v, datetime):
        return {"dt": v.isoformat()}
    if isinstance(v, date):
        return {"d": v.isoformat()}
    if isinstance(v, time):
        return {"t": v.isoformat()}
    if isinstance(v, timedelta):
        return {"td": [v.days, v.seconds, v.microseconds]}
    raise _NotCacheable(type(v).__name__)


def _dec_value(v):
    if not isinstance(v, dict):
        return v
    if "dt" in v:
        return datetime.fromisoformat(v["dt"])
    if "d" in v:
        return date.fromisoformat(v["d"])
    if "t" in v:
        return time.fromisoformat(v["t"])
    days, seconds, micro = v["td"]
    return timedelta(days=days, seconds=seconds, microseconds=micro)


def _enc_row(row: SourceRow) -> list:
    cells = []
    for c in row.cells:
        if c.value is None and not c.text and c.number_format is None and not (c.bold or c.strike or c.fill) \
                and c.merged_anchor is None:
            cells.append(0)  # 何もない空欄
            continue
        flags = (1 if c.bold else 0) | (2 if c.strike else 0) | (4 if c.fill else 0)
        cells.append([_enc_value(c.value), c.text, c.number_format, flags,
                      list(c.merged_anchor) if c.merged_anchor else None])
    return [row.index, row.hidden, cells]


def _dec_row(data: list) -> SourceRow:
    cells = []
    for c in data[2]:
        if c == 0:
            cells.append(CellInfo(value=None, text=""))
            continue
        value, text, number_format, flags, anchor = c
        cells.append(CellInfo(value=_dec_value(value), text=text, number_format=number_format, bold=bool(flags & 1),
                              strike=bool(flags & 2), fill=bool(flags & 4),
                              merged_anchor=tuple(anchor) if anchor else None))
    return SourceRow(index=data[0], cells=cells, hidden=data[1])


def _dec_layout(data: dict) -> LayoutGuess:
    d = dict(data)
    d["row_classes"] = [RowClass(**rc) for rc in d["row_classes"]]
    return LayoutGuess(**d)


def layout_key(sheet: str, anchors=None, header_row=None, data_end=None, header_rows=None, max_scan_rows=None) -> str:
    """guess_layout の結果を左右する引数だけで鍵を作る（見出し行を指定したときはアンカー語・見出し候補は使われない）。"""
    header_rows = sorted(int(r) for r in header_rows) if header_rows else None
    header_row = int(header_row) if header_row and not header_rows else None
    anchors = list(anchors) if anchors and not header_rows and not header_row else None
    return json.dumps([sheet, anchors, header_row, header_rows, int(data_end) if data_end else None,
                       int(max_scan_rows) if max_scan_rows else None], ensure_ascii=False)


# ---- 控えのファイル ----------------------------------------------------------------------------

_CODE_MODULES = ("source.py", "excel_source.py", "csv_source.py", "detect.py", "dictionary.py", "source_cache.py")


@functools.lru_cache(maxsize=1)
def _code_version() -> str:
    """読み取り・判定のコードの中身のハッシュ。コードを直したら古い控えを使わない。"""
    digest = hashlib.sha256()
    base = Path(__file__).resolve().parent
    for name in _CODE_MODULES:
        try:
            digest.update((base / name).read_bytes())
        except OSError:
            digest.update(name.encode("utf-8"))
    return digest.hexdigest()[:16]


class ImportSource:
    """元の表ソースの代わりに使う、控え付きの表ソース（kind・file_name・sheets・rows は元と同じ）。

    open_real(sheet_stats) で元のファイルを開く（Excel はシートの大きさの控えを渡すと数え直さない）。
    real に開いた元のソースを渡せば、それを使う（ジョブで読み込みと同じソースを使うとき）。
    """

    def __init__(self, directory: Path, key: str, kind: str, file_name: str,
                 open_real: Callable[[dict | None], object], real=None):
        self.directory = Path(directory)
        self.path = self.directory / CACHE_FILE
        try:
            # 置き場所はここで用意する（消したあとに書き戻さないよう _save では作らない）
            self.directory.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        self.key = key
        self.kind = kind
        self.file_name = file_name
        self._open_real = open_real
        self._real = real
        self._data = self._load()
        if real is not None:
            self._remember_stats(real)

    # ---- 控えの読み書き ----

    def _load(self) -> dict:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            data = None
        version = [CACHE_VERSION, _code_version()]
        if not isinstance(data, dict) or data.get("version") != version or data.get("key") != self.key:
            data = {"version": version, "key": self.key}
        return data

    def _save(self) -> None:
        """一時ファイルに書いて置き換える。書けなくても処理は続ける（次に作り直すだけ）。

        ここではフォルダを作らない: 取り込みを消した（purge_table_import が imports/<id>/ を消した）あとに
        別のタブの画面処理が控えを書くと、元の表の先頭行・見本の行が入った控えごとフォルダが
        復活してしまう（design.md 3.3「データを残さない」）。フォルダを作るのは __init__ だけで、
        消えたあとに書こうとした分は捨てる。
        """
        if not self.directory.is_dir():
            return
        tmp = self.path.with_name(f".{self.path.name}.{os.getpid()}.{threading.get_ident()}.tmp")
        try:
            tmp.write_text(json.dumps(self._data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
            os.replace(tmp, self.path)
        except OSError:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass

    def _put(self, section: str, key: str, value, limit: int | None = None) -> None:
        # 別のリクエストやジョブが足した分を消さないよう、書く直前に読み直して足す
        latest = self._load()
        entries = dict(latest.get(section) or {})
        entries.pop(key, None)
        entries[key] = value
        if limit is not None:
            while len(entries) > limit:
                entries.pop(next(iter(entries)))
        latest[section] = entries
        self._data = latest
        self._save()

    # ---- 元のソース ----

    @property
    def real(self):
        """元のファイルを開いたソース（初めて使うときに開く）。"""
        if self._real is None:
            stats = self._data.get("sheet_stats") if self.kind == "excel" else None
            self._real = self._open_real(stats)
            self._remember_stats(self._real)
        return self._real

    def _remember_stats(self, real) -> None:
        if self.kind == "excel" and not self._data.get("sheet_stats") and hasattr(real, "sheet_stats"):
            self._data = self._load()
            self._data["sheet_stats"] = real.sheet_stats()
            self._save()

    def __getattr__(self, name):
        # date1904・replaced_rows・sniff など、控えにない属性は元のソースから
        if name.startswith("_"):
            raise AttributeError(name)
        return getattr(self.real, name)

    # ---- TableSource ----

    def sheets(self) -> list[SheetInfo]:
        cached = self._data.get("sheets")
        if cached is None:
            infos = self.real.sheets()
            self._data = self._load()
            self._data["sheets"] = [asdict(s) for s in infos]
            self._save()
            return infos
        return [SheetInfo(**s) for s in cached]

    def rows(self, sheet: str, start: int = 1, limit: int | None = None) -> Iterator[SourceRow]:
        if limit is None or start < 1 or start + limit - 1 > HEAD_ROWS:
            return self.real.rows(sheet, start, limit)
        return iter([r for r in self._head(sheet) if start <= r.index < start + limit])

    def _head(self, sheet: str) -> list[SourceRow]:
        """先頭 HEAD_ROWS 行（見出しの探索・範囲の画面の表示・年度の手がかりに使う範囲）。"""
        cached = (self._data.get("head") or {}).get(sheet)
        if cached is not None:
            return [_dec_row(r) for r in cached]
        rows = list(self.real.rows(sheet, 1, HEAD_ROWS))
        try:
            self._put("head", sheet, [_enc_row(r) for r in rows])
        except _NotCacheable:
            return rows
        return rows

    # ---- 計算結果の控え ----

    def layout(self, sheet: str, anchors=None, header_row=None, data_end=None, header_rows=None,
               max_scan_rows=None) -> LayoutGuess:
        """guess_layout と同じ結果（同じ引数で前に計算していれば控えから）。"""
        key = layout_key(sheet, anchors, header_row, data_end, header_rows, max_scan_rows)
        cached = (self._data.get("layouts") or {}).get(key)
        if cached is not None:
            return _dec_layout(cached)
        memo = dict(self._data.get("scans") or {})
        known = set(memo)
        layout = guess_layout(self, sheet, anchors=anchors, header_row=header_row, data_end=data_end,
                              header_rows=header_rows, max_scan_rows=max_scan_rows, scan_memo=memo)
        for scan_key in set(memo) - known:
            self._put("scans", scan_key, memo[scan_key], MAX_LAYOUTS)
        self._put("layouts", key, asdict(layout), MAX_LAYOUTS)
        return layout

    def list_kind(self, sheet: str) -> str:
        """detect.list_kind と同じ判定（簡易判定の表の形を控えから使う）。"list" / "crosstab" / ""。"""
        try:
            layout = self.layout(sheet, max_scan_rows=200)
        except Exception:
            return ""
        if layout.table_kind in ("list", "crosstab") and layout.counts.get("data", 0) >= 10:
            return layout.table_kind
        return ""

    def looks_like_list(self, sheet: str) -> bool:
        """detect.looks_like_list と同じ判定（クロス集計も含む）。"""
        return bool(self.list_kind(sheet))

    def sample_rows(self, sheet: str, layout: LayoutGuess, n: int = 200) -> list[SourceRow]:
        """sample_data_rows と同じ行（同じ表の形なら控えから）。"""
        key = hashlib.sha256(json.dumps([sheet, asdict(layout), n], ensure_ascii=False, sort_keys=True)
                             .encode("utf-8")).hexdigest()
        cached = (self._data.get("samples") or {}).get(key)
        if cached is not None:
            return [_dec_row(r) for r in cached]
        rows = sample_data_rows(self, sheet, layout, n)
        try:
            self._put("samples", key, [_enc_row(r) for r in rows], MAX_SAMPLES)
        except _NotCacheable:
            pass
        return rows
