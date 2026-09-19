"""表ソースの共通モデル。ExcelもCSVも読み込み後はこの形にそろえる。

見出しの検出・列の対応づけ・正規化は、すべてこのインターフェースだけを使う。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterator, Protocol

try:
    from core.files import UploadError
except ImportError:  # WP-core 未統合の間の代替
    class UploadError(Exception):
        """利用者に見せる読み込みエラー（日本語メッセージ）。"""


EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
CSV_EXTENSIONS = {".csv", ".tsv", ".txt"}

# 制御文字と見えない文字（ゼロ幅スペース・BOM・向き指定・ソフトハイフン。excel.text._INVISIBLE_RE と同じ集合）
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b-\u200d\u2060\ufeff\u202a-\u202e\u2066-\u2069\u00ad]")
_EXCEL_ERRORS = {"#N/A", "#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#GETTING_DATA"}


@dataclass
class CellInfo:
    value: object
    text: str
    number_format: str | None = None
    bold: bool = False
    strike: bool = False
    fill: bool = False
    merged_anchor: tuple[int, int] | None = None  # 結合範囲の左上 (行, 列)。1始まり。左上セル自身も自分を指す


@dataclass
class SourceRow:
    index: int  # 1始まり。CSVはレコード番号（Excelで開いたときの行番号と同じ）
    cells: list[CellInfo]
    hidden: bool | None = None  # CSVは不明なので None

    def text(self, col: int) -> str:
        """0始まりの列位置の文字列。範囲外は空。"""
        return self.cells[col].text if 0 <= col < len(self.cells) else ""

    def cell(self, col: int) -> CellInfo | None:
        return self.cells[col] if 0 <= col < len(self.cells) else None

    @property
    def is_blank(self) -> bool:
        return not any(c.text for c in self.cells)


@dataclass
class SheetInfo:
    name: str
    hidden: bool
    max_row: int
    max_col: int
    table_ranges: list[str]
    date1904: bool = False
    hidden_columns: list[int] = field(default_factory=list)  # 1始まりの列番号
    freeze_panes: str | None = None
    auto_filter: str | None = None


class TableSource(Protocol):
    kind: str  # "csv" | "excel"
    file_name: str

    def sheets(self) -> list[SheetInfo]: ...

    def rows(self, sheet: str, start: int = 1, limit: int | None = None) -> Iterator[SourceRow]: ...


def open_source(path: Path, file_name: str, options: dict | None = None) -> TableSource:
    """拡張子で CSV / Excel を振り分けて開く。"""
    ext = Path(file_name or str(path)).suffix.lower()
    if ext in EXCEL_EXTENSIONS:
        from tables.excel_source import ExcelSource

        return ExcelSource(path, file_name, options)
    if ext in CSV_EXTENSIONS:
        from tables.csv_source import CsvSource

        return CsvSource(path, file_name, options)
    raise UploadError(f"対応していない形式です（{ext or '拡張子なし'}）。.xlsx / .xlsm / .csv / .tsv / .txt を選んでください")


def clean_text(text: str) -> str:
    """改行をLFにそろえ、制御文字を除いて前後の空白を削る。"""
    if not text:
        return ""
    s = text
    if "_x000D_" in s:
        s = s.replace("_x000D_", "")
    if "\r" in s:
        s = s.replace("\r\n", "\n").replace("\r", "\n")
    if _CONTROL_RE.search(s):  # 多くのセルは置き換えるものがないので、先に確かめてから置き換える
        s = _CONTROL_RE.sub("", s)
    return s.strip()


def cell_text(value) -> str:
    """セル値を表示用の文字列にする。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        if value.time() == time(0):
            return value.date().isoformat()
        if value.second:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M:%S" if value.second else "%H:%M")
    if isinstance(value, timedelta):
        minutes = int(value.total_seconds() // 60)
        return f"{minutes // 60}:{minutes % 60:02d}"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(round(value, 10))
    return clean_text(str(value))


# ---- 値の種類の判定（見出しらしさ・型推定で共用） ----

NA_TOKENS = {"-", "－", "―", "‐", "ー", "N/A", "n/a", "NA", "#N/A"}
_NUM_RE = re.compile(r"^[△▲+\-]?[¥]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?[%]?-?$")
_DATE_RE = re.compile(r"^(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?(\s*\(?[月火水木金土日]?\)?)?"
                      r"((?:\s+|T)\d{1,2}:\d{2}(:\d{2}(\.\d{1,7})?)?(\s*(Z|[+\-]\d{2}:?\d{2}))?)?$")
_ERA_RE = re.compile(r"^(R|H|S|令和|平成|昭和)\s*(\d{1,2}|元)[./年](\d{1,2})[./月](\d{1,2})日?(\s+\d{1,2}:\d{2}(:\d{2})?)?$")
_MD_RE = re.compile(r"^(\d{1,2})/(\d{1,2})$")
_YMD8_RE = re.compile(r"^(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$")
_TIME_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
_CODE_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-_/.#]*$")


def value_kind(value, text: str) -> str:
    """セル値の種類: blank/number/date/datetime/time/text/code/string/bool/error"""
    if not text:
        return "blank"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, datetime):
        return "date" if value.time() == time(0) else "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, time):
        return "time"
    if isinstance(value, (int, float, timedelta)):
        return "number"
    s = unicodedata.normalize("NFKC", text).strip()
    if s in _EXCEL_ERRORS:
        return "error"
    if "\n" in s or len(s) > 40:
        return "text"
    if _YMD8_RE.match(s):
        return "date"
    m = _DATE_RE.match(s) or _ERA_RE.match(s)
    if m:
        return "datetime" if ":" in s else "date"
    if _MD_RE.match(s):
        return "date"
    if _TIME_RE.match(s):
        return "time"
    if _NUM_RE.match(s) and not (len(s) > 1 and s[0] == "0" and s[1].isdigit()):
        return "number"
    if _CODE_RE.match(s) and any(ch.isdigit() for ch in s):
        return "code"
    return "string"
