"""行の読み込みと正規化（ステップ5）。

1行ずつ: 行の分類 → 継続行の連結 → 結合範囲と「空欄＝上と同じ」列の補完 →
型の変換（和暦・年のない日付・8桁日付・6桁時刻・△▲/末尾マイナス・桁区切り・%・timedelta→分・単位換算）→
記録キー（出現順の番号付き）。値は NFKC だけで揃える（名寄せ辞書は使わない）。
数値・日付以外の文章列は元の値を残さない（NFKC と空白の畳み込みだけなので）。
"""
from __future__ import annotations

import re
import time as _time
import unicodedata
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel

from tables.checks import Issue
from tables.detect import classify_rows, split_header_unit
from tables.spec import TableSpec, resolve_columns

MAX_ROW_ISSUES = 2000
PROGRESS_EVERY = 500

_EXCEL_ERRORS = {"#N/A", "#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#GETTING_DATA"}
_SPACES = re.compile(r"[^\S\n]+")
_ALL_SPACES = re.compile(r"\s+")
_WEEKDAY = re.compile(r"\s*[(（]\s*[月火水木金土日](?:曜日?)?\s*[)）]\s*")
_TIME_PART = r"(?:\s*T?\s*(\d{1,2})\s*[:時]\s*(\d{1,2})\s*(?:[:分]\s*(\d{1,2})\s*秒?)?分?)?"
_YMD = re.compile(r"^(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})\s*日?" + _TIME_PART + r"$")
_ERA_YMD = re.compile(r"^(令和|平成|昭和|R|H|S)\s*(\d{1,2}|元)\s*[./年\-]\s*(\d{1,2})\s*[./月\-]\s*(\d{1,2})\s*日?"
                      + _TIME_PART + r"$", re.I)
_MD = re.compile(r"^(\d{1,2})\s*[/月]\s*(\d{1,2})\s*日?" + _TIME_PART + r"$")
_YMD8 = re.compile(r"^((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?:\s*(\d{2})(\d{2})(\d{2})?|\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$")
_TIME_TEXT = re.compile(r"^(\d{1,2})\s*[:時]\s*(\d{1,2})\s*(?:[:分]\s*(\d{1,2})\s*秒?)?分?$")
_TIME_DIGITS = re.compile(r"^(\d{2})(\d{2})(\d{2})?$")
_NUMBER = re.compile(
    r"^(?P<sign>[△▲\-−+])?\s*[¥￥]?\s*(?P<int>\d{1,3}(?:,\d{3})+|\d+)?(?P<dec>\.\d+)?(?:[eE](?P<exp>[+\-]?\d+))?"
    r"\s*(?P<tail>-)?\s*(?P<unit>[^\d\s,.\-+]{1,6})?$")
_ERA_BASE = {"令和": 2018, "R": 2018, "平成": 1988, "H": 1988, "昭和": 1925, "S": 1925}

# 単位の換算表（列の単位ごと。値は「元の単位1つが列の単位でいくつか」）
_UNIT_ALIASES = {"h": "時間", "hr": "時間", "hrs": "時間", "hour": "時間", "hours": "時間", "H": "時間",
                 "min": "分", "mins": "分", "sec": "秒", "yen": "円", "¥": "円", "￥": "円",
                 "％": "%"}
_BUILTIN_CONVERSIONS = {
    "分": {"分": 1, "時間": 60, "秒": 1 / 60, "日": 1440},
    "時間": {"時間": 1, "分": 1 / 60, "秒": 1 / 3600, "日": 24},
    "秒": {"秒": 1, "分": 60, "時間": 3600},
    "円": {"円": 1, "千円": 1000, "万円": 10000, "百万円": 1000000},
    "千円": {"千円": 1, "円": 0.001, "万円": 10},
    "%": {"%": 1},
}

_TYPE_LABELS = {"date": "日付", "datetime": "日時", "time": "時刻", "number": "数値"}


# ---- データ構造 ---------------------------------------------------------------------

@dataclass
class RecordRow:
    key: str
    values: dict[str, object]
    originals: dict[str, str]
    source: dict  # {"file", "sheet", "row"}
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"key": self.key, "values": self.values, "originals": self.originals,
                "source": self.source, "warnings": self.warnings}


@dataclass
class ImportStats:
    file: str = ""
    sheet: str = ""
    rows_scanned: int = 0
    data_rows: int = 0
    records: int = 0
    blank_rows: int = 0
    excluded: dict = field(default_factory=dict)  # 理由 → 行数
    continuation_merged: int = 0
    subtotal_rows: int = 0
    included_hidden: int = 0
    type_errors: dict = field(default_factory=dict)
    type_checked: dict = field(default_factory=dict)
    year_inferred: int = 0
    year_missing: int = 0
    year_context: dict = field(default_factory=dict)
    year_context_label: str = ""
    error_values: int = 0
    filled_down: dict = field(default_factory=dict)
    unused_headers: list = field(default_factory=list)
    missing_required: list = field(default_factory=list)
    missing_optional: list = field(default_factory=list)
    positions: dict = field(default_factory=dict)
    duplicate_keys: dict = field(default_factory=dict)
    reconcile: list = field(default_factory=list)
    replaced_rows: list = field(default_factory=list)
    date_min: str | None = None
    date_max: str | None = None
    months: dict = field(default_factory=dict)
    layout_warnings: list = field(default_factory=list)
    row_issues_truncated: int = 0
    key_columns: list = field(default_factory=list)
    elapsed_sec: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)


# ---- 年・年度の手がかり -------------------------------------------------------------------

_FY_PATTERNS = [
    (re.compile(r"(?:FY|ＦＹ)\s*'?(\d{4}|\d{2})(?!\d)", re.I), lambda m: _year4(m.group(1))),
    (re.compile(r"(?<!\d)(\d{4})\s*年度"), lambda m: int(m.group(1))),
    (re.compile(r"(令和|平成|R|H)\s*(\d{1,2}|元)\s*年度"), lambda m: _era_year(m.group(1), m.group(2))),
]
_YM_PATTERNS = [
    re.compile(r"(?<!\d)(\d{4})\s*年\s*(\d{1,2})\s*月"),
    re.compile(r"(?<!\d)(\d{4})[/\-._](\d{1,2})(?![\d/\-.])"),
    re.compile(r"(?<!\d)(\d{4})(0[1-9]|1[0-2])(?!\d)"),
]
_Y_PATTERN = re.compile(r"(?<!\d)(\d{4})\s*年(?!度)")


def _year4(text: str) -> int:
    n = int(text)
    return n + 2000 if n < 100 else n


def _era_year(era: str, num: str) -> int:
    base = _ERA_BASE.get(era.upper() if len(era) == 1 else era, 2018)
    return base + (1 if num == "元" else int(num))


def find_year_context(sources: list[tuple[str, str]]) -> dict:
    """タイトル行・シート名・ファイル名から年度・年月を探す。sources: [(出どころの名前, 文字列)]（優先順）。

    戻り値: {"fiscal_year", "year", "month", "source"}（見つからなければ値は None）
    """
    for label, text in sources:
        s = unicodedata.normalize("NFKC", str(text or ""))
        if not s.strip():
            continue
        for pattern, conv in _FY_PATTERNS:
            m = pattern.search(s)
            if m:
                fy = conv(m)
                if 1950 <= fy <= 2100:
                    return {"fiscal_year": fy, "year": None, "month": None, "source": label}
        for pattern in _YM_PATTERNS:
            m = pattern.search(s)
            if m and 1950 <= int(m.group(1)) <= 2100 and 1 <= int(m.group(2)) <= 12:
                return {"fiscal_year": None, "year": int(m.group(1)), "month": int(m.group(2)), "source": label}
        m = _Y_PATTERN.search(s)
        if m and 1950 <= int(m.group(1)) <= 2100:
            return {"fiscal_year": None, "year": int(m.group(1)), "month": None, "source": label}
    return {"fiscal_year": None, "year": None, "month": None, "source": ""}


def year_context_for(source, sheet: str, layout, file_name: str) -> dict:
    """表の上のタイトル行 → シート名 → ファイル名の順で年度・年月を探す。"""
    sources: list[tuple[str, str]] = []
    top = (layout.header_rows[0] - 1) if getattr(layout, "header_rows", None) else 0
    if top > 0:
        texts = []
        for row in source.rows(sheet, 1, min(top, 20)):
            texts.append(" ".join(c.text for c in row.cells if c.text))
        sources.append(("タイトル行", "\n".join(texts)))
    if getattr(source, "kind", "") == "excel":
        sources.append(("シート名", sheet))
    sources.append(("ファイル名", Path(file_name or "").stem))
    return find_year_context(sources)


# ---- 値の変換 ---------------------------------------------------------------------------

@dataclass
class ConvertContext:
    fiscal_year: int | None = None
    year: int | None = None
    month: int | None = None
    fiscal_start: int = 4
    date1904: bool = False
    na_tokens: frozenset = frozenset()


def _nfkc(text) -> str:
    return unicodedata.normalize("NFKC", str(text or ""))


def nfkc_text(text) -> str:
    """NFKC＋行内の空白の畳み込み（改行は残す）。"""
    if text is None:
        return ""
    s = _nfkc(text).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [_SPACES.sub(" ", line).strip() for line in s.split("\n")]
    return "\n".join(lines).strip("\n")


def _fmt_dt(y: int, mo: int, d: int, hh=None, mm=None, ss=None) -> tuple[str | None, str | None]:
    try:
        day = date(int(y), int(mo), int(d))
    except (TypeError, ValueError):
        return None, None
    if hh is None or hh == "":
        return day.isoformat(), None
    try:
        t = time(int(hh), int(mm or 0), int(ss or 0))
    except ValueError:
        return None, None
    return day.isoformat(), _fmt_time(t)


def _fmt_time(t: time) -> str:
    return t.strftime("%H:%M:%S") if t.second else t.strftime("%H:%M")


def _join_dt(d: str | None, t: str | None, type_: str) -> str | None:
    if d is None:
        return None
    if type_ == "datetime" and t and t not in ("00:00",):
        return f"{d} {t}"
    return d


def parse_date_text(text: str, cctx: ConvertContext) -> tuple[str | None, str | None, str | None]:
    """文字列の日付・日時。戻り値: (日付, 時刻, 印)。印は year_inferred / year_missing / None。"""
    s = _WEEKDAY.sub(" ", _nfkc(text)).strip()
    s = re.sub(r"\s+", " ", s)
    m = _YMD.match(s)
    if m:
        d, t = _fmt_dt(m.group(1), m.group(2), m.group(3), m.group(4), m.group(5), m.group(6))
        return d, t, None
    m = _YMD8.match(s)
    if m:
        hh, mm, ss = (m.group(4), m.group(5), m.group(6)) if m.group(4) else (m.group(7), m.group(8), m.group(9))
        d, t = _fmt_dt(m.group(1), m.group(2), m.group(3), hh, mm, ss)
        return d, t, None
    m = _ERA_YMD.match(s)
    if m:
        era = m.group(1)
        y = _era_year(era.upper() if len(era) == 1 else era, m.group(2))
        d, t = _fmt_dt(y, m.group(3), m.group(4), m.group(5), m.group(6), m.group(7))
        return d, t, None
    m = _MD.match(s)
    if m:
        mo = int(m.group(1))
        if cctx.fiscal_year:
            y = cctx.fiscal_year if mo >= cctx.fiscal_start else cctx.fiscal_year + 1
        elif cctx.year:
            y = cctx.year
        else:
            return None, None, "year_missing"
        d, t = _fmt_dt(y, mo, m.group(2), m.group(3), m.group(4), m.group(5))
        return d, t, ("year_inferred" if d else None)
    return None, None, None


def _excel_serial_date(value: float, date1904: bool) -> datetime | None:
    low = 3000
    if not (low < value < 80000):
        return None
    try:
        dt = from_excel(value, epoch=MAC_EPOCH if date1904 else WINDOWS_EPOCH)
    except (ValueError, OverflowError):
        return None
    return dt if isinstance(dt, datetime) else None


def _parse_time_text(text: str) -> str | None:
    s = _nfkc(text).strip()
    m = _TIME_TEXT.match(s) or _TIME_DIGITS.match(s)
    if not m:
        return None
    try:
        t = time(int(m.group(1)), int(m.group(2)), int(m.group(3) or 0))
    except ValueError:
        return None
    return _fmt_time(t)


def _unit_norm(unit: str) -> str:
    u = _nfkc(unit).strip()
    return _UNIT_ALIASES.get(u, _UNIT_ALIASES.get(u.lower(), u))


def unit_factor(src_unit: str, target_unit: str, conversions: dict | None = None) -> float | None:
    """src_unit の値を target_unit に直す倍率。換算できなければ None。"""
    if not src_unit:
        return 1.0
    conversions = conversions or {}
    for name in (src_unit, _unit_norm(src_unit)):
        if name in conversions:
            return float(conversions[name])
    src, dst = _unit_norm(src_unit), _unit_norm(target_unit)
    if not dst or src == dst:
        return 1.0
    return _BUILTIN_CONVERSIONS.get(dst, {}).get(src)


def _clean_number(v: float) -> int | float:
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, int):
        return v
    v = round(float(v), 9)
    if v.is_integer() and abs(v) < 1e15:
        return int(v)
    return v


def parse_number_text(text: str) -> tuple[float | None, str]:
    """数値の文字列。戻り値: (値, 末尾の単位)。△▲・末尾マイナス・桁区切り・全角数字に対応。"""
    s = _nfkc(text).strip().replace(" ", "")
    if not s:
        return None, ""
    m = _NUMBER.match(s)
    if not m or (m.group("int") is None and m.group("dec") is None):
        return None, ""
    number = float((m.group("int") or "0").replace(",", "") + (m.group("dec") or ""))
    if m.group("exp"):
        number *= 10 ** int(m.group("exp"))
    if (m.group("sign") and m.group("sign") in "△▲-−") or m.group("tail"):
        number = -number
    return number, m.group("unit") or ""


def _zero_pad(value: float, number_format: str | None) -> str:
    n = int(value)
    fmt = (number_format or "").strip()
    if fmt and set(fmt) == {"0"} and len(fmt) > 1:
        return str(n).zfill(len(fmt))
    return str(n)


def convert_cell(value, text: str, col, cctx: ConvertContext, number_format: str | None = None,
                 header_unit: str = "") -> tuple[object, str | None, str | None]:
    """セル1つを列の型に変換する。戻り値: (値, エラー文, 印)。印: year_inferred / year_missing / error_value"""
    if value is None and not text:
        return None, None, None
    stripped = _nfkc(text).strip()
    if isinstance(value, str) or value is None:
        if not stripped:
            return None, None, None
        if stripped in _EXCEL_ERRORS:
            return None, None, "error_value"
        if stripped in cctx.na_tokens:
            return None, None, None
    t = col.type
    if t in ("date", "datetime"):
        return _convert_date(value, text, t, cctx)
    if t == "time":
        return _convert_time(value, text)
    if t == "number":
        return _convert_number(value, text, col, number_format, header_unit)
    if t == "code":
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            out = _zero_pad(value, number_format) if float(value).is_integer() else str(_clean_number(value))
        else:
            out = _SPACES.sub(" ", nfkc_text(text)).strip()
        if "upper" in col.normalize:
            out = out.upper()
        return out, None, None
    if t == "text":
        out = nfkc_text(text)
        return (out or None), None, None
    # string / enum / status
    out = nfkc_text(text) if "nfkc" in col.normalize or not col.normalize else str(text).strip()
    if "upper" in col.normalize:
        out = out.upper()
    if col.value_map:
        mapped = col.value_map.get(out)
        if mapped is None:
            mapped = {nfkc_text(k): v for k, v in col.value_map.items()}.get(out)
        if mapped is not None:
            out = str(mapped)
    return (out or None), None, None


def _convert_date(value, text, type_, cctx):
    label = _TYPE_LABELS[type_]
    if isinstance(value, datetime):
        d, tm = value.date().isoformat(), (_fmt_time(value.time()) if value.time() != time(0) else None)
        return _join_dt(d, tm, type_), None, None
    if isinstance(value, date):
        return value.isoformat(), None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, int) or float(value).is_integer():
            d, tm, _ = parse_date_text(str(int(value)), cctx)
            if d:
                return _join_dt(d, tm, type_), None, None
        dt = _excel_serial_date(float(value), cctx.date1904)
        if dt is not None:
            tm = _fmt_time(dt.time()) if dt.time() != time(0) else None
            return _join_dt(dt.date().isoformat(), tm, type_), None, None
        return str(_clean_number(value)), f"「{text}」を{label}に変換できません", None
    if isinstance(value, (time, timedelta)):
        return text, f"「{text}」を{label}に変換できません", None
    d, tm, flag = parse_date_text(text, cctx)
    if d:
        return _join_dt(d, tm, type_), None, flag
    if flag == "year_missing":
        return nfkc_text(text), f"「{text}」は年がないため{label}にできません", flag
    return nfkc_text(text), f"「{_short(text)}」を{label}に変換できません", None


def _convert_time(value, text):
    if isinstance(value, datetime):
        return _fmt_time(value.time()), None, None
    if isinstance(value, time):
        return _fmt_time(value), None, None
    if isinstance(value, timedelta):
        secs = int(round(value.total_seconds()))
        if 0 <= secs < 86400:
            return _fmt_time(time(secs // 3600, secs % 3600 // 60, secs % 60)), None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = float(value)
        if 0 <= v < 1:
            secs = int(round(v * 86400))
            secs = min(secs, 86399)
            return _fmt_time(time(secs // 3600, secs % 3600 // 60, secs % 60)), None, None
        parsed = _parse_time_text(str(int(v))) if v.is_integer() else None
        if parsed:
            return parsed, None, None
        return str(_clean_number(v)), f"「{text}」を時刻に変換できません", None
    parsed = _parse_time_text(text)
    if parsed:
        return parsed, None, None
    return nfkc_text(text), f"「{_short(text)}」を時刻に変換できません", None


def _convert_number(value, text, col, number_format, header_unit):
    target = col.unit or ""
    if isinstance(value, bool):
        return None, f"「{text}」を数値に変換できません", None
    if isinstance(value, timedelta) or isinstance(value, time):
        minutes = value.total_seconds() / 60 if isinstance(value, timedelta) else value.hour * 60 + value.minute + value.second / 60
        factor = unit_factor("分", target or "分", col.unit_conversions)
        if factor is None:
            return None, f"「{text}」（時間）を{target}に換算できません", None
        return _clean_number(minutes * factor), None, None
    if isinstance(value, (int, float)):
        v = float(value)
        if number_format and "%" in number_format:
            v *= 100
        factor = unit_factor(header_unit, target, col.unit_conversions) if header_unit else 1.0
        if factor is None:
            factor = 1.0
        return _clean_number(v * factor), None, None
    number, unit = parse_number_text(text)
    if number is None:
        return nfkc_text(text), f"「{_short(text)}」を数値に変換できません", None
    src_unit = unit or header_unit
    factor = unit_factor(src_unit, target, col.unit_conversions)
    if factor is None:
        if unit and not target:
            factor = 1.0  # 列に単位がなければ数値部分だけを使う
        else:
            return nfkc_text(text), f"「{_short(text)}」の単位「{unit or header_unit}」を{target}に換算できません", None
    return _clean_number(number * factor), None, None


def _short(text, n: int = 30) -> str:
    s = str(text or "").replace("\n", " ")
    return s if len(s) <= n else s[:n] + "…"


# ---- 読み込み本体 ---------------------------------------------------------------------------

def _is_text_column(col) -> bool:
    return col.type == "text" or col.role in ("text", "log")


def read_records(source, source_opts: dict | None, layout, spec: TableSpec, on_progress=None) -> tuple[list[RecordRow], list[Issue], ImportStats]:
    """表ソースの全行を読み、正規化した記録・問題・件数を返す。

    source_opts: {"sheet", "file_name", "fiscal_year"（年度を手で指定するとき）, "key_columns"（継続行の判定列、0始まり）}
    on_progress(done, total): 一定行ごとに呼ぶ（ジョブの進捗・中止確認に使う）
    """
    started = _time.monotonic()
    opts = dict(source_opts or {})
    sheet = opts.get("sheet") or layout.sheet
    file_name = opts.get("file_name") or getattr(source, "file_name", "") or ""
    is_csv = getattr(source, "kind", "") == "csv"
    stats = ImportStats(file=file_name, sheet="" if is_csv else sheet,
                        layout_warnings=list(getattr(layout, "warnings", []) or []))
    issues: list[Issue] = []

    res = resolve_columns(spec, list(layout.headers))
    stats.positions = dict(res.positions)
    stats.unused_headers = list(res.unused_headers)
    stats.missing_required = list(res.missing_required)
    stats.missing_optional = [d for d in res.missing if d not in res.missing_required]

    ctx_year = year_context_for(source, sheet, layout, file_name)
    if opts.get("fiscal_year"):
        ctx_year = {"fiscal_year": int(opts["fiscal_year"]), "year": None, "month": None, "source": "指定"}
    stats.year_context = ctx_year
    if ctx_year.get("fiscal_year"):
        stats.year_context_label = f"{ctx_year['source']}の{ctx_year['fiscal_year']}年度"
    elif ctx_year.get("year"):
        stats.year_context_label = f"{ctx_year['source']}の{ctx_year['year']}年"
    na = frozenset(_nfkc(t).strip() for t in (spec.na_tokens or []))
    cctx = ConvertContext(fiscal_year=ctx_year.get("fiscal_year"), year=ctx_year.get("year"), month=ctx_year.get("month"),
                          fiscal_start=int(spec.fiscal_year_start_month or 4),
                          date1904=bool(getattr(source, "date1904", False)), na_tokens=na)

    columns = [(c, res.positions[c.key]) for c in spec.columns if c.key in res.positions]
    header_units = {c.key: split_header_unit(layout.headers[pos])[1] for c, pos in columns}
    measure_cols = [(c, pos) for c, pos in columns if c.type == "number"]
    policies = spec.exclude or {}

    anchors: dict[tuple[int, int], tuple[object, str, str | None]] = {}
    prev_raw: dict[str, tuple[object, str, str | None]] = {}
    pending: list[RecordRow] = []
    last_record: RecordRow | None = None
    running: dict[int, float] = {}
    grand: dict[int, float] = {}
    total_rows = max(1, layout.data_end - layout.data_start + 1)
    row_issue_count = 0

    def add_row_issue(issue: Issue):
        nonlocal row_issue_count
        row_issue_count += 1
        if row_issue_count <= MAX_ROW_ISSUES:
            issues.append(issue)
        else:
            stats.row_issues_truncated += 1

    def cell_raw(row, pos: int) -> tuple[object, str, str | None]:
        cell = row.cell(pos)
        if cell is None:
            return None, "", None
        if not cell.text and cell.merged_anchor and cell.merged_anchor != (row.index, pos + 1):
            found = anchors.get(cell.merged_anchor)
            if found is not None:
                return found
        return cell.value, cell.text, cell.number_format

    def convert(col, raw, row_index: int, record_warnings: list[str], originals: dict, header_unit: str = ""):
        value, text, number_format = raw
        out, error, flag = convert_cell(value, text, col, cctx, number_format, header_unit)
        if col.type in ("date", "datetime", "time", "number"):
            if value is not None or text:
                stats.type_checked[col.key] = stats.type_checked.get(col.key, 0) + (0 if out is None and not error else 1)
        if flag == "error_value":
            stats.error_values += 1
        elif flag == "year_inferred":
            stats.year_inferred += 1
            record_warnings.append(f"{col.display}の年を{stats.year_context_label or '年度'}から補った")
        elif flag == "year_missing":
            stats.year_missing += 1
        if error:
            stats.type_errors[col.key] = stats.type_errors.get(col.key, 0) + 1
            record_warnings.append(error)
            add_row_issue(Issue("warning", "type_error", error, row=row_index, column=col.display))
        if out is not None and not _is_text_column(col) and text and str(out) != text:
            originals[col.key] = text
        return out

    def reconcile_row(row, label: str, level: str):
        acc = running if level == "小計" else grand
        positions = [(c.display, pos, c, header_units.get(c.key, "")) for c, pos in measure_cols]
        for display, pos, col, unit in positions:
            value, text, fmt = cell_raw(row, pos)
            number, error, _flag = convert_cell(value, text, col, cctx, fmt, unit)
            if error or not isinstance(number, (int, float)) or isinstance(number, bool):
                continue
            stats.reconcile.append({"row": row.index, "label": f"{label}の行" if label else level,
                                    "column": display, "expected": _clean_number(number),
                                    "actual": _clean_number(acc.get(pos, 0.0))})
        if level == "小計":
            running.clear()

    def accumulate(pos: int, number):
        if isinstance(number, (int, float)):
            running[pos] = running.get(pos, 0.0) + number
            grand[pos] = grand.get(pos, 0.0) + number

    # 継続行の判定に使う列: 指定 → 設定のキー・日付の列 → 自動判定
    key_columns = list(opts.get("key_columns") or [])
    if not key_columns:
        key_columns = sorted(pos for c, pos in columns if c.role in ("key", "date"))
    if not key_columns:
        key_columns = list(layout.key_columns or [])
    stats.key_columns = key_columns
    for n, (row, rc) in enumerate(classify_rows(source, sheet, layout, key_columns=key_columns)):
        stats.rows_scanned += 1
        if on_progress is not None and n % PROGRESS_EVERY == 0:
            on_progress(n, total_rows)
        for i, cell in enumerate(row.cells):
            if cell.text and cell.merged_anchor == (row.index, i + 1):
                anchors[cell.merged_anchor] = (cell.value, cell.text, cell.number_format)
        kind, reason = rc.kind, rc.reason
        if kind == "blank":
            stats.blank_rows += 1
            continue
        if kind == "excluded" and reason in ("非表示の行", "取り消し線の行"):
            policy = policies.get("hidden_rows" if reason == "非表示の行" else "strike_rows", "exclude_with_warning")
            if policy == "include":
                kind = "data"
                stats.included_hidden += 1
        if kind == "subtotal":
            stats.subtotal_rows += 1
            stats.excluded[reason or "小計"] = stats.excluded.get(reason or "小計", 0) + 1
            label = next((c.text for c in row.cells if c.text), "")
            reconcile_row(row, _nfkc(label).strip()[:20], "合計" if reason == "合計" else "小計")
            continue
        if kind in ("note", "title", "header", "excluded"):
            key = reason or {"note": "注記", "title": "表題", "header": "見出し"}.get(kind, "除外")
            stats.excluded[key] = stats.excluded.get(key, 0) + 1
            continue
        if kind == "continuation":
            if spec.continuation_rows == "merge_into_previous" and last_record is not None:
                merged = False
                for col, pos in columns:
                    if not _is_text_column(col):
                        continue
                    text = row.text(pos)
                    if not text:
                        continue
                    addition = nfkc_text(text)
                    current = last_record.values.get(col.key)
                    last_record.values[col.key] = f"{current}\n{addition}" if current else addition
                    merged = True
                if merged:
                    stats.continuation_merged += 1
                    continue
            kind = "data"

        stats.data_rows += 1
        record_warnings: list[str] = []
        originals: dict[str, str] = {}
        values: dict[str, object] = {}
        for col, pos in columns:
            raw = cell_raw(row, pos)
            if col.fill_down_blank and not raw[1] and raw[0] is None:
                if col.key in prev_raw:
                    raw = prev_raw[col.key]
                    stats.filled_down[col.key] = stats.filled_down.get(col.key, 0) + 1
            elif raw[1] or raw[0] is not None:
                prev_raw[col.key] = raw
            out = convert(col, raw, row.index, record_warnings, originals, header_units.get(col.key, ""))
            if out is not None:
                values[col.key] = out
        src = {"file": file_name, "sheet": stats.sheet, "row": row.index}

        for col, pos in measure_cols:
            accumulate(pos, values.get(col.key))
        record = RecordRow("", values, originals, src, record_warnings)
        pending.append(record)
        last_record = record

    replaced = getattr(source, "replaced_rows", None)
    if replaced:
        stats.replaced_rows = list(replaced)
    _assign_keys(pending, spec, stats)
    date_key = spec.date_key
    months: Counter[str] = Counter()
    dates = []
    for rec in pending:
        v = rec.values.get(date_key)
        if v:
            s = str(v)
            dates.append(s[:10])
            if len(s) >= 7 and s[4] == "-":
                months[s[:7]] += 1
    stats.months = dict(sorted(months.items()))
    stats.date_min = min(dates) if dates else None
    stats.date_max = max(dates) if dates else None
    stats.records = len(pending)
    if stats.row_issues_truncated:
        issues.append(Issue("warning", "issues_truncated",
                            f"行ごとの問題が多いため、{MAX_ROW_ISSUES}件を超えた分（{stats.row_issues_truncated}件）は省略しました"))
    if on_progress is not None:
        on_progress(total_rows, total_rows)
    stats.elapsed_sec = round(_time.monotonic() - started, 2)
    return pending, issues, stats


def _key_part(values: dict, part: str) -> str:
    name, _, length = str(part).partition(":")
    value = values.get(name)
    if value in (None, ""):
        return ""
    s = str(value)
    if length.isdigit():
        s = _ALL_SPACES.sub("", s)[: int(length)]
    return s


def _assign_keys(records: list[RecordRow], spec: TableSpec, stats: ImportStats) -> None:
    """記録キー。キー列がそろわない行は代わりのキー。同じキーは出現順に #2, #3 を付ける。"""
    record = spec.record or {}
    key_parts = [str(p) for p in (record.get("key") or [])]
    fallback = [str(p) for p in (record.get("fallback_key") or [])]
    bases: list[tuple[str, bool]] = []
    for rec in records:
        base, primary = "", False
        if key_parts:
            parts = [_key_part(rec.values, p) for p in key_parts]
            if all(parts):
                base, primary = "|".join(parts), True
        if not base and fallback:
            parts = [_key_part(rec.values, p) for p in fallback]
            if any(parts):
                base = "|".join(parts)
        if not base:
            base = f"行{rec.source.get('row')}"
        bases.append((base, primary))
    counts = Counter(b for b, _ in bases)
    seen: Counter[str] = Counter()
    for rec, (base, primary) in zip(records, bases):
        seen[base] += 1
        rec.key = base if seen[base] == 1 else f"{base}#{seen[base]}"
        if counts[base] > 1 and primary:
            stats.duplicate_keys[base] = counts[base]
