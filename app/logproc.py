"""「経過の記録」の列（画面の役割名。コードでは log）のルール処理。

分割・日時・記入者・識別子・マスク・用語集・時系列の描画。純粋関数のみ。
この層の名前（logproc / log / date_log）は変えていない（画面の言葉だけ 2026-09-21 に変えた）。
"""
from __future__ import annotations

import re
import threading
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date, timedelta
from functools import lru_cache
from typing import Iterable



# ====================================================================================================
# 元 logproc/models.py
# 「経過の記録」（対応内容など。コードでは log）のルール処理で使うデータ構造。
# ====================================================================================================

# 「1.5mm」「2.5A」のような寸法・電気量は日付にしない。ただし単位の後ろに英数字・「-」・カタカナが続くときは
# 「4.3 AGV」「4.3 ALM-2031」「4.3 Aライン」「4.3 Vベルト」のように日付＋設備名などなので、単位とみなさない
DEFAULT_NOT_DATE_PATTERNS = [r"\d+\.\d+\s*(?:mm|MPa|V|A)(?![A-Za-z0-9\-\u30A0-\u30FF])", r"納期\s*\d+/\d+"]


@dataclass
class WhenInfo:
    text: str                      # 原文の日時表現（例「4/1 10:00」「翌週」）。継承時は空か時刻のみ
    date: str | None               # ISO日付（範囲なら開始日）
    time: str | None               # "HH:MM"
    date_to: str | None            # 範囲の終了日（翌週など）
    shift: str | None              # 夜勤/日勤/2直/夕方 など
    estimated: bool
    note: str = ""                 # 要確認に出す説明（「原文「翌週」から推定した（基準は…）」など）
    how: str = ""                  # explicit/year_inferred/relative/inherited/base/unresolved/year_unknown/none


@dataclass
class AuthorInfo:
    raw: str                       # 原文の表記（「K.T」「保全G 高橋」など）
    name: str | None               # 特定できた名前。イニシャル未登録などは None
    estimated: bool
    note: str = ""


@dataclass
class Segment:
    id: str                        # s1..（原文の並び順）
    raw: str                       # 原文（クリーニング後テキストの start:end）
    body: str                      # 日時・記入者の部分を除いた原文
    start: int
    end: int
    when: WhenInfo | None
    author: AuthorInfo | None
    identifiers: list[str] = field(default_factory=list)
    quantities: list[str] = field(default_factory=list)
    plans: list[str] = field(default_factory=list)
    marks: list[str] = field(default_factory=list)   # email/header_cell/bullet/checklist/note/sentence_split/reference/correction
    label: str = ""                # 見出し型セルの見出し語（【現象】→「現象」）


@dataclass
class LogParse:
    segments: list[Segment]
    order: str                     # asc/desc/unknown
    kind: str                      # log/header_cell/single/empty
    warnings: list[str] = field(default_factory=list)
    text: str = ""                 # 分割に使ったテキスト（_x000D_ と CRLF を除いたもの。start/end はこの位置）

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class SplitOptions:
    order: str = "auto"                    # auto/asc/desc
    sentence_split_min_chars: int = 120    # 目印のない長いセグメントを「。」で分ける長さ
    header_cells: str = "detect"           # detect/off
    extra_anchors: list[str] = field(default_factory=list)       # 行頭の区切りを正規表現で追加
    not_date_patterns: list[str] = field(default_factory=lambda: list(DEFAULT_NOT_DATE_PATTERNS))
    time_only_lines: str = "separate"      # separate（別セグメント、日付は直前から）/ join（直前につなぐ）
    order_tolerance_days: int = 60         # この日数以内の逆行は年を変えず「前後している」とする

    @classmethod
    def from_dict(cls, d: dict | None) -> "SplitOptions":
        d = d or {}
        opts = cls()
        for key in ("order", "header_cells", "time_only_lines"):
            if d.get(key):
                setattr(opts, key, str(d[key]))
        split = d.get("sentence_split")
        if isinstance(split, dict) and split.get("min_chars"):
            opts.sentence_split_min_chars = int(split["min_chars"])
        if d.get("sentence_split_min_chars"):
            opts.sentence_split_min_chars = int(d["sentence_split_min_chars"])
        if d.get("extra_anchors"):
            opts.extra_anchors = [str(x) for x in d["extra_anchors"]]
        if d.get("not_date_patterns") is not None:
            opts.not_date_patterns = [str(x) for x in d["not_date_patterns"]]
        if d.get("order_tolerance_days") is not None:
            opts.order_tolerance_days = int(d["order_tolerance_days"])
        return opts


# ====================================================================================================
# 元 logproc/text.py
# 判定用テキスト（影テキスト）の作成。原文と文字位置が1対1で対応する。
# ====================================================================================================

# NFKC で数字に化けると区切りの判定ができなくなる丸数字などは残す
_KEEP_RANGES = ((0x2460, 0x24FF), (0x2776, 0x2793))
EMPTY_TOKENS = {"", "-", "－", "ー", "―", "—", "‐", "−", "n/a", "N/A", "なし", "無し", "特になし", "特に無し"}


def clean_log_text(text) -> str:
    """セル値をログ処理用の文字列にする（_x000D_ と CR を除く。他は変えない）。"""
    if text is None:
        return ""
    s = str(text).replace("_x000D_", "")
    return s.replace("\r\n", "\n").replace("\r", "\n")


class _ShadowTable(dict):
    """str.translate 用の1文字の写しの表。初めて出た文字だけ NFKC を計算して覚える（文字の種類は有限）。"""

    def __missing__(self, code: int) -> str:
        if code < 0x80 or any(lo <= code <= hi for lo, hi in _KEEP_RANGES):
            ch = chr(code)
        else:
            ch = chr(code)
            n = unicodedata.normalize("NFKC", ch)
            ch = n if len(n) == 1 else ch
        self[code] = ch
        return ch


_SHADOW_TABLE = _ShadowTable()


@lru_cache(maxsize=4096)
def _shadow_cached(text: str) -> str:
    return text.translate(_SHADOW_TABLE)


def shadow(text: str) -> str:
    """1文字ずつ NFKC をかけた写し。結果が1文字にならない文字は元のまま残し、長さを保つ。

    同じ行のセルを何度も（空判定・マスク・区切り・識別子の抜き出しで）写すので、結果を覚えておく。
    """
    if text.isascii():
        return text
    return _shadow_cached(text)


def is_empty_log(text: str) -> bool:
    return shadow(text).strip() in EMPTY_TOKENS


def nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")


def dedupe(items) -> list[str]:
    """順番を保ったまま重複を落とす（空の要素は入れない）。"""
    seen, out = set(), []
    for x in items:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


# ====================================================================================================
# 元 logproc/dates.py
# ログのエントリ先頭にある日時表現の読み取りと、年・記入順の決定。
#
# - 読み取るのはセグメント先頭（行頭・【】の中・「／」「→」の直後）の表現だけ。
#   本文中の「納期4/8」「1.5mm」「4/10に伺います」は出来事の日付にしない。
# - 年は、まず記入順（古い順／新しい順）を判定し、時間の流れに沿って前のエントリと矛盾しない年を選ぶ。
# ====================================================================================================

ERA_BASE = {"令和": 2018, "R": 2018, "平成": 1988, "H": 1988, "昭和": 1925, "S": 1925}

_ERA_RE = re.compile(
    r"(?<![A-Za-z])(令和|平成|昭和|[RHS])\s*(\d{1,2}|元)\s*"
    r"(?:年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?|[./\-]\s*(\d{1,2})\s*[./\-]\s*(\d{1,2}))(?![\d.])"
)
_YMD_RE = re.compile(
    r"(\d{4})\s*(?:年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?|([/.\-])\s*(\d{1,2})\s*\4\s*(\d{1,2}))(?![\d.])"
)
_YY_RE = re.compile(r"(\d{2})/(\d{1,2})/(\d{1,2})(?![\d/.])")
_MD_KANJI_RE = re.compile(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_MD_SLASH_RE = re.compile(r"(\d{1,2})/(\d{1,2})(?![\d/.%])")
_MMDD_RE = re.compile(r"(0[1-9]|1[0-2])([0-3]\d)(?=[ \t]|$)")
_MDOT_RE = re.compile(r"(\d{1,2})\.(\d{1,2})(?=[ \t]|$)")
_WEEKDAY_RE = re.compile(r"\s*\(\s*[月火水木金土日](?:曜日?)?\s*\)")
_RANGE_RE = re.compile(r"\s*[〜~]\s*(?:(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})|(\d{1,2})/(\d{1,2})|(\d{1,2})\s*月\s*(\d{1,2})\s*日)(?![\d/.])")
_TIME_RE = re.compile(
    r"(?:(AM|PM|午前|午後)\s*)?(?:(\d{1,2}):(\d{2})(?!\d)|(\d{1,2})時(?!間)(?:(半)|(\d{1,2})分)?)"
    r"(?:\s*[〜~\-]\s*(?:\d{1,2}:\d{2}(?!\d)|\d{1,2}時(?!間)(?:半|\d{1,2}分)?))?"
)
_SHIFT_RE = re.compile(
    r"(夜勤|日勤|[123一二三]直|午前中|午前|午後|朝一|昼過ぎ|夕方|深夜|定時後|朝|昼|夜)(?=[\s:)\]】、。,]|$)"
)
_REL_RE = re.compile(
    r"(同日|翌々日|翌日|翌朝|翌週|週明け|後日|連休明け|前日|昨日|本日|今日|今朝|明日|(\d{1,2})日後|(\d{1,2})週間後)"
    r"(?![のにはもをがで])"
)
UNRESOLVABLE = {"昨日", "本日", "今日", "今朝", "明日", "後日", "連休明け", "前日"}


@dataclass
class HeadWhen:
    """セグメント先頭の日時表現（年を決める前）。位置は影テキスト上。"""
    start: int
    end: int
    year: int | None = None
    month: int | None = None
    day: int | None = None
    full: bool = False
    to: tuple[int | None, int, int] | None = None
    time: str | None = None
    shift: str | None = None
    relative: str | None = None
    rel_days: int | None = None
    rel_weeks: int | None = None

    @property
    def has_date(self) -> bool:
        return self.month is not None


def _valid(y: int, m: int, d: int) -> bool:
    try:
        date(y, m, d)
        return True
    except ValueError:
        return False


def _match_date(sh: str, p: int, line_start: bool):
    """(year|None, month, day, full, end) を返す。"""
    m = _ERA_RE.match(sh, p)
    if m:
        n = 1 if m[2] == "元" else int(m[2])
        mo, d = (m[3], m[4]) if m[3] else (m[5], m[6])
        y = ERA_BASE[m[1]] + n
        if _valid(y, int(mo), int(d)):
            return y, int(mo), int(d), True, m.end()
    m = _YMD_RE.match(sh, p)
    if m:
        mo, d = (m[2], m[3]) if m[2] else (m[5], m[6])
        if 1900 < int(m[1]) < 2200 and _valid(int(m[1]), int(mo), int(d)):
            return int(m[1]), int(mo), int(d), True, m.end()
    m = _YY_RE.match(sh, p)
    if m and _valid(2000 + int(m[1]), int(m[2]), int(m[3])):
        return 2000 + int(m[1]), int(m[2]), int(m[3]), True, m.end()
    for rx in (_MD_KANJI_RE, _MD_SLASH_RE):
        m = rx.match(sh, p)
        if m and _valid(2024, int(m[1]), int(m[2])):
            return None, int(m[1]), int(m[2]), False, m.end()
    if line_start:
        for rx in (_MMDD_RE, _MDOT_RE):
            m = rx.match(sh, p)
            if m and _valid(2024, int(m[1]), int(m[2])):
                return None, int(m[1]), int(m[2]), False, m.end()
    return None


def _format_time(m: re.Match) -> str | None:
    ampm = m[1]
    if m[2] is not None:
        h, mi = int(m[2]), int(m[3])
    else:
        h = int(m[4])
        mi = 30 if m[5] else int(m[6] or 0)
    if ampm in ("PM", "午後") and h < 12:
        h += 12
    if h > 29 or mi > 59:
        return None
    return f"{h:02d}:{mi:02d}"


def _skip_ws(sh: str, p: int) -> int:
    while p < len(sh) and sh[p] in " \t":
        p += 1
    return p


# 利用者が設定で書いた正規表現（日付ではない書き方・区切りの目印）に渡す長さの上限。
# 行頭の目印を見るだけなので、長い行の全体を渡さない（書き方によっては時間がかかる正規表現の被害を抑える）
USER_PATTERN_WINDOW = 200


def parse_when_at(sh: str, pos: int, *, line_start: bool = True, not_date_res=(), end: int | None = None) -> HeadWhen | None:
    """影テキストの pos から始まる日時表現を読む。何もなければ None。"""
    end = len(sh) if end is None else end
    view = sh[:end]
    p = pos
    w = HeadWhen(start=pos, end=pos)
    found = False
    if not (not_date_res and any(r.match(view, p, p + USER_PATTERN_WINDOW) for r in not_date_res)):
        got = _match_date(view, p, line_start)
        if got:
            w.year, w.month, w.day, w.full, p = got
            found = True
            m = _WEEKDAY_RE.match(view, p)
            if m:
                p = m.end()
            m = _RANGE_RE.match(view, p)
            if m:
                if m[1]:
                    w.to = (int(m[1]), int(m[2]), int(m[3]))
                elif m[4]:
                    w.to = (None, int(m[4]), int(m[5]))
                else:
                    w.to = (None, int(m[6]), int(m[7]))
                if _valid(w.to[0] or 2024, w.to[1], w.to[2]):
                    p = m.end()
                else:
                    w.to = None
    for _ in range(4):
        q = _skip_ws(view, p) if found else p
        m = _TIME_RE.match(view, q) if w.time is None else None
        if m:
            t = _format_time(m)
            if t is None:
                break
            w.time = t
            p = m.end()
            found = True
            continue
        m = _SHIFT_RE.match(view, q) if w.shift is None else None
        if m:
            w.shift = m[1]
            p = m.end()
            found = True
            continue
        m = _REL_RE.match(view, q) if (not w.has_date and w.relative is None and w.time is None) else None
        if m:
            w.relative = m[1]
            if m[2]:
                w.rel_days = int(m[2])
            if m[3]:
                w.rel_weeks = int(m[3])
            if w.relative == "翌朝":
                w.shift = w.shift or "朝"
            p = m.end()
            found = True
            continue
        break
    if not found:
        return None
    w.end = p
    return w


def head_kind(w: HeadWhen | None) -> str | None:
    if w is None:
        return None
    if w.has_date:
        return "date"
    if w.relative:
        return "relative"
    if w.time:
        return "time"
    return "shift"


# ---- 年と記入順の決定 ----

def _mk(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d)
    except ValueError:
        return None


def _pick_year(m: int, d: int, prev: date | None, base: date | None, tol: int) -> tuple[date | None, bool]:
    """時間の流れに沿って年を選ぶ。戻り値: (日付, 前後の逆行があったか)"""
    if prev:
        c = _mk(prev.year, m, d)
        if c and c >= prev:
            return c, False
        if c and (prev - c).days <= tol:
            return c, True
        for y in (prev.year + 1, prev.year + 2, prev.year + 4):
            c = _mk(y, m, d)
            if c:
                return c, False
        return None, False
    if base:
        cands = [c for c in (_mk(base.year + k, m, d) for k in (-1, 0, 1)) if c]
        if cands:
            return min(cands, key=lambda c: abs((c - base).days)), False
    return None, False


def _step(a: HeadWhen, b: HeadWhen) -> int:
    """a→b が進む向きなら 1、戻る向きなら -1、同じ日なら 0。年がない日付は1年の円周上で近い向き。"""
    if a.full and b.full:
        da, db = date(a.year, a.month, a.day), date(b.year, b.month, b.day)
        return (db > da) - (db < da)
    da = date(2024, a.month, a.day).timetuple().tm_yday
    db = date(2024, b.month, b.day).timetuple().tm_yday
    f = (db - da) % 366
    if f == 0:
        return 0
    return 1 if f < 183 else -1


def detect_order(heads: list[HeadWhen | None]) -> tuple[str, bool]:
    """明示された日付の並びから記入順を判定する。戻り値: (asc/desc/unknown, 向きが混在するか)"""
    dated = [h for h in heads if h is not None and h.has_date]
    fwd = bwd = 0
    for a, b in zip(dated, dated[1:]):
        s = _step(a, b)
        fwd += s > 0
        bwd += s < 0
    if fwd == 0 and bwd == 0:
        # 明示の日付で決まらなくても「翌日」「3日後」などが続けば古い順
        if len(dated) >= 1 and any(h is not None and h.relative and h.relative not in UNRESOLVABLE for h in heads):
            return "asc", False
        return "unknown", False
    return ("desc" if bwd > fwd else "asc"), (fwd > 0 and bwd > 0)


def first_full_date(heads: list[HeadWhen | None]) -> date | None:
    for h in heads:
        if h is not None and h.full:
            return date(h.year, h.month, h.day)
    return None


def _relative_range(h: HeadWhen, ref: date) -> tuple[date, date | None, bool]:
    """相対表現を (開始日, 終了日, 推定か) にする。"""
    r = h.relative
    if r == "同日":
        return ref, None, False
    if r in ("翌日", "翌朝"):
        return ref + timedelta(days=1), None, True
    if r == "翌々日":
        return ref + timedelta(days=2), None, True
    if h.rel_days is not None:
        return ref + timedelta(days=h.rel_days), None, True
    if h.rel_weeks is not None:
        return ref + timedelta(weeks=h.rel_weeks), None, True
    if r == "翌週":
        monday = ref + timedelta(days=7 - ref.weekday())
        return monday, monday + timedelta(days=6), True
    if r == "週明け":
        return ref + timedelta(days=7 - ref.weekday()), None, True
    return ref, None, True


def resolve_whens(
    heads: list[HeadWhen | None],
    texts: list[str],
    base_date: date | None,
    order_option: str = "auto",
    tolerance_days: int = 60,
    skip: set[int] | None = None,
) -> tuple[list[WhenInfo | None], str, list[str]]:
    """セグメントごとの先頭日時から WhenInfo を決める。

    heads/texts はセグメントの原文順。skip の位置（見出し型など）は None のまま返す。
    戻り値: (WhenInfo のリスト, 記入順, 警告)
    """
    skip = skip or set()
    n = len(heads)
    active = [h if i not in skip else None for i, h in enumerate(heads)]
    warnings: list[str] = []
    detected, mixed = detect_order(active)
    order = order_option if order_option in ("asc", "desc") else detected
    if mixed and order_option not in ("asc", "desc"):
        warnings.append("記入順が一部前後している（古い順と新しい順が混在）")
    base = base_date or first_full_date(active)

    resolved: dict[int, WhenInfo] = {}
    chron = list(range(n))
    if order == "desc":
        chron.reverse()
    prev: date | None = None
    year_unknown = False
    for i in chron:
        h = active[i]
        if h is None:
            continue
        text = texts[i]
        if h.has_date:
            if h.full:
                d, anomaly, how = date(h.year, h.month, h.day), False, "explicit"
            else:
                d, anomaly = _pick_year(h.month, h.day, prev, base, tolerance_days)
                how = "year_inferred"
            if d is None:
                year_unknown = True
                resolved[i] = WhenInfo(text, None, h.time, None, h.shift, False,
                                       f"原文「{text}」の年を決められない（発生日が空で、セル内に年を含む日付がない）", "year_unknown")
                continue
            if anomaly:
                warnings.append(f"記入順が一部前後している（「{text}」）")
            d_to = None
            if h.to:
                y_to = h.to[0] or d.year
                d_to = _mk(y_to, h.to[1], h.to[2])
                if d_to and d_to < d and not h.to[0]:
                    d_to = _mk(y_to + 1, h.to[1], h.to[2])
            resolved[i] = WhenInfo(text, d.isoformat(), h.time, d_to.isoformat() if d_to else None, h.shift, False, "", how)
            prev = d
        elif h.relative:
            if h.relative in UNRESOLVABLE:
                resolved[i] = WhenInfo(text, None, h.time, None, h.shift, True,
                                       f"原文「{h.relative}」は基準の日が分からないため日付にできない", "unresolved")
                continue
            ref = prev or base
            if ref is None:
                resolved[i] = WhenInfo(text, None, h.time, None, h.shift, True,
                                       f"原文「{h.relative}」の基準になる日付がない", "unresolved")
                continue
            d, d_to, est = _relative_range(h, ref)
            basis = "直前の記録" if prev else "発生日"
            note = f"原文「{h.relative}」から推定した（基準は{basis}の{ref.isoformat()}）" if est else ""
            resolved[i] = WhenInfo(text, d.isoformat(), h.time, d_to.isoformat() if d_to else None, h.shift, est, note, "relative")
            prev = d

    out: list[WhenInfo | None] = []
    last: WhenInfo | None = None
    for i in range(n):
        if i in skip:
            out.append(None)
            continue
        info = resolved.get(i)
        if info is None:
            h = heads[i]
            time = h.time if h else None
            shift = h.shift if h else None
            text = texts[i] if h else ""
            if last is not None:
                info = WhenInfo(text, last.date, time, last.date_to, shift, last.estimated,
                                "日付の記載がなく、直前の記録の日付を使った", "inherited")
            elif base is not None:
                info = WhenInfo(text, base.isoformat(), time, None, shift, True,
                                f"日付の記載がなく、発生日（{base.isoformat()}）を仮に使った", "base")
            else:
                info = WhenInfo(text, None, time, None, shift, False, "日付の記載がない", "none")
        if info.date:
            last = info
        out.append(info)
    if year_unknown:
        warnings.append("年を決められない日付がある（発生日が空で、セル内に年を含む日付がない）")
    if any(h is not None and h.relative in UNRESOLVABLE for h in active):
        warnings.append("「昨日」「本日」など基準の日が分からない表現は日付にしていない")
    return out, order, warnings


# ====================================================================================================
# 元 logproc/extract.py
# 識別子（型番・アラームコード・ロット）、数量・回数、予定句の抜き出し。
# ====================================================================================================

_ID_RE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z0-9](?:[A-Za-z0-9._\-]*[A-Za-z0-9])?")
_UNITS = sorted(
    "個 本 枚 台 回 件 箇所 ヶ所 か所 カ所 セット 式 袋 缶 巻 m mm cm μm MPa kPa Pa ℃ °C V kV mA A kW W L mL ml "
    "分 秒 時間 日間 週間 か月 ヶ月 カ月 kg g t MΩ Ω M % rpm Hz P 万円 円 人 Torr mTorr sccm slm degC".split(),
    key=len, reverse=True,
)
_UNIT_ALT = "|".join(re.escape(u) for u in _UNITS)
# 英字で終わる単位は後ろに英字が続かないこと（「5mm」と「5mmHg」などの区別）。和文の単位は制限しない
_QTY_UNIT = "(?:" + "|".join(re.escape(u) + ("(?![A-Za-z])" if u[-1].isascii() and u[-1].isalpha() else "") for u in _UNITS) + ")"
_ASCII_UNIT_RE = re.compile(rf"\d+(?:\.\d+)?(?:E[-+]?\d+)?(?:{_UNIT_ALT})", re.IGNORECASE)
# ガス・化学式（N2パージ、NF3 など）は型番ではない。E1・P2 のようなアラームコードと区別できないので、決まった一覧だけ除く
_FORMULAS = frozenset("N2 O2 H2 NF3 CF4 SF6 CO2 H2O NH3 Cl2 C4F8 C2F6 CHF3 CH4 SiH4 WF6 BCl3 N2O O3 H2O2".split())
# 識別子にしない形（日付・時刻・和暦・番号・マスク）
_NOT_ID_RES = [
    re.compile(r"[RHS]\d{1,2}(?:\.\d{1,2}){2}"),
    re.compile(r"(?:AM|PM)\d{1,2}", re.IGNORECASE),
    re.compile(r"No\.?\d+", re.IGNORECASE),
    re.compile(r"(?:NG|OK)\d+", re.IGNORECASE),
    re.compile(r"\d{4}-\d{1,2}-\d{1,2}T.*"),
    re.compile(r".*(?:XXX|xxx).*"),
]
_QTY_RE = re.compile(
    rf"(?<![A-Za-z])[×xX]\s*\d+(?:\.\d+)?"
    rf"|¥\s*\d[\d,]*"
    rf"|(?<![\d.\-,±])\d+(?:[.,]\d+)*(?:E[-+]?\d+)?(?:\s*[〜~]\s*\d+(?:\.\d+)?)?\s*{_QTY_UNIT}"
)
_HAS_ALPHA_RE = re.compile(r"[A-Za-z]")
_HAS_DIGIT_RE = re.compile(r"\d")
_PLAN_KEY_RE = re.compile(r"納期|予定|目処|目途|見込み|までに")
_CLAUSE_SEP_RE = re.compile(r"[、。,，()（）\s→・「」【】]+")


def identifier_spans(text: str) -> list[tuple[int, int, str]]:
    """英字と数字を両方含む語の位置。戻り値の語は NFKC 後の表記。"""
    return list(_identifier_spans(text))


@lru_cache(maxsize=4096)
def _identifier_spans(text: str) -> tuple[tuple[int, int, str], ...]:
    # 同じ本文を識別子・数量の抜き出しで続けて調べるので、結果を覚えておく（変更されないよう tuple で持つ）
    sh = shadow(text)
    out = []
    for m in _ID_RE.finditer(sh):
        tok = m.group(0)
        if not (_HAS_ALPHA_RE.search(tok) and _HAS_DIGIT_RE.search(tok)):
            continue
        if tok in _FORMULAS or _ASCII_UNIT_RE.fullmatch(tok) or any(r.fullmatch(tok) for r in _NOT_ID_RES):
            continue
        out.append((m.start(), m.end(), tok))
    return tuple(out)


def extract_identifiers(text: str) -> list[str]:
    return dedupe(tok for _, _, tok in _identifier_spans(text))


def extract_quantities(text: str) -> list[str]:
    """数値＋単位、×1、2回、金額。原文の表記（NFKC 後）で返す。"""
    sh = shadow(text)
    id_spans = _identifier_spans(text)
    out = []
    for m in _QTY_RE.finditer(sh):
        if any(s <= m.start() and m.end() <= e for s, e, _ in id_spans):
            continue
        out.append(m.group(0).strip())
    return dedupe(out)


def extract_plans(text: str) -> list[str]:
    """「納期1週間」「6月末目処」「交換予定」などの予定句（原文の表記）。"""
    # 予定の語は区切り記号を含まないので、全体に無ければどの句にも無い（句ごとに写しを作らない）
    if not text or not _PLAN_KEY_RE.search(shadow(text)):
        return []
    out = []
    for piece in _CLAUSE_SEP_RE.split(text):
        if piece and _PLAN_KEY_RE.search(shadow(piece)):
            out.append(piece)
    return dedupe(out)


# ====================================================================================================
# 元 logproc/people.py
# 記入者の判定。人物一覧（別名・イニシャル）と担当列の名前で照合する。
#
# 人物一覧はアプリ内だけで使い、AIには送らない。
# 位置の形（日付直後の「田中：」、行末の「（田中）」、「田中→佐藤」など）に合うときだけ人名として扱う。
# ====================================================================================================

DEFAULT_GROUPS = [
    "保全G", "保全課", "保全", "製造課", "製造", "品証", "品質保証", "施設課", "技術", "生技",
    "メーカーFE", "メーカー", "業者", "夜勤者", "班長", "課長",
]

# 辞書にない名前でも「日付 名前 本文」の形で人名とみなす一般的な姓
COMMON_SURNAMES = set("""
佐藤 鈴木 高橋 田中 伊藤 渡辺 渡部 渡邊 山本 中村 小林 加藤 吉田 山田 佐々木 山口 松本 井上 木村 林 斎藤 斉藤
清水 山崎 森 池田 橋本 阿部 石川 山下 中島 石井 小川 前田 岡田 長谷川 藤田 後藤 近藤 村上 遠藤 青木 坂本
福田 太田 西村 藤井 金子 岡本 藤原 中野 三浦 原田 中川 松田 竹内 小野 田村 中山 和田 石田 森田 上田 原
柴田 酒井 工藤 横山 宮崎 宮本 内田 高木 安藤 島田 谷口 大野 高田 丸山 今井 河野 藤本 村田 武田 上野 杉山
増田 小山 大塚 平野 菅原 久保 松井 千葉 岩崎 桜井 木下 野口 松尾 菊地 野村 新井 小西 大西 西田 北村 石原
永井 荒木 本田 久保田 中西 浅野 服部 市川 飯田 片山 小島 水野 岡崎 西川 伊東 五十嵐 松下 吉川 山内 北川
""".split())

# 「〇〇：」の〇〇が人名ではない語
_COLON_STOP = set("""
原因 現象 対応 処置 暫定 恒久 結果 備考 内容 状況 理由 対策 再発防止 補足 注意 注記 予定 回答 連絡 報告 確認
作業 停止 復旧 部品 調査 点検 判断 結論 経過 方針 課題 問題 要因 件名 場所 時間 日時 期間 費用 金額 担当
記入者 対応者 設備 工程 品番 型番 数量 状態 目的 依頼 指示 保留 完了 未 済 注 訂正 追記 水平展開 異常 不具合
""".split())

_KANJI = r"[一-鿿々ヶ]"
TOKEN_RE = re.compile(rf"(?:{_KANJI}{{1,4}}(?:\({_KANJI}{{1,2}}\))?|[A-Za-z][A-Za-z.]{{0,14}}[A-Za-z.]?)")
_INITIALS_RE = re.compile(r"[A-Z]\.[A-Z]\.?")
_LATIN_NAME_RE = re.compile(r"[A-Z][a-z]{2,}")


def key(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text or "")).casefold()


@dataclass
class Person:
    name: str
    aliases: list[str] = field(default_factory=list)
    org: str = ""


class PeopleIndex:
    """登録人物・担当列の名前・部署名の索引。"""

    def __init__(self, registered: Iterable[dict | Person] | None = None, column_names: Iterable[str] = (),
                 groups: Iterable[str] | None = None):
        self.persons: list[Person] = []
        self._exact: dict[str, list[Person]] = {}     # 氏名・別名
        self._surname: dict[str, list[Person]] = {}   # 姓
        for item in registered or []:
            p = item if isinstance(item, Person) else Person(
                str(item.get("name", "")).strip(), [str(a) for a in item.get("aliases", []) or []], str(item.get("org", "") or ""))
            if p.name:
                self._add(p)
        for raw in column_names or []:
            for part in re.split(r"[/／、,・]", str(raw or "")):
                part = part.strip()
                if not part or len(part) > 12:
                    continue
                k = key(part)
                if k in self._exact or k in self._surname:
                    continue
                self._add(Person(part))
        self.groups = list(groups) if groups is not None else list(DEFAULT_GROUPS)
        self._group_keys = {key(g): g for g in self.groups}
        # 本文先頭で区切りなしに照合する表記（長い順）
        forms = {unicodedata.normalize("NFKC", f) for f in list(self._exact_forms()) + self.groups}
        self._forms = sorted((f for f in forms if len(f) >= 2), key=len, reverse=True)
        # known_at 用に先頭の文字で引けるようにする（各リストの中は長い順のまま）
        self._forms_by_head: dict[str, list[str]] = {}
        for f in self._forms:
            self._forms_by_head.setdefault(f[0], []).append(f)
        # 部署名の照合用（長い順・NFKC）。セグメントごとに並べ直さない
        self._group_forms = [unicodedata.normalize("NFKC", g) for g in sorted(self.groups, key=len, reverse=True)]

    def _add(self, p: Person) -> None:
        self.persons.append(p)
        for form in [p.name, *p.aliases]:
            self._exact.setdefault(key(form), []).append(p)
        parts = unicodedata.normalize("NFKC", p.name).split()
        if len(parts) >= 2:
            self._surname.setdefault(key(parts[0]), []).append(p)

    def _exact_forms(self):
        for p in self.persons:
            yield p.name
            yield from p.aliases
            parts = p.name.split()
            if len(parts) >= 2:
                yield parts[0]
                yield "".join(parts)

    # ---- 照合 ----
    def is_group(self, token: str) -> bool:
        return key(token) in self._group_keys

    def is_known(self, token: str) -> bool:
        k = key(token)
        return k in self._exact or k in self._surname or k in self._group_keys

    def group_at(self, sh: str, p: int, end: int) -> str | None:
        """p から始まる部署名（後ろが空白・「:」・行末のもの）。"""
        for gn in self._group_forms:
            q = p + len(gn)
            if q <= end and sh.startswith(gn, p) and (q == end or sh[q] in " \t:"):
                return gn
        return None

    def known_at(self, sh: str, p: int, end: int) -> str | None:
        """p から区切りなしで始まる登録済みの名前（「4/3佐藤エンコーダ…」用）。"""
        if p >= end:
            return None
        for f in self._forms_by_head.get(sh[p], ()):
            if p + len(f) <= end and sh.startswith(f, p) and not self.is_group(f):
                return f
        return None

    def names(self) -> list[str]:
        """マスク用の名前一覧（氏名・姓・別名）。"""
        out = set()
        for f in self._exact_forms():
            if len(f) >= 2:
                out.add(f)
        return sorted(out, key=len, reverse=True)

    def accept_colon(self, token: str) -> bool:
        if token in _COLON_STOP:
            return False
        if self.is_known(token) or _INITIALS_RE.fullmatch(token) or _LATIN_NAME_RE.fullmatch(token):
            return True
        base = re.sub(r"\(.*\)$", "", token)
        return bool(re.fullmatch(rf"{_KANJI}{{2,4}}", base)) and base not in _COLON_STOP

    def accept_weak(self, token: str) -> bool:
        if self.is_known(token) or _INITIALS_RE.fullmatch(token):
            return True
        return re.sub(r"\(.*\)$", "", token) in COMMON_SURNAMES

    def resolve(self, raw: str) -> AuthorInfo:
        """表記から AuthorInfo を作る（未登録でも返す）。"""
        k = key(raw)
        if k in self._group_keys:
            return AuthorInfo(raw, self._group_keys[k], False, "")
        cands = self._exact.get(k)
        if cands:
            uniq = _uniq(cands)
            if len(uniq) == 1:
                p = uniq[0]
                note = f"人物一覧の別名「{raw}」から特定した" if key(p.name) != k else ""
                return AuthorInfo(raw, p.name, False, note)
            return AuthorInfo(raw, raw, False, _ambiguous_note(uniq))
        cands = self._surname.get(k)
        if cands:
            uniq = _uniq(cands)
            if len(uniq) == 1:
                return AuthorInfo(raw, uniq[0].name, False, "")
            return AuthorInfo(raw, raw, False, _ambiguous_note(uniq))
        if _INITIALS_RE.fullmatch(unicodedata.normalize("NFKC", raw)):
            return AuthorInfo(raw, None, False, "人物一覧にないイニシャルのため、誰か特定していない")
        return AuthorInfo(raw, raw, False, "")


def _uniq(persons: list[Person]) -> list[Person]:
    seen, out = set(), []
    for p in persons:
        if id(p) not in seen:
            seen.add(id(p))
            out.append(p)
    return out


def _ambiguous_note(persons: list[Person]) -> str:
    names = "／".join(p.name for p in persons)
    return f"人物一覧に{len(persons)}人（{names}）いるため、どちらか特定していない"


def _skip(sh: str, p: int, end: int, chars: str) -> int:
    while p < end and sh[p] in chars:
        p += 1
    return p


def _token_at(sh: str, p: int, end: int) -> str | None:
    m = TOKEN_RE.match(sh[:end], p)
    return m.group(0) if m else None


def detect_head_author(sh: str, p: int, end: int, index: PeopleIndex) -> tuple[AuthorInfo | None, int]:
    """本文先頭（日時の直後）の記入者。戻り値: (記入者, 本文の開始位置)"""
    p = _skip(sh, p, end, " \t:")
    g = index.group_at(sh, p, end)
    if g:
        q = p + len(g)
        q2 = _skip(sh, q, end, " \t")
        t = _token_at(sh, q2, end) if q2 > q else None
        if t and index.accept_weak(t) and not index.is_group(t):
            after = q2 + len(t)
            if after == end or sh[after] in " \t:":
                info = index.resolve(t)
                info.raw = sh[p:after]
                return info, _skip(sh, after, end, " \t:")
        return index.resolve(g), _skip(sh, q, end, " \t:")
    t = _token_at(sh, p, end)
    if t:
        q = p + len(t)
        m = re.compile(r"\s*→\s*").match(sh[:end], q)
        if m:
            t2 = _token_at(sh, m.end(), end)
            if t2 and index.accept_weak(t) and index.accept_weak(t2):
                after = m.end() + len(t2)
                info = index.resolve(t)
                info.raw = sh[p:after]
                info.note = (info.note + "。" if info.note else "") + f"原文は「{sh[p:after]}」（引継ぎまたは連絡の相手は{t2}）"
                return info, _skip(sh, after, end, " \t:")
        r = _skip(sh, q, end, " \t")
        if r < end and sh[r] == ":" and index.accept_colon(t):
            return index.resolve(t), _skip(sh, r + 1, end, " \t")
        if (r > q or r == end) and index.accept_weak(t):
            return index.resolve(t), r
    k = index.known_at(sh, p, end)
    if k:
        return index.resolve(k), p + len(k)
    return None, p


_TAIL_PAREN_RE = re.compile(rf"(?:(\d{{1,2}}/\d{{1,2}})\s*)?({TOKEN_RE.pattern}|[^\s()]{{2,8}})?")
_TAIL_BARE_RE = re.compile(rf"(?<=[\s)）])({_KANJI}{{1,4}})$")


def detect_tail_author(sh: str, start: int, end: int, index: PeopleIndex) -> tuple[AuthorInfo | None, int, tuple[int, int] | None]:
    """末尾の「（田中）」「(4/3 西村)」「…）小西」。戻り値: (記入者, 本文の終了位置, 括弧内の日付の位置)"""
    e = end
    while e > start and sh[e - 1] in " \t。、/→":
        e -= 1
    if e - start >= 3 and sh[e - 1] == ")":
        o = sh.rfind("(", start, e - 1)
        if o > start:
            inner_start = _skip(sh, o + 1, e - 1, " \t")
            inner = sh[inner_start:e - 1].rstrip()
            m = _TAIL_PAREN_RE.fullmatch(inner)
            if m and (m.group(1) or m.group(2)):
                tok = m.group(2)
                if tok and not (index.accept_weak(tok) or index.is_group(tok)):
                    return None, end, None
                info = index.resolve(tok) if tok else None
                dspan = (inner_start + m.start(1), inner_start + m.end(1)) if m.group(1) else None
                return info, o, dspan
    m = _TAIL_BARE_RE.search(sh[start:e])
    if m and index.accept_weak(m.group(1)) and start + m.start(1) > start:
        return index.resolve(m.group(1)), start + m.start(1), None
    return None, end, None


def inherit_authors(authors: list[AuthorInfo | None], skip: set[int] | None = None) -> list[AuthorInfo | None]:
    """記入者の書かれていないセグメントに直前の記入者を（推定）で付ける。"""
    skip = skip or set()
    out: list[AuthorInfo | None] = []
    last: AuthorInfo | None = None
    for i, a in enumerate(authors):
        if i in skip:
            out.append(a)
            continue
        if a is None and last is not None:
            label = last.name or last.raw
            a = AuthorInfo("", label, True, f"原文になく、直前の{label}を引き継いだ")
        if a is not None:
            last = a
        out.append(a)
    return out


# ====================================================================================================
# 元 logproc/mask.py
# 個人情報のマスク（電話番号・メールアドレス、任意で人名・金額）。
#
# AI に送る前と md 出力の両方で使う。判定は1文字ずつ NFKC した写しで行い、位置は原文と同じ。
# ====================================================================================================

REPLACEMENTS = {"phone": "［電話番号］", "email": "［メール］", "person": "［人名］", "amount": "［金額］"}
_RULE_ALIASES = {
    "phone": "phone", "電話番号": "phone", "tel": "phone",
    "email": "email", "mail": "email", "メールアドレス": "email", "メール": "email",
    "person": "person", "人名": "person", "name": "person",
    "amount": "amount", "金額": "amount", "money": "amount",
}
DEFAULT_RULES = ["phone", "email"]

_SEP = r"[\-‐‑–—―−ー \t]"
_PHONE_RE = re.compile(
    r"(?<![\dA-Za-z.\-/])(?:"
    rf"\+81{_SEP}?\(?0?\d{{1,4}}\)?{_SEP}?\d{{1,4}}{_SEP}?\d{{4}}"
    rf"|\(0\d{{1,4}}\){_SEP}?\d{{1,4}}{_SEP}?\d{{4}}"
    rf"|0(?:120|800){_SEP}?\d{{2,3}}{_SEP}?\d{{3,4}}"
    rf"|0\d{{1,4}}{_SEP}\d{{1,4}}{_SEP}\d{{3,4}}"
    r"|0\d{9,10}"
    r")(?!\d)"
)
_EXT_RE = re.compile(r"(?:内線|ext\.?|EXT\.?|Ext\.?)\s*[:：]?\s*\d{2,6}(?!\d)")
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
_AMOUNT_RE = re.compile(r"¥\s*\d[\d,]*(?:\.\d+)?|(?<![\d.,])\d[\d,]*(?:\.\d+)?\s*(?:万円|円)")


@dataclass
class MaskSpan:
    start: int          # 原文上の位置
    end: int
    kind: str           # phone/email/person/amount
    text: str           # 元の文字列
    replacement: str


def normalize_rules(rules: Iterable[str] | None) -> list[str]:
    out = []
    for r in (DEFAULT_RULES if rules is None else rules):
        k = _RULE_ALIASES.get(str(r).strip().lower(), _RULE_ALIASES.get(str(r).strip()))
        if k and k not in out:
            out.append(k)
    return out


def find_mask_spans(text: str, rules: Iterable[str] | None = None, names: Iterable[str] = ()) -> list[MaskSpan]:
    text = text or ""
    sh = shadow(text)
    kinds = normalize_rules(rules)
    found: list[MaskSpan] = []

    def add(kind: str, s: int, e: int):
        found.append(MaskSpan(s, e, kind, text[s:e], REPLACEMENTS[kind]))

    if "email" in kinds:
        for m in _EMAIL_RE.finditer(sh):
            add("email", m.start(), m.end())
    if "phone" in kinds:
        for rx in (_PHONE_RE, _EXT_RE):
            for m in rx.finditer(sh):
                add("phone", m.start(), m.end())
    if "amount" in kinds:
        for m in _AMOUNT_RE.finditer(sh):
            add("amount", m.start(), m.end())
    if "person" in kinds:
        forms = sorted({unicodedata.normalize("NFKC", n).strip() for n in names if n and len(n.strip()) >= 2},
                       key=len, reverse=True)
        for form in forms:
            for variant in {form, form.replace(" ", "")}:
                start = 0
                while (i := sh.find(variant, start)) >= 0:
                    add("person", i, i + len(variant))
                    start = i + len(variant)
    # 重なりは先に始まる方・長い方を採る
    found.sort(key=lambda s: (s.start, -(s.end - s.start)))
    out: list[MaskSpan] = []
    for s in found:
        if out and s.start < out[-1].end:
            continue
        out.append(s)
    return out


def mask_text(text: str, rules: Iterable[str] | None = None, names: Iterable[str] = ()) -> tuple[str, list[MaskSpan]]:
    """マスクした文字列と、置き換えた箇所（原文上の位置）を返す。rules は phone/email/person/amount（日本語名も可）。"""
    text = text or ""
    spans = find_mask_spans(text, rules, names)
    parts, pos = [], 0
    for s in spans:
        parts.append(text[pos:s.start])
        parts.append(s.replacement)
        pos = s.end
    parts.append(text[pos:])
    return "".join(parts), spans


# ====================================================================================================
# 元 logproc/glossary.py
# 用語集による言い換え（様子見→経過観察 など）。
#
# - 長い語から順に照合する（1回の走査。置き換えた結果を再度置き換えない）。
# - 識別子（FE-100 など）とマスク記号の範囲は置き換えない。
# - ascii_boundary: 前後が英数字・「-」「_」なら置き換えない。英字だけの語は既定でオン。
# ====================================================================================================

_MASK_TOKEN_RE = re.compile(r"[［\[][^］\]]{1,10}[］\]]")
_ASCII_WORD_CHARS = re.compile(r"[A-Za-z0-9_\-]")


@dataclass
class GlossaryHit:
    start: int
    end: int
    term: str
    to: str


def _entries(glossary: dict) -> list[tuple[str, str, bool]]:
    out = []
    for term, spec in (glossary or {}).items():
        term = str(term)
        if not term:
            continue
        if isinstance(spec, dict):
            to = str(spec.get("to", ""))
            boundary = spec.get("ascii_boundary")
        else:
            to, boundary = str(spec), None
        term_sh = shadow(term)
        if boundary is None:
            boundary = bool(re.fullmatch(r"[A-Za-z0-9_\-]+", term_sh))
        out.append((term_sh, to, bool(boundary)))
    out.sort(key=lambda e: len(e[0]), reverse=True)
    return out


def _span_pairs(protected_spans) -> list[tuple[int, int]]:
    pairs = []
    for sp in protected_spans or []:
        if isinstance(sp, (tuple, list)):
            pairs.append((int(sp[0]), int(sp[1])))
        else:
            pairs.append((int(sp.start), int(sp.end)))
    return pairs


def glossary_hits(text: str, glossary: dict, protected_spans=None) -> list[GlossaryHit]:
    """置き換える箇所の一覧（分割プレビューでの表示用）。protected_spans が None なら識別子を自動で保護する。"""
    text = text or ""
    entries = _entries(glossary)
    if not entries or not text:
        return []
    sh = shadow(text)
    if protected_spans is None:
        protected = [(s, e) for s, e, _ in identifier_spans(text)]
    else:
        protected = _span_pairs(protected_spans)
    protected += [(m.start(), m.end()) for m in _MASK_TOKEN_RE.finditer(sh)]
    # すでに言い換え後の表記になっている箇所（「短時間停止（チョコ停）」）は触らない
    for term_sh, to, _ in entries:
        to_sh = shadow(to)
        if to_sh and term_sh in to_sh:
            start = 0
            while (i := sh.find(to_sh, start)) >= 0:
                protected.append((i, i + len(to_sh)))
                start = i + len(to_sh)

    by_first: dict[str, list[tuple[str, str, bool]]] = {}
    for e in entries:
        by_first.setdefault(e[0][0], []).append(e)
    hits: list[GlossaryHit] = []
    i = 0
    while i < len(sh):
        cands = by_first.get(sh[i])
        matched = None
        for term_sh, to, boundary in cands or []:
            j = i + len(term_sh)
            if not sh.startswith(term_sh, i):
                continue
            if any(s < j and i < e for s, e in protected):
                continue
            if boundary and ((i > 0 and _ASCII_WORD_CHARS.match(sh[i - 1])) or (j < len(sh) and _ASCII_WORD_CHARS.match(sh[j]))):
                continue
            matched = GlossaryHit(i, j, text[i:j], to)
            break
        if matched:
            hits.append(matched)
            i = matched.end
        else:
            i += 1
    return hits


def apply_glossary(text: str, glossary: dict, protected_spans=None) -> str:
    """用語集を適用した文字列を返す。"""
    text = text or ""
    parts, pos = [], 0
    for h in glossary_hits(text, glossary, protected_spans):
        parts.append(text[pos:h.start])
        parts.append(h.to)
        pos = h.end
    parts.append(text[pos:])
    return "".join(parts)


# ====================================================================================================
# 元 logproc/render.py
# ルール処理結果から「対応の時系列」の行を作る（keep: 本文は原文＋用語集）。
# ====================================================================================================

def format_when(when: WhenInfo | None) -> str:
    """「2024-04-01 10:00」「2024-04-08〜2024-04-14（原文「翌週」、推定）」など。"""
    if when is None:
        return "日付不明"
    if not when.date:
        if when.how == "year_unknown" and when.text:
            return f"{when.text}（年不明）"
        return f"日付不明（原文「{when.text}」）" if when.text else "日付不明"
    s = when.date
    if when.date_to and when.date_to != when.date:
        s += f"〜{when.date_to}"
    if when.time:
        s += f" {when.time}"
    if when.shift:
        s += f" {when.shift}"
    if when.estimated:
        rel = _relative_word(when)
        s += f"（原文「{rel}」、推定）" if rel else "（推定）"
    return s


def _relative_word(when: WhenInfo) -> str:
    m = re.search(r"原文「([^」]+)」", when.note or "")
    return m.group(1) if (m and when.how == "relative") else ""


def format_author(author: AuthorInfo | None) -> str:
    if author is None:
        return ""
    label = author.name or author.raw
    if not label:
        return ""
    if author.name and author.raw and key(author.raw) != key(author.name) and key(author.raw) not in key(author.name) \
            and not author.estimated:
        label = f"{author.name}（{author.raw}）"
    if author.estimated:
        label += "（推定）"
    return label


def _one_line(text: str) -> str:
    return re.sub(r"\s*\n\s*", " ", text or "").strip()


def timeline_order(parse: LogParse) -> list[Segment]:
    """描画順。新しい順の記入は、日付を持つ記録ごとの塊で古い順に並べ替える（塊の中の順は保つ）。"""
    if parse.order != "desc":
        return list(parse.segments)
    blocks: list[list[Segment]] = []
    for seg in parse.segments:
        own = seg.when is not None and seg.when.how not in ("inherited", "base", "none")
        if own or not blocks:
            blocks.append([seg])
        else:
            blocks[-1].append(seg)
    return [s for block in reversed(blocks) for s in block]


def render_timeline(parse: LogParse, entity_label: str, types: dict[str, list[str]] | None = None,
                    glossary: dict | None = None) -> list[str]:
    """時系列の行。例: "1. 2024-04-01 10:00［連絡・初動］田中｜搬送ロボット2号機: 本文"

    types はセグメントID→種別（AI の結果。なければ［］を付けない）。
    見出し型セルは「現象: 起動しない」の形（番号なし）で返す。
    """
    types = types or {}
    lines: list[str] = []
    if parse.kind == "empty":
        return lines
    if parse.kind == "header_cell":
        for seg in parse.segments:
            body = _one_line(apply_glossary(seg.body, glossary) if glossary else seg.body)
            lines.append(f"{seg.label}: {body}" if seg.label else body)
        return lines
    for n, seg in enumerate(timeline_order(parse), start=1):
        body = seg.body or seg.raw
        if glossary:
            body = apply_glossary(body, glossary)
        body = _one_line(body)
        head = format_when(seg.when)
        seg_types = [t for t in types.get(seg.id, []) if t]
        head += f"［{'・'.join(seg_types)}］" if seg_types else " "
        who = format_author(seg.author)
        entity = (entity_label or "").strip()
        if who and entity:
            mid = f"{who}｜{entity}"
        else:
            mid = who or entity
        lines.append(f"{n}. {head}{mid}: {body}" if mid else f"{n}. {head.rstrip()}: {body}")
    return lines


def review_notes(parse: LogParse) -> list[str]:
    """要確認に出す文（推定した日付・引き継いだ記入者・特定できない人物・警告）。番号は render_timeline と同じ。"""
    notes: list[str] = []
    if parse.kind in ("empty", "header_cell"):
        return list(parse.warnings)
    for n, seg in enumerate(timeline_order(parse), start=1):
        w = seg.when
        if w is not None and w.note and (w.estimated or w.how in ("unresolved", "base")):
            if w.how != "inherited":
                notes.append(f"{n}の日付は{w.note}。")
        a = seg.author
        if a is not None and a.note:
            if a.estimated:
                notes.append(f"{n}の記入者は{a.note}。")
            elif "特定していない" in a.note:
                notes.append(f"「{a.raw}」は{a.note}。")
    notes.extend(parse.warnings)
    seen, out = set(), []
    for x in notes:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


# ====================================================================================================
# 元 logproc/segment.py
# 「経過の記録」（date_log）の分割と、parse_log（分割・日時・記入者・抜き出しの一括処理）。
#
# 区切り: 行頭の日付・時刻・相対日・勤務帯、【】、・などの箇条書き、①、「／」「→」直後の日付、
# 全角空白の後の「10:15：」。※行・→行・目印のない行は直前につなぐ。メール転記は1つの塊にする。
# 目印のない長いセグメントは「。」でも分ける。【現象】【原因】だけのセルは見出し型（header_cell）。
# ====================================================================================================

_BULLETS = "・●■◆◇□○▪►*"
_CIRCLED_RE = re.compile(r"[\u2460-\u2473\u2776-\u277f]|\(?\d{1,2}\)(?!\d)|\d{1,2}\.(?=[ \t])")
_EMAIL_START_RE = re.compile(
    r"\s*(?:-{3,}\s*(?:Original Message|元のメッセージ|Forwarded message)|(?:From|差出人)\s*:|>)", re.IGNORECASE
)
_SENTENCE_RE = re.compile(r"[^。]*。|[^。]+$")
_REFERENCE_RE = re.compile(r"同上|と同じ|別紙|参照")
_CORRECTION_RE = re.compile(r"訂正|撤回|ではなく")
_HEADER_LABEL_MAX = 10
_MIN_SENTENCE = 15


@dataclass
class _Head:
    pos: int                       # 本文（記入者判定）の開始位置
    when: HeadWhen | None = None
    when_text: str = ""
    label: str = ""
    marks: list[str] = field(default_factory=list)

    @property
    def kind(self) -> str | None:
        if self.when is not None:
            return head_kind(self.when)
        if self.label:
            return "label"
        if "bullet" in self.marks or "checklist" in self.marks:
            return "bullet"
        return None


@dataclass
class _Piece:
    start: int
    end: int
    line_start: bool = True
    marks: list[str] = field(default_factory=list)


def _skip_ws_until(sh: str, p: int, end: int) -> int:
    while p < end and sh[p] in " \t":
        p += 1
    return p


_HEAD_MEMO = threading.local()


def _parse_head(clean: str, sh: str, start: int, end: int, line_start: bool, not_date_res) -> _Head:
    """セグメント先頭の目印・日時を読む。

    1つのセルの処理（parse_log）の中で、同じ行を見出し型の判定・区切りの判定・分割後の先頭の読み取りで
    何度も読むので、同じセル（同じ clean・sh・not_date_res の組）の間だけ結果を覚える。戻り値は変更しないこと。
    """
    memo = getattr(_HEAD_MEMO, "value", None)
    if memo is None or memo[0] is not clean or memo[1] is not sh or memo[2] is not not_date_res:
        memo = (clean, sh, not_date_res, {})
        _HEAD_MEMO.value = memo
    k = (start, end, line_start)
    head = memo[3].get(k)
    if head is None:
        head = memo[3][k] = _read_head(clean, sh, start, end, line_start, not_date_res)
    return head


def _read_head(clean: str, sh: str, start: int, end: int, line_start: bool, not_date_res) -> _Head:
    p = _skip_ws_until(sh, start, end)
    head = _Head(pos=p)
    if p < end and (sh[p] in _BULLETS or (sh[p] == "-" and p + 1 < end and sh[p + 1] in " \t")):
        head.marks.append("bullet")
        p = _skip_ws_until(sh, p + 1, end)
    m = _CIRCLED_RE.match(sh[:end], p)
    if m:
        head.marks.append("checklist")
        p = _skip_ws_until(sh, m.end(), end)
    if p < end and sh[p] in "【[":
        close = sh.find("】" if sh[p] == "【" else "]", p + 1, min(end, p + 32))
        if close > 0:
            inner_s = _skip_ws_until(sh, p + 1, close)
            w = parse_when_at(sh, inner_s, line_start=True, not_date_res=not_date_res, end=close)
            if w is not None and _skip_ws_until(sh, w.end, close) == close:
                head.when, head.when_text = w, clean[inner_s:w.end].strip()
                head.pos = _skip_ws_until(sh, close + 1, end)
                return head
            inner = sh[inner_s:close].strip()
            if inner and len(inner) <= _HEADER_LABEL_MAX and not re.search(r"\d", inner):
                head.label = clean[inner_s:close].strip()
                head.pos = _skip_ws_until(sh, close + 1, end)
                return head
    w = parse_when_at(sh, p, line_start=line_start and not head.marks, not_date_res=not_date_res, end=end)
    if w is not None:
        head.when, head.when_text = w, clean[w.start:w.end].strip()
        p = w.end
    head.pos = p
    return head


def _line_anchor(clean: str, sh: str, s: int, e: int, options: SplitOptions, not_date_res, extra_res) -> str | None:
    p = _skip_ws_until(sh, s, e)
    if p >= e:
        return None
    if sh[p] == "※":
        return "note"
    if any(r.match(sh, p, min(e, p + USER_PATTERN_WINDOW)) for r in extra_res):
        return "extra"
    head = _parse_head(clean, sh, s, e, True, not_date_res)
    kind = head.kind
    if kind == "relative":
        # 相対日は、後ろが空白・時刻・区切り・漢字/カタカナのときだけ区切りにする
        q = head.when.end
        if q < e and sh[q] in "のにはもをがでと":
            return None
    return kind


def _is_email_end(clean: str, sh: str, s: int, e: int, not_date_res) -> bool:
    """メール転記の塊を終える行（行頭に年月日＋空白/「:」がある行）。"""
    p = _skip_ws_until(sh, s, e)
    w = parse_when_at(sh, p, line_start=True, not_date_res=not_date_res, end=e)
    return w is not None and w.has_date and (w.end >= e or sh[w.end] in " \t:")


def _detect_header_cell(clean: str, sh: str, lines, not_date_res) -> bool:
    labels = dated = 0
    for s, e in lines:
        if not sh[s:e].strip():
            continue
        head = _parse_head(clean, sh, s, e, True, not_date_res)
        if head.label:
            labels += 1
        elif head.when is not None:
            dated += 1
    return labels >= 2 and dated == 0


def _compile_all(patterns) -> list[re.Pattern]:
    """設定の正規表現をまとめてコンパイルする。検証より前に保存された設定で書けない形があっても、
    処理全体を止めずにその形だけ使わない。"""
    out = []
    for x in patterns or ():
        try:
            out.append(re.compile(x))
        except (re.error, TypeError, RecursionError):
            continue
    return out


def _lines(text: str) -> list[tuple[int, int]]:
    out, pos = [], 0
    for part in text.split("\n"):
        out.append((pos, pos + len(part)))
        pos += len(part) + 1
    return out


def _split_pieces(clean: str, sh: str, options: SplitOptions, not_date_res, header_mode: bool) -> list[_Piece]:
    extra_res = _compile_all(options.extra_anchors)
    pieces: list[_Piece] = []
    cur: _Piece | None = None
    in_email = False
    blank = False
    for s, e in _lines(clean):
        if not sh[s:e].strip():
            blank = not in_email
            continue
        if in_email:
            if not _is_email_end(clean, sh, s, e, not_date_res):
                cur.end = e
                continue
            in_email = False
        if _EMAIL_START_RE.match(sh[s:e]):
            cur = _Piece(s, e, True, ["email"])
            pieces.append(cur)
            in_email, blank = True, False
            continue
        if header_mode:
            head = _parse_head(clean, sh, s, e, True, not_date_res)
            anchor = "label" if head.label else None
        else:
            anchor = _line_anchor(clean, sh, s, e, options, not_date_res, extra_res)
        new = cur is None or (anchor not in (None, "note")) or (blank and anchor != "note")
        if anchor == "time" and options.time_only_lines == "join" and cur is not None and not blank:
            new = False
        if new:
            cur = _Piece(s, e, True, [])
            pieces.append(cur)
        else:
            cur.end = e
        if anchor == "note" and "note" not in cur.marks:
            cur.marks.append("note")
        blank = False
    if header_mode:
        for pc in pieces:
            if "email" not in pc.marks:
                pc.marks.append("header_cell")
        return pieces
    out: list[_Piece] = []
    for pc in pieces:
        if "email" in pc.marks:
            out.append(pc)
            continue
        out.extend(_split_inline(clean, sh, pc, not_date_res))
    return _split_sentences(clean, sh, out, options, not_date_res)


def _split_inline(clean: str, sh: str, pc: _Piece, not_date_res) -> list[_Piece]:
    """「／」「→」直後の日付、全角空白の後の「10:15：」で分ける。"""
    # 区切りの候補（直前が「/」「→」か全角空白）がなければ1文字ずつ調べない
    last = pc.end - 1
    if ("/" not in sh[pc.start:last] and "→" not in sh[pc.start:last]
            and "　" not in clean[pc.start:last]):
        return [pc]
    cuts = []
    for i in range(pc.start + 1, pc.end):
        ch, prev = sh[i], sh[i - 1]
        j = None
        if prev in "/→" and not (prev == "/" and i >= 2 and sh[i - 2].isdigit()):
            j = _skip_ws_until(sh, i, pc.end)
            w = parse_when_at(sh, j, line_start=False, not_date_res=not_date_res, end=pc.end)
            if w is None or not w.has_date:
                j = None
        elif clean[i - 1] == "\u3000" and ch.isdigit():
            w = parse_when_at(sh, i, line_start=False, not_date_res=not_date_res, end=pc.end)
            if w is not None and (w.has_date or (w.time and w.end < pc.end and sh[w.end] == ":")):
                j = i
        if j is not None and j > pc.start and (not cuts or j > cuts[-1]):
            line_head = clean.rfind("\n", pc.start, j)
            if sh[max(pc.start, line_head + 1):j].strip():
                cuts.append(j)
    if not cuts:
        return [pc]
    out, s = [], pc.start
    for j in cuts:
        out.append(_Piece(s, j, s == pc.start and pc.line_start, list(pc.marks)))
        s = j
    out.append(_Piece(s, pc.end, False, list(pc.marks)))
    return out


def _split_sentences(clean: str, sh: str, pieces: list[_Piece], options: SplitOptions, not_date_res) -> list[_Piece]:
    out: list[_Piece] = []
    limit = options.sentence_split_min_chars
    for pc in pieces:
        text = clean[pc.start:pc.end]
        if "email" in pc.marks or len(text.strip()) <= limit or "。" not in text.rstrip("。"):
            out.append(pc)
            continue
        head = _parse_head(clean, sh, pc.start, pc.end, pc.line_start, not_date_res)
        if head.kind is not None:
            out.append(pc)
            continue
        parts: list[_Piece] = []
        for m in _SENTENCE_RE.finditer(text):
            if not m.group(0).strip():
                continue
            s, e = pc.start + m.start(), pc.start + m.end()
            if parts and len(clean[s:e].strip()) < _MIN_SENTENCE:
                parts[-1].end = e
            elif parts and len(clean[parts[-1].start:parts[-1].end].strip()) < _MIN_SENTENCE:
                parts[-1].end = e
            else:
                parts.append(_Piece(s, e, pc.line_start and not parts, list(pc.marks) + ["sentence_split"]))
        out.extend(parts if len(parts) > 1 else [pc])
    return out


def _trim(clean: str, s: int, e: int) -> tuple[int, int]:
    while s < e and clean[s] in " \t　\n":
        s += 1
    while e > s and clean[e - 1] in " \t　\n":
        e -= 1
    return s, e


def parse_log(text, base_date: date | None = None, people: PeopleIndex | None = None,
              options: SplitOptions | None = None) -> LogParse:
    """セル1つのログを分割し、日時・記入者・識別子・数量・予定句を付ける（ルールのみ・決定的）。"""
    options = options or SplitOptions()
    people = people or PeopleIndex()
    clean = clean_log_text(text)
    if is_empty_log(clean):
        return LogParse([], "unknown", "empty", [], clean)
    sh = shadow(clean)
    not_date_res = _compile_all(options.not_date_patterns)
    header_mode = options.header_cells != "off" and _detect_header_cell(clean, sh, _lines(clean), not_date_res)
    pieces = _split_pieces(clean, sh, options, not_date_res, header_mode)

    segs: list[Segment] = []
    heads: list[HeadWhen | None] = []
    texts: list[str] = []
    authors = []
    skip: set[int] = set()
    for pc in pieces:
        s, e = _trim(clean, pc.start, pc.end)
        if s >= e:
            continue
        idx = len(segs)
        marks = list(pc.marks)
        label = ""
        head_when, when_text, author = None, "", None
        body_s, body_e = s, e
        if "email" in marks:
            skip_author = True
        else:
            skip_author = False
            head = _parse_head(clean, sh, s, e, pc.line_start, not_date_res)
            for mk in head.marks:
                if mk not in marks:
                    marks.append(mk)
            label = head.label
            head_when, when_text = head.when, head.when_text
            body_s = head.pos
        if "header_cell" in marks:
            skip.add(idx)
            head_when, when_text = None, ""
        elif not skip_author:
            author, body_s = detect_head_author(sh, body_s, e, people)
            tail_author, tail_end, dspan = detect_tail_author(sh, body_s, e, people)
            if author is None and tail_author is not None:
                author = tail_author
            if tail_author is not None or dspan is not None:
                if author is tail_author or dspan is not None:
                    body_e = tail_end
            if dspan is not None and head_when is None:
                w = parse_when_at(sh, dspan[0], line_start=False, not_date_res=not_date_res, end=dspan[1])
                if w is not None and w.has_date:
                    head_when, when_text = w, clean[dspan[0]:dspan[1]]
        body = clean[body_s:body_e].strip(" \t　:：、")
        body = body.rstrip(" \t　／/→\n").strip()
        raw = clean[s:e]
        sh_body = sh[body_s:body_e]
        if _REFERENCE_RE.search(sh_body) and "reference" not in marks:
            marks.append("reference")
        if _CORRECTION_RE.search(sh_body) and "correction" not in marks:
            marks.append("correction")
        segs.append(Segment(
            id=f"s{idx + 1}", raw=raw, body=body, start=s, end=e, when=None, author=author,
            identifiers=extract_identifiers(body), quantities=extract_quantities(body),
            plans=extract_plans(body), marks=marks, label=label,
        ))
        heads.append(head_when)
        texts.append(when_text)
        authors.append(author)

    warnings: list[str] = []
    whens, order, date_warnings = resolve_whens(
        heads, texts, base_date, options.order, options.order_tolerance_days, skip)
    warnings.extend(date_warnings)
    authors = inherit_authors(authors, skip)
    for seg, w, a in zip(segs, whens, authors):
        seg.when, seg.author = w, a
    if header_mode:
        kind = "header_cell"
        order = "unknown"
    elif len(segs) <= 1:
        kind = "single"
    else:
        kind = "log"
    if any("email" in s.marks for s in segs):
        warnings.append("メール転記を含む（1つのセグメントにまとめた）")
    if any("reference" in s.marks for s in segs):
        warnings.append("他の記録や資料への参照がある（内容は推測しない）")
    return LogParse(segs, order, kind, dedupe(warnings), clean)
