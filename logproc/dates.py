"""ログのエントリ先頭にある日時表現の読み取りと、年・記入順の決定。

- 読み取るのはセグメント先頭（行頭・【】の中・「／」「→」の直後）の表現だけ。
  本文中の「納期4/8」「1.5mm」「4/10に伺います」は出来事の日付にしない。
- 年は、まず記入順（古い順／新しい順）を判定し、時間の流れに沿って前のエントリと矛盾しない年を選ぶ。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, timedelta

from logproc.models import WhenInfo

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


def parse_when_at(sh: str, pos: int, *, line_start: bool = True, not_date_res=(), end: int | None = None) -> HeadWhen | None:
    """影テキストの pos から始まる日時表現を読む。何もなければ None。"""
    end = len(sh) if end is None else end
    view = sh[:end]
    p = pos
    w = HeadWhen(start=pos, end=pos)
    found = False
    if not (not_date_res and any(r.match(view, p) for r in not_date_res)):
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
