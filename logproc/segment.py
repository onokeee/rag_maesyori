"""追記ログ（date_log）の分割と、parse_log（分割・日時・記入者・抜き出しの一括処理）。

区切り: 行頭の日付・時刻・相対日・勤務帯、【】、・などの箇条書き、①、「／」「→」直後の日付、
全角空白の後の「10:15：」。※行・→行・目印のない行は直前につなぐ。メール転記は1つの塊にする。
目印のない長いセグメントは「。」でも分ける。【現象】【原因】だけのセルは見出し型（header_cell）。
"""
from __future__ import annotations

import re
import threading
from dataclasses import dataclass, field
from datetime import date

from logproc.dates import HeadWhen, head_kind, parse_when_at, resolve_whens
from logproc.extract import extract_identifiers, extract_plans, extract_quantities
from logproc.models import LogParse, Segment, SplitOptions
from logproc.people import PeopleIndex, detect_head_author, detect_tail_author, inherit_authors
from logproc.text import clean_log_text, is_empty_log, shadow

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


def _skip_ws(sh: str, p: int, end: int) -> int:
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
    p = _skip_ws(sh, start, end)
    head = _Head(pos=p)
    if p < end and (sh[p] in _BULLETS or (sh[p] == "-" and p + 1 < end and sh[p + 1] in " \t")):
        head.marks.append("bullet")
        p = _skip_ws(sh, p + 1, end)
    m = _CIRCLED_RE.match(sh[:end], p)
    if m:
        head.marks.append("checklist")
        p = _skip_ws(sh, m.end(), end)
    if p < end and sh[p] in "【[":
        close = sh.find("】" if sh[p] == "【" else "]", p + 1, min(end, p + 32))
        if close > 0:
            inner_s = _skip_ws(sh, p + 1, close)
            w = parse_when_at(sh, inner_s, line_start=True, not_date_res=not_date_res, end=close)
            if w is not None and _skip_ws(sh, w.end, close) == close:
                head.when, head.when_text = w, clean[inner_s:w.end].strip()
                head.pos = _skip_ws(sh, close + 1, end)
                return head
            inner = sh[inner_s:close].strip()
            if inner and len(inner) <= _HEADER_LABEL_MAX and not re.search(r"\d", inner):
                head.label = clean[inner_s:close].strip()
                head.pos = _skip_ws(sh, close + 1, end)
                return head
    w = parse_when_at(sh, p, line_start=line_start and not head.marks, not_date_res=not_date_res, end=end)
    if w is not None:
        head.when, head.when_text = w, clean[w.start:w.end].strip()
        p = w.end
    head.pos = p
    return head


def _line_anchor(clean: str, sh: str, s: int, e: int, options: SplitOptions, not_date_res, extra_res) -> str | None:
    p = _skip_ws(sh, s, e)
    if p >= e:
        return None
    if sh[p] == "※":
        return "note"
    if any(r.match(sh[:e], p) for r in extra_res):
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
    p = _skip_ws(sh, s, e)
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
            j = _skip_ws(sh, i, pc.end)
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
    return LogParse(segs, order, kind, _dedupe(warnings), clean)


def _dedupe(items: list[str]) -> list[str]:
    seen, out = set(), []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out
