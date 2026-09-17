"""用語集による言い換え（様子見→経過観察 など）。

- 長い語から順に照合する（1回の走査。置き換えた結果を再度置き換えない）。
- 識別子（FE-100 など）とマスク記号の範囲は置き換えない。
- ascii_boundary: 前後が英数字・「-」「_」なら置き換えない。英字だけの語は既定でオン。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from logproc.extract import identifier_spans
from logproc.text import shadow

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
