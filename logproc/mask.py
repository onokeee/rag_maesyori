"""個人情報のマスク（電話番号・メールアドレス、任意で人名・金額）。

AI に送る前と md 出力の両方で使う。判定は1文字ずつ NFKC した写しで行い、位置は原文と同じ。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Iterable

from logproc.text import shadow

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
