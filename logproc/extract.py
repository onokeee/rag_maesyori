"""識別子（型番・アラームコード・ロット）、数量・回数、予定句の抜き出し。"""
from __future__ import annotations

import re

from logproc.text import shadow

_ID_RE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z0-9](?:[A-Za-z0-9._\-]*[A-Za-z0-9])?")
_UNITS = sorted(
    "個 本 枚 台 回 件 箇所 ヶ所 か所 カ所 セット 式 袋 缶 巻 m mm cm μm MPa kPa Pa ℃ °C V kV mA A kW W L mL ml "
    "分 秒 時間 日間 週間 か月 ヶ月 カ月 kg g t MΩ Ω M % rpm Hz P 万円 円 人".split(),
    key=len, reverse=True,
)
_UNIT_ALT = "|".join(re.escape(u) for u in _UNITS)
# 英字で終わる単位は後ろに英字が続かないこと（「5mm」と「5mmHg」などの区別）。和文の単位は制限しない
_QTY_UNIT = "(?:" + "|".join(re.escape(u) + ("(?![A-Za-z])" if u[-1].isascii() and u[-1].isalpha() else "") for u in _UNITS) + ")"
_ASCII_UNIT_RE = re.compile(rf"\d+(?:\.\d+)?(?:E[-+]?\d+)?(?:{_UNIT_ALT})", re.IGNORECASE)
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
_PLAN_KEY_RE = re.compile(r"納期|予定|目処|目途|見込み|までに")
_CLAUSE_SEP_RE = re.compile(r"[、。,，()（）\s→・「」【】]+")


def identifier_spans(text: str) -> list[tuple[int, int, str]]:
    """英字と数字を両方含む語の位置。戻り値の語は NFKC 後の表記。"""
    sh = shadow(text)
    out = []
    for m in _ID_RE.finditer(sh):
        tok = m.group(0)
        if not (re.search(r"[A-Za-z]", tok) and re.search(r"\d", tok)):
            continue
        if _ASCII_UNIT_RE.fullmatch(tok) or any(r.fullmatch(tok) for r in _NOT_ID_RES):
            continue
        out.append((m.start(), m.end(), tok))
    return out


def extract_identifiers(text: str) -> list[str]:
    return _dedupe(tok for _, _, tok in identifier_spans(text))


def extract_quantities(text: str) -> list[str]:
    """数値＋単位、×1、2回、金額。原文の表記（NFKC 後）で返す。"""
    sh = shadow(text)
    id_spans = identifier_spans(text)
    out = []
    for m in _QTY_RE.finditer(sh):
        if any(s <= m.start() and m.end() <= e for s, e, _ in id_spans):
            continue
        out.append(m.group(0).strip())
    return _dedupe(out)


def extract_plans(text: str) -> list[str]:
    """「納期1週間」「6月末目処」「交換予定」などの予定句（原文の表記）。"""
    out = []
    for piece in _CLAUSE_SEP_RE.split(text or ""):
        if piece and _PLAN_KEY_RE.search(shadow(piece)):
            out.append(piece)
    return _dedupe(out)


def _dedupe(items) -> list[str]:
    seen, out = set(), []
    for x in items:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out
