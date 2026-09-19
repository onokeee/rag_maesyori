"""Markdown 出力のテキスト処理（帳票・一覧表で共通）。

決まり（docs/design.md 6章）: UTF-8・LF、レコード内に空行を入れない、パイプ表を使わない（`- 項目: 値`）、
複数行の値は2文字下げの連続行。
"""
from __future__ import annotations

import re
import unicodedata

_SPACES = re.compile(r"[^\S\n]+")          # 改行以外の空白の連続
_HEADING = re.compile(r"^(#{1,6})(\s|$)")
_LIST = re.compile(r"^([-*+])(\s|$)")
_ORDERED = re.compile(r"^(\d{1,9})([.)])(\s|$)")
_RULE = re.compile(r"^(?:-[ \t]*){3,}$|^(?:=[ \t]*){3,}$|^(?:\*[ \t]*){3,}$|^(?:_[ \t]*){3,}$")
_FENCE = re.compile(r"^(```|~~~)")


def _enclosed_marks() -> str:
    """NFKC で囲みが外れてしまう囲み文字（丸数字 ①、丸英字 Ⓐ、丸カナ ㋐、丸漢字 ㊤ など）を集める。"""
    marks = []
    for start, end in ((0x2460, 0x24FF), (0x3240, 0x32FF), (0x1F100, 0x1F1FF)):
        for cp in range(start, end + 1):
            ch = chr(cp)
            if "CIRCLED" in unicodedata.name(ch, "") and unicodedata.normalize("NFKC", ch) != ch:
                marks.append(ch)
    return "".join(marks)


# ①→1 のように囲みが外れると「①破損…」が「1破損…」になり、番号と本文の区切りが消えて読めなくなる。
# 丸数字は帳票の手順・項目番号でごく普通に使われるので、囲み文字だけは NFKC をかけずに残す。
# （㈱→(株)、⑴→(1) のように区切りが残る表記はそのまま NFKC で正規化する）
_ENCLOSED = re.compile(f"([{re.escape(_enclosed_marks())}])")


def nfkc_keep_enclosed(text: str) -> str:
    """NFKC 正規化。ただし囲み文字（①Ⓐ㋐㊤…）はそのまま残す。

    囲み文字は前後の文字と結合しないので、そこで区切って正規化しても結果は変わらない。
    """
    if not _ENCLOSED.search(text):
        return unicodedata.normalize("NFKC", text)
    return "".join(part if _ENCLOSED.fullmatch(part) else unicodedata.normalize("NFKC", part)
                   for part in _ENCLOSED.split(text))


def nfkc_value(text) -> str:
    """値の正規化。NFKC（囲み文字は残す）＋空白の畳み込み（改行は保持、行末の空白と前後の空行は除く）。"""
    if text is None:
        return ""
    s = nfkc_keep_enclosed(str(text)).replace("\r\n", "\n").replace("\r", "\n")
    lines = [_SPACES.sub(" ", line).strip() for line in s.split("\n")]
    return "\n".join(lines).strip("\n")


def escape_md_line(line: str) -> str:
    """行頭の見出し・箇条書き・引用・番号付きリスト、区切り線、コードフェンスとして解釈されないようにする。"""
    body = line.lstrip(" \t")
    indent = line[: len(line) - len(body)]
    if not body:
        return line
    if _FENCE.match(body):
        body = "\\" + body[0] + "\\" + body[1] + "\\" + body[2:]
    elif _RULE.match(body):
        body = "\\" + body
    elif _HEADING.match(body) or _LIST.match(body) or body.startswith(">"):
        body = "\\" + body
    else:
        m = _ORDERED.match(body)
        if m:
            body = m.group(1) + "\\" + body[len(m.group(1)):]
    return indent + body


def _value_lines(value) -> list[str]:
    if value is None:
        return []
    items = value if isinstance(value, (list, tuple)) else [value]
    lines: list[str] = []
    for item in items:
        if item is None:
            continue
        text = str(item).replace("\r\n", "\n").replace("\r", "\n")
        lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
    return lines


def md_bullet(label: str, value) -> list[str]:
    """`- label: value`。複数行は `- label:` の後に2文字下げで続ける。値が空なら出さない（[]）。"""
    lines = _value_lines(value)
    if not lines:
        return []
    label = " ".join(str(label).split())
    if len(lines) == 1:
        return [f"- {label}: {lines[0]}"]
    return [f"- {label}:"] + [f"  {escape_md_line(ln)}" for ln in lines]


def estimate_tokens(text: str) -> int:
    """推定トークン数（実トークン以上になる見積もり）。非ASCII 1文字=1.1、ASCII 2文字=1。

    LightRAG の o200k_base と比べた実測（実出力 6,150 ブロック）で、この式は実/推定の p95 0.98・最大 1.05。
    旧式（非ASCII 1、ASCII 3文字=1）は中央値で 13%・最大 29% 少なく見積もっていた。
    """
    if not text:
        return 0
    ascii_count = len(text.encode("ascii", "ignore"))  # ASCII の文字数（1文字ずつ数えるより速い）
    return -(-(11 * (len(text) - ascii_count) + 5 * ascii_count) // 10)  # 整数だけで ceil する（浮動小数の誤差を避ける）


def join_blocks(blocks: list[list[str]]) -> str:
    """ブロック間に空行1つ、ブロック内の空行は除く、末尾改行1つ、LF。"""
    parts: list[str] = []
    for block in blocks:
        lines: list[str] = []
        for line in block or []:
            for ln in str(line).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
                if ln.strip():
                    lines.append(ln.rstrip())
        if lines:
            parts.append("\n".join(lines))
    return "\n\n".join(parts) + "\n" if parts else ""
