"""ルール処理結果から「対応の時系列」の行を作る（keep: 本文は原文＋用語集）。"""
from __future__ import annotations

import re

from logproc.glossary import apply_glossary
from logproc.models import AuthorInfo, LogParse, Segment, WhenInfo
from logproc.people import key


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
