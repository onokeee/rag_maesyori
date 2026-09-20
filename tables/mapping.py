"""列の対応づけ候補と、取り込み設定の候補選び。

照合の順番: 取り込み設定の見出し候補 → 表用の標準キー辞書 → 似た文字列（候補として出すだけ）。
設定との照合は見出し帯の中だけで行う（データ値で点数が上がらないように）。
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass

from tables.detect import split_header_unit
from tables.dictionary import lookup_header, norm_header
from tables.source import NA_TOKENS, cell_text, value_kind

_LOG_LINE_RE = re.compile(
    r"^\s*[【\[(（<]?\s*(\d{4}[/\-.年])?\d{1,2}[/\-.月]\d{1,2}日?|^\s*(R|令和)\d{1,2}\.\d{1,2}\.\d{1,2}",
    re.M,
)
_MERGED = object()  # 結合範囲の左上以外のセル（空欄ではない扱い）
_NUMBER_TYPES = {"number"}
_DATE_TYPES = {"date", "datetime"}

@dataclass
class ColumnSuggestion:
    index: int
    header: str
    unit: str
    key: str | None
    display: str
    type: str
    role: str
    examples: list[str]
    type_error_rate: float
    blank_rate: float
    md: str  # body/attribute/omit
    fill_down_blank: bool = False
    log: bool = False
    matched_by: str = ""  # template/dictionary/similar/none
    inferred_type: str = ""  # 値から推定した型（type と違えば画面で知らせる）
    # md が "omit" になった理由（画面の説明用）: blank（空欄だけ）/template（取り込み設定のとおり）
    omit_reason: str = ""


def suggest_columns(headers: list[str], sample_rows, template_spec=None) -> list[ColumnSuggestion]:
    """見出しと先頭のデータ行から、列ごとの標準キー・型・役割などの候補を作る。

    sample_rows: SourceRow のリスト、または値のリストのリスト。
    template_spec: TableSpec（dataclass）か、その dict。あれば設定の見出し候補を先に照合する。
    """
    grid = [_row_values(r) for r in sample_rows]
    template_columns = _get(template_spec, "columns", []) if template_spec is not None else []
    out: list[ColumnSuggestion] = []
    for i, header in enumerate(headers):
        unit = split_header_unit(header)[1]
        values = [row[i] if i < len(row) else (None, "") for row in grid]
        stats = _column_stats(values)

        spec_col = _match_template_column(header, template_columns)
        if spec_col is not None:
            s = ColumnSuggestion(
                index=i, header=header, unit=_get(spec_col, "unit", "") or unit,
                key=_get(spec_col, "key", None), display=_get(spec_col, "display", "") or _display_name(header),
                type=_get(spec_col, "type", "") or stats["type"], role=_get(spec_col, "role", "attribute"),
                examples=stats["examples"], type_error_rate=0.0, blank_rate=stats["blank_rate"],
                md=_get(spec_col, "md", "attribute"), fill_down_blank=bool(_get(spec_col, "fill_down_blank", False)),
                matched_by="template", inferred_type=stats["type"],
            )
        else:
            found = lookup_header(header)
            if found and found[1] == "similar" and found[0].role == "person" and _looks_like_prose(values, stats):
                # 「対応」「作業」が似た語の「対応者」「作業者」に当たっても、値が文章なら人名の列ではない
                found = None
            if found:
                std, how = found
                s = ColumnSuggestion(
                    index=i, header=header, unit=unit or std.unit, key=std.key,
                    display=_display_name(header), type=std.type, role=std.role,
                    examples=stats["examples"], type_error_rate=0.0, blank_rate=stats["blank_rate"],
                    md=std.md, matched_by=how, inferred_type=stats["type"],
                )
            else:
                role = "text" if stats["type"] == "text" else "attribute"
                s = ColumnSuggestion(
                    index=i, header=header, unit=unit, key=None, display=_display_name(header), type=stats["type"], role=role,
                    examples=stats["examples"], type_error_rate=0.0, blank_rate=stats["blank_rate"],
                    md="body" if role == "text" else "attribute", matched_by="none", inferred_type=stats["type"],
                )
        if s.md == "omit":
            s.omit_reason = "template"
        s.type_error_rate = _type_error_rate(values, s.type)
        if stats["log_like"] and s.role in ("text", "log", "attribute") and s.type == "text":
            s.role, s.log = "log", True
        if s.role == "log":
            s.log = True
        if s.blank_rate >= 1.0 and s.matched_by != "template":
            s.md, s.omit_reason = "omit", "blank"
        if s.matched_by != "template" and s.role in ("entity", "entity_label") and _looks_filled_down(values):
            s.fill_down_blank = True
        out.append(s)
    _keep_one_log(out)
    _dedupe_keys(out)
    return out


def _keep_one_log(out: list[ColumnSuggestion]) -> None:
    """AI整形の対象（追記ログ）は1列だけ。取り込み設定・辞書で当たった列、なければ最初の列を残し、
    ほかは長文にする（何も変えずに保存しただけで「AI整形の対象は1列だけ」と断られないように）。"""
    logs = [s for s in out if s.role == "log"]
    if len(logs) <= 1:
        return
    rank = {"template": 0, "dictionary": 1, "similar": 2}
    best = min(logs, key=lambda s: (rank.get(s.matched_by, 3), s.index))
    for s in logs:
        if s is not best:
            s.role, s.log = "text", False
            if s.md != "omit":
                s.md = "body"


def match_templates(headers: list[str], sheet_or_file_name: str, specs: list) -> list[tuple[object, int, int]]:
    """取り込み設定の候補。戻り値: (設定, 一致した必須列数, 必須列数) を点数の高い順に。

    点数 = 0.25×名前 + 0.6×必須列が見つかった割合 + 0.15×列の並び順。
    必須列が1つもない設定は、全列を必須列とみなして数える。
    """
    header_norms = [_header_norms(h) for h in headers]
    name_norm = norm_header(sheet_or_file_name)
    scored = []
    for spec in specs:
        columns = list(_get(spec, "columns", []) or [])
        required = [c for c in columns if _get(c, "required", False)] or columns
        positions: list[int] = []
        matched = 0
        for col in required:
            pos = _find_header(col, header_norms)
            if pos is not None:
                matched += 1
                positions.append(pos)
        total = len(required)
        ratio = matched / total if total else 0.0
        patterns = [norm_header(p) for p in (_get(spec, "name_patterns", []) or []) if p]
        name_hit = 1.0 if any(p and p in name_norm for p in patterns) else 0.0
        order = _increasing_share(positions)
        score = 0.25 * name_hit + 0.6 * ratio + 0.15 * order
        scored.append((score, spec, matched, total))
    scored.sort(key=lambda t: -t[0])
    return [(spec, matched, total) for _, spec, matched, total in scored]


# ---- 内部 ----

def _display_name(header: str) -> str:
    """表示名の候補。2段見出しは下段、単位は除く（「交換部品_単価(円)」→「単価」）。"""
    from tables.detect import is_month_label

    if is_month_label(header):
        return header  # クロス集計の年月見出しは年を残す
    lower = header.rsplit("_", 1)[-1] if "_" in header else header
    name = split_header_unit(lower)[0]
    return name or header


def _get(obj, name: str, default=None):
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _row_values(row) -> list[tuple[object, str]]:
    cells = getattr(row, "cells", None)
    if cells is not None:
        index = getattr(row, "index", 0)
        return [
            (_MERGED, "") if (not c.text and c.merged_anchor and c.merged_anchor != (index, i + 1)) else (c.value, c.text)
            for i, c in enumerate(cells)
        ]
    return [(v, cell_text(v)) for v in row]


def _header_norms(header: str) -> set[str]:
    name, _unit = split_header_unit(header)
    norms = {norm_header(header), norm_header(name)}
    if "_" in header:  # 2段見出しは下段だけでも照合する
        lower = header.split("_")[-1]
        norms |= {norm_header(lower), norm_header(split_header_unit(lower)[0])}
    return {n for n in norms if n}


def _match_template_column(header: str, columns) -> object | None:
    norms = _header_norms(header)
    for col in columns:
        candidates = list(_get(col, "headers", []) or []) + [_get(col, "display", "")]
        for cand in candidates:
            if cand and (norm_header(cand) in norms or norm_header(split_header_unit(cand)[0]) in norms):
                return col
    return None


def _find_header(col, header_norms: list[set[str]]) -> int | None:
    candidates = list(_get(col, "headers", []) or []) + [_get(col, "display", "")]
    cand_norms = {norm_header(c) for c in candidates if c} | {norm_header(split_header_unit(c)[0]) for c in candidates if c}
    cand_norms.discard("")
    for pos, norms in enumerate(header_norms):
        if norms & cand_norms:
            return pos
    return None


def _increasing_share(positions: list[int]) -> float:
    """見つかった列が設定の順に並んでいる割合（最長増加部分列 / 件数）。"""
    if not positions:
        return 0.0
    tails: list[int] = []
    for p in positions:
        lo, hi = 0, len(tails)
        while lo < hi:
            mid = (lo + hi) // 2
            if tails[mid] < p:
                lo = mid + 1
            else:
                hi = mid
        if lo == len(tails):
            tails.append(p)
        else:
            tails[lo] = p
    return len(tails) / len(positions)


def _is_blank(value, text: str) -> bool:
    if value is _MERGED:
        return False
    return not text or text.strip() in NA_TOKENS


def _looks_like_prose(values: list[tuple[object, str]], stats: dict) -> bool:
    """値が人名ではなく文章か（長文、または平均で人名より明らかに長い）。"""
    if stats["type"] == "text":
        return True
    texts = [t for v, t in values if not _is_blank(v, t) and v is not _MERGED]
    return bool(texts) and sum(len(t) for t in texts) / len(texts) > 15


def _column_stats(values: list[tuple[object, str]]) -> dict:
    total = len(values)
    filled = [(v, t) for v, t in values if not _is_blank(v, t)]
    nonblank = [(v, t) for v, t in filled if v is not _MERGED]
    kinds = Counter(value_kind(v, t) for v, t in nonblank)
    examples: list[str] = []
    for _v, t in nonblank:
        short = t if len(t) <= 60 else t[:60] + "…"
        if short not in examples:
            examples.append(short)
        if len(examples) >= 3:
            break
    log_hits = sum(1 for _v, t in nonblank if len(_LOG_LINE_RE.findall(t)) >= 2)
    return {
        "type": _infer_type(kinds, nonblank),
        "blank_rate": round(1 - len(filled) / total, 3) if total else 1.0,
        "examples": examples,
        "log_like": bool(nonblank) and log_hits >= max(2, 0.3 * len(nonblank)),
    }


def _infer_type(kinds: Counter, nonblank: list[tuple[object, str]]) -> str:
    n = sum(kinds.values())
    if not n:
        return "string"
    if kinds.get("text", 0) >= 0.2 * n:
        return "text"
    top, count = kinds.most_common(1)[0]
    if top in ("date", "datetime"):
        return "datetime" if kinds.get("datetime", 0) >= 0.2 * n else "date"
    if top == "time":
        return "time"
    if top == "number":
        # 先頭ゼロ付きの数字（コード）が混ざるならコード
        return "code" if kinds.get("code", 0) >= 0.2 * n else "number"
    if top == "code":
        return "code"
    distinct = len({unicodedata.normalize("NFKC", t) for _v, t in nonblank})
    if n >= 20 and distinct <= 12 and distinct <= 0.2 * n:
        return "enum"
    return "string"


def _type_error_rate(values: list[tuple[object, str]], type_: str) -> float:
    nonblank = [(v, t) for v, t in values if v is not _MERGED and not _is_blank(v, t)]
    if not nonblank:
        return 0.0
    if type_ == "number":
        ok_kinds = _NUMBER_TYPES
    elif type_ in ("date", "datetime"):
        ok_kinds = _DATE_TYPES
    elif type_ == "time":
        ok_kinds = {"time", "datetime"}
    elif type_ == "code":
        ok_kinds = {"code", "number", "string"}
    else:
        return 0.0
    errors = sum(1 for v, t in nonblank if value_kind(v, t) not in ok_kinds)
    return round(errors / len(nonblank), 3)


def _looks_filled_down(values: list[tuple[object, str]]) -> bool:
    """空欄が「上と同じ」の意味で使われていそうか（値の後に空欄が続く）。"""
    texts = [t for v, t in values if v is not _MERGED]
    blanks = sum(1 for t in texts if not t)
    if not texts or blanks < 0.2 * len(texts) or blanks > 0.9 * len(texts):
        return False
    return bool(texts[0])


def _dedupe_keys(items: list[ColumnSuggestion]) -> None:
    """同じ標準キーが複数の列に付いたら、辞書で完全一致した列を優先し、残りは番号を付ける。"""
    by_key: dict[str, list[ColumnSuggestion]] = {}
    for s in items:
        if s.key:
            by_key.setdefault(s.key, []).append(s)
    rank = {"template": 0, "dictionary": 1, "similar": 2}
    for key, group in by_key.items():
        if len(group) < 2:
            continue
        group.sort(key=lambda s: (rank.get(s.matched_by, 3), s.index))
        for n, s in enumerate(group[1:], start=2):
            s.key = f"{key}_{n}"
