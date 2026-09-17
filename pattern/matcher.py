"""アップロードされたExcelがどのテンプレートに合うかを判定する（ルールベース）。

AI判定を追加する場合は rank_patterns の結果（上位候補が僅差のとき等）をAIに渡して並べ替える。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from excel.text import normalize_sheet_name
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.model import PatternDef

_NUMBER_LIKE = re.compile(r"^[\d\s,.\-/:+%]+$")


@dataclass
class PatternMatch:
    pattern: PatternDef
    confidence: int
    sheet_names: list[str]
    found_fields: int
    total_fields: int


def rank_patterns(info: WorkbookInfo, patterns: list[PatternDef]) -> list[PatternMatch]:
    matches = [match_pattern(info, p) for p in patterns]
    return sorted(matches, key=lambda m: m.confidence, reverse=True)


def match_pattern(info: WorkbookInfo, pattern: PatternDef) -> PatternMatch:
    found_by_sheet = {
        name: {fd.field_name for fd in pattern.fields if grid.find_labels(fd.label_norms())}
        for name, grid in info.grids.items()
    }
    total = len(pattern.fields) or 1

    chosen: list[str] = []
    if pattern.sheets:
        name_scores = []
        for sd in pattern.sheets:
            best = max(
                info.grids,
                key=lambda n: _sheet_name_score(sd.sheet_name, n) * 0.5 + len(found_by_sheet[n]) / total,
                default=None,
            )
            score = _sheet_name_score(sd.sheet_name, best) if best else 0.0
            if best is None or (score == 0 and not found_by_sheet[best]):
                name_scores.append(0.0)
                continue
            name_scores.append(score)
            if best not in chosen:
                chosen.append(best)
        sheet_score = sum(name_scores) / len(name_scores)
    else:
        best = max(info.grids, key=lambda n: len(found_by_sheet[n]), default=None)
        if best and found_by_sheet[best]:
            chosen.append(best)
        sheet_score = 1.0 if chosen else 0.0

    found = set().union(*(found_by_sheet[n] for n in chosen)) if chosen else set()
    weights = {fd.field_name: 2 if fd.required else 1 for fd in pattern.fields}
    field_score = sum(weights[f] for f in found) / sum(weights.values()) if weights else 0.0
    confidence = round(100 * (0.25 * sheet_score + 0.75 * field_score))
    return PatternMatch(pattern, confidence, chosen, len(found), len(pattern.fields))


def _sheet_name_score(expected: str, actual: str) -> float:
    e, a = normalize_sheet_name(expected), normalize_sheet_name(actual)
    if e == a:
        return 1.0
    if e and a and (e in a or a in e):
        return 0.6
    return 0.0


# ---- 一覧表らしさ（帳票フローから「一覧表の取り込みへ」を提案する判定） ----

MIN_TABLE_ROWS = 10
MIN_HEADER_CELLS = 3
DOMINANT_TABLE_ROWS = 30
SIMILAR_ROW_RATIO = 0.6


def looks_like_table_sheet(target: WorkbookInfo | SheetGrid, sheet_names: list[str] | None = None) -> bool:
    """見出し行の下に、同じ列の並び（空でない列の組）の行が10行以上続くシートがあれば True。

    帳票の中の小さな明細表（時系列・部品表）で誤判定しないよう、その表がシートの行の半分以上を
    占める（または30行以上続く）ことも条件にする。非表示シート（入力規則用のリスト等）は見ない。
    target に WorkbookInfo を渡すと sheet_names（省略時は全シート）のどれかが当てはまるかを返す。
    """
    if isinstance(target, SheetGrid):
        return _is_table_like(target)
    return bool(table_like_sheets(target, sheet_names))


def table_like_sheets(info: WorkbookInfo, sheet_names: list[str] | None = None) -> list[str]:
    names = sheet_names if sheet_names is not None else info.sheet_names
    return [n for n in names if n in info.grids and _is_table_like(info.grids[n])]


def _is_table_like(grid: SheetGrid) -> bool:
    if getattr(grid, "hidden", False):
        return False
    run, text_rows = _table_run_length(grid)
    return run >= MIN_TABLE_ROWS and (run >= DOMINANT_TABLE_ROWS or run * 2 >= text_rows)


def _table_run_length(grid: SheetGrid) -> tuple[int, int]:
    """戻り値: (見出し行の直下から続く「同じ形の行」の最大行数, 値のある行数)"""
    rows: dict[int, set[int]] = {}
    header_like: set[int] = set()
    by_row: dict[int, list[Cell]] = {}
    for cell in grid.text_cells():
        by_row.setdefault(cell.row, []).append(cell)
    for r, cells in by_row.items():
        rows[r] = {c.col for c in cells}
        strings = [c for c in cells if isinstance(c.value, str) and not _NUMBER_LIKE.match(c.norm)]
        if len(cells) >= MIN_HEADER_CELLS and len(strings) >= len(cells) * 0.8:
            header_like.add(r)

    best = 0
    for header_row in sorted(header_like):
        start = header_row + 1
        if start not in rows or start in header_like:
            continue
        base, run, r = rows[start], 0, start
        # 空行をはさまず、列の組が先頭データ行と似ている行を数える
        while r in rows and len(rows[r]) >= 2 and _similar(rows[r], base):
            run += 1
            r += 1
        best = max(best, run)
    return best, len(rows)


def _similar(a: set[int], b: set[int]) -> bool:
    return len(a & b) >= SIMILAR_ROW_RATIO * len(a | b)
