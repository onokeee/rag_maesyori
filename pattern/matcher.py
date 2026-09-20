"""アップロードされたExcelがどのテンプレートに合うかを判定する（ルールベース）。

AI判定を追加する場合は rank_patterns の結果（上位候補が僅差のとき等）をAIに渡して並べ替える。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from excel.extractor import locate_value
from excel.text import normalize_sheet_name
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.dictionary import DICTIONARY_NORMS
from pattern.model import PatternDef

_NUMBER_LIKE = re.compile(r"^[\d\s,.\-/:+%]+$")
MIN_CONTINUATION_FIELDS = 2


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

    if chosen:
        _add_continuation_sheets(info, pattern, chosen, found_by_sheet)
    found = set().union(*(found_by_sheet[n] for n in chosen)) if chosen else set()
    field_score = len(found) / len(pattern.fields) if pattern.fields else 0.0
    confidence = round(100 * (0.25 * sheet_score + 0.75 * field_score))
    return PatternMatch(pattern, confidence, chosen, len(found), len(pattern.fields))


def _add_continuation_sheets(info: WorkbookInfo, pattern: PatternDef, chosen: list[str],
                             found_by_sheet: dict[str, set[str]]) -> None:
    """「8D報告(1)」「8D報告(2)」のように1件の帳票が複数シートに分かれている場合、続きのシートも選ぶ。

    選んだシートに無い項目の値が2つ以上（項目数の1割以上）あるシートを足す。一覧表らしいシート・非表示シートと、
    見出しだけで値の無いシート（記入要領など）は足さない。選んだシートはブック内の順に並べ直す。
    """
    stop_labels = pattern.label_norms() | DICTIONARY_NORMS

    def with_value(name: str) -> set[str]:
        grid = info.grids[name]
        return {fd.field_name for fd in pattern.fields
                if fd.field_name in found_by_sheet[name] and locate_value(grid, fd, stop_labels)[1]}

    covered = set().union(*(with_value(n) for n in chosen))
    need = max(MIN_CONTINUATION_FIELDS, 0.1 * len(pattern.fields))
    candidates = []
    for name, grid in info.grids.items():
        if name in chosen or grid.hidden or len(found_by_sheet[name] - covered) < need:
            continue
        candidates.append((name, with_value(name)))
    for name, values in sorted(candidates, key=lambda c: -len(c[1])):
        if len(values - covered) >= need and not _is_table_like(info.grids[name]):
            chosen.append(name)
            covered |= values
    order = {name: i for i, name in enumerate(info.grids)}
    chosen.sort(key=order.get)


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
