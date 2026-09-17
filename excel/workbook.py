"""Excelの構造解析。結合セルを考慮した「値のあるセル」のグリッドを作る。"""
from __future__ import annotations

import warnings
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import MAC_EPOCH

from excel.image_detector import detect_images
from excel.text import cell_text, normalize_label, split_inline


@dataclass
class Cell:
    """値を持つセル。結合セルの場合は左上セルが範囲全体を代表する。"""

    row: int
    col: int
    max_row: int
    max_col: int
    value: object
    text: str
    norm: str
    inline: tuple[str, str] | None
    bold: bool = False
    filled: bool = False

    @property
    def coord(self) -> str:
        start = f"{get_column_letter(self.col)}{self.row}"
        if (self.max_row, self.max_col) == (self.row, self.col):
            return start
        return f"{start}:{get_column_letter(self.max_col)}{self.max_row}"

    def is_label(self, label_norms: set[str]) -> bool:
        return self.norm in label_norms or (self.inline is not None and self.inline[0] in label_norms)


class SheetGrid:
    def __init__(self, ws):
        self.name: str = ws.title
        self.hidden: bool = getattr(ws, "sheet_state", "visible") != "visible"
        self._bounds: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for rng in ws.merged_cells.ranges:
            bounds = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._bounds[(r, c)] = bounds

        self.cells: dict[tuple[int, int], Cell] = {}
        self.max_row = self.max_col = 0
        for (r, c), xl in ws._cells.items():
            text = cell_text(xl.value)
            if not text:
                continue
            top, left, bottom, right = self.bounds(r, c)
            if (top, left) != (r, c):
                continue
            self.cells[(r, c)] = Cell(
                row=r, col=c, max_row=bottom, max_col=right,
                value=xl.value, text=text, norm=normalize_label(text), inline=split_inline(text),
                bold=bool(xl.font and xl.font.b),
                filled=bool(xl.fill and xl.fill.fill_type == "solid"),
            )
            self.max_row = max(self.max_row, bottom)
            self.max_col = max(self.max_col, right)

        self._by_norm: dict[str, list[Cell]] = defaultdict(list)
        self._by_inline: dict[str, list[Cell]] = defaultdict(list)
        for cell in self.text_cells():
            self._by_norm[cell.norm].append(cell)
            if cell.inline:
                self._by_inline[cell.inline[0]].append(cell)

    def bounds(self, row: int, col: int) -> tuple[int, int, int, int]:
        """(top, left, bottom, right) を返す。結合されていなければ自セルのみ。"""
        return self._bounds.get((row, col), (row, col, row, col))

    def cell_at(self, row: int, col: int) -> Cell | None:
        top, left, _, _ = self.bounds(row, col)
        return self.cells.get((top, left))

    def text_cells(self) -> list[Cell]:
        return [self.cells[key] for key in sorted(self.cells)]

    def find_labels(self, label_norms: set[str]) -> list[Cell]:
        """ラベル候補に一致するセルを上→下、左→右の順で返す（セル内ラベルも含む）。"""
        hits = {}
        for norm in label_norms:
            for cell in self._by_norm.get(norm, []) + self._by_inline.get(norm, []):
                hits[(cell.row, cell.col)] = cell
        return [hits[key] for key in sorted(hits)]


@dataclass
class WorkbookInfo:
    path: Path
    grids: dict[str, SheetGrid]
    images: list[dict]
    date1904: bool = False  # 1904年基準のブック（日付シリアル値の起点が違う）

    @property
    def sheet_names(self) -> list[str]:
        return list(self.grids)

    def images_in(self, sheet_names: list[str]) -> list[dict]:
        return [img for img in self.images if img["sheet"] in sheet_names]


def load_workbook_info(path: str | Path) -> WorkbookInfo:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(path, data_only=True)
    try:
        grids = {ws.title: SheetGrid(ws) for ws in wb.worksheets}
        date1904 = wb.epoch == MAC_EPOCH
    finally:
        wb.close()
    return WorkbookInfo(path=Path(path), grids=grids, images=detect_images(path), date1904=date1904)
