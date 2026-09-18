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
from excel.text import (MAX_LABEL_LENGTH, cell_text, label_base, label_parts, normalize_label, paren_stripped,
                        section_stripped, split_inline)


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
    fill: str = ""  # 塗りつぶし色の識別キー（"rgb:FFDDEEFF" / "theme:4:0.6" など。塗りなしは ""）
    alt_norm: str = ""  # 先頭の項番を除いたラベル（「３．暫定対策」→「暫定対策」）。無ければ ""
    part_norms: tuple[str, ...] = ()  # 「ライン／工程」のように2つのラベルをまとめた見出しの各部分

    @property
    def coord(self) -> str:
        start = f"{get_column_letter(self.col)}{self.row}"
        if (self.max_row, self.max_col) == (self.row, self.col):
            return start
        return f"{start}:{get_column_letter(self.max_col)}{self.max_row}"

    @property
    def label_keys(self) -> tuple[str, ...]:
        """セル全体をラベルとみなすときの比較キー。"""
        return (self.norm, self.alt_norm) if self.alt_norm else (self.norm,)

    def matches(self, label_norms: set[str]) -> bool:
        """セル全体がラベル候補に一致するか（セル内ラベル「設備番号：EQ-001」は含まない）。"""
        return self.norm in label_norms or (bool(self.alt_norm) and self.alt_norm in label_norms)

    def is_label(self, label_norms: set[str]) -> bool:
        return self.matches(label_norms) or (self.inline is not None and self.inline[0] in label_norms)


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
                fill=fill_key(xl),
                alt_norm=section_stripped(text) if _label_like(xl.value, text) else "",
                part_norms=label_parts(text) if _label_like(xl.value, text) else (),
            )
            self.max_row = max(self.max_row, bottom)
            self.max_col = max(self.max_col, right)

        self._by_norm: dict[str, list[Cell]] = defaultdict(list)
        self._by_inline: dict[str, list[Cell]] = defaultdict(list)
        self._by_base: dict[str, list[str]] = defaultdict(list)  # 括弧書きを除いたラベル → 元のキー
        self._by_variant: dict[str, list[str]] = defaultdict(list)  # 括弧書き・「内容」を除いたラベル → 元のキー
        self._by_part: dict[str, list[Cell]] = defaultdict(list)  # 「ライン／工程」の「ライン」「工程」
        for cell in self.text_cells():
            for key in cell.part_norms:
                self._by_part[key].append(cell)
            for key in cell.label_keys:
                self._by_norm[key].append(cell)
            if cell.inline:
                self._by_inline[cell.inline[0]].append(cell)
            for key in (*cell.label_keys, *(cell.inline[:1] if cell.inline else ())):
                base = paren_stripped(key)
                if base:
                    self._by_base[base].append(key)
                # 表記違いの照合は見出しらしいセル（塗りつぶし・太字・「ラベル：値」）だけ。値の「完了」を「完了日時」と見ない
                if cell.filled or cell.bold or (cell.inline and key == cell.inline[0]):
                    self._by_variant[label_base(key) or key].append(key)

    def bounds(self, row: int, col: int) -> tuple[int, int, int, int]:
        """(top, left, bottom, right) を返す。結合されていなければ自セルのみ。"""
        return self._bounds.get((row, col), (row, col, row, col))

    def cell_at(self, row: int, col: int) -> Cell | None:
        top, left, _, _ = self.bounds(row, col)
        return self.cells.get((top, left))

    def text_cells(self) -> list[Cell]:
        return [self.cells[key] for key in sorted(self.cells)]

    def resolve_labels(self, label_norms: set[str]) -> set[str]:
        """シートで探すラベルの集合。候補どおりのラベルが無いときだけ、末尾の括弧書きの違いを許す。

        例: 候補「停止時間」でシートに「停止時間(分)」だけがある（または逆）。「原因(推定)」と「原因(確定)」の
        ように括弧書きで区別しているラベルは、候補どおりのものがあればそちらだけを使う。
        """
        if any(n in self._by_norm or n in self._by_inline or n in self._by_part for n in label_norms):
            return label_norms
        bases = {paren_stripped(n) or n for n in label_norms}
        extra = {key for base in bases for key in self._by_base.get(base, [])}
        extra |= {base for base in bases if base in self._by_norm or base in self._by_inline}
        return label_norms | extra if extra else label_norms

    def label_variants(self, label_norms: set[str]) -> set[str]:
        """シートにある、候補と表記だけが違うラベル（候補どおりのラベルから値が取れなかったときに使う）。

        「発生原因」と「発生原因（なぜ起きたか）」、「応急処置」と「応急処置内容」。
        括弧書きどうしが違うもの（「原因（推定）」と「原因（確定）」）は別の項目なので含めない。
        """
        found = set()
        for norm in label_norms:
            qualified = bool(paren_stripped(norm))
            for key in self._by_variant.get(label_base(norm) or norm, []):
                if key in label_norms or (qualified and paren_stripped(key)):
                    continue
                found.add(key)
        return found

    def find_labels(self, label_norms: set[str]) -> list[Cell]:
        """ラベル候補に一致するセルを上→下、左→右の順で返す（セル内ラベルも含む）。

        明細表の列見出し（時系列表の「対応内容」、部品表の「備考」など）は、同じラベルが表の外にもあれば後回しにする。
        """
        from excel.tables import table_header_keys  # excel.tables が SheetGrid を使うため

        label_norms = self.resolve_labels(label_norms)
        hits = {}
        for norm in label_norms:
            for cell in self._by_norm.get(norm, []) + self._by_inline.get(norm, []) + self._by_part.get(norm, []):
                hits[(cell.row, cell.col)] = cell
        if len(hits) > 1:
            headers = table_header_keys(self)
            return [hits[key] for key in sorted(hits, key=lambda k: (k in headers, k))]
        return [hits[key] for key in sorted(hits)]


def _label_like(value, text: str) -> bool:
    return isinstance(value, str) and "\n" not in text and len(text) <= MAX_LABEL_LENGTH + 4


def fill_key(xl) -> str:
    """セルの塗りつぶし色を比較用の文字列にする（ラベル欄と同じ色かの判定に使う）。"""
    fill = getattr(xl, "fill", None)
    if fill is None or fill.fill_type != "solid":
        return ""
    color = fill.fgColor
    try:
        if color.type == "theme":
            return f"theme:{color.theme}:{round(color.tint or 0, 3)}"
        return f"{color.type}:{color.value}"
    except (AttributeError, TypeError, ValueError):
        return "solid"


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
