"""テンプレート定義に従ってExcelから項目を抽出する。

セル番地を固定せず、「ラベル候補に一致するセルを探す → その右 or 下にある値を取る」
という方式なので、帳票ごとにセル位置がズレていても同じJSONになる。
"""
from __future__ import annotations

from dataclasses import asdict, dataclass

from excel.text import to_date, to_number
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.dictionary import DICTIONARY_NORMS
from pattern.model import FieldDef, PatternDef

MAX_RIGHT_STEPS = 10
MAX_TEXT_ROWS = 30


@dataclass
class FieldResult:
    field_name: str
    display_name: str
    data_type: str
    required: bool
    value: object = None
    sheet: str | None = None
    label_cell: str | None = None
    value_cell: str | None = None
    label_found: bool = False
    warning: str | None = None
    edited: bool = False
    ai_filled: bool = False
    unit: str = ""
    rag_output: str = "show"


def extract_document(info: WorkbookInfo, pattern: PatternDef, sheet_names: list[str]) -> dict:
    stop_labels = pattern.label_norms() | DICTIONARY_NORMS
    results = [_extract_field(info, fd, sheet_names, stop_labels) for fd in pattern.fields]
    extraction = {
        "pattern": {
            "id": pattern.id,
            "name": pattern.name,
            "version": pattern.version,
            "version_no": pattern.version_no,
            "image_processing": pattern.image_processing,
            "title_fields": list(pattern.title_fields),
            "md_options": dict(pattern.md_options or {}),
        },
        "sheets": sheet_names,
        "fields": [asdict(r) for r in results],
        "attachments": info.images_in(sheet_names),
    }
    refresh_summary(extraction)
    return extraction


def refresh_summary(extraction: dict) -> None:
    """fields から values / missing_required を再計算する。"""
    fields = extraction["fields"]
    extraction["values"] = {f["field_name"]: f["value"] for f in fields}
    extraction["missing_required"] = [
        f["display_name"] for f in fields if f["required"] and f["value"] in (None, "")
    ]


def apply_manual_values(extraction: dict, form) -> None:
    """プレビュー画面で人が修正した値を反映する。"""
    for f in extraction["fields"]:
        key = f"value-{f['field_name']}"
        if key not in form:
            continue
        text = form[key].replace("\r\n", "\n").strip()
        if f["data_type"] == "date" and text:
            value, warning = to_date(text, text)
        elif f["data_type"] == "number" and text:
            value, warning = to_number(text, text)
        else:
            value, warning = (text or None), None
        if value != f["value"]:
            f["value"], f["warning"], f["edited"], f["ai_filled"] = value, warning, True, False
    refresh_summary(extraction)


def _extract_field(info: WorkbookInfo, fd: FieldDef, sheet_names: list[str], stop_labels: set[str]) -> FieldResult:
    result = FieldResult(fd.field_name, fd.display_name, fd.data_type, fd.required,
                         unit=fd.unit, rag_output=fd.rag_output)
    for name in sheet_names:
        grid = info.grids.get(name)
        if grid is None:
            continue
        label, values, inline_value = locate_value(grid, fd, stop_labels)
        if label is None:
            continue
        if not result.label_found:
            result.label_found, result.sheet, result.label_cell = True, name, label.coord
        if not values:
            continue
        result.sheet, result.label_cell = name, label.coord
        result.value_cell = values[0].coord if inline_value is None else label.coord
        result.value, result.warning = _convert(fd.data_type, values, inline_value, info.date1904)
        return result
    if not result.label_found:
        result.warning = "ラベルが見つかりません"
    else:
        result.warning = "ラベルはありますが値が空です"
    return result


def locate_value(grid: SheetGrid, fd: FieldDef, stop_labels: set[str]) -> tuple[Cell | None, list[Cell], str | None]:
    """戻り値: (ラベルセル, 値セルのリスト, セル内ラベルの場合の値テキスト)"""
    norms = fd.label_norms()
    first_label = None
    for cell in grid.find_labels(norms):
        first_label = first_label or cell
        if cell.norm in norms:
            if fd.direction == "same_cell":
                continue
            values: list[Cell] = []
            if fd.direction in ("auto", "right"):
                values = scan_right(grid, cell, stop_labels)
            if not values and fd.direction in ("auto", "below"):
                values = scan_below(grid, cell, stop_labels, multi=fd.data_type == "text")
            if values:
                return cell, values, None
        elif fd.direction in ("auto", "same_cell"):
            return cell, [cell], cell.inline[1]
    return first_label, [], None


def scan_right(grid: SheetGrid, label: Cell, stop_labels: set[str]) -> list[Cell]:
    """ラベルの右側で最初に値があるセルを探す。別のラベルに当たったら値なし。"""
    row, col = label.row, label.max_col + 1
    for _ in range(MAX_RIGHT_STEPS):
        if col > grid.max_col:
            break
        top, left, _, right = grid.bounds(row, col)
        cell = grid.cells.get((top, left))
        if cell is not None:
            return [] if cell.is_label(stop_labels) else [cell]
        col = right + 1
    return []


def scan_below(grid: SheetGrid, label: Cell, stop_labels: set[str], multi: bool) -> list[Cell]:
    """ラベルの下の値を探す。文章型(multi)は空行か別ラベルまでの連続セルをまとめて取る。"""
    col, row = label.col, label.max_row + 1
    values: list[Cell] = []
    skipped_empty = False
    while row <= grid.max_row and len(values) < MAX_TEXT_ROWS:
        _, _, bottom, _ = grid.bounds(row, col)
        cell = grid.cell_at(row, col)
        if cell is None:
            if values or skipped_empty:
                break
            skipped_empty = True
        else:
            if cell.is_label(stop_labels):
                break
            values.append(cell)
            if not multi:
                break
        row = bottom + 1
    return values


def _convert(data_type: str, cells: list[Cell], inline_value: str | None, date1904: bool = False):
    if inline_value is not None:
        raw, texts = inline_value, [inline_value]
    else:
        raw, texts = cells[0].value, [c.text for c in cells]
    if data_type == "text":
        return "\n".join(texts), None
    if data_type == "date":
        return to_date(raw, texts[0], date1904)
    if data_type == "number":
        return to_number(raw, texts[0])
    return texts[0], None
