"""Excel(.xlsx/.xlsm) の読み込み。openpyxl の通常モード（data_only=True）で読む。

結合セル・行の非表示・太字/取り消し線/塗りを CellInfo に載せる。
通常モードはシート全体をメモリに読むので、セル数に上限を設ける。
"""
from __future__ import annotations

import re
import warnings
import zipfile
from pathlib import Path
from typing import Iterator

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, range_boundaries
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from tables.source import CellInfo, SheetInfo, SourceRow, UploadError, cell_text

DEFAULT_MAX_CELLS = 500_000
_DIMENSION_RE = re.compile(rb'<dimension ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"')


def _too_many_cells(count: int, limit: int) -> UploadError:
    return UploadError(
        f"セル数が上限（{limit:,}セル）を超えています（約{count:,}セル）。"
        "CSVで保存してから取り込んでください（書式による判定は使えなくなります）"
    )


def estimate_cell_count(path) -> int:
    """シートXMLの <dimension> から、読み込み前にセル数の目安を出す。"""
    total = 0
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
                    continue
                with zf.open(name) as f:
                    head = f.read(4096)
                m = _DIMENSION_RE.search(head)
                if not m:
                    continue
                ref = m.group(1).decode()
                if ":" not in ref:
                    total += 1
                    continue
                min_col, min_row, max_col, max_row = range_boundaries(ref)
                total += (max_row - min_row + 1) * (max_col - min_col + 1)
    except (zipfile.BadZipFile, OSError, ValueError):
        return 0
    return total


class ExcelSource:
    kind = "excel"

    def __init__(self, path, file_name: str | None = None, options: dict | None = None):
        self.path = Path(path)
        self.file_name = file_name or self.path.name
        options = dict(options or {})
        self.max_cells = int(options.get("max_cells") or DEFAULT_MAX_CELLS)

        # 範囲情報（dimension）は書式だけの行で大きく出ることがあるので、目安の4倍を超えるときだけ先に止める
        estimated = estimate_cell_count(self.path)
        if estimated > self.max_cells * 4:
            raise _too_many_cells(estimated, self.max_cells)
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                self.wb = load_workbook(self.path, data_only=True)
        except Exception as e:  # openpyxl は形式ごとに様々な例外を出す
            raise UploadError("Excelファイルとして開けませんでした。.xlsx 形式で保存し直してください") from e

        actual = sum(len(ws._cells) for ws in self.wb.worksheets)
        if actual > self.max_cells:
            raise _too_many_cells(actual, self.max_cells)
        self.date1904 = self.wb.epoch == CALENDAR_MAC_1904
        self._cache: dict[str, dict] = {}

    def close(self) -> None:
        pass

    def sheets(self) -> list[SheetInfo]:
        result = []
        for ws in self.wb.worksheets:
            meta = self._meta(ws.title)
            result.append(SheetInfo(
                name=ws.title,
                hidden=ws.sheet_state != "visible",
                max_row=meta["max_row"],
                max_col=meta["max_col"],
                table_ranges=[t.ref for t in ws.tables.values()],
                date1904=self.date1904,
                hidden_columns=meta["hidden_columns"],
                freeze_panes=ws.freeze_panes,
                auto_filter=ws.auto_filter.ref or None,
            ))
        return result

    def _meta(self, sheet: str) -> dict:
        if sheet in self._cache:
            return self._cache[sheet]
        if sheet not in self.wb.sheetnames:
            raise UploadError(f"シート「{sheet}」が見つかりません")
        ws = self.wb[sheet]
        max_row = max_col = 0
        for (r, c), cell in ws._cells.items():
            if cell.value is not None and cell.value != "":
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        anchors: dict[tuple[int, int], tuple[int, int]] = {}
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    anchors[(r, c)] = (rng.min_row, rng.min_col)
        hidden_columns: list[int] = []
        for key, dim in ws.column_dimensions.items():
            if dim.hidden:
                lo = dim.min or column_index_from_string(key)
                hi = dim.max or lo
                hidden_columns.extend(range(lo, hi + 1))
        meta = {"ws": ws, "max_row": max_row, "max_col": max_col, "anchors": anchors,
                "hidden_columns": sorted(set(hidden_columns))}
        self._cache[sheet] = meta
        return meta

    def rows(self, sheet: str, start: int = 1, limit: int | None = None) -> Iterator[SourceRow]:
        meta = self._meta(sheet)
        ws = meta["ws"]
        cells = ws._cells
        anchors = meta["anchors"]
        row_dims = ws.row_dimensions
        max_col = meta["max_col"]
        end = meta["max_row"] if limit is None else min(meta["max_row"], start + limit - 1)
        for r in range(max(1, start), end + 1):
            row_cells = []
            for c in range(1, max_col + 1):
                xl = cells.get((r, c))
                anchor = anchors.get((r, c))
                if xl is None:
                    row_cells.append(CellInfo(value=None, text="", merged_anchor=anchor))
                    continue
                value = xl.value
                font = xl.font
                fill = xl.fill
                row_cells.append(CellInfo(
                    value=value,
                    text=cell_text(value),
                    number_format=xl.number_format,
                    bold=bool(font is not None and font.b),
                    strike=bool(font is not None and font.strike),
                    fill=bool(fill is not None and fill.fill_type not in (None, "none")),
                    merged_anchor=anchor,
                ))
            dim = row_dims.get(r)
            yield SourceRow(index=r, cells=row_cells, hidden=bool(dim is not None and dim.hidden))
