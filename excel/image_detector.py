"""Excel内の画像の「存在と位置」だけを検出する（画像の内容は解析しない）。

openpyxlの画像読み込みはPillowに依存し、グループ化された図などを取りこぼすため、
xlsx(zip)内の drawing XML を直接読む。
"""
from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl.utils import get_column_letter

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
ANCHOR_TAGS = {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}


def detect_images(path: str | Path) -> list[dict]:
    """戻り値: [{"type": "image", "sheet": "...", "location": "H10:M20", "name": "..."}]"""
    results: list[dict] = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        workbook_part = next(
            (target for rtype, target in _rels(zf, names, "").values() if rtype.endswith("/officeDocument")),
            "xl/workbook.xml",
        )
        if workbook_part not in names:
            return results
        workbook = ET.fromstring(zf.read(workbook_part))
        workbook_rels = _rels(zf, names, workbook_part)
        # 複数のシートが同じ drawing を指すことがあるので、drawing ごとに1回だけ読む
        parsed: dict[str, ET.Element] = {}

        for sheet in workbook.findall("main:sheets/main:sheet", NS):
            sheet_part = workbook_rels.get(sheet.get(R_ID), ("", ""))[1]
            if sheet_part not in names:
                continue
            for rtype, drawing_part in _rels(zf, names, sheet_part).values():
                if not rtype.endswith("/drawing") or drawing_part not in names:
                    continue
                if drawing_part not in parsed:
                    parsed[drawing_part] = ET.fromstring(zf.read(drawing_part))
                results.extend(_images_in_drawing(parsed[drawing_part], sheet.get("name")))
    return results


def _rels(zf: zipfile.ZipFile, names: set[str], part: str) -> dict[str, tuple[str, str]]:
    """パーツのリレーションを {rId: (type, 解決済みパス)} で返す。"""
    folder, filename = posixpath.split(part)
    rels_path = posixpath.join(folder, "_rels", filename + ".rels")
    if rels_path not in names:
        return {}
    out = {}
    for rel in ET.fromstring(zf.read(rels_path)).findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = (rel.get("Type", ""), resolved)
    return out


def _images_in_drawing(root: ET.Element, sheet_name: str) -> list[dict]:
    images = []
    for anchor in root:
        if anchor.tag.rsplit("}", 1)[-1] not in ANCHOR_TAGS:
            continue
        location = _anchor_location(anchor)
        for pic in anchor.findall(".//xdr:pic", NS):
            props = pic.find("xdr:nvPicPr/xdr:cNvPr", NS)
            images.append({
                "type": "image",
                "sheet": sheet_name,
                "location": location,
                "name": props.get("name", "") if props is not None else "",
            })
    return images


def _anchor_location(anchor: ET.Element) -> str:
    start = _marker(anchor.find("xdr:from", NS))
    end = _marker(anchor.find("xdr:to", NS))
    if start and end and end != start:
        return f"{start}:{end}"
    return start or ""


def _marker(marker: ET.Element | None) -> str:
    if marker is None:
        return ""
    col, row = marker.find("xdr:col", NS), marker.find("xdr:row", NS)
    if col is None or row is None:
        return ""
    return f"{get_column_letter(int(col.text) + 1)}{int(row.text) + 1}"
