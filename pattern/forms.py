"""帳票の種類の編集画面の行（dict）と PatternDef の相互変換。"""
from __future__ import annotations

from pattern.model import DEFAULT_MD_OPTIONS, RAG_OUTPUTS, FieldDef, PatternDef, SheetDef


def pattern_to_rows(pattern: PatternDef) -> tuple[list[dict], list[dict]]:
    sheet_rows = [{"use": True, "sheet_name": s.sheet_name} for s in pattern.sheets]
    field_rows = [
        {
            "use": True,
            "field_name": f.field_name,
            "display_name": f.display_name,
            "candidates": "\n".join(f.candidates),
            "data_type": f.data_type,
            "direction": f.direction,
            "unit": f.unit,
            "rag_output": f.rag_output,
            "table_columns": "\n".join(f.table_columns),
            "section": f.section,
            "sheet_name": getattr(f, "sheet_name", "") or "",
            "label_cell": getattr(f, "label_cell", "") or "",
            "cell": getattr(f, "cell", "") or "",
            "renamed": bool(getattr(f, "renamed", False)),
        }
        for f in pattern.fields
    ]
    return sheet_rows, field_rows


def pattern_to_meta(pattern: PatternDef) -> dict:
    """編集画面の初期値用のメタ情報（rows_to_pattern に渡す meta と同じ形）。"""
    return {
        "name": pattern.name,
        "version": pattern.version,
        "description": pattern.description,
        "image_processing": pattern.image_processing,
        "title_fields": list(pattern.title_fields),
        "md_options": {**DEFAULT_MD_OPTIONS, **(pattern.md_options or {})},
        "version_no": pattern.version_no,
    }


def rows_to_pattern(pattern_id: int, meta: dict, sheet_rows: list[dict], field_rows: list[dict]) -> PatternDef:
    fields = [
        FieldDef(
            field_name=r["field_name"],
            display_name=r["display_name"],
            # 「値だけ」の項目（クリックで作った、見出しの無い項目）は探す見出しを持たない
            candidates=r["candidates"].splitlines() or ([] if r.get("cell") else [r["display_name"]]),
            data_type=r["data_type"],
            direction=r["direction"],
            unit=r.get("unit", "") or "",
            rag_output=r.get("rag_output", "show") if r.get("rag_output") in RAG_OUTPUTS else "show",
            table_columns=(r.get("table_columns") or "").splitlines() if r["data_type"] == "table" else [],
            section=_section_value(r.get("section")),
            sheet_name=r.get("sheet_name", "") or "",
            label_cell=r.get("label_cell", "") or "",
            cell=r.get("cell", "") or "",
            renamed=bool(r.get("renamed")),
        )
        for r in field_rows
        if r["use"]
    ]
    names = {f.field_name for f in fields}
    return PatternDef(
        id=pattern_id,
        name=meta["name"],
        version=meta.get("version", "v1"),
        description=meta.get("description", ""),
        image_processing=meta.get("image_processing", "none"),
        sheets=[SheetDef(r["sheet_name"]) for r in sheet_rows if r["use"]],
        fields=fields,
        title_fields=[n for n in dict.fromkeys(meta.get("title_fields") or []) if n in names],
        md_options={**DEFAULT_MD_OPTIONS, **(meta.get("md_options") or {})},
        version_no=int(meta.get("version_no") or 1),
    )


def _section_value(text) -> str:
    """探す区画（「回答欄」「▼ 回答欄」のどちらで入力しても、見出しと同じ比較用の名前にする）。"""
    from excel.tables import section_name

    return section_name(text) if str(text or "").strip() else ""
