"""帳票の種類の編集フォームと PatternDef の相互変換。"""
from __future__ import annotations

import re

from pattern.model import (DATA_TYPES, DEFAULT_MD_OPTIONS, DIRECTIONS, IMAGE_PROCESSING, RAG_OUTPUTS, FieldDef,
                           PatternDef, SheetDef)

FIELD_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def pattern_to_rows(pattern: PatternDef) -> tuple[list[dict], list[dict]]:
    sheet_rows = [{"use": True, "sheet_name": s.sheet_name, "required": s.required} for s in pattern.sheets]
    field_rows = [
        {
            "use": True,
            "field_name": f.field_name,
            "display_name": f.display_name,
            "candidates": "\n".join(f.candidates),
            "data_type": f.data_type,
            "required": f.required,
            "direction": f.direction,
            "unit": f.unit,
            "rag_output": f.rag_output,
            "table_columns": "\n".join(f.table_columns),
            "section": f.section,
            "sheet_name": getattr(f, "sheet_name", "") or "",
            "label_cell": getattr(f, "label_cell", "") or "",
            "cell": getattr(f, "cell", "") or "",
        }
        for f in pattern.fields
    ]
    return sheet_rows, field_rows


def pattern_to_meta(pattern: PatternDef) -> dict:
    """編集画面の初期値用のメタ情報（parse_pattern_form の meta と同じ形）。"""
    return {
        "name": pattern.name,
        "version": pattern.version,
        "description": pattern.description,
        "image_processing": pattern.image_processing,
        "title_fields": list(pattern.title_fields),
        "md_options": {**DEFAULT_MD_OPTIONS, **(pattern.md_options or {})},
        "version_no": pattern.version_no,
    }


def parse_pattern_form(form) -> tuple[dict, list[dict], list[dict], list[str]]:
    """戻り値: (メタ情報, シート行, 項目行, エラー)

    タイトル項目は title_fields（複数値、またはカンマ・改行区切り）で受け取る。
    """
    errors: list[str] = []
    meta = {
        "name": form.get("name", "").strip(),
        "version": form.get("version", "").strip() or "v1",
        "description": form.get("description", "").strip(),
        "image_processing": form.get("image_processing", "none"),
        "title_fields": _get_list(form, "title_fields"),
        "md_options": dict(DEFAULT_MD_OPTIONS),
    }
    if not meta["name"]:
        errors.append("帳票の種類の名前を入力してください")
    if meta["image_processing"] not in IMAGE_PROCESSING:
        meta["image_processing"] = "none"

    sheet_rows = []
    for i in _row_indices(form, "sheets"):
        p = f"sheets-{i}-"
        name = form.get(p + "sheet_name", "").strip()
        if name:
            sheet_rows.append({"use": form.get(p + "use") == "on", "sheet_name": name,
                               "required": form.get(p + "required") == "on"})

    field_rows = []
    for i in _row_indices(form, "fields"):
        p = f"fields-{i}-"
        row = {
            "use": form.get(p + "use") == "on",
            "field_name": form.get(p + "field_name", "").strip(),
            "display_name": form.get(p + "display_name", "").strip(),
            "candidates": "\n".join(_split_candidates(form.get(p + "candidates", ""))),
            "data_type": form.get(p + "data_type", "string"),
            "required": form.get(p + "required") == "on",
            "direction": form.get(p + "direction", "auto"),
            "unit": form.get(p + "unit", "").strip(),
            "rag_output": form.get(p + "rag_output", "show"),
            "table_columns": "\n".join(_split_candidates(form.get(p + "table_columns", ""))),
            "section": _section_value(form.get(p + "section", "")),
        }
        if not row["display_name"] and not row["candidates"]:
            continue
        if row["data_type"] not in DATA_TYPES:
            row["data_type"] = "string"
        if row["direction"] not in DIRECTIONS:
            row["direction"] = "auto"
        if row["rag_output"] not in RAG_OUTPUTS:
            row["rag_output"] = "show"
        field_rows.append(row)

    used_fields = [r for r in field_rows if r["use"]]
    if not used_fields:
        errors.append("読み取る項目を1つ以上選択してください")
    seen_names: set[str] = set()
    for row in used_fields:
        if not row["display_name"]:
            errors.append(f"項目名が空の行があります（候補: {row['candidates'].splitlines()[0]}）")
        if not row["field_name"]:
            errors.append(f"「{row['display_name']}」のキー名を入力してください")
        elif not FIELD_NAME_RE.match(row["field_name"]):
            errors.append(f"キー名「{row['field_name']}」は半角英数字と _ で入力してください")
        elif row["field_name"] in seen_names:
            errors.append(f"キー名「{row['field_name']}」が重複しています")
        seen_names.add(row["field_name"])
    # 使わない項目を指したタイトル項目は黙って外す
    meta["title_fields"] = [name for name in meta["title_fields"] if name in seen_names]
    return meta, sheet_rows, field_rows, errors


def rows_to_pattern(pattern_id: int, meta: dict, sheet_rows: list[dict], field_rows: list[dict]) -> PatternDef:
    fields = [
        FieldDef(
            field_name=r["field_name"],
            display_name=r["display_name"],
            # 「値だけ」の項目（クリックで作った、見出しの無い項目）は探す見出しを持たない
            candidates=r["candidates"].splitlines() or ([] if r.get("cell") else [r["display_name"]]),
            required=r["required"],
            data_type=r["data_type"],
            direction=r["direction"],
            unit=r.get("unit", "") or "",
            rag_output=r.get("rag_output", "show") if r.get("rag_output") in RAG_OUTPUTS else "show",
            table_columns=(r.get("table_columns") or "").splitlines() if r["data_type"] == "table" else [],
            section=_section_value(r.get("section")),
            sheet_name=r.get("sheet_name", "") or "",
            label_cell=r.get("label_cell", "") or "",
            cell=r.get("cell", "") or "",
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
        sheets=[SheetDef(r["sheet_name"], r["required"]) for r in sheet_rows if r["use"]],
        fields=fields,
        title_fields=[n for n in dict.fromkeys(meta.get("title_fields") or []) if n in names],
        md_options={**DEFAULT_MD_OPTIONS, **(meta.get("md_options") or {})},
        version_no=int(meta.get("version_no") or 1),
    )


def _get_list(form, key: str) -> list[str]:
    values = form.getlist(key) if hasattr(form, "getlist") else form.get(key, [])
    if isinstance(values, str):
        values = [values]
    items: list[str] = []
    for value in values or []:
        items.extend(re.split(r"[\n,、]", str(value).replace("\r", "")))
    return list(dict.fromkeys(i.strip() for i in items if i.strip()))


def _row_indices(form, prefix: str) -> list[int]:
    indices = set()
    for key in form.keys():
        parts = key.split("-")
        if len(parts) == 3 and parts[0] == prefix and parts[1].isdigit():
            indices.add(int(parts[1]))
    return sorted(indices)


def _section_value(text) -> str:
    """探す区画（「回答欄」「▼ 回答欄」のどちらで入力しても、見出しと同じ比較用の名前にする）。"""
    from excel.tables import section_name

    return section_name(text) if str(text or "").strip() else ""


def _split_candidates(text: str) -> list[str]:
    items = re.split(r"[\n,、]", text.replace("\r", ""))
    return list(dict.fromkeys(i.strip() for i in items if i.strip()))
