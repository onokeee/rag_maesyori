"""帳票の Markdown（docs/design.md 6.1）: タイトル・ファイル名・定型文なし・NFKC・単位・出さない項目・決定性。"""
import copy
import hashlib

import pytest

from excel.extractor import extract_document, refresh_summary
from export.formats import build_json, build_markdown, markdown_filename
from pattern.builder import suggest_rows, suggest_title_fields
from pattern.forms import rows_to_pattern
from pattern.matcher import rank_patterns

META = {"name": "設備修理報告書", "version": "v1", "description": "", "image_processing": "none"}


def _pattern(infos, **meta):
    sheets, fields = suggest_rows(infos)
    return rows_to_pattern(1, {**META, "title_fields": suggest_title_fields(fields), **meta}, sheets, fields)


def _doc(info, doc_id=1, file_hash="0123456789abcdef"):
    return {"id": doc_id, "file_name": "修理報告書_標準.xlsx" if doc_id else info.path.name, "file_hash": file_hash}


@pytest.fixture
def standard(repair_infos):
    info = repair_infos[0]
    pattern = _pattern(repair_infos)
    extraction = extract_document(info, pattern, rank_patterns(info, [pattern])[0].sheet_names)
    return info, pattern, extraction


def _field(extraction, name):
    return next(f for f in extraction["fields"] if f["field_name"] == name)


def test_title_basic_block_and_source(standard):
    info, _, extraction = standard
    md = build_markdown(_doc(info), extraction)
    lines = md.split("\n")
    assert lines[0] == "# 設備修理報告書 R2026-00123｜CMP装置（EQ-001）｜2026-09-14"
    assert "- 帳票の種類: 設備修理報告書" in lines
    assert "- 報告番号: R2026-00123" in lines
    assert "- 設備: CMP装置（EQ-001）" in lines
    assert "- 発生日: 2026-09-14（2026年9月）" in lines
    assert "## 故障内容" in lines  # 短い帳票は見出しに識別子を入れない
    assert "- 添付画像: 1枚" in lines
    assert lines[-2] == "- 出典: 修理報告書_標準.xlsx（報告番号 R2026-00123）"


def test_title_falls_back_to_identifier_keys_when_not_configured(repair_infos):
    info = repair_infos[2]
    pattern = _pattern(repair_infos, title_fields=[])
    extraction = extract_document(info, pattern, rank_patterns(info, [pattern])[0].sheet_names)
    md = build_markdown(_doc(info), extraction)
    assert md.startswith("# 設備修理報告書 R2026-00125｜露光装置（EQ-003）｜2026-08-30\n")


def test_configured_title_fields_are_used_in_order(standard):
    info, _, extraction = standard
    extraction["pattern"]["title_fields"] = ["equipment_id", "occurred_date"]
    md = build_markdown(_doc(info), extraction)
    assert md.startswith("# 設備修理報告書 EQ-001｜2026-09-14\n")
    assert markdown_filename(_doc(info), extraction) == "設備修理報告書_EQ-001_2026-09-14.md"


def test_no_boilerplate_internal_ids_or_cell_coordinates(standard):
    info, _, extraction = standard
    md = build_markdown(_doc(info, doc_id=42), extraction)
    for banned in ("文書ID", "42", "解析していません", "## 原本", "## 基本情報", "テンプレート", "v1", "J3",
                   "修理報告書シート", "A8", "## 添付画像", "原本を参照"):
        assert banned not in md, banned


def test_filename_is_stable_and_independent_of_document_id(standard):
    info, _, extraction = standard
    a = markdown_filename(_doc(info, doc_id=1), extraction)
    b = markdown_filename(_doc(info, doc_id=999, file_hash="ffffffffffffffff"), copy.deepcopy(extraction))
    assert a == b == "設備修理報告書_R2026-00123_EQ-001_CMP装置_2026-09-14.md"


def test_filename_uses_file_hash_when_title_values_are_empty(standard):
    info, _, extraction = standard
    for name in ("report_id", "equipment_id", "equipment_name", "occurred_date"):
        _field(extraction, name)["value"] = None
    refresh_summary(extraction)
    assert markdown_filename(_doc(info), extraction) == "設備修理報告書_01234567.md"


def test_filename_sanitizes_hint_brackets_and_unsafe_chars(standard):
    info, _, extraction = standard
    extraction["pattern"]["name"] = "修理報告書.[legacy-R(chunk_ts=800)]"
    extraction["pattern"]["title_fields"] = ["report_id"]
    _field(extraction, "report_id")["value"] = "R/2026:00123 [改]\t*?"
    name = markdown_filename(_doc(info), extraction)
    assert name.endswith(".md") and name.count(".") == 1
    for ch in ('.[', '[', ']', '/', '\\', ':', '*', '?', '"', '<', '>', '|', ' ', '\t'):
        assert ch not in name, (ch, name)
    assert name.startswith("修理報告書")
    # 全角英数は NFKC で半角に
    _field(extraction, "report_id")["value"] = "Ｒ２０２６－００１２３"
    assert "R2026-00123" in markdown_filename(_doc(info), extraction)


def test_values_are_nfkc_normalized(standard):
    info, _, extraction = standard
    _field(extraction, "equipment_name")["value"] = "ＣＭＰ　　装置"
    _field(extraction, "equipment_id")["value"] = "ＥＱ－００１"
    _field(extraction, "cause")["value"] = "ﾎﾟﾝﾌﾟの  ｺﾈｸﾀ緩み。\n\n\n再締結した。"
    md = build_markdown(_doc(info), extraction)
    assert md.startswith("# 設備修理報告書 R2026-00123｜CMP 装置（EQ-001）｜2026-09-14\n")
    assert "- 設備: CMP 装置（EQ-001）" in md
    assert "## 原因\nポンプの コネクタ緩み。\n再締結した。\n" in md  # レコード内に空行を入れない


def test_units_and_ai_mark(standard):
    info, _, extraction = standard
    assert "- 作業時間: 2.5時間" in build_markdown(_doc(info), extraction)
    work = _field(extraction, "work_hours")
    work["unit"], work["value"], work["ai_filled"] = "h", 3.0, True
    md = build_markdown(_doc(info), extraction)
    assert "- 作業時間: 3h（AI入力）" in md
    work["value"] = "3時間くらい"  # 数値にできなかった値には単位を足さない
    assert "- 作業時間: 3時間くらい（AI入力）" in build_markdown(_doc(info), extraction)


def test_person_fields_are_omitted_by_default(standard):
    info, _, extraction = standard
    md = build_markdown(_doc(info), extraction)
    assert "報告者" not in md and "山田" not in md
    extraction["pattern"]["md_options"] = {"omit_person_fields": False}
    assert "- 報告者: 山田 太郎" in build_markdown(_doc(info), extraction)


def test_rag_output_omit_fields_are_not_written(standard):
    info, _, extraction = standard
    _field(extraction, "cause")["rag_output"] = "omit"
    _field(extraction, "work_hours")["rag_output"] = "omit"
    md = build_markdown(_doc(info), extraction)
    assert "## 原因" not in md and "コネクタ接触不良" not in md and "作業時間" not in md
    # JSON 側には残す
    data = build_json(_doc(info), extraction)
    cause = next(f for f in data["fields"] if f["field_name"] == "cause")
    assert cause["rag_output"] == "omit" and cause["value"]
    assert next(f for f in data["fields"] if f["field_name"] == "work_hours")["unit"] == "時間"


def test_long_documents_get_identifier_headings(standard):
    info, _, extraction = standard
    _field(extraction, "symptom")["value"] = "搬送アームが停止した。\n" * 120
    md = build_markdown(_doc(info), extraction)
    assert "## 故障内容（R2026-00123／EQ-001 CMP装置／2026-09-14）" in md
    assert "## 原因（R2026-00123／EQ-001 CMP装置／2026-09-14）" in md


def test_markdown_syntax_in_values_is_escaped(standard):
    info, _, extraction = standard
    _field(extraction, "cause")["value"] = "# 見出しではない\n---\n> 引用ではない"
    md = build_markdown(_doc(info), extraction)
    assert "\n# 見出し" not in md and "\n---\n" not in md and "\n> 引用" not in md
    assert md.count("\n# ") == 0 and md.startswith("# ")


def test_output_is_deterministic_bytes(repair_infos):
    digests = set()
    for _ in range(2):
        info = repair_infos[1]
        pattern = _pattern(repair_infos)
        extraction = extract_document(info, pattern, rank_patterns(info, [pattern])[0].sheet_names)
        md = build_markdown(_doc(info, doc_id=None), extraction)
        digests.add(hashlib.sha256(md.encode("utf-8")).hexdigest())
        assert "\r" not in md and md.endswith("\n") and not md.endswith("\n\n") and "\n\n\n" not in md
        assert md.startswith("# 設備修理報告書 R2026-00124｜CVD装置（EQ-002）｜2026-09-10\n")
        assert "- 添付画像: 2枚" in md
    assert len(digests) == 1


def test_local_fallback_matches_core_helpers(standard, monkeypatch):
    """core/mdtext・core/naming が無い環境でも同じ md とファイル名になる。"""
    from export import formats

    if formats._core_mdtext is None or formats._core_naming is None:
        pytest.skip("core が未導入")
    info, _, extraction = standard
    _field(extraction, "cause")["value"] = "# 見出し\n1. 手順\n- 箇条\n---\n> 引用"
    _field(extraction, "report_id")["value"] = "R/2026 [改].[x]"
    extraction["pattern"]["md_options"] = {"omit_person_fields": False}
    with_core = (build_markdown(_doc(info), extraction), markdown_filename(_doc(info), extraction))
    monkeypatch.setattr(formats, "_core_mdtext", None)
    monkeypatch.setattr(formats, "_core_naming", None)
    assert (build_markdown(_doc(info), extraction), markdown_filename(_doc(info), extraction)) == with_core
