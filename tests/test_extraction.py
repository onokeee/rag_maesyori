from excel.extractor import extract_document
from excel.text import normalize_label, to_date, to_number
from excel.workbook import load_workbook_info
from pattern.builder import suggest_rows
from pattern.forms import rows_to_pattern
from pattern.matcher import rank_patterns

META = {"name": "設備修理報告書", "version": "v1", "description": "", "image_processing": "none"}


def build_repair_pattern(infos):
    sheets, fields = suggest_rows(infos)
    return rows_to_pattern(1, META, sheets, fields)


def test_normalize_label_absorbs_notation_differences():
    assert normalize_label("設備№：") == normalize_label(" 設備 No. ") == normalize_label("【設備NO】") == "設備no"


def test_value_conversion():
    assert to_date("2026/9/1", "2026/9/1") == ("2026-09-01", None)
    assert to_date("令和8年9月14日", "令和8年9月14日") == ("2026-09-14", None)
    assert to_number("1.5時間", "1.5時間")[0] == 1.5
    assert to_number(3.0, "3") == (3, None)


def test_builder_suggests_standard_fields_from_samples(repair_infos):
    sheets, fields = suggest_rows(repair_infos)
    used_sheets = [s["sheet_name"] for s in sheets if s["use"]]
    assert used_sheets == ["修理報告書"]  # マスタ・参考資料・(2) の重複は選ばない

    used = {f["field_name"] for f in fields if f["use"]}
    assert {"report_id", "equipment_id", "equipment_name", "occurred_date", "reporter",
            "symptom", "cause", "repair", "result"} <= used
    equipment = next(f for f in fields if f["field_name"] == "equipment_id")
    assert equipment["seen"] == "3/3"
    assert "装置番号" in equipment["candidates"].splitlines()


def test_same_json_from_different_layouts(repair_infos):
    pattern = build_repair_pattern(repair_infos)
    expected = [
        {"report_id": "R2026-00123", "equipment_id": "EQ-001", "occurred_date": "2026-09-14", "work_hours": 2.5},
        {"report_id": "R2026-00124", "equipment_id": "EQ-002", "occurred_date": "2026-09-10", "work_hours": 1.5,
         "equipment_name": "CVD装置"},
        {"report_id": "R2026-00125", "equipment_id": "EQ-003", "occurred_date": "2026-08-30", "cause": "リニアガイドの潤滑不足。"},
    ]
    for info, values in zip(repair_infos, expected):
        match = rank_patterns(info, [pattern])[0]
        extraction = extract_document(info, pattern, match.sheet_names)
        for key, value in values.items():
            assert extraction["values"][key] == value, (info.path.name, key)
        assert "Robot Position Error" in extraction["values"]["symptom"] or info is not repair_infos[0]


def test_images_are_detected_with_location(repair_infos):
    standard, shifted, table = repair_infos
    assert [(i["sheet"], i["location"]) for i in standard.images] == [("修理報告書", "J3")]
    assert len(shifted.images) == 2
    assert table.images == []


def test_matcher_prefers_correct_pattern(repair_infos, sample_dir):
    inspection_info = load_workbook_info(sample_dir / "inspection.xlsx")
    repair = build_repair_pattern(repair_infos)
    sheets, fields = suggest_rows([inspection_info])
    inspection = rows_to_pattern(2, {**META, "name": "点検記録表"}, sheets, fields)

    assert rank_patterns(repair_infos[0], [inspection, repair])[0].pattern.name == "設備修理報告書"
    ranked = rank_patterns(inspection_info, [repair, inspection])
    assert ranked[0].pattern.name == "点検記録表"
    assert ranked[1].confidence < 50


# ---- 種類定義の拡張（単位・Markdownでの扱い・タイトル項目） ----

def test_to_date_supports_1904_workbooks():
    assert to_date(45000, "45000") == ("2023-03-15", None)
    assert to_date(45000 - 1462, "43538", date1904=True) == ("2023-03-15", None)


def test_split_label_unit_and_value_unit():
    from excel.text import nfkc_value, split_label_unit, value_unit

    assert split_label_unit("作業時間(h)") == ("作業時間", "時間")
    assert split_label_unit("停止時間（分）") == ("停止時間", "分")
    assert split_label_unit("設備番号") == ("設備番号", "")
    assert value_unit("1.5時間") == "時間" and value_unit("EQ-001") == ""
    assert nfkc_value(" ＣＭＰ　　装置 \r\nﾎﾟﾝﾌﾟ") == "CMP 装置\nポンプ"
    # ラベル比較用の正規化は従来どおり空白を消す
    assert normalize_label("ＣＭＰ　装置") == "cmp装置"


def test_builder_suggests_units_and_title_fields(repair_infos, tmp_path):
    from openpyxl import Workbook
    from pattern.builder import suggest_title_fields

    sheets, fields = suggest_rows(repair_infos)
    work_hours = next(f for f in fields if f["field_name"] == "work_hours")
    assert work_hours["unit"] == "時間"  # 値「1.5時間」から
    assert all(f["rag_output"] == "show" for f in fields)
    assert suggest_title_fields(fields) == ["report_id", "equipment_id", "equipment_name", "occurred_date"]

    wb = Workbook()
    ws = wb.active
    ws.title = "報告書"
    ws["A1"], ws["B1"] = "報告番号", "R-1"
    ws["A2"], ws["B2"] = "作業時間(h)", 3
    ws["A3"], ws["B3"] = "停止時間（分）", 45
    wb.save(tmp_path / "unit.xlsx")
    _, rows = suggest_rows([load_workbook_info(tmp_path / "unit.xlsx")])
    by_name = {r["field_name"]: r for r in rows}
    assert by_name["work_hours"]["unit"] == "時間" and by_name["work_hours"]["display_name"] == "作業時間"
    assert "作業時間(h)" in by_name["work_hours"]["candidates"].splitlines()
    assert by_name["downtime"]["unit"] == "分"


def test_form_rows_roundtrip_unit_rag_output_and_title_fields():
    from pattern.forms import parse_pattern_form, pattern_to_meta, pattern_to_rows

    form = {
        "name": "設備修理報告書", "version": "v1", "title_fields": "report_id, equipment_id, unknown",
        "md_options_form": "1",
        "fields-0-use": "on", "fields-0-field_name": "report_id", "fields-0-display_name": "報告番号",
        "fields-0-candidates": "報告番号", "fields-0-data_type": "string", "fields-0-direction": "auto",
        "fields-1-use": "on", "fields-1-field_name": "equipment_id", "fields-1-display_name": "設備番号",
        "fields-1-candidates": "設備番号", "fields-1-data_type": "string", "fields-1-direction": "auto",
        "fields-2-use": "on", "fields-2-field_name": "work_hours", "fields-2-display_name": "作業時間",
        "fields-2-candidates": "作業時間", "fields-2-data_type": "number", "fields-2-direction": "auto",
        "fields-2-unit": "時間", "fields-2-rag_output": "omit",
    }
    meta, sheet_rows, field_rows, errors = parse_pattern_form(form)
    assert errors == []
    assert meta["title_fields"] == ["report_id", "equipment_id"]
    assert meta["md_options"]["omit_person_fields"] is False  # チェックなしで送信された
    pattern = rows_to_pattern(3, meta, sheet_rows, field_rows)
    work = pattern.fields[2]
    assert (work.unit, work.rag_output) == ("時間", "omit")
    assert pattern.title_fields == ["report_id", "equipment_id"] and pattern.version_no == 1
    _, rows = pattern_to_rows(pattern)
    assert rows[2]["unit"] == "時間" and rows[2]["rag_output"] == "omit"
    assert pattern_to_meta(pattern)["title_fields"] == ["report_id", "equipment_id"]

    meta2, _, _, _ = parse_pattern_form({k: v for k, v in form.items() if k != "md_options_form"})
    assert meta2["md_options"]["omit_person_fields"] is True  # 設定欄のない画面からは既定のまま


def test_extraction_carries_unit_rag_output_and_title_fields(repair_infos):
    pattern = build_repair_pattern(repair_infos)
    pattern.title_fields = ["report_id", "occurred_date"]
    pattern.version_no = 4
    pattern.fields[0].rag_output = "omit"
    info = repair_infos[0]
    extraction = extract_document(info, pattern, rank_patterns(info, [pattern])[0].sheet_names)
    assert extraction["pattern"]["title_fields"] == ["report_id", "occurred_date"]
    assert extraction["pattern"]["version_no"] == 4
    work = next(f for f in extraction["fields"] if f["field_name"] == "work_hours")
    assert work["unit"] == "時間" and work["rag_output"] == "show"
    assert extraction["fields"][0]["rag_output"] == "omit"


# ---- 一覧表らしさ ----

def test_looks_like_table_sheet(repair_infos, sample_dir, tmp_path):
    from openpyxl import Workbook
    from pattern.matcher import looks_like_table_sheet, table_like_sheets

    for info in repair_infos + [load_workbook_info(sample_dir / "inspection.xlsx")]:
        assert not looks_like_table_sheet(info), info.path.name

    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    ws["A1"] = "トラブル対応一覧"
    ws.append([])
    ws.append(["管理No", "発生日", "設備番号", "現象", "停止時間(分)"])
    for i in range(12):
        ws.append([f"TR-{i:03d}", f"2026-08-{i + 1:02d}", "EQ-001", "停止" if i % 3 else None, i * 5])
    notes = wb.create_sheet("記入要領")
    notes["A1"] = "この様式の書き方"
    wb.save(tmp_path / "list.xlsx")
    info = load_workbook_info(tmp_path / "list.xlsx")
    assert looks_like_table_sheet(info)
    assert table_like_sheets(info) == ["一覧"]
    assert looks_like_table_sheet(info.grids["一覧"]) and not looks_like_table_sheet(info, ["記入要領"])

    # 9行しかない表は一覧表とみなさない
    ws.delete_rows(13, 3)
    wb.save(tmp_path / "short.xlsx")
    assert not looks_like_table_sheet(load_workbook_info(tmp_path / "short.xlsx"))
