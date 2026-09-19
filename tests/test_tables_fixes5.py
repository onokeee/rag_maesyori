"""一覧表の不具合修正（5巡目）の確認。"""
from __future__ import annotations

import io
import json
import time
from datetime import timedelta
from pathlib import Path

import pytest
from openpyxl import Workbook

from tables import pipeline, store
from tables.markdown import _Names, record_block, record_title, render_all
from tables.normalize import ConvertContext, convert_cell, read_records
from tables.spec import ColumnSpec, spec_from_dict, validate_spec
from tests.test_tables_pipe import FakeCtx, _new_import, _rec, list_spec_dict


# ---- R5T-1: 大文字・小文字だけ違う設備のファイル名が Windows で上書きし合わない ------------------------------

def test_names_differing_only_in_case_get_a_suffix():
    names = _Names()
    a = names.make(["a", "ETC-302号機"])
    b = names.make(["a", "Etc-302号機"])
    assert a.casefold() != b.casefold()


def test_entities_differing_only_in_case_keep_every_record(tmp_path):
    d = list_spec_dict(group_by="entity_month")
    for c in d["columns"]:
        if c["key"] == "equipment_id":
            c["normalize"] = ["nfkc"]   # 大文字にそろえない設定
    d["columns"] = [c for c in d["columns"] if c["key"] != "equipment_name"]
    spec = spec_from_dict(d)
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="ETC-302号機"),
               _rec("A-2", record_no="A-2", occurred_at="2026-08-04", equipment_id="Etc-302号機")]
    files = render_all(spec, records, {}, {})
    assert len({f.name.casefold() for f in files}) == len(files)
    target = tmp_path / "imp"
    target.mkdir()
    pipeline._write_md_dir(target / "md", files)
    written = list((target / "md").glob("*.md"))
    assert len(written) == len(files)
    text = "".join(p.read_text(encoding="utf-8") for p in written)
    assert "【A-1】" in text and "【A-2】" in text


# ---- R5T-2: CSV の「1:30」（[h]:mm を書き出した値）を分として読む ------------------------------------------

@pytest.mark.parametrize("text, unit, header_unit, expected", [
    ("1:30", "分", "", 90),
    ("1:03", "分", "", 63),
    ("01:30:30", "分", "", 90.5),
    ("1:30", "時間", "", 1.5),
    ("1時間30分", "分", "", 90),
    ("1:30", "", "分", 90),
])
def test_duration_text_is_read_as_minutes(text, unit, header_unit, expected):
    col = ColumnSpec("c", "停止時間", type="number", unit=unit)
    out, error, _flag = convert_cell(None, text, col, ConvertContext(), None, header_unit)
    assert error is None and out == expected


def test_duration_text_in_a_non_time_column_is_still_an_error():
    col = ColumnSpec("c", "費用", type="number", unit="円")
    _out, error, _flag = convert_cell(None, "1:30", col, ConvertContext(), None, "")
    assert error is not None


def test_csv_minutes_column_with_hmm_values_has_no_type_errors(tmp_path):
    from tables.detect import guess_layout
    from tables.source import open_source

    lines = ["管理No,発生日,停止時間(分)"] + [f"A-{i},2026/08/{i + 1:02d},{i + 1}:{i * 3:02d}" for i in range(10)]
    path = tmp_path / "t.csv"
    path.write_bytes(("\r\n".join(lines) + "\r\n").encode("cp932"))
    source = open_source(path, path.name)
    sheet = source.sheets()[0].name
    layout = guess_layout(source, sheet)
    spec = spec_from_dict({"name": "t", "columns": [
        {"key": "record_no", "display": "管理No", "headers": ["管理No"], "type": "code", "role": "key"},
        {"key": "occurred_at", "display": "発生日", "headers": ["発生日"], "type": "date", "role": "date"},
        {"key": "downtime", "display": "停止時間", "headers": ["停止時間(分)", "停止時間"], "type": "number",
         "role": "measure", "unit": "分"}], "record": {"key": ["record_no"]}})
    records, _issues, stats = read_records(source, {"sheet": sheet}, layout, spec)
    assert stats.type_errors == {}
    assert [r.values["downtime"] for r in records][:3] == [60, 123, 186]


# ---- R5T-3: 名前を変えるときに、前の名前を手で入れたファイル名の先頭は残す ----------------------------------

def test_renaming_keeps_an_explicitly_typed_old_name_as_prefix(app, client):
    from tests.test_endpoints_settings import _json_only_spec
    from tests.test_tables_fixes4 import _editor_body, _saved

    spec = _json_only_spec()
    spec.markdown["file_prefix"] = ""
    template_id, body = _editor_body(app, client, spec)
    old_name = body["name"]
    body["name"] = old_name + "_第2工場"
    body["file_prefix"] = old_name
    after = _saved(app, client, template_id, body)
    assert after.markdown["file_prefix"] == old_name and after.file_prefix == old_name


def test_renaming_still_clears_a_prefix_stored_as_the_old_name(app, client):
    """以前の版で空欄を設定名として保存した設定は、名前を変えると新しい名前に付いていく（これまでどおり）。"""
    from tests.test_endpoints_settings import _json_only_spec
    from tests.test_tables_fixes4 import _editor_body, _saved

    spec = _json_only_spec()
    spec.markdown["file_prefix"] = spec.name
    template_id, body = _editor_body(app, client, spec)
    body["name"] = spec.name + "_新"
    after = _saved(app, client, template_id, body)
    assert after.markdown["file_prefix"] == "" and after.file_prefix == spec.name + "_新"


# ---- R5T-4: 手で書き換えた取り込み設定の JSON は、例外ではなく日本語の問題点にする ---------------------------------

_COL = {"key": "a", "header": "a", "display": "a", "type": "string", "role": "attribute"}


@pytest.mark.parametrize("extra, message", [
    ({"columns": [{**_COL, "key": 1}]}, "キー「1」"),
    ({"columns": [{**_COL, "key": {"x": 1}}]}, "半角英数字"),
    ({"columns": [{**_COL, "type": 5}]}, "型「5」"),
    ({"columns": [_COL], "checks": {"type_error_rate": 5}}, "type_error_rate"),
    ({"columns": [_COL], "custom_stages": [{"id": 5, "inputs": ["a"]}]}, "ID「5」"),
    ({"columns": [_COL], "custom_stages": [{"id": "s", "inputs": 7}]}, "入力列「7」"),
    ({"columns": [_COL], "header": {"rows": "二"}}, "header.rows"),
    ({"columns": [_COL], "header": {"rows": 0}}, "header.rows"),
    ({"columns": [_COL], "header": {"search_rows": None}}, "header.search_rows"),
])
def test_malformed_spec_fields_become_japanese_errors(extra, message):
    errors = validate_spec(spec_from_dict({"name": "x", **extra}))
    assert any(message in e for e in errors), errors


def test_valid_spec_has_no_header_errors():
    assert validate_spec(spec_from_dict({"name": "x", "columns": [_COL], "header": {"rows": 2}})) == []


def test_importing_a_malformed_template_json_is_refused_without_500(client):
    data = {"spec": {"name": "壊れた設定", "columns": [{**_COL, "key": 1}], "checks": {"type_error_rate": 5}}}
    res = client.post("/settings/table-templates/import",
                      data={"file": (io.BytesIO(json.dumps(data).encode()), "t.json")},
                      content_type="multipart/form-data", follow_redirects=True)
    assert res.status_code == 200 and "設定の内容に問題があります" in res.get_data(as_text=True)
    with client.application.app_context():
        assert store.list_templates() == []


# ---- R5-MD-1 / R5-MD-2: 時系列の重複の削除（番号付きの列・1件だけのログ） ------------------------------------

def _timeline(action: str, log: str) -> list[str]:
    spec = spec_from_dict(list_spec_dict())
    rec = _rec("TR-001", record_no="TR-001", occurred_at="2026-08-03", equipment_id="CMP-101",
               symptom="スラリー流量低下アラームで研磨停止", action=action, response_log=log)
    lines = record_block(rec, spec)
    return [ln.strip() for ln in lines[lines.index("- 対応の時系列:") + 1:] if ln.startswith("  ")]


@pytest.mark.parametrize("action", [
    "1. フィルター交換\n2. 流量の再校正を実施",
    "①フィルター交換\n②流量の再校正を実施",
    "(1) フィルター交換\n(2) 流量の再校正を実施",
    "（１）フィルター交換\n（２）流量の再校正を実施",
    "・フィルター交換\n・流量の再校正を実施",
])
def test_numbered_column_lines_are_deduped_from_the_timeline(action):
    log = ("8/3 10:00 田中：スラリー流量低下アラームで研磨停止。保全へ連絡\n"
           "8/3 11:00 田中：フィルター交換、流量の再校正を実施\n")
    timeline = _timeline(action, log)
    assert timeline[1].endswith(": （処置と同じ）")


def test_single_entry_log_is_deduped_too():
    log = "8/3 10:00 田中：スラリー流量低下アラームで研磨停止。フィルター交換、流量の再校正を実施\n"
    timeline = _timeline("フィルター交換、流量の再校正を実施", log)
    assert len(timeline) == 1
    assert "スラリー流量低下" not in timeline[0] and "フィルター交換" not in timeline[0]
    assert timeline[0].endswith("と同じ）")


# ---- R5-MD-4: 1行目が札と日付だけの現象は、次の行を見出しに使う ----------------------------------------------

@pytest.mark.parametrize("symptom", [
    "発生:2024-04-28 14:50(休日)\n設備:CMP-103\n現象:MESとの通信が断続的に切断",
    "【発生】R06.04.28 14:50\n【現象】MESとの通信が断続的に切断",
])
def test_title_uses_the_next_line_when_the_first_is_only_a_date(symptom):
    spec = spec_from_dict(list_spec_dict())
    title = record_title({"record_no": "CA-1", "occurred_at": "2024-05-01", "equipment_id": "CMP-103",
                          "symptom": symptom}, spec)
    assert "MESとの通信が断続的に切断" in title and "設備:" not in title


def test_title_of_a_plain_symptom_is_unchanged():
    spec = spec_from_dict(list_spec_dict())
    title = record_title({"record_no": "CA-1", "occurred_at": "2024-05-01", "equipment_id": "CMP-103",
                          "symptom": "MES通信断\n詳細は別紙"}, spec)
    assert "MES通信断" in title and "詳細" not in title


# ---- SEC5-3: 選択（|）を含むグループに量指定子が付く区切りの正規表現を断る -------------------------------------

def test_overlapping_alternation_pattern_is_refused():
    from tests.test_tables_fixes4 import _log_spec

    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"extra_anchors": [r"(?:\d|\d)*年"]})))
    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"not_date_patterns": [r"(a|ab)+c"]})))
    assert validate_spec(_log_spec(splitter={"not_date_patterns": [r"\d+\.\d+\s*(?:mm|MPa)"]})) == []


# ---- R5UX-3: 管理用ファイルを保存しない設定では、保存の確認文でそう知らせる ------------------------------------

def test_table_save_confirm_mentions_admin_files_when_not_saved(app, client, tmp_path):
    from tests.test_output_folder import _set_folder
    from tests.test_retention import _confirmed_import
    from views.tables import SAVE_ADMIN_OFF_NOTE

    folder = tmp_path / "out"
    folder.mkdir()
    import_id = _confirmed_import(app, client)
    _set_folder(client, folder)
    assert SAVE_ADMIN_OFF_NOTE in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)
    _set_folder(client, folder, save_admin=True)
    assert SAVE_ADMIN_OFF_NOTE not in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)


# ---- R5-FUZZ-1: 遠くのセル1つ（XFD1 / XFD1048576）で読み込みが止まらない ---------------------------------------

def _far_book(path: Path, far: str, rows: int) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["管理No", "発生日", "設備", "現象", "処置"])
    for i in range(rows):
        ws.append([f"A-{i}", f"2026/08/{i % 28 + 1:02d}", "CMP-1", f"現象{i}", "処置"])
    ws[far] = "メモ"
    wb.save(path)
    return path


def test_far_stray_header_cell_does_not_widen_the_table(tmp_path):
    from tables.detect import guess_layout
    from tables.source import open_source

    path = _far_book(tmp_path / "far.xlsx", "XFD1", 3000)
    started = time.monotonic()
    source = open_source(path, path.name)
    assert sum(1 for _ in source.rows("S")) == 3001
    layout = guess_layout(source, "S")
    assert time.monotonic() - started < 10
    assert len(layout.headers) == 5 and any("XFD" in w for w in layout.warnings)


def test_far_last_cell_upload_finishes_and_reads_the_table(app, client, tmp_path):
    path = _far_book(tmp_path / "far2.xlsx", "XFD1048576", 50)
    started = time.monotonic()
    res = client.post("/tables/upload", data={"file": (io.BytesIO(path.read_bytes()), "far2.xlsx")},
                      content_type="multipart/form-data")
    assert time.monotonic() - started < 30
    assert res.status_code == 302 and "/tables/imports/" in res.headers["Location"]


# ---- R5C-1: 確定の処理の途中で渡し終えて消えた取り込みのフォルダを作り直さない ---------------------------------

def test_render_does_not_recreate_a_purged_import(app, monkeypatch):
    from core import purge

    with app.app_context():
        _tid, import_id = _new_import(app)
        pipeline.run_read(FakeCtx(), import_id)
        original = pipeline.load_rows

        def load_then_purge(*a, **k):
            rows = original(*a, **k)
            purge.purge_table_import(import_id)   # 別のタブの保存・ダウンロードが渡し終えた
            return rows

        monkeypatch.setattr(pipeline, "load_rows", load_then_purge)
        with pytest.raises(pipeline.PipelineError):
            pipeline.run_render(FakeCtx(), import_id)
        assert not pipeline.import_dir(import_id).exists()


def test_write_md_dir_does_not_create_a_missing_import_folder(tmp_path):
    with pytest.raises(OSError):
        pipeline._write_md_dir(tmp_path / "gone" / "md", [])
    assert not (tmp_path / "gone").exists()
