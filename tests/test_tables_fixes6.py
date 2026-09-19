"""一覧表の不具合修正（6巡目）の確認。"""
from __future__ import annotations

import csv
import io
import json
import re
import shutil
import subprocess
import time
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from tables import pipeline, store
from tables.spec import spec_from_dict, validate_spec
from tests.test_tables_pipe import FakeCtx, _new_import, _rec


# ---- R6T-1: 設定の編集を何も変えずに保存しても、JSON で選んだ期間の列（2列目の日付）を置き換えない -------------

def test_noop_save_keeps_the_period_date_column(app, client):
    from tests.test_endpoints_settings import _json_only_spec
    from tests.test_tables_fixes4 import _editor_body, _saved
    from tables.spec import ColumnSpec

    spec = _json_only_spec()
    spec.columns.append(ColumnSpec("completed_at", "完了日", headers=["完了日"], type="date", role="date"))
    spec.period["date_column"] = "completed_at"
    assert validate_spec(spec) == []
    template_id, body = _editor_body(app, client, spec)
    after = _saved(app, client, template_id, body)
    assert after.period["date_column"] == "completed_at"

    # 日付の役割の列を変えたときは画面の役割から決め直す
    for r in body["columns"]:
        if r.get("key") == "completed_at":
            r["role"] = "attribute"
    after = _saved(app, client, template_id, body)
    assert after.period["date_column"] == "occurred_at"


# ---- R6T-2: 80行目より下の見出し行を手で指定しても見出しが読める ------------------------------------------------

def _deep_book(path: Path, header_row: int, sheet: str = "一覧") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.cell(1, 1, "トラブル対応一覧（前置き）")
    ws.cell(header_row, 1, "管理No")
    ws.cell(header_row, 2, "発生日")
    ws.cell(header_row, 3, "現象")
    for i in range(1, 40):
        r = header_row + i
        ws.cell(r, 1, f"TR-{i:03d}")
        ws.cell(r, 2, f"2026/08/{(i % 28) + 1:02d}")
        ws.cell(r, 3, f"搬送ロボット{i}号機が停止")
    wb.save(path)
    return path


def test_header_row_below_the_head_rows_is_read(tmp_path):
    from tables.detect import guess_layout
    from tables.excel_source import ExcelSource

    src = ExcelSource(_deep_book(tmp_path / "deep90.xlsx", 91))
    layout = guess_layout(src, "一覧", header_rows=[91])
    assert layout.header_rows == [91] and layout.data_start == 92 and layout.data_end == 130
    assert layout.headers[:3] == ["管理No", "発生日", "現象"]
    single = guess_layout(src, "一覧", header_row=91)
    assert single.headers[:3] == ["管理No", "発生日", "現象"]


# ---- R6T-3: 見出しが自動で見つからなくても、保存した取り込み設定を選べる ------------------------------------------

def test_step2_lists_saved_settings_and_checks_the_imports_own(app, client, tmp_path):
    path = _deep_book(tmp_path / "deep40.xlsx", 41)
    res = client.post("/tables/upload", data={"file": (io.BytesIO(path.read_bytes()), "deep40.xlsx")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    spec_a = spec_from_dict({"name": "トラブル一覧", "columns": [
        {"key": "record_no", "display": "管理No", "headers": ["管理No"], "type": "code", "role": "key"}]})
    spec_b = spec_from_dict({"name": "別名", "columns": [
        {"key": "no", "display": "番号", "headers": ["番号"], "type": "code", "role": "key"}]})
    with app.app_context():
        store.create_template("トラブル一覧", spec_a, "")
        t2, v2 = store.create_template("別名", spec_b, "")
        imp = store.get_import(import_id)
        src = dict(imp["source"] or {})
        src["header_rows"] = [41]
        store.update_import(import_id, source=src, template_id=t2, template_version_id=v2)
    page = client.get(f"/tables/imports/{import_id}/source").get_data(as_text=True)
    radios = re.findall(r'<input type="radio" name="template" value="([^"]+)"\s*(checked)?', page)
    values = [v for v, _c in radios]
    assert str(t2) in values and "new" in values and len(values) == 3
    assert [v for v, c in radios if c] == [str(t2)]
    assert "管理No" in page   # 保存した見出し行（41行目）で見出しを読んでいる


# ---- R6T-4: 見出し行の入力は画面の下見とサーバーで同じように読む ---------------------------------------------------

_ROW_INPUTS = ["６,７", "3-4", "1，2", "1、2 3", "3a", "5-3", "1-20", "", "0,2", "１２０", "+5", "2-2"]


def test_row_inputs_are_read_the_same_way_on_the_server():
    from views.tables import _int_list, _row_no

    assert _int_list("６,７") == [6, 7]
    assert _int_list("3-4") == [3, 4]
    assert _int_list("1，2") == [1, 2]
    assert _int_list("3a") == [] and _int_list("5-3") == [] and _int_list("1-20") == []
    assert _int_list([3, 1]) == [1, 3]
    assert _row_no("１２０") == 120 and _row_no("12x") is None and _row_no("") is None and _row_no(None) is None


@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
def test_row_inputs_are_read_the_same_way_in_the_browser():
    from views.tables import _int_list, _row_no

    js = (Path(__file__).resolve().parents[1] / "static" / "tables.js").read_text(encoding="utf-8")
    start = js.index("    const parseEnd = ")
    end = js.index("    const apply = ")
    script = js[start:end] + f"""
const inputs = {json.dumps(_ROW_INPUTS, ensure_ascii=False)};
process.stdout.write(JSON.stringify({{rows: inputs.map(parseRows), ends: inputs.map(parseEnd)}}));
"""
    out = subprocess.run(["node", "-e", script], capture_output=True, check=True, timeout=30, encoding="utf-8")
    got = json.loads(out.stdout)
    assert got["rows"] == [_int_list(x) for x in _ROW_INPUTS]
    assert got["ends"] == [_row_no(x) for x in _ROW_INPUTS]


# ---- R6T-5: 既定の候補で AI整形の対象（追記ログ）は1列だけ ---------------------------------------------------------

def test_default_suggestions_tick_only_one_log_column():
    from tables.mapping import suggest_columns

    log = "4/1 10:00 停止を確認。\n4/2 11:00 センサーを交換。\n4/3 復旧を確認した。"
    headers = ["管理No", "対応内容", "対応内容_2", "対応内容_3"]
    rows = [[f"TR-{i}", log, log + "追記", log + "再発"] for i in range(10)]
    sugg = suggest_columns(headers, rows)
    logs = [s for s in sugg if s.role == "log"]
    assert [s.header for s in logs] == ["対応内容"]
    others = [s for s in sugg if s.header in ("対応内容_2", "対応内容_3")]
    assert all(s.role == "text" and not s.log and s.md in ("body", "omit") for s in others)


def test_untouched_columns_page_with_two_log_like_columns_can_be_saved(app, client):
    from tests.test_endpoints_settings import _collect

    log = "4/1 10:00 停止を確認。\n4/2 11:00 センサーを交換。\n4/3 復旧を確認した。"
    text = "管理No,発生日,対応内容,対応内容_2\r\n" + "".join(
        f"TR-{i},2026/08/{i + 1:02d},\"{log}\",\"{log}追記\"\r\n" for i in range(10))
    res = client.post("/tables/upload", data={"file": (io.BytesIO(text.encode("utf-8")), "log2.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "utf-8", "delimiter": ",", "template": "new", "new_template_name": "T"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    body = _collect(client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True))
    assert sum(1 for r in body["columns"] if r.get("ai")) == 1
    res = client.post(f"/tables/imports/{import_id}/columns", json=body)
    assert res.status_code == 200, res.get_json()


# ---- R6T-6: JSON の AI整形の上限・実行条件が数値でなければ保存の前に断る -------------------------------------------

def _log_spec(**stage):
    return spec_from_dict({"name": "x", "columns": [{"key": "log", "display": "ログ", "type": "text", "role": "log"}],
                           "log_stage": {"column": "log", **stage}})


@pytest.mark.parametrize("stage", [
    {"limits": {"max_segments": "多い"}},
    {"limits": {"max_input_tokens": 0}},
    {"limits": {"max_segments": True}},
    {"run_if": {"min_chars": "x"}},
    {"run_if": {"any": [{"min_segments": 2}, {"min_chars": "x"}]}},
    {"run_if": {"all": "min_chars"}},
])
def test_non_numeric_limits_and_run_if_are_refused(stage):
    assert validate_spec(_log_spec(**stage))


@pytest.mark.parametrize("stage", [
    {"limits": {"max_segments": 30, "max_input_tokens": "4000"}},
    {"run_if": {"any": [{"min_segments": 2}, {"min_chars": 80}], "contains": "停止"}},
    {"run_if": {}},
])
def test_numeric_limits_and_run_if_are_accepted(stage):
    assert validate_spec(_log_spec(**stage)) == []


# ---- R6-MD-1: 番号だけの行が先にあっても、集計の表示名は同じ番号の行の名前から付ける ---------------------------

def _code_entity_spec():
    return spec_from_dict({"name": "T", "columns": [
        {"key": "record_no", "display": "管理No", "type": "code", "role": "key"},
        {"key": "occurred_at", "display": "発生日", "type": "date", "role": "date"},
        {"key": "equipment", "display": "設備", "type": "code", "role": "entity"},
    ], "period": {"date_column": "occurred_at"}})


def test_summary_display_name_comes_from_the_group_not_the_first_row():
    from tables.spec import SummarySpec
    from tables.summaries import entity_fiscal_year_summaries, month_summaries

    spec = _code_entity_spec()
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment="CLN-502"),
               _rec("A-2", record_no="A-2", occurred_at="2026-08-04", equipment="枚葉洗浄 2号機（CLN-502）"),
               _rec("A-3", record_no="A-3", occurred_at="2026-09-04", equipment="CLN-502")]
    fy = entity_fiscal_year_summaries(records, spec, None, SummarySpec("entity_fiscal_year"))
    assert [(s["entity"], s["display"]) for s in fy] == [("CLN-502", "枚葉洗浄 2号機（CLN-502）")]
    months = month_summaries(records, spec, None, SummarySpec("month"))
    assert [m["top"][0]["display"] for m in months] == ["枚葉洗浄 2号機（CLN-502）"] * 2


# ---- R6-MD-2: 整数だけの列の平均を整数に丸めない ----------------------------------------------------------------

def test_average_of_integer_values_keeps_one_decimal():
    from tables.summaries import fmt_average

    assert fmt_average(9, 18, True) == 0.5
    assert fmt_average(57, 39, True) == 1.5
    assert fmt_average(150, 2, True) == 75 and isinstance(fmt_average(150, 2, True), int)
    assert fmt_average(0, 3, True) == 0


# ---- R6-MD-4: 「0:28以降、…」の時刻は見出しに残す ------------------------------------------------------------------

@pytest.mark.parametrize("text, expected", [
    ("0:28以降、装置がオフライン表示。", "0:28以降、装置がオフライン表示。"),
    ("12:30から停止", "12:30から停止"),
    ("12:30以降、装置がオフライン表示", "12:30以降、装置がオフライン表示"),
    ("R05.04.01 11:45(休日)、搬送停止", "搬送停止"),
    ("0:28、搬送停止", "搬送停止"),
])
def test_title_keeps_a_time_that_is_part_of_the_sentence(text, expected):
    from tables.markdown import _title_text_line

    assert _title_text_line(text) == expected


# ---- R6-1: 渡し終えた取り込みに、読み込み直しが行データ・問題一覧を書き戻さない ------------------------------------

def test_reread_does_not_write_back_into_a_purged_import(app, monkeypatch):
    from core import purge

    with app.app_context():
        _template_id, import_id = _new_import(app)
        real_checks = pipeline.run_checks

        def purge_midway(records, spec, stats):
            purge.purge_table_import(import_id)   # 読み込みの途中で渡し終えて消えた
            return real_checks(records, spec, stats)

        monkeypatch.setattr(pipeline, "run_checks", purge_midway)
        with pytest.raises(pipeline.PipelineError):
            pipeline.run_read(FakeCtx(), import_id)
        assert store.get_import(import_id) is None
        assert not pipeline.import_dir(import_id).exists()


def test_write_helpers_do_not_create_a_missing_folder(tmp_path):
    with pytest.raises(OSError):
        pipeline.write_atomic(tmp_path / "gone" / "issues.csv", b"x")
    with pytest.raises(OSError):
        pipeline.write_rows(tmp_path / "gone" / "rows.jsonl.gz", [])
    assert not (tmp_path / "gone").exists()


# ---- R6-SEC-1: 全角の ＝ ＋ － ＠ で始まるシート名・ファイル名も式にしない --------------------------------------------

def test_full_width_formula_prefixes_are_guarded():
    from tables.outputs import guard_formula, import_report_items, normalized_csv, report_csv

    for s in ("＝SUM(1,1)", "＋1", "－1", "＠SUM(1,1)"):
        assert guard_formula(s) == "'" + s
    assert guard_formula("設備") == "設備"
    spec = _code_entity_spec()
    rec = {"key": "A-1", "values": {"record_no": "A-1"}, "source": {"file": "＋1.xlsx", "sheet": "＝SUM(1,1)", "row": 2}}
    rows = list(csv.reader(io.StringIO(normalized_csv(spec, [rec]).decode("utf-8-sig"))))
    assert rows[1][-3:] == ["'＋1.xlsx", "'＝SUM(1,1)", "2"]
    imp = {"id": 1, "file_name": "＋1.xlsx", "stats": {"sheet": "＝SUM(1,1)"}}
    report = dict(csv.reader(io.StringIO(report_csv(import_report_items(imp, spec, 1)).decode("utf-8-sig"))))
    assert report["ファイル名"] == "'＋1.xlsx" and report["シート"] == "'＝SUM(1,1)"


# ---- UX6-3: 保存先フォルダがあるときの案内と、試し実行中の断りの文 --------------------------------------------------

def test_preview_note_mentions_the_save_folder_and_busy_message_mentions_saving(app, client, monkeypatch):
    from services import output_folder
    from views import tables as tables_view

    assert "保存" in tables_view.TRIAL_BUSY_MESSAGE
    with app.app_context():
        _template_id, import_id = _new_import(app)
        pipeline.run_read(FakeCtx(), import_id)
    monkeypatch.setattr(output_folder, "configured_folder", lambda: "D:\\LightRAG\\inputs")
    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    if "files" not in page or "保存先フォルダに保存したときも同じです" not in page:
        # プレビューのジョブを待ってからもう一度開く
        with app.app_context():
            from core import jobs
            job = jobs.latest_job("table_import", import_id, kind="table_preview")
            if job is not None:
                jobs.wait_job(job["id"], timeout=60)
        page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "保存先フォルダに保存したときも同じです" in page


# ---- FUZZ: 手で作った Excel（大きな非表示列・上限を超える行・列が多すぎる表） ---------------------------------------

def _rewrite_sheet(path: Path, fn) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data = fn(data)
            dst.writestr(info, data)
    path.write_bytes(buf.getvalue())


def _small_book(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    ws.append(["管理No", "発生日", "現象"])
    for i in range(5):
        ws.append([f"TR-{i}", "2026/08/01", "停止"])
    wb.save(path)
    return path


def test_huge_hidden_column_span_is_clamped(tmp_path):
    from tables.excel_source import ExcelSource

    path = _small_book(tmp_path / "cols.xlsx")
    _rewrite_sheet(path, lambda d: d.replace(
        b"<sheetData>", b'<cols><col min="1" max="20000000" hidden="1" width="5"/></cols><sheetData>', 1))
    t = time.monotonic()
    info = ExcelSource(path).sheets()[0]
    assert time.monotonic() - t < 5
    assert len(info.hidden_columns) <= 16384


def test_cell_beyond_the_excel_row_limit_is_refused_quickly(tmp_path):
    from tables.excel_source import ExcelSource
    from tables.source import UploadError

    path = _small_book(tmp_path / "far.xlsx")
    _rewrite_sheet(path, lambda d: d.replace(
        b"</sheetData>",
        b'<row r="4294967296"><c r="D4294967296" t="inlineStr"><is><t>z</t></is></c></row></sheetData>', 1))
    t = time.monotonic()
    with pytest.raises(UploadError, match="Excel の上限"):
        ExcelSource(path)
    assert time.monotonic() - t < 5


def test_xlsx_with_too_many_columns_is_refused(tmp_path):
    from tables.csv_source import MAX_COLUMNS
    from tables.excel_source import ExcelSource
    from tables.source import UploadError

    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    ws.append([f"項目{i}" for i in range(MAX_COLUMNS + 1)])
    ws.append([f"v{i}" for i in range(MAX_COLUMNS + 1)])
    path = tmp_path / "wide.xlsx"
    wb.save(path)
    src = ExcelSource(path)
    with pytest.raises(UploadError, match="列数が上限"):
        list(src.rows("一覧", 1, 2))


def test_far_memo_cell_is_not_counted_as_many_columns(tmp_path):
    from tables.excel_source import ExcelSource

    path = _small_book(tmp_path / "memo.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    ws.append(["管理No", "発生日", "現象"])
    ws.append(["TR-1", "2026/08/01", "停止"])
    ws["XFD1"] = "メモ"
    wb.save(path)
    rows = list(ExcelSource(path).rows("一覧", 1, 2))
    assert len(rows) == 2


# ---- R6C-1: 確定した取り込みでは試し実行を断る（zip・保存と下書きが食い違わないように） ---------------------------------

def test_trial_is_refused_after_confirm(app, client):
    with app.app_context():
        _template_id, import_id = _new_import(app)
        pipeline.run_read(FakeCtx(), import_id)
        pipeline.run_render(FakeCtx(), import_id)
        assert store.get_import(import_id)["status"] == "confirmed"
    res = client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 400
    assert "確定後は試し実行できません" in res.get_json()["error"]


# ---- R6C-2: 渡している間に別のタブで読み込み直しが始まったら、「ダウンロード済み」とは言わない ----------------------------

@pytest.mark.parametrize("method, url, builder", [
    ("get", "download.zip", "build_download"),
    ("post", "save-to-folder", "build_download_files"),
])
def test_reread_during_hand_out_goes_back_to_the_preview(app, client, monkeypatch, tmp_path, method, url, builder):
    from services import output_folder

    with app.app_context():
        _template_id, import_id = _new_import(app)
        pipeline.run_read(FakeCtx(), import_id)
        pipeline.run_render(FakeCtx(), import_id)

    def reread_started(*_a, **_k):
        store.update_import(import_id, status="preview")   # 別のタブの読み込み直しが md を消した
        raise FileNotFoundError("md")

    monkeypatch.setattr(pipeline, builder, reread_started)
    monkeypatch.setattr(output_folder, "configured_folder", lambda: str(tmp_path))
    res = getattr(client, method)(f"/tables/imports/{import_id}/{url}")
    assert res.status_code == 302 and res.headers["Location"].endswith(f"/tables/imports/{import_id}/preview")
    with app.app_context():
        assert store.get_import(import_id) is not None


# ---- R6B-1: 入力したキーが重なったら黙って「_2」にせず断る -------------------------------------------------------

def test_duplicate_typed_key_is_refused_and_does_not_carry_another_columns_settings():
    from tables.spec import spec_from_suggestions
    from views.tables import _build_spec

    base = spec_from_suggestions("T", {"header_rows": [1]}, [
        {"header": "台帳No", "key": "record_no", "role": "key", "type": "code"},
        {"header": "起票日", "key": "date", "role": "date", "type": "datetime"}])
    base.columns[0].value_map = {"A": "B"}
    base.columns[0].headers = ["台帳No", "管理番号"]
    rows = [{"use": 1, "header": "台帳No", "key": "record_no", "role": "key", "type": "code"},
            {"use": 1, "header": "起票日", "key": "record_no", "role": "date", "type": "datetime"}]
    _spec, errors = _build_spec({"name": "T", "columns": rows}, base)
    assert "キー「record_no」が重複しています" in errors

    # キーを入れ替えただけなら保存できるが、別の列の見出しの別名・値の置き換えは引き継がない
    swapped = [{"use": 1, "header": "台帳No", "key": "date", "role": "key", "type": "code"},
               {"use": 1, "header": "起票日", "key": "record_no", "role": "date", "type": "datetime"}]
    spec, errors = _build_spec({"name": "T", "columns": swapped}, base)
    assert errors == []
    dated = spec.column("record_no")
    assert dated.value_map == {} and "台帳No" not in dated.headers and "管理番号" not in dated.headers


def test_duplicate_typed_key_is_refused_by_the_editor(app, client):
    from tests.test_endpoints_settings import _json_only_spec
    from tests.test_tables_fixes4 import _editor_body

    template_id, body = _editor_body(app, client, _json_only_spec())
    keys = [r["key"] for r in body["columns"] if r.get("use")]
    body["columns"][1]["key"] = keys[0]
    res = client.post(f"/tables/templates/{template_id}", json=body)
    assert res.status_code == 400
    assert f"キー「{keys[0]}」が重複しています" in json.dumps(res.get_json(), ensure_ascii=False)


# ---- R6B-2: md のパスが Windows の上限を超えるときは分かる言葉で止める --------------------------------------------------

def test_too_long_md_path_gives_a_clear_message(tmp_path, monkeypatch):
    from services import output_folder
    from tables.markdown import MdFile

    monkeypatch.setattr(output_folder, "_long_paths_enabled", lambda: False)
    target = tmp_path / "md"
    tmp_path.mkdir(exist_ok=True)
    name = "あ" * (output_folder.MAX_PATH_CHARS - len(str(target)) + 10) + ".md"
    files = [MdFile(name=name, text="x", kind="records")]
    with pytest.raises(pipeline.PipelineError, match="ファイル名が長すぎ"):
        pipeline._write_md_dir(target, files)
    assert not target.exists()
