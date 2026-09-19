"""一覧表の不具合修正（2巡目）の確認。"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from core import jobs
from tables import pipeline, store
from tables.checks import has_blocking, run_checks
from tables.csv_source import MAX_COLUMNS, CsvSource, sniff_csv
from tables.detect import guess_layout, sample_data_rows
from tables.mapping import suggest_columns
from tables.normalize import parse_number_text, read_records
from tables.source import UploadError, open_source
from tables.spec import spec_from_dict, spec_from_suggestions
from tests.test_tables_flow import COLUMNS, CSV_TEXT, _wait_import_job
from tests.test_tables_pipe import _rec, list_spec_dict


def _auto(src, sheet):
    layout = guess_layout(src, sheet)
    spec = spec_from_suggestions("テスト", layout, suggest_columns(layout.headers, sample_data_rows(src, sheet, layout)), {})
    records, _issues, stats = read_records(src, {"sheet": sheet}, layout, spec)
    return layout, spec, records, stats


# ---- 読み込み（CSV） ----------------------------------------------------------------------------

def test_cp932_file_with_one_bad_byte_stays_cp932(tmp_path):
    """CP932 の拡張文字（髙・﨑）があるファイルに読めないバイトが1つあっても、Shift_JIS-2004 にして字を化けさせない。"""
    rows = "".join(f"TR-00{i},2026-08-01,髙砂ポンプ{i},㈱山﨑製作所で停止\r\n" for i in range(5))
    p = tmp_path / "a.csv"
    p.write_bytes(("管理No,発生日,設備名,内容\r\n" + rows).encode("cp932") + b"TR-999,2026-08-02,\x86\x9f,x\r\n")
    sniff = sniff_csv(p)
    assert sniff.encoding == "cp932" and sniff.decode_error_line == 7 and sniff.warnings
    src = CsvSource(p, "a.csv", {"encoding": "cp932", "delimiter": ",", "errors": "replace"})
    texts = [c.text for r in src.rows() for c in r.cells]
    assert "髙砂ポンプ0" in texts and "㈱山﨑製作所で停止" in texts


def test_long_but_closed_quoted_value_is_a_warning_not_an_error(tmp_path):
    """" が正しく閉じた長い値（600行の貼り付けログ）は確定を止めない。閉じ忘れの疑いとして警告だけ出す。"""
    log = "\n".join(f"{i}: 確認" for i in range(600))
    text = ("管理No,発生日,対応内容\r\nTR-001,2026-08-01,a\r\nTR-002,2026-08-02,b\r\nTR-003,2026-08-03,c\r\n"
            f'TR-004,2026-08-04,"{log}"\r\nTR-005,2026-08-05,e\r\n')
    p = tmp_path / "long.csv"
    p.write_bytes(text.encode("utf-8"))
    src = open_source(p, "long.csv")
    assert len(list(src.rows())) == 6 and src.unclosed_quote_row is None and src.long_record_row == 5
    layout, spec, records, stats = _auto(src, "long.csv")
    assert len(records) == 5 and stats.long_record_row == 5
    issues = run_checks(records, spec, stats)
    assert not has_blocking(issues)
    assert any(i.code == "long_record" and i.level == "warning" and "5行目" in i.message for i in issues)


def test_csv_wider_than_the_column_limit_is_refused(tmp_path):
    p = tmp_path / "wide.csv"
    header = ",".join(f"列{i}" for i in range(MAX_COLUMNS + 1))
    p.write_text(header + "\r\n" + ",".join("1" for _ in range(MAX_COLUMNS + 1)) + "\r\n", encoding="utf-8")
    with pytest.raises(UploadError, match="列数が上限"):
        open_source(p, "wide.csv")


# ---- 表の範囲・見出し ----------------------------------------------------------------------------

def test_row_with_its_own_no_and_date_is_not_a_total_row(tmp_path):
    """点検項目が「稼働時間累計」でも、自分の点検No・点検日がある行はデータ（合計行で表を切らない）。"""
    lines = ["点検No,点検日,点検項目,測定値,判定"]
    for i in range(40):
        item = "稼働時間累計" if i % 4 == 3 else f"振動{i}"
        lines.append(f"P-{i:03d},2026/04/{i % 28 + 1:02d},{item},{i * 10},{'' if i % 4 == 3 else 'OK'}")
    p = tmp_path / "check.csv"
    p.write_bytes(("\r\n".join(lines) + "\r\n").encode("utf-8"))
    src = open_source(p, "check.csv")
    layout = guess_layout(src, "check.csv")
    assert layout.counts.get("data") == 40 and not layout.counts.get("subtotal") and layout.warnings == []
    # 番号・日付のない本当の合計行は、これまでどおり合計
    p2 = tmp_path / "total.csv"
    p2.write_bytes(("\r\n".join(lines[:11] + [",,合計,450,"]) + "\r\n").encode("utf-8"))
    src2 = open_source(p2, "total.csv")
    assert guess_layout(src2, "total.csv").counts.get("data") == 10


def test_unlabeled_last_column_is_kept(tmp_path):
    """右端の見出しが空欄でも値のある列は捨てない（「列6」として取り込み、Markdown にも出る）。"""
    text = "管理No,発生日,設備名,現象,停止時間(分),\r\n" + "".join(
        f"T{i},2026/04/{i + 1:02d},ポンプ,停止,{i},ベアリング交換済み{i}\r\n" for i in range(20))
    p = tmp_path / "blank.csv"
    p.write_bytes(text.encode("utf-8"))
    src = open_source(p, "blank.csv")
    layout, spec, records, _stats = _auto(src, "blank.csv")
    assert layout.headers[-1] == "列6"
    col = next(c for c in spec.columns if "列6" in (c.headers or [c.display]))
    assert records[3].values.get(col.key) == "ベアリング交換済み3"


def test_repeated_headers_keep_their_name_as_display():
    """同じ見出し「備考」が3つあっても、表示名が「2」「3」にならない。"""
    from tables.detect import _dedupe
    from tables.mapping import _display_name

    headers = _dedupe(["管理No", "発生日", "設備名", "備考", "備考", "備考"])
    assert headers[3:] == ["備考", "備考(2)", "備考(3)"]
    assert [_display_name(h) for h in headers[3:]] == ["備考", "備考(2)", "備考(3)"]
    sugg = suggest_columns(headers, [["T1", "2026/04/01", "ポンプ", "a", "b", "c"]])
    displays = [s["display"] if isinstance(s, dict) else s.display for s in (sugg.values() if isinstance(sugg, dict) else sugg)]
    assert "2" not in displays and "3" not in displays


# ---- 数値の変換 --------------------------------------------------------------------------------

def test_excel_number_under_an_unconvertible_unit_is_reported():
    from tables.normalize import _convert_number
    from tables.spec import ColumnSpec

    col = ColumnSpec(key="downtime", display="停止時間", type="number", role="measure", unit="分")
    assert _convert_number(7200, "7200", col, None, "s") == (120, None, None)   # 秒→分
    assert _convert_number(3, "3", col, None, "時") == (180, None, None)        # 時→分
    value, error, _ = _convert_number(5, "5", col, None, "kg")
    assert value == 5 and error and "換算できません" in error


def test_huge_exponent_is_a_cell_issue_not_a_crash(tmp_path):
    assert parse_number_text("1e309") == (None, "")
    assert parse_number_text("9" * 400) == (None, "")
    assert parse_number_text("1e308") == (1e308, "")
    text = "管理No,発生日,停止時間(分)\r\n" + "".join(f"T{i},2026/04/{i + 1:02d},{i}\r\n" for i in range(10)) + "T99,2026/04/20,1e999\r\n"
    p = tmp_path / "big.csv"
    p.write_bytes(text.encode("utf-8"))
    src = open_source(p, "big.csv")
    layout = guess_layout(src, "big.csv")
    spec = spec_from_suggestions("テスト", layout, suggest_columns(layout.headers, sample_data_rows(src, "big.csv", layout)), {})
    num = next(c for c in spec.columns if c.display.startswith("停止時間"))
    num.type, num.role = "number", "measure"
    records, issues, _stats = read_records(src, {"sheet": "big.csv"}, layout, spec)
    assert len(records) == 11
    assert any("1e999" in i.message and "数値に変換できません" in i.message for i in issues)


def test_overflowing_sum_renders_without_error():
    from tables.markdown import render_all
    from tables.summaries import fmt_number

    assert fmt_number(float("inf")) == "inf"
    spec = spec_from_dict(list_spec_dict(group_by="entity_month"))
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="CMP-101", cost=1e308),
               _rec("A-2", record_no="A-2", occurred_at="2026-08-10", equipment_id="CMP-101", cost=1e308)]
    assert render_all(spec, records, {}, {})


# ---- 集計の Markdown ---------------------------------------------------------------------------

def test_month_with_only_blank_measures_is_not_written_as_zero():
    from tables.markdown import render_all

    spec = spec_from_dict(list_spec_dict(group_by="entity_month"))
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="CMP-101", cost=500, downtime=10),
               _rec("A-2", record_no="A-2", occurred_at="2026-09-10", equipment_id="CMP-101", downtime=20)]
    files = render_all(spec, records, {}, {})
    summary = next(f for f in files if f.kind == "summary" and "設備別" in f.name).text
    assert "2026年9月: 1件" in summary
    line = next(line for line in summary.split("\n") if line.startswith("- 2026年9月"))
    assert "費用 値なし" in line and "費用 0円" not in line
    # 1年分すべて空欄の測定値は「合計0」と書かない
    records = [_rec("A-3", record_no="A-3", occurred_at="2026-08-03", equipment_id="CMP-101", downtime=10)]
    summary = next(f for f in render_all(spec, records, {}, {}) if f.kind == "summary" and "設備別" in f.name).text
    assert "費用の値はありません" in summary and "費用の合計は0円" not in summary


def test_placeholder_equipment_is_not_an_equipment():
    """対象設備が「調査中」だけの記録は、設備別の集計・ファイル分けに入れない（本文には原文のまま出す）。"""
    from tables.markdown import render_all
    from tables.summaries import dataset_counts, entity_value

    d = list_spec_dict(group_by="entity_month")
    d["columns"] = [c for c in d["columns"] if c["key"] != "equipment_name"]
    spec = spec_from_dict(d)
    assert entity_value({"equipment_id": "調査中"}, spec) == ("", "")
    assert entity_value({"equipment_id": "CMP-101"}, spec) == ("CMP-101", "")
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="調査中", symptom="停止"),
               _rec("A-2", record_no="A-2", occurred_at="2026-08-10", equipment_id="CMP-101", symptom="異音")]
    files = render_all(spec, records, {}, {})
    assert not any("調査中" in f.name for f in files)
    assert "- 設備番号: 調査中" in "".join(f.text for f in files if f.kind == "records")
    assert dataset_counts(records, spec)["entities"] == 1


# ---- 画面（取り込み・削除・ダウンロード） -----------------------------------------------------------------

def _confirmed_import(app, client) -> int:
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "トラブル一覧.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "トラブル対応一覧"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "トラブル対応一覧", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                            "fill_down_blank": False, "ai": False, "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)
    client.get(f"/tables/imports/{import_id}/preview")
    with app.app_context():
        job = jobs.latest_job("table_import", import_id, kind="table_preview")
        if job is not None:
            jobs.wait_job(job["id"], timeout=60)
    client.post(f"/tables/imports/{import_id}/confirm")
    assert _wait_import_job(app, import_id)["status"] == "confirmed"
    return import_id


def test_template_in_use_by_an_undownloaded_import_cannot_be_deleted(app, client):
    import_id = _confirmed_import(app, client)
    with app.app_context():
        tid = store.get_import(import_id)["template_id"]
    res = client.post(f"/tables/templates/{tid}/delete")
    assert res.status_code == 302
    with app.app_context():
        assert store.get_template(tid) is not None
        assert store.get_import(import_id)["template_version_id"] is not None
    res = client.get(f"/tables/imports/{import_id}/download.zip")
    assert res.status_code == 200 and res.mimetype == "application/zip"
    # ダウンロードで取り込みが消えたあとは削除できる
    assert client.post(f"/tables/templates/{tid}/delete").status_code == 302
    with app.app_context():
        assert store.get_template(tid) is None


def test_second_download_after_purge_is_404_not_500(app, client, monkeypatch):
    import_id = _confirmed_import(app, client)

    def gone(*args, **kwargs):
        raise FileNotFoundError("purged")

    monkeypatch.setattr(pipeline, "build_download", gone)
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 404


def test_delete_is_refused_while_the_preview_draft_is_being_made(app, client):
    from tests.test_tables_flow import _queued_job, _uploaded_csv

    import_id = _uploaded_csv(client)
    _queued_job(app, import_id, "table_preview")
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 302 and res.headers["Location"].endswith("/preview")
    with app.app_context():
        assert store.get_import(import_id) is not None


def test_failed_preview_draft_can_be_made_again(app, client, monkeypatch):
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "t.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "再作成テスト"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "再作成テスト", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "", "md": "attribute", "fill_down_blank": False, "ai": False,
                            "description": ""} for i, (k, h, t, r) in enumerate(COLUMNS)]}
    client.post(f"/tables/imports/{import_id}/columns", json=payload)
    _wait_import_job(app, import_id)

    real = pipeline.render_files
    calls = {"n": 0}

    def flaky(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            raise OSError("一時的に使えない")
        return real(*args, **kwargs)

    monkeypatch.setattr(pipeline, "render_files", flaky)

    def wait_preview():
        with app.app_context():
            jobs.wait_job(jobs.latest_job("table_import", import_id, kind="table_preview")["id"], timeout=60)

    client.get(f"/tables/imports/{import_id}/preview")
    wait_preview()
    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "Markdownの下書きを作れませんでした" in page and "もう一度作る" in page
    assert "表を読み込めませんでした" not in page
    res = client.get(f"/tables/imports/{import_id}/preview?retry=1")
    assert res.status_code == 302 and "retry" not in res.headers["Location"]
    wait_preview()
    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "確定してMarkdownを作成" in page


def test_import_status_is_written_before_the_job_can_finish(app, monkeypatch):
    """ジョブがすぐ終わって書いた状態を、あとから reading / confirming で上書きしない。"""
    with app.app_context():
        import_id = store.create_import("x.csv", "h" * 64, "tables/x.csv", {"encoding": "utf-8"})

        def instant(kind, ref_type, ref_id, fn, params):
            # ワーカーが先に終わった場合と同じ: 呼び出し元が状態を書く前にジョブが状態を書く
            store.update_import(ref_id, status="preview" if kind == "table_read" else "confirmed")
            return 12345

        monkeypatch.setattr(pipeline.jobs, "start_job", instant)
        assert pipeline.start_read_job(import_id) == 12345
        imp = store.get_import(import_id)
        assert imp["status"] == "preview" and imp["job_id"] == 12345
        pipeline.start_render_job(import_id)
        assert store.get_import(import_id)["status"] == "confirmed"


# ---- 読み込み（Excel） --------------------------------------------------------------------------

def _small_book(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.append(["管理No", "発生日", "設備名", "現象", "停止時間"])
    for i in range(10):
        ws.append([f"T{i}", f"2026/04/{i + 1:02d}", "ポンプ", "停止", i + 1])
    return wb, ws


def test_format_only_far_cell_does_not_block_upload(app, client, tmp_path):
    from tables.excel_source import ExcelSource

    wb, ws = _small_book(tmp_path)
    ws["XFD1048576"].fill = PatternFill("solid", fgColor="FFFF00")
    path = tmp_path / "far.xlsx"
    wb.save(path)
    ExcelSource(path).close()
    res = client.post("/tables/upload", data={"file": (io.BytesIO(path.read_bytes()), "far.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and res.headers["Location"].endswith("/source")


def test_workbook_with_unreadable_values_is_refused_at_upload(app, client, tmp_path):
    wb, _ws = _small_book(tmp_path)
    wb.save(tmp_path / "ok.xlsx")
    src = zipfile.ZipFile(tmp_path / "ok.xlsx")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<v>3</v>", b"<v>NaN</v>", data, count=1)
            z.writestr(item, data)
    res = client.post("/tables/upload", data={"file": (io.BytesIO(out.getvalue()), "nan.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and res.headers["Location"].endswith("/tables/new")
    with app.app_context():
        assert store.list_imports(limit=10) == []
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


def test_retention_note_does_not_ask_for_the_csv_first():
    from views.tables import DELETE_ON_DOWNLOAD_NOTE

    assert "より先" not in DELETE_ON_DOWNLOAD_NOTE and "管理用_RAGには入れない" in DELETE_ON_DOWNLOAD_NOTE


def test_confirmed_import_can_be_deleted_from_home_and_done_page(app, client):
    import_id = _confirmed_import(app, client)
    action = f'action="/tables/imports/{import_id}/delete"'
    assert action in client.get("/").get_data(as_text=True)
    assert action in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)
    res = client.post(f"/tables/imports/{import_id}/delete", data={"next": "/"})
    assert res.status_code == 302
    with app.app_context():
        assert store.get_import(import_id) is None


def test_preview_headers_show_the_unit_and_template_list_label(app, client):
    import_id = _confirmed_import(app, client)
    html = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "<th>停止時間（分）</th>" in html
    listing = client.get("/settings/table-templates").get_data(as_text=True)
    assert "まだダウンロードしていない取り込み（件）" in listing
    assert "ダウンロード待ちの取り込み" not in listing


def test_404_for_a_missing_setting_does_not_blame_a_download(client):
    for path, back in (("/settings/form-types/999", "/settings/form-types"),
                       ("/tables/templates/999", "/settings/table-templates")):
        res = client.get(path)
        assert res.status_code == 404
        html = res.get_data(as_text=True)
        assert "この設定は見つかりません" in html and "ダウンロード済み" not in html
        assert f'href="{back}' in html
    html = client.get("/tables/imports/999/done").get_data(as_text=True)
    assert "ダウンロード済み" in html and "この設定は見つかりません" not in html
