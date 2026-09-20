"""表の取り込みの直し（記録0件のまま確定できる・日付の列が選べない・見出しと並び順）。

このまわりで見つかった不具合:
  - 記録0件でも「確定」でき、④のダウンロードを押すと空の画面に戻されて取り込みが消える
  - 「日付の列を1つ選んでください」と言うのに、年月だけの列を選ぶと保存が400で拒否され直せない
  - 識別番号・日付・設備・長文のどれも無い表は、全記録の見出しが「（見出しなし）」になる
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font

from tables.checks import has_blocking, run_checks
from tables.markdown import record_title
from tables.spec import spec_from_dict
from tests.tables_helpers import (columns_payload, csv_source, panel, panel_html, preview_panel, save_columns,
                                  save_layout, upload_bytes, upload_csv, wait_import_job)
from tests.test_tables_pipe import _rec, list_spec_dict


def _struck_book(path):
    """見出し1行＋データ5行。データ行は全部取り消し線（＝取り込める行が1件も無い表）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    for col, header in enumerate(["管理No", "発生日", "設備番号", "現象"], start=1):
        ws.cell(row=1, column=col, value=header).font = Font(bold=True)
    for i in range(1, 6):
        for col, value in enumerate([f"TR-{i:03d}", f"2026-08-{i:02d}", "EQ-01", f"アラーム{i}"], start=1):
            ws.cell(row=1 + i, column=col, value=value).font = Font(strike=True)
    wb.save(path)
    return path


# 「実施月」が年月だけ（2025-03）なので、日付として読めない表
MONTH_CSV = "管理No,実施月,設備番号,現象,対応内容\r\n" + "".join(
    f"K-{i:03d},2025-0{i},EQ-0{i},現象{i},対応{i}\r\n" for i in range(1, 6))

def _month_columns(name: str, date_role: str = "attribute") -> dict:
    roles = ["key", date_role, "attribute", "attribute", "attribute"]
    return {"name": name, "columns": [{"index": i, "use": True, "role": r} for i, r in enumerate(roles)]}


def _prepare(client, csv_name: str, text: str) -> int:
    import_id = upload_csv(client, csv_name, text)
    csv_source(client, import_id)
    assert save_layout(client, import_id).status_code == 200
    return import_id


def _struck_import(app, client, tmp_path, name: str) -> int:
    """全部の行が取り消し線の表を、④の保存まで通した取り込み。"""
    path = _struck_book(tmp_path / name)
    import_id = upload_bytes(client, path.read_bytes(), name)
    assert client.post(f"/tables/imports/{import_id}/source", json={"sheet": "一覧"}).status_code == 200
    assert save_layout(client, import_id).status_code == 200
    payload = {"name": "取り消し線の表",
               "columns": [{"index": i, "use": True, "role": r}
                           for i, r in enumerate(["key", "date", "entity", "attribute"])]}
    assert save_columns(client, import_id, payload).status_code == 200, save_columns
    wait_import_job(app, import_id)
    return import_id


# ---- 記録0件のまま確定させない -----------------------------------------------------------------

def test_a_table_with_no_usable_rows_cannot_be_confirmed():
    """取り込める行が1件も無いときは、確定を止めて理由を出す（警告のまま通さない）。

    警告のままだと⑥で［確定してMarkdownを作成］が押せてしまい、⑦のダウンロードで
    「先に［確定してMarkdownを作成］を押してください」と言われて空の画面に戻されていた。
    """
    spec = spec_from_dict(list_spec_dict())
    issues = run_checks([], spec, {"excluded": {"取り消し線の行": 5}})
    no_records = [i for i in issues if i.code == "no_records"]
    assert len(no_records) == 1
    assert no_records[0].level == "error"
    assert "取り消し線・非表示などで5行を除外しました" in no_records[0].message
    assert "表の範囲か元のファイルを見直してください" in no_records[0].message
    assert has_blocking(issues) is True


def test_the_preview_step_blocks_the_confirm_button_when_nothing_can_be_read(app, client, tmp_path):
    """⑥の［確定してMarkdownを作成］が押せなくなり、理由が画面に出る。"""
    import_id = _struck_import(app, client, tmp_path, "取り消し線.xlsx")

    html = preview_panel(app, client, import_id)["html"]
    assert "取り込める行がありません" in html
    assert "取り消し線・非表示などで5行を除外しました" in html
    assert "data-confirm-run disabled" in html
    assert "エラーが残っているため確定できません" in html
    # 確定も断る
    res = client.post(f"/tables/imports/{import_id}/confirm", json={})
    assert res.status_code == 400 and "エラーが残っているため確定できません" in res.get_json()["error"]


def test_a_confirmed_import_with_no_markdown_does_not_offer_a_download(app, client, tmp_path):
    """記録0件のまま確定された取り込みは、⑦でダウンロードのボタンを出さない。"""
    from tables import store

    import_id = _struck_import(app, client, tmp_path, "取り消し線2.xlsx")
    with app.app_context():
        store.update_import(import_id, status="confirmed")
    data = panel(client, import_id, "done")
    assert "この取り込みから作られた Markdown はありません" in data["locked"]
    assert "download.zip" not in data["html"]

    res = client.get(f"/tables/imports/{import_id}/download.zip")
    assert res.status_code == 302
    page = client.get("/tables/new").get_data(as_text=True)
    assert "作成された Markdown がありません" in page
    assert "先に [確定してMarkdownを作成] を押してください" not in page


# ---- 日付にできない列は「日付」の選択肢を出さない -------------------------------------------------

def test_a_column_that_is_not_a_date_has_no_date_role_to_choose(app, client):
    """年月だけの列（2025-03）に「日付」の選択肢を出さない（選ぶと保存が必ず400になるため）。"""
    import_id = _prepare(client, "月報.csv", MONTH_CSV)
    html = panel_html(client, import_id, "columns")

    rows = html.split("<tr data-col")[1:]
    assert len(rows) == 5
    month = next(r for r in rows if 'data-header="実施月"' in r)
    assert '<option value="date"' not in month          # 日付として読めないので出さない
    assert '<option value="key"' in month and '<option value="attribute"' in month


def test_a_table_without_any_date_column_is_not_asked_to_pick_one(app, client):
    """日付として読める列が1つも無い表では、「日付の列を1つ選んでください」と求めない。"""
    import_id = _prepare(client, "月報2.csv", MONTH_CSV)
    html = panel_html(client, import_id, "columns")
    assert "日付の列が決まっていません。1つ選んでください" not in html


def test_choosing_a_date_role_for_a_non_date_column_says_what_is_wrong(app, client):
    """古い画面から送られても、「型を日付にしてください」（直せない指示）とは返さない。"""
    import_id = _prepare(client, "月報3.csv", MONTH_CSV)
    res = save_columns(client, import_id, _month_columns("月報", date_role="date"))
    assert res.status_code == 400
    error = res.get_json()["error"]
    assert "列「実施月」の値は日付として読めないので、日付の列にはできません" in error
    assert "型を日付にしてください" not in error


def test_a_real_date_column_can_still_be_chosen(app, client):
    """型の推定が当たっている日付の列には、これまでどおり「日付」を出す。"""
    from tests.tables_helpers import CSV_TEXT

    import_id = _prepare(client, "一覧.csv", CSV_TEXT)
    html = panel_html(client, import_id, "columns")
    rows = html.split("<tr data-col")[1:]
    occurred = next(r for r in rows if 'data-header="発生日"' in r)
    assert '<option value="date"' in occurred
    assert save_columns(client, import_id, columns_payload("トラブル対応一覧")).status_code == 200


# ---- 見出しと並び順（見出しの材料が何も無い表） --------------------------------------------------

def _code_spec():
    d = list_spec_dict()
    d["columns"] = [
        {"key": "kind", "display": "コード種別", "type": "code", "role": "attribute", "md": "attribute"},
        {"key": "code", "display": "コード", "type": "code", "role": "attribute", "md": "attribute"},
        {"key": "name", "display": "名称", "type": "string", "role": "attribute", "md": "attribute"},
    ]
    d["record"] = {"key": ["code"], "fallback_key": ["code"]}
    d["period"] = {"date_column": ""}
    d["markdown"] = {k: v for k, v in (d.get("markdown") or {}).items() if k != "title_columns"}
    return spec_from_dict(d)


def test_records_without_a_title_column_get_a_heading_from_their_values():
    """見出しの材料が無い表でも、記録ごとに違う見出しにする（全部「（見出しなし）」にしない）。"""
    spec = _code_spec()
    first = record_title({"kind": "区分1", "code": "C001", "name": "停止"}, spec, {"row": 2})
    second = record_title({"kind": "区分2", "code": "C002", "name": "故障"}, spec, {"row": 3})
    assert first != second and "（見出しなし）" not in (first, second)
    assert "C001" in first and "C002" in second


def test_a_completely_empty_record_still_says_which_row_it_came_from():
    """値が1つも無い行だけ、行番号を見出しにする（同じ見出しが並ばない）。"""
    spec = _code_spec()
    assert record_title({}, spec, {"row": 42}) == "42行目の記録"
    assert record_title({}, spec) == "（見出しなし）"


def test_records_without_a_date_keep_the_order_of_the_source_table():
    """日付の列が無い表は、元の表の行の順に並べる（記録キーの文字くらべにしない）。"""
    from tables.markdown import render_all

    spec = _code_spec()
    records = []
    for row in (2, 10, 11, 100, 101):
        rec = _rec(f"行{row}", code=f"C{row:03d}", kind="区分", name=f"名称{row}")
        rec["source"] = {"file": "コード表.csv", "row": row}
        records.append(rec)
    files = render_all(spec, list(reversed(records)), {})
    assert len(files) == 1
    rows = [int(line.split("（")[1].split("行目")[0])
            for line in files[0].text.splitlines() if line.startswith("- 出典:")]
    assert rows == [2, 10, 11, 100, 101]


def test_records_with_a_date_are_still_ordered_by_date(app, client):
    """日付のある表の並びは今までどおり（同じ日付のときだけ元の行の順になる）。"""
    from tables.markdown import render_all

    spec = spec_from_dict(list_spec_dict())
    early = _rec("A-2", record_no="A-2", occurred_at="2026-08-01", equipment_id="EQ-1")
    late = _rec("A-1", record_no="A-1", occurred_at="2026-08-09", equipment_id="EQ-1")
    early["source"], late["source"] = {"row": 9}, {"row": 2}
    text = "\n".join(f.text for f in render_all(spec, [late, early], {}))
    assert text.index("A-2") < text.index("A-1")
