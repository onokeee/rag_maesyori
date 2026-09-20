"""設定まわりの経路: 取り込み設定の編集保存・JSONの書き出しと読み込み、帳票の種類の見本と状態、
帳票の JSON・元ファイルのダウンロード、Excel のシートの選択。"""
import copy
import io
import json
import re
from html.parser import HTMLParser

import openpyxl

from core import jobs
from models import database as db
from tables import store
from tables.spec import spec_from_dict, spec_to_dict
from tests.test_aiproc import SPEC
from tests.test_tables_flow import COLUMNS, CSV_TEXT


# ---- 取り込み設定の編集画面（static/tables.js の collect() と同じ値を HTML から集める） ----------------

class _EditorParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.settings, self.cols, self.cur, self.sel, self.header_rows = {}, [], None, None, 1

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "option" and self.sel:
            target, key = self.sel
            if target[key] is None or "selected" in a:
                target[key] = a.get("value")
            return
        if "data-columns-editor" in a:
            self.header_rows = int(a.get("data-header-rows-count") or 1)
        if tag == "tr" and "data-col" in a:
            self.cur = {"index": int(a["data-index"]), "header": a["data-header"]}
            self.cols.append(self.cur)
        target = key = None
        if "data-setting" in a:
            target, key = self.settings, a["data-setting"]
        elif "data-field" in a and self.cur is not None:
            target, key = self.cur, a["data-field"]
        if target is None:
            return
        if tag == "input":
            target[key] = ("checked" in a) if a.get("type") == "checkbox" else a.get("value", "")
        elif tag == "select":
            self.sel = (target, key)
            target[key] = None

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        if tag == "select":
            self.sel = None


def _collect(html: str) -> dict:
    p = _EditorParser()
    p.feed(html)
    return {**p.settings, "header_rows_count": p.header_rows, "columns": p.cols}


def _json_only_spec():
    """JSON で読み込んだときにだけ付く設定（見出しの別名・単位換算・値の置き換えなど）を持つ取り込み設定。"""
    d = copy.deepcopy(SPEC)
    for col in d["columns"]:
        if col["key"] == "occurred_at":
            col["headers"] = ["発生日", "発生日時", "日付"]
        if col["key"] == "equipment_name":
            col["value_map"] = {"ロボ2": "搬送ロボット2号機"}
            col["allowed"] = ["搬送ロボット1号機", "搬送ロボット2号機"]
            col["normalize"] = ["nfkc", "upper"]
    d["columns"].append({"key": "downtime", "display": "停止時間", "headers": ["停止時間(分)", "停止時間"],
                         "type": "number", "role": "measure", "unit": "分", "unit_conversions": {"h": 60, "時間": 60}})
    d["markdown"] = {"summaries": [{"id": "month", "metrics": ["count", "sum:downtime"], "top_n": 10},
                                   {"id": "entity_fiscal_year", "metrics": ["count", "avg:downtime"]}],
                     "title_columns": ["equipment_name", "symptom"], "file_prefix": "トラブル", "omit_person": True,
                     "dataset_card": False}
    d["header"] = {"rows": 1, "anchors": ["管理No"], "search_rows": 50}
    return spec_from_dict(d)


def test_saving_the_template_editor_unchanged_keeps_json_only_settings(app, client):
    with app.app_context():
        template_id, _ = store.create_template("トラブル対応一覧", _json_only_spec(), "説明")
        before = spec_to_dict(store.get_template(template_id)["spec"])
    body = _collect(client.get(f"/tables/templates/{template_id}").get_data(as_text=True))
    res = client.post(f"/tables/templates/{template_id}", json=body)
    assert res.status_code == 200, res.get_json()
    with app.app_context():
        after = spec_to_dict(store.get_template(template_id)["spec"])

    assert after["columns"] == before["columns"]
    assert after["markdown"]["summaries"] == before["markdown"]["summaries"]
    assert after["markdown"]["dataset_card"] is False
    assert after["markdown"]["title_columns"] == ["equipment_name", "symptom"]
    # search_rows は判定に使っていないので読み捨てる（R6-T-04）。anchors は残る
    assert "search_rows" not in after["header"] and after["header"]["anchors"] == ["管理No"]
    assert after["custom_stages"] == before["custom_stages"]
    assert after["log_stage"] == before["log_stage"]
    assert after["period"]["grain"] == "month"


def test_template_editor_changes_are_saved_and_stale_metrics_dropped(app, client):
    with app.app_context():
        template_id, _ = store.create_template("トラブル対応一覧", _json_only_spec(), "")
    body = _collect(client.get(f"/tables/templates/{template_id}").get_data(as_text=True))
    for row in body["columns"]:
        if row["key"] == "downtime":
            row["use"] = False         # 集計に使っていた列を外す
        if row["key"] == "equipment_name":
            row["type"], row["role"] = "code", "entity"    # 型と役割を変えた列は画面の既定に合わせる
    body["file_prefix"] = "新しい接頭辞"
    res = client.post(f"/tables/templates/{template_id}", json=body)
    assert res.status_code == 200, res.get_json()
    with app.app_context():
        spec = store.get_template(template_id)["spec"]
    assert spec.column("downtime") is None
    assert [s.metrics for s in spec.summaries()] == [["count"], ["count"]]
    assert spec.markdown["file_prefix"] == "新しい接頭辞"
    eq = spec.column("equipment_name")
    assert eq.role == "entity" and eq.value_map == {"ロボ2": "搬送ロボット2号機"}   # 値の置き換えは引き継ぐ


def _upload_csv(client, text: str, name: str) -> int:
    res = client.post("/tables/upload", data={"file": (io.BytesIO(text.encode("cp932")), name)},
                      content_type="multipart/form-data")
    return int(res.headers["Location"].split("/")[3])


def _payload(name: str, headers: dict | None = None) -> dict:
    headers = headers or {}
    return {"name": name, "group_by": "month", "max_records_per_file": 300, "omit_person": True,
            "lightrag_hint": True,
            "columns": [{"index": i, "header": headers.get(k, h), "use": True, "key": k,
                         "display": headers.get(k, h.split("(")[0]), "type": t, "role": r,
                         "unit": "分" if k == "downtime" else "", "md": "attribute", "fill_down_blank": False,
                         "ai": False, "description": ""} for i, (k, h, t, r) in enumerate(COLUMNS)]}


def _wait(app, import_id):
    with app.app_context():
        job = jobs.wait_job(store.get_import(import_id)["job_id"], timeout=60)
    assert job["status"] == "done", job


def test_renamed_header_keeps_the_old_header_in_the_template(app, client):
    first = _upload_csv(client, CSV_TEXT, "a.csv")
    client.post(f"/tables/imports/{first}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "T"})
    client.post(f"/tables/imports/{first}/layout", data={"header_rows": "1"})
    assert client.post(f"/tables/imports/{first}/columns", json=_payload("T")).status_code == 200
    _wait(app, first)
    with app.app_context():
        template_id = store.get_import(first)["template_id"]

    # 新しいファイルで見出し「設備番号」が「機番」に変わった
    second = _upload_csv(client, CSV_TEXT.replace("設備番号", "機番", 1), "b.csv")
    client.post(f"/tables/imports/{second}/source", data={"encoding": "cp932", "delimiter": ",",
                                                          "template": str(template_id)})
    res = client.post(f"/tables/imports/{second}/layout", data={"header_rows": "1"})
    assert res.headers["Location"].endswith("/columns")
    res = client.post(f"/tables/imports/{second}/columns", json=_payload("T", {"equipment_id": "機番"}))
    assert res.status_code == 200, res.get_json()
    _wait(app, second)
    with app.app_context():
        headers = store.get_template(template_id)["spec"].column("equipment_id").headers
    assert headers[0] == "機番" and "設備番号" in headers

    # 前の見出しのファイルも、列の対応づけをやり直さずに読める
    third = _upload_csv(client, CSV_TEXT, "c.csv")
    client.post(f"/tables/imports/{third}/source", data={"encoding": "cp932", "delimiter": ",",
                                                         "template": str(template_id)})
    res = client.post(f"/tables/imports/{third}/layout", data={"header_rows": "1"})
    assert not res.headers["Location"].endswith("/columns")
    _wait(app, third)


# ---- 取り込み設定の JSON の書き出し・読み込み ------------------------------------------------------

def test_export_then_import_under_a_new_name(app, client):
    with app.app_context():
        template_id, _ = store.create_template("トラブル対応一覧", _json_only_spec(), "説明文")
        original = spec_to_dict(store.get_template(template_id)["spec"])
    res = client.get(f"/settings/table-templates/{template_id}/export.json")
    assert res.status_code == 200 and "attachment" in res.headers["Content-Disposition"]
    data = json.loads(res.data)
    assert data["format"] == "rag_maesyori.table_template" and data["name"] == "トラブル対応一覧"

    res = client.post("/settings/table-templates/import", data={"file": (io.BytesIO(res.data), "t.json"),
                                                                "name": "別名の設定"},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and res.headers["Location"].endswith("/settings/table-templates")
    with app.app_context():
        imported = [t for t in store.list_templates() if t["name"] == "別名の設定"]
        assert len(imported) == 1
        spec = spec_to_dict(store.get_template(imported[0]["id"])["spec"])
        assert store.get_template(imported[0]["id"])["description"] == "説明文"
    assert spec.pop("name") == "別名の設定"
    original.pop("name")
    assert spec == original

    # 同じ名前ではもう読み込めない
    res = client.post("/settings/table-templates/import", data={"file": (io.BytesIO(json.dumps(data).encode()), "t.json"),
                                                                "name": "別名の設定"},
                      content_type="multipart/form-data", follow_redirects=True)
    assert "同じ名前の取り込み設定「別名の設定」があります" in res.get_data(as_text=True)
    with app.app_context():
        assert len([t for t in store.list_templates() if t["name"] == "別名の設定"]) == 1
    assert client.get("/settings/table-templates/9999/export.json").status_code == 404


def test_import_rejects_a_file_that_is_not_a_template(client):
    res = client.post("/settings/table-templates/import", data={"file": (io.BytesIO(b"[1, 2]"), "x.json")},
                      content_type="multipart/form-data", follow_redirects=True)
    assert res.status_code == 200
    with client.application.app_context():
        assert store.list_templates() == []


# ---- 帳票の種類: 見本・状態・削除、帳票の JSON と元ファイル ------------------------------------------

def _upload(client, url, path, field, extra=None):
    data = {**(extra or {}), field: (io.BytesIO(path.read_bytes()), path.name)}
    return client.post(url, data=data, content_type="multipart/form-data")


def _make_form_type(app, client, sample_dir) -> int:
    res = _upload(client, "/settings/form-types/new", sample_dir / "standard.xlsx", "samples", {"name": "設備修理報告書"})
    pattern_id = int(re.search(r"/form-types/(\d+)/review", res.headers["Location"]).group(1))
    with app.app_context():
        from excel.workbook import load_workbook_info
        from pattern.builder import suggest_rows
        _, rows = suggest_rows([load_workbook_info(sample_dir / "standard.xlsx")])
    form = {"name": "設備修理報告書", "version": "v1", "mode": "review", "md_options_form": "1",
            "image_processing": "none"}
    for i, r in enumerate(rows):
        for key in ("field_name", "display_name", "candidates", "data_type", "direction", "unit", "rag_output"):
            form[f"fields-{i}-{key}"] = r[key]
        if r["use"]:
            form[f"fields-{i}-use"] = "on"
    assert client.post(f"/settings/form-types/{pattern_id}/save", data=form).status_code == 302
    return pattern_id


def _upload_files(tmp_path) -> set[str]:
    return {p.name for p in (tmp_path / "uploads").rglob("*") if p.is_file()}


def test_form_type_samples_status_and_delete(app, client, sample_dir, tmp_path):
    pattern_id = _make_form_type(app, client, sample_dir)
    other_id = _make_form_type(app, client, sample_dir)

    res = client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "active"}, follow_redirects=True)
    assert "使用を開始しました" in res.get_data(as_text=True)
    with app.app_context():
        assert db.load_pattern(pattern_id).status == "active"
    assert client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "zzz"}).status_code == 400
    assert client.post("/settings/form-types/9999/status", data={"status": "active"}).status_code == 404

    # 見本の追加（Excel でないファイルはエラー）
    res = _upload(client, f"/settings/form-types/{pattern_id}/samples", sample_dir / "shifted.xlsx", "samples")
    assert res.status_code == 302
    assert "見本ファイルを1件追加しました" in client.get(res.headers["Location"]).get_data(as_text=True)
    before = _upload_files(tmp_path)
    res = client.post(f"/settings/form-types/{pattern_id}/samples",
                      data={"samples": (io.BytesIO(b"not excel"), "x.xlsx")},
                      content_type="multipart/form-data", follow_redirects=True)
    page = res.get_data(as_text=True)
    assert "1件目のファイル: Excelファイル" in page and "追加しました" not in page
    assert _upload_files(tmp_path) == before          # 読めなかったファイルは残さない
    with app.app_context():
        samples = db.list_samples(pattern_id)
        other_samples = db.list_samples(other_id)
    assert len(samples) == 2

    # 他の種類の見本は消せない
    assert client.post(f"/settings/form-types/{pattern_id}/samples/{other_samples[0]['id']}/delete").status_code == 404
    assert client.post(f"/settings/form-types/{pattern_id}/samples/9999/delete").status_code == 404
    with app.app_context():
        assert len(db.list_samples(other_id)) == 1

    stored = samples[-1]["stored_path"].split("/")[-1]
    assert stored in _upload_files(tmp_path)
    res = client.post(f"/settings/form-types/{pattern_id}/samples/{samples[-1]['id']}/delete", follow_redirects=True)
    assert "見本ファイルを削除しました" in res.get_data(as_text=True)
    assert stored not in _upload_files(tmp_path)
    with app.app_context():
        assert [s["id"] for s in db.list_samples(pattern_id)] == [samples[0]["id"]]

    res = client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "inactive"}, follow_redirects=True)
    assert "使用を停止しました" in res.get_data(as_text=True)
    with app.app_context():
        assert db.load_pattern(pattern_id).status == "inactive"

    # 種類を削除すると見本ファイルも消える
    remaining = samples[0]["stored_path"].split("/")[-1]
    res = client.post(f"/settings/form-types/{pattern_id}/delete", follow_redirects=True)
    assert "を削除しました" in res.get_data(as_text=True)
    assert remaining not in _upload_files(tmp_path)
    with app.app_context():
        assert db.list_samples(pattern_id) == []
        assert len(db.list_samples(other_id)) == 1
    assert client.post(f"/settings/form-types/{pattern_id}/delete").status_code == 404


def test_form_json_and_original_downloads(app, client, sample_dir):
    pattern_id = _make_form_type(app, client, sample_dir)
    client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "active"})
    docs = []
    for _ in range(2):
        res = _upload(client, "/forms/upload", sample_dir / "standard.xlsx", "file")
        doc_id = int(re.search(r"/forms/(\d+)/type", res.headers["Location"]).group(1))
        client.post(f"/forms/{doc_id}/read", data={"pattern_id": pattern_id, "sheets": ["修理報告書"]})
        docs.append(doc_id)
    confirmed, working = docs
    assert client.post(f"/forms/{confirmed}/confirm", data={}).status_code == 302

    res = client.get(f"/forms/{confirmed}/download.json")
    assert res.status_code == 200 and res.mimetype == "application/json"
    assert "attachment" in res.headers["Content-Disposition"]
    body = json.loads(res.data)
    assert body["source"]["file_name"] == "standard.xlsx" and body["source"]["sheets"] == ["修理報告書"]
    assert client.get(f"/forms/{working}/download.json").status_code == 404    # 確定前は出さない

    res = client.get(f"/forms/{working}/original")
    assert res.status_code == 200 and res.data == (sample_dir / "standard.xlsx").read_bytes()
    assert "standard.xlsx" in res.headers["Content-Disposition"]
    with app.app_context():
        assert db.get_document(working) is not None     # 元ファイルのダウンロードでは消さない
    assert client.get("/forms/9999/original").status_code == 404


# ---- 一覧表: Excel のシートの選択 ---------------------------------------------------------------

def test_save_source_uses_the_chosen_excel_sheet(app, client, tmp_path):
    book = openpyxl.Workbook()
    memo = book.active
    memo.title = "メモ"
    memo["A1"] = "このブックの説明"
    sheet = book.create_sheet("一覧")
    sheet.append(["管理No", "発生日", "現象"])
    for i in range(1, 6):
        sheet.append([f"TR-{i:03d}", f"2026-08-{i:02d}", f"アラーム停止{i}"])
    path = tmp_path / "二つのシート.xlsx"
    book.save(path)

    res = client.post("/tables/upload", data={"file": (io.BytesIO(path.read_bytes()), path.name)},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    page = client.get(f"/tables/imports/{import_id}/source").get_data(as_text=True)
    assert "メモ" in page and "一覧" in page

    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"sheet": "一覧", "template": "new", "new_template_name": "シート選択"})
    assert res.headers["Location"].endswith("/layout")
    with app.app_context():
        assert store.get_import(import_id)["source"]["sheet"] == "一覧"
    page = client.get(f"/tables/imports/{import_id}/layout").get_data(as_text=True)
    assert "TR-001" in page and "このブックの説明" not in page
    detect = client.post(f"/tables/imports/{import_id}/layout/detect", json={"header_rows": [1]}).get_json()
    assert detect["data_end"] == 6
    res = client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1"})
    assert res.headers["Location"].endswith("/columns")
    assert "発生日" in client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True)

    # シートを選び直すと、前のシートで決めた見出し行・範囲は使わない
    client.post(f"/tables/imports/{import_id}/source",
                data={"sheet": "メモ", "template": "new", "new_template_name": "シート選択"})
    with app.app_context():
        src = store.get_import(import_id)["source"]
    assert src["sheet"] == "メモ" and "header_rows" not in src
    assert "このブックの説明" in client.get(f"/tables/imports/{import_id}/layout").get_data(as_text=True)


def test_form_type_section_is_saved_from_the_editor_and_shown_again(app, client, sample_dir):
    """画面で入れた「探す区画」が保存され、編集画面を開き直しても残る（保存で消えると F5 の回答欄の指定が効かない）。"""
    res = _upload(client, "/settings/form-types/new", sample_dir / "standard.xlsx", "samples", {"name": "設備修理報告書"})
    pattern_id = int(re.search(r"/form-types/(\d+)/review", res.headers["Location"]).group(1))
    with app.app_context():
        from excel.workbook import load_workbook_info
        from pattern.builder import suggest_rows
        _, rows = suggest_rows([load_workbook_info(sample_dir / "standard.xlsx")])
    form = {"name": "設備修理報告書", "version": "v1", "mode": "review", "md_options_form": "1", "image_processing": "none"}
    for i, r in enumerate(rows):
        for key in ("field_name", "display_name", "candidates", "data_type", "direction", "unit", "rag_output"):
            form[f"fields-{i}-{key}"] = r[key]
        if r["use"]:
            form[f"fields-{i}-use"] = "on"
    first = next(i for i, r in enumerate(rows) if r["use"])
    form[f"fields-{first}-section"] = "▼ 回答欄（宛先記入）"
    assert client.post(f"/settings/form-types/{pattern_id}/save", data=form).status_code == 302
    with app.app_context():
        saved = {f.field_name: f.section for f in db.load_pattern(pattern_id).fields}
    assert saved[rows[first]["field_name"]] == "回答"
    assert sum(1 for v in saved.values() if v) == 1
    page = client.get(f"/settings/form-types/{pattern_id}/edit").get_data(as_text=True)
    assert f'name="fields-{first}-section" value="回答"' in page
