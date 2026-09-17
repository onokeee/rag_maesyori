"""一覧表フローの画面（CSV を選ぶ→設定→範囲→列→AI整形画面→確認→確定→zip）の通し確認。"""
import io
import zipfile

from core import jobs
from tables import store

CSV_TEXT = "管理No,発生日,設備番号,設備名,現象,対応内容,停止時間(分),担当者\r\n" + "".join(
    f'TR-{i:03d},2026-08-{i:02d},EQ-{i % 3 + 1:02d},搬送ロボット{i % 3 + 1}号機,アラーム停止{i},'
    f'"8/{i} 10:00 田中: 停止の連絡あり。\n8/{i} 11:00 佐藤: 再起動で復旧。",{i * 5},田中\r\n'
    for i in range(1, 13))

COLUMNS = [
    ("record_no", "管理No", "code", "key"), ("occurred_at", "発生日", "date", "date"),
    ("equipment_id", "設備番号", "code", "entity"), ("equipment_name", "設備名", "string", "entity_label"),
    ("symptom", "現象", "text", "text"), ("response_log", "対応内容", "text", "log"),
    ("downtime", "停止時間", "number", "measure"), ("worker", "担当者", "string", "person"),
]


def _wait_import_job(app, import_id):
    with app.app_context():
        job = jobs.wait_job(store.get_import(import_id)["job_id"], timeout=60)
        assert job["status"] == "done", job
        return store.get_import(import_id)


def test_csv_import_flow(app, client):
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "トラブル一覧.csv")},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and "/source" in res.headers["Location"]
    import_id = int(res.headers["Location"].split("/")[3])

    assert client.get(f"/tables/imports/{import_id}/source").status_code == 200
    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "トラブル対応一覧"})
    assert res.headers["Location"].endswith("/layout")
    page = client.get(f"/tables/imports/{import_id}/layout")
    assert page.status_code == 200 and "row-header" in page.get_data(as_text=True)
    detect = client.post(f"/tables/imports/{import_id}/layout/detect", json={"header_rows": [1]}).get_json()
    assert detect["data_end"] == 13 and detect["table_kind"] == "list"
    res = client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    assert res.headers["Location"].endswith("/columns")

    assert "対応内容" in client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True)
    payload = {"name": "トラブル対応一覧", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                            "fill_down_blank": False, "ai": r == "log", "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    res = client.post(f"/tables/imports/{import_id}/columns", json=payload)
    assert res.status_code == 200, res.get_json()
    assert res.get_json()["redirect"].endswith("/ai")
    imp = _wait_import_job(app, import_id)
    assert imp["status"] == "preview" and imp["stats"]["records"] == 12

    assert client.get(f"/api/jobs/{imp['job_id']}").get_json()["status"] == "done"
    assert "分割プレビュー" in client.get(f"/tables/imports/{import_id}/ai").get_data(as_text=True)
    split = client.post(f"/tables/imports/{import_id}/ai/split-preview", json={"row_key": "TR-001"}).get_json()
    assert len(split["segments"]) == 2

    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "確定してMarkdownを作成" in page and "トラブル対応一覧_2026-08.md" in page
    text = client.get(f"/tables/imports/{import_id}/preview/file?name=トラブル対応一覧_2026-08.md").get_json()["text"]
    assert "TR-001" in text and "対応の時系列" in text

    res = client.post(f"/tables/imports/{import_id}/confirm")
    assert res.headers["Location"].endswith("/done")
    imp = _wait_import_job(app, import_id)
    assert imp["status"] == "confirmed"
    assert "Markdownをまとめてダウンロード" in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)

    zdata = client.get(f"/tables/imports/{import_id}/download.zip").data
    names = zipfile.ZipFile(io.BytesIO(zdata)).namelist()
    assert "RAG投入用/トラブル対応一覧_2026-08.md" in names
    assert {"管理用_RAGには入れない/正規化データ.csv", "管理用_RAGには入れない/問題一覧.csv",
            "管理用_RAGには入れない/取込レポート.csv"} <= set(names)
    assert client.get(f"/tables/imports/{import_id}/normalized.csv").status_code == 200

    # 2回目: 同じ設定を選べば列の対応づけを飛ばす
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "トラブル一覧2.csv")},
                      content_type="multipart/form-data")
    second = int(res.headers["Location"].split("/")[3])
    tid = imp["template_id"]
    assert "必須列" in client.get(f"/tables/imports/{second}/source").get_data(as_text=True)
    client.post(f"/tables/imports/{second}/source", data={"encoding": "cp932", "delimiter": ",", "template": str(tid)})
    res = client.post(f"/tables/imports/{second}/layout", data={"header_rows": "1"})
    assert res.headers["Location"].endswith("/ai")
    assert _wait_import_job(app, second)["status"] == "preview"
    assert client.get(f"/tables/templates/{tid}").status_code == 200
    assert client.get("/settings/table-templates").status_code == 200
