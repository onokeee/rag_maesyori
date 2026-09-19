"""一覧表フローの画面（CSV を選ぶ→設定→範囲→列→AI整形画面→確認→確定→zip）の通し確認。"""
import io
import zipfile

import pytest

from core import jobs
from tables import pipeline, store

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


def _preview_page(app, client, import_id) -> str:
    """確認画面。初回は md を作るジョブの待ち画面が出るので、終わってから開き直す。"""
    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    if "Markdownの下書きを作っています" in page:
        with app.app_context():
            job = jobs.wait_job(jobs.latest_job("table_import", import_id, kind="table_preview")["id"], timeout=60)
            assert job["status"] == "done", job
        page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    return page


def test_csv_import_flow(app, client, monkeypatch):
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "トラブル一覧.csv")},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and "/source" in res.headers["Location"]
    import_id = int(res.headers["Location"].split("/")[3])

    page = client.get(f"/tables/imports/{import_id}/source")
    assert page.status_code == 200
    # 取り込み設定の名前は残り続けるので、アップロードしたファイル名を初期値にしない（design.md 3.3）
    assert 'name="new_template_name" class="input" value=""' in page.get_data(as_text=True)
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
    # 2回目からは取り込みの控えで画面を作る（CSV を開き直さない）
    monkeypatch.setattr(pipeline, "open_import_source", lambda *a, **k: pytest.fail("CSV を開き直した"))
    page = client.get(f"/tables/imports/{import_id}/layout")
    assert page.status_code == 200 and "row-header" in page.get_data(as_text=True)
    assert "対応内容" in client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True)
    assert client.get(f"/tables/imports/{import_id}/source").status_code == 200
    monkeypatch.undo()
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

    page = _preview_page(app, client, import_id)
    assert "確定してMarkdownを作成" in page and "トラブル対応一覧_2026-08.md" in page
    text = client.get(f"/tables/imports/{import_id}/preview/file?name=トラブル対応一覧_2026-08.md").get_json()["text"]
    assert "TR-001" in text and "対応の時系列" in text

    res = client.post(f"/tables/imports/{import_id}/confirm")
    assert res.headers["Location"].endswith("/done")
    imp = _wait_import_job(app, import_id)
    assert imp["status"] == "confirmed"
    assert "Markdownをまとめてダウンロード" in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)

    # 正規化CSV は単独でも取れる（zip の 管理用_RAGには入れない フォルダにも入る）。
    # zip をダウンロードするとこの取り込みのデータは消えるので、単独で取るのは zip の前だけ
    assert client.get(f"/tables/imports/{import_id}/normalized.csv").status_code == 200
    zdata = client.get(f"/tables/imports/{import_id}/download.zip").data
    names = zipfile.ZipFile(io.BytesIO(zdata)).namelist()
    assert "RAG投入用/トラブル対応一覧_2026-08.md" in names
    assert {"管理用_RAGには入れない/正規化データ.csv", "管理用_RAGには入れない/問題一覧.csv",
            "管理用_RAGには入れない/取込レポート.csv"} <= set(names)
    with app.app_context():
        assert store.get_import(import_id) is None      # ダウンロードしたら残さない
        assert store.get_template(imp["template_id"]) is not None   # 取り込み設定は残る

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


# ---- 読み込み中の中止（待ち画面の［中止］） ------------------------------------------------

def _pending_import(app, status="reading"):
    """読み込み中に見える取り込みを1件つくる（ジョブは待機中のまま）。"""
    from models import database
    with app.app_context():
        import_id = store.create_import("大きい表.csv", "h" * 64, "uploads/大きい表.csv", {"encoding": "utf-8"})
        conn = database.get_db()
        ts = database.now()
        cur = conn.execute("""INSERT INTO jobs (kind, ref_type, ref_id, status, params_json, progress_json,
                              created_at, updated_at) VALUES ('table_read', 'table_import', ?, 'queued', '{}', '{}', ?, ?)""",
                           (import_id, ts, ts))
        conn.commit()
        store.update_import(import_id, status=status, job_id=cur.lastrowid)
        return import_id, cur.lastrowid


def test_wait_screen_offers_cancel(app, client):
    import_id, _job_id = _pending_import(app)
    page = client.get(f"/tables/imports/{import_id}/preview").get_data(as_text=True)
    assert "表を読み込んでいます" in page
    assert f'action="/tables/imports/{import_id}/cancel"' in page and ">中止<" in page


def test_cancel_read_job_returns_to_layout(app, client):
    import_id, job_id = _pending_import(app)
    res = client.post(f"/tables/imports/{import_id}/cancel")
    assert res.status_code == 302
    with app.app_context():
        assert jobs.get_job(job_id)["status"] == "cancelled"
        # 待機中のまま中止したときは取り込みの状態も戻す（読み込み中のままにしない）
        assert store.get_import(import_id)["status"] == "uploaded"
    # 読み込み中の待ち画面ではなく、選び直せる画面に戻る
    assert client.get(f"/tables/imports/{import_id}/preview").status_code == 302


def test_cancel_render_job_returns_to_preview(app, client):
    from models import database

    import_id, job_id = _pending_import(app, status="confirming")
    with app.app_context():
        conn = database.get_db()
        conn.execute("UPDATE jobs SET kind = 'table_render' WHERE id = ?", (job_id,))
        conn.commit()
    client.post(f"/tables/imports/{import_id}/cancel")
    with app.app_context():
        assert jobs.get_job(job_id)["status"] == "cancelled"
        assert store.get_import(import_id)["status"] == "preview"


def test_cancel_without_running_job_is_reported(app, client):
    with app.app_context():
        import_id = store.create_import("表.csv", "h" * 64, "uploads/表.csv", {})
    assert client.post(f"/tables/imports/{import_id}/cancel").status_code == 302
    with client.session_transaction() as session:
        assert any("中止できる処理がありません" in text for _level, text in session["_flashes"])


# ---- 範囲の確認画面の表示 ----------------------------------------------------------------

def test_layout_screen_without_headers(app, client):
    res = client.post("/tables/upload", data={"file": (io.BytesIO("あ\r\nい\r\n".encode("utf-8")), "1列.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "utf-8", "delimiter": ",", "template": "new", "new_template_name": "1列"})
    page = client.get(f"/tables/imports/{import_id}/layout").get_data(as_text=True)
    # 見出し行が見つからないときに「1〜0行目」のような存在しない行番号を出さない
    assert "見出し行が見つかりません" in page
    assert "データの行が見つかりません" in page and "〜0行目" not in page
    # 見出し行を指定すれば進める（［この範囲で次へ］は押せるまま）
    assert 'role="radiogroup"' in page
    res = client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    assert res.status_code == 302 and res.headers["Location"].endswith("/columns")


# ---- 取り込みの削除・画面の表示 ------------------------------------------------------------

def _uploaded_csv(client, name="間違い.csv"):
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), name)},
                      content_type="multipart/form-data")
    return int(res.headers["Location"].split("/")[3])


def test_delete_import_removes_it_from_home(app, client):
    """間違えて取り込んだファイルを消せる（ホームの「作業中の一覧表」に残り続けない）。"""
    import_id = _uploaded_csv(client)
    assert f'/tables/imports/{import_id}/delete' in client.get(f"/tables/imports/{import_id}/layout").get_data(as_text=True)
    with app.app_context():
        base = pipeline.import_dir(import_id)
        base.mkdir(parents=True, exist_ok=True)
        (base / "rows.jsonl.gz").write_bytes(b"x")
        stored = store.get_import(import_id)["stored_path"]
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 302
    with app.app_context():
        from core.files import upload_path
        assert store.get_import(import_id) is None
        assert not pipeline.import_dir(import_id).exists()
        assert not upload_path(stored).exists()
    assert str(import_id) not in [str(i["id"]) for i in _recent_ids(app)]


def _recent_ids(app):
    with app.app_context():
        return store.list_imports(limit=10)


def test_delete_is_refused_while_a_job_is_running(app, client):
    import_id, _job_id = _pending_import(app)
    assert client.post(f"/tables/imports/{import_id}/delete").status_code == 302
    with app.app_context():
        assert store.get_import(import_id) is not None


def test_source_screen_rejects_an_encoding_that_is_not_offered(app, client):
    """画面にない文字コードを保存させない（保存できると読み込みで落ちて画面が開けなくなる）。"""
    import_id = _uploaded_csv(client, "文字コード.csv")
    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"encoding": "rot13", "delimiter": ",", "template": "new", "new_template_name": "x"})
    assert res.headers["Location"].endswith("/source")
    with app.app_context():
        assert store.get_import(import_id)["source"].get("encoding") != "rot13"
    # 画面にある文字コードは保存できる
    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"encoding": "utf-16", "delimiter": ",", "template": "new", "new_template_name": "x"})
    assert res.headers["Location"].endswith("/layout")
    # 中身と合わない文字コードでも、500 ではなく案内を出して選び直せる
    assert client.get(f"/tables/imports/{import_id}/source").status_code == 200
    assert client.get(f"/tables/imports/{import_id}/layout").status_code in (200, 302)


def test_column_editor_shows_japanese_type_names(app, client):
    """型が合っていないかもしれないという注意書きを、英語キー（date / number / enum）のまま出さない。"""
    import re

    text = "管理No,発生日,設備番号,設備名,状態コード,対応内容,停止時間\r\n" + "".join(
        f"MS-{i:04d},2026-08-{i:02d},EQ-01,搬送ロボット1号機,{i % 3},点検した,{i * 5}\r\n" for i in range(1, 29))
    res = client.post("/tables/upload", data={"file": (io.BytesIO(text.encode("cp932")), "型.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "型"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    page = client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True)
    shown = re.findall(r"値は「([^」]+)」らしい", page)
    assert shown, "型が合っていない列の注意書きが出ていない"
    assert not ({"date", "number", "enum", "string", "text", "code"} & set(shown)), shown


def test_preview_screen_says_it_is_already_confirmed(app, client):
    """確定済みの取り込みを開き直したときに、黙って作り直せるボタンだけを出さない。"""
    import_id = _uploaded_csv(client, "確定済み.csv")
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "確定済み"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "確定済み", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                            "fill_down_blank": False, "ai": False, "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)
    _preview_page(app, client, import_id)
    client.post(f"/tables/imports/{import_id}/confirm")
    assert _wait_import_job(app, import_id)["status"] == "confirmed"
    page = _preview_page(app, client, import_id)
    assert "この取り込みは確定済みです" in page
    assert "確定し直してMarkdownを作り直す" in page and "data-confirm" in page


def test_source_screen_marks_a_crosstab_sheet(app, client):
    """クロス集計のシートに「一覧表らしい」と出さない（3画面目で初めて止められると分かりにくい）。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "FY2023"
    ws.append(["設備"] + [f"2023-{m:02d}" for m in range(4, 13)] + [f"2024-{m:02d}" for m in range(1, 4)])
    for i in range(1, 16):
        ws.append([f"EQ-{i:02d}"] + [i * m for m in range(1, 13)])
    buf = io.BytesIO()
    wb.save(buf)
    res = client.post("/tables/upload", data={"file": (io.BytesIO(buf.getvalue()), "月別集計.xlsx")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    page = client.get(f"/tables/imports/{import_id}/source").get_data(as_text=True)
    assert "クロス集計らしい（対応していません）" in page and "・一覧表らしい" not in page
    # 3画面目の判定は今までどおり（この範囲では進めない）
    res = client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    assert res.status_code == 302 and res.headers["Location"].endswith("/layout")


def test_preview_warns_when_the_filename_hint_is_off(app, client):
    """ヒント無しのまま確定すると LightRAG 側で記録が切られるので、確認画面で知らせる。"""
    import_id = _uploaded_csv(client, "ヒント.csv")
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "ヒント"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "ヒント", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                            "fill_down_blank": False, "ai": False, "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)
    assert "ファイル名に LightRAG の分割ヒントを付けていません" in _preview_page(app, client, import_id)

    payload["lightrag_hint"] = True
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)
    assert "ファイル名に LightRAG の分割ヒントを付けていません" not in _preview_page(app, client, import_id)


def test_column_editor_shows_the_new_markdown_options(app, client):
    """列の対応づけ画面に、分割ヒントと「対応の時系列から、他の列と同じ内容の文を省く」のチェックを出す。"""
    import_id = _uploaded_csv(client, "出し方.csv")
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "出し方"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    page = client.get(f"/tables/imports/{import_id}/columns").get_data(as_text=True)
    assert "対応の時系列から、他の列と同じ内容の文を省く" in page
    assert "ファイル名に LightRAG の分割ヒントを付ける（推奨・既定はオン）" in page
    # 新しい設定では、どちらのチェックも最初からオンで出す（design.md 5.3 の markdown 既定）
    assert "checked" in page.split('data-setting="dedupe_timeline"')[1][:12]
    assert "checked" in page.split('data-setting="lightrag_hint"')[1][:12]


def test_markdown_options_are_saved_and_shown_again(app, client):
    """画面で外したチェックが取り込み設定に保存され、編集画面を開き直しても外れたまま出る。"""
    import_id = _uploaded_csv(client, "出し方2.csv")
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "出し方2"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "出し方2", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "lightrag_hint": False, "dedupe_timeline": False,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                            "role": r, "unit": "", "md": "attribute", "fill_down_blank": False, "ai": False,
                            "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)

    with app.app_context():
        imp = store.get_import(import_id)
        spec = pipeline.spec_for_import(imp)
        template_id = imp["template_id"]
    assert spec.markdown["lightrag_hint"] is False and spec.markdown["dedupe_timeline"] is False
    page = client.get(f"/tables/templates/{template_id}").get_data(as_text=True)
    assert "checked" not in page.split('data-setting="dedupe_timeline"')[1][:12]
    assert "checked" not in page.split('data-setting="lightrag_hint"')[1][:12]


# ---- 処理中の変更・削除を止める（R1） -----------------------------------------------------------

def _queued_job(app, import_id, kind):
    from models import database
    with app.app_context():
        conn = database.get_db()
        ts = database.now()
        cur = conn.execute("""INSERT INTO jobs (kind, ref_type, ref_id, status, params_json, progress_json,
                              created_at, updated_at) VALUES (?, 'table_import', ?, 'queued', '{}', '{}', ?, ?)""",
                           (kind, import_id, ts, ts))
        conn.commit()
        return cur.lastrowid


def test_source_is_not_changed_while_reading(app, client):
    """読み込み中に2画面目を保存しても、文字コード・区切り文字を書き換えない（古い設定の読み込み結果で確定させない）。"""
    import_id, _job_id = _pending_import(app)
    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"encoding": "utf-8", "delimiter": ";", "template": "new", "new_template_name": "x"})
    assert res.status_code == 302 and res.headers["Location"].endswith("/preview")
    res = client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "2"})
    assert res.headers["Location"].endswith("/preview")
    assert client.post(f"/tables/imports/{import_id}/columns", json={"name": "x", "columns": []}).status_code == 409
    with app.app_context():
        imp = store.get_import(import_id)
        assert imp["status"] == "reading" and imp["source"] == {"encoding": "utf-8"}


def test_utf16_without_bom_can_pass_the_source_screen(app, client):
    """BOMなしの UTF-16 と判定された CSV も、2画面目をそのまま保存して先へ進める。"""
    data = CSV_TEXT.encode("utf-16-le")
    res = client.post("/tables/upload", data={"file": (io.BytesIO(data), "u16.csv")}, content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    with app.app_context():
        assert store.get_import(import_id)["source"]["encoding"] == "utf-16-le"
    page = client.get(f"/tables/imports/{import_id}/source").get_data(as_text=True)
    assert 'value="utf-16-le" selected' in page or 'value="utf-16-le"' in page
    res = client.post(f"/tables/imports/{import_id}/source",
                      data={"encoding": "utf-16-le", "delimiter": ",", "template": "new", "new_template_name": "u16"})
    assert res.headers["Location"].endswith("/layout")
    assert client.get(f"/tables/imports/{import_id}/layout").status_code == 200


def test_delete_and_confirm_are_refused_while_ai_formatting_runs(app, client, monkeypatch):
    """AI整形の実行中に削除すると、動いている呼び出しが消したあとの DB に結果を書き戻すので、止める。"""
    from types import SimpleNamespace

    from views import tables as tables_view

    import_id = _uploaded_csv(client, "AI中.csv")
    with app.app_context():
        store.update_import(import_id, status="preview")
    _queued_job(app, import_id, "ai_format")
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 302 and res.headers["Location"].endswith("/ai")
    with app.app_context():
        assert store.get_import(import_id) is not None

    monkeypatch.setattr(tables_view, "_spec_for", lambda imp: SimpleNamespace(log_stage=object()))
    monkeypatch.setattr(pipeline, "start_render_job", lambda *a, **k: pytest.fail("AI整形の実行中に確定した"))
    res = client.post(f"/tables/imports/{import_id}/confirm")
    assert res.headers["Location"].endswith("/ai")


def test_trial_run_needs_the_external_confirmation(app, client, monkeypatch):
    """試し実行も、外部の AI に送るときは画面の確認のチェックがなければ送らない（サーバー側で確かめる）。"""
    from types import SimpleNamespace

    from aiproc import runner
    from services import llm
    from views import tables as tables_view

    import_id = _uploaded_csv(client, "試し.csv")
    monkeypatch.setattr(tables_view, "_spec_for", lambda imp: SimpleNamespace(log_stage=object()))
    monkeypatch.setattr(llm, "job_client_settings", lambda *a, **k: {"chat_url": "https://api.example.com/v1/chat"})
    monkeypatch.setattr(runner, "trial_row", lambda *a, **k: pytest.fail("確認なしで外部に送った"))
    monkeypatch.setattr(runner, "load_rows_for_ai", lambda *a, **k: pytest.fail("確認なしで外部に送った"))
    res = client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 400 and "外部のAIサービス" in res.get_json()["error"]


def test_workbook_without_sheets_is_rejected(app, client, tmp_path):
    """シートのないブックは取り込みにせず、選び直しの案内を出す（2画面目が 500 にならない）。"""
    import re

    from openpyxl import Workbook

    wb = Workbook()
    wb.save(tmp_path / "s.xlsx")
    src = zipfile.ZipFile(tmp_path / "s.xlsx")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = re.sub(rb"<sheets>.*?</sheets>", b"<sheets/>", data, flags=re.S)
            z.writestr(item, data)
    res = client.post("/tables/upload", data={"file": (io.BytesIO(out.getvalue()), "シートなし.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 302 and res.headers["Location"].endswith("/tables/new")
    with app.app_context():
        assert store.list_imports(limit=10) == []


@pytest.mark.parametrize("url, target", [("/tables/upload", "views.tables.open_source"),
                                         ("/forms/upload", "views.forms.precheck_excel")])
def test_unexpected_error_during_upload_leaves_no_file(app, client, monkeypatch, tmp_path, url, target):
    """読み込みの途中で思わぬエラーが出ても、アップロードしたファイルは残さない（design.md 3.3）。"""
    from pathlib import Path

    from openpyxl import Workbook

    def boom(*args, **kwargs):
        raise RuntimeError("想定外")

    monkeypatch.setattr(target, boom)
    wb = Workbook()
    wb.active["A1"] = "管理No"
    wb.save(tmp_path / "a.xlsx")
    with pytest.raises(RuntimeError):
        client.post(url, data={"file": (io.BytesIO((tmp_path / "a.xlsx").read_bytes()), "a.xlsx")},
                    content_type="multipart/form-data")
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []
