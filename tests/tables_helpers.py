"""表の取り込み（1画面）のテスト用ヘルパー。

画面は /tables の1枚だけで、段（panel）の中身と保存はすべて fetch でやりとりする
（views/tables.py・static/tables.js）。テストも同じ JSON のやりとりで進める。
"""
from __future__ import annotations

import copy
import io
from html.parser import HTMLParser

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


# ---- ファイルを置く ---------------------------------------------------------------------

def upload(client, data: bytes, name: str):
    """POST /tables/upload（生の応答を返す。断られる場合の確認に使う）。"""
    return client.post("/tables/upload", data={"file": (io.BytesIO(data), name)},
                       content_type="multipart/form-data")


def upload_csv(client, name: str = "一覧.csv", text: str = CSV_TEXT, encoding: str = "cp932") -> int:
    res = upload(client, text.encode(encoding), name)
    assert res.status_code == 200, res.get_json()
    return res.get_json()["import_id"]


def upload_bytes(client, data: bytes, name: str) -> int:
    res = upload(client, data, name)
    assert res.status_code == 200, res.get_json()
    return res.get_json()["import_id"]


# ---- 段の中身（panel） ------------------------------------------------------------------

def panel(client, import_id: int, name: str, **query) -> dict:
    res = client.get(f"/tables/imports/{import_id}/panel/{name}", query_string=query or None)
    assert res.status_code == 200, res.status_code
    return res.get_json()


def panel_html(client, import_id: int, name: str, **query) -> str:
    return panel(client, import_id, name, **query)["html"]


# ---- 保存（JSON を返す POST） -------------------------------------------------------------

def save_source(client, import_id: int, **fields):
    return client.post(f"/tables/imports/{import_id}/source", json=fields)

def name_source(client, import_id: int, template_name: str, **fields):
    """新しい取り込み設定を作る、いつもの読み取り方の保存（CSV は cp932・カンマ）。"""
    payload = {"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": template_name}
    payload.update(fields)
    res = save_source(client, import_id, **payload)
    assert res.status_code == 200, res.get_json()
    return res.get_json()


def save_layout(client, import_id: int, header_rows="1", data_end_row=""):
    return client.post(f"/tables/imports/{import_id}/layout",
                       json={"header_rows": header_rows, "data_end_row": data_end_row})


def columns_payload(name: str, columns=COLUMNS, ai_role: str | None = None, **extra) -> dict:
    """列の対応づけの保存に送る JSON。ai_role に役割名を渡すと、その列を AI整形の対象にする。"""
    payload = {
        "name": name, "group_by": "month", "description": "", "file_prefix": "",
        "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h.split("(")[0], "type": t,
                     "role": r, "unit": "分" if k == "downtime" else "", "md": "attribute",
                     "fill_down_blank": False, "ai": r == ai_role, "description": ""}
                    for i, (k, h, t, r) in enumerate(columns)],
    }
    payload.update(extra)
    return payload


def save_columns(client, import_id: int, payload: dict):
    return client.post(f"/tables/imports/{import_id}/columns", json=payload)


# ---- ジョブの待ち合わせ ------------------------------------------------------------------

def wait_import_job(app, import_id: int) -> dict:
    """読み込み・確定のジョブ（取り込みの job_id）が終わるまで待つ。"""
    with app.app_context():
        job = jobs.wait_job(store.get_import(import_id)["job_id"], timeout=60)
        assert job["status"] == "done", job
        return store.get_import(import_id)


def preview_panel(app, client, import_id: int, **query) -> dict:
    """「内容の確認」の段。初回は md の下書きを作るジョブが動くので、終わってから取り直す。"""
    data = panel(client, import_id, "preview", **query)
    if data.get("building") and data.get("job"):
        with app.app_context():
            job = jobs.wait_job(data["job"]["id"], timeout=60)
            assert job["status"] == "done", job
        data = panel(client, import_id, "preview", **query)
    return data


# ---- 通しで「内容の確認」まで進める -------------------------------------------------------

def imported(app, client, name: str, template_name: str, text: str = CSV_TEXT, ai_role: str | None = None,
             **payload_extra) -> int:
    """CSV を置いて読み取り方・範囲・列を保存し、読み込みが終わった取り込みを作る。"""
    import_id = upload_csv(client, name, text)
    name_source(client, import_id, template_name)
    res = save_layout(client, import_id)
    assert res.status_code == 200, res.get_json()
    payload = columns_payload(template_name, ai_role=ai_role, **payload_extra)
    res = save_columns(client, import_id, payload)
    assert res.status_code == 200, res.get_json()
    wait_import_job(app, import_id)
    return import_id


def confirmed(app, client, name: str, template_name: str, **kwargs) -> int:
    """確定まで済ませた取り込み（zip をダウンロードできる状態）。"""
    import_id = imported(app, client, name, template_name, **kwargs)
    preview_panel(app, client, import_id)
    res = client.post(f"/tables/imports/{import_id}/confirm")
    assert res.status_code == 200, res.get_json()
    assert wait_import_job(app, import_id)["status"] == "confirmed"
    return import_id


# ---- 「列の対応づけ」の段を、画面と同じ形で読み書きする -------------------------------------
# static/tables.js の collect() と同じ値を、段の HTML から集める。

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


def collect_editor(html: str) -> dict:
    p = _EditorParser()
    p.feed(html)
    return {**p.settings, "header_rows_count": p.header_rows, "columns": p.cols}


def json_only_spec():
    """JSON で読み込んだときにだけ付く設定（見出しの別名・単位換算・値の置き換えなど）を持つ取り込み設定。"""
    from tables.spec import spec_from_dict
    from tests.test_aiproc import SPEC

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
                     "title_columns": ["equipment_name", "symptom"], "file_prefix": "トラブル", "dataset_card": False}
    d["header"] = {"rows": 1, "anchors": ["管理No"], "search_rows": 50}
    return spec_from_dict(d)


# json_only_spec の列がそのまま当たる CSV（「列の対応づけ」の段をその設定で開くため）
SPEC_CSV = "管理No,発生日,設備名,現象,原因,対応内容,担当,停止時間(分)\r\n" + "".join(
    f'TR-{i:03d},2026-08-{i:02d},搬送ロボット{i % 2 + 1}号機,アラーム停止{i},摩耗,'
    f'"8/{i} 10:00 田中: 確認した。",田中,{i * 5}\r\n' for i in range(1, 13))


def template_import(app, client, spec, name: str = "編集.csv", text: str = SPEC_CSV) -> tuple[int, int]:
    """spec を取り込み設定として保存し、その設定で読む取り込みを1件作る。戻り値: (import_id, template_id)"""
    with app.app_context():
        template_id, _version_id = store.create_template(spec.name, spec, "")
    import_id = upload_csv(client, name, text)
    res = save_source(client, import_id, encoding="cp932", delimiter=",", template=str(template_id))
    assert res.status_code == 200, res.get_json()
    res = save_layout(client, import_id)
    assert res.status_code == 200, res.get_json()
    if res.get_json().get("job"):
        wait_import_job(app, import_id)   # 読み込み中は段が待ち画面になるので、終わらせてから開く
    return import_id, template_id


def editor_body(client, import_id: int) -> dict:
    """「列の対応づけ」の段が出す表を、画面が送る形にして返す。"""
    return collect_editor(panel_html(client, import_id, "columns"))
