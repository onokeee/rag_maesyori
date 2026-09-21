"""pytest の共通部品: fixture・偽の OpenAI 互換サーバー（旧 fake_servers）・表の取り込みの操作（旧 tables_helpers）。"""
from __future__ import annotations

import io
import json
import re
import threading
import time
from dataclasses import dataclass, field
from html.parser import HTMLParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest
from flask.testing import FlaskClient

import core
import tables
from app import create_app
from forms import load_workbook_info
from scripts.make_samples import make_inspection, make_repair_shifted, make_repair_standard, make_repair_table



# ====================================================================================================
# 元 tests/conftest.py
# ====================================================================================================

@pytest.fixture(scope="session")
def sample_dir(tmp_path_factory) -> Path:
    directory = tmp_path_factory.mktemp("samples")
    make_repair_standard(directory / "standard.xlsx")
    make_repair_shifted(directory / "shifted.xlsx")
    make_repair_table(directory / "table.xlsx")
    make_inspection(directory / "inspection.xlsx")
    return directory


@pytest.fixture(scope="session")
def repair_infos(sample_dir):
    return [load_workbook_info(sample_dir / n) for n in ("standard.xlsx", "shifted.xlsx", "table.xlsx")]


def make_config(tmp_path, **extra) -> dict:
    return {
        "TESTING": True,
        "SECRET_KEY": "test-secret",
        "DATABASE": tmp_path / "app.db",
        "UPLOAD_DIR": tmp_path / "uploads",
        "DATA_DIR": tmp_path / "data",
        # env ファイルに本物のAPIキーがあっても、テストが外部のAIに接続しないようにする
        # （AIを使うテストは偽サーバーの URL とキーで上書きする）
        "OPENAI_BASE_URL": "http://127.0.0.1:9/v1",
        "OPENAI_API_KEY": "",
        **extra,
    }


@pytest.fixture
def app(tmp_path):
    return create_app(make_config(tmp_path))


class BufferedClient(FlaskClient):
    """本文を最後まで読んで応答を閉じるテスト用クライアント（本番のサーバと同じ扱い）。

    ダウンロードしたデータを消すのは「本文を送り終えて応答を閉じたとき」なので（core.purge.purge_after_send）、
    応答を閉じないテストクライアントでは消える処理が動かない。buffered=True で毎回閉じる。
    """

    def open(self, *args, **kwargs):
        kwargs.setdefault("buffered", True)
        return super().open(*args, **kwargs)


@pytest.fixture
def client(app):
    app.test_client_class = BufferedClient
    return app.test_client()


# ---- 出来上がったデータを作る（画面と同じ道すじで） ------------------------------------------
# 画面は「帳票取り込み・表の取り込み・帳票登録」の3つだけで、どの段も fetch で進む（2026-09-20 の作り直し）。
# 「確定済みの帳票」「確定済みの一覧表の取り込み」は片付け・ダウンロードのテストで何度も要るので、ここに置く。

EXTRACTION = {
    "pattern": {"id": 1, "name": "設備修理報告書", "version": "v1"},
    "values": {"equipment_id": "EQ-001"},
    "fields": [{"field_name": "equipment_id", "display_name": "設備番号", "data_type": "string", "value": "EQ-001",
                "sheet": "修理報告書", "label_cell": "A4", "value_cell": "B4", "edited": False}],
    "attachments": [], "sheets": ["修理報告書"],
}


def extraction_json(value: str = "EQ-001") -> str:
    import json

    data = json.loads(json.dumps(EXTRACTION))
    data["values"]["equipment_id"] = value
    data["fields"][0]["value"] = value
    return json.dumps(data, ensure_ascii=False)


def add_confirmed_document(app, name: str, *, value: str = "EQ-001", batch_id: str = "",
                           order: int = 0) -> tuple[int, Path]:
    """確定済みの帳票を1件作る（アップロードしたファイルの実体も置く）。"""
    import database as db

    with app.app_context():
        stored = f"documents/{name}"
        path = Path(app.config["UPLOAD_DIR"]) / stored
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"dummy-excel")
        doc_id = db.create_document(name, "0" * 64, stored, batch_id=batch_id, batch_order=order)
        db.update_document(doc_id, data_json=extraction_json(value), confirmed_json=extraction_json(value),
                           title=f"{value} 確定")
    return doc_id, path


def confirmed_import(app, client, file_name: str = "トラブル一覧.csv", template_name: str = "トラブル対応一覧") -> int:
    """CSV を置いて確定まで進めた取り込み（zip をダウンロードできる状態）。

    段ごとの fetch は tests/tables_helpers.py（表の取り込みのテスト用ヘルパー）と同じものを使う。
    """

    return confirmed(app, client, file_name, template_name)


# ====================================================================================================
# 元 tests/fake_servers.py
# テスト用: OpenAI互換API の最小限の偽物。
#
# 既存の使い方（chat_replies に応答を積む）はそのまま。AI整形のテスト用に次を足している。
# - responder: リクエスト内容で応答を変える関数 (body, server) -> Reply | str | None（None なら chat_replies を使う）
# - Reply: ステータス・ヘッダー・遅延・finish_reason を指定した応答
# - モデル名が "noschema" で始まると json_schema を 400 で拒否、"plain" で始まると json_object も拒否
# - keep_scenario: 対応内容の文面（目印の語）で keep の応答を作り分ける responder
# ====================================================================================================

OPENAI_KEY = "sk-test-123456"


@dataclass
class Reply:
    content: str | None = None
    status: int = 200
    body: dict | None = None             # status != 200 のときのエラー本文
    headers: dict = field(default_factory=dict)
    delay: float = 0.0
    finish_reason: str = "stop"
    drop: bool = False                   # 何も返さずに接続を切る（要求を送った後の切断）
    html: str | None = None              # 200 で HTML を返す（プロキシのブロック画面など）


class FakeServer:
    def __init__(self):
        self.requests: list[dict] = []
        self.chat_replies: list[str] = []
        self.responder = None
        self.counters: dict[str, int] = {}
        self.lock = threading.Lock()
        server = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def _send(self, status: int, body: dict, headers: dict | None = None):
                data = json.dumps(body).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                for k, v in (headers or {}).items():
                    self.send_header(k, str(v))
                self.end_headers()
                try:
                    self.wfile.write(data)
                except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                    pass

            def _body(self) -> dict:
                length = int(self.headers.get("Content-Length") or 0)
                return json.loads(self.rfile.read(length) or b"{}")

            def do_GET(self):
                if self.path == "/v1/models":
                    if self.headers.get("Authorization") != f"Bearer {OPENAI_KEY}":
                        return self._send(401, {"error": {"message": "Incorrect API key"}})
                    return self._send(200, {"object": "list", "data": [
                        {"id": m, "object": "model", "created": 0, "owned_by": "test"}
                        for m in ("gpt-test", "picky-model", "catalog-only")]})
                self._send(404, {"detail": "not found"})

            def do_POST(self):
                body = self._body()
                with server.lock:
                    server.requests.append({"path": self.path, "headers": dict(self.headers), "body": body})
                if self.path == "/v1/chat/completions":
                    if self.headers.get("Authorization") != f"Bearer {OPENAI_KEY}":
                        return self._send(401, {"error": {"message": "Incorrect API key"}})
                    model = body.get("model", "")
                    if model.startswith("picky") and "temperature" in body:
                        return self._send(400, {"error": {
                            "message": "Unsupported value: 'temperature' does not support 0 with this model. "
                                       "Only the default (1) value is supported.",
                            "type": "invalid_request_error", "param": "temperature", "code": "unsupported_value"}})
                    rf_type = (body.get("response_format") or {}).get("type")
                    if (rf_type == "json_schema" and model.startswith(("noschema", "plain"))) or \
                            (rf_type == "json_object" and model.startswith("plain")):
                        return self._send(400, {"error": {
                            "message": f"response_format type '{rf_type}' is not supported by this model.",
                            "type": "invalid_request_error", "param": "response_format"}})
                    reply = server.responder(body, server) if server.responder else None
                    if reply is None:
                        with server.lock:
                            reply = server.chat_replies.pop(0) if server.chat_replies else "{}"
                    if isinstance(reply, str):
                        reply = Reply(content=reply)
                    if reply.delay:
                        time.sleep(reply.delay)
                    if reply.drop:
                        self.close_connection = True
                        return
                    if reply.html is not None:
                        data = reply.html.encode()
                        self.send_response(200)
                        self.send_header("Content-Type", "text/html; charset=utf-8")
                        self.send_header("Content-Length", str(len(data)))
                        self.end_headers()
                        self.wfile.write(data)
                        return
                    if reply.status != 200:
                        return self._send(reply.status, reply.body or {"error": {"message": "error"}}, reply.headers)
                    content = reply.content or ""
                    prompt_len = sum(len(str(m.get("content") or "")) for m in body.get("messages") or [])
                    return self._send(200, {
                        "id": "chatcmpl-test", "object": "chat.completion", "created": 0, "model": model,
                        "choices": [{"index": 0, "finish_reason": reply.finish_reason,
                                     "message": {"role": "assistant", "content": content}}],
                        "usage": {"prompt_tokens": prompt_len, "completion_tokens": len(content),
                                  "total_tokens": prompt_len + len(content)},
                    }, {"x-ratelimit-limit-tokens": "200000", **reply.headers})
                self._send(404, {"detail": "not found"})

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.url = f"http://127.0.0.1:{self.httpd.server_address[1]}"
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)

    def chat_requests(self) -> list[dict]:
        with self.lock:
            return [r for r in self.requests if r["path"] == "/v1/chat/completions"]

    def count(self, name: str) -> int:
        """名前ごとの呼び出し回数を1つ進めて、進める前の値を返す。"""
        with self.lock:
            n = self.counters.get(name, 0)
            self.counters[name] = n + 1
            return n

    def __enter__(self):
        self.thread.start()
        return self

    def __exit__(self, *exc):
        self.httpd.shutdown()
        self.httpd.server_close()


# ---- AI整形（keep）用の応答づくり ----------------------------------------------------------

_SEG_LINE = re.compile(r"^\[(s\d+)｜[^\]]*\] ?(.*)$")


def user_messages(body: dict) -> list[str]:
    return [str(m.get("content") or "") for m in body.get("messages") or [] if m.get("role") == "user"]


def is_repair(body: dict) -> bool:
    users = user_messages(body)
    return bool(users) and users[-1].startswith("前回のJSONに次の問題")


def segments_of(body: dict) -> list[tuple[str, str]]:
    """<segments> の [sN｜…] 行 → [(id, 本文)]（再依頼でも元の依頼から読む）。"""
    for text in user_messages(body):
        if "<segments>" in text:
            block = text.split("<segments>", 1)[1].split("</segments>", 1)[0]
            out = []
            for line in block.strip().splitlines():
                m = _SEG_LINE.match(line)
                if m:
                    out.append((m.group(1), m.group(2)))
                elif out:
                    out[-1] = (out[-1][0], out[-1][1] + "\n" + line)
            return out
    return []


def keep_base(segs: list[tuple[str, str]]) -> dict:
    """1セグメント＝1エントリの正しい keep 応答（要点は空）。"""
    return {
        "entries": [{"id": f"e{i}", "segs": [sid], "t": ["メモ"]} for i, (sid, _) in enumerate(segs, start=1)],
        "ignored": [],
        "incident": {"root_cause": None, "temporary_actions": [], "permanent_actions": [], "parts": [],
                     "recurrence": None, "final_state": None},
    }


def _entry_with(segs, word: str) -> str:
    for i, (_, text) in enumerate(segs, start=1):
        if word in text:
            return f"e{i}"
    return "e1"


def keep_scenario(body: dict, server: FakeServer):
    """文面の目印で応答を変える。

    - 方式判定（{"ok": true}）→ そのまま返す
    - custom 段（<inputs>）: 「緩み」→ 締結緩み、「謎」→ 原文にない引用、それ以外 → null
    - 「清掃」: 暫定処置「センサー清掃」（正しい）
    - 「型番違い」: 部品の型番を原文にない表記にする（再依頼でも直さない）＋正しい恒久処置
    - 「日付捏造」: 暫定処置に日付を書く（再依頼でも直さない）＋正しい暫定処置
    - 「手配のみ」: 手配しか書かれていないのに恒久処置「ケーブル交換済」＋正しい部品
    - 「ID抜け」: 最後のセグメントをどのエントリにも入れない（再依頼でも直さない）
    - 「壊れJSON」: 最初は壊れたJSON、再依頼で正しい応答
    - 「混雑」: 最初は 429（Retry-After: 1）、次は正しい応答
    - 「遅い応答」: 0.8秒待ってから返す
    """
    users = user_messages(body)
    if users and '{"ok": true}' in users[-1]:
        return Reply(content='{"ok": true}')
    if users and "<inputs>" in users[-1]:
        text = users[-1]
        if "緩み" in text:
            return Reply(content=json.dumps({"v": "締結緩み", "q": "緩み"}, ensure_ascii=False))
        if "謎" in text:
            return Reply(content=json.dumps({"v": "摩耗", "q": "すり減り"}, ensure_ascii=False))
        return Reply(content='{"v": null, "q": null}')
    segs = segments_of(body)
    if not segs:
        return None
    joined = "\n".join(t for _, t in segs)
    repair = is_repair(body)
    out = keep_base(segs)
    inc = out["incident"]
    delay = 0.0
    if "清掃" in joined:
        inc["temporary_actions"].append({"v": "センサー清掃", "src": [_entry_with(segs, "清掃")]})
    if "型番違い" in joined:
        inc["permanent_actions"].append({"v": "コネクタの増し締め", "src": [_entry_with(segs, "増し締め")]})
        inc["parts"].append({"name": "エンコーダケーブル", "model": "RB-ENC-5M", "qty_q": "×1",
                             "src": [_entry_with(segs, "RB-ENC")]})
    if "日付捏造" in joined:
        inc["temporary_actions"].append({"v": "原点復帰", "src": [_entry_with(segs, "原点復帰")]})
        inc["temporary_actions"].append({"v": "4/8に再起動", "src": [_entry_with(segs, "再起動")]})
    if "手配のみ" in joined:
        e = _entry_with(segs, "手配")
        inc["permanent_actions"].append({"v": "ケーブル交換済", "src": [e]})
        inc["parts"].append({"name": "ケーブル", "model": "RB-ENC-05M", "qty_q": "×1", "src": [e]})
    if "ID抜け" in joined:
        out["entries"] = out["entries"][:-1]
    if "壊れJSON" in joined and not repair:
        return Reply(content='{"entries": [{"id": "e1", "segs": ["s1"')
    if "混雑" in joined and server.count("混雑") == 0:
        return Reply(status=429, body={"error": {"message": "Rate limit reached", "type": "rate_limit_exceeded"}},
                     headers={"retry-after": "1"})
    if "遅い応答" in joined:
        delay = 0.8
    return Reply(content=json.dumps(out, ensure_ascii=False), delay=delay)


# ====================================================================================================
# 元 tests/tables_helpers.py
# 表の取り込み（1画面）のテスト用ヘルパー。
#
# 画面は /tables の1枚だけで、段（panel）の中身と保存はすべて fetch でやりとりする
# （views/tables.py・static/tables.js）。テストも同じ JSON のやりとりで進める。
# 取り込み設定は保存しないので、列の対応づけは取り込みごとに「使う・役割」だけを送る。
# ====================================================================================================

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

def csv_source(client, import_id: int, **fields):
    """いつもの読み取り方の保存（CSV は cp932・カンマ）。"""
    payload = {"encoding": "cp932", "delimiter": ","}
    payload.update(fields)
    res = save_source(client, import_id, **payload)
    assert res.status_code == 200, res.get_json()
    return res.get_json()


def save_layout(client, import_id: int, header_rows="1", data_end_row=""):
    return client.post(f"/tables/imports/{import_id}/layout",
                       json={"header_rows": header_rows, "data_end_row": data_end_row})


# 画面で選べる役割（views.tables.SCREEN_ROLES）。ここに無い役割は「その他」にして候補の役割を活かす
SCREEN_ROLES = {"key", "date", "entity", "log"}


def screen_role(role: str, ai_role: str | None = None) -> str:
    if role == "log":
        return "log" if ai_role == "log" else "attribute"
    return role if role in SCREEN_ROLES else "attribute"


def columns_payload(name: str, columns=COLUMNS, ai_role: str | None = None, **extra) -> dict:
    """列の対応づけの保存に送る JSON。ai_role="log" を渡すとその列を役割「経過の記録」にする。"""
    payload = {
        "name": name,
        "columns": [{"index": i, "use": True, "role": screen_role(r, ai_role)}
                    for i, (_k, _h, _t, r) in enumerate(columns)],
    }
    payload.update(extra)
    return payload


def save_columns(client, import_id: int, payload: dict):
    return client.post(f"/tables/imports/{import_id}/columns", json=payload)


# ---- ジョブの待ち合わせ ------------------------------------------------------------------

def wait_import_job(app, import_id: int) -> dict:
    """読み込み・確定のジョブ（取り込みの job_id）が終わるまで待つ。"""
    with app.app_context():
        job = core.wait_job(tables.get_import(import_id)["job_id"], timeout=60)
        assert job["status"] == "done", job
        return tables.get_import(import_id)


def preview_panel(app, client, import_id: int, **query) -> dict:
    """「内容の確認」の段。初回は md の下書きを作るジョブが動くので、終わってから取り直す。"""
    data = panel(client, import_id, "preview", **query)
    if data.get("building") and data.get("job"):
        with app.app_context():
            job = core.wait_job(data["job"]["id"], timeout=60)
            assert job["status"] == "done", job
        data = panel(client, import_id, "preview", **query)
    return data


# ---- 通しで「内容の確認」まで進める -------------------------------------------------------

def imported(app, client, name: str, table_name: str, text: str = CSV_TEXT, ai_role: str | None = None,
             **payload_extra) -> int:
    """CSV を置いて読み取り方・範囲・列を保存し、読み込みが終わった取り込みを作る。"""
    import_id = upload_csv(client, name, text)
    csv_source(client, import_id)
    res = save_layout(client, import_id)
    assert res.status_code == 200, res.get_json()
    payload = columns_payload(table_name, ai_role=ai_role, **payload_extra)
    res = save_columns(client, import_id, payload)
    assert res.status_code == 200, res.get_json()
    wait_import_job(app, import_id)
    return import_id


def confirmed(app, client, name: str, table_name: str, **kwargs) -> int:
    """確定まで済ませた取り込み（zip をダウンロードできる状態）。"""
    import_id = imported(app, client, name, table_name, **kwargs)
    preview_panel(app, client, import_id)
    res = client.post(f"/tables/imports/{import_id}/confirm")
    assert res.status_code == 200, res.get_json()
    assert wait_import_job(app, import_id)["status"] == "confirmed"
    return import_id


# ---- 「列の対応づけ」の段を、画面と同じ形で読み書きする -------------------------------------
# static/tables.js の saveColumns が集める値（表の名前と、列ごとの「使う・役割」）を段の HTML から集める。


class _EditorParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.settings, self.cols, self.cur, self.sel = {}, [], None, None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "option" and self.sel:
            target, key = self.sel
            if target[key] is None or "selected" in a:
                target[key] = a.get("value")
            return
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
    return {**p.settings, "columns": p.cols}


def editor_body(client, import_id: int) -> dict:
    """「列の対応づけ」の段が出す表を、画面が送る形（そのまま save_columns に渡せる形）にして返す。"""
    return collect_editor(panel_html(client, import_id, "columns"))


def set_role(body: dict, header: str, role: str) -> dict:
    """段から集めた表の、その見出しの列の役割を変える（画面で選び直したのと同じ）。"""
    for col in body["columns"]:
        if col["header"] == header:
            col["role"] = role
    return body


# 見出しがはっきりした、もう1枚の表（列の対応づけの段をいろいろ試すため）
SPEC_CSV = "管理No,発生日,設備名,現象,原因,対応内容,担当,停止時間(分)\r\n" + "".join(
    f'TR-{i:03d},2026-08-{i:02d},搬送ロボット{i % 2 + 1}号機,アラーム停止{i},摩耗,'
    f'"8/{i} 10:00 田中: 確認した。",田中,{i * 5}\r\n' for i in range(1, 13))


def spec_import(app, client, name: str = "編集.csv", text: str = SPEC_CSV) -> int:
    """SPEC_CSV を置いて読み取り方と範囲まで決めた取り込み（「列の対応づけ」の段を開ける状態）。"""
    import_id = upload_csv(client, name, text)
    res = save_source(client, import_id, encoding="cp932", delimiter=",")
    assert res.status_code == 200, res.get_json()
    res = save_layout(client, import_id)
    assert res.status_code == 200, res.get_json()
    return import_id
