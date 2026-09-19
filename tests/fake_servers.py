"""テスト用: OpenAI互換API の最小限の偽物。

既存の使い方（chat_replies に応答を積む）はそのまま。AI整形のテスト用に次を足している。
- responder: リクエスト内容で応答を変える関数 (body, server) -> Reply | str | None（None なら chat_replies を使う）
- Reply: ステータス・ヘッダー・遅延・finish_reason を指定した応答
- モデル名が "noschema" で始まると json_schema を 400 で拒否、"plain" で始まると json_object も拒否
- keep_scenario: 対応内容の文面（目印の語）で keep の応答を作り分ける responder
"""
from __future__ import annotations

import json
import re
import threading
import time
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

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
