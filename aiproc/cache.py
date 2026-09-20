"""llm_calls：AIの生の応答のキャッシュ。

- キーは sha256(実際に送る messages 全文＋モデル＋パラメータ＋スキーマ＋方式＋接続先URL)。
  接続先を別のサーバーに変えたら、同じモデル名でも前の応答は使わない。
- 受け取ったらその場で1件ずつコミットする（落ちても払い済みの呼び出しを失わない）。
- 照合と描画は読み出すたびにやり直す（閾値や md の形を変えても再課金しない）。
- conn を渡さなければ database.connect() で開いて閉じる（app_context が必要）。
"""
from __future__ import annotations

import json

from aiproc.common import sha256_json
from models import database


def cache_key(messages: list[dict], model: str, params: dict | None = None, schema: dict | None = None,
              structured_mode: str | None = None, endpoint: str | None = None) -> str:
    payload = {
        "messages": [{"role": m.get("role"), "content": m.get("content")} for m in messages],
        "model": model or "",
        "params": params or {},
        "schema": schema,
        "structured_mode": structured_mode or "",
    }
    ep = str(endpoint or "").strip().rstrip("/")
    if ep:   # 接続先の指定が無いときは従来と同じキー（既存のキャッシュを無駄にしない）
        payload["endpoint"] = ep
    return sha256_json(payload)


def _run(conn, fn):
    if conn is not None:
        return fn(conn)
    own = database.connect()
    try:
        return fn(own)
    finally:
        own.close()


def get(key: str, conn=None) -> dict | None:
    """保存済みの応答（raw_text, parsed, finish_reason, tokens_in/out, latency_ms ...）。無ければ None。"""
    def run(c):
        row = c.execute("SELECT * FROM llm_calls WHERE cache_key = ?", (key,)).fetchone()
        if row is None:
            return None
        item = dict(row)
        item["params"] = json.loads(item.get("params_json") or "{}")
        try:
            item["parsed"] = json.loads(item["parsed_json"]) if item.get("parsed_json") else None
        except ValueError:
            item["parsed"] = None
        return item
    return _run(conn, run)


def put(key: str, raw_text: str, *, model: str, params: dict | None = None, structured_mode: str = "",
        parsed: dict | None = None, finish_reason: str | None = None, tokens_in: int | None = None,
        tokens_out: int | None = None, latency_ms: int | None = None, import_id: int | None = None,
        conn=None) -> None:
    """応答を保存してすぐコミットする（同じキーは上書き）。

    import_id は「この応答を払った取り込み」。ai_items から参照されない応答（再依頼で直した行の1回目・
    一時停止の直前に受け取った分）も、この列があればその取り込みを消すときに一緒に消せる（design.md 3.3）。
    """
    def run(c):
        c.execute(
            """INSERT OR REPLACE INTO llm_calls (cache_key, raw_text, parsed_json, model, params_json, structured_mode,
                   finish_reason, tokens_in, tokens_out, latency_ms, created_at, import_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (key, raw_text, json.dumps(parsed, ensure_ascii=False) if parsed is not None else None, model,
             json.dumps(params or {}, ensure_ascii=False, default=str), structured_mode, finish_reason,
             tokens_in, tokens_out, latency_ms, database.now(), int(import_id) if import_id else None),
        )
        c.commit()
    _run(conn, run)


def put_result(key: str, result, *, model: str, structured_mode: str, parsed: dict | None = None,
               import_id: int | None = None, conn=None) -> None:
    """services.llm.ChatResult をそのまま保存する。"""
    put(key, result.text, model=model, params=result.params, structured_mode=structured_mode, parsed=parsed,
        finish_reason=result.finish_reason, tokens_in=result.tokens_in, tokens_out=result.tokens_out,
        latency_ms=result.latency_ms, import_id=import_id, conn=conn)


def exists(key: str, conn=None) -> bool:
    return _run(conn, lambda c: c.execute("SELECT 1 FROM llm_calls WHERE cache_key = ?", (key,)).fetchone() is not None)


def count(conn=None) -> int:
    return _run(conn, lambda c: c.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0])


def clear_all(conn=None) -> int:
    """「生の応答を消す」。消すと再実行で再課金になる。"""
    def run(c):
        cur = c.execute("DELETE FROM llm_calls")
        c.commit()
        return cur.rowcount
    return _run(conn, run)
