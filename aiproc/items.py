"""ai_items：行×段の最新状態（ジョブの再開・「エラーだけ再実行」・古くなった判定）。

ハッシュの定義:
- source_hash:   マスク後の対象列の写し（custom 段は入力列の値）
- context_hash:  その段が送る文脈列の値
- segments_hash: その行の分割結果（分割ルールを変えても結果が同じ行は古くならない）
版（template_version_id）かハッシュのどれかが変わった結果は「古い（outdated）」。
"""
from __future__ import annotations

import json

from aiproc.common import sha256_json, sha256_text
from logproc import LogParse
from models import database

STATUSES = ("pending", "ok", "flagged", "rule_only", "error", "skipped", "outdated", "excluded")
STATUS_LABELS = {
    "pending": "未処理", "ok": "照合OK", "flagged": "要確認", "rule_only": "ルールのみ", "error": "エラー",
    "skipped": "対象外", "outdated": "古い結果", "excluded": "除外",
}
OVERRIDES = ("rule_only", "excluded")   # 人が決めた扱い（再実行で上書きしない）


def source_hash(text) -> str:
    return sha256_text(str(text or ""))


def context_hash(context) -> str:
    return sha256_json(context or {})


def segments_hash(parse: LogParse | None) -> str:
    """分割結果のハッシュ（ID・本文・日時・印。記入者は送らないので含めない）。"""
    if parse is None:
        return sha256_text("")
    data = [{"id": s.id, "body": s.body, "when": [s.when.date, s.when.time, s.when.date_to, s.when.shift,
                                                   s.when.estimated] if s.when else None,
             "marks": sorted(s.marks or [])} for s in parse.segments]
    return sha256_json({"kind": parse.kind, "segments": data})


def _run(conn, fn):
    if conn is not None:
        return fn(conn)
    own = database.connect()
    try:
        return fn(own)
    finally:
        own.close()


def _decode(row) -> dict | None:
    if row is None:
        return None
    item = dict(row)
    for col, key, default in (("result_json", "result", None), ("checks_json", "checks", None)):
        try:
            item[key] = json.loads(item[col]) if item.get(col) else default
        except ValueError:
            item[key] = default
    return item


def get_item(template_id: int, stage_id: str, row_key: str, conn=None) -> dict | None:
    return _run(conn, lambda c: _decode(c.execute(
        "SELECT * FROM ai_items WHERE template_id = ? AND stage_id = ? AND row_key = ?",
        (template_id, stage_id, row_key)).fetchone()))


def items_by_key(template_id: int, stage_id: str, conn=None) -> dict[str, dict]:
    def run(c):
        rows = c.execute("SELECT * FROM ai_items WHERE template_id = ? AND stage_id = ?", (template_id, stage_id))
        return {r["row_key"]: _decode(r) for r in rows}
    return _run(conn, run)


def upsert_item(template_id: int, stage_id: str, row_key: str, *, status: str, template_version_id: int | None = None,
                source_hash: str | None = None, context_hash: str | None = None, segments_hash: str | None = None,
                cache_key: str | None = None, result: dict | None = None, checks: dict | None = None,
                attempts: int | None = None, error: str | None = None, job_id: int | None = None,
                conn=None, commit: bool = True) -> None:
    """行×段の状態を保存する。override（人の判断）は変えない。"""
    if status not in STATUSES:
        raise ValueError(f"不明な状態です: {status}")

    def run(c):
        c.execute(
            """INSERT INTO ai_items (template_id, stage_id, row_key, template_version_id, source_hash, context_hash,
                   segments_hash, cache_key, status, result_json, checks_json, attempts, error, job_id, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT (template_id, stage_id, row_key) DO UPDATE SET
                   template_version_id = excluded.template_version_id, source_hash = excluded.source_hash,
                   context_hash = excluded.context_hash, segments_hash = excluded.segments_hash,
                   cache_key = excluded.cache_key, status = excluded.status, result_json = excluded.result_json,
                   checks_json = excluded.checks_json,
                   attempts = CASE WHEN ? IS NULL THEN ai_items.attempts ELSE excluded.attempts END,
                   error = excluded.error, job_id = excluded.job_id, updated_at = excluded.updated_at""",
            (template_id, stage_id, row_key, template_version_id, source_hash, context_hash, segments_hash, cache_key,
             status, json.dumps(result, ensure_ascii=False) if result is not None else None,
             json.dumps(checks, ensure_ascii=False) if checks is not None else None, attempts or 0, error, job_id,
             database.now(), attempts),
        )
        if commit:
            c.commit()
    _run(conn, run)


def is_outdated(item: dict | None, template_version_id, source_hash_: str, context_hash_: str,
                segments_hash_: str) -> bool:
    """保存済みの結果が、今の版・送る文面と合わなければ True（結果が無い行は False）。"""
    if not item or item.get("status") in ("pending", None):
        return False
    return (item.get("template_version_id") != template_version_id or item.get("source_hash") != source_hash_
            or item.get("context_hash") != context_hash_ or item.get("segments_hash") != segments_hash_)


def mark_outdated(template_id: int, stage_id: str, current: dict[str, dict], template_version_id,
                  conn=None) -> int:
    """current = {row_key: {source_hash, context_hash, segments_hash}} と比べて古い結果を outdated にする。件数を返す。"""
    def run(c):
        n = 0
        for key, item in items_by_key(template_id, stage_id, conn=c).items():
            h = current.get(key)
            if h is None or item["status"] in ("outdated", "pending"):
                continue
            if is_outdated(item, template_version_id, h.get("source_hash"), h.get("context_hash"),
                           h.get("segments_hash")):
                c.execute("UPDATE ai_items SET status = 'outdated', updated_at = ? WHERE id = ?",
                          (database.now(), item["id"]))
                n += 1
        c.commit()
        return n
    return _run(conn, run)


def counts(template_id: int, stage_id: str | None = None, conn=None) -> dict[str, int]:
    sql, args = "SELECT status, COUNT(*) AS n FROM ai_items WHERE template_id = ?", [template_id]
    if stage_id:
        sql += " AND stage_id = ?"
        args.append(stage_id)
    return _run(conn, lambda c: {r["status"]: r["n"] for r in c.execute(sql + " GROUP BY status", args)})


def results_for_render(template_id: int, stage_id: str, conn=None) -> dict[str, dict]:
    """Markdown 描画用：照合に通った結果（ok / flagged）だけを {row_key: accepted} で返す。

    override が rule_only / excluded の行、outdated・error の行は含めない（ルール出力に戻す）。
    """
    out = {}
    for key, item in items_by_key(template_id, stage_id, conn=conn).items():
        if item.get("override") in OVERRIDES or item["status"] not in ("ok", "flagged"):
            continue
        if item.get("result") is not None:
            out[key] = item["result"]
    return out
