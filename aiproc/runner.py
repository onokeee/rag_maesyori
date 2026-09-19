"""AI整形のジョブ本体（run_ai_job）と試し実行（trial_row）。

流れ（1行×段ごと）:
  ルール前処理（マスク→分割）→ 振り分け（対象外 / ルールのみ / AI）→ messages 組み立て
  → キャッシュ照会 → LLM（1セル1回）→ 即時キャッシュ保存 → 照合 →（1回だけ再依頼）→ ai_items に保存
並列: core.jobs のワーカースレッド内で ThreadPoolExecutor を使う。一時停止中は新しい呼び出しを出さない
（送信中の呼び出しは応答を待って保存する）。429 などの待機は Event.wait なので止めるとすぐ抜ける。
エラー: 401/403・モデルなし・接続拒否はジョブを止める。429/5xx/タイムアウトは待って再試行。
        壊れたJSON・長さ超過（再依頼しても直らない）はその行だけエラーにして続ける。
"""
from __future__ import annotations

import gzip
import json
import re
import threading
import time
from collections import deque
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable

from flask import current_app

from aiproc import cache, custom, items, prompts
from aiproc.common import (DEFAULT_LIMITS, DEFAULT_RUN_IF, DEFAULT_SUMMARY_TOKENS, nfkc, sget)
from aiproc.verify import VerifyReport, verify_log_result
from core.jobs import JobCancelled, JobError
from core.mdtext import estimate_tokens
from logproc import PeopleIndex, SplitOptions, mask_text, parse_log, render_timeline, review_notes
from models import database
from services import llm

LOG_STAGE_ID = "log"
MAX_RETRIES = 5          # 429/5xx/タイムアウトの再試行回数（1呼び出しあたり）
MAX_WAIT = 120.0         # 1回の待機の上限（秒）
DEFAULT_CONCURRENCY = {"local": 1, "cloud": 4}
SCOPES = ("pending", "all", "errors", "flagged", "changed")
SCOPE_LABELS = {"pending": "未処理のみ", "all": "全件", "errors": "エラーだけ", "flagged": "要確認だけ",
                "changed": "変更行のみ"}


class AIJobError(JobError):
    """ジョブを止めるエラー（日本語のメッセージ。そのまま画面に出る）。"""


class SettingsChanged(AIJobError):
    """再開時に AI 接続の設定（指紋）が変わっていた。"""


# ---- 入力行の読み込み（アダプタ） ----------------------------------------------------

@dataclass
class ImportData:
    import_id: int
    template_id: int | None
    template_version_id: int | None
    spec: object                    # TableSpec または dict
    rows: list[dict]                # {"key", "values": {列キー: 値}, "originals", "source"}
    file_name: str = ""


# 差し替えられる読み込み口（None なら既定の読み込み）
ROW_LOADER: Callable[[int], ImportData] | None = None


def _tables_dir() -> Path:
    cfg = current_app.config
    return Path(cfg.get("TABLES_DIR") or Path(cfg["DATA_DIR"]) / "tables")


def _read_jsonl(path: Path) -> list[dict]:
    opener = gzip.open if path.suffix == ".gz" else open
    out = []
    with opener(path, "rt", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def _normalize_row(raw: dict, n: int) -> dict:
    if "values" in raw and isinstance(raw["values"], dict):
        row = dict(raw)
    else:
        row = {"values": {k: v for k, v in raw.items() if not str(k).startswith("_")}}
    row["key"] = str(raw.get("key") or raw.get("row_key") or raw.get("_key") or f"row{n}")
    row.setdefault("originals", {})
    row.setdefault("source", {})
    return row


def load_spec(template_version_id: int, conn=None):
    """版の spec_json を TableSpec に（tables.spec が無い・読めないときは dict のまま）。"""
    def run(c):
        row = c.execute("SELECT spec_json FROM table_template_versions WHERE id = ?", (template_version_id,)).fetchone()
        return json.loads(row["spec_json"]) if row else None
    d = _db(conn, run)
    if d is None:
        return None
    try:
        from tables.spec import spec_from_dict
        return spec_from_dict(d)
    except Exception:
        return d


def load_rows_for_ai(import_id: int) -> ImportData:
    """AI整形の入力行を読む。

    既定: table_imports の rows_path（無ければ TABLES_DIR/imports/<id>/rows.jsonl.gz）の JSON Lines。
    テストなどで別の読み込み口を使うときは ROW_LOADER に差し替える。
    """
    if ROW_LOADER is not None:
        return ROW_LOADER(import_id)

    def run(c):
        return c.execute("SELECT * FROM table_imports WHERE id = ?", (import_id,)).fetchone()
    imp = _db(None, run)
    if imp is None:
        raise AIJobError("取り込みが見つかりません。")
    spec = load_spec(imp["template_version_id"]) if imp["template_version_id"] else None
    if spec is None:
        raise AIJobError("取り込み設定が決まっていないため、AI整形を実行できません。")
    path = Path(imp["rows_path"]) if imp["rows_path"] else _tables_dir() / "imports" / str(import_id) / "rows.jsonl.gz"
    if not path.is_absolute():
        path = _tables_dir() / path
    if not path.exists():
        raise AIJobError("読み取った行のデータが見つかりません。表の読み取りからやり直してください。")
    raw_rows = _read_jsonl(path)
    rows = [_normalize_row(r, n) for n, r in enumerate(raw_rows, start=1)]
    return ImportData(import_id, imp["template_id"], imp["template_version_id"], spec, rows, imp["file_name"])


def _db(conn, fn):
    if conn is not None:
        return fn(conn)
    own = database.connect()
    try:
        return fn(own)
    finally:
        own.close()


# ---- ルール前処理と振り分け ----------------------------------------------------------

@dataclass
class StageWork:
    stage_id: str
    kind: str                       # log / custom
    row_key: str
    route: str                      # ai / rule_only / skipped
    reason: str = ""
    messages: list[dict] = field(default_factory=list)
    schema: dict | None = None
    max_tokens: int | None = None
    want_summary: bool = False
    parse: object = None            # LogParse（log 段）
    stage: object = None
    inputs: dict = field(default_factory=dict)       # custom 段の入力 {表示名: 値}
    context: dict = field(default_factory=dict)      # 文脈列 {表示名: 値}
    context_text: str = ""
    sent_text: str = ""             # 実際に送る user の文面（照合用）
    people_names: list[str] = field(default_factory=list)
    entity_label: str = ""
    source_hash: str = ""
    context_hash: str = ""
    segments_hash: str = ""
    key: str = ""                   # キャッシュキー（方式が決まってから付ける）

    def hashes(self) -> dict:
        return {"source_hash": self.source_hash, "context_hash": self.context_hash,
                "segments_hash": self.segments_hash}


def _columns(spec) -> list[tuple[str, str, str]]:
    out = []
    for c in sget(spec, "columns", []) or []:
        out.append((str(sget(c, "key", "")), str(sget(c, "display", "") or sget(c, "key", "")), str(sget(c, "role", ""))))
    return out


def _resolve_column(spec, name: str) -> str:
    for key, display, _ in _columns(spec):
        if name in (key, display):
            return key
    for c in sget(spec, "columns", []) or []:
        if name in (sget(c, "headers", []) or []):
            return str(sget(c, "key", ""))
    return name


def _labels(spec) -> dict[str, str]:
    return {key: display for key, display, _ in _columns(spec)}


_DATE_RE = re.compile(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})")


def _base_date(row: dict, spec) -> date | None:
    period = sget(spec, "period", {}) or {}
    keys = [sget(period, "date_column", None)] + [k for k, _, role in _columns(spec) if role == "date"] + ["occurred_at"]
    for k in keys:
        if not k:
            continue
        m = _DATE_RE.search(str(row["values"].get(k) or ""))
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                continue
    return None


def _entity_label(row: dict, spec) -> str:
    vals = row["values"]
    for role in ("entity_label", "entity"):
        for k, _, r in _columns(spec):
            if r == role and str(vals.get(k) or "").strip():
                return nfkc(vals[k]).strip()
    return ""


def people_index(data: ImportData, stage) -> PeopleIndex:
    """人物一覧（設定）＋ person 役割の列のユニーク値。AIには送らない。"""
    person_keys = [k for k, _, role in _columns(data.spec) if role == "person"]
    names = []
    seen = set()
    for row in data.rows:
        for k in person_keys:
            v = str(row["values"].get(k) or "").strip()
            if v and v not in seen:
                seen.add(v)
                names.append(v)
    groups = sget(stage, "groups", []) or None
    return PeopleIndex(sget(stage, "people", []) or [], column_names=names, groups=groups)


def _run_if(rule, parse, text: str) -> bool:
    """決まった形の JSON 条件だけを評価する（式の文字列は使わない）。"""
    if not rule:
        return True
    if isinstance(rule, dict):
        if "any" in rule:
            return any(_run_if(r, parse, text) for r in rule["any"] or [])
        if "all" in rule:
            return all(_run_if(r, parse, text) for r in rule["all"] or [])
        ok = True
        if "min_segments" in rule:
            ok = ok and len(parse.segments) >= int(rule["min_segments"])
        if "min_chars" in rule:
            ok = ok and len(text.strip()) >= int(rule["min_chars"])
        if "contains" in rule:
            words = rule["contains"] if isinstance(rule["contains"], list) else [rule["contains"]]
            ok = ok and any(str(w) in text for w in words)
        return ok
    return True


def _prepare_log(row: dict, data: ImportData, stage, people: PeopleIndex) -> StageWork:
    spec = data.spec
    col = _resolve_column(spec, str(sget(stage, "column", "")))
    raw = row["values"].get(col)
    raw_text = "" if raw is None else str(raw)
    # マスクと分割は tables.markdown.parse_log_cell と同じ規則（セグメントIDをそろえる）
    rules = sget(stage, "mask", None)
    rules = ["phone", "email"] if rules is None else list(rules)
    masked = raw_text
    if rules:
        names = people.names() if any(r in ("person", "人名") for r in rules) else ()
        masked, _ = mask_text(raw_text, rules, names=names)
    parse = parse_log(masked, _base_date(row, spec), people, SplitOptions.from_dict(sget(stage, "splitter", {}) or {}))
    labels = _labels(spec)
    context = {}
    for c in sget(stage, "context_columns", []) or []:
        k = _resolve_column(spec, str(c))
        v = row["values"].get(k)
        if v is not None and str(v).strip():
            context[labels.get(k, k)] = nfkc(v).strip()
    work = StageWork(LOG_STAGE_ID, "log", row["key"], "ai", parse=parse, stage=stage, context=context,
                     context_text=prompts.context_text(context), people_names=people.names(),
                     entity_label=_entity_label(row, spec) or next(iter(context.values()), ""),
                     source_hash=items.source_hash(masked), context_hash=items.context_hash(context),
                     segments_hash=items.segments_hash(parse))
    limits = dict(DEFAULT_LIMITS)
    limits.update(sget(stage, "limits", {}) or {})
    if parse.kind == "empty":
        work.route, work.reason = "skipped", "対応内容の記載なし"
        return work
    if parse.kind == "header_cell":
        work.route, work.reason = "rule_only", "見出し型のセル（ルールのみ）"
        return work
    if parse.segments and all("reference" in (s.marks or []) for s in parse.segments):
        work.route, work.reason = "rule_only", "他の記録の参照だけ（要確認）"
        return work
    if len(parse.segments) > int(limits["max_segments"]):
        work.route, work.reason = "rule_only", f"区切りが多すぎる（{len(parse.segments)}件）"
        return work
    if not _run_if(sget(stage, "run_if", None) or DEFAULT_RUN_IF, parse, parse.text):
        work.route, work.reason = "rule_only", "短い記載（ルールのみ）"
        return work
    summary_if = sget(stage, "summary_if", {}) or {}
    over = int(sget(summary_if, "record_tokens_over", DEFAULT_SUMMARY_TOKENS))
    work.want_summary = estimate_tokens("\n".join(render_timeline(parse, work.entity_label))) > over
    work.messages = prompts.build_log_messages(parse, context, stage, want_summary=work.want_summary)
    work.schema = prompts.log_output_schema(stage, work.want_summary)
    work.max_tokens = prompts.log_max_tokens(stage, parse, work.want_summary)
    work.sent_text = work.messages[-1]["content"]
    tokens = estimate_tokens("".join(m["content"] for m in work.messages))
    if tokens > int(limits["max_input_tokens"]):
        work.route, work.reason = "rule_only", f"入力が長すぎる（推定{tokens}トークン）"
        work.messages = []
    return work


def _prepare_custom(row: dict, data: ImportData, stage) -> StageWork:
    labels = _labels(data.spec)
    inputs = {}
    for c in sget(stage, "inputs", []) or []:
        k = _resolve_column(data.spec, str(c))
        inputs[labels.get(k, k)] = row["values"].get(k)
    sid = str(sget(stage, "id", "custom"))
    work = StageWork(sid, "custom", row["key"], "ai", stage=stage, inputs=inputs,
                     source_hash=items.source_hash(prompts.custom_user_content(inputs)),
                     context_hash=items.context_hash({}), segments_hash="")
    run_if = sget(stage, "run_if", None)
    if not any(str(v or "").strip() for v in inputs.values()):
        work.route, work.reason = "skipped", "入力が空"
    elif isinstance(run_if, dict) and run_if.get("not_empty"):
        k = _resolve_column(data.spec, str(run_if["not_empty"]))
        if not str(row["values"].get(k) or "").strip():
            work.route, work.reason = "skipped", "入力が空"
    if work.route == "ai":
        work.messages, work.schema = custom.messages_for(stage, inputs)
        work.sent_text = work.messages[-1]["content"]
        work.max_tokens = 200 + int(sget(stage, "max_chars", 80)) * 2
    return work


def enabled_stages(spec, stage_ids=None) -> list[tuple[str, str, object]]:
    """(stage_id, kind, stage)。stage_ids を渡せば enabled_ai に関係なくその段を使う。"""
    out = []
    log_stage = sget(spec, "log_stage", None)
    if log_stage is not None and (stage_ids is None and sget(log_stage, "enabled_ai", False)
                                  or stage_ids is not None and LOG_STAGE_ID in stage_ids):
        out.append((LOG_STAGE_ID, "log", log_stage))
    for st in sget(spec, "custom_stages", []) or []:
        sid = str(sget(st, "id", ""))
        if stage_ids is None or sid in stage_ids:
            out.append((sid, "custom", st))
    return out


def prepare_works(data: ImportData, stage_ids=None, row_keys=None) -> list[StageWork]:
    stages = enabled_stages(data.spec, stage_ids)
    people = None
    works = []
    wanted = set(row_keys) if row_keys else None
    for row in data.rows:
        if wanted is not None and row["key"] not in wanted:
            continue
        for sid, kind, stage in stages:
            if kind == "log":
                people = people or people_index(data, stage)
                works.append(_prepare_log(row, data, stage, people))
            else:
                works.append(_prepare_custom(row, data, stage))
    return works


def assign_keys(works: list[StageWork], settings: dict, mode: str) -> None:
    for w in works:
        if w.route == "ai":
            w.key = cache.cache_key(w.messages, settings.get("model", ""), settings.get("params") or {},
                                    w.schema if mode == "json_schema" else None, mode)


def selected(work: StageWork, item: dict | None, template_version_id, scope: str, forced: bool) -> bool:
    """範囲の指定で、この行×段を処理するか。人の判断（override）がある行は処理しない。"""
    if item and item.get("override") in items.OVERRIDES:
        return False
    if forced or scope == "all":
        return True
    outdated = items.is_outdated(item, template_version_id, work.source_hash, work.context_hash, work.segments_hash)
    status = (item or {}).get("status")
    if scope == "errors":
        return status == "error"
    if scope == "flagged":
        return status == "flagged"
    if scope == "changed":
        return outdated or status == "outdated"
    return item is None or status in ("pending", "error", "outdated") or outdated


# ---- 1行の呼び出しと照合（ワーカースレッド） ---------------------------------------------

class _Stopped(Exception):
    """一時停止・中止で再試行の待機を抜けた。"""


@dataclass
class Outcome:
    work: StageWork
    status: str = "pending"
    result: dict | None = None
    checks: dict | None = None
    error: str | None = None
    calls: int = 0                  # 実際に送った HTTP 呼び出しの数
    cached: bool = False            # キャッシュだけで済んだ
    attempts: int = 0
    tokens_in: int = 0
    tokens_out: int = 0
    latency_ms: int = 0
    headers: dict = field(default_factory=dict)
    raw_text: str = ""
    fatal: llm.LLMCallError | None = None
    stopped: bool = False
    key: str = ""


def call_with_retry(settings: dict, messages, response_format, max_tokens, stop_event: threading.Event | None = None,
                    max_retries: int = MAX_RETRIES, max_wait: float = MAX_WAIT):
    """429/5xx/タイムアウトは待って再試行（Retry-After に従う）。待機中に止められたら _Stopped。"""
    attempt = 0
    stop_event = stop_event or threading.Event()
    while True:
        try:
            return llm.chat_raw(settings, messages, response_format=response_format, max_tokens=max_tokens)
        except llm.LLMCallError as e:
            if e.kind != "retry" or attempt >= max_retries:
                raise
            attempt += 1
            sec = e.retry_after if e.retry_after is not None else min(2.0 ** attempt, 60.0)
            if stop_event.wait(min(max(sec, 0.0), max_wait)):
                raise _Stopped()


def _evaluate(work: StageWork, text: str, finish_reason: str | None):
    """(status, result, checks, problems, parsed)。problems があれば再依頼の対象。"""
    try:
        parsed = llm.parse_json_text(text)
    except ValueError as e:
        return "error", None, {"issues": [{"level": "fatal", "path": "output", "message": str(e), "code": "json"}]}, \
            ["JSONとして読めませんでした。JSONだけを出力してください。"], None
    if work.kind == "log":
        rep: VerifyReport = verify_log_result(parsed, work.parse, work.sent_text, spec=work.stage,
                                              context_text=work.context_text, people_names=work.people_names,
                                              finish_reason=finish_reason, want_summary=work.want_summary)
        status = rep.status()
        if status == "rule_only":
            status = "error" if finish_reason == "length" else "flagged"
        return status, rep.accepted, rep.to_dict(), rep.repair_problems(), parsed
    if finish_reason == "length":
        return "error", None, {"issues": [{"level": "fatal", "path": "output", "message": "出力が上限で打ち切られました。",
                                           "code": "length"}]}, ["出力が上限で打ち切られました。短くしてください。"], parsed
    res = custom.verify_custom_result(work.stage, parsed, work.inputs)
    problems = [i.message for i in res.issues if i.level in ("error", "fatal")]
    return res.status(), res.to_dict(), {"issues": res.to_dict()["issues"]}, problems, parsed


def _one_call(work: StageWork, messages, key: str, settings: dict, mode: str, stop_event, out: Outcome,
              max_tokens, use_cache: bool = True) -> tuple[str, str | None]:
    """キャッシュを見て、無ければ呼んで即時保存する。(text, finish_reason) を返す。"""
    hit = cache.get(key) if use_cache else None
    if hit is not None:
        out.tokens_in += hit.get("tokens_in") or 0
        out.tokens_out += hit.get("tokens_out") or 0
        return hit.get("raw_text") or "", hit.get("finish_reason")
    rf = llm.response_format_for(mode, work.schema, "log_keep" if work.kind == "log" else "custom")
    res = call_with_retry(settings, messages, rf, max_tokens, stop_event)
    out.calls += 1
    out.tokens_in += res.tokens_in or 0
    out.tokens_out += res.tokens_out or 0
    out.latency_ms += res.latency_ms
    out.headers = res.headers
    parsed = None
    try:
        parsed = llm.parse_json_text(res.text)
    except ValueError:
        pass
    cache.put_result(key, res, model=settings.get("model", ""), structured_mode=mode, parsed=parsed)
    return res.text, res.finish_reason


def _repair_request(work: StageWork, text: str, problems: list, settings: dict, mode: str) -> tuple[list, str]:
    """再依頼のメッセージとキャッシュキー。"""
    r_messages = prompts.build_repair_messages(work.messages, text, problems)
    return r_messages, cache.cache_key(r_messages, settings.get("model", ""), settings.get("params") or {},
                                       work.schema if mode == "json_schema" else None, mode)


def cached_usable(work: StageWork, settings: dict, mode: str, key: str | None = None) -> bool:
    """キャッシュだけで AI を呼ばずに済むか（見積もり用）。execute_work と同じ判断をする：
    キャッシュが無い・再依頼が要るのに再依頼の応答が無い・キャッシュの応答がエラー（壊れたJSON・打ち切り。
    実行時は聞き直す）なら False。"""
    hit = cache.get(key or work.key)
    if hit is None:
        return False
    text, finish = hit.get("raw_text") or "", hit.get("finish_reason")
    status, _, _, problems, _ = _evaluate(work, text, finish)
    if problems:
        r_hit = cache.get(_repair_request(work, text, problems, settings, mode)[1])
        if r_hit is None:
            return False
        r_status = _evaluate(work, r_hit.get("raw_text") or "", r_hit.get("finish_reason"))[0]
        if not (r_status == "error" and status != "error"):
            status = r_status
    return status != "error"


def execute_work(work: StageWork, settings: dict, mode: str, stop_event: threading.Event | None = None,
                 use_cache: bool = True, repair: bool = True) -> Outcome:
    """1行×段を処理する（キャッシュ→呼び出し→照合→1回だけ再依頼）。app_context 内で呼ぶ。"""
    out = Outcome(work, key=work.key)
    try:
        text, finish = _one_call(work, work.messages, work.key, settings, mode, stop_event, out, work.max_tokens,
                                 use_cache)
        out.attempts = 1
        out.raw_text = text
        status, result, checks, problems, _ = _evaluate(work, text, finish)
        if problems and repair:
            # 一時停止・中止を頼まれていたら再依頼を送らない（1回目の応答は保存済みなので、再開時は再依頼から）
            if stop_event is not None and stop_event.is_set():
                raise _Stopped()
            r_messages, r_key = _repair_request(work, text, problems, settings, mode)
            r_tokens = int(work.max_tokens * 1.5) if (finish == "length" and work.max_tokens) else work.max_tokens
            r_text, r_finish = _one_call(work, r_messages, r_key, settings, mode, stop_event, out, r_tokens, use_cache)
            out.attempts = 2
            r_status, r_result, r_checks, r_problems, _ = _evaluate(work, r_text, r_finish)
            # 再依頼で壊れた（JSON が読めない）ときは最初の結果を残す
            if not (r_status == "error" and status != "error"):
                status, result, checks, problems, text, out.key = r_status, r_result, r_checks, r_problems, r_text, r_key
                out.raw_text = r_text
            if checks is not None:
                checks["repaired"] = True
        out.status, out.result, out.checks = status, result, checks
        if status == "error":
            msgs = [i.get("message") for i in (checks or {}).get("issues", []) if i.get("level") == "fatal"]
            out.error = msgs[0] if msgs else "AIの応答を使えませんでした。"
        out.cached = out.calls == 0
        if status == "error" and out.cached and use_cache:
            # キャッシュの応答だけでエラーになった（壊れたJSON・打ち切り）。同じ応答を再生しても直らないので
            # 聞き直す（「エラーだけ再実行」で直せるように）。再依頼で直った応答はこれまでどおり使い回す
            return execute_work(work, settings, mode, stop_event, use_cache=False, repair=repair)
    except _Stopped:
        out.stopped = True
    except llm.LLMCallError as e:
        if e.kind == "fatal":
            out.fatal = e
        out.status, out.error = "error", str(e)
    return out


# ---- ジョブ本体 ------------------------------------------------------------------

def start_ai_job(import_id: int, scope: str = "pending", concurrency: int | None = None, stage_ids=None,
                 row_keys=None, model: str | None = None) -> int:
    """画面から呼ぶ：設定を固定してジョブを登録する（app_context 内）。"""
    from core.jobs import start_job

    settings = llm.job_client_settings(model=model)
    params = {"import_id": import_id, "scope": scope, "concurrency": concurrency, "stage_ids": stage_ids,
              "row_keys": row_keys, "model": settings["model"], "fingerprint": settings["fingerprint"],
              "settings": llm.public_settings(settings)}
    return start_job("ai_format", "table_import", import_id,
                     lambda ctx: run_ai_job(ctx, import_id, scope, concurrency), params)


def check_resume(job: dict, settings: dict | None = None) -> tuple[bool, str]:
    """中断・中止したジョブを同じ設定で続けられるか（指紋の比較）。"""
    settings = settings or llm.job_client_settings(model=(job.get("params") or {}).get("model"))
    expected = (job.get("params") or {}).get("fingerprint")
    if expected and expected != settings["fingerprint"]:
        return False, ("AI接続の設定（接続先・モデル・パラメータ）が前回の実行から変わっています。"
                       "新しい設定で残りを別の実行として始めてください。")
    return True, ""


def run_ai_job(ctx, import_id: int, scope: str | None = None, concurrency: int | None = None, *,
               stage_ids=None, row_keys=None, settings: dict | None = None) -> dict:
    """AI整形のジョブ本体。ctx は core.jobs.JobContext。戻り値は件数のまとめ（jobs の result になる）。"""
    params = getattr(ctx, "params", {}) or {}
    scope = scope or params.get("scope") or "pending"
    if scope not in SCOPES:
        raise AIJobError(f"範囲の指定が正しくありません: {scope}")
    stage_ids = stage_ids if stage_ids is not None else params.get("stage_ids")
    row_keys = row_keys if row_keys is not None else params.get("row_keys")
    settings = settings or llm.job_client_settings(model=params.get("model"))
    expected = params.get("fingerprint")
    if expected and expected != settings["fingerprint"]:
        raise SettingsChanged("AI接続の設定（接続先・モデル・パラメータ）が開始時から変わっています。"
                              "新しい設定で残りを別の実行として始めてください。")
    concurrency = int(concurrency or params.get("concurrency")
                      or DEFAULT_CONCURRENCY["local" if settings.get("local") else "cloud"])
    concurrency = max(1, min(concurrency, 16))

    ctx.progress(phase="準備", done=0, total=0)
    data = load_rows_for_ai(import_id)
    works = prepare_works(data, stage_ids, row_keys)
    tv_id = data.template_version_id
    template_id = data.template_id or 0
    existing = {}
    for sid in {w.stage_id for w in works}:
        existing[sid] = items.items_by_key(template_id, sid, import_id=import_id)   # 取り込みごと（design.md 3.3）

    job_id = getattr(ctx, "job_id", None)
    stats = {"total": 0, "done": 0, "ok": 0, "flagged": 0, "error": 0, "rule_only": 0, "skipped": 0,
             "already": 0, "cache_hits": 0, "calls": 0, "tokens_in": 0, "tokens_out": 0}
    todo: list[StageWork] = []
    conn = database.connect()
    try:
        for w in works:
            item = existing[w.stage_id].get(w.row_key)
            if not selected(w, item, tv_id, scope, bool(row_keys)):
                stats["already"] += 1
                continue
            if w.route == "ai":
                todo.append(w)
                continue
            result = custom.fallback_result(w.stage, w.reason).to_dict() if w.kind == "custom" else None
            items.upsert_item(template_id, w.stage_id, w.row_key, status=w.route, template_version_id=tv_id,
                              **w.hashes(), result=result, checks={"reason": w.reason}, job_id=job_id,
                              import_id=import_id, conn=conn, commit=False)
            stats[w.route] += 1
        conn.commit()
        stats["total"] = len(todo) + stats["rule_only"] + stats["skipped"]
        stats["done"] = stats["rule_only"] + stats["skipped"]
        ctx.progress(phase="AI整形", **stats)
        if not todo:
            return _summary(stats, settings, None)

        stop_event = threading.Event()
        mode = _detect_mode(ctx, settings, stop_event)
        assign_keys(todo, settings, mode)
        # 同じ文面の行は1回だけ呼ぶ（重複排除）
        groups: dict[str, list[StageWork]] = {}
        for w in todo:
            groups.setdefault(w.key, []).append(w)
        queue = deque(groups.keys())
        app = current_app._get_current_object()
        started = time.monotonic()
        fatal: llm.LLMCallError | None = None

        def task(w: StageWork) -> Outcome:
            with app.app_context():
                return execute_work(w, settings, mode, stop_event)

        executor = ThreadPoolExecutor(max_workers=concurrency, thread_name_prefix=f"ai-job-{job_id}")
        inflight = {}
        try:
            while queue or inflight:
                if _import_gone(conn, import_id):
                    # ダウンロード・削除で取り込みが消えた: 新しい呼び出しを出さず、結果も書かない（design.md 3.3）
                    stop_event.set()
                    raise JobCancelled()
                stopping = fatal is not None or ctx.should_stop()
                if stopping:
                    stop_event.set()
                while queue and not stopping and len(inflight) < concurrency:
                    key = queue.popleft()
                    inflight[executor.submit(task, groups[key][0])] = key
                if inflight:
                    done, _ = wait(list(inflight), timeout=0.2, return_when=FIRST_COMPLETED)
                    for fut in done:
                        key = inflight.pop(fut)
                        out = fut.result()
                        if out.stopped:
                            queue.appendleft(key)       # 止めたので後でもう一度
                            continue
                        if out.fatal is not None:
                            fatal = fatal or out.fatal
                            queue.appendleft(key)
                            continue
                        _save_outcome(out, groups[key], template_id, tv_id, job_id, conn, stats,
                                      import_id=import_id)
                    if done:
                        conn.commit()
                        _report(ctx, stats, started)
                    continue
                if fatal is not None:
                    break
                if stopping:
                    ctx.progress(**stats)
                    ctx.check_cancel()           # 一時停止なら再開まで待つ。中止なら JobCancelled
                    stop_event.clear()
        finally:
            stop_event.set()
            executor.shutdown(wait=True, cancel_futures=True)
            conn.commit()
        if fatal is not None:
            ctx.progress(**stats)
            raise AIJobError(f"AI整形を止めました: {fatal}")
        return _summary(stats, settings, mode)
    finally:
        conn.close()


def _import_gone(conn, import_id: int) -> bool:
    return conn.execute("SELECT 1 FROM table_imports WHERE id = ?", (import_id,)).fetchone() is None


def _detect_mode(ctx, settings: dict, stop_event: threading.Event) -> str:
    for attempt in range(MAX_RETRIES + 1):
        try:
            return llm.detect_structured_mode(settings)
        except llm.LLMCallError as e:
            if e.kind == "fatal":
                raise AIJobError(f"AI整形を始められません: {e}") from e
            if e.kind != "retry" or attempt >= MAX_RETRIES:
                raise AIJobError(f"AIの出力方式を判定できませんでした: {e}") from e
            sec = e.retry_after if e.retry_after is not None else min(2.0 ** (attempt + 1), 60.0)
            deadline = time.monotonic() + min(sec, MAX_WAIT)
            while time.monotonic() < deadline:
                if ctx.should_stop():
                    ctx.check_cancel()
                time.sleep(min(0.2, max(0.0, deadline - time.monotonic())))
    raise AIJobError("AIの出力方式を判定できませんでした。")


def _save_outcome(out: Outcome, group: list[StageWork], template_id, tv_id, job_id, conn, stats: dict,
                  import_id: int | None = None) -> None:
    for i, w in enumerate(group):
        items.upsert_item(template_id, w.stage_id, w.row_key, status=out.status, template_version_id=tv_id,
                          **w.hashes(), cache_key=out.key, result=out.result, checks=out.checks,
                          attempts=out.attempts, error=out.error, job_id=job_id, import_id=import_id,
                          conn=conn, commit=False)
        stats[out.status] = stats.get(out.status, 0) + 1
        stats["done"] += 1
        if i > 0 or out.cached:
            stats["cache_hits"] += 1
    stats["calls"] += out.calls
    stats["tokens_in"] += out.tokens_in
    stats["tokens_out"] += out.tokens_out


def _report(ctx, stats: dict, started: float) -> None:
    elapsed = max(time.monotonic() - started, 0.001)
    ai_done = stats["ok"] + stats["flagged"] + stats["error"]
    rate = ai_done / elapsed * 60
    remaining = stats["total"] - stats["done"]
    ctx.progress(**stats, rows_per_min=round(rate, 1),
                 remaining_sec=int(remaining / rate * 60) if rate > 0 else None)


def _summary(stats: dict, settings: dict, mode: str | None) -> dict:
    return {**stats, "model": settings.get("model"), "structured_mode": mode, "fingerprint": settings.get("fingerprint")}


# ---- 試し実行（1行ずつ同期） ------------------------------------------------------------

def trial_row(import_id: int, row_key: str, stage_ids=None, settings: dict | None = None, use_cache: bool = True,
              data: ImportData | None = None) -> dict:
    """1行を同期で処理して結果を返す（画面から1行ずつ POST する）。ai_items にも保存する。

    致命的なエラー（キー拒否・接続不可など）は llm.LLMCallError のまま投げる。
    """
    data = data or load_rows_for_ai(import_id)
    if not any(r["key"] == row_key for r in data.rows):
        raise AIJobError(f"行「{row_key}」が見つかりません。")
    if stage_ids is None:
        stage_ids = [sid for sid, _, _ in enumerate_stages(data.spec)]
    works = prepare_works(data, stage_ids, [row_key])
    settings = settings or llm.job_client_settings()
    template_id = data.template_id or 0
    mode = None
    out_stages = []
    for w in works:
        entry = {"stage_id": w.stage_id, "kind": w.kind, "route": w.route, "reason": w.reason,
                 "status": w.route if w.route != "ai" else "pending", "messages": w.messages,
                 "prompt_text": prompts.messages_text(w.messages) if w.messages else "", "raw_text": "",
                 "result": None, "checks": None, "cached": False, "calls": 0, "tokens_in": 0, "tokens_out": 0,
                 "latency_ms": 0}
        if w.route == "ai":
            if mode is None:
                mode = llm.detect_structured_mode(settings)
            assign_keys([w], settings, mode)
            out = execute_work(w, settings, mode, None, use_cache=use_cache)
            if out.fatal is not None:
                raise out.fatal
            entry.update(status=out.status, result=out.result, checks=out.checks, raw_text=out.raw_text,
                         cached=out.cached, calls=out.calls, tokens_in=out.tokens_in, tokens_out=out.tokens_out,
                         latency_ms=out.latency_ms, error=out.error, cache_key=out.key, headers=out.headers)
            items.upsert_item(template_id, w.stage_id, w.row_key, status=out.status,
                              template_version_id=data.template_version_id, **w.hashes(), cache_key=out.key,
                              result=out.result, checks=out.checks, attempts=out.attempts, error=out.error,
                              import_id=import_id)
        else:
            result = custom.fallback_result(w.stage, w.reason).to_dict() if w.kind == "custom" else None
            entry["result"] = result
            items.upsert_item(template_id, w.stage_id, w.row_key, status=w.route,
                              template_version_id=data.template_version_id, **w.hashes(), result=result,
                              checks={"reason": w.reason}, import_id=import_id)
        if w.kind == "log" and w.parse is not None:
            types = (entry["result"] or {}).get("types") if entry["status"] in ("ok", "flagged") else None
            entry["timeline"] = render_timeline(w.parse, w.entity_label, types=types,
                                                glossary=sget(w.stage, "glossary", {}) or None)
            entry["notes"] = review_notes(w.parse)
            entry["segments"] = [{"id": s.id, "start": s.start, "end": s.end, "body": s.body} for s in w.parse.segments]
        out_stages.append(entry)
    return {"row_key": row_key, "model": settings.get("model"), "structured_mode": mode, "stages": out_stages}


def enumerate_stages(spec) -> list[tuple[str, str, object]]:
    """設定にある全段（enabled_ai に関係なく）。"""
    out = []
    if sget(spec, "log_stage", None) is not None:
        out.append((LOG_STAGE_ID, "log", sget(spec, "log_stage")))
    for st in sget(spec, "custom_stages", []) or []:
        out.append((str(sget(st, "id", "")), "custom", st))
    return out


def trial_stats(trials: list[dict]) -> list[dict]:
    """trial_row の戻り値の一覧 → 見積もり用の実測値（AIを呼んだ段だけ）。"""
    out = []
    for t in trials:
        for st in t.get("stages", []):
            if st.get("route") == "ai" and st.get("calls"):
                out.append({"tokens_in": st.get("tokens_in") or 0, "tokens_out": st.get("tokens_out") or 0,
                            "latency_ms": st.get("latency_ms") or 0, "stage_id": st.get("stage_id"),
                            "headers": st.get("headers") or {}})
    return out
