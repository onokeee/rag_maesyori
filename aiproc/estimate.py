"""AI整形の見積もり（呼び出し数・トークン・所要時間）。AIは呼ばない。

計算: 試し実行の実測値の75パーセンタイル（入力・出力トークン、秒/回）×
      （AI対象の行 − 同じ文面の重複 − キャッシュ済み − 処理済み・人が決めた行）。
所要時間は「同時実行数から」と「TPM 上限から」のうち遅い方。
"""
from __future__ import annotations

import math

from aiproc import cache, items
from aiproc.runner import assign_keys, cached_usable, load_rows_for_ai, prepare_works, selected
from core.mdtext import estimate_tokens
from models import database

DEFAULT_SEC_PER_CALL = {"local": 15.0, "cloud": 8.0}


def percentile(values, p: float = 75) -> float:
    vals = sorted(float(v) for v in values if v is not None)
    if not vals:
        return 0.0
    k = (len(vals) - 1) * p / 100
    lo, hi = math.floor(k), math.ceil(k)
    if lo == hi:
        return vals[int(k)]
    return vals[lo] + (vals[hi] - vals[lo]) * (k - lo)


def duration_text(minutes: float) -> str:
    """見積もり時間の表示。1分に満たなければ「約0分」ではなく「1分未満」。"""
    if minutes < 1:
        return "1分未満"
    return f"約{math.ceil(round(minutes, 1))}分"


def estimate_from_counts(calls: int, trial_stats: list[dict] | None = None, *, concurrency: int = 1,
                         tpm: int | None = None, default_tokens_in: int = 0, default_tokens_out: int = 0,
                         local: bool = False) -> dict:
    """呼び出し数と実測値から見積もる（純粋関数）。"""
    stats = [s for s in (trial_stats or []) if s]
    if stats:
        tin = percentile([s.get("tokens_in") for s in stats])
        tout = percentile([s.get("tokens_out") for s in stats])
        sec = percentile([(s.get("latency_ms") or 0) / 1000 for s in stats])
        basis = "trial"
    else:
        tin, tout = float(default_tokens_in), float(default_tokens_out)
        sec = DEFAULT_SEC_PER_CALL["local" if local else "cloud"]
        basis = "default"
    concurrency = max(1, int(concurrency or 1))
    minutes_conc = calls * sec / concurrency / 60
    minutes_tpm = None
    if tpm and (tin + tout) > 0:
        per_min = tpm / (tin + tout)
        minutes_tpm = calls / per_min if per_min > 0 else None
    minutes = max(minutes_conc, minutes_tpm or 0.0)
    return {
        "calls": int(calls),
        "tokens_in": int(round(calls * tin)),
        "tokens_out": int(round(calls * tout)),
        "minutes": round(minutes, 1),
        "duration_text": duration_text(minutes),
        "minutes_by_concurrency": round(minutes_conc, 1),
        "minutes_by_tpm": round(minutes_tpm, 1) if minutes_tpm is not None else None,
        "per_call": {"tokens_in": round(tin), "tokens_out": round(tout), "seconds": round(sec, 2)},
        "basis": basis,
    }


def tpm_from_headers(trial_stats: list[dict] | None) -> int | None:
    for s in trial_stats or []:
        raw = (s.get("headers") or {}).get("x-ratelimit-limit-tokens")
        if raw:
            try:
                return int(float(raw))
            except ValueError:
                continue
    return None


def estimate(import_id: int, trial_stats: list[dict] | None = None, *, scope: str = "pending",
             concurrency: int | None = None, tpm: int | None = None, settings: dict | None = None,
             structured_mode: str | None = None, stage_ids=None) -> dict:
    """取り込み全体の見積もり。app_context 内で呼ぶ。

    戻り値: {"rows", "calls", "tokens_in", "tokens_out", "minutes", ...内訳}
    settings（llm.job_client_settings()）を渡すとキャッシュ済みの行も差し引く。
    """
    data = load_rows_for_ai(import_id)
    works = prepare_works(data, stage_ids)
    template_id = data.template_id or 0
    # DBは1回だけ開いて使い回す（行ごとに開き直すと1万行規模で数十秒かかる）
    conn = database.connect()
    try:
        existing = {sid: items.items_by_key(template_id, sid, conn, import_id=import_id)
                    for sid in {w.stage_id for w in works}}
        counts = {"ai": 0, "rule_only": 0, "skipped": 0, "already": 0, "duplicates": 0, "cached": 0}
        targets = []
        for w in works:
            if not selected(w, existing[w.stage_id].get(w.row_key), data.template_version_id, scope, False):
                counts["already"] += 1
                continue
            if w.route != "ai":
                counts[w.route] += 1
                continue
            counts["ai"] += 1
            targets.append(w)

        uniq: dict[str, object] = {}
        if settings:
            modes = [structured_mode] if structured_mode else ["json_schema", "json_object", "prompt_only"]
            for w in targets:
                keys = []
                for m in modes:
                    assign_keys([w], settings, m)
                    keys.append((m, w.key))
                dedupe = keys[0][1]
                if dedupe in uniq:
                    counts["duplicates"] += 1
                    continue
                uniq[dedupe] = w
                # 保存済みでも、使えない応答（壊れたJSON・打ち切り）や再依頼の応答が無いものは実行時に AI を呼ぶ
                if any(cache.exists(k, conn) and cached_usable(w, settings, m, k, conn) for m, k in keys):
                    counts["cached"] += 1
        else:
            for w in targets:
                dedupe = "\n".join(m["content"] for m in w.messages)
                if dedupe in uniq:
                    counts["duplicates"] += 1
                else:
                    uniq[dedupe] = w
    finally:
        conn.close()
    calls = len(uniq) - counts["cached"]
    to_call = list(uniq.values())
    avg_in = (sum(estimate_tokens("".join(m["content"] for m in w.messages)) for w in to_call) / len(to_call)
              if to_call else 0)
    avg_out = (sum((w.max_tokens or 0) * 0.5 for w in to_call) / len(to_call)) if to_call else 0
    local = bool((settings or {}).get("local"))
    conc = concurrency or (1 if local else 4)
    result = estimate_from_counts(calls, trial_stats, concurrency=conc, tpm=tpm or tpm_from_headers(trial_stats),
                                  default_tokens_in=int(avg_in), default_tokens_out=int(avg_out), local=local)
    result.update(rows=len(data.rows), ai_rows=counts["ai"], rule_only_rows=counts["rule_only"],
                  skipped_rows=counts["skipped"], already_rows=counts["already"], duplicates=counts["duplicates"],
                  cached=counts["cached"], concurrency=conc)
    return result
