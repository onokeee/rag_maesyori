"""aiproc 内で共有する小さな道具（設定の読み出し・ハッシュ・正規化）。"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata

# 選択肢の既定（取り込み設定に無ければこれを使う）
DEFAULT_ENTRY_TYPES = ["連絡", "初動", "調査", "原因判明", "部品手配", "待ち", "暫定処置", "恒久処置", "試運転",
                       "経過観察", "再発", "打合せ", "品質処置", "再発防止", "クローズ", "訂正", "引継ぎ", "メモ"]
DEFAULT_CERTAINTY = ["確定", "疑い", "不明"]
DEFAULT_FINAL_STATES = ["完了", "経過観察中", "部品待ち", "メーカー回答待ち", "承認待ち", "暫定対応中", "未着手", "不明"]
DEFAULT_LIMITS = {"max_segments": 40, "max_input_tokens": 6000}
DEFAULT_RUN_IF = {"any": [{"min_segments": 2}, {"min_chars": 60}, {"contains": ["Original Message", "訂正"]}]}
DEFAULT_OUTPUT_TOKENS = {"base": 400, "per_segment": 40, "max": 2000, "summary": 300}
DEFAULT_SUMMARY_TOKENS = 1500


def sget(obj, name: str, default=None):
    """dataclass でも dict でも同じように値を読む（None は既定値に置き換える）。"""
    if obj is None:
        return default
    value = obj.get(name, None) if isinstance(obj, dict) else getattr(obj, name, None)
    return default if value is None else value


def stable_json(value) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def sha256_text(text: str) -> str:
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def sha256_json(value) -> str:
    return sha256_text(stable_json(value))


def norm(text) -> str:
    """照合用：NFKC＋空白をすべて除く。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or "")))


def nfkc(text) -> str:
    return unicodedata.normalize("NFKC", str(text or ""))
