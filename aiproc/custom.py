"""custom 段：列に登録した指示文で、短い文（text）か選択肢1つ（choice）を作る。

照合に落ちたら fallback の値にする（行は要確認）。同じ入力は同じ messages になるので、キャッシュで1回の呼び出しに済む。
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field

from aiproc.common import nfkc, norm, sget
from aiproc.prompts import build_custom_messages, custom_output_schema
from aiproc.verify import VerifyIssue, _DATE_RE, _check_flip_words, _identifiers, _numbers, _strip_identifiers


@dataclass
class CustomResult:
    value: str                     # 採用した値（落ちたら fallback）
    source: str                    # ai / fallback
    quote: str | None = None
    issues: list[VerifyIssue] = field(default_factory=list)

    def status(self) -> str:
        if self.source == "ai" and not self.issues:
            return "ok"
        return "flagged"

    def to_dict(self) -> dict:
        return {"value": self.value, "source": self.source, "q": self.quote, "issues": [asdict(i) for i in self.issues]}


def messages_for(stage, inputs: dict) -> tuple[list[dict], dict]:
    return build_custom_messages(stage, inputs), custom_output_schema(stage)


def verify_custom_result(stage, result, inputs: dict) -> CustomResult:
    """出力を照合する。text: 字数・引用・識別子・数・日付・完了否定語。choice: 選択肢・引用。"""
    fallback = str(sget(stage, "fallback", "不明"))
    source_text = "\n".join(str(v or "") for v in inputs.values())
    issues: list[VerifyIssue] = []
    if not isinstance(result, dict):
        return CustomResult(fallback, "fallback", None,
                            [VerifyIssue("fatal", "v", "JSONオブジェクトになっていません。", "structure")])
    v, q = result.get("v"), result.get("q")
    if v is None or str(v).strip() == "":
        return CustomResult(fallback, "fallback", None, [])
    v = nfkc(v).strip()
    quote_required = bool(sget(stage, "quote_required", True))
    if q is not None and str(q).strip():
        if norm(q) not in norm(source_text):
            issues.append(VerifyIssue("error", "q", f"引用「{q}」は入力にありません。", "quote"))
    elif quote_required:
        issues.append(VerifyIssue("error", "q", "根拠の引用（q）がありません。", "quote"))

    if sget(stage, "output_type", "text") == "choice":
        choices = list(sget(stage, "choices", []) or [])
        if v not in choices and v != fallback:
            issues.append(VerifyIssue("error", "v", f"「{v}」は選択肢にありません。", "choice"))
    else:
        limit = int(sget(stage, "max_chars", 80))
        if len(v) > limit:
            issues.append(VerifyIssue("error", "v", f"{len(v)}字あります（{limit}字以内）。", "length"))
        src_ids = {norm(t) for t in _identifiers(source_text)}
        for tok in _identifiers(v):
            if norm(tok) not in src_ids:
                issues.append(VerifyIssue("error", "v", f"「{tok}」は入力にありません。", "identifier"))
        missing = sorted(_numbers(v) - _numbers(source_text))
        if missing:
            issues.append(VerifyIssue("error", "v", f"数「{missing[0]}」は入力にありません。", "quantity"))
        m = _DATE_RE.search(_strip_identifiers(v))
        if m and norm(m.group(0)) not in norm(source_text):
            issues.append(VerifyIssue("error", "v", f"日付・時刻「{m.group(0)}」は入力にありません。", "date"))
        _check_flip_words("custom", "v", v, source_text, issues)

    if any(i.level in ("error", "fatal") for i in issues):
        return CustomResult(fallback, "fallback", q, issues)
    return CustomResult(v, "ai", q, issues)


def fallback_result(stage, reason: str = "") -> CustomResult:
    issues = [VerifyIssue("warning", "v", reason, "skipped")] if reason else []
    return CustomResult(str(sget(stage, "fallback", "不明")), "fallback", None, issues)
