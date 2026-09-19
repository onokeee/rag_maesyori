"""AI出力（keep）の原文照合。すべてコードで行い、AIの自己申告は使わない。

- 照合は実際に送った文面（マスク後・エスケープ後）に対して行う。
- 構造・セグメントの不備は fatal（再依頼 → だめならセル全体をルール出力）。
- それ以外は項目単位：error の項目は出さない（accepted から外す）、warning は出すが要確認。
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

from aiproc.common import DEFAULT_CERTAINTY, DEFAULT_ENTRY_TYPES, DEFAULT_FINAL_STATES, nfkc, norm, sget
from aiproc.prompts import segment_body, stage_choices
from logproc import LogParse
from logproc.extract import identifier_spans

MAX_CHARS = 80
FLIP_WORDS = ["ではなく", "予定", "待ち", "手配", "なし", "不可", "完了", "済", "未", "OK", "NG"]
DROP_WORDS = ["ではなく", "予定", "待ち", "手配", "なし", "不可", "未", "NG"]   # 根拠にあって出力で落ちたら重大
SPECULATION_WORDS = ["疑い", "可能性", "と思われ", "思われる", "らしい", "かもしれ", "おそらく", "恐らく", "推定", "模様"]
ACTION_GROUPS = [
    ["交換", "取替", "取り替え", "取換", "取り換え"], ["増し締め", "増締め", "締め直し", "増締"], ["清掃", "掃除"],
    ["調整"], ["リセット"], ["再起動"], ["修理", "補修"], ["給油", "注油", "給脂"], ["洗浄"], ["校正"], ["溶接"],
    ["再設定"], ["研磨"],
]
# 「1/2に調整」「1/4回転」（分数）と「1日1回」「1日あたり」（頻度）は日付にしない。
# 年まである「4/1/2024」と「月」の付いた日付はいつでも日付
_DATE_RE = re.compile(
    r"\d{1,4}\s*[/／]\s*\d{1,2}\s*[/／]\s*\d{1,4}"
    r"|\d{1,4}\s*[/／]\s*\d{1,2}(?!\d|\s*(?:回転|開度?|程度|以下|以上|まで(?:開|閉|絞|下げ|上げ|減)"
    r"|に(?:調整|設定|変更|絞|下げ|上げ|減|開|閉)|の(?:開度|量|流量|速度|回転|圧力)))"
    r"|\d{4}-\d{1,2}-\d{1,2}|\d{1,2}月\d{1,2}日|\d{1,2}月"
    r"|\d{1,2}日(?!間|\s*\d+\s*回|あたり|当たり|おき)|\d{1,2}\s*[:：]\s*\d{2}|\d{1,2}時(?!間)|令和|平成|昭和"
    r"|翌日|翌週|翌月|翌朝|前日|昨日|本日|今日|今朝|明日|先週|来週|今週|先月|来月|今月|月末|月初|週末|年内|週明け|同日"
    r"|\d+日後|\d+日前|午前|午後"
)
_HONORIFIC_RE = re.compile(r"(?<![お皆各])([一-鿿々ァ-ヶー]{1,6})(さん|様|氏|殿)")
_HONORIFIC_OK = {"客", "業者", "メーカー", "先方", "担当者", "ご担当者", "皆", "各位"}
_NUM_RE = re.compile(r"\d+(?:\.\d+)?")
_KANJI_DIGITS = {"〇": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_KANJI_NUM_RE = re.compile(r"([一二三四五六七八九十]{1,3})(?=本|個|回|度|枚|台|件|箇所|ヶ所|か所|セット|式|袋|缶|巻|人|日|時間|分|秒|週|か月|ヶ月)")
_VAGUE_COUNT = ("数回", "複数", "何回", "数本", "数個", "数枚", "数台", "何度")
# 最後の状態（final_state.v）ごとに、根拠のエントリに書かれているはずの語（どれか1つ）。
# 選択肢にない独自の状態と「不明」は照合しない
FINAL_STATE_WORDS = {
    "完了": ["完了", "クローズ", "CLOSE", "済", "終了", "解決"],
    "経過観察中": ["経過観察", "様子見", "観察"],
    "部品待ち": ["待ち", "待", "手配", "入荷", "納期"],
    "メーカー回答待ち": ["待ち", "待", "問い合わせ", "問合せ", "問合わせ", "回答", "照会"],
    "承認待ち": ["待ち", "待", "承認", "申請"],
    "暫定対応中": ["暫定", "仮", "応急"],
    "未着手": ["未"],
}
_NOT_DONE_RE = re.compile(r"未(?:完了|解決|終了|済)")          # 「未完了」を「完了」の根拠にしない
# 「完了予定」「完了していない」「まだ終わっていない」は、まだ終わっていない（「完了」の根拠にしない）
_NOT_YET_RE = re.compile(
    r"(?:完了|終了|解決|クローズ|済み?)\s*(?:予定|見込み?|次第|待ち|していない|しておらず|せず|できず|できていない|しない|前)"
    r"|まだ[^。]{0,6}?(?:完了|終了|解決|済)")
# 「手配済」「発注済」「連絡済」は段取りが済んだだけで、不具合の対応の完了ではない
_ARRANGED_RE = re.compile(r"(?:手配|発注|連絡|依頼|申請|問い?合わ?せ|注文|見積)\s*済み?")
# 「済」だけが根拠のとき、「完了」と両立しない、まだ終わっていないことを示す語
_PENDING_RE = re.compile(r"待ち|入荷待|納期")
# 「再発なし」の根拠になる言い方（「復旧しない」のような別の否定は根拠にしない）
_RECUR_NO_RE = re.compile(
    r"再発\s*[はがも]?\s*(?:なし|無し|無|せず|しない|していない|しておらず|ない|見られ(?:ない|ず))"
    r"|(?:以降|その後|以後|現在|今のところ)[^。]{0,8}?(?:異常|問題|不具合|症状|発生|再発)\s*(?:なし|無し|無|ない|せず|しない|していない)"
    r"|(?:異常|問題|不具合|症状)\s*(?:なし|無し)")
# 再発の記録（「再発なし」「再発はなし」「再発は見られない」「再発防止」は除く）。
# 「再度」「再び」は、すぐ後に起きたことを示す語があるときだけ（「再度測定し正常」は再発ではない）
_RECUR_YES_RE = re.compile(
    r"再発(?![はがも]?\s*(?:なし|無し|無|せず|しない|していない|しておらず|ない|見られ|防止|対策))"
    r"|(?:再度|再び)[^。、]{0,4}?(?:発生|停止|エラー|異常|同様|同じ)|再燃")
# 原因が分かっていないことを示す語（「確定」の原因の根拠にならない）
UNKNOWN_WORDS = ["不明", "未特定", "調査中", "調査継続", "特定できず", "特定できない", "わからない", "分からない"]
_CONTENT_RUN_RE = re.compile(r"[一-鿿々]{2,}|[ァ-ヶー]{2,}")


def _final_state_words(state: str, glossary: dict) -> list[str]:
    """最後の状態の根拠になる語。用語集でその語に言い換えられる現場の言い方（様子見→経過観察など）も含める。"""
    words = list(FINAL_STATE_WORDS.get(state, []))
    for term, spec in (glossary or {}).items():
        to = spec.get("to", "") if isinstance(spec, dict) else str(spec)
        if words and any(w in to for w in words):
            words.append(str(term))
    return words


@dataclass
class VerifyIssue:
    level: str        # fatal / error / warning
    path: str         # entries / incident.parts[0] / summary[1] など
    message: str      # 再依頼と要確認に使う日本語
    code: str = ""


@dataclass
class VerifyReport:
    ok_items: list[str] = field(default_factory=list)
    failed_items: list[str] = field(default_factory=list)
    issues: list[VerifyIssue] = field(default_factory=list)
    accepted: dict = field(default_factory=dict)   # 照合に通った項目だけ（描画に使う）
    fatal: bool = False

    @property
    def warnings(self) -> list[VerifyIssue]:
        return [i for i in self.issues if i.level == "warning"]

    def status(self) -> str:
        """ai_items の状態: 構造が壊れていれば rule_only 相当、落ちた項目・警告があれば flagged。"""
        if self.fatal:
            return "rule_only"
        return "flagged" if (self.failed_items or self.warnings) else "ok"

    def repair_problems(self) -> list[str]:
        """再依頼に書く問題（fatal と error だけ）。"""
        out = []
        for i in self.issues:
            if i.level in ("fatal", "error") and i.message not in out:
                out.append(i.message)
        return out

    def to_dict(self) -> dict:
        return {"ok_items": self.ok_items, "failed_items": self.failed_items, "fatal": self.fatal,
                "issues": [asdict(i) for i in self.issues], "status": self.status()}


# ---- 文字列の道具 ------------------------------------------------------------------

def _kanji_to_int(s: str) -> int | None:
    if s == "十":
        return 10
    if "十" in s:
        head, _, tail = s.partition("十")
        tens = _KANJI_DIGITS.get(head, 1) if head else 1
        ones = _KANJI_DIGITS.get(tail, 0) if tail else 0
        return tens * 10 + ones
    if len(s) == 1:
        return _KANJI_DIGITS.get(s)
    return None


def _numbers(text: str) -> set[str]:
    """数値の集合（漢数字＋助数詞も数字にする。識別子の中の数字は除く）。"""
    t = _strip_identifiers(nfkc(text))
    t = _KANJI_NUM_RE.sub(lambda m: str(_kanji_to_int(m.group(1)) if _kanji_to_int(m.group(1)) is not None
                                        else m.group(1)), t)
    t = t.replace(",", "")
    out = set()
    for m in _NUM_RE.finditer(t):
        v = m.group(0)
        out.add(str(float(v)) if "." in v else str(int(v)))
    return out


def _strip_identifiers(text: str) -> str:
    spans = identifier_spans(text)
    if not spans:
        return text
    out, pos = [], 0
    for s, e, _ in spans:
        out.append(text[pos:s])
        out.append(" ")
        pos = e
    out.append(text[pos:])
    return "".join(out)


def _identifiers(text: str) -> list[str]:
    return [tok for _, _, tok in identifier_spans(nfkc(text))]


def _anchor_before(text: str, pos: int) -> str:
    """位置の直前のカタカナ・漢字の連なり（「ケーブル手配」の「ケーブル」）。"""
    i = pos
    while i > 0 and re.match(r"[一-鿿々ァ-ヶーA-Za-z0-9]", text[i - 1]) and pos - i < 8:
        i -= 1
    return text[i:pos]


def _anchor_after(text: str, pos: int) -> str:
    i = pos
    while i < len(text) and re.match(r"[一-鿿々ァ-ヶーA-Za-z0-9]", text[i]) and i - pos < 8:
        i += 1
    return text[pos:i]


def _has_word(text: str, word: str) -> bool:
    t = norm(text).upper()
    w = norm(word).upper()
    if w == "なし":
        return "なし" in t or "無し" in t
    if w == "済":
        return "済" in t
    return w in t


# ---- 照合本体 ----------------------------------------------------------------------

class _Ctx:
    def __init__(self, parse: LogParse, sent_text: str, context_text: str, people_names, glossary, choices):
        self.parse = parse
        self.seg_ids = [s.id for s in parse.segments]
        self.seg_index = {sid: i for i, sid in enumerate(self.seg_ids)}
        self.seg_text = {s.id: segment_body(s) for s in parse.segments}
        self.seg_by_id = {s.id: s for s in parse.segments}
        self.sent_text = sent_text or "\n".join(self.seg_text.values())
        self.context_text = context_text or ""
        self.all_ids = {norm(t) for t in _identifiers(self.sent_text + "\n" + self.context_text)}
        self.people = [n for n in (norm(x) for x in people_names or []) if len(n) >= 2]
        self.glossary = glossary or {}
        self.choices = choices
        self.entry_segs: dict[str, list[str]] = {}
        self.entry_types: dict[str, list[str]] = {}


def verify_log_result(result, parse: LogParse, sent_text: str = "", *, spec=None, context_text: str = "",
                      people_names=(), finish_reason: str | None = None, want_summary: bool = False) -> VerifyReport:
    """keep の出力を照合する。spec は LogStageSpec（選択肢・用語集・incident の有無）か dict。"""
    stage = sget(spec, "log_stage", None) or spec
    choices = stage_choices(stage) if stage is not None else {
        "entry_types": DEFAULT_ENTRY_TYPES, "certainty": DEFAULT_CERTAINTY, "final_states": DEFAULT_FINAL_STATES}
    cx = _Ctx(parse, sent_text, context_text, people_names, sget(stage, "glossary", {}), choices)
    rep = VerifyReport()

    if finish_reason == "length":
        _fatal(rep, "output", "出力が上限で打ち切られました。短くまとめてください。", "length")
        return rep
    if not isinstance(result, dict):
        _fatal(rep, "output", "JSONオブジェクトになっていません。", "structure")
        return rep

    entries = _check_entries(result, cx, rep)
    if rep.fatal:
        return rep
    accepted: dict = {"entries": entries, "types": {sid: t for e in entries for sid in e["segs"]
                                                     for t in [e["t"]] if t}}
    rep.ok_items.append("entries")

    if sget(stage, "incident", True):
        inc = result.get("incident")
        if inc is not None and not isinstance(inc, dict):
            _fatal(rep, "incident", "incident がオブジェクトになっていません。", "structure")
            return rep
        accepted["incident"] = _check_incident(inc or {}, cx, rep)

    summary = result.get("summary")
    if summary is not None or want_summary:
        accepted["summary"] = _check_summary(summary, cx, rep, want_summary)
    rep.accepted = accepted
    return rep


def _fatal(rep: VerifyReport, path: str, message: str, code: str) -> None:
    rep.fatal = True
    rep.issues.append(VerifyIssue("fatal", path, message, code))
    if path not in rep.failed_items:
        rep.failed_items.append(path)


def _check_entries(result: dict, cx: _Ctx, rep: VerifyReport) -> list[dict]:
    entries = result.get("entries")
    if not isinstance(entries, list) or not entries:
        _fatal(rep, "entries", "entries がありません。", "structure")
        return []
    used: dict[str, str] = {}
    out = []
    for n, e in enumerate(entries):
        if not isinstance(e, dict) or not isinstance(e.get("id"), str) or not isinstance(e.get("segs"), list):
            _fatal(rep, f"entries[{n}]", f"entries[{n}] の形が正しくありません（id と segs が必要）。", "structure")
            continue
        eid = e["id"].strip()
        if not eid or eid in cx.entry_segs:
            _fatal(rep, f"entries[{n}]", f"エントリID「{eid}」が空か重複しています。", "structure")
            continue
        segs = [str(s).strip() for s in e["segs"]]
        for sid in segs:
            if sid not in cx.seg_index:
                _fatal(rep, f"entries[{eid}]", f"{eid} の segs にある {sid} は存在しないセグメントIDです。", "segments")
            elif sid in used:
                _fatal(rep, f"entries[{eid}]", f"{sid} が {used[sid]} と {eid} の両方に入っています。", "segments")
            else:
                used[sid] = eid
        idx = sorted(cx.seg_index[s] for s in segs if s in cx.seg_index)
        if not segs:
            _fatal(rep, f"entries[{eid}]", f"{eid} の segs が空です。", "segments")
        elif idx and idx != list(range(idx[0], idx[0] + len(idx))):
            _fatal(rep, f"entries[{eid}]", f"{eid} の segs（{', '.join(segs)}）は隣り合っていません。", "segments")
        types_raw = e.get("t") if isinstance(e.get("t"), list) else []
        types = []
        for t in types_raw:
            t = str(t).strip()
            if t in cx.choices["entry_types"]:
                if t not in types:
                    types.append(t)
            else:
                rep.issues.append(VerifyIssue("error", f"entries[{eid}].t",
                                              f"{eid} の種別「{t}」は選択肢にありません。", "choice"))
                if f"entries[{eid}].t" not in rep.failed_items:
                    rep.failed_items.append(f"entries[{eid}].t")
        cx.entry_segs[eid] = segs
        cx.entry_types[eid] = types
        out.append({"id": eid, "segs": segs, "t": types})
    ignored = result.get("ignored") or []
    if not isinstance(ignored, list):
        _fatal(rep, "ignored", "ignored が配列になっていません。", "structure")
        ignored = []
    for sid in (str(x).strip() for x in ignored):
        seg = cx.seg_by_id.get(sid)
        if seg is None:
            _fatal(rep, "ignored", f"ignored の {sid} は存在しないセグメントIDです。", "segments")
            continue
        if sid in used:
            _fatal(rep, "ignored", f"{sid} がエントリと ignored の両方に入っています。", "segments")
            continue
        used[sid] = "ignored"
        body = cx.seg_text[sid]
        if not any(m in (seg.marks or []) for m in ("signature", "greeting")) or re.search(r"\d", nfkc(body)) \
                or _identifiers(body):
            _fatal(rep, "ignored", f"{sid} は署名・挨拶ではないため ignored に入れられません。", "ignored")
    for sid in cx.seg_ids:
        if sid not in used:
            _fatal(rep, "entries", f"{sid} がどのエントリにも入っていません。", "segments")
    return out


def _evidence(cx: _Ctx, src, path: str, rep: VerifyReport, item_issues: list) -> tuple[list[str], str] | None:
    if isinstance(src, str):
        src = [src]
    if not isinstance(src, list) or not src:
        item_issues.append(VerifyIssue("error", path, f"{path} に根拠のエントリID（src）がありません。", "src"))
        return None
    segs = []
    for eid in (str(x).strip() for x in src):
        if eid not in cx.entry_segs:
            item_issues.append(VerifyIssue("error", path, f"{path} の根拠 {eid} は存在しないエントリIDです。", "src"))
            return None
        segs += cx.entry_segs[eid]
    return segs, "\n".join(cx.seg_text[s] for s in segs)


def _check_text_value(cx: _Ctx, path: str, label: str, value: str, evidence: str, issues: list,
                      check_words: bool = True, limit: int = MAX_CHARS, check_dropped: bool = False) -> None:
    """v などの本文の照合（識別子・数量・日付・人名・完了否定語・長さ）。"""
    v = nfkc(value).strip()
    if len(v) > limit:
        issues.append(VerifyIssue("error", path, f"{path}.{label} が長すぎます（{len(v)}字。{limit}字以内）。", "length"))
    for tok in _identifiers(v):
        if norm(tok) not in cx.all_ids:
            near = _similar_identifier(tok, cx)
            hint = f"（原文の表記は「{near}」）" if near else ""
            issues.append(VerifyIssue("error", path, f"{path}.{label} の「{tok}」は原文にありません{hint}。", "identifier"))
    stripped = _strip_identifiers(v)
    m = _DATE_RE.search(stripped)
    if m:
        issues.append(VerifyIssue("error", path, f"{path}.{label} に日付・時刻「{m.group(0)}」を書かないでください。", "date"))
    for name in cx.people:
        if name in norm(v):
            issues.append(VerifyIssue("error", path, f"{path}.{label} に人名を書かないでください。", "person"))
            break
    else:
        for hm in _HONORIFIC_RE.finditer(v):
            if hm.group(1) not in _HONORIFIC_OK and not hm.group(1).endswith(tuple(_HONORIFIC_OK)):
                issues.append(VerifyIssue("error", path, f"{path}.{label} に人名（「{hm.group(0)}」）を書かないでください。",
                                          "person"))
                break
    out_nums = _numbers(v)
    if out_nums:
        ev_nums = _numbers(evidence)
        cell_nums = _numbers(_DATE_RE.sub(" ", "\n".join(cx.seg_text.values()) + "\n" + cx.context_text))
        for num in sorted(out_nums):
            if num in ev_nums:
                continue
            if num in cell_nums and not any(w in evidence for w in _VAGUE_COUNT):
                issues.append(VerifyIssue("warning", path, f"{path}.{label} の数「{num}」は根拠のエントリにはなく、"
                                                           "同じセルの別の記録にあります。", "quantity_other"))
            else:
                issues.append(VerifyIssue("error", path, f"{path}.{label} の数「{num}」は原文にありません。数量・回数は _q に"
                                                         "原文の語句を引用してください。", "quantity"))
    if check_words:
        _check_flip_words(path, label, v, evidence, issues, check_dropped)


def _model_in(model, text: str) -> bool:
    """型番が原文にそのまま書かれているか。識別子は「丸ごと」一致だけを認める
    （原文 RB-ENC-05M に対して RB-ENC-05・ENC のような切れ端は通さない）。"""
    q = norm(str(model))
    if not q or q not in norm(text):
        return False
    ids = {norm(t) for t in _identifiers(str(model))}
    if ids:
        return ids <= {norm(t) for t in _identifiers(text)}
    # 識別子の形でない型番（Φ10 など）は、原文の識別子の中の一部分でなければよい
    return q in norm(_strip_identifiers(nfkc(text)))


def _similar_identifier(tok: str, cx: _Ctx) -> str:
    t = norm(tok).upper().replace("-", "")
    for cand in _identifiers(cx.sent_text):
        c = norm(cand).upper().replace("-", "")
        if c != t and (c.replace("0", "") == t.replace("0", "") or c[:4] == t[:4]):
            return cand
    return ""


def _check_flip_words(path: str, label: str, v: str, evidence: str, issues: list, check_dropped: bool = True) -> None:
    """根拠にない完了・否定語が出た（手配→交換済）か、処置の文で根拠の否定・予定語が落ちた（ケーブル手配→ケーブル交換）か。"""
    ev = nfkc(evidence)
    for w in FLIP_WORDS:
        if _has_word(v, w) and not _has_word(ev, w):
            issues.append(VerifyIssue("error", path, f"{path}.{label} の「{w}」は根拠のエントリに書かれていません"
                                                     "（完了・否定の語を変えないでください）。", "flip_added"))
            return
    if not check_dropped:
        return
    vn = norm(v)
    for w in DROP_WORDS:
        for m in re.finditer(re.escape(w), ev):
            anchors = [a for a in (_anchor_before(ev, m.start()), _anchor_after(ev, m.end())) if len(a) >= 2]
            if any(norm(a) in vn for a in anchors) and not _has_word(v, w):
                issues.append(VerifyIssue("error", path, f"{path}.{label} では根拠の「{w}」が落ちています"
                                                         f"（根拠には「{anchors[0]}{w}」と書かれています）。", "flip_dropped"))
                return


def _check_quote(cx: _Ctx, path: str, label: str, quote, evidence: str, issues: list) -> None:
    if quote is None or str(quote).strip() == "":
        return
    q = norm(quote)
    if q in norm(evidence):
        return
    if q in norm(cx.sent_text):
        issues.append(VerifyIssue("warning", path, f"{path}.{label} の引用「{quote}」は根拠のエントリではなく、"
                                                   "同じセルの別の記録にあります。", "quote_other"))
        return
    issues.append(VerifyIssue("error", path, f"{path}.{label} の「{quote}」は原文にありません（原文の語句をそのまま引用してください）。",
                              "quote"))


def _action_words(text: str, glossary: dict) -> list[list[str]]:
    groups = [list(g) for g in ACTION_GROUPS]
    for term, spec in (glossary or {}).items():
        to = spec.get("to", "") if isinstance(spec, dict) else str(spec)
        for g in groups:
            if term in g or any(x in to for x in g):
                g.append(str(term))
    t = norm(text)
    return [g for g in groups if any(norm(x) in t for x in g)]


def _check_content(cx: _Ctx, path: str, label: str, value: str, evidence: str, issues: list) -> None:
    """中身の語（2字以上の漢字・カタカナの連なり）が根拠に1つも無ければ警告（根拠に無いことを作った疑い）。

    言い換え（用語集の言い換えを含む）もあるので error にはしない（要確認にするだけ）。
    """
    runs = _CONTENT_RUN_RE.findall(nfkc(value))
    if not runs:
        return
    ev = norm(evidence)
    extra = []
    for term, spec in (cx.glossary or {}).items():
        to = spec.get("to", "") if isinstance(spec, dict) else str(spec)
        if term and norm(term) in ev:
            extra.append(norm(to))
        if to and norm(to) in ev:
            extra.append(norm(term))
    ev += " " + " ".join(extra)
    if not any(norm(r) in ev for r in runs):
        issues.append(VerifyIssue("warning", path, f"{path}.{label}「{nfkc(value).strip()}」の語は根拠のエントリに"
                                                   "1つも書かれていません。", "content"))


def _finish(rep: VerifyReport, path: str, issues: list) -> bool:
    rep.issues.extend(issues)
    if any(i.level == "error" for i in issues):
        rep.failed_items.append(path)
        return False
    rep.ok_items.append(path)
    return True


def _check_incident(inc: dict, cx: _Ctx, rep: VerifyReport) -> dict:
    out: dict = {}
    certainty_choices = cx.choices["certainty"]

    rc = inc.get("root_cause")
    if isinstance(rc, dict):
        path, issues = "incident.root_cause", []
        ev = _evidence(cx, rc.get("src"), path, rep, issues)
        v = rc.get("v")
        if not isinstance(v, str) or not v.strip():
            issues.append(VerifyIssue("error", path, f"{path}.v がありません。", "structure"))
        certainty = str(rc.get("certainty") or "")
        if certainty not in certainty_choices:
            issues.append(VerifyIssue("error", path, f"{path}.certainty「{certainty}」は選択肢にありません。", "choice"))
        if ev and isinstance(v, str):
            segs, text = ev
            _check_quote(cx, path, "q", rc.get("q"), text, issues)
            if certainty == "確定" and (rc.get("q") is None or not str(rc.get("q")).strip()):
                issues.append(VerifyIssue("error", path, f"{path}.certainty が「確定」のときは、q に根拠の原文の語句を"
                                                         "引用してください。", "quote"))
            _check_text_value(cx, path, "v", v, text, issues)
            _check_content(cx, path, "v", v, text, issues)
            unknown_in_ev = [w for w in UNKNOWN_WORDS if w in nfkc(text)]
            if certainty == "確定" and unknown_in_ev:
                issues.append(VerifyIssue("error", path, f"{path}.certainty が「確定」ですが、根拠 {', '.join(rc.get('src') or [])} "
                                                         f"には「{unknown_in_ev[0]}」と書かれています。", "speculation"))
            spec_in_ev = [w for w in SPECULATION_WORDS if w in nfkc(text)]
            if spec_in_ev:
                if certainty == "確定":
                    issues.append(VerifyIssue("error", path, f"{path}.certainty が「確定」ですが、根拠 {', '.join(rc.get('src') or [])} "
                                                             f"には「{spec_in_ev[0]}」と書かれています。", "speculation"))
                elif certainty != "疑い" and not any(w in nfkc(v) for w in SPECULATION_WORDS):
                    issues.append(VerifyIssue("warning", path, f"{path}.v で根拠の「{spec_in_ev[0]}」（推量）が消えています。",
                                              "speculation_lost"))
        if _finish(rep, path, issues):
            out["root_cause"] = {"q": rc.get("q"), "v": nfkc(v).strip(), "certainty": certainty,
                                 "src": list(rc.get("src") or []), "segs": ev[0] if ev else []}
    elif rc is not None:
        _finish(rep, "incident.root_cause", [VerifyIssue("error", "incident.root_cause",
                                                         "incident.root_cause の形が正しくありません。", "structure")])

    for key in ("temporary_actions", "permanent_actions"):
        items = inc.get(key) or []
        kept = []
        if not isinstance(items, list):
            _finish(rep, f"incident.{key}", [VerifyIssue("error", f"incident.{key}", f"incident.{key} が配列になっていません。",
                                                         "structure")])
            items = []
        for n, a in enumerate(items):
            path, issues = f"incident.{key}[{n}]", []
            if not isinstance(a, dict) or not isinstance(a.get("v"), str) or not a["v"].strip():
                _finish(rep, path, [VerifyIssue("error", path, f"{path}.v がありません。", "structure")])
                continue
            ev = _evidence(cx, a.get("src"), path, rep, issues)
            if ev:
                segs, text = ev
                _check_text_value(cx, path, "v", a["v"], text, issues, check_dropped=True)
                _check_content(cx, path, "v", a["v"], text, issues)
                for group in _action_words(a["v"], cx.glossary):
                    if not any(norm(x) in norm(text) for x in group):
                        issues.append(VerifyIssue("warning", path, f"{path}.v の「{group[0]}」は根拠のエントリに書かれていません。",
                                                  "action_word"))
                        break
            if _finish(rep, path, issues):
                kept.append({"v": nfkc(a["v"]).strip(), "src": list(a.get("src") or []), "segs": ev[0] if ev else []})
        out[key] = kept

    parts = inc.get("parts") or []
    kept_parts = []
    if not isinstance(parts, list):
        parts = []
    for n, p in enumerate(parts):
        path, issues = f"incident.parts[{n}]", []
        if not isinstance(p, dict) or not isinstance(p.get("name"), str) or not p["name"].strip():
            _finish(rep, path, [VerifyIssue("error", path, f"{path}.name がありません。", "structure")])
            continue
        ev = _evidence(cx, p.get("src"), path, rep, issues)
        if ev:
            segs, text = ev
            _check_text_value(cx, path, "name", p["name"], text, issues)
            _check_content(cx, path, "name", p["name"], text, issues)
            model = p.get("model")
            if model:
                if not _model_in(model, text) and not _model_in(model, cx.sent_text):
                    near = _similar_identifier(str(model), cx)
                    hint = f"（原文の表記は「{near}」）" if near else ""
                    issues.append(VerifyIssue("error", path, f"{path}.model の「{model}」は原文にありません{hint}。", "identifier"))
                elif not _model_in(model, text):
                    issues.append(VerifyIssue("warning", path, f"{path}.model の「{model}」は根拠のエントリにはありません。",
                                              "identifier_other"))
            _check_quote(cx, path, "qty_q", p.get("qty_q"), text, issues)
        if _finish(rep, path, issues):
            kept_parts.append({"name": nfkc(p["name"]).strip(), "model": p.get("model") or None,
                               "qty_q": p.get("qty_q") or None, "src": list(p.get("src") or []),
                               "segs": ev[0] if ev else []})
    out["parts"] = kept_parts

    rec = inc.get("recurrence")
    if isinstance(rec, dict):
        path, issues = "incident.recurrence", []
        ev = _evidence(cx, rec.get("src"), path, rep, issues)
        v = str(rec.get("v") or "").strip()
        if v not in ("あり", "なし"):
            issues.append(VerifyIssue("error", path, f"{path}.v は「あり」か「なし」にしてください。", "choice"))
        if ev:
            segs, text = ev
            _check_quote(cx, path, "count_q", rec.get("count_q"), text, issues)
            if v == "なし" and not _RECUR_NO_RE.search(nfkc(text)):
                issues.append(VerifyIssue("error", path, f"{path}.v「なし」の根拠が書かれていません。", "flip_added"))
            # 「あり」は根拠に再発の記録（「再発なし」「再発せず」ではないもの）か、根拠にある回数の引用が要る
            # 「再発はなし」「再発は見られない」など、再発しなかった書き方の部分は「あり」の根拠にしない
            count_q = str(rec.get("count_q") or "").strip()
            yes_text = _RECUR_NO_RE.sub(" ", nfkc(text))
            if v == "あり" and not _RECUR_YES_RE.search(yes_text) and not (count_q and norm(count_q) in norm(text)):
                issues.append(VerifyIssue("error", path, f"{path}.v「あり」の根拠（再発の記録）が書かれていません。",
                                          "flip_added"))
        if _finish(rep, path, issues):
            out["recurrence"] = {"v": v, "count_q": rec.get("count_q") or None, "src": list(rec.get("src") or []),
                                 "segs": ev[0] if ev else []}

    fs = inc.get("final_state")
    if isinstance(fs, dict):
        path, issues = "incident.final_state", []
        ev = _evidence(cx, fs.get("src"), path, rep, issues)
        v = str(fs.get("v") or "").strip()
        if v not in cx.choices["final_states"]:
            issues.append(VerifyIssue("error", path, f"{path}.v「{v}」は選択肢にありません。", "choice"))
        elif ev:
            words = _final_state_words(v, cx.glossary)
            ev_text = norm(ev[1] if v == "未着手" else _NOT_DONE_RE.sub(" ", nfkc(ev[1]))).upper()
            if v == "完了":
                # 「完了予定」「完了していない」「手配済」は完了の根拠にしない
                ev_text = norm(_NOT_YET_RE.sub(" ", _NOT_DONE_RE.sub(" ", nfkc(ev[1])))).upper()
                ev_text = _ARRANGED_RE.sub(" ", ev_text)
            hits = [w for w in words if norm(w).upper() in ev_text]
            if v == "完了" and hits == ["済"] and _PENDING_RE.search(ev_text):
                hits = []   # 「済」だけで「入荷待ち」も書かれている根拠は、完了の根拠にしない
            if words and not hits:
                issues.append(VerifyIssue("error", path, f"{path}.v「{v}」の根拠が書かれていません（根拠のエントリに"
                                                         f"「{words[0]}」などの語がありません）。", "flip_added"))
        if _finish(rep, path, issues):
            out["final_state"] = {"v": v, "src": list(fs.get("src") or []), "segs": ev[0] if ev else []}

    _check_missing_identifiers(out, cx, rep)
    return out


def _check_missing_identifiers(out: dict, cx: _Ctx, rep: VerifyReport) -> None:
    """抜けの疑い：部品・処置の根拠セグメントにある型番が、要点のどこにも出ていない（警告）。"""
    segs = set()
    for key in ("temporary_actions", "permanent_actions", "parts"):
        for item in out.get(key, []):
            segs.update(item.get("segs") or [])
    if not segs:
        return
    written = norm(" ".join(str(i.get("v") or "") + " " + str(i.get("name") or "") + " " + str(i.get("model") or "")
                            for key in ("temporary_actions", "permanent_actions", "parts") for i in out.get(key, [])))
    for sid in cx.seg_ids:
        if sid not in segs:
            continue
        for tok in _identifiers(cx.seg_text[sid]):
            if norm(tok) not in written and not re.fullmatch(r"(?i)ALM|ERR|E|AL", re.sub(r"[-\d]", "", tok)):
                rep.issues.append(VerifyIssue("warning", "incident.parts", f"{sid} の型番「{tok}」が部品・処置のどこにも出ていません。",
                                              "missing"))


def _check_summary(summary, cx: _Ctx, rep: VerifyReport, want_summary: bool) -> list[dict]:
    if summary is None:
        if want_summary:
            rep.issues.append(VerifyIssue("error", "summary", "summary がありません。", "structure"))
            rep.failed_items.append("summary")
        return []
    if not isinstance(summary, list):
        _finish(rep, "summary", [VerifyIssue("error", "summary", "summary が配列になっていません。", "structure")])
        return []
    kept = []
    for n, s in enumerate(summary):
        path, issues = f"summary[{n}]", []
        if isinstance(s, str):
            s = {"v": s, "src": []}
        if not isinstance(s, dict) or not isinstance(s.get("v"), str) or not s["v"].strip():
            _finish(rep, path, [VerifyIssue("error", path, f"{path}.v がありません。", "structure")])
            continue
        ev = _evidence(cx, s.get("src"), path, rep, issues)
        dates: list[str] = []
        if ev:
            segs, text = ev
            _check_text_value(cx, path, "v", s["v"], text, issues, limit=200, check_dropped=True)
            for sid in segs:
                when = cx.seg_by_id[sid].when
                d = when.date if when else None
                if d and d not in dates:
                    dates.append(d)
            if len(dates) > 1:
                issues.append(VerifyIssue("warning", path, f"{path} の根拠が複数の日（{min(dates)}〜{max(dates)}）にまたがっています。",
                                          "summary_days"))
        if _finish(rep, path, issues):
            item = {"v": nfkc(s["v"]).strip(), "src": list(s.get("src") or []), "segs": ev[0] if ev else []}
            if len(dates) == 1:
                item["date"] = dates[0]
            elif dates:
                item["date_from"], item["date_to"] = min(dates), max(dates)
            kept.append(item)
    return kept
