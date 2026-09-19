"""一覧表の Markdown 生成（記録ファイル・集計ファイル・データセット説明）。

決まり（docs/design.md 6章）: 同じ入力から同じバイト列。生成日時・取込ID・行番号の一覧を本文に書かない。
レコード内に空行を入れない。パイプ表を使わない。人名（person 役割）は既定で出さない。
追記ログ列は logproc のルール出力（対応の時系列）＋ AI 照合に通った要点だけを出す。
"""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date

from core.mdtext import escape_md_line, estimate_tokens, join_blocks, md_bullet
from core.naming import LIGHTRAG_HINT_RECORDS, hint_chunk_tokens, md_filename
from tables.spec import TableSpec
from tables.summaries import (
    category_column, dataset_counts, entity_columns, entity_display, entity_fiscal_year_summaries, entity_value,
    fmt_measure, fmt_number, is_month, measure_columns, month_first_day, month_label, month_last_day, month_summaries,
    unit_label,
)

# 記録1件の上限。LightRAG のヒントの chunk_ts から余裕（100トークン）を引く＝1件が1チャンクに収まる大きさ
RECORD_TOKEN_MARGIN = 100
RECORD_TOKEN_BUDGET = (hint_chunk_tokens() or 1500) - RECORD_TOKEN_MARGIN
TITLE_TEXT_CHARS = 40


@dataclass
class MdFile:
    name: str
    text: str
    kind: str  # dataset/records/summary

    @property
    def data(self) -> bytes:
        return self.text.encode("utf-8")


# ---- 追記ログ（logproc） ----------------------------------------------------------------

def people_index_for(spec: TableSpec, records: list[dict]):
    """記入者の判定に使う人物の索引（人物一覧＋担当列の値）。"""
    from logproc import PeopleIndex

    stage = spec.log_stage
    person_keys = [c.key for c in spec.columns if c.role == "person"]
    names = sorted({str(r.get("values", {}).get(k)) for r in records for k in person_keys
                    if r.get("values", {}).get(k)})
    return PeopleIndex(stage.people if stage else None, column_names=names,
                       groups=(stage.groups or None) if stage else None)


def _base_date(values: dict, spec: TableSpec) -> date | None:
    s = str(values.get(spec.date_key) or "")[:10]
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def parse_log_cell(spec: TableSpec, values: dict, people=None):
    """ログ列のセルをルールで分割する（マスク → parse_log）。AI 整形も同じ関数で分割して ID をそろえる。"""
    from logproc import PeopleIndex, SplitOptions, mask_text, parse_log

    stage = spec.log_stage
    if stage is None:
        return None
    text = values.get(stage.column)
    if text in (None, ""):
        text = ""
    people = people or PeopleIndex(stage.people, groups=stage.groups or None)
    rules = list(stage.mask or [])
    if rules:
        names = people.names() if any(r in ("person", "人名") for r in rules) else ()
        text, _spans = mask_text(str(text), rules, names=names)
    options = SplitOptions.from_dict(stage.splitter or {})
    return parse_log(str(text), base_date=_base_date(values, spec), people=people, options=options)


def _ai_item(ai_results: dict | None, key: str) -> tuple[str | None, dict]:
    item = (ai_results or {}).get(key)
    if not isinstance(item, dict):
        return None, {}
    status = str(item.get("status") or "ok")
    result = item.get("result") if isinstance(item.get("result"), dict) else item
    return status, result


def _glossary(text: str, glossary: dict | None) -> str:
    if not glossary or not text:
        return text
    from logproc import apply_glossary

    return apply_glossary(text, glossary)


def ai_point_lines(result: dict, glossary: dict | None = None) -> list[str]:
    """AI の incident から「対応の要点」の行を作る（照合に通った項目だけが result に残っている前提）。"""
    inc = result.get("incident") if isinstance(result.get("incident"), dict) else {}
    lines: list[str] = []

    def g(text) -> str:
        return _glossary(" ".join(str(text or "").split()), glossary)

    rc = inc.get("root_cause") if isinstance(inc.get("root_cause"), dict) else {}
    if rc.get("v"):
        lines.append(f"原因: {g(rc['v'])}" + (f"（{rc['certainty']}）" if rc.get("certainty") else ""))
    for field_name, label in (("temporary_actions", "暫定処置"), ("permanent_actions", "恒久処置")):
        items = [g(a.get("v")) for a in inc.get(field_name) or [] if isinstance(a, dict) and a.get("v")]
        if items:
            lines.append(f"{label}: {'、'.join(items)}")
    parts = []
    for p in inc.get("parts") or []:
        if not isinstance(p, dict):
            continue
        text = " ".join(x for x in (g(p.get("name")), str(p.get("model") or "").strip(),
                                     str(p.get("qty_q") or "").strip()) if x)
        if text:
            parts.append(text)
    if parts:
        lines.append(f"使用部品: {'、'.join(parts)}")
    rec = inc.get("recurrence") if isinstance(inc.get("recurrence"), dict) else {}
    if rec.get("v"):
        lines.append(f"再発: {g(rec['v'])}" + (f"（{rec['count_q']}）" if rec.get("count_q") else ""))
    fs = inc.get("final_state") if isinstance(inc.get("final_state"), dict) else {}
    if fs.get("v"):
        lines.append(f"最終状態: {g(fs['v'])}")
    return [escape_md_line(line) for line in lines]


def _segment_types(result: dict) -> dict[str, list[str]]:
    types: dict[str, list[str]] = {}
    for entry in result.get("entries") or []:
        if not isinstance(entry, dict):
            continue
        t = [str(x) for x in entry.get("t") or [] if x]
        for seg in entry.get("segs") or []:
            types.setdefault(str(seg), [])
            for x in t:
                if x not in types[str(seg)]:
                    types[str(seg)].append(x)
    return types


# ---- 値の書き方 ------------------------------------------------------------------------

def _one_line(text) -> str:
    return " ".join(str(text or "").split())


def _time_column(spec: TableSpec):
    date_key = spec.date_key
    for key in ("occurred_time", f"{date_key.removesuffix('_at').removesuffix('_date')}_time"):
        col = spec.column(key)
        if col is not None and col.type == "time":
            return col
    return None


def format_value(col, value) -> str:
    if value is None or value == "":
        return ""
    if col.type == "number" and isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{fmt_number(value)}{unit_label(col.unit)}"
    return str(value)


def _date_text(value, time_value=None) -> str:
    s = str(value)
    if time_value and len(s) == 10:
        s = f"{s} {time_value}"
    if is_month(s):
        s += f"（{month_label(s[:7])}）"
    return s


def _entity_label_name(spec: TableSpec) -> str:
    entity, _label = entity_columns(spec)
    if entity is None:
        return "設備"
    return "設備" if entity.key.startswith("equipment") else entity.display


def _is_hidden(col, spec: TableSpec) -> bool:
    if col.md == "omit":
        return True
    return col.role == "person" and bool((spec.markdown or {}).get("omit_person", True))


_BRACKETS = {"（": "）", "(": ")", "「": "」", "『": "』", "【": "】", "［": "］", "[": "]", "〔": "〕", "《": "》",
             "〈": "〉", "｛": "｝", "{": "}", "“": "”"}
_TITLE_TAG = re.compile(r"^[【\[][^】\]]{0,12}[】\]][ 　]*|^■?[ 　]*発生(?:日時)?[ 　]*[:：][ 　]*")
# 行頭の「R05.04.01 11:45(休日)、」のような日付・時刻（見出しの末尾に日付が入るので、現象の字数を使わない）
_TITLE_LEAD_DATE = re.compile(
    r"^(?:(?:[RHSrhs][ 　]?\d{1,2}|\d{2,4})[./年\-][ 　]?\d{1,2}[./月\-][ 　]?\d{1,2}日?"
    r"|\d{1,2}[:：]\d{2}(?:[:：]\d{2})?"
    r"|[（(][^）)]{0,10}[)）]"
    r"|[\s頃、,。.・~〜\-]+)+")


def _clip_title_text(text: str, limit: int) -> str:
    """見出し用に切り詰める。閉じない括弧の前で切り、切ったことが分かるように「…」を付ける。"""
    if limit <= 0 or len(text) <= limit:
        return text
    cut = text[:limit]
    opened: list[int] = []
    for i, ch in enumerate(cut):
        if ch in _BRACKETS:
            opened.append(i)
        elif opened and ch == _BRACKETS[cut[opened[-1]]]:
            opened.pop()
    if opened:
        cut = cut[:opened[0]]
    cut = cut.rstrip().rstrip("、,。.")
    return (cut or text[:limit].rstrip()) + "…"


def _title_text_line(raw) -> str:
    """見出しに使う文章。1行目の「【発生】」「発生日時:」などの札と、それに続く日付・時刻を外した残り。

    「【発生】R05.04.01 11:45(休日)」のように日付だけの1行目は空になる（見出しの日付と同じものを繰り返さない）。
    """
    first = _one_line(str(raw).split("\n")[0])
    body = _TITLE_TAG.sub("", first).strip()
    return _TITLE_LEAD_DATE.sub("", body).strip()


def record_title(values: dict, spec: TableSpec) -> str:
    """見出し: 【管理No】設備名（設備番号）現象の先頭40字｜日付"""
    md = spec.markdown or {}
    pieces: list[str] = []
    title_columns = list(md.get("title_columns") or [])
    entity, label = entity_columns(spec)
    if title_columns:
        for item in title_columns:
            key, _, length = str(item).partition(":")
            col = spec.column(key)
            if entity is not None and key in (entity.key, label.key if label else None):
                text = entity_display(values, spec)[2]
            else:
                raw = values.get(key)
                text = _title_text_line(raw) if raw not in (None, "") else ""
                limit = int(length) if length.isdigit() else TITLE_TEXT_CHARS
                if col is not None and col.type in ("text", "string") and len(text) > limit:
                    text = _clip_title_text(text, limit)
                if col is not None and col.role == "key" and text:
                    text = f"【{text}】"
            if text and text not in pieces:
                pieces.append(text)
    else:
        key_col = spec.first_role("key")
        if key_col is not None and values.get(key_col.key):
            pieces.append(f"【{_one_line(values[key_col.key])}】")
        display = entity_display(values, spec)[2]
        if display:
            pieces.append(display)
        text_col = spec.column("symptom") or spec.first_role("text")
        if text_col is not None and values.get(text_col.key):
            first = _title_text_line(values[text_col.key])
            if first:
                pieces.append(_clip_title_text(first, TITLE_TEXT_CHARS))
    title = ""
    for piece in pieces:
        if title and not title.endswith(("】", "）")):
            title += " "
        title += piece
    date_value = values.get(spec.date_key)
    if date_value:
        title += f"｜{str(date_value)[:10]}"
    return _one_line(title) or "（見出しなし）"


def record_block(record: dict, spec: TableSpec, ai_results: dict | None = None, people=None) -> list[str]:
    """1件分の行（見出し＋箇条書き）。空行は入れない。

    推定トークンが RECORD_TOKEN_BUDGET を超える場合だけ、時系列を切って収める（要点と時系列の合計で判定する）。
    """
    lines, timeline_tokens = _record_lines(record, spec, ai_results, people, None)
    if timeline_tokens:
        total = estimate_tokens("\n".join(lines))
        if total > RECORD_TOKEN_BUDGET:
            budget = RECORD_TOKEN_BUDGET - (total - timeline_tokens)
            lines, _ = _record_lines(record, spec, ai_results, people, max(0, budget))
    return lines


def _record_lines(record: dict, spec: TableSpec, ai_results: dict | None, people,
                  timeline_budget: int | None) -> tuple[list[str], int]:
    values = record.get("values", {}) or {}
    key = record.get("key", "")
    lines = [f"## {record_title(values, spec)}"]
    entity, label = entity_columns(spec)
    time_col = _time_column(spec)
    date_key = spec.date_key
    log_key = spec.log_stage.column if spec.log_stage else None
    entity_done = False
    timeline_tokens = 0
    status, result = _ai_item(ai_results, key)
    for col in spec.columns:
        if _is_hidden(col, spec):
            continue
        if time_col is not None and col.key == time_col.key and values.get(date_key):
            continue
        if entity is not None and label is not None and col.key in (entity.key, label.key):
            if entity_done:
                continue
            entity_done = True
            eid, name, display = entity_display(values, spec)
            if eid and name and not _is_hidden(entity, spec) and not _is_hidden(label, spec):
                lines += md_bullet(_entity_label_name(spec), display)
                continue
            for c in (entity, label):
                if not _is_hidden(c, spec):
                    lines += md_bullet(c.display, format_value(c, values.get(c.key)))
            continue
        if entity is not None and label is None and col.key == entity.key:
            # 設備名の列がない台帳。見出し・集計と同じ「設備名（設備番号）」の書き方にそろえる（列の名前はそのまま）
            _eid, name, display = entity_display(values, spec)
            if name and display:
                lines += md_bullet(col.display, display)
                continue
        value = values.get(col.key)
        if col.key == log_key:
            log, timeline_tokens = _log_lines(col, values, spec, status, result, people, timeline_budget)
            lines += log
            continue
        if value in (None, ""):
            continue
        if col.key == date_key:
            time_value = values.get(time_col.key) if time_col is not None else None
            lines += md_bullet(col.display, _date_text(value, time_value))
        else:
            lines += md_bullet(col.display, format_value(col, value))
    if status == "ok":
        for stage in spec.custom_stages:
            item = (result.get("custom") or {}).get(stage.id) if isinstance(result.get("custom"), dict) else None
            v = item.get("v") if isinstance(item, dict) else item
            if v:
                target = spec.column(stage.target_key) if stage.target_key else None
                lines += md_bullet(f"{target.display if target else stage.id}（AI分類）", _one_line(v))
    source = record.get("source", {}) or {}
    lines += md_bullet("出典", _source_text(values, spec, source))
    return lines, timeline_tokens


# ---- 時系列（他の列と重複する文を省く。取り込み設定の markdown.dedupe_timeline で切り替える） ----------

_SENTENCE_END = re.compile(r"(?<=[。、])")
_LEADING_NO = re.compile(r"^\s*\d{1,3}\s*[.)．）、]\s*")
_DUPLICATE_MIN_CHARS = 6  # これより短い文は偶然一致しうるので省かない


def _norm_sentence(text: str) -> str:
    """文の比較用。行頭の番号・空白・区切り記号を落として NFKC。"""
    s = unicodedata.normalize("NFKC", _LEADING_NO.sub("", str(text or "")))
    return "".join(s.split()).strip("。、.,:：;；")


def _column_sentences(values: dict, spec: TableSpec, log_key: str) -> dict[str, str]:
    """同じ記録の他の列（原因・処置内容・使用部品など）にある文 → その列の表示名。"""
    out: dict[str, str] = {}
    for col in spec.columns:
        if col.key == log_key or _is_hidden(col, spec) or col.type not in ("text", "string"):
            continue
        raw = values.get(col.key)
        if raw in (None, ""):
            continue
        for piece in re.split(r"[。、\n]+", str(raw)):
            key = _norm_sentence(piece)
            if len(key) >= _DUPLICATE_MIN_CHARS:
                out.setdefault(key, col.display)
    return out


def _dedupe_parse(parse, duplicates: dict[str, str]):
    """時系列の本文から、同じ記録の他の列と同じ文を省く（全部同じなら「（処置内容と同じ）」に縮める）。"""
    if not duplicates or parse.kind != "log":
        return parse
    changed = False
    segments = []
    for seg in parse.segments:
        body = seg.body
        if not body:
            segments.append(seg)
            continue
        kept, hit = [], []
        for piece in _SENTENCE_END.split(body):
            name = duplicates.get(_norm_sentence(piece))
            if name:
                hit.append(name)
                # 省いた文が「。」で終わっていたら、その「。」は残す（前後の文がつながって元にない1文になるため）
                prev = kept[-1].rstrip() if kept else ""
                if piece.rstrip().endswith("。") and prev and not prev.endswith("。"):
                    kept[-1] = prev.rstrip("、") + "。"
            else:
                kept.append(piece)
        if not hit:
            segments.append(seg)
            continue
        text = "".join(kept).strip().strip("、")
        segments.append(replace(seg, body=text or f"（{hit[0]}と同じ）"))
        changed = True
    return replace(parse, segments=segments) if changed else parse


def _fit_timeline(timeline: list[str], limit: int, budget: int) -> list[str]:
    """時系列を、設定の件数と残りの推定トークンに収める（決定的）。1件は必ず残す。"""
    costs = [estimate_tokens(line) + 1 for line in timeline]  # +1 は2字下げと改行の分
    n = min(len(timeline), max(1, limit))
    while n > 1 and sum(costs[:n]) + estimate_tokens(f"（以降{len(timeline) - n}件は管理用の正規化CSVに収録）") > budget:
        n -= 1
    if n >= len(timeline):
        return timeline
    return timeline[:n] + [f"（以降{len(timeline) - n}件は管理用の正規化CSVに収録）"]


def _log_lines(col, values: dict, spec: TableSpec, status, result: dict, people,
               timeline_budget: int | None) -> tuple[list[str], int]:
    stage = spec.log_stage
    parse = parse_log_cell(spec, values, people)
    if parse is None or parse.kind == "empty":
        return [], 0
    out: list[str] = []
    types: dict[str, list[str]] = {}
    if status == "ok" and result:
        points = ai_point_lines(result, stage.glossary)
        if points:
            out.append("- 対応の要点（AI抽出）:")
            out += [f"  {p}" for p in points]
        types = _segment_types(result)
    entity_label = entity_display(values, spec)[1] or entity_display(values, spec)[2]
    from logproc import render_timeline

    if (spec.markdown or {}).get("dedupe_timeline", True):
        parse = _dedupe_parse(parse, _column_sentences(values, spec, col.key))
    timeline = render_timeline(parse, entity_label, types=types or None, glossary=stage.glossary or None)
    if parse.kind == "header_cell":
        out += md_bullet(f"{col.display}（見出しごと）", timeline)
        return out, 0
    if not timeline:
        return out, 0
    if timeline_budget is not None:
        timeline = _fit_timeline(timeline, max(1, int(stage.max_timeline_entries or 20)), timeline_budget)
    out.append("- 対応の時系列:")
    out += [f"  {line}" for line in timeline]
    return out, estimate_tokens("\n".join(timeline))


def _source_text(values: dict, spec: TableSpec, source: dict) -> str:
    file_name = str(source.get("file") or "")
    key_col = spec.first_role("key")
    if key_col is not None and values.get(key_col.key):
        return f"{file_name}（{key_col.display} {_one_line(values[key_col.key])}）"
    row = source.get("row")
    sheet = source.get("sheet")
    if row:
        where = f"シート「{sheet}」{row}行目" if sheet else f"{row}行目"
        return f"{file_name}（{where}）"
    return file_name


# ---- ファイルの組み立て ---------------------------------------------------------------------

def _sort_key(record: dict, spec: TableSpec, time_col) -> tuple:
    values = record.get("values", {}) or {}
    d = str(values.get(spec.date_key) or "")
    t = str(values.get(time_col.key) or "") if time_col is not None else ""
    return (0 if d else 1, d, t, str(record.get("key", "")))


def _hint(spec: TableSpec) -> str | None:
    return LIGHTRAG_HINT_RECORDS if (spec.markdown or {}).get("lightrag_hint") else None


class _Names:
    """ファイル名の重複を避ける（安全化で同じ名前になった場合だけ _2 を付ける）。"""

    def __init__(self):
        self.used: set[str] = set()

    def make(self, parts: list[str], hint: str | None = None) -> str:
        name = md_filename(parts, hint)
        n = 2
        while name in self.used:
            name = md_filename(parts + [str(n)], hint)
            n += 1
        self.used.add(name)
        return name


def render_all(spec: TableSpec, records: list[dict], ai_results: dict | None, meta: dict | None) -> list[MdFile]:
    """全 md ファイルを作る（決定的）。records は確定済み全行（RecordRow.to_dict の形）。

    meta: {"coverage": {"start": "YYYY-MM", "end": "YYYY-MM"}}（取り込み範囲。なければ記録の日付から）
    """
    meta = dict(meta or {})
    md = spec.markdown or {}
    names = _Names()
    prefix = spec.file_prefix
    coverage = _coverage(meta.get("coverage"), records, spec)
    files: list[MdFile] = []
    time_col = _time_column(spec)
    ordered = sorted(records, key=lambda r: _sort_key(r, spec, time_col))

    if md.get("dataset_card", True):
        name = names.make([prefix, "00", "データセット説明"])
        files.append(MdFile(name, render_dataset_card(spec, ordered, coverage), "dataset"))

    if md.get("records", True):
        files += _record_files(spec, ordered, ai_results or {}, names)

    for summary in spec.summaries():
        if summary.id == "month":
            for item in month_summaries(ordered, spec, coverage, summary):
                name = names.make([prefix, "集計", "月次", item["month"]])
                files.append(MdFile(name, _render_month_summary(spec, item, coverage), "summary"))
        elif summary.id == "entity_fiscal_year":
            entity, _label = entity_columns(spec)
            group_word = "設備別" if (entity is not None and entity.key.startswith("equipment")) or entity is None \
                else f"{entity.display}別"
            for item in entity_fiscal_year_summaries(ordered, spec, coverage, summary):
                name = names.make([prefix, "集計", group_word, item["entity"], f"{item['fiscal_year']}年度"])
                files.append(MdFile(name, _render_entity_fy(spec, item, coverage), "summary"))
    return files


def _coverage(coverage: dict | None, records: list[dict], spec: TableSpec) -> dict:
    counts = dataset_counts(records, spec)
    start = (coverage or {}).get("start") or (counts["months"][0] if counts["months"] else None)
    end = (coverage or {}).get("end") or (counts["months"][-1] if counts["months"] else None)
    if counts["months"]:
        start = min(start, counts["months"][0]) if start else counts["months"][0]
        end = max(end, counts["months"][-1]) if end else counts["months"][-1]
    # 最初と最後の記録の日（月の途中で始まる・終わるデータを、月全体と書かないため）
    first = counts["date_min"] if len(str(counts["date_min"] or "")) == 10 else None
    last = counts["date_max"] if len(str(counts["date_max"] or "")) == 10 else None
    return {"start": start, "end": end, "first_date": first, "last_date": last}


def _period_text(start_month: str, end_month: str, coverage: dict) -> str:
    """start_month〜end_month の期間。取り込んだ記録の最初・最後の月なら、実際の最初・最後の日で書く。"""
    first, last = coverage.get("first_date"), coverage.get("last_date")
    a = first if first and first[:7] == start_month else month_first_day(start_month)
    b = last if last and last[:7] == end_month else month_last_day(end_month)
    return f"{a}〜{b}"


def _record_files(spec: TableSpec, ordered: list[dict], ai_results: dict, names: _Names) -> list[MdFile]:
    md = spec.markdown or {}
    prefix = spec.file_prefix
    per_file = max(1, int(md.get("max_records_per_file") or 300))
    by_entity = md.get("group_by") == "entity_month"
    entity, _label = entity_columns(spec)
    groups: dict[tuple, list[dict]] = defaultdict(list)
    for rec in ordered:
        values = rec.get("values", {}) or {}
        d = str(values.get(spec.date_key) or "")
        month = d[:7] if is_month(d) else ""
        eid = entity_value(values, spec)[0] if (by_entity and entity is not None) else ""
        groups[(eid, month) if by_entity else ("", month)].append(rec)
    people = people_index_for(spec, ordered) if spec.log_stage else None
    hint = _hint(spec)
    files: list[MdFile] = []
    for (eid, month) in sorted(groups, key=lambda g: (g[0], g[1] == "", g[1])):
        recs = groups[(eid, month)]
        chunks = [recs[i:i + per_file] for i in range(0, len(recs), per_file)]
        for n, chunk in enumerate(chunks, start=1):
            parts = [prefix]
            if by_entity:
                parts.append(eid or "設備不明")
            parts.append(month or "日付なし")
            if n > 1:
                parts.append(f"part{n}")
            name = names.make(parts, hint)
            start_no = (n - 1) * per_file + 1
            header = _record_file_header(spec, recs, chunk, month, eid, n, len(chunks), start_no)
            blocks = [header] + [record_block(r, spec, ai_results, people) for r in chunk]
            files.append(MdFile(name, join_file(blocks), "records"))
    return files


def _record_file_header(spec: TableSpec, group: list[dict], chunk: list[dict], month: str, eid: str,
                        part: int, parts: int, start_no: int) -> list[str]:
    name = spec.name
    scope_month = month_label(month) if month else "日付なし"
    entity_disp = ""
    if eid:
        entity_disp = entity_display(group[0].get("values", {}), spec)[2] or eid
    title = f"# {name}"
    if entity_disp:
        title += f" {entity_disp}"
    title += f" {scope_month}の記録" if month else " 日付なしの記録"
    if parts > 1:
        title += f"（{part}/{parts}）"
    body = [f"- データ種別: {name}（1行＝1件）の記録"]
    if entity_disp:
        body += md_bullet(_entity_label_name(spec), entity_disp)
    if month:
        body.append(f"- 対象期間: {month_first_day(month)}〜{month_last_day(month)}")
    scope = f"{eid}の{scope_month}" if eid else scope_month
    if parts > 1:
        end_no = start_no + len(chunk) - 1
        body.append(f"- このファイルの記録: {len(chunk):,}件（{scope}の全 {len(group):,}件のうち {start_no:,}〜{end_no:,}件目）")
    else:
        body.append(f"- このファイルの記録: {len(chunk):,}件（{scope}の全件）")
    return _Block(title, body)


class _Block(list):
    """ファイル先頭の「# 見出し / 空行 / 本文」。join_blocks はブロック内の空行を落とすので別ブロックに分けて出す。"""

    def __init__(self, title: str, body: list[str]):
        super().__init__([title] + body)
        self.title = title
        self.body = body


def join_file(blocks: list[list[str]]) -> str:
    expanded: list[list[str]] = []
    for block in blocks:
        if isinstance(block, _Block):
            expanded.append([block.title])
            expanded.append(block.body)
        else:
            expanded.append(block)
    return join_blocks(expanded)


# ---- データセット説明 ------------------------------------------------------------------------

def _range_text(coverage: dict) -> str:
    start, end = coverage.get("start"), coverage.get("end")
    if not start or not end:
        return "日付なし"
    return _period_text(start, end, coverage)


def render_dataset_card(spec: TableSpec, records: list[dict], coverage: dict) -> str:
    counts = dataset_counts(records, spec)
    name = spec.name
    entity, label = entity_columns(spec)
    measures = measure_columns(spec)
    cat = category_column(spec)
    entity_word = _entity_label_name(spec)
    head = [f"# データセット説明：{name}"]
    body: list[str] = []
    body.append(f"- データ種別: 表データ（1行＝1件の{name}）をRAG用に変換した資料群の説明")
    body.append(f"- 取り込み範囲: {_range_text(coverage)}（記録 {counts['records']:,}件）")
    if spec.description:
        body += md_bullet("説明", spec.description)
    if entity is not None:
        body.append(f"- 記録に出てくる{entity_word}: {counts['entities']:,}件（期間中に1件以上の記録がある{entity_word}だけです。"
                    f"{entity_word}の一覧ではありません）")

    structure = ["## 資料の構成"]
    md = spec.markdown or {}
    if md.get("records", True):
        per_file = int(md.get("max_records_per_file") or 300)
        if md.get("group_by") == "entity_month":
            structure.append(f"- 記録（{entity_word}×月）: 1件ごとの内容。各ファイルには、その{entity_word}・その月の記録を全件載せています"
                             f"（{per_file:,}件を超える場合は複数ファイルに分けています）。")
        else:
            structure.append(f"- 記録（月ごと）: 1件ごとの内容。各ファイルには、その月の記録を全件載せています"
                             f"（{per_file:,}件を超える月は複数ファイルに分けています）。")
    measure_words = "・".join(d for _k, d, _u in measures)
    for summary in spec.summaries():
        if summary.id == "month":
            text = f"- 月次集計: 月ごとの全{entity_word}の件数"
            if measure_words:
                text += f"、{measure_words}の合計"
            if entity is not None:
                text += f"、上位{summary.top_n}件の{entity_word}"
            if cat is not None:
                text += f"、{cat.display}の内訳"
            structure.append(text + "。")
        elif summary.id == "entity_fiscal_year":
            text = f"- {entity_word}別年度集計: {entity_word}ごとの年度内の件数"
            if measure_words:
                text += f"、{measure_words}の合計・平均"
            text += "、月別の内訳"
            if cat is not None:
                text += f"、{cat.display}の内訳"
            structure.append(text + "。")

    notes = ["## 数値についての注意"]
    units = [s.id for s in spec.summaries()]
    unit_words = []
    if "entity_fiscal_year" in units:
        unit_words.append(f"{entity_word}×年度")
    if "month" in units:
        unit_words.append(f"月×全{entity_word}")
    if unit_words:
        notes.append(f"- 件数・合計・平均は、集計ファイルにある単位（{'、'.join(unit_words)}）でのみ、アプリが元データから計算しています。")
    notes.append("- それ以外の条件（例: 特定の区分だけの月別の平均）の件数・合計は、この資料群からは確定できません。")
    if entity is not None:
        notes.append(f"- 記録が1件もない{entity_word}は、この資料群には出てきません。")

    columns = ["## 列の意味"]
    for col in spec.columns:
        if _is_hidden(col, spec):
            continue
        text = col.description or f"元の見出し「{col.headers[0] if col.headers else col.display}」"
        if col.unit and col.type == "number":
            text += f"（単位: {unit_label(col.unit)}）"
        columns += md_bullet(col.display, text)
    # 出していない列も名前だけは書く（コードのままでは意味が分からない列を既定で外しているため）。理由が違うので人名の列は分ける
    people_cols = [col.display for col in spec.columns if col.role == "person" and _is_hidden(col, spec)]
    omitted = [col.display for col in spec.columns if col.md == "omit" and col.role != "person"]
    if omitted:
        columns.append(f"- 記録に出していない列: {'、'.join(omitted)}"
                       "（意味の分からないコード値・管理用の列などのため。元の値は管理用の正規化CSVにあります）")
    if people_cols:
        columns.append(f"- 記録に出していない列（人名）: {'、'.join(people_cols)}"
                       "（人名のため出していません。取り込み設定で出すようにできます。元の値は管理用の正規化CSVにあります）")

    questions = ["## 答えられる質問の例"]
    no_questions = ["## 答えられない質問の例"]
    m0 = measures[0][1] if measures else None
    questions.append(f"- 特定の{entity_word}で過去に起きた記録と、その内容（検索で取り出された記録の範囲。全件の列挙は保証しません）")
    if "month" in units:
        questions.append(f"- ある月の{name}の件数" + (f"、{m0}の合計、{m0}が長い{entity_word}（上位{spec.summaries()[0].top_n}件）" if m0 else ""))
    if "entity_fiscal_year" in units and entity is not None:
        questions.append(f"- ある{entity_word}のある年度の件数" + (f"と{m0}の合計" if m0 else "") + "、月別の件数")
    no_questions.append("- 集計ファイルにない条件の件数・合計・平均（例: 複数の条件を組み合わせた件数）")
    if entity is not None:
        no_questions.append(f"- 記録が1件もなかった{entity_word}（{entity_word}の一覧を持っていません）")
    return join_file([_Block(head[0], body), structure, notes, columns, questions, no_questions])


# ---- 集計ファイル --------------------------------------------------------------------------

def _measure_meta(spec: TableSpec) -> dict[str, tuple[str, str]]:
    return {k: (d, u) for k, d, u in measure_columns(spec)}


def _render_month_summary(spec: TableSpec, item: dict, coverage: dict) -> str:
    name = spec.name
    month = item["month"]
    ml = month_label(month)
    mm = _measure_meta(spec)
    entity, _label = entity_columns(spec)
    entity_word = _entity_label_name(spec)
    cat = category_column(spec)
    head = f"# {name} 月次集計 {ml}"
    body = [
        f"- データ種別: {name}からアプリが計算した集計値（AIは使っていません）",
        f"- 集計対象: {_period_text(month, month, coverage)} の記録（{name}の取り込み範囲: {_range_text(coverage)}）",
    ]
    overview = ["## 概要", f"- {ml}の{name}の記録は{item['count']:,}件です。"]
    for key, total in item["sums"].items():
        display, unit = mm.get(key, (key, ""))
        overview.append(f"- {ml}の{display}の合計は{fmt_measure(total, unit, with_hours=True)}です。")
    blocks: list = [_Block(head, body), overview]
    if item["count"] and item["top"]:
        rank_key = item["rank_key"]
        if rank_key:
            display, unit = mm.get(rank_key, (rank_key, ""))
            top = [f"## {display}が長い{entity_word}（上位{len(item['top'])}件）"]
        else:
            top = [f"## 記録が多い{entity_word}（上位{len(item['top'])}件）"]
        for n, t in enumerate(item["top"], start=1):
            parts = []
            if rank_key:
                display, unit = mm.get(rank_key, (rank_key, ""))
                parts.append(f"{display} {fmt_measure(t['sum'] or 0, unit)}")
            parts.append(f"{t['count']:,}件")
            if t["main_category"] and cat is not None:
                parts.append(f"主な{cat.display}: {t['main_category']}")
            top.append(f"- {n}位: {t['display']} {'、'.join(parts)}")
        blocks.append(top)
    if item["count"] and item["categories"] and cat is not None:
        cats = [f"## {cat.display}の内訳"]
        rank_key = item["rank_key"]
        for c in item["categories"]:
            text = f"{c['count']:,}件"
            if rank_key and c["sum"] is not None:
                display, unit = mm.get(rank_key, (rank_key, ""))
                text += f"、{display} {fmt_measure(c['sum'], unit)}"
            cats += md_bullet(c["name"], text)
        blocks.append(cats)
    return join_file(blocks)


def _render_entity_fy(spec: TableSpec, item: dict, coverage: dict) -> str:
    name = spec.name
    fy = item["fiscal_year"]
    mm = _measure_meta(spec)
    entity, label = entity_columns(spec)
    eid = item["entity"]
    start, end = item["range"]
    period_words = f"{month_label(start)}〜{month_label(end)}"
    head = f"# {name} {_entity_label_name(spec)}別年度集計 {item['display']} {fy}年度"
    body = [f"- データ種別: {name}からアプリが計算した集計値（AIは使っていません）"]
    body += md_bullet(entity.display, eid)
    if label is not None and item["name"]:
        body += md_bullet(label.display, item["name"])
    body.append(f"- 集計対象: {_period_text(start, end, coverage)} の記録（{name}の取り込み範囲: {_range_text(coverage)}）")
    overview = ["## 概要"]
    metrics = item["metrics"]
    overview.append(f"- {eid}の{fy}年度（{period_words}）の記録は{item['count']:,}件です。")
    for key in item["keys"]:
        display, unit = mm.get(key, (key, ""))
        st = item["stats"][key]
        if not st["n"]:
            overview.append(f"- {display}の値はありません。")  # 空欄は 0 ではない
            continue
        text = f"- {display}の合計は{fmt_measure(st['sum'], unit, with_hours=True)}"
        if key in metrics["avg"] and st["avg"] is not None:
            base = f"（値のある{st['n']:,}件）" if st["n"] < st.get("rows", st["n"]) else ""
            text += f"、1件あたり平均{base}{fmt_measure(st['avg'], unit)}"
        overview.append(text + "です。")
        if key in metrics["max"] and st["max"]:
            overview.append(f"- 1件の{display}の最大は{fmt_measure(st['max']['value'], unit)}（{month_label(st['max']['month'])}）です。")
    if item["categories"]:
        cat = category_column(spec)
        detail = "、".join(f"{c['name']} {c['count']:,}件" for c in item["categories"])
        overview.append(f"- {cat.display}の内訳: {detail}。")
    months = ["## 月別"]
    for mi in item["months"]:
        parts = [f"{mi['count']:,}件"]
        for key in item["keys"]:
            display, unit = mm.get(key, (key, ""))
            if key in metrics["sum"]:
                if mi["count"] and not mi["has_value"][key]:
                    parts.append(f"{display} 値なし")  # 記録はあるが値がすべて空欄（0 と書かない）
                else:
                    parts.append(f"{display} {fmt_measure(mi['sums'][key], unit)}")
        months.append(f"- {month_label(mi['month'])}: {'、'.join(parts)}")
    return join_file([_Block(head, body), overview, months])


__all__ = ["MdFile", "ai_point_lines", "parse_log_cell", "people_index_for", "record_block",
           "record_title", "render_all", "render_dataset_card"]
