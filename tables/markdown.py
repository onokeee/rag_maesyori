"""一覧表の Markdown 生成（記録ファイル・集計ファイル・データセット説明）。

決まり（docs/design.md 6章）: 同じ入力から同じバイト列。生成日時・取込ID・行番号の一覧を本文に書かない。
レコード内に空行を入れない。パイプ表を使わない。出す列の中身は1文字も削らない（人名・コードの列も出す）。
追記ログ列は logproc のルール出力（対応の時系列）＋ AI 照合に通った要点だけを出す。
大きい記録は「（続きn/m）」に分けて、どの部分にも管理No・設備・日付を書く（切られても身元が分かるように）。
"""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date

from core.mdtext import escape_md_line, estimate_tokens, join_blocks, md_bullet
from core.naming import md_filename
from tables.spec import TableSpec, base_date_from
from tables.summaries import (
    category_column, dataset_counts, entity_columns, entity_display, entity_fiscal_year_summaries, entity_value,
    fmt_measure, fmt_number, is_month, measure_columns, month_first_day, month_label, month_last_day, month_summaries,
    unit_label,
)

# 記録1件の上限（推定トークン）。これを超えたら「（続きn/m）」に分ける。
# 取り込み側の設定は見えないので、狭い固定窓（600/overlap 50）でも記録が途中で切られない大きさにする。
# オフライン評価（T1・T2・T5 × F1200/100・F600/50・P2000）で 300/400/600 を比べて 400 を選んだ（docs/design.md 6.5）。
RECORD_TOKEN_BUDGET = 400
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
    # AI整形（aiproc.runner）と同じ探し方にする。片方だけ日付が出て md が「年不明」になるのを防ぐ
    return base_date_from(values, spec)


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
    """出さない列か。画面で「出さない」にした列だけ（人名・コードの列も既定では出す）。"""
    return col.md == "omit"


_BRACKETS = {"（": "）", "(": ")", "「": "」", "『": "』", "【": "】", "［": "］", "[": "]", "〔": "〕", "《": "》",
             "〈": "〉", "｛": "｝", "{": "}", "“": "”"}
_TITLE_TAG = re.compile(r"^[【\[][^】\]]{0,12}[】\]][ 　]*|^■?[ 　]*発生(?:日時)?[ 　]*[:：][ 　]*")
# 行頭の「R05.04.01 11:45(休日)、」のような日付・時刻（見出しの末尾に日付が入るので、現象の字数を使わない）
_TITLE_LEAD_DATE = re.compile(
    r"^(?:(?:[RHSrhs][ 　]?\d{1,2}|\d{2,4})[./年\-][ 　]?\d{1,2}[./月\-][ 　]?\d{1,2}日?"
    r"|\d{1,2}[:：]\d{2}(?:[:：]\d{2})?"
    r"|[（(][^）)]{0,10}[)）]"
    r"|[\s頃、,。.・~〜\-]+)+")
# 日付・時刻のすぐ後にこれが続くときは、時刻が文の一部（「0:28以降、…」）なので外さない
_TITLE_LEAD_KEEP = ("以降", "以後", "以前", "から", "まで", "より", "前後", "過ぎ", "すぎ")


def _strip_lead_date(text: str) -> str:
    m = _TITLE_LEAD_DATE.match(text)
    if m is None or text[m.end():].startswith(_TITLE_LEAD_KEEP):
        return text
    return text[m.end():]


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


# 2行目以降の行頭の「現象:」「設備：」のような項目名（数字で始まるもの＝時刻は含めない）
_TITLE_LABEL = re.compile(r"^(?![\d０-９])[^\s:：、。]{1,8}[ 　]*[:：][ 　]*")


def _title_text_line(raw) -> str:
    """見出しに使う文章。1行目の「【発生】」「発生日時:」などの札と、それに続く日付・時刻を外した残り。

    「【発生】R05.04.01 11:45(休日)」のように日付だけの1行目は使わず（見出しの日付と同じものを繰り返さない）、
    次の行（「設備:」の行は飛ばす）の札・項目名・日付を外した残りを使う。
    """
    lines = str(raw).split("\n")
    first = _one_line(lines[0])
    body = _strip_lead_date(_TITLE_TAG.sub("", first).strip()).strip()
    if body or not first:
        return body
    # 1行目が札と日付だけ（「発生:2024-04-28 14:50」）なら、次の行から探す。「設備:」の行は見出しの設備と重なるので飛ばす
    for line in lines[1:]:
        line = _one_line(line)
        tag = _TITLE_TAG.match(line)
        text = _strip_lead_date(_TITLE_TAG.sub("", line).strip()).strip()
        label = _TITLE_LABEL.match(text)
        if (tag and "設備" in tag.group(0)) or (label and "設備" in label.group(0)):
            continue
        text = _strip_lead_date(_TITLE_LABEL.sub("", text).strip()).strip()
        if text:
            return text
    return ""


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


def record_blocks(record: dict, spec: TableSpec, ai_results: dict | None = None, people=None) -> list[list[str]]:
    """1件分のブロック。推定トークンが RECORD_TOKEN_BUDGET を超えたら「（続きn/m）」に分ける。

    分けても文字は1つも消さない。2つ目以降には管理No・設備・日付を書き直す（その部分だけで身元が分かるように）。
    """
    lines, repeat = _record_lines(record, spec, ai_results, people)
    joined = "\n".join(lines)
    # 推定は1文字あたり最大1.1トークン。それでも上限以下なら数えずに済む（大半の記録。判定の結果は同じ）
    if len(joined) * 11 <= RECORD_TOKEN_BUDGET * 10 or estimate_tokens(joined) <= RECORD_TOKEN_BUDGET:
        return [lines]
    return _split_record(lines[0], lines[1:], repeat)


def record_block(record: dict, spec: TableSpec, ai_results: dict | None = None, people=None) -> list[str]:
    """1件分の行（画面の下書き表示用）。分かれる記録は続きの見出しも含めて続けて返す。"""
    return [line for block in record_blocks(record, spec, ai_results, people) for line in block]


def _record_lines(record: dict, spec: TableSpec, ai_results: dict | None, people) -> tuple[list[str], list[str]]:
    """1件分の行と、分けたときに書き直す行（管理No・設備・日付）。"""
    values = record.get("values", {}) or {}
    key = record.get("key", "")
    lines = [f"## {record_title(values, spec)}"]
    entity, label = entity_columns(spec)
    key_col = spec.first_role("key")
    time_col = _time_column(spec)
    date_key = spec.date_key
    log_key = spec.log_stage.column if spec.log_stage else None
    entity_done = False
    repeat: list[str] = []
    status, result = _ai_item(ai_results, key)

    def add(col, bullet: list[str]) -> None:
        """行を足す。管理No・設備・日付の行は、記録を分けたときに書き直すので控えておく。"""
        lines.extend(bullet)
        if col is not None and (col.key == date_key or (key_col is not None and col.key == key_col.key)
                                or (entity is not None and col.key in (entity.key, label.key if label else ""))):
            repeat.extend(bullet)

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
                add(entity, md_bullet(_entity_label_name(spec), display))
                continue
            for c in (entity, label):
                if not _is_hidden(c, spec):
                    add(c, md_bullet(c.display, format_value(c, values.get(c.key))))
            continue
        if entity is not None and label is None and col.key == entity.key:
            # 設備名の列がない台帳。見出し・集計と同じ「設備名（設備番号）」の書き方にそろえる（列の名前はそのまま）
            _eid, name, display = entity_display(values, spec)
            if name and display:
                add(col, md_bullet(col.display, display))
                continue
        value = values.get(col.key)
        if col.key == log_key:
            lines += _log_lines(col, values, spec, status, result, people)
            continue
        if value in (None, ""):
            continue
        if col.key == date_key:
            time_value = values.get(time_col.key) if time_col is not None else None
            add(col, md_bullet(col.display, _date_text(value, time_value)))
        else:
            add(col, md_bullet(col.display, format_value(col, value)))
    if status == "ok":
        for stage in spec.custom_stages:
            item = (result.get("custom") or {}).get(stage.id) if isinstance(result.get("custom"), dict) else None
            v = item.get("v") if isinstance(item, dict) else item
            if v:
                target = spec.column(stage.target_key) if stage.target_key else None
                lines += md_bullet(f"{target.display if target else stage.id}（AI分類）", _one_line(v))
    source = record.get("source", {}) or {}
    lines += md_bullet("出典", _source_text(values, spec, source))
    return lines, repeat


# ---- 大きい記録を「（続きn/m）」に分ける（文字は1つも消さない） ------------------------------

_MIN_PART_TOKENS = 80          # 見出し・書き直す行を引いても、これだけは中身に使う
_BULLET_LABEL = re.compile(r"^- ([^:]{1,40}): ")


def _tokens(lines: list[str]) -> int:
    return estimate_tokens("\n".join(lines)) if lines else 0


def _bullet_groups(body: list[str]) -> list[list[str]]:
    """箇条書きのまとまり（`- 項目:` の行と、それに続く2字下げの行）に分ける。"""
    groups: list[list[str]] = []
    for line in body:
        if line.startswith("- ") or not groups:
            groups.append([line])
        else:
            groups[-1].append(line)
    return groups


def _split_group(group: list[str], budget: int) -> list[list[str]]:
    """1つの箇条書きが budget に収まらないときに小分けにする（行も文も消さない）。"""
    if _tokens(group) <= budget:
        return [group]
    head = group[0]
    if len(group) > 1:
        # 複数行の値（対応の時系列など）。見出しの行を「（続き）」で繰り返して2字下げの行を分ける
        cont = f"{head[:-1]}（続き）:" if head.endswith(":") else head
        out, cur = [], [head]
        for line in group[1:]:
            if len(cur) > 1 and _tokens(cur) + estimate_tokens(line) > budget:
                out.append(cur)
                cur = [cont]
            cur.append(line)
        return out + [cur]
    m = _BULLET_LABEL.match(head)
    if m is None:
        return [group]  # 項目名が読めない1行。切らずにそのまま出す
    # 1行の長い値。「。」の後ろで分け、項目名を「（続き）」で繰り返す
    label, text = m.group(1), head[m.end():]
    pieces = [p for p in re.split(r"(?<=。)", text) if p]
    out, cur = [], ""
    for piece in pieces:
        if cur and estimate_tokens(f"- {label}（続き）: {cur}{piece}") > budget:
            out.append([f"- {label}: {cur}" if not out else f"- {label}（続き）: {cur}"])
            cur = ""
        cur += piece
    if cur:
        out.append([f"- {label}: {cur}" if not out else f"- {label}（続き）: {cur}"])
    return out or [group]


def _split_record(title: str, body: list[str], repeat: list[str]) -> list[list[str]]:
    """見出し・本文を、1つあたり RECORD_TOKEN_BUDGET に収まるブロックの並びにする。"""
    # 2つ目以降は「見出し（続きn/m）」＋管理No・設備・日付を書き直すので、その分を引いた残りが中身に使える
    overhead = estimate_tokens(f"{title}（続き00/00）") + _tokens(repeat)
    budget = max(_MIN_PART_TOKENS, RECORD_TOKEN_BUDGET - overhead)
    groups = [g for group in _bullet_groups(body) for g in _split_group(group, budget)]
    parts: list[list[str]] = []
    cur: list[str] = []
    for g in groups:
        if cur and _tokens(cur) + _tokens(g) > budget:
            parts.append(cur)
            cur = []
        cur += g
    if cur or not parts:
        parts.append(cur)
    if len(parts) == 1:
        return [[title] + parts[0]]
    total = len(parts)
    out = []
    for i, part in enumerate(parts, start=1):
        head = f"{title}（{i}/{total}）" if i == 1 else f"{title}（続き{i}/{total}）"
        again = [] if i == 1 else [ln for ln in repeat if ln not in part]
        out.append([head] + again + part)
    return out


# ---- 追記ログ列の行 -------------------------------------------------------------------


def _log_lines(col, values: dict, spec: TableSpec, status, result: dict, people) -> list[str]:
    """ログ列の行（対応の要点＋対応の時系列）。書かれた文は1つも省かない。"""
    stage = spec.log_stage
    parse = parse_log_cell(spec, values, people)
    if parse is None or parse.kind == "empty":
        return []
    out: list[str] = []
    types: dict[str, list[str]] = {}
    if status == "ok" and result:
        points = ai_point_lines(result, stage.glossary)
        if points:
            out.append("- 対応の要点（AI抽出）:")
            out += [f"  {p}" for p in points]
        types = _segment_types(result)
    _eid, entity_name, entity_text = entity_display(values, spec)
    entity_label = entity_name or entity_text
    from logproc import render_timeline

    timeline = render_timeline(parse, entity_label, types=types or None, glossary=stage.glossary or None)
    if parse.kind == "header_cell":
        out += md_bullet(f"{col.display}（見出しごと）", timeline)
        return out
    if not timeline:
        return out
    out.append("- 対応の時系列:")
    out += [f"  {line}" for line in timeline]
    return out


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


class _Names:
    """ファイル名の重複を避ける（安全化で同じ名前になった場合だけ _2 を付ける）。

    Windows のフォルダは大文字・小文字を区別しないので、比較は casefold して行う
    （「ETC-302号機」と「Etc-302号機」が同じファイルに上書きされて記録が消えるのを防ぐ）。
    """

    def __init__(self):
        self.used: set[str] = set()

    def make(self, parts: list[str]) -> str:
        name = md_filename(parts)
        n = 2
        while name.casefold() in self.used:
            name = md_filename(parts + [str(n)])
            n += 1
        self.used.add(name.casefold())
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
    """記録ファイル（月ごと、または設備×月ごとに1ファイル。件数では分けない）。"""
    md = spec.markdown or {}
    prefix = spec.file_prefix
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
    files: list[MdFile] = []
    for (eid, month) in sorted(groups, key=lambda g: (g[0], g[1] == "", g[1])):
        recs = groups[(eid, month)]
        parts = [prefix]
        if by_entity:
            parts.append(eid or "設備不明")
        parts.append(month or "日付なし")
        blocks: list[list[str]] = [_record_file_header(spec, recs, month, eid)]
        for r in recs:
            blocks += record_blocks(r, spec, ai_results, people)
        files.append(MdFile(names.make(parts), join_file(blocks), "records"))
    return files


def _record_file_header(spec: TableSpec, group: list[dict], month: str, eid: str) -> list[str]:
    name = spec.name
    scope_month = month_label(month) if month else "日付なし"
    entity_disp = ""
    if eid:
        entity_disp = entity_display(group[0].get("values", {}), spec)[2] or eid
    title = f"# {name}"
    if entity_disp:
        title += f" {entity_disp}"
    title += f" {scope_month}の記録" if month else " 日付なしの記録"
    body = [f"- データ種別: {name}（1行＝1件）の記録"]
    if entity_disp:
        body += md_bullet(_entity_label_name(spec), entity_disp)
    if month:
        body.append(f"- 対象期間: {month_first_day(month)}〜{month_last_day(month)}")
    scope = f"{eid}の{scope_month}" if eid else scope_month
    body.append(f"- このファイルの記録: {len(group):,}件（{scope}の全件）")
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
        if md.get("group_by") == "entity_month":
            structure.append(f"- 記録（{entity_word}×月）: 1件ごとの内容。各ファイルには、その{entity_word}・その月の記録を全件載せています"
                             "（長い記録は「（続きn/m）」に分かれていますが、内容は削っていません）。")
        else:
            structure.append("- 記録（月ごと）: 1件ごとの内容。各ファイルには、その月の記録を全件載せています"
                             "（長い記録は「（続きn/m）」に分かれていますが、内容は削っていません）。")
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
    # 画面で「出さない」にした列は名前だけ書く（出さない列があること自体は分かるように）
    omitted = [col.display for col in spec.columns if col.md == "omit"]
    if omitted:
        columns.append(f"- 記録に出していない列: {'、'.join(omitted)}"
                       "（取り込み設定で「出さない」にした列です。元の値は管理用の正規化CSVにあります）")

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


__all__ = ["MdFile", "ai_point_lines", "parse_log_cell", "people_index_for", "record_block", "record_blocks",
           "record_title", "render_all", "render_dataset_card"]
