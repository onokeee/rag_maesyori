"""一覧表の取り込み設定（TableSpec）。JSON で保存する（dataclass ⇔ dict、検証、spec_hash）。

- 列の定義（ColumnSpec）、追記ログ列の段（LogStageSpec）、AI の custom 段、集計の種類（SummarySpec）を持つ。
- 取り込み時の見出しとの照合（resolve_columns）と、画面の候補からの設定作成（spec_from_suggestions）もここに置く。
"""
from __future__ import annotations

import copy
import hashlib
import json
import re
from dataclasses import asdict, dataclass, field, fields

COLUMN_TYPES = ("code", "string", "text", "date", "datetime", "time", "number", "enum", "status")
COLUMN_ROLES = ("key", "date", "entity", "entity_label", "category", "measure", "text", "log", "person", "attribute")
MD_MODES = ("body", "attribute", "omit")
GROUP_BY = ("month", "entity_month")
SUMMARY_IDS = ("month", "entity_fiscal_year")
ROW_POLICIES = ("exclude_with_warning", "include")
CONTINUATION_POLICIES = ("merge_into_previous", "keep")
CUSTOM_OUTPUT_TYPES = ("text", "choice")

DEFAULT_NA_TOKENS = ["-", "－", "―", "‐", "N/A", "n/a", "NA", "#N/A", "該当なし"]
_KEY_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,63}$")
_METRIC_RE = re.compile(r"^(count|(sum|avg|max):([A-Za-z_][A-Za-z0-9_]*))$")


def _default_header() -> dict:
    return {"anchors": [], "rows": 1, "search_rows": 30}


def _default_data_end() -> dict:
    return {"blank_rows": 3, "stop_first_col": ["合計", "総計"], "stop_prefix": ["※", "注"]}


def _default_exclude() -> dict:
    return {"aggregate_keywords": ["小計", "計", "合計", "平均"],
            "hidden_rows": "exclude_with_warning", "strike_rows": "exclude_with_warning"}


def _default_record() -> dict:
    return {"key": [], "fallback_key": ["occurred_at", "equipment_id", "symptom:20"]}


def _default_period() -> dict:
    # 期間の置き換えはしない（取り込みごとにその内容だけで md を作る）。日付の列だけを持つ
    return {"date_column": "occurred_at"}


def _default_checks() -> dict:
    return {"type_error_rate": {"warn": 0.02, "block": 0.10}, "reconcile_tolerance": 0}


def _default_markdown() -> dict:
    # lightrag_hint: 新しい取り込み設定の既定はオン（ヒント無しだと LIGHTRAG_PARSER 未設定のサーバーで記録が途中で切られる）。
    # 保存済みの設定は JSON に値を持っているので、この既定では書き換わらない。
    # dedupe_timeline: 対応の時系列から、同じ記録の他の列と同じ文を省く（既定オン＝これまでの動き）。
    return {"file_prefix": "", "group_by": "month", "max_records_per_file": 300, "lightrag_hint": True,
            "dataset_card": True, "records": True, "dedupe_timeline": True,
            "summaries": [SummarySpec("month"), SummarySpec("entity_fiscal_year")],
            "title_columns": [], "omit_person": True}


@dataclass
class ColumnSpec:
    key: str
    display: str
    headers: list[str] = field(default_factory=list)
    type: str = "string"  # code/string/text/date/datetime/time/number/enum/status
    role: str = "attribute"  # key/date/entity/entity_label/category/measure/text/log/person/attribute
    unit: str = ""
    unit_conversions: dict = field(default_factory=dict)  # {"h": 60} = 1h を列の単位で60
    required: bool = False
    md: str = "attribute"  # body/attribute/omit
    fill_down_blank: bool = False
    normalize: list[str] = field(default_factory=lambda: ["nfkc"])  # nfkc / upper
    allowed: list[str] = field(default_factory=list)
    value_map: dict = field(default_factory=dict)
    description: str = ""


@dataclass
class LogStageSpec:
    column: str
    enabled_ai: bool = False
    context_columns: list[str] = field(default_factory=list)
    people: list[dict] = field(default_factory=list)  # {name, aliases, org}。アプリ内だけで使う
    groups: list[str] = field(default_factory=list)
    glossary: dict = field(default_factory=dict)
    entry_types: list[str] = field(default_factory=list)
    instruction: str = ""
    incident: bool = True
    run_if: dict = field(default_factory=dict)
    limits: dict = field(default_factory=dict)
    splitter: dict = field(default_factory=dict)  # logproc.SplitOptions.from_dict の形
    mask: list[str] = field(default_factory=lambda: ["phone", "email"])
    max_timeline_entries: int = 20  # 推定トークンが多いレコードで時系列を切る件数


@dataclass
class CustomStageSpec:
    id: str
    inputs: list[str] = field(default_factory=list)
    prompt: str = ""
    output_type: str = "text"  # text/choice
    choices: list[str] = field(default_factory=list)
    max_chars: int = 80
    fallback: str = "不明"
    target_key: str = ""
    quote_required: bool = True


@dataclass
class SummarySpec:
    id: str  # entity_fiscal_year / month
    metrics: list[str] = field(default_factory=lambda: ["count"])  # count, sum:<key>, avg:<key>, max:<key>
    top_n: int = 5


@dataclass
class TableSpec:
    name: str
    description: str = ""
    file_types: list[str] = field(default_factory=lambda: ["xlsx", "xlsm", "csv"])
    name_patterns: list[str] = field(default_factory=list)
    header: dict = field(default_factory=_default_header)
    data_end: dict = field(default_factory=_default_data_end)
    exclude: dict = field(default_factory=_default_exclude)
    continuation_rows: str = "merge_into_previous"
    na_tokens: list[str] = field(default_factory=lambda: list(DEFAULT_NA_TOKENS))
    fiscal_year_start_month: int = 4
    columns: list[ColumnSpec] = field(default_factory=list)
    record: dict = field(default_factory=_default_record)
    period: dict = field(default_factory=_default_period)
    log_stage: LogStageSpec | None = None
    custom_stages: list[CustomStageSpec] = field(default_factory=list)
    markdown: dict = field(default_factory=_default_markdown)
    checks: dict = field(default_factory=_default_checks)

    # ---- 参照の補助 ----
    def column(self, key: str) -> ColumnSpec | None:
        for col in self.columns:
            if col.key == key:
                return col
        return None

    def columns_with_role(self, role: str) -> list[ColumnSpec]:
        return [c for c in self.columns if c.role == role]

    def first_role(self, *roles: str) -> ColumnSpec | None:
        for role in roles:
            for col in self.columns:
                if col.role == role:
                    return col
        return None

    @property
    def date_key(self) -> str:
        key = (self.period or {}).get("date_column") or ""
        if key and self.column(key):
            return key
        col = self.first_role("date")
        return col.key if col else key

    @property
    def file_prefix(self) -> str:
        return str((self.markdown or {}).get("file_prefix") or self.name)

    def summaries(self) -> list[SummarySpec]:
        return [s if isinstance(s, SummarySpec) else _summary_from(s) for s in (self.markdown or {}).get("summaries", [])]


# ---- dict ⇔ dataclass ----------------------------------------------------------------

def _pick(cls, d: dict) -> dict:
    names = {f.name for f in fields(cls)}
    return {k: copy.deepcopy(v) for k, v in (d or {}).items() if k in names}


def _merged(default: dict, value) -> dict:
    out = copy.deepcopy(default)
    if isinstance(value, dict):
        out.update(copy.deepcopy(value))
    return out


def _as_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


def _summary_from(d) -> SummarySpec:
    if isinstance(d, SummarySpec):
        return d
    if isinstance(d, str):
        return SummarySpec(d)
    s = SummarySpec(**_pick(SummarySpec, d))
    s.metrics = [str(m) for m in _as_list(s.metrics)] or ["count"]
    s.top_n = int(s.top_n or 5)
    return s


def _str_field(obj, name: str) -> None:
    value = getattr(obj, name)
    if value is not None and not isinstance(value, str):
        setattr(obj, name, str(value))


def _custom_stage_from(d: dict) -> CustomStageSpec:
    stage = CustomStageSpec(**_pick(CustomStageSpec, d))
    for name in ("id", "output_type", "target_key"):
        _str_field(stage, name)
    stage.inputs = [str(v) for v in _as_list(stage.inputs)]
    stage.choices = [str(v) for v in _as_list(stage.choices)]
    return stage


def _column_from(d: dict) -> ColumnSpec:
    if isinstance(d, ColumnSpec):
        return d
    col = ColumnSpec(**_pick(ColumnSpec, d))
    # 手で書いた JSON の「"key": 1」なども文字列にして検証に回す（検証の途中で例外にしない）
    for name in ("key", "display", "type", "role", "md"):
        _str_field(col, name)
    col.headers = [str(h) for h in _as_list(col.headers) if str(h).strip()]
    col.normalize = [str(n) for n in _as_list(col.normalize)]
    col.allowed = [str(a) for a in _as_list(col.allowed)]
    col.value_map = dict(col.value_map or {})
    col.unit_conversions = dict(col.unit_conversions or {})
    col.required = bool(col.required)
    col.fill_down_blank = bool(col.fill_down_blank)
    col.unit = str(col.unit or "")
    col.description = str(col.description or "")
    return col


def spec_from_dict(d: dict) -> TableSpec:
    """dict（JSON）から TableSpec を作る。足りない項目は既定値。"""
    if not isinstance(d, dict):
        raise ValueError("取り込み設定は JSON のオブジェクトで指定してください")
    data = copy.deepcopy(d)
    spec = TableSpec(**_pick(TableSpec, {k: v for k, v in data.items()
                                         if k not in ("columns", "log_stage", "custom_stages", "markdown", "header",
                                                      "data_end", "exclude", "record", "period", "checks")}))
    spec.name = str(spec.name or "")
    spec.columns = [_column_from(c) for c in _as_list(data.get("columns"))]
    spec.header = _merged(_default_header(), data.get("header"))
    spec.data_end = _merged(_default_data_end(), data.get("data_end"))
    spec.exclude = _merged(_default_exclude(), data.get("exclude"))
    spec.record = _merged(_default_record(), data.get("record"))
    spec.period = _merged(_default_period(), data.get("period"))
    spec.checks = _merged(_default_checks(), data.get("checks"))
    markdown = _merged(_default_markdown(), data.get("markdown"))
    markdown["summaries"] = [_summary_from(s) for s in _as_list(markdown.get("summaries"))]
    spec.markdown = markdown
    log = data.get("log_stage")
    spec.log_stage = LogStageSpec(**_pick(LogStageSpec, log)) if isinstance(log, dict) and log.get("column") else None
    if spec.log_stage is not None:
        st = spec.log_stage
        # 手で書いた JSON の「"mask": "email"」「"phone,email"」も規則の並びとして読む（1文字ずつにしない）
        if isinstance(st.mask, str):
            st.mask = [m for m in re.split(r"[,、，\s]+", st.mask) if m]
        st.mask = [str(m) for m in _as_list(st.mask)]
        for name in ("context_columns", "groups", "entry_types"):
            setattr(st, name, [str(v) for v in _as_list(getattr(st, name))])
    spec.custom_stages = [_custom_stage_from(c) for c in _as_list(data.get("custom_stages")) if isinstance(c, dict)]
    spec.na_tokens = [str(t) for t in _as_list(spec.na_tokens)]
    spec.name_patterns = [str(p) for p in _as_list(spec.name_patterns)]
    spec.file_types = [str(t) for t in _as_list(spec.file_types)]
    spec.fiscal_year_start_month = int(spec.fiscal_year_start_month or 4)
    return spec


def spec_to_dict(spec: TableSpec) -> dict:
    return asdict(spec)


def spec_json(spec: TableSpec) -> str:
    """保存用の JSON（キー順固定）。"""
    return json.dumps(spec_to_dict(spec), ensure_ascii=False, sort_keys=True)


def spec_hash(spec: TableSpec) -> str:
    canonical = json.dumps(spec_to_dict(spec), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


# ---- 検証 ---------------------------------------------------------------------------

def validate_spec(spec: TableSpec) -> list[str]:
    """設定の問題点（日本語）。空なら保存・取り込みに使える。"""
    errors: list[str] = []
    if not str(spec.name or "").strip():
        errors.append("設定名を入力してください")
    if not spec.columns:
        errors.append("列を1つ以上設定してください")
    keys: set[str] = set()
    for col in spec.columns:
        label = col.display or col.key
        if not _KEY_RE.match(col.key or ""):
            errors.append(f"列「{label}」のキー「{col.key}」は半角英数字と _ で指定してください")
        elif col.key in keys:
            errors.append(f"キー「{col.key}」が重複しています")
        keys.add(col.key)
        if not str(col.display or "").strip():
            errors.append(f"列「{col.key}」の表示名を入力してください")
        if col.type not in COLUMN_TYPES:
            errors.append(f"列「{label}」の型「{col.type}」は使えません")
        if col.role not in COLUMN_ROLES:
            errors.append(f"列「{label}」の役割「{col.role}」は使えません")
        if col.md not in MD_MODES:
            errors.append(f"列「{label}」のmdでの扱い「{col.md}」は使えません")
        for unit, factor in (col.unit_conversions or {}).items():
            if not isinstance(factor, (int, float)) or isinstance(factor, bool) or factor <= 0:
                errors.append(f"列「{label}」の単位換算「{unit}」の倍率が正しくありません")
    if len([c for c in spec.columns if c.role == "entity"]) > 1:
        errors.append("役割「entity（設備など）」の列は1つだけにしてください")

    all_keys = set(keys)
    record = spec.record or {}
    for part in _as_list(record.get("key")):
        if str(part) not in all_keys:
            errors.append(f"記録キーの列「{part}」がありません")
    date_column = (spec.period or {}).get("date_column")
    col = spec.column(date_column) if date_column else None
    if col is not None and col.type not in ("date", "datetime"):
        errors.append(f"日付の列「{col.display}」の型を日付にしてください")
    if not 1 <= int(spec.fiscal_year_start_month or 0) <= 12:
        errors.append("年度の開始月は1〜12で指定してください")
    if spec.continuation_rows not in CONTINUATION_POLICIES:
        errors.append("継続行の扱いが正しくありません")
    header = spec.header if isinstance(spec.header, dict) else {}
    for name, label in (("rows", "見出しの行数（header.rows）"), ("search_rows", "見出しを探す行数（header.search_rows）")):
        value = header.get(name)
        if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= 1000:
            errors.append(f"{label}は1以上の整数で指定してください")
    for name in ("hidden_rows", "strike_rows"):
        if (spec.exclude or {}).get(name) not in ROW_POLICIES:
            errors.append("非表示行・取り消し線の行の扱いが正しくありません")
            break

    md = spec.markdown or {}
    if md.get("group_by") not in GROUP_BY:
        errors.append("記録ファイルのまとめ方は month / entity_month から選んでください")
    elif md.get("group_by") == "entity_month" and not spec.first_role("entity"):
        errors.append("設備×月でまとめるには、役割「entity（設備など）」の列が必要です")
    try:
        if int(md.get("max_records_per_file") or 0) < 1:
            errors.append("1ファイルの記録数の上限は1以上にしてください")
    except (TypeError, ValueError):
        errors.append("1ファイルの記録数の上限は数値で指定してください")
    for key in _as_list(md.get("title_columns")):
        base = str(key).split(":")[0]
        if base not in all_keys:
            errors.append(f"見出しに使う列「{base}」がありません")
    for s in spec.summaries():
        if s.id not in SUMMARY_IDS:
            errors.append(f"集計の種類「{s.id}」は使えません")
        for metric in s.metrics:
            m = _METRIC_RE.match(metric)
            if not m:
                errors.append(f"集計の指標「{metric}」は count / sum:列 / avg:列 / max:列 で指定してください")
            elif m.group(3):
                target = m.group(3)
                col = spec.column(target)
                if col is None or col.type != "number":
                    errors.append(f"集計の指標「{metric}」の列が数値の列ではありません")

    if spec.log_stage is not None:
        col = spec.column(spec.log_stage.column)
        if col is None:
            errors.append(f"AI整形の対象列「{spec.log_stage.column}」がありません")
        for key in spec.log_stage.context_columns:
            if key not in keys:
                errors.append(f"AI整形に添える列「{key}」がありません")
        errors.extend(_log_stage_errors(spec.log_stage))
    try:
        tolerance = float((spec.checks or {}).get("reconcile_tolerance") or 0)
        if not 0 <= tolerance < float("inf"):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("突き合わせの許容差（checks.reconcile_tolerance）は0以上の数値で指定してください")
    stage_ids: set[str] = set()
    for stage in spec.custom_stages:
        if not _KEY_RE.match(stage.id or "") or stage.id in stage_ids:
            errors.append(f"AIの追加処理のID「{stage.id}」が正しくないか重複しています")
        stage_ids.add(stage.id)
        if stage.output_type not in CUSTOM_OUTPUT_TYPES:
            errors.append(f"AIの追加処理「{stage.id}」の出力の種類が正しくありません")
        if stage.output_type == "choice" and not stage.choices:
            errors.append(f"AIの追加処理「{stage.id}」の選択肢を入力してください")
        for key in stage.inputs:
            if key not in keys:
                errors.append(f"AIの追加処理「{stage.id}」の入力列「{key}」がありません")
    rates = (spec.checks or {}).get("type_error_rate") or {}
    if not isinstance(rates, dict):
        errors.append("型エラーの割合の上限（checks.type_error_rate）は warn と block を持つ形で指定してください")
        rates = {}
    try:
        if not 0 <= float(rates.get("warn", 0.02)) <= float(rates.get("block", 0.10)) <= 1:
            errors.append("型エラーの割合の上限は 0〜1 で、警告 ≦ 確定を止める にしてください")
    except (TypeError, ValueError):
        errors.append("型エラーの割合の上限は数値で指定してください")
    return errors


# 区切りの正規表現（JSON で取り込んだ設定だけが持つ）の上限。画面からは設定しない
_MAX_SPLIT_PATTERNS = 20
_MAX_PATTERN_CHARS = 200
# 量指定子を含むグループにさらに量指定子が付く形（(.+)+ など）と、選択（|）を含むグループに量指定子が付く形
# （(?:\d|\d)* など。選択肢が重なると同じく極端に遅くなる）。どちらも受け付けない
_NESTED_QUANTIFIER_RE = re.compile(r"\([^)]*[+*|][^)]*\)[+*{]")


def _log_stage_errors(stage: LogStageSpec) -> list[str]:
    """AI整形の設定（JSON で取り込んだときだけ画面に出ない項目）の問題点。"""
    from logproc.mask import normalize_rules

    errors: list[str] = []
    bad = [m for m in stage.mask if not normalize_rules([m])]
    if bad:
        errors.append(f"AI整形の伏せ字の規則「{'、'.join(bad)}」は使えません（phone / email / person / amount か"
                      "電話番号 / メール / 人名 / 金額）")
    if not isinstance(stage.people, list) or not all(
            isinstance(p, dict) and isinstance(p.get("name"), str) and p["name"].strip()
            and isinstance(p.get("aliases", []), list) for p in stage.people):
        errors.append("AI整形の人名一覧（people）は、name（名前）を持つ項目の並びで指定してください")
    for name, label in (("glossary", "用語集"), ("splitter", "区切り"), ("limits", "上限"), ("run_if", "実行条件")):
        if not isinstance(getattr(stage, name), dict):
            errors.append(f"AI整形の{label}（{name}）の書き方が正しくありません")
    try:
        n = stage.max_timeline_entries
        if isinstance(n, bool) or int(n) != n or not 1 <= int(n) <= 1000:
            raise ValueError
    except (TypeError, ValueError):
        errors.append("時系列の最大件数（max_timeline_entries）は1〜1000の整数で指定してください")
    splitter = stage.splitter if isinstance(stage.splitter, dict) else {}
    for key in ("extra_anchors", "not_date_patterns"):
        patterns = splitter.get(key)
        if patterns is None:
            continue
        if not isinstance(patterns, list) or len(patterns) > _MAX_SPLIT_PATTERNS:
            errors.append(f"AI整形の区切りの正規表現（{key}）は{_MAX_SPLIT_PATTERNS}個までの並びで指定してください")
            continue
        for pat in patterns:
            text = str(pat)[:40]
            if not isinstance(pat, str) or len(pat) > _MAX_PATTERN_CHARS:
                errors.append(f"AI整形の区切りの正規表現「{text}」は{_MAX_PATTERN_CHARS}文字以内の文字列にしてください")
                continue
            try:
                re.compile(pat)
            except re.error:
                errors.append(f"AI整形の区切りの正規表現「{text}」が正しくありません")
                continue
            if _NESTED_QUANTIFIER_RE.search(pat):
                errors.append(f"AI整形の区切りの正規表現「{text}」は処理が極端に遅くなる形（(…+)+ など）です")
    for key in ("sentence_split_min_chars", "order_tolerance_days"):
        if splitter.get(key) is not None:
            try:
                int(splitter[key])
            except (TypeError, ValueError):
                errors.append(f"AI整形の区切りの {key} は整数で指定してください")
    # 上限・実行条件の数値（aiproc/runner.py が int() で読む。数値でないと分割プレビュー・AI整形が止まる）
    limits = stage.limits if isinstance(stage.limits, dict) else {}
    for key in ("max_segments", "max_input_tokens"):
        if key in limits and not _is_int_at_least(limits[key], 1):
            errors.append(f"AI整形の上限（limits）の {key} は1以上の整数で指定してください")
    if isinstance(stage.run_if, dict) and not _run_if_ok(stage.run_if, 0):
        errors.append("AI整形の実行条件（run_if）の min_chars・min_segments は0以上の整数で、"
                      "any・all は条件の並びで指定してください")
    return errors


def _is_int_at_least(value, low: int) -> bool:
    if isinstance(value, bool):
        return False
    try:
        n = int(value)
    except (TypeError, ValueError):
        return False
    return n >= low and (not isinstance(value, float) or value == n)


def _run_if_ok(rule, depth: int) -> bool:
    """実行条件（aiproc.runner._run_if の形）の数値と入れ子を確かめる。"""
    if not isinstance(rule, dict):
        return True   # 条件でないものは runner が「条件なし」として扱う
    if depth > 10:
        return False
    for key in ("any", "all"):
        if key in rule:
            items = rule[key]
            if items is not None and not isinstance(items, list):
                return False
            if not all(_run_if_ok(r, depth + 1) for r in items or []):
                return False
    return all(_is_int_at_least(rule[k], 0) for k in ("min_chars", "min_segments") if k in rule)


# ---- 見出しとの照合 --------------------------------------------------------------------

@dataclass
class ColumnResolution:
    positions: dict[str, int]  # 列キー → 見出しの位置（0始まり）
    missing_required: list[str]  # 見つからない必須列の表示名
    missing: list[str]  # 見つからない列の表示名（必須以外も含む）
    unused_headers: list[str]  # どの列にも対応しない見出し


def resolve_columns(spec: TableSpec, headers: list[str]) -> ColumnResolution:
    """設定の列を、読み取った見出しの位置に対応づける。"""
    from tables.detect import split_header_unit
    from tables.dictionary import norm_header

    def norms_of(header: str) -> set[str]:
        name = split_header_unit(header)[0]
        out = {norm_header(header), norm_header(name)}
        if "_" in header:
            lower = header.rsplit("_", 1)[-1]
            out |= {norm_header(lower), norm_header(split_header_unit(lower)[0])}
        return {n for n in out if n}

    header_norms = [norms_of(h) for h in headers]
    exact = [norm_header(h) for h in headers]
    positions: dict[str, int] = {}
    used: set[int] = set()

    # 完全一致を先に、その後に単位・2段見出しの下段で一致したものを割り当てる
    for strict in (True, False):
        for col in spec.columns:
            if col.key in positions:
                continue
            candidates = [c for c in (list(col.headers) + [col.display]) if c]
            cand_exact = [norm_header(c) for c in candidates]
            cand_loose = {norm_header(split_header_unit(c)[0]) for c in candidates} | set(cand_exact)
            cand_loose.discard("")
            for pos in range(len(headers)):
                if pos in used:
                    continue
                hit = exact[pos] in cand_exact if strict else bool(header_norms[pos] & cand_loose)
                if hit:
                    positions[col.key] = pos
                    used.add(pos)
                    break
    missing = [c.display for c in spec.columns if c.key not in positions]
    missing_required = [c.display for c in spec.columns if c.required and c.key not in positions]
    unused = [h for pos, h in enumerate(headers)
              if pos not in used and not (h.startswith("列") and h[1:].isdigit())]
    return ColumnResolution(positions, missing_required, missing, unused)


# ---- 候補からの設定作成 -----------------------------------------------------------------

def _get(obj, name, default=None):
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def spec_from_suggestions(name: str, layout, suggestions, options: dict | None = None) -> TableSpec:
    """見出しの判定結果（LayoutGuess）と列の候補（ColumnSuggestion）から取り込み設定を作る。

    options: description, name_patterns, file_types, group_by, fiscal_year_start_month, max_records_per_file
    """
    options = dict(options or {})
    spec = TableSpec(name=name)  # ファイル名の先頭は空（＝設定名。TableSpec.file_prefix）
    spec.description = str(options.get("description") or "")
    if options.get("name_patterns"):
        spec.name_patterns = [str(p) for p in options["name_patterns"]]
    if options.get("file_types"):
        spec.file_types = [str(t) for t in options["file_types"]]
    if options.get("fiscal_year_start_month"):
        spec.fiscal_year_start_month = int(options["fiscal_year_start_month"])
    header_rows = list(_get(layout, "header_rows", []) or [])
    spec.header["rows"] = max(1, len(header_rows))

    columns: list[ColumnSpec] = []
    keys: set[str] = set()
    for s in suggestions:
        header = _get(s, "header", "")
        key = _get(s, "key") or f"col{int(_get(s, 'index', len(columns))) + 1}"
        base, n = key, 2
        while key in keys:
            key = f"{base}_{n}"
            n += 1
        keys.add(key)
        type_ = _get(s, "type", "string") or "string"
        if type_ not in COLUMN_TYPES:
            type_ = "string"
        role = _get(s, "role", "attribute") or "attribute"
        if role not in COLUMN_ROLES:
            role = "attribute"
        md = _get(s, "md", "attribute") or "attribute"
        columns.append(ColumnSpec(
            key=key, display=_get(s, "display", "") or header, headers=[header] if header else [],
            type=type_, role=role, unit=_get(s, "unit", "") or "", md=md if md in MD_MODES else "attribute",
            fill_down_blank=bool(_get(s, "fill_down_blank", False)),
            normalize=["nfkc", "upper"] if (role == "entity" and type_ == "code") else ["nfkc"],
        ))
    # 役割の重複を直す（entity・entity_label・key は1列だけ）
    for role in ("entity", "entity_label", "key"):
        for extra in [c for c in columns if c.role == role][1:]:
            extra.role = "attribute"
    spec.columns = columns

    entity = spec.first_role("entity")
    date_col = spec.first_role("date")
    key_col = spec.first_role("key")
    text_col = spec.first_role("text")
    if key_col is not None:
        key_col.required = True
    if date_col is not None:
        date_col.required = True
    spec.record = {
        "key": [key_col.key] if key_col else [],
        "fallback_key": [c for c in (
            date_col.key if date_col else None,
            entity.key if entity else None,
            f"{text_col.key}:20" if text_col else None,
        ) if c],
    }
    spec.period = {"date_column": date_col.key if date_col else ""}
    log_col = next((c for c in columns if c.role == "log"), None)
    if log_col is not None:
        spec.log_stage = LogStageSpec(column=log_col.key, context_columns=[
            c.key for c in (spec.first_role("entity_label"), entity, spec.columns_with_role("text")[0]
                            if spec.columns_with_role("text") else None) if c is not None])
    if options.get("group_by") in GROUP_BY:
        spec.markdown["group_by"] = options["group_by"]
    if options.get("max_records_per_file"):
        spec.markdown["max_records_per_file"] = int(options["max_records_per_file"])
    return spec
