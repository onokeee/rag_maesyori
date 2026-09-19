"""追記ログ（対応内容など）のルール処理で使うデータ構造。"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field

# 「1.5mm」「2.5A」のような寸法・電気量は日付にしない。ただし単位の後ろに英数字・「-」・カタカナが続くときは
# 「4.3 AGV」「4.3 ALM-2031」「4.3 Aライン」「4.3 Vベルト」のように日付＋設備名などなので、単位とみなさない
DEFAULT_NOT_DATE_PATTERNS = [r"\d+\.\d+\s*(?:mm|MPa|V|A)(?![A-Za-z0-9\-\u30A0-\u30FF])", r"納期\s*\d+/\d+"]


@dataclass
class WhenInfo:
    text: str                      # 原文の日時表現（例「4/1 10:00」「翌週」）。継承時は空か時刻のみ
    date: str | None               # ISO日付（範囲なら開始日）
    time: str | None               # "HH:MM"
    date_to: str | None            # 範囲の終了日（翌週など）
    shift: str | None              # 夜勤/日勤/2直/夕方 など
    estimated: bool
    note: str = ""                 # 要確認に出す説明（「原文「翌週」から推定した（基準は…）」など）
    how: str = ""                  # explicit/year_inferred/relative/inherited/base/unresolved/year_unknown/none


@dataclass
class AuthorInfo:
    raw: str                       # 原文の表記（「K.T」「保全G 高橋」など）
    name: str | None               # 特定できた名前。イニシャル未登録などは None
    estimated: bool
    note: str = ""


@dataclass
class Segment:
    id: str                        # s1..（原文の並び順）
    raw: str                       # 原文（クリーニング後テキストの start:end）
    body: str                      # 日時・記入者の部分を除いた原文
    start: int
    end: int
    when: WhenInfo | None
    author: AuthorInfo | None
    identifiers: list[str] = field(default_factory=list)
    quantities: list[str] = field(default_factory=list)
    plans: list[str] = field(default_factory=list)
    marks: list[str] = field(default_factory=list)   # email/header_cell/bullet/checklist/note/sentence_split/reference/correction
    label: str = ""                # 見出し型セルの見出し語（【現象】→「現象」）


@dataclass
class LogParse:
    segments: list[Segment]
    order: str                     # asc/desc/unknown
    kind: str                      # log/header_cell/single/empty
    warnings: list[str] = field(default_factory=list)
    text: str = ""                 # 分割に使ったテキスト（_x000D_ と CRLF を除いたもの。start/end はこの位置）

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class SplitOptions:
    order: str = "auto"                    # auto/asc/desc
    sentence_split_min_chars: int = 120    # 目印のない長いセグメントを「。」で分ける長さ
    header_cells: str = "detect"           # detect/off
    extra_anchors: list[str] = field(default_factory=list)       # 行頭の区切りを正規表現で追加
    not_date_patterns: list[str] = field(default_factory=lambda: list(DEFAULT_NOT_DATE_PATTERNS))
    time_only_lines: str = "separate"      # separate（別セグメント、日付は直前から）/ join（直前につなぐ）
    order_tolerance_days: int = 60         # この日数以内の逆行は年を変えず「前後している」とする

    @classmethod
    def from_dict(cls, d: dict | None) -> "SplitOptions":
        d = d or {}
        opts = cls()
        for key in ("order", "header_cells", "time_only_lines"):
            if d.get(key):
                setattr(opts, key, str(d[key]))
        split = d.get("sentence_split")
        if isinstance(split, dict) and split.get("min_chars"):
            opts.sentence_split_min_chars = int(split["min_chars"])
        if d.get("sentence_split_min_chars"):
            opts.sentence_split_min_chars = int(d["sentence_split_min_chars"])
        if d.get("extra_anchors"):
            opts.extra_anchors = [str(x) for x in d["extra_anchors"]]
        if d.get("not_date_patterns") is not None:
            opts.not_date_patterns = [str(x) for x in d["not_date_patterns"]]
        if d.get("order_tolerance_days") is not None:
            opts.order_tolerance_days = int(d["order_tolerance_days"])
        return opts
