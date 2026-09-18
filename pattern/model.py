"""帳票の種類（旧: テンプレート／パターン）の定義。DBやFlaskに依存しない。"""
from __future__ import annotations

from dataclasses import dataclass, field

from excel.text import normalize_label

DATA_TYPES = {
    "string": "文字列",
    "text": "文章（複数行）",
    "date": "日付",
    "number": "数値",
    "table": "明細表（列見出しと行）",
}

DIRECTIONS = {
    "auto": "自動（右→下）",
    "right": "ラベルの右",
    "below": "ラベルの下",
    "same_cell": "同じセル（設備番号：EQ-001）",
}

# 画像の扱い。今は枚数だけ Markdown に書く（"vision" は未実装。DB に残っている古い値は読める）
IMAGE_PROCESSING = {
    "none": "読み取らない（枚数だけ Markdown に書きます）",
}

# Markdown に出すかどうか
RAG_OUTPUTS = {
    "show": "出す",
    "omit": "出さない",
}

# タイトル項目が未設定のときに使う「識別らしい項目」（この順で並べる）
DEFAULT_TITLE_KEYS = ("report_id", "equipment_id", "equipment_name", "occurred_date")

DEFAULT_MD_OPTIONS = {"omit_person_fields": True}


@dataclass
class SheetDef:
    sheet_name: str
    required: bool = True


@dataclass
class FieldDef:
    field_name: str
    display_name: str
    candidates: list[str] = field(default_factory=list)
    required: bool = False
    data_type: str = "string"
    direction: str = "auto"
    unit: str = ""
    rag_output: str = "show"
    # 明細表: 見本で見た列見出し。見出し（アンカー）の書き方が違う帳票でも、列見出しの並びが似た表を探すのに使う
    table_columns: list[str] = field(default_factory=list)

    def search_labels(self) -> list[str]:
        labels = [c.strip() for c in self.candidates if c.strip()] or [self.display_name]
        return list(dict.fromkeys(labels))

    def label_norms(self) -> set[str]:
        return {normalize_label(label) for label in self.search_labels()} - {""}


@dataclass
class PatternDef:
    name: str
    version: str = "v1"
    description: str = ""
    image_processing: str = "none"
    status: str = "draft"
    sheets: list[SheetDef] = field(default_factory=list)
    fields: list[FieldDef] = field(default_factory=list)
    id: int | None = None
    title_fields: list[str] = field(default_factory=list)
    md_options: dict = field(default_factory=lambda: dict(DEFAULT_MD_OPTIONS))
    version_no: int = 1

    @property
    def label(self) -> str:
        return f"{self.name} {self.version}"

    def label_norms(self) -> set[str]:
        norms: set[str] = set()
        for fd in self.fields:
            norms |= fd.label_norms()
        return norms

    def output_label_norms(self) -> set[str]:
        """この帳票の見出し語（候補ラベル・表示名・明細表の列見出し）の正規化形。値がこれに当たる行は Markdown に出さない。"""
        norms = self.label_norms()
        for fd in self.fields:
            norms |= {normalize_label(x) for x in (fd.display_name, *(fd.table_columns or []))}
        return norms - {""}
