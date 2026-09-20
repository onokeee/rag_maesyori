"""一覧表用の標準キー辞書。

帳票の pattern/dictionary.py は、保存済み帳票の停止ラベルにも使われているので変更しない。
キー名は意味が同じものだけ帳票とそろえる（equipment_id, equipment_name, downtime, symptom, cause, alarm など）。
"""
from __future__ import annotations

import difflib
import re
import unicodedata
from dataclasses import dataclass, field


@dataclass(frozen=True)
class StdColumn:
    key: str
    display: str
    type: str  # code/string/text/date/datetime/time/number/enum/status
    role: str  # key/date/entity/entity_label/category/measure/text/log/person/attribute
    synonyms: tuple[str, ...] = field(default_factory=tuple)
    unit: str = ""
    md: str = "attribute"  # body/attribute/omit
    log_candidate: bool = False  # 追記ログの列になりやすい（AI整形の対象候補）


def _c(key, display, type_, role, synonyms, unit="", md="attribute", log_candidate=False) -> StdColumn:
    return StdColumn(key, display, type_, role, tuple(synonyms), unit, md, log_candidate)


STANDARD_COLUMNS: list[StdColumn] = [
    # 識別・日時
    _c("record_no", "管理No", "code", "key",
       ["管理No", "管理番号", "管理NO", "故障No", "トラブルNo", "トラブル番号", "台帳No", "台帳番号", "報告書No",
        "報告番号", "記録No", "記録番号", "整理番号", "受付番号", "伝票No", "No", "番号"]),
    _c("occurred_at", "発生日時", "datetime", "date",
       ["発生日時", "発生日", "故障発生日", "故障発生日時", "発生年月日", "日付", "年月日", "作業日", "実施日",
        "起票日", "発見日", "発見日時"]),
    _c("occurred_time", "発生時刻", "time", "attribute", ["発生時刻", "時刻", "発生時間"]),
    _c("reported_at", "報告日時", "datetime", "attribute", ["報告日時", "報告日", "連絡日時"]),
    _c("started_at", "対応開始日時", "datetime", "attribute", ["対応開始日時", "対応開始日", "着手日", "作業開始日時"]),
    _c("completed_at", "完了日", "datetime", "attribute", ["完了日", "完了日時", "復旧日時", "復旧日", "終了日", "終了日時"]),
    _c("due_date", "期限", "date", "attribute", ["期限", "対策期限", "完了予定日", "予定日"]),
    # 設備・場所
    _c("equipment_id", "設備番号", "code", "entity",
       ["設備番号", "設備No", "設備NO", "設備コード", "設備ID", "装置番号", "装置No", "装置ID", "装置コード",
        "機番", "号機", "対象設備", "設備"]),
    _c("equipment_name", "設備名", "string", "entity_label", ["設備名", "装置名", "設備名称", "装置名称", "機器名"]),
    _c("equipment_class", "設備分類", "enum", "attribute", ["設備分類", "設備分類コード", "設備区分", "装置区分", "設備種別"]),
    _c("line", "ライン", "string", "attribute", ["ライン", "ラインコード", "ライン名", "製造ライン"]),
    _c("process", "工程", "string", "attribute", ["工程", "工程名", "工程コード"]),
    _c("location", "設置場所", "string", "attribute", ["設置場所", "場所", "エリア"]),
    _c("part_location", "部位", "string", "attribute", ["部位", "部位コード", "故障部位", "ユニット"]),
    # 区分
    _c("record_type", "記録区分", "enum", "category", ["記録区分", "記録種別"]),
    _c("failure_category", "故障区分", "enum", "category",
       ["故障区分", "故障区分コード", "故障分類", "区分", "不具合区分", "作業区分", "保全区分", "トラブル区分"]),
    _c("severity", "重要度", "enum", "attribute", ["重要度", "重要度コード", "重大度", "影響度", "ランク"]),
    _c("shift", "シフト", "enum", "attribute", ["シフト", "シフトコード", "勤務帯", "直"]),
    _c("status", "状態", "status", "category", ["状態", "状態コード", "ステータス", "進捗状況", "対応状況"]),
    _c("recurrence", "再発", "enum", "attribute", ["再発", "再発フラグ", "再発有無"]),
    _c("judgement", "判定", "enum", "attribute", ["判定", "効果判定", "評価"]),
    # 文章
    _c("alarm", "アラーム", "code", "attribute", ["アラーム", "アラームコード", "エラーコード", "アラーム番号", "警報"]),
    _c("symptom", "現象", "text", "text",
       ["現象", "故障内容", "不具合内容", "不具合現象", "症状", "トラブル内容", "事象", "故障現象"], md="body"),
    _c("initial_action", "初期対応", "text", "text", ["初期対応", "応急処置", "暫定対策", "暫定処置", "初動"], md="body"),
    _c("investigation", "調査内容", "text", "text", ["調査内容", "調査結果", "調査"], md="body"),
    _c("cause", "原因", "text", "text", ["原因", "推定原因", "故障原因", "真因", "原因内容", "真因(なぜなぜ要約)"], md="body"),
    _c("cause_category", "原因区分", "enum", "category", ["原因区分", "原因区分コード", "原因分類"]),
    _c("why_analysis", "なぜなぜ分析", "text", "text", ["なぜなぜ分析", "なぜなぜ"], md="body"),
    _c("action", "処置", "text", "text",
       ["処置", "処置内容", "対策", "対策内容", "作業内容", "修理内容", "処置・対策", "作業内容・処置"],
       md="body", log_candidate=True),
    _c("response_log", "対応内容", "text", "log",
       ["対応内容", "対応履歴", "対応経過", "経過", "経緯", "対応記録", "経過記録", "進捗", "対応メモ"],
       md="body", log_candidate=True),
    _c("permanent_action", "恒久対策", "text", "text", ["恒久対策", "再発防止策", "再発防止対策", "恒久処置"], md="body"),
    _c("horizontal_deployment", "水平展開", "text", "text", ["水平展開"], md="body"),
    _c("result", "結果", "text", "text", ["結果", "効果確認", "処置結果"], md="body"),
    _c("remarks", "備考", "text", "attribute", ["備考", "特記事項", "備考・特記", "メモ", "コメント"]),
    # 部品・数量・金額・時間
    _c("parts_used", "使用部品", "text", "attribute", ["使用部品", "使用部品一覧"]),
    _c("part_name", "部品名", "string", "attribute", ["部品名", "交換部品", "交換部品名", "部品"]),
    _c("part_no", "品番", "code", "attribute", ["品番", "部品番号", "部品コード", "型番"]),
    _c("quantity", "数量", "number", "measure", ["数量", "個数", "使用数", "交換数"]),
    _c("unit_price", "単価", "number", "measure", ["単価"], unit="円"),
    _c("cost", "費用", "number", "measure", ["費用", "金額", "部品代", "修理費", "コスト"], unit="円"),
    _c("downtime", "停止時間", "number", "measure", ["停止時間", "ダウンタイム", "設備停止時間", "ライン停止時間"], unit="分"),
    _c("work_hours", "作業工数", "number", "measure", ["作業工数", "工数", "作業時間", "時間"], unit="h"),
    _c("scrap_qty", "廃棄枚数", "number", "measure", ["廃棄枚数", "廃棄数", "不良数"]),
    _c("attachments", "添付数", "number", "attribute", ["添付数", "添付"]),
    # 人・組織
    _c("worker", "担当者", "string", "person",
       ["担当者", "担当", "作業者", "実施者", "対応者", "担当者社員番号", "保全担当"]),
    _c("reporter", "報告者", "string", "person", ["報告者", "起票者", "報告者社員番号", "連絡者"]),
    _c("approver", "承認者", "string", "person", ["承認者", "確認者"]),
    _c("department", "部署", "string", "attribute", ["部署", "担当部署", "起票部署", "部門", "部門コード", "担当部門コード", "報告者部門コード"]),
    # 関連・管理用
    _c("related_no", "関連番号", "code", "attribute", ["関連管理番号", "関連トラブルNo", "関連No", "関連番号"]),
    _c("lot", "対象ロット", "string", "attribute", ["対象ロット", "ロット", "ロットNo"]),
    _c("internal_code", "内部コード", "code", "attribute", ["内部コード"]),
    _c("registered_at", "登録日時", "datetime", "attribute", ["登録日時", "作成日時"]),
    _c("updated_at", "更新日時", "datetime", "attribute", ["更新日時", "最終更新日時"]),
    _c("registered_by", "登録者", "string", "attribute", ["登録者ID", "登録者", "更新者ID", "更新者"]),
]

_STRIP_RE = re.compile(r"[\s・()（）\[\]【】「」『』_\-‐－/／:：.。、,，#＃*＊]")


def norm_header(text) -> str:
    """見出し比較用の正規化（NFKC・小文字・空白と区切り記号の除去）。"""
    s = unicodedata.normalize("NFKC", str(text or "")).lower()
    return _STRIP_RE.sub("", s)


_SYNONYMS: dict[str, StdColumn] = {}
for _col in STANDARD_COLUMNS:
    for _syn in (_col.display, *_col.synonyms):
        _SYNONYMS.setdefault(norm_header(_syn), _col)
_SIMILAR_KEYS = sorted(
    (k for k in _SYNONYMS if len(k) >= 2 and not (k.isascii() and len(k) < 4)),
    key=len, reverse=True,
)


def lookup_header(header: str) -> tuple[StdColumn, str] | None:
    """見出しから標準キーを探す。戻り値: (標準列, "dictionary" | "similar")"""
    from tables.detect import split_header_unit

    name, _unit = split_header_unit(header)
    candidates = [header, name]
    if "_" in str(header):  # 2段見出し「交換部品_品番」は下段でも照合する
        lower = str(header).rsplit("_", 1)[-1]
        candidates += [lower, split_header_unit(lower)[0]]
    for candidate in candidates:
        if norm_header(candidate) in _SYNONYMS:
            return _SYNONYMS[norm_header(candidate)], "dictionary"
    n = norm_header(name)
    if not n:
        return None
    for key in _SIMILAR_KEYS:  # 長い同義語から順に、見出しに含まれるものを探す
        if key in n:
            return _SYNONYMS[key], "similar"
    close = difflib.get_close_matches(n, list(_SYNONYMS), n=1, cutoff=0.8)
    if close:
        return _SYNONYMS[close[0]], "similar"
    return None
