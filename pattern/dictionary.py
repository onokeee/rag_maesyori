"""帳票でよく使われる項目名の辞書。

サンプルExcelからテンプレートを作るときの候補提示と、
値の探索時に「隣のセルが別のラベルかどうか」の判定に使う。
"""
from __future__ import annotations

from dataclasses import dataclass

from excel.text import normalize_label


@dataclass(frozen=True)
class StandardField:
    field_name: str
    display_name: str
    data_type: str
    synonyms: tuple[str, ...]


STANDARD_FIELDS = [
    StandardField("report_id", "報告番号", "string",
                  ("報告番号", "報告No", "報告書番号", "報告書No", "修理番号", "修理No", "管理番号", "Report No")),
    StandardField("subject", "件名", "string", ("件名", "タイトル", "表題")),
    StandardField("equipment_id", "設備番号", "string",
                  ("設備番号", "装置番号", "設備No", "装置No", "設備コード", "装置コード", "Equipment ID")),
    StandardField("equipment_name", "設備名", "string", ("設備名", "装置名", "設備名称", "装置名称")),
    StandardField("process", "工程", "string", ("工程", "発生工程", "工程名")),
    StandardField("location", "発生場所", "string", ("発生場所", "設置場所", "場所", "ライン")),
    StandardField("department", "部署", "string", ("部署", "所属", "担当部署", "部署名")),
    StandardField("occurred_date", "発生日", "date", ("発生日", "発生日時", "故障発生日", "発生年月日")),
    StandardField("completed_date", "修理完了日", "date", ("修理完了日", "完了日", "復旧日", "修理日")),
    StandardField("reporter", "報告者", "string", ("報告者", "作成者", "記入者", "担当者", "作業者")),
    StandardField("approver", "承認者", "string", ("承認者", "承認", "確認者")),
    StandardField("work_hours", "作業時間", "number", ("作業時間", "修理時間", "対応時間")),
    StandardField("downtime", "停止時間", "number", ("停止時間", "ダウンタイム", "設備停止時間")),
    StandardField("symptom", "故障内容", "text", ("故障内容", "症状", "不具合内容", "異常内容", "現象", "故障状況")),
    StandardField("alarm", "アラーム", "string", ("アラーム", "アラーム内容", "エラー", "エラーコード", "エラー内容")),
    StandardField("investigation", "原因調査", "text", ("原因調査", "調査内容", "調査結果")),
    StandardField("cause", "原因", "text", ("原因", "故障原因", "推定原因", "発生原因")),
    StandardField("repair", "修理内容", "text", ("修理内容", "処置内容", "処置", "対応内容", "作業内容", "対策")),
    StandardField("result", "修理結果", "text", ("修理結果", "結果", "確認結果", "処置結果")),
    StandardField("prevention", "再発防止策", "text", ("再発防止策", "再発防止対策", "恒久対策")),
    StandardField("parts", "使用部品", "text", ("使用部品", "交換部品", "部品")),
    StandardField("remarks", "備考", "text", ("備考", "特記事項", "その他")),
]

LOOKUP: dict[str, StandardField] = {
    normalize_label(synonym): sf for sf in STANDARD_FIELDS for synonym in sf.synonyms
}
DICTIONARY_NORMS: set[str] = set(LOOKUP)

# 人名が入る項目（Markdown で既定では出さない）
PERSON_FIELD_NAMES = {"reporter", "approver", "worker", "person", "inspector", "creator", "checker", "author"}
_PERSON_LABEL_SUFFIXES = ("者", "担当", "氏名", "名前", "検印", "承認", "確認印")


def is_person_field(field_name: str, display_name: str = "") -> bool:
    """報告者・承認者・担当者など人名の項目か。「担当部署」「設備名」は含めない。"""
    if field_name in PERSON_FIELD_NAMES:
        return True
    label = normalize_label(display_name)
    return bool(label) and label.endswith(_PERSON_LABEL_SUFFIXES)
