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
                  ("報告番号", "報告No", "報告書番号", "報告書No", "修理番号", "修理No", "管理番号", "管理No", "連絡票No",
                   "連絡No", "連絡番号", "帳票No", "Report No", "No")),
    StandardField("subject", "件名", "string", ("件名", "タイトル", "表題")),
    StandardField("equipment_id", "設備番号", "string",
                  ("設備番号", "装置番号", "設備No", "装置No", "設備コード", "装置コード", "設備ID", "機番", "号機",
                   "Equipment ID")),
    StandardField("equipment_name", "設備名", "string", ("設備名", "装置名", "設備名称", "装置名称", "機器名")),
    StandardField("process", "工程", "string", ("工程", "発生工程", "工程名")),
    StandardField("location", "発生場所", "string", ("発生場所", "設置場所", "場所", "ライン", "生産ライン", "設置ライン")),
    StandardField("department", "部署", "string", ("部署", "所属", "担当部署", "部署名", "起票部署", "発行部署", "発行元",
                   "発信部署")),
    StandardField("occurred_date", "発生日", "date", ("発生日", "発生日時", "故障発生日", "発生年月日", "故障発生日時", "不具合発生日時",
                   "異常発生日時")),
    StandardField("completed_date", "修理完了日", "date", ("修理完了日", "完了日", "復旧日", "修理日", "復旧日時", "完了日時", "復旧完了日時",
                   "生産復帰日時")),
    StandardField("reporter", "報告者", "string", ("報告者", "作成者", "記入者", "担当者", "作業者", "起票者", "発信者", "連絡者",
                   "発行者")),
    StandardField("approver", "承認者", "string", ("承認者", "承認", "確認者", "完了承認者")),
    StandardField("work_hours", "作業時間", "number", ("作業時間", "修理時間", "対応時間")),
    StandardField("downtime", "停止時間", "number", ("停止時間", "ダウンタイム", "設備停止時間", "設備停止")),
    StandardField("symptom", "故障内容", "text", ("故障内容", "症状", "不具合内容", "異常内容", "現象", "故障状況", "異常の内容",
                   "不具合現象", "発生現象", "問題の記述", "問題の概要")),
    StandardField("severity", "重要度", "string", ("重要度", "重大度", "影響度", "ランク")),
    StandardField("alarm", "アラーム", "string", ("アラーム", "アラーム内容", "エラー", "エラーコード", "エラー内容", "アラームNo",
                   "アラームコード", "ALM No", "発生アラーム", "発報アラーム")),
    StandardField("investigation", "原因調査", "text", ("原因調査", "調査内容", "調査結果", "確認・調査結果", "調査・検証結果")),
    StandardField("cause", "原因", "text", ("原因", "故障原因", "推定原因", "発生原因", "発生要因", "真因", "根本原因", "原因(確定)")),
    StandardField("repair", "修理内容", "text", ("修理内容", "処置内容", "処置", "対応内容", "作業内容", "対策", "応急処置", "暫定対策",
                   "復旧処置")),
    StandardField("result", "修理結果", "text", ("修理結果", "結果", "確認結果", "処置結果", "復旧確認", "効果判定", "効果の確認結果")),
    StandardField("prevention", "再発防止策", "text", ("再発防止策", "再発防止対策", "恒久対策", "再発防止", "歯止め")),
    StandardField("parts", "使用部品", "text", ("使用部品", "交換部品", "部品")),
    StandardField("remarks", "備考", "text", ("備考", "特記事項", "その他")),
]

LOOKUP: dict[str, StandardField] = {
    normalize_label(synonym): sf for sf in STANDARD_FIELDS for synonym in sf.synonyms
}
DICTIONARY_NORMS: set[str] = set(LOOKUP)
BY_FIELD_NAME: dict[str, StandardField] = {sf.field_name: sf for sf in STANDARD_FIELDS}

# 設備番号と設備名を1つのセルにまとめて書く見出し（「対象設備：CMP-108　STI-CMP 8号機」）。
# 値が「番号＋名前」に分けられるときだけ、設備番号（番号の部分）と設備名（名前の部分）のラベルとして使う
COMBINED_EQUIPMENT_LABELS = ("対象設備", "使用設備", "対象装置", "使用装置", "設備", "装置")
COMBINED_EQUIPMENT_NORMS: set[str] = {normalize_label(label) for label in COMBINED_EQUIPMENT_LABELS}
# 分けた値のどちらを使うか（0: 番号, 1: 名前）
COMBINED_EQUIPMENT_PARTS = {"equipment_id": 0, "equipment_name": 1}

# 人名が入る項目（Markdown で既定では出さない）
PERSON_FIELD_NAMES = {"reporter", "approver", "worker", "person", "inspector", "creator", "checker", "author"}
_PERSON_LABEL_SUFFIXES = ("者", "担当", "氏名", "名前", "検印", "承認", "確認印")
# 押印欄の見出し（「確認」「作成」）。前方の語を含む「効果確認」「作成日」まで人名にしないよう、完全一致だけで見る
_PERSON_LABEL_EXACT = {"確認", "作成", "立会", "立ち会い", "審査", "署名", "サイン", "印"}


def is_person_field(field_name: str, display_name: str = "") -> bool:
    """報告者・承認者・担当者など人名の項目か。「担当部署」「設備名」「効果確認」は含めない。"""
    if field_name in PERSON_FIELD_NAMES:
        return True
    label = normalize_label(display_name)
    if not label:
        return False
    return label in _PERSON_LABEL_EXACT or is_person_label(display_name)


# 単位が書かれていないと意味が変わる数値項目（見出しの語 → 画面に出す説明）。
# 「件数」「回数」「人数」「枚数」のように数えるだけの項目は、単位が無いのが普通なので入れない
# （ここに無い項目は単位が空でも要確認にしない＝確認画面が警告だらけにならないようにする）。
AMBIGUOUS_UNIT_HINTS: tuple[tuple[tuple[str, ...], str], ...] = (
    (("時間", "工数", "所要", "ダウンタイム", "期間", "リードタイム", "タクト"), "分か時間かで意味が変わります"),
    (("金額", "費用", "コスト", "価格", "単価", "原価", "経費", "予算", "損失額", "修理費", "部品費"),
     "円か千円かで意味が変わります"),
    (("長さ", "寸法", "距離", "厚み", "厚さ", "板厚", "幅", "直径", "外径", "内径", "隙間", "クリアランス", "変位", "摩耗量"),
     "mmかmかで意味が変わります"),
    (("重量", "質量", "重さ"), "gかkgかで意味が変わります"),
    (("温度", "温度差"), "℃かKかで意味が変わります"),
    (("圧力", "真空度"), "MPaかkPaかで意味が変わります"),
    (("流量", "風量"), "L/minかm3/hかで意味が変わります"),
    (("電流", "電圧", "電力", "消費電力"), "AかmA（VかmV）かで意味が変わります"),
)
# 単位の付かない数え方・割合の項目（「不良件数」のように上の語を含んでいても、こちらが優先）
_COUNT_SUFFIXES = ("件数", "回数", "人数", "枚数", "個数", "台数", "本数", "点数", "数量", "員数", "率", "割合", "%")


def ambiguous_unit_hint(field_name: str, display_name: str = "") -> str:
    """単位が書かれていないと意味が変わる数値項目か。当てはまれば画面に出す説明、当てはまらなければ ""。

    単位の無い数値をすべて要確認にすると「件数」「回数」「人数」で確認画面が埋まるので、
    辞書が「単位で意味が変わる」と知っている項目だけを要確認にする。
    """
    standard = BY_FIELD_NAME.get(field_name)
    labels = [normalize_label(x) for x in (display_name, standard.display_name if standard else "")]
    labels = [x for x in labels if x]
    if not labels or any(x.endswith(_COUNT_SUFFIXES) for x in labels):
        return ""
    for words, hint in AMBIGUOUS_UNIT_HINTS:
        if any(normalize_label(w) in label for label in labels for w in words):
            return hint
    return ""


def is_person_label(display_name: str) -> bool:
    """見出しの語だけで人名の欄とわかるか（明細表の列見出し「担当」「氏名」用）。

    押印欄の「確認」「作成」は、明細表では判定記号の列でもありうるので含めない。
    """
    label = normalize_label(display_name)
    return bool(label) and label.endswith(_PERSON_LABEL_SUFFIXES)
