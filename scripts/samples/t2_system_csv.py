"""T2 設備故障履歴（保全管理システムからのCSVエクスポート）を生成する。

    python -m scripts.samples.t2_system_csv

出力（samples/tables/ 配下）:
- T2_設備故障履歴_システム出力.csv         : CP932 / CRLF。先頭4行が出力条件、最終行が「合計件数」トレーラ
- T2_コード表.csv                          : CP932 / CRLF。本CSVで使うコード値の定義（人が読んで復号する用）
- T2_設備故障履歴_システム出力_README.md   : 構造・文字コード・件数・意図的な不規則性の説明（UTF-8）

内容は domain.standard_incidents()（設備故障）と domain.minor_stops()（チョコ停）を
発生日時順にマージしたもの。名称ではなくコード（設備コード・故障区分コード・重要度コード・
社員番号・部門コード など）で出力する、いかにも基幹システムらしい形にしている。
管理番号・設備コード・社員番号は domain と同一なので、帳票サンプルと突き合わせできる。

意図的に入れている「現場のシステム出力らしさ」:
- ヘッダ前の出力条件4行と、末尾の合計件数行（表として素直に読めない）
- 処置内容のセル内改行（LF）、カンマを含むクォート付き項目、"" エスケープ
- チョコ停行では故障区分・重要度・原因などが空欄
- 旧システムからの移行データ（2023年4〜9月に発生し移行時点で完了済みの故障）は自由記述が全角英数、
  対応開始日時が空欄、費用が「634,600」形式の文字列、改行が「／」に置換
"""
from __future__ import annotations

import collections
import csv
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

from . import domain as D

# ---------------------------------------------------------------------------
# 基本設定
# ---------------------------------------------------------------------------
TABLE_DIR = "tables"
CSV_NAME = "T2_設備故障履歴_システム出力.csv"
CODE_NAME = "T2_コード表.csv"
README_NAME = "T2_設備故障履歴_システム出力_README.md"

EXPORT_AT = datetime(2026, 9, 14, 18, 5, 33)          # 出力日時（固定。完了日時・更新日時の最大値より後）
EXPORT_USER = next(p for p in D.people() if p.employee_id == "M10544")   # 保全1係 係長が出力した想定
PERIOD_FROM = datetime(2023, 4, 1, 0, 0, 0)
PERIOD_TO = datetime(2026, 8, 31, 23, 59, 59)
LEGACY_BEFORE = datetime(2023, 10, 1)                   # これより前に発生した故障は旧システムからの移行データ
MIGRATED_AT = datetime(2023, 10, 1, 2, 10, 0)           # 移行バッチの実行日時
SYS_MES = "SYS-MES"                                     # MES連携（チョコ停の自動登録）
SYS_MIG = "SYS-MIG"                                     # 旧システム移行バッチ

DT_FMT = "%Y/%m/%d %H:%M:%S"

HEADER = [
    "管理番号", "記録区分", "設備コード", "設備分類コード", "ラインコード",
    "発生日時", "報告日時", "対応開始日時", "完了日時", "シフトコード", "発見区分コード",
    "報告者社員番号", "報告者部門コード", "担当者社員番号", "担当部門コード",
    "故障区分コード", "部位コード", "重要度コード", "状態コード", "アラームコード",
    "現象", "初期対応", "調査内容", "原因", "原因区分コード", "なぜなぜ分析", "処置内容", "使用部品",
    "停止時間(分)", "作業工数(H)", "費用(円)", "対象ロット", "廃棄枚数", "復旧区分コード",
    "再発フラグ", "関連管理番号", "結果", "再発防止策", "水平展開", "備考", "添付数",
    "登録日時", "登録者ID", "更新日時", "更新者ID",
]

# ---------------------------------------------------------------------------
# コード定義（固定の並び。domain 側に想定外の値が出たら末尾に採番して追加する）
# ---------------------------------------------------------------------------
RECORD_KIND = {"設備故障": "1", "チョコ停": "2"}
FAILURE_CAT = {"機械": "F01", "電気": "F02", "制御": "F03", "ソフト": "F04", "ユーティリティ": "F05",
               "人為": "F06", "品質": "F07", "外部要因": "F08"}
SEVERITY = {"重大": "1", "大": "2", "中": "3", "小": "4"}
STATUS = {"対応中": "1", "保留": "2", "経過観察": "3", "完了": "9"}
SHIFT = {"日勤": "1", "夜勤": "2", "休日": "3"}
RECOVERY = {"自動復帰": "1", "製造復旧": "2", "保全復旧": "3", "メーカー復旧": "4"}
DETECTED = {
    "装置アラーム": "H01", "オペレーター": "H02", "オペレーター申告": "H03", "保全巡回": "H04", "中央監視": "H05",
    "FDC監視": "H06", "SPC異常": "H10", "SPC異常(CD)": "H11", "SPC異常(膜厚)": "H12", "SPC異常(欠陥数)": "H13",
    "SPC異常(パーティクル)": "H14", "SPC異常(シート抵抗)": "H15", "SPC異常(重ね合わせ)": "H16",
    "インライン欠陥検査": "H20", "インライン膜厚測定": "H21", "マクロ検査": "H22", "パーティクルQC": "H23",
    "QC(E/R均一性)": "H24", "後工程からの連絡": "H30", "使用先装置からの連絡": "H31", "技術者の気付き": "H40",
}
CAUSE_CAT = {"摩耗": "C01", "劣化": "C02", "汚れ": "C03", "異物": "C04", "締結緩み": "C05", "断線": "C06",
             "調整不良": "C07", "設定ミス": "C08", "作業ミス": "C09", "施工不良": "C10", "ソフト不具合": "C11",
             "設計起因": "C12", "能力不足": "C13", "前工程起因": "C14", "外部要因": "C15", "不明": "C99"}
EQ_CAT = {"CMP装置": "K01", "CVD装置": "K02", "エッチング装置": "K03", "露光装置": "K04", "洗浄装置": "K05",
          "イオン注入装置": "K06", "検査装置": "K07", "搬送系": "K08", "ユーティリティ": "K09"}
LINE = {"L1": "LN01", "L2": "LN02", "L3": "LN03", "L4": "LN04", "L5": "LN05", "L6": "LN06",
        "L1-L3": "LN13", "L4-L6": "LN46", "Fab1共通": "UT01", "Fab2共通": "UT02", "全Fab共通": "UT00"}
LINE_NOTE = {"LN13": "Fab1 全ライン（搬送系）", "LN46": "Fab2 全ライン（搬送系）", "UT01": "Fab1 ユーティリティ",
             "UT02": "Fab2 ユーティリティ", "UT00": "全Fab共通ユーティリティ"}
# 部門コード（部・課・係の階層を5桁で表す。設備メーカーは取引先扱いで9xxxx）
SECTION = {
    ("製造部", "設備保全課"): "21000", ("製造部", "設備保全課 保全1係"): "21010", ("製造部", "設備保全課 保全2係"): "21020",
    ("製造部", "設備保全課 施設係"): "21030", ("製造部", "製造課 A班"): "22010", ("製造部", "製造課 B班"): "22020",
    ("製造部", "製造課 C班"): "22030", ("製造部", "製造課 D班"): "22040", ("製造部", "生産技術課"): "23000",
    ("品質保証部", "品質保証課"): "31000",
}


def _code_of(table: dict, value: str, prefix: str, width: int) -> str:
    """table に無い値は prefix + 連番で追加採番する（domain 側の拡張でも落ちないように）。"""
    if value not in table:
        n = len(table) + 1
        code = f"{prefix}{n:0{width}d}"
        while code in table.values():
            n += 1
            code = f"{prefix}{n:0{width}d}"
        table[value] = code
    return table[value]


def _section_code(p: D.Person) -> str:
    key = (p.department, p.section)
    if key not in SECTION:
        if p.department == "設備メーカー":
            SECTION[key] = f"9{sum(1 for k in SECTION if k[0] == '設備メーカー') + 1:02d}00"
        else:
            SECTION[key] = f"8{len(SECTION):02d}00"
    return SECTION[key]


def _subsystem_codes(incidents: list, stops: list) -> dict:
    """部位（サブシステム）コード。3分類以上で使う部位は共通 B0xx、それ以外は最初の設備分類の B{分類}xx。"""
    cats_of: dict[str, set] = collections.defaultdict(set)
    for inc in incidents:
        cats_of[inc.subsystem].add(inc.equipment.category)
    for eq in D.equipment_master():
        for a in D.alarm_codes_for(eq):
            cats_of[a.subsystem].add(eq.category)
    for ms in stops:
        if ms.alarm:
            cats_of[ms.alarm.subsystem].add(ms.equipment.category)
    for part in D.parts_catalog():
        cats_of[part.subsystem].update(part.applicable_categories)
    cat_order = list(EQ_CAT)
    groups: dict[int, list] = collections.defaultdict(list)
    for sub, cats in cats_of.items():
        if len(cats) >= 3:
            groups[0].append(sub)
        else:
            groups[min(cat_order.index(c) for c in cats) + 1].append(sub)
    out = {}
    for g in sorted(groups):
        for i, sub in enumerate(sorted(groups[g]), 1):
            out[sub] = f"B{g}{i:02d}"
    return out


# ---------------------------------------------------------------------------
# 文字列ヘルパー
# ---------------------------------------------------------------------------
# CP932 に無い文字の置換（Å は CP932 では U+212B に割り当たっている）。見た目が同じ文字が多いのでエスケープで書く
_CP932_FIX = {
    "\u00c5": "\u212b",   # オングストローム（ラテン文字Å -> 単位記号）
    "\u301c": "\uff5e",   # 波ダッシュ -> 全角チルダ
    "\u2212": "\uff0d",   # 数学マイナス -> 全角ハイフンマイナス
    "\u2014": "\u2015",   # EMダッシュ -> 水平線
    "\u00a0": " ",        # NBSP -> 半角スペース
    "\u2016": "\u2225",   # 双柱 -> 平行記号
    "\u00a2": "\uffe0", "\u00a3": "\uffe1", "\u00ac": "\uffe2",   # ￠ ￡ ￢
}
_unencodable = collections.Counter()


def _cp932(s: str) -> str:
    """CP932 で書けない文字を置換する。どうしても無い文字はシステム同様 '?' にする（件数は記録）。"""
    try:
        s.encode("cp932")
        return s
    except UnicodeEncodeError:
        pass
    out = []
    for ch in s:
        ch = _CP932_FIX.get(ch, ch)
        try:
            ch.encode("cp932")
        except UnicodeEncodeError:
            _unencodable[ch] += 1
            ch = "?"
        out.append(ch)
    return "".join(out)


def _dt(v: datetime | None) -> str:
    return v.strftime(DT_FMT) if v else ""


def _legacy_text(s: str) -> str:
    """旧システムは全角入力を強制していた：ASCII英数記号・スペースを全角化し、改行は「／」で連結。"""
    if not s:
        return s
    s = s.replace("\r\n", "\n").replace("\n", "／")
    return D.to_zenkaku(s).replace(" ", "　")


_DELAY_RE = re.compile(r"報告書記入が(\d+)日遅れ")


def _report_delay_days(remarks: str) -> int | None:
    m = _DELAY_RE.search(unicodedata.normalize("NFKC", remarks))
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# 行の組み立て
# ---------------------------------------------------------------------------
class _Builder:
    def __init__(self):
        self.r = D.rng("t2_system_csv")
        self.person_by_id = {p.employee_id: p for p in D.people()}
        self.person_by_name = {p.name: p for p in D.people()}
        self.used = collections.Counter()          # (種別, コード) -> 使用件数
        self.legacy_n = 0
        self.migrated_open_n = 0                   # 移行時点で未完了だった（SYS-MIG 登録・社員が更新）件数
        self.minor_n = 0
        self.migrated_seq = 0

    def use(self, kind: str, code: str) -> str:
        if code:
            self.used[(kind, code)] += 1
        return code

    def equipment_cols(self, eq: D.Equipment) -> list[str]:
        return [self.use("設備", eq.equipment_id), self.use("設備分類", _code_of(EQ_CAT, eq.category, "K", 2)),
                self.use("ライン", _code_of(LINE, eq.line, "LX", 2))]

    def chief_of(self, p: D.Person) -> D.Person | None:
        for q in D.people():
            if q.section == p.section and q.role == "係長":
                return q
        return None

    # --- 設備故障 ---
    def incident_row(self, inc: D.Incident, sub_codes: dict) -> list[str]:
        r = self.r
        # 移行データ＝旧システム時代に発生し、移行時点で完了済みだったもの（未完了分は新システムで更新が続いた扱い）
        legacy = inc.occurred_at < LEGACY_BEFORE and inc.completed_at is not None and inc.completed_at < MIGRATED_AT
        eq = inc.equipment
        assignees = inc.assignees
        primary = assignees[0] if assignees else None
        for p in [inc.reporter] + list(assignees):
            self.use("社員", p.employee_id)
            self.use("部門", _section_code(p))

        # 復旧区分: 完了系のみ。FEが入っていればメーカー復旧
        recovery = ""
        if inc.completed_at is not None:
            recovery = RECOVERY["メーカー復旧"] if any(p.role == "FE" for p in assignees) else RECOVERY["保全復旧"]

        parts = ",".join(f"{p.part_no}x{q}" for p, q in inc.parts_used)
        for p, _q in inc.parts_used:
            self.use("部品", p.part_no)

        remarks = inc.remarks
        if inc.alarm is not None and not legacy and r.random() < 0.03:
            # 装置画面の表示文をそのまま貼り付けた記入（"" エスケープが発生する）
            quoted = f'装置画面表示 "{inc.alarm.message}"'
            remarks = f"{remarks}。{quoted}" if remarks else quoted

        # 登録・更新
        if legacy:
            self.migrated_seq += 1
            reg_at = MIGRATED_AT + timedelta(seconds=self.migrated_seq * 2)
            reg_by = upd_by = SYS_MIG
            upd_at = reg_at
        else:
            if inc.occurred_at < LEGACY_BEFORE:
                # 移行時点で未完了だった旧システムの案件: 移行バッチが登録し、以後は新システムで担当者が更新した
                self.migrated_seq += 1
                self.migrated_open_n += 1
                reg_at = MIGRATED_AT + timedelta(seconds=self.migrated_seq * 2)
                reg_by = SYS_MIG
            else:
                reg_at = inc.reported_at + timedelta(minutes=r.randint(2, 95), seconds=r.randint(0, 59))
                reg_at = max(reg_at, MIGRATED_AT + timedelta(minutes=r.randint(20, 90)))   # 移行直後の発生分も新システム稼働後に登録
                reg_by = inc.reporter.employee_id
            if inc.completed_at is not None:
                delay = _report_delay_days(remarks)
                if delay is not None:
                    upd_at = inc.completed_at + timedelta(days=delay, minutes=r.randint(0, 600))
                else:
                    upd_at = inc.completed_at + timedelta(minutes=r.randint(10, 60 * 30), seconds=r.randint(0, 59))
                chief = self.chief_of(primary) if primary else None
                upd_by = chief.employee_id if chief and r.random() < 0.3 else (primary or inc.reporter).employee_id
            elif inc.status == "対応中":
                # 対応中は調査・部品手配の進捗をこまめに更新している（直近2週間以内）
                upd_at = EXPORT_AT - timedelta(minutes=r.randint(60, 60 * 24 * 14), seconds=r.randint(0, 59))
                upd_by = (primary or inc.reporter).employee_id
            else:
                # 保留は月1回程度の棚卸しで状況を更新（直近45日以内、係長が更新することも多い）
                upd_at = EXPORT_AT - timedelta(days=r.randint(1, 45), minutes=r.randint(0, 600))
                chief = self.chief_of(primary) if primary else None
                upd_by = chief.employee_id if chief and r.random() < 0.4 else (primary or inc.reporter).employee_id
            if reg_by == SYS_MIG:
                upd_at = max(upd_at, MIGRATED_AT + timedelta(hours=r.randint(6, 72)))      # 更新は移行より後
            # 出力日時より後にはならないように丸める（ただし完了・登録より前には戻さない）
            floor = max(reg_at, inc.completed_at or reg_at)
            upd_at = max(min(upd_at, EXPORT_AT - timedelta(minutes=r.randint(30, 600))), floor)

        texts = [inc.symptom, inc.first_response, inc.investigation, inc.cause, " → ".join(inc.why_why), inc.action]
        tail_texts = [inc.result, inc.prevention, inc.horizontal_deployment, remarks]
        cost = str(inc.cost_yen)
        if legacy:
            self.legacy_n += 1
            texts = [_legacy_text(t) for t in texts]
            tail_texts = [_legacy_text(t) for t in tail_texts]
            cost = f"{inc.cost_yen:,}"

        cause_cat = _code_of(CAUSE_CAT, inc.cause_category, "C", 2) if inc.cause_category else ""
        return [
            inc.incident_id, self.use("記録区分", RECORD_KIND["設備故障"]), *self.equipment_cols(eq),
            _dt(inc.occurred_at), _dt(inc.reported_at), "" if legacy else _dt(inc.response_started_at), _dt(inc.completed_at),
            self.use("シフト", SHIFT[inc.shift]), self.use("発見区分", _code_of(DETECTED, inc.detected_by, "H9", 1)),
            inc.reporter.employee_id, _section_code(inc.reporter),
            ",".join(p.employee_id for p in assignees), _section_code(primary) if primary else "",
            self.use("故障区分", _code_of(FAILURE_CAT, inc.category, "F", 2)), self.use("部位", sub_codes[inc.subsystem]),
            self.use("重要度", SEVERITY[inc.severity]), self.use("状態", STATUS[inc.status]),
            self.use("アラーム", inc.alarm.code if inc.alarm else ""),
            *texts[:4], self.use("原因区分", cause_cat), *texts[4:], parts,
            str(inc.downtime_min), f"{inc.work_hours:.2f}", cost, ",".join(inc.lots_affected), str(inc.scrap_wafers),
            self.use("復旧区分", recovery), "1" if inc.recurrence else "0", inc.related_incident_id or "",
            *tail_texts, str(len(inc.photos)),
            _dt(reg_at), reg_by, _dt(upd_at), upd_by,
        ]

    # --- チョコ停 ---
    def minor_row(self, ms: D.MinorStop, sub_codes: dict) -> list[str]:
        r = self.r
        self.minor_n += 1
        eq = ms.equipment
        end =ms.occurred_at + timedelta(minutes=ms.duration_min)
        person = self.person_by_name.get(ms.recovered_by)
        if person is None:
            recovery = RECOVERY["自動復帰"]
        elif person.section.startswith("製造課"):
            recovery = RECOVERY["製造復旧"]
        else:
            recovery = RECOVERY["保全復旧"]
        if person:
            self.use("社員", person.employee_id)
            self.use("部門", _section_code(person))
        if ms.alarm is not None:
            detected = "装置アラーム"
        elif person is not None:
            detected = "オペレーター"
        else:
            detected = ""
        reg_at = end + timedelta(seconds=r.randint(5, 59))
        upd_at, upd_by = reg_at, SYS_MES
        if person is not None and r.random() < 0.06:
            # 復旧者が後からメモを追記したもの
            upd_at = reg_at + timedelta(minutes=r.randint(5, 60 * 10))
            upd_by = person.employee_id
        return [
            ms.event_id, self.use("記録区分", RECORD_KIND["チョコ停"]), *self.equipment_cols(eq),
            _dt(ms.occurred_at), _dt(ms.occurred_at), "", _dt(end),
            self.use("シフト", SHIFT[D.shift_of(ms.occurred_at)]),
            self.use("発見区分", _code_of(DETECTED, detected, "H9", 1)) if detected else "",
            "", "", person.employee_id if person else "", _section_code(person) if person else "",
            "", self.use("部位", sub_codes[ms.alarm.subsystem]) if ms.alarm else "", "",
            self.use("状態", STATUS["完了"]), self.use("アラーム", ms.alarm.code if ms.alarm else ""),
            ms.note, "", "", "", "", "", "", "",
            str(ms.duration_min), "", "", "", "", self.use("復旧区分", recovery), "0", "",
            "", "", "", "", "0",
            _dt(reg_at), SYS_MES, _dt(upd_at), upd_by,
        ]


# ---------------------------------------------------------------------------
# コード表
# ---------------------------------------------------------------------------
def _code_table_rows(b: _Builder, sub_codes: dict) -> list[list[str]]:
    rows: list[list[str]] = []

    def add(kind, code, name, rel="", note=""):
        rows.append([kind, code, name, rel, note, str(b.used.get((kind, code), 0))])

    for name, code in RECORD_KIND.items():
        add("記録区分", code, name, note="1=保全管理システム手入力 / 2=MES連携で自動登録")
    for name, code in FAILURE_CAT.items():
        add("故障区分", code, name, note="チョコ停は空欄")
    for name, code in SEVERITY.items():
        add("重要度", code, name, note={"1": "生産停止・安全/品質影響大", "2": "半日以上停止", "3": "数時間停止", "4": "軽微"}[code])
    for name, code in STATUS.items():
        add("状態", code, name)
    for name, code in SHIFT.items():
        add("シフト", code, name, note={"1": "8:00-20:00", "2": "20:00-翌8:00", "3": "土日・連休（基準日で判定）"}[code])
    for name, code in DETECTED.items():
        add("発見区分", code, name)
    for name, code in CAUSE_CAT.items():
        add("原因区分", code, name)
    for name, code in RECOVERY.items():
        add("復旧区分", code, name)
    for name, code in EQ_CAT.items():
        add("設備分類", code, name)
    for name, code in LINE.items():
        add("ライン", code, name, note=LINE_NOTE.get(code, ""))
    for sub, code in sorted(sub_codes.items(), key=lambda kv: kv[1]):
        g = code[1]
        add("部位", code, sub, note="共通部位" if g == "0" else list(EQ_CAT)[int(g) - 1])
    for (dept, sect), code in sorted(SECTION.items(), key=lambda kv: kv[1]):
        add("部門", code, f"{dept} {sect}", note="取引先（設備メーカー）" if dept == "設備メーカー" else "")
    for p in D.people():
        add("社員", p.employee_id, p.name, rel=_section_code(p), note=p.role)
    # システムユーザの使用件数は登録者IDとしての件数
    rows.append(["社員", SYS_MES, "MES連携バッチ", "", "システムユーザ（チョコ停の自動登録）", str(b.minor_n)])
    rows.append(["社員", SYS_MIG, "旧システム移行バッチ", "", f"システムユーザ（{MIGRATED_AT:%Y/%m/%d} 移行）", str(b.legacy_n + b.migrated_open_n)])
    for eq in D.equipment_master():
        add("設備", eq.equipment_id, eq.name, rel=f"{EQ_CAT[eq.category]}/{LINE[eq.line]}",
            note=f"{eq.maker} {eq.model}／{eq.area}／設置 {eq.installed_date:%Y/%m/%d}／重要度ランク {eq.criticality}")
    for a in D.alarm_codes():
        add("アラーム", a.code, a.message, rel=sub_codes.get(a.subsystem, ""), note=a.category)
    for p in D.parts_catalog():
        add("部品", p.part_no, p.name, rel=sub_codes.get(p.subsystem, ""),
            note=f"{p.maker}／単価 {p.unit_price_yen}円／適用: {'・'.join(p.applicable_categories)}")
    return rows


# ---------------------------------------------------------------------------
# README
# ---------------------------------------------------------------------------
def _readme(stats: dict) -> str:
    cols = "\n".join(f"| {i} | {h} | {stats['col_desc'].get(h, '')} |" for i, h in enumerate(HEADER, 1))
    kinds = "、".join(f"{k}（{n}件）" for k, n in stats["code_kinds"].items())
    return f"""# T2 設備故障履歴（システム出力CSV）

保全管理システムの「故障履歴照会」画面から CSV エクスポートした、という想定のサンプル。
設備故障（手入力の故障報告）とチョコ停（MES連携の自動登録）を1ファイルにマージし、
名称ではなく**コード値**で出力している。コードの意味は `{CODE_NAME}` で引ける。

生成: `python -m scripts.samples.t2_system_csv`（固定シード。再実行しても同一ファイル）

## ファイル

| ファイル | 文字コード / 改行 | 内容 |
|---|---|---|
| `{CSV_NAME}` | CP932（Shift_JIS）/ CRLF | 本体。{stats['size_kb']:,} KB |
| `{CODE_NAME}` | CP932（Shift_JIS）/ CRLF | コード表 {stats['code_rows']:,} 行 |
| `{README_NAME}` | UTF-8 / LF | このファイル |

## 本体CSVの構造

```
1行目  出力日時,{EXPORT_AT:%Y/%m/%d %H:%M:%S}
2行目  出力者,{EXPORT_USER.employee_id},{EXPORT_USER.name}
3行目  抽出条件,期間=... ,記録区分=... ,部門=... ,（複数セル）
4行目  （空行）
5行目  ヘッダ（{len(HEADER)}列）
6行目〜 データ {stats['rows']:,} レコード（発生日時の昇順）
最終行  合計件数,{stats['rows']}
```

- レコード数: **{stats['rows']:,}**（設備故障 {stats['n_inc']:,} ＋ チョコ停 {stats['n_ms']:,}）
- 物理行数: {stats['physical_lines']:,}（処置内容のセル内改行を含むため、レコード数＋6 とは一致しない）
- 期間: {PERIOD_FROM:%Y/%m/%d} 〜 {PERIOD_TO:%Y/%m/%d}（発生日時で抽出）
- 日時書式: `YYYY/MM/DD HH:MM:SS`（故障は秒が常に 00、チョコ停は秒まで記録）
- クォート: 必要な項目だけダブルクォートで囲む（カンマ・改行・`"` を含む項目）。`"` は `""` にエスケープ

## 列定義

| # | 列名 | 内容 |
|---|---|---|
{cols}

## コード表 `{CODE_NAME}`

列: `コード種別, コード, 名称, 関連コード, 補足, 使用件数`（使用件数＝本体CSVでそのコードが出現した件数。
社員・部門は報告者＋担当者としての出現数、システムユーザは登録者IDとしての件数、部品は使用部品に現れたレコード行数）。
種別: {kinds}。

- 部位コード `B{{g}}{{nn}}`: g=設備分類（K0g）の部位、g=0 は3分類以上で共通の部位（ホスト通信・電源/受電など）
- 部門コード: 5桁の階層コード（21000 設備保全課 / 21010 保全1係 …）。9xxxx は設備メーカー（取引先）
- 社員には `SYS-MES`（MES連携）/ `SYS-MIG`（旧システム移行）のシステムユーザを含む
- 設備の関連コードは `設備分類コード/ラインコード`、アラーム・部品の関連コードは部位コード

## 他サンプルとの整合

`scripts/samples/domain.py` の正準データ（`standard_incidents()` / `minor_stops()`）をそのまま出力しているので、
管理番号（`TR-YYYY-NNNNN` / `MS-YYYY-NNNNN`）、設備コード、社員番号、アラームコード、部品品番は
Excel帳票系サンプルと一致する。

## 意図的な不規則性（前処理のテスト観点）

1. **ヘッダ前のメタ情報4行と末尾トレーラ**: 1〜4行目は出力条件（列数不定・4行目は空行）、最終行 `合計件数,{stats['rows']}` は2列だけ。
2. **セル内改行**: 処置内容はクォート内に LF 改行を含む（{stats['multiline']:,} レコード）。レコード区切りは CRLF、セル内は LF。
3. **カンマを含むクォート項目**: 担当者社員番号（複数名は `M12410,FE-0213`）、使用部品（`PV62-1166x1,PV53-1182x2`）、対象ロット、
   「10,300Å」のような数値を含む自由記述など。カンマを含むフィールドは {stats['comma_fields']:,} 個。
4. **`""` エスケープ**: 備考に装置画面の表示文を `"..."` 付きで貼り付けたレコードが {stats['dq_rows']:,} 件。
5. **空欄が多い**: チョコ停行は報告者・故障区分・重要度・原因・工数・費用などが空欄。
   故障でも対応中/保留は完了日時が空欄（{stats['no_completed']:,} 件）、対応中は原因・なぜなぜ分析が空欄（原因空欄 {stats['no_cause']:,} 件）。
6. **旧システム移行データ**（{LEGACY_BEFORE:%Y/%m/%d} より前に発生し、{MIGRATED_AT:%Y/%m/%d %H:%M} の移行時点で完了済みだった設備故障 {stats['legacy']:,} 件）:
   自由記述の英数字・記号・スペースが全角（例 `ＭＦＣ`、`ＣＶＤ－２０５`）、改行は `／` に置換、
   対応開始日時が空欄、費用(円)が `634,600` のようなカンマ区切り文字列（クォートされる）、登録者/更新者は `SYS-MIG`。
   移行時点で未完了だった {stats['migrated_open']:,} 件は、登録者 `SYS-MIG`・登録日時が移行日時で、更新者は移行後に更新した社員番号（自由記述は新システムで書き直されており半角のまま）。
7. **全角・半角の混在**: 記入者の癖による全角数字・半角カナ（`ｱﾗｰﾑ`）・丸数字の手順番号・誤変換（異常→以上 等）が自由記述に残っている。
8. **CP932 化の置換**: `Å`(U+00C5) は CP932 に無いため `Å`(U+212B) に置換。置換できず `?` になった文字: {stats['unencodable']}。
   （これとは別に、元の記入に由来する半角 `?` が自由記述中に {stats['qmarks']:,} 箇所ある）
9. **コードとIDの形式ゆれ**: 社員番号は `M12410`（社員）/ `FE-0213`（メーカーFE）/ `SYS-MES`（システム）が混在。
   重要度・状態・シフトは1桁数字、故障区分は `F01`、ラインは `LN01`/`LN13`/`UT00` と体系が異なる。
10. **数値の書式**: 作業工数(H)は小数2桁（`3.00`）、停止時間(分)は整数、費用は通常は整数・移行データのみカンマ付き。
11. **未完了案件の更新日時**: 対応中は直近2週間以内、保留は直近45日以内に更新されている（棚卸しで状況を更新）。
    対応中 {stats['n_ongoing']:,} 件のうち発生から60日以内が {stats['ongoing_recent']:,} 件、保留 {stats['n_hold']:,} 件のうち発生から半年超が {stats['hold_old']:,} 件。

## 件数の内訳

- 記録区分: 設備故障 {stats['n_inc']:,} / チョコ停 {stats['n_ms']:,}
- 年別（発生日時）: {stats['by_year']}
- 重要度（故障のみ）: {stats['by_sev']}
- 状態: {stats['by_status']}
- 自由記述の異なり数: {stats['distinct']}
"""


_COL_DESC = {
    "管理番号": "TR-YYYY-NNNNN（設備故障）/ MS-YYYY-NNNNN（チョコ停）。年ごとの連番",
    "記録区分": "1=設備故障, 2=チョコ停",
    "設備コード": "CMP-101 等（コード表: 設備）",
    "設備分類コード": "K01〜K09",
    "ラインコード": "LN01〜LN06, LN13, LN46, UT00〜UT02",
    "発生日時": "YYYY/MM/DD HH:MM:SS",
    "報告日時": "チョコ停は発生日時と同じ",
    "対応開始日時": "チョコ停・移行データは空欄",
    "完了日時": "対応中/保留は空欄。チョコ停は発生＋停止時間",
    "シフトコード": "1=日勤, 2=夜勤, 3=休日",
    "発見区分コード": "H01〜（コード表: 発見区分）。アラーム無し自動復帰のチョコ停は空欄",
    "報告者社員番号": "チョコ停は空欄",
    "報告者部門コード": "5桁（コード表: 部門）",
    "担当者社員番号": "複数名はカンマ区切り（クォート）。チョコ停は復旧者、自動復帰は空欄",
    "担当部門コード": "先頭担当者の部門",
    "故障区分コード": "F01〜F08。チョコ停は空欄",
    "部位コード": "B001〜（コード表: 部位）",
    "重要度コード": "1=重大, 2=大, 3=中, 4=小。チョコ停は空欄",
    "状態コード": "1=対応中, 2=保留, 3=経過観察, 9=完了",
    "アラームコード": "A-3102 等。無い場合は空欄",
    "現象": "自由記述。チョコ停はMESのメモ",
    "初期対応": "自由記述",
    "調査内容": "自由記述（複数文）",
    "原因": "自由記述",
    "原因区分コード": "C01〜C15, C99=不明",
    "なぜなぜ分析": "「 → 」区切り",
    "処置内容": "番号付き手順。セル内改行（LF）あり",
    "使用部品": "品番x数量 をカンマ区切り",
    "停止時間(分)": "整数",
    "作業工数(H)": "小数2桁",
    "費用(円)": "整数（移行データはカンマ区切り文字列）",
    "対象ロット": "ロットIDをカンマ区切り",
    "廃棄枚数": "整数",
    "復旧区分コード": "1=自動復帰, 2=製造復旧, 3=保全復旧, 4=メーカー復旧。未完了の故障は空欄",
    "再発フラグ": "0/1",
    "関連管理番号": "再発元・瞬低の同時多発の親など",
    "結果": "自由記述", "再発防止策": "自由記述", "水平展開": "自由記述", "備考": "自由記述",
    "添付数": "写真などの添付ファイル数",
    "登録日時": "システム登録日時", "登録者ID": "社員番号 / SYS-MES / SYS-MIG",
    "更新日時": "最終更新日時", "更新者ID": "社員番号 / SYS-MES / SYS-MIG",
}


# ---------------------------------------------------------------------------
# 生成
# ---------------------------------------------------------------------------
def generate(output_root: Path | None = None) -> list[Path]:
    root = Path(output_root) if output_root is not None else D.OUTPUT_ROOT
    out_dir = root / TABLE_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    _unencodable.clear()

    incidents = D.standard_incidents()
    stops = D.minor_stops()
    sub_codes = _subsystem_codes(incidents, stops)
    b = _Builder()

    # 発生日時 → 記録区分 → 管理番号 の順にマージ
    merged = [(inc.occurred_at, 1, inc.incident_id, inc) for inc in incidents] + \
             [(ms.occurred_at, 2, ms.event_id, ms) for ms in stops]
    merged.sort(key=lambda x: x[:3])
    rows = []
    for _at, kind, _id, obj in merged:
        row = b.incident_row(obj, sub_codes) if kind == 1 else b.minor_row(obj, sub_codes)
        assert len(row) == len(HEADER), (obj, len(row))
        rows.append([_cp932(v) for v in row])

    # --- 本体CSV ---
    csv_path = out_dir / CSV_NAME
    with open(csv_path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
        w.writerow(["出力日時", EXPORT_AT.strftime(DT_FMT)])
        w.writerow(["出力者", EXPORT_USER.employee_id, EXPORT_USER.name])
        w.writerow(["抽出条件", f"期間={PERIOD_FROM:%Y/%m/%d %H:%M:%S}～{PERIOD_TO:%Y/%m/%d %H:%M:%S}",
                    "記録区分=1:設備故障,2:チョコ停", "部門=21000(設備保全課) 配下すべて", "削除済み=含まない"])
        f.write("\r\n")
        w.writerow(HEADER)
        w.writerows(rows)
        w.writerow(["合計件数", str(len(rows))])

    # --- コード表 ---
    code_rows = _code_table_rows(b, sub_codes)
    code_path = out_dir / CODE_NAME
    with open(code_path, "w", encoding="cp932", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(["コード種別", "コード", "名称", "関連コード", "補足", "使用件数"])
        w.writerows([[_cp932(v) for v in row] for row in code_rows])

    # --- README（件数は実データから集計） ---
    idx = {h: i for i, h in enumerate(HEADER)}
    inc_rows = [r_ for r_ in rows if r_[idx["記録区分"]] == "1"]
    text_cols = ["現象", "初期対応", "調査内容", "原因", "なぜなぜ分析", "処置内容", "結果", "再発防止策", "水平展開", "備考"]
    rev = lambda t: {v: k for k, v in t.items()}  # noqa: E731
    stats = dict(
        rows=len(rows), n_inc=len(incidents), n_ms=len(stops),
        size_kb=round(csv_path.stat().st_size / 1024),
        code_rows=len(code_rows),
        physical_lines=csv_path.read_bytes().count(b"\n"),   # セル内LFも数える（テキストエディタで見た行数）
        multiline=sum(1 for r_ in rows if any("\n" in v for v in r_)),
        comma_fields=sum(1 for r_ in rows for v in r_ if "," in v),
        dq_rows=sum(1 for r_ in rows if any('"' in v for v in r_)),
        no_completed=sum(1 for r_ in inc_rows if not r_[idx["完了日時"]]),
        no_cause=sum(1 for r_ in inc_rows if not r_[idx["原因"]]),
        legacy=b.legacy_n,
        migrated_open=b.migrated_open_n,
        n_ongoing=sum(1 for i in incidents if i.status == "対応中"),
        ongoing_recent=sum(1 for i in incidents if i.status == "対応中" and (PERIOD_TO - i.occurred_at).days <= 60),
        n_hold=sum(1 for i in incidents if i.status == "保留"),
        hold_old=sum(1 for i in incidents if i.status == "保留" and (PERIOD_TO - i.occurred_at).days > 182),
        qmarks=sum(v.count("?") for r_ in rows for v in r_),
        unencodable=("なし" if not _unencodable else "、".join(f"U+{ord(c):04X}×{n}" for c, n in _unencodable.items())),
        by_year="、".join(f"{y}年 {n:,}" for y, n in sorted(collections.Counter(x[0].year for x in merged).items())),
        by_sev="、".join(f"{k} {n:,}" for k, n in sorted(collections.Counter(i.severity for i in incidents).items(),
                                                           key=lambda kv: SEVERITY[kv[0]])),
        by_status="、".join(f"{rev(STATUS)[k]}({k}) {n:,}" for k, n in sorted(collections.Counter(r_[idx["状態コード"]] for r_ in rows).items())),
        distinct="、".join(f"{c} {len({r_[idx[c]] for r_ in rows if r_[idx[c]]}):,}" for c in text_cols),
        code_kinds=dict(collections.Counter(row[0] for row in code_rows)),
        col_desc=_COL_DESC,
    )
    readme_path = out_dir / README_NAME
    readme_path.write_text(_readme(stats), encoding="utf-8", newline="\n")
    return [csv_path, code_path, readme_path]


if __name__ == "__main__":
    import time

    t0 = time.time()
    for p in generate():
        print(f"{p}  ({p.stat().st_size / 1024:,.0f} KB)")
    print(f"done in {time.time() - t0:.1f}s")
