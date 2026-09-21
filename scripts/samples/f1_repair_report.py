"""F1 設備修理報告書（Excel帳票）のサンプルを作る。

    python -m scripts.samples.f1_repair_report     # samples/forms/F1_設備修理報告書/ に出力

domain.standard_incidents() のうち、重要度「大」「重大」または部品を使ったトラブルから30件を選び、
1件＝1ファイルの修理報告書にする。様式は改訂の経緯に合わせて3版が混在する。

- Rev.1（v1）: 左ラベル・右値の2列組み。文章項目は見出し行の下に結合セル。重要度はチェックボックス表記
- Rev.2（v2）: B列始まりで列がズレる。項目名の揺れ（装置No./設備No.、故障内容→症状 など）、
               「設備番号：CMP-101」のような1セル内ラベル、承認欄は下部
- Rev.3（v3）: 上部は見出し行＋値行の表形式、文章項目は縦書きの縦結合ラベル＋右の結合セル

あわせて、抽出結果の正解データ _expected.jsonl と説明 _README.md を書き出す。
乱数は domain.rng() からのみ取り、ZIP内のタイムスタンプも固定するため、何度実行しても同一バイトのファイルになる。
"""
from __future__ import annotations

import json
import math
import shutil
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.writer.excel import ExcelWriter
from PIL import Image as PILImage, ImageDraw, ImageFont

from . import domain as D

# ---------------------------------------------------------------------------
# 基本設定
# ---------------------------------------------------------------------------
FORM_DIR = "F1_設備修理報告書"
N_FILES = 30
FORM_NO = "様式MT-012"
FONT_NAME = "ＭＳ ゴシック"
TODAY = date(2026, 9, 14)            # 報告日の上限（サンプル作成時点）
ZIP_TIME = (2026, 9, 1, 0, 0, 0)     # xlsx(ZIP)内エントリの固定タイムスタンプ

REV_INFO = {
    "v1": ("Rev.1", "2019.10制定"),
    "v2": ("Rev.2", "2024.04改訂"),
    "v3": ("Rev.3", "2025.07改訂"),
}

# Windows のファイル名に使えない文字
_NG_CHARS = '.:/\\*?"<>|'


def _version_dir(version: str) -> str:
    """版ごとのサブフォルダ名（例: Rev.1・2019.10制定 → Rev1_2019制定）。

    名前の付け方は F1〜F5 で共通で、「版の名前＋いつから」。ローマ字は使わず、
    Windows のフォルダ名に使えない文字（. : / \\ * ? " < > |）も使わない（`Rev.1` → `Rev1`）。
    """
    label, since = REV_INFO[version]
    year, kind = since[:4], since.lstrip("0123456789.")     # 2019.10制定 → 2019 / 制定
    return "".join(c for c in f"{label}_{year}{kind}" if c not in _NG_CHARS)


THIN = Side(style="thin", color="000000")
MEDIUM = Side(style="medium", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
STAMP_RED = "C00000"

# 係ごとの確認者（係長／主任）と課長
SECTION_HEAD = {"設備保全課 保全1係": "M10544", "設備保全課 保全2係": "M10688", "設備保全課 施設係": "M10902"}
MANAGER_ID = "M10231"

# 1ファイル内で「必須項目の空欄」にする候補（版ごと）
BLANK_CANDIDATES = {
    "v1": ["report_date", "assignee", "downtime_min", "cause", "result", "recovered_at"],
    "v2": ["report_date", "assignee", "cause", "result", "downtime_min", "reporter"],
    "v3": ["report_date", "assignee", "cause", "prevention", "downtime_min", "recovered_at"],
}

# v2 の項目名ゆれ（ファイルごとに1つ選ぶ）
V2_SYNONYMS = {
    "report_id": ["報告No.", "報告書No.", "No."],
    "report_date": ["報告日", "作成日", "起票日"],
    "incident_id": ["トラブルNo.", "TR No.", "管理番号"],
    "reporter": ["起票者", "報告者", "連絡者"],
    "equipment_id": ["装置No.", "設備No.", "設備番号"],
    "equipment_name": ["装置名", "設備名"],
    "line": ["ライン/工程", "ライン"],
    "assignee": ["作業者", "対応者", "担当者"],
    "occurred_at": ["発生日時", "故障発生", "発生"],
    "recovered_at": ["復旧完了", "復旧日時", "完了日時"],
    "downtime_min": ["ダウンタイム", "停止時間", "設備停止"],
    "work_hours": ["作業工数", "工数"],
    "category": ["故障区分", "区分"],
    "severity": ["ランク", "重要度", "影響度"],
    "alarm": ["ALM No.", "アラーム", "エラーコード"],
    "symptom": ["症状", "現象", "不具合内容"],
    "first_response": ["応急処置", "初期対応"],
    "investigation": ["調査内容", "調査結果", "原因調査"],
    "cause": ["故障原因", "原因", "推定原因"],
    "action": ["処置内容", "修理内容", "対策内容"],
    "parts": ["使用部品", "交換部品"],
    "result": ["確認結果", "修理結果", "結果"],
    "prevention": ["恒久対策", "再発防止", "再発防止策"],
    "lots": ["影響ロット", "影響Lot"],
    "photos": ["写真", "添付写真"],
    "remarks": ["特記事項", "備考", "その他"],
    "creator": ["担当", "作成"],
}


# ---------------------------------------------------------------------------
# データ構造
# ---------------------------------------------------------------------------
@dataclass
class FileSpec:
    """1ファイル分の設計（どのトラブルを、どの版で、どんなイレギュラー付きで書くか）。"""
    inc: D.Incident
    version: str
    report_id: str = ""
    report_date: date | None = None
    photo_sheet: bool = False
    ref_sheet: bool = False
    hidden: str | None = None          # "リスト" / "作業メモ" / None
    zenkaku: bool = False
    blank_field: str | None = None
    photos_outside: bool = False       # v2: 写真を印刷範囲外（右側）に貼付
    default_sheet_name: bool = False   # v2: シート名が Sheet1 のまま
    filename: str = ""


@dataclass
class Record:
    """_expected.jsonl に書く正解データ。"""
    values: dict = field(default_factory=dict)
    labels: dict = field(default_factory=dict)
    irregularities: list = field(default_factory=list)
    images: dict = field(default_factory=dict)
    extra_sheets: dict = field(default_factory=dict)
    hidden_sheets: list = field(default_factory=list)
    main_sheet: str = ""


# ---------------------------------------------------------------------------
# 対象トラブルの選定・ファイル設計
# ---------------------------------------------------------------------------
def _is_candidate(i: D.Incident) -> bool:
    return i.severity in ("大", "重大") or bool(i.parts_used)


@lru_cache(maxsize=None)
def _candidate_pool() -> tuple:
    return tuple(i for i in D.standard_incidents() if _is_candidate(i))


def _select_incidents() -> list[D.Incident]:
    """期間・設備カテゴリ・重要度・状態が偏らないよう30件を選ぶ。"""
    r = D.rng("f1-repair-report:select")
    pool = list(_candidate_pool())
    by_id = {i.incident_id: i for i in pool}
    chosen: list[D.Incident] = []
    used_ids: set[str] = set()
    used_scen: set[str] = set()

    def ok(i: D.Incident) -> bool:
        # 瞬低の同時多発（従属側）は報告書が別にまとめられるため対象外
        return (i.incident_id not in used_ids and i.scenario_id not in used_scen
                and "同時刻に工場内複数設備が停止" not in i.symptom)

    def take(i: D.Incident) -> None:
        chosen.append(i)
        used_ids.add(i.incident_id)
        used_scen.add(i.scenario_id)

    # 1) 同一設備の再発ペア（前回報告書と今回報告書が両方ある）
    pairs = [(by_id[b.related_incident_id], b) for b in pool
             if b.recurrence and b.related_incident_id in by_id and "同時刻" not in b.symptom
             and by_id[b.related_incident_id].equipment.equipment_id == b.equipment.equipment_id
             and b.status == "完了" and by_id[b.related_incident_id].status == "完了"
             and (b.occurred_at - by_id[b.related_incident_id].occurred_at).days >= 7]
    a, b = r.choice(pairs)
    take(a)
    chosen.append(b)          # 同じシナリオでも再発なので採用する
    used_ids.add(b.incident_id)

    # 2) 状態・重要度の特殊ケース
    def pick(pred, n: int) -> None:
        cands = [i for i in pool if ok(i) and pred(i)]
        cats_done = {i.equipment.category for i in chosen}
        for _ in range(n):
            fresh = [i for i in cands if ok(i) and i.equipment.category not in cats_done] or [i for i in cands if ok(i)]
            if not fresh:
                return
            x = r.choice(fresh)
            take(x)
            cats_done.add(x.equipment.category)

    pick(lambda i: i.severity == "重大" and i.status == "完了" and bool(i.photos), 3)
    pick(lambda i: i.status == "対応中" and (i.severity in ("大", "重大") or i.parts_used), 2)
    pick(lambda i: i.status == "保留", 1)
    pick(lambda i: i.status == "経過観察" and i.severity == "大", 2)

    # 3) 期間ごとの枠を、カテゴリが偏らないように埋める（大/重大 と 部品使用の中/小 を 6:4）
    periods = [(date(2023, 4, 1), date(2024, 3, 31), 8), (date(2024, 4, 1), date(2025, 6, 30), 10),
               (date(2025, 7, 1), date(2026, 8, 31), 12)]
    cats = sorted({e.category for e in D.equipment_master()})
    special = set(used_ids)

    def in_period(i: D.Incident, p) -> bool:
        return p[0] <= i.occurred_at.date() <= p[1]

    for p in periods:
        have = sum(1 for i in chosen if in_period(i, p))
        while have < p[2]:
            count = {c: sum(1 for i in chosen if i.equipment.category == c) for c in cats}
            cat = min(cats, key=lambda c: (count[c], r.random()))
            big = r.random() < 0.6
            base = [i for i in pool if ok(i) and in_period(i, p)]
            cands = ([i for i in base if i.equipment.category == cat and i.status == "完了" and (i.severity in ("大", "重大")) == big]
                     or [i for i in base if i.equipment.category == cat] or base)
            # 写真付きの記録をやや優先
            x = r.choices(cands, weights=[2.0 if i.photos else 1.0 for i in cands])[0]
            take(x)
            have += 1
    # 特殊ケースが枠を超えた場合は、超過の大きい期間から通常枠の記録を外して30件にそろえる
    while len(chosen) > N_FILES:
        over = max(periods, key=lambda p: sum(1 for i in chosen if in_period(i, p)) - p[2])
        removable = [i for i in chosen if i.incident_id not in special and in_period(i, over)] or \
                    [i for i in chosen if i.incident_id not in special]
        chosen.remove(r.choice(removable))
    chosen.sort(key=lambda i: i.occurred_at)
    return chosen


def _report_numbers() -> dict[str, dict[str, str]]:
    """版ごとの報告番号体系で、候補全件に通し番号を振る（選ばれた30件はその一部）。"""
    fy_seq: dict[int, int] = {}
    y_seq: dict[int, int] = {}
    m_seq: dict[tuple, int] = {}
    out = {}
    for i in _candidate_pool():
        at = i.occurred_at
        fy = at.year if at.month >= 4 else at.year - 1
        fy_seq[fy] = fy_seq.get(fy, 0) + 1
        y_seq[at.year] = y_seq.get(at.year, 0) + 1
        m_seq[(at.year, at.month)] = m_seq.get((at.year, at.month), 0) + 1
        out[i.incident_id] = {
            "v1": f"保全-{fy}-{fy_seq[fy]:04d}",
            "v2": f"R{at.year}-{y_seq[at.year]:04d}",
            "v3": f"MR-{at.year % 100:02d}{at.month:02d}-{m_seq[(at.year, at.month)]:03d}",
        }
    return out


def _spread(r, groups: dict[str, list[int]], k: int, allowed=lambda idx: True, exclude: set | None = None) -> list[int]:
    """版をまたいで均等になるように k 件のファイル番号を選ぶ。"""
    exclude = exclude or set()
    pools = {v: [i for i in idx if allowed(i) and i not in exclude] for v, idx in groups.items()}
    for v in pools:
        r.shuffle(pools[v])
    out, order = [], list(groups)
    while len(out) < k and any(pools.values()):
        for v in order:
            if pools[v] and len(out) < k:
                out.append(pools[v].pop())
    return sorted(out)


def _file_specs() -> list[FileSpec]:
    incs = _select_incidents()
    r = D.rng("f1-repair-report:flags")
    nums = _report_numbers()

    # 版: 期間で決まる（Rev.2 は 2024.04〜、Rev.3 は 2025.07〜）。ただし旧様式を使い続ける人もいる
    specs = []
    for i in incs:
        d = i.occurred_at.date()
        ver = "v1" if d < date(2024, 4, 1) else ("v2" if d < date(2025, 7, 1) else "v3")
        specs.append(FileSpec(inc=i, version=ver))
    v2_idx = [n for n, s in enumerate(specs) if s.version == "v2"]
    v3_idx = [n for n, s in enumerate(specs) if s.version == "v3"]
    for n in r.sample(v2_idx, 1):
        specs[n].version = "v1"
    old = r.sample(v3_idx, 2)
    specs[old[0]].version, specs[old[1]].version = "v1", "v2"

    for s in specs:
        inc = s.inc
        s.report_id = nums[inc.incident_id][s.version]
        if inc.completed_at is not None:
            rd = inc.completed_at.date() + timedelta(days=r.choice([0, 0, 1, 1, 1, 2, 3, 4, 7]))
        else:
            rd = inc.reported_at.date() + timedelta(days=r.choice([0, 1, 1, 2]))
        s.report_date = min(rd, TODAY)

    groups = {v: [n for n, s in enumerate(specs) if s.version == v] for v in ("v1", "v2", "v3")}
    complete = lambda n: specs[n].inc.status in ("完了", "経過観察")   # noqa: E731
    for n in _spread(r, groups, 7):
        specs[n].photo_sheet = True
    for n in _spread(r, groups, 7):
        specs[n].ref_sheet = True
    for j, n in enumerate(_spread(r, groups, 6)):
        specs[n].hidden = "リスト" if j % 2 == 0 else "作業メモ"
    for n in _spread(r, groups, 6):
        specs[n].zenkaku = True
    for j, n in enumerate(_spread(r, groups, 6, allowed=complete)):
        cands = BLANK_CANDIDATES[specs[n].version]
        specs[n].blank_field = cands[j % len(cands)]
    v2_now = [n for n, s in enumerate(specs) if s.version == "v2"]
    for n in r.sample(v2_now, 2):
        specs[n].photos_outside = True
    specs[r.choice(v2_now)].default_sheet_name = True

    # ファイル名（現場でありがちな付け方のゆれ）
    used = set()
    for s in specs:
        inc, eq = s.inc, s.inc.equipment
        eid, ymd = eq.equipment_id, inc.occurred_at.strftime("%Y%m%d")
        cands = {
            "v1": [f"設備修理報告書_{s.report_id}_{eid}", f"修理報告書_{eid}_{ymd}"],
            "v2": [f"{inc.occurred_at:%y%m%d}_{eid}_修理報告", f"修理報告書_{inc.incident_id}",
                   f"{eid}修理報告書（{inc.occurred_at.month}月{inc.occurred_at.day}日）"],
            "v3": [f"設備修理報告書_{eid}_{inc.incident_id}", f"{s.report_id}_{eid}_修理報告書"],
        }[s.version]
        name = r.choice(cands)
        if inc.severity == "重大":
            name = "【重大】" + name
        if inc.status in ("対応中", "保留"):
            name += "_中間報告"
        while name in used:
            name += "_2"
        used.add(name)
        s.filename = name + ".xlsx"
    return specs


# ---------------------------------------------------------------------------
# 表示文字列のヘルパー
# ---------------------------------------------------------------------------
def _disp_width(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(c) in "FWA" else 1 for c in s)


def _short_section(p: D.Person) -> str:
    if p.role == "FE":
        return D.MAKER_SHORT.get(p.section, p.section) + "FE"
    sec = p.section
    if sec.startswith("製造課 "):
        return "製造" + sec.split(" ")[1]
    if sec.startswith("設備保全課 "):
        return sec.split(" ")[1]
    return {"生産技術課": "生技", "品質保証課": "品証"}.get(sec, sec)


def _person_full(p: D.Person) -> str:
    if p.role == "FE":
        return f"{p.name}（{D.MAKER_SHORT.get(p.section, p.section)}FE）"
    return p.name


def _numbered_action(text: str) -> str:
    """「・」書きの処置は、様式の注記（実施順に番号を付ける）に合わせて番号を振り直す。"""
    lines = text.split("\n")
    if lines and all(x.startswith("・") for x in lines):
        return "\n".join(f"{n}. {x[1:]}" for n, x in enumerate(lines, 1))
    return text


def _stamps(spec: FileSpec) -> tuple[str, str, str]:
    """(承認, 確認, 作成) の押印者（姓）。未承認は空欄、確認者が作成者本人の場合は斜線「／」。"""
    inc = spec.inc
    ppl = {p.employee_id: p for p in D.people()}
    internal = [p for p in inc.assignees if p.section.startswith("設備保全課")]
    if internal:
        writer = internal[0]
    else:
        sect = "設備保全課 施設係" if D.kind_of(inc.equipment) in D.UTILITY_KINDS else (
            "設備保全課 保全1係" if D.fab_of(inc.equipment) == "Fab1" else "設備保全課 保全2係")
        writer = ppl[SECTION_HEAD[sect]]
    head = ppl[SECTION_HEAD.get(writer.section, "M10544")]
    creator = D.surname(writer)
    checker = "／" if writer.employee_id in (head.employee_id, MANAGER_ID) else D.surname(head)
    complete = inc.status in ("完了", "経過観察")
    approver = D.surname(ppl[MANAGER_ID]) if complete and spec.report_date <= date(2026, 9, 7) else ""
    return approver, checker, creator


def _writer_person(spec: FileSpec) -> D.Person:
    internal = [p for p in spec.inc.assignees if p.section.startswith("設備保全課")]
    return internal[0] if internal else next(p for p in D.people() if p.employee_id == MANAGER_ID)


class Fmt:
    """版とイレギュラーに応じて、日時・数値の「セルに入れる値」と「人が読む表示」を作る。"""

    def __init__(self, spec: FileSpec, r):
        self.spec, self.r, self.v = spec, r, spec.version
        self.dt_style = r.choice([0, 1, 3, 4, 5, 6])       # v2 の日時文字列の書き方
        self.date_style = r.choice([0, 2, 3, 4])            # v2 の日付文字列の書き方

    def z(self, s: str) -> str:
        return D.to_zenkaku(s, digits_only=True) if self.spec.zenkaku else s

    def datetime_(self, dt: datetime | None):
        """(表示, セル値, 表示形式)。"""
        if dt is None:
            return "", None, None
        y, m, d, hh, mm = dt.year, dt.month, dt.day, dt.hour, dt.minute
        if self.spec.zenkaku:
            s = self.z(f"{y}年{m}月{d}日 {hh}:{mm:02d}")
            return s, s, None
        if self.v == "v1":
            return f"{y}/{m:02d}/{d:02d} {hh:02d}:{mm:02d}", dt, "yyyy/mm/dd hh:mm"
        if self.v == "v3":
            return f"{y}/{m}/{d} {hh}:{mm:02d}", dt, "yyyy/m/d h:mm"
        s = D.fmt_datetime_variants(dt)[self.dt_style]
        return s, s, None

    def date_(self, d: date):
        y, m, dd = d.year, d.month, d.day
        if self.spec.zenkaku:
            s = self.z(f"{y}年{m}月{dd}日")
            return s, s, None
        if self.v == "v1":
            return f"{y}年{m}月{dd}日", datetime(y, m, dd), 'yyyy"年"m"月"d"日"'
        if self.v == "v3":
            return f"{y}/{m:02d}/{dd:02d}", datetime(y, m, dd), "yyyy/mm/dd"
        s = D.fmt_date(d, style=self.date_style)
        return s, s, None

    def downtime(self, inc: D.Incident):
        n = inc.downtime_min
        if inc.status == "対応中":
            s = {"v1": "継続中", "v2": "―（未復旧）", "v3": "集計中"}[self.v]
            return s, s, None
        if self.spec.zenkaku:
            s = self.z(f"{n:,}分")
            return s, s, None
        if self.v == "v1":
            return f"{n:,}分", n, '#,##0"分"'
        if self.v == "v3":
            return f"{n:,}", n, "#,##0"
        s = f"{n / 60:.1f}h" if self.r.random() < 0.5 else f"{n // 60}時間{n % 60:02d}分"
        if inc.status == "保留":
            s += "（暫定復旧まで）"
        return s, s, None

    def recovered(self, inc: D.Incident):
        if inc.status == "対応中":
            s = {"v1": "未復旧", "v2": "―", "v3": "未復旧（対応中）"}[self.v]
            return s, s, None
        if inc.completed_at is None:   # 保留: 暫定処置で運転再開した時刻
            disp, raw, fmt = self.datetime_(inc.occurred_at + timedelta(minutes=inc.downtime_min))
            s = "暫定 " + disp
            return s, s, None
        return self.datetime_(inc.completed_at)

    def hours(self, h: float) -> str:
        return self.z(f"{h:.2f}".rstrip("0").rstrip(".") + ("人時" if self.v == "v1" else "h"))

    def yen(self, n: int):
        if self.spec.zenkaku:
            s = self.z(f"{n:,}円")
            return s, s, None
        if self.v == "v1":
            return f"¥{n:,}", n, '"¥"#,##0'
        return f"{n:,}円", n, '#,##0"円"'

    def qty(self, q: int) -> str:
        return self.z(f"{q}個" if self.v == "v2" else str(q))


def _severity_checkbox(sev: str) -> str:
    return "　".join(("■" if s == sev else "□") + s for s in ("重大", "大", "中", "小"))


def _alarm_text(inc: D.Incident, v: str) -> str:
    a = inc.alarm
    if a is None:
        return {"v1": "なし", "v2": "－", "v3": f"アラームなし（{inc.detected_by}で検知）"}[v]
    return {"v1": f"{a.code}：{a.message}", "v2": f"{a.code}（{a.message}）", "v3": f"{a.code} {a.message}"}[v]


def _lots_text(inc: D.Incident) -> str:
    if not inc.lots_affected:
        return "なし"
    s = "、".join(inc.lots_affected)
    return s + (f"（廃棄 {inc.scrap_wafers}枚）" if inc.scrap_wafers else "")


def _remarks_text(inc: D.Incident) -> str:
    rm = inc.remarks
    if inc.recurrence and inc.related_incident_id and inc.related_incident_id not in rm:
        rm = (rm + "。" if rm else "") + f"{inc.related_incident_id} の再発"
    return rm


def _photo_captions(inc: D.Incident, r) -> list[str]:
    caps = list(inc.photos)
    if not caps:
        part = inc.parts_used[0][0].name if inc.parts_used else None
        pool = ["装置外観（作業前）", "アラーム画面", "処置後の状態"] + ([f"交換した{part}"] if part else [])
        caps = r.sample(pool, k=r.choice([1, 2]))
    return caps[:3]


# ---------------------------------------------------------------------------
# 写真（Pillow で簡単な設備スケッチを描く）
# ---------------------------------------------------------------------------
@lru_cache(maxsize=None)
def _pil_font(size: int):
    for name in ("msgothic.ttc", "meiryo.ttc", "YuGothM.ttc", "BIZ-UDGothicR.ttc"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _sketch(dr: ImageDraw.ImageDraw, kind: str, r, W: int, H: int) -> list[tuple]:
    """設備種別ごとの簡単な外観スケッチを描き、赤枠で囲む候補領域を返す。"""
    steel, dark, edge = (176, 182, 188), (92, 98, 104), (60, 60, 60)
    j = lambda n: r.randint(-n, n)   # noqa: E731  描くたびに少しずらす
    if kind == "CMP":
        dr.rectangle([30 + j(8), 150, 450 + j(8), 290], fill=steel, outline=edge, width=2)
        dr.ellipse([90, 120 + j(6), 300, 215], fill=dark, outline=edge, width=2)          # プラテン
        dr.ellipse([105, 130, 285, 205], fill=(70, 110, 130))                               # 研磨パッド
        dr.line([(380, 70), (240, 150)], fill=edge, width=10)                               # ヘッドアーム
        dr.ellipse([200, 135, 262, 180], fill=(150, 150, 160), outline=edge, width=2)      # 研磨ヘッド
        dr.line([(330, 110), (190, 165)], fill=(230, 230, 230), width=4)                   # スラリーノズル
        dr.ellipse([280, 185, 318, 212], fill=(120, 120, 125), outline=edge)               # コンディショナ
        return [(200, 135, 262, 180), (180, 150, 335, 175), (280, 185, 318, 212), (90, 120, 300, 215)]
    if kind in ("CVD", "ETC"):
        dr.rectangle([140, 90 + j(5), 330, 260], fill=steel, outline=edge, width=3)       # チャンバー
        dr.ellipse([140, 70, 330, 110], fill=(190, 196, 200), outline=edge, width=2)
        dr.rectangle([195, 30, 285, 72], fill=(90, 90, 100), outline=edge, width=2)       # RFマッチャー
        glow = (210, 120, 230) if kind == "ETC" else (120, 170, 230)
        dr.ellipse([205, 145, 262, 200], fill=glow, outline=edge, width=3)                 # ビューポート
        for k, y in enumerate((120, 165, 210)):                                             # ガスライン
            dr.line([(20, y), (140, y)], fill=(210, 210, 210), width=6)
            dr.rectangle([60 + k * 10, y - 9, 80 + k * 10, y + 9], fill=(60, 90, 160), outline=edge)
        dr.rectangle([200, 268, 280, 312], fill=(80, 84, 90), outline=edge, width=2)       # ポンプ
        dr.line([(240, 260), (240, 268)], fill=edge, width=8)
        return [(205, 145, 262, 200), (195, 30, 285, 72), (60, 111, 90, 129), (200, 268, 280, 312), (70, 156, 100, 174)]
    if kind == "LIT":
        dr.rectangle([50, 40, 430, 300], fill=(238, 238, 232), outline=edge, width=3)
        dr.rectangle([215, 50, 265, 160], fill=(120, 120, 130), outline=edge, width=2)     # 投影レンズ
        dr.rectangle([120, 170, 360, 280], fill=(200, 205, 210), outline=edge, width=2)    # ステージ窓
        for x in range(130, 360, 20):
            dr.line([(x, 175), (x, 275)], fill=(170, 175, 180))
        dr.ellipse([200, 195, 280, 260], fill=(90, 120, 160), outline=edge)                # ウェーハ
        dr.rectangle([70, 60, 140, 110], fill=(210, 210, 200), outline=edge)              # レチクルローダ
        return [(200, 195, 280, 260), (215, 50, 265, 160), (70, 60, 140, 110), (120, 170, 360, 280)]
    if kind == "CLN":
        for k in range(3):
            x = 40 + k * 110
            dr.rectangle([x, 140, x + 90, 280], fill=(215, 225, 230), outline=edge, width=2)
            dr.rectangle([x + 5, 190, x + 85, 275], fill=(150, 200, 225))
        dr.ellipse([370, 150, 450, 230], fill=(190, 195, 200), outline=edge, width=2)     # スピンカップ
        dr.line([(460, 90), (410, 170)], fill=edge, width=6)                               # ノズルアーム
        dr.line([(40, 110), (450, 110)], fill=(200, 200, 200), width=5)
        return [(395, 150, 450, 190), (150, 140, 240, 280), (370, 150, 450, 230), (40, 100, 200, 120)]
    if kind == "IMP":
        dr.rectangle([30, 150, 120, 250], fill=steel, outline=edge, width=2)              # イオン源
        dr.pieslice([110, 70, 290, 250], 180, 270, fill=(120, 80, 60), outline=edge)      # 分析マグネット
        dr.line([(120, 200), (200, 160), (420, 160)], fill=(100, 100, 110), width=14)     # ビームライン
        dr.rectangle([380, 110, 460, 250], fill=(200, 200, 205), outline=edge, width=2)   # エンドステーション
        dr.polygon([(60, 120), (90, 120), (75, 95)], fill=(250, 210, 0), outline=edge)    # 高電圧注意
        return [(30, 150, 120, 250), (110, 70, 200, 160), (380, 110, 460, 250), (55, 92, 95, 122)]
    if kind == "INS":
        dr.rectangle([80, 60, 400, 290], fill=(230, 232, 235), outline=edge, width=3)
        dr.rectangle([220, 70, 260, 150], fill=(60, 60, 70), outline=edge)                # 対物レンズ
        dr.ellipse([180, 170, 300, 250], fill=(95, 125, 170), outline=edge)               # ウェーハ
        dr.rectangle([300, 80, 390, 150], fill=(20, 30, 40), outline=edge)                # モニタ
        dr.ellipse([315, 88, 375, 142], outline=(90, 200, 90))
        for _ in range(12):
            x, y = r.randint(322, 368), r.randint(95, 135)
            dr.point((x, y), fill=(255, 80, 80))
        return [(220, 70, 260, 150), (180, 170, 300, 250), (300, 80, 390, 150)]
    if kind == "OHT":
        dr.rectangle([0, 40, W, 58], fill=(110, 110, 115))                                 # 走行レール
        dr.rectangle([160, 60, 320, 140], fill=(235, 235, 230), outline=edge, width=2)    # 台車
        for x in (200, 280):
            dr.line([(x, 140), (x, 225)], fill=(40, 40, 40), width=3)                     # ホイストベルト
        dr.rectangle([190, 225, 290, 295], fill=(210, 215, 220), outline=edge, width=2)   # FOUP
        dr.ellipse([165, 50, 190, 70], fill=dark)
        dr.ellipse([290, 50, 315, 70], fill=dark)
        return [(190, 140, 290, 225), (160, 45, 320, 75), (190, 215, 290, 240), (160, 60, 320, 140)]
    if kind == "AGV":
        dr.rectangle([100, 180, 380, 270], fill=(240, 200, 60), outline=edge, width=2)
        for x in (140, 340):
            dr.ellipse([x - 22, 250, x + 22, 294], fill=(40, 40, 40))
        dr.ellipse([220, 150, 260, 185], fill=(30, 30, 30))                               # LiDAR
        dr.rectangle([140, 110, 340, 178], fill=(200, 205, 210), outline=edge)            # マガジン
        return [(118, 250, 362, 294), (220, 150, 260, 185), (100, 180, 380, 270)]
    if kind == "ROB":
        dr.rectangle([190, 220, 290, 300], fill=steel, outline=edge, width=2)
        dr.line([(240, 220), (170, 150), (290, 110)], fill=(90, 95, 100), width=18)
        dr.rectangle([285, 95, 350, 125], fill=(210, 210, 215), outline=edge)             # ハンド
        for x in (40, 380):
            dr.rectangle([x, 150, x + 70, 260], fill=(225, 228, 232), outline=edge, width=2)
        return [(285, 95, 350, 125), (150, 130, 200, 170), (380, 150, 450, 260)]
    if kind == "STK":
        for c in range(5):
            for rr in range(4):
                x, y = 40 + c * 85, 50 + rr * 62
                dr.rectangle([x, y, x + 70, y + 50], fill=(215, 218, 222), outline=edge)
        dr.rectangle([225, 20, 245, 320], fill=(90, 90, 95))                               # クレーン
        return [(225, 120, 245, 200), (125, 112, 195, 162), (40, 50, 110, 100)]
    if kind == "EXH":
        dr.ellipse([120, 60, 360, 300], fill=steel, outline=edge, width=3)                # ファンケーシング
        for a in range(0, 360, 60):
            x = 240 + 90 * math.cos(math.radians(a + j(10)))
            y = 180 + 90 * math.sin(math.radians(a))
            dr.line([(240, 180), (x, y)], fill=dark, width=12)
        dr.rectangle([370, 150, 460, 220], fill=(70, 90, 120), outline=edge)              # モーター
        return [(215, 155, 265, 205), (370, 150, 460, 220), (330, 170, 380, 200)]
    # UPW / PCW / VAC / SCR: ポンプと配管
    dr.rectangle([60, 180, 200, 260], fill=(60, 100, 150), outline=edge, width=2)         # モーター
    dr.ellipse([190, 160, 300, 280], fill=steel, outline=edge, width=3)                    # ポンプケーシング
    dr.line([(245, 160), (245, 60), (460, 60)], fill=(200, 200, 200), width=16)           # 吐出配管
    dr.rectangle([330, 45, 350, 75], fill=(180, 40, 40), outline=edge)                     # バルブ
    dr.ellipse([380, 85, 430, 135], fill=(245, 245, 245), outline=edge, width=3)          # 圧力計
    dr.line([(405, 110), (420, 95)], fill=(200, 0, 0), width=2)
    return [(380, 85, 430, 135), (190, 160, 300, 280), (330, 45, 350, 75), (195, 250, 245, 285)]


def _photo_png(kind: str, caption: str, eid: str, shot: datetime, seed: str, size=(480, 360)) -> bytes:
    r = D.rng(seed)
    W, H = size
    base = (240, 228, 160) if kind == "LIT" else ((198, 203, 206) if kind in D.UTILITY_KINDS else (222, 226, 230))
    img = PILImage.new("RGB", (W, H), base)
    dr = ImageDraw.Draw(img)
    phase = r.random() * 6.28
    for y in range(H):   # 照明ムラ
        k = 1.0 - 0.16 * (y / H) + 0.03 * math.sin(y / 27 + phase)
        dr.line([(0, y), (W, y)], fill=tuple(max(0, min(255, int(c * k))) for c in base))
    floor = int(H * 0.8)
    dr.rectangle([0, floor, W, H], fill=(150, 152, 150))
    for x in range(0, W, 16):   # グレーチング床
        dr.line([(x, floor), (x, H)], fill=(125, 127, 125))
    boxes = _sketch(dr, kind, r, W, H)
    x0, y0, x1, y1 = r.choice(boxes)
    pad = r.randint(5, 12)
    dr.rectangle([x0 - pad, y0 - pad, x1 + pad, y1 + pad], outline=(230, 20, 20), width=4)
    for _ in range(700):   # 粒状ノイズ
        g = r.randint(90, 250)
        dr.point((r.randrange(W), r.randrange(H)), fill=(g, g, g))
    f14, f17 = _pil_font(14), _pil_font(17)
    dr.text((8, 6), shot.strftime("%Y/%m/%d %H:%M"), fill=(255, 150, 40), font=f14)
    tw = dr.textlength(eid, font=f14)
    dr.rectangle([W - tw - 16, 4, W - 4, 24], fill=(40, 40, 40))
    dr.text((W - tw - 10, 6), eid, fill=(255, 255, 255), font=f14)
    dr.rectangle([0, H - 32, W, H], fill=(28, 28, 28))
    cap = caption
    while dr.textlength(cap, font=f17) > W - 20 and len(cap) > 4:
        cap = cap[:-2] + "…"
    dr.text((10, H - 26), cap, fill=(255, 255, 255), font=f17)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _img_rows(h_px: int, row_pt: float = 15.0) -> int:
    """高さ h_px の画像を載せるのに必要な行数（1行 row_pt ポイント、上下の余白込み）。"""
    return math.ceil((h_px * 0.75 + 10) / row_pt)


def _col_px(width_units: float) -> int:
    """列幅（文字数単位）→ ピクセル（標準フォント前提の近似）。"""
    return int(width_units * 7 + 5)


# ---------------------------------------------------------------------------
# 帳票書き込みヘルパー
# ---------------------------------------------------------------------------
class Form:
    """1シート分の帳票を組み立てる。ラベル・値を置くと正解データ（Record）にも記録する。"""

    def __init__(self, ws, widths: dict[str, float], fill: str, rec: Record | None):
        self.ws, self.rec = ws, rec
        self.fill = PatternFill("solid", fgColor=fill)
        self.widths = widths
        self.break_rows: set[int] = set()
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    # --- 基本 ---
    def width_of(self, ref: str) -> float:
        a, b = (ref.split(":") + [ref])[:2]
        c1 = column_index_from_string("".join(ch for ch in a if ch.isalpha()))
        c2 = column_index_from_string("".join(ch for ch in b if ch.isalpha()))
        return sum(self.widths.get(get_column_letter(c), 8.43) for c in range(c1, c2 + 1))

    @staticmethod
    def _rows(ref: str) -> tuple[int, int]:
        a, b = (ref.split(":") + [ref])[:2]
        return int("".join(ch for ch in a if ch.isdigit())), int("".join(ch for ch in b if ch.isdigit()))

    def put(self, ref: str, value, *, size: float = 10, bold: bool = False, color: str | None = None, fill: PatternFill | None = None,
            h: str = "left", v: str = "center", wrap: bool = True, border: Border | None = BOX, fmt: str | None = None,
            rotation: int = 0):
        top_left = ref.split(":")[0]
        c = self.ws[top_left]
        c.value = value
        c.font = Font(name=FONT_NAME, size=size, bold=bold, color=color)
        # 日付・数値は折り返せず「####」になるため、縮小して全体を表示する
        numeric = isinstance(value, (datetime, int, float)) and not isinstance(value, bool)
        c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap and not numeric, shrink_to_fit=numeric,
                                text_rotation=rotation)
        if fill is not None:
            c.fill = fill
        if border is not None:
            c.border = border
        if fmt:
            c.number_format = fmt
        if ":" in ref:
            self.ws.merge_cells(ref)
        return c

    def label(self, ref: str, text: str, key: str | None = None, *, vertical: bool = False, h: str = "center", size: float = 10,
              bold: bool = False):
        self.put(ref, text, size=size, bold=bold, fill=self.fill, h=h, rotation=255 if vertical else 0)
        if key and self.rec is not None:
            self.rec.labels[key] = text

    def value(self, ref: str, display: str, key: str | None = None, *, raw=None, fmt: str | None = None, h: str = "left",
              v: str = "center", size: float = 10, color: str | None = None, bold: bool = False):
        cell_value = raw if raw is not None else (display if display != "" else None)
        self.put(ref, cell_value, size=size, h=h, v=v, fmt=fmt, color=color, bold=bold)
        if key and self.rec is not None:
            self.rec.values[key] = display

    def fit(self, ref: str, text: str, *, min_row_pt: float = 15.0, line_pt: float = 13.5, min_total_pt: float = 0):
        """結合範囲に文章が収まるよう行の高さを決める（全角2・半角1で幅を見積もる）。"""
        cap = max(8.0, self.width_of(ref) * 0.95 - 1)
        lines = sum(max(1, math.ceil(_disp_width(p) / cap)) for p in (text or "").split("\n"))
        r1, r2 = self._rows(ref)
        n = r2 - r1 + 1
        total = max(lines * line_pt + 6, n * min_row_pt, min_total_pt)
        for rr in range(r1, r2 + 1):
            self.ws.row_dimensions[rr].height = round(total / n, 1)

    def height(self, row: int, pt: float):
        self.ws.row_dimensions[row].height = pt

    def breakable(self, row: int):
        """この行の上で改ページしてよい（ブロックの先頭）ことを記録する。"""
        self.break_rows.add(row)

    def image(self, png: bytes, col: str, row: int, w: int, h: int, x_off: int = 6, y_off: int = 4):
        img = XLImage(BytesIO(png))
        img.width, img.height = w, h
        marker = AnchorMarker(col=column_index_from_string(col) - 1, colOff=pixels_to_EMU(x_off), row=row - 1, rowOff=pixels_to_EMU(y_off))
        img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
        self.ws.add_image(img)

    def page_setup(self, area: str, *, gridlines: bool = False):
        """A4縦・横1ページ合わせの印刷設定。縦はブロックの切れ目で手動改ページを入れる。"""
        ws = self.ws
        ws.print_area = area
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.5
        ws.page_margins.top, ws.page_margins.bottom = 0.6, 0.6
        ws.print_options.horizontalCentered = True
        ws.sheet_view.showGridLines = gridlines
        ws.oddFooter.center.text = "&P / &N"
        ws.oddFooter.center.size = 8

        # 改ページ位置: 印刷倍率（横幅合わせ）を見積もり、ページ高さを超える手前のブロック先頭で切る
        c1, c2 = (x.rstrip("0123456789") for x in area.split(":"))
        r_last = self._rows(area)[1]
        sheet_w_pt = sum(_col_px(self.widths.get(get_column_letter(c), 8.43))
                         for c in range(column_index_from_string(c1), column_index_from_string(c2) + 1)) * 0.75
        scale = min(1.0, (8.27 - 1.0) * 72 / sheet_w_pt)
        page_pt = 735 / scale
        acc, last_break, candidate = 0.0, 1, None
        for rr in range(1, r_last + 1):
            hgt = ws.row_dimensions[rr].height or 15.0
            if rr in self.break_rows and rr > last_break:
                candidate = rr
            if acc + hgt > page_pt and candidate is not None:
                ws.row_breaks.append(Break(id=candidate - 1))
                acc = sum((ws.row_dimensions[x].height or 15.0) for x in range(candidate, rr))
                last_break, candidate = candidate, None
            acc += hgt


# ---------------------------------------------------------------------------
# 共通の値準備
# ---------------------------------------------------------------------------
@dataclass
class Ctx:
    spec: FileSpec
    r: object
    fmt: Fmt
    rec: Record
    wb: Workbook
    approver: str = ""
    checker: str = ""
    creator: str = ""
    captions: list = field(default_factory=list)

    @property
    def inc(self) -> D.Incident:
        return self.spec.inc

    def blank(self, key: str, display: str, raw=None, fmt=None):
        """必須項目の空欄イレギュラー: 対象キーなら値を消す。"""
        if self.spec.blank_field == key:
            return "", None, None
        return display, raw, fmt


def _photo_shot_time(inc: D.Incident, n: int) -> datetime:
    return inc.response_started_at + timedelta(minutes=20 + n * 17)


def _photo_bytes(ctx: Ctx, caption: str, n: int, tag: str, size=(480, 360)) -> bytes:
    inc = ctx.inc
    return _photo_png(D.kind_of(inc.equipment), caption, inc.equipment.equipment_id, _photo_shot_time(inc, n),
                      f"f1-photo:{inc.incident_id}:{tag}:{n}", size=size)


def _text_or_blank(ctx: Ctx, key: str, text: str) -> str:
    return "" if ctx.spec.blank_field == key else text


# ---------------------------------------------------------------------------
# v1（Rev.1）: 左ラベル・右値の2列組み＋見出し下の結合テキスト
# ---------------------------------------------------------------------------
def _build_v1(ctx: Ctx):
    inc, eq, fm, spec = ctx.inc, ctx.inc.equipment, ctx.fmt, ctx.spec
    ws = ctx.wb.active
    ws.title = "修理報告書"
    ctx.rec.main_sheet = ws.title
    f = Form(ws, {"A": 13, "B": 11, "C": 11, "D": 11, "E": 13, "F": 11, "G": 11, "H": 11}, "DDEBF7", ctx.rec)
    head_fill = PatternFill("solid", fgColor="BDD7EE")
    rev = REV_INFO["v1"][0]
    interim = inc.status in ("対応中", "保留")

    f.put("G1:H1", f"{FORM_NO}({rev[-1]})", size=8, h="right", border=None)
    f.height(1, 14)
    f.put("A2:H2", "設備修理報告書" + ("（中間報告）" if interim else ""), size=18, bold=True, h="center",
          border=Border(bottom=MEDIUM))
    for c in "BCDEFGH":
        ws[f"{c}2"].border = Border(bottom=MEDIUM)
    f.height(2, 34)
    f.put("A3:D3", "製造部　設備保全課", size=9, border=None)
    f.height(3, 16)

    # 報告番号・報告日・管理No. と 承認欄
    f.label("A4", "報告番号", "report_id")
    f.value("B4:D4", spec.report_id, "report_id")
    disp, raw, nf = ctx.blank("report_date", *fm.date_(spec.report_date))
    f.label("A5", "報告日", "report_date")
    f.value("B5:D5", disp, "report_date", raw=raw, fmt=nf)
    f.label("A6", "管理No.", "incident_id")
    f.value("B6:D6", inc.incident_id, "incident_id")
    for col, text, key, who in (("F", "承認", "approver", ctx.approver), ("G", "確認", "checker", ctx.checker),
                                ("H", "作成", "creator", ctx.creator)):
        f.label(f"{col}4", text, key)
        f.value(f"{col}5:{col}6", who, key, h="center", size=12, bold=True, color=STAMP_RED)
    for rr, pt in ((4, 20), (5, 22), (6, 22), (7, 8)):
        f.height(rr, pt)

    # 設備・日時など（左右2列）
    rows = [
        (("設備番号", "equipment_id", eq.equipment_id, None, None), ("設備名", "equipment_name", eq.name, None, None)),
        (("ライン", "line", eq.line, None, None), ("メーカー・型式", "maker", f"{eq.maker}　{eq.model}", None, None)),
        (("発生日時", "occurred_at", *fm.datetime_(inc.occurred_at)), ("復旧日時", "recovered_at", *ctx.blank("recovered_at", *fm.recovered(inc)))),
        (("停止時間", "downtime_min", *ctx.blank("downtime_min", *fm.downtime(inc))), ("重要度", "severity", inc.severity, _severity_checkbox(inc.severity), None)),
        (("報告者", "reporter", inc.reporter.name, None, None),
         ("担当者", "assignee", _text_or_blank(ctx, "assignee", "、".join(_person_full(p) for p in inc.assignees)), None, None)),
        (("作業工数", "work_hours", fm.hours(inc.work_hours), None, None), ("修理費用", "cost", *fm.yen(inc.cost_yen))),
    ]
    r0 = 8
    for n, (left, right) in enumerate(rows):
        rr = r0 + n
        for (lab, key, disp, raw, nf), (lc, vc) in zip((left, right), (("A", "B{r}:D{r}"), ("E", "F{r}:H{r}"))):
            f.label(f"{lc}{rr}", lab, key)
            f.value(vc.format(r=rr), disp, key, raw=raw, fmt=nf)
        f.fit(f"F{rr}:H{rr}", right[2], min_row_pt=21)
    rr = r0 + len(rows)
    f.label(f"A{rr}", "アラーム", "alarm")
    alarm = _alarm_text(inc, "v1")
    f.value(f"B{rr}:H{rr}", alarm, "alarm")
    f.fit(f"B{rr}:H{rr}", alarm, min_row_pt=21)
    f.height(rr + 1, 8)
    row = rr + 2

    def section(title: str, key: str, text: str, n_rows: int) -> None:
        nonlocal row
        f.breakable(row)
        f.put(f"A{row}:H{row}", title, bold=True, fill=head_fill, h="left")
        ctx.rec.labels[key] = title
        f.height(row, 19)
        ref = f"A{row + 1}:H{row + n_rows}"
        f.value(ref, text, key, v="top")
        f.fit(ref, text)
        row += n_rows + 1

    section("故障内容（発生状況）", "symptom", inc.symptom, 3)
    section("原因調査", "investigation", inc.investigation, 3)
    section("原因", "cause", _text_or_blank(ctx, "cause", inc.cause), 2)
    section("修理内容（実施順に記入）", "action", _numbered_action(inc.action), 4)

    # 使用部品（小さな表）
    f.breakable(row)
    f.put(f"A{row}:H{row}", "使用部品", bold=True, fill=head_fill, h="left")
    ctx.rec.labels["parts"] = "使用部品"
    f.height(row, 19)
    row += 1
    for ref, text, key in ((f"A{row}", "No.", None), (f"B{row}:C{row}", "品番", "part_no"), (f"D{row}:F{row}", "品名", "name"),
                           (f"G{row}", "数量", "qty"), (f"H{row}", "金額(円)", "amount")):
        f.put(ref, text, size=9, fill=f.fill, h="center")
        if key:
            ctx.rec.labels[f"parts.{key}"] = text
    f.height(row, 17)
    parts = []
    for n in range(max(3, len(inc.parts_used))):
        row += 1
        f.height(row, 18)
        if n < len(inc.parts_used):
            p, q = inc.parts_used[n]
            amount = p.unit_price_yen * q
            f.put(f"A{row}", n + 1, size=9, h="center")
            f.put(f"B{row}:C{row}", p.part_no, size=9)
            f.put(f"D{row}:F{row}", p.name, size=9)
            f.fit(f"D{row}:F{row}", p.name, min_row_pt=18)
            f.put(f"G{row}", fm.qty(q), size=9, h="center")
            f.put(f"H{row}", amount, size=9, h="right", fmt="#,##0")
            parts.append({"part_no": p.part_no, "name": p.name, "qty": fm.qty(q), "amount": f"{amount:,}"})
        else:
            for ref in (f"A{row}", f"B{row}:C{row}", f"D{row}:F{row}", f"G{row}", f"H{row}"):
                f.put(ref, "－" if (n == 0 and ref.startswith("D")) else None, size=9, h="center" if n == 0 else "left")
    ctx.rec.values["parts"] = parts
    row += 1

    section("修理結果", "result", _text_or_blank(ctx, "result", inc.result), 2)
    section("再発防止策", "prevention", inc.prevention, 2)

    # 写真（2枚ずつ横並び）
    f.breakable(row)
    f.put(f"A{row}:H{row}", "写真", bold=True, fill=head_fill, h="left")
    ctx.rec.labels["photos"] = "写真"
    f.height(row, 19)
    row += 1
    row = _photo_grid(f, ctx, row, [("A", "D"), ("E", "H")], (280, 210), "main", border_cols=("A", "H"), mark_first=False)
    section("備考", "remarks", _remarks_text(inc), 2)

    f.page_setup(f"A1:H{row - 1}")
    return f


def _photo_grid(f: Form, ctx: Ctx, row: int, slots: list[tuple[str, str]], size: tuple[int, int], tag: str,
                border_cols: tuple[str, str], captions: list[str] | None = None, prefix: str = "写真", mark_first: bool = True, split_ok: bool = True) -> int:
    """写真を slots（列範囲）に横並びで貼り、下にキャプションを書く。次の行番号を返す。

    mark_first: 先頭の写真行の上で改ページしてよいか / split_ok: 写真の段と段の間で改ページしてよいか
    （左に縦結合ラベルがある版では、ラベルがページで分断されないよう段の間では切らない）
    """
    caps = ctx.captions if captions is None else captions
    w, h = size
    rows_img = _img_rows(h)
    out = []
    for k in range(0, len(caps), len(slots)):
        chunk = caps[k:k + len(slots)]
        if (k > 0 and split_ok) or (k == 0 and mark_first):
            f.breakable(row)
        for rr in range(row, row + rows_img):
            f.height(rr, 15)
        # 写真枠（外周だけ罫線）
        c1, c2 = column_index_from_string(border_cols[0]), column_index_from_string(border_cols[1])
        for rr in range(row, row + rows_img):
            for cc in range(c1, c2 + 1):
                ws_cell = f.ws.cell(rr, cc)
                ws_cell.border = Border(left=THIN if cc == c1 else None, right=THIN if cc == c2 else None,
                                        top=THIN if rr == row else None)
        cap_row = row + rows_img
        for n, cap in enumerate(chunk):
            a, b = slots[n]
            slot_px = int(f.width_of(f"{a}1:{b}1") * 7)
            x_off = max(4, (slot_px - w) // 2)
            num = k + n + 1
            f.image(_photo_bytes(ctx, cap, num, tag), a, row, w, h, x_off=x_off, y_off=6)
            text = f"{prefix}{num}：{cap}"
            f.put(f"{a}{cap_row}:{b}{cap_row}", text, size=9, h="center")
            out.append(text)
        for n in range(len(chunk), len(slots)):
            a, b = slots[n]
            f.put(f"{a}{cap_row}:{b}{cap_row}", None, size=9)
        f.height(cap_row, 18)
        row = cap_row + 1
    if captions is None:
        ctx.rec.values["photos"] = out
        ctx.rec.images[f.ws.title] = ctx.rec.images.get(f.ws.title, 0) + len(caps)
    return row


# ---------------------------------------------------------------------------
# v2（Rev.2）: 列ズレ・項目名ゆれ・1セル内ラベル・承認欄は下
# ---------------------------------------------------------------------------
def _build_v2(ctx: Ctx):
    inc, eq, fm, spec, r = ctx.inc, ctx.inc.equipment, ctx.fmt, ctx.spec, ctx.r
    ws = ctx.wb.active
    ws.title = "Sheet1" if spec.default_sheet_name else r.choice(["報告書", "修理報告書", f"{inc.occurred_at:%m%d}"])
    ctx.rec.main_sheet = ws.title
    if spec.default_sheet_name:
        ctx.rec.irregularities.append("シート名が既定の「Sheet1」のまま")
    widths = {"A": 2, "B": 8, "C": 8, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 11, "L": 2}
    f = Form(ws, widths, "E2EFDA", ctx.rec)
    lab = {k: r.choice(v) for k, v in V2_SYNONYMS.items()}
    if lab["equipment_id"] == "装置No.":
        lab["equipment_name"] = "装置名"
    inline = set(r.sample(["incident_id", "equipment_id", "equipment_name"], k=r.choice([1, 2, 2, 3])))
    sep = r.choice(["：", ":", "　："])
    rev = REV_INFO["v2"][0]
    interim = inc.status in ("対応中", "保留")

    f.put("B1:G1", r.choice(["修理報告書", "設備修理報告書"]), size=16, bold=True, border=None)
    f.put("I1:K1", f"{FORM_NO} {rev}", size=8, h="right", border=None)
    f.height(1, 28)
    f.put("B2:E2", "製造部 設備保全課", size=9, border=None)
    # 正解値はチェックボックスで■の付いた選択肢の文字（F4 の「■定期点検」→「定期点検」と同じ扱い）
    rtype = "中間" if interim else "完了"
    f.put("H2:K2", f"報告区分{sep}" + ("□完了　■中間" if interim else "■完了　□中間"), size=9, border=None, h="right")
    ctx.rec.labels["report_type"] = "報告区分"
    ctx.rec.values["report_type"] = rtype
    f.height(2, 18)
    f.height(3, 6)

    def pair(row: int, key: str, disp: str, raw=None, nf=None, *, left: bool = True, wide: bool = False):
        """左（B:C→D:F）または右（H→I:K）にラベルと値を置く。inline 指定なら1セルに「ラベル：値」。"""
        text = lab[key]
        if left:
            lab_ref, val_ref, whole = f"B{row}:C{row}", (f"D{row}:K{row}" if wide else f"D{row}:F{row}"), f"B{row}:F{row}"
        else:
            lab_ref, val_ref, whole = f"H{row}", f"I{row}:K{row}", f"H{row}:K{row}"
        if key in inline:
            f.put(whole, f"{text}{sep}{disp}")
            ctx.rec.labels[key] = text
            ctx.rec.values[key] = disp
            return
        f.label(lab_ref, text, key)
        f.value(val_ref, disp, key, raw=raw, fmt=nf)

    pair(4, "report_id", spec.report_id)
    pair(4, "report_date", *ctx.blank("report_date", *fm.date_(spec.report_date)), left=False)
    pair(5, "incident_id", inc.incident_id)
    rep = _text_or_blank(ctx, "reporter", f"{D.surname(inc.reporter)}（{_short_section(inc.reporter)}）")
    pair(5, "reporter", rep, left=False)
    pair(6, "equipment_id", eq.equipment_id)
    pair(6, "equipment_name", eq.name, left=False)
    line = f"{eq.line} / {eq.process}" if lab["line"] == "ライン/工程" else eq.line
    pair(7, "line", line)
    asg = _text_or_blank(ctx, "assignee", "、".join(
        f"{D.surname(p)}({D.MAKER_SHORT.get(p.section, p.section)})" if p.role == "FE" else D.surname(p) for p in inc.assignees))
    pair(7, "assignee", asg, left=False)
    pair(8, "occurred_at", *fm.datetime_(inc.occurred_at))
    pair(8, "recovered_at", *fm.recovered(inc), left=False)
    pair(9, "downtime_min", *ctx.blank("downtime_min", *fm.downtime(inc)))
    pair(9, "work_hours", fm.hours(inc.work_hours), left=False)
    pair(10, "category", inc.category)
    pair(10, "severity", inc.severity, left=False)
    alarm = _alarm_text(inc, "v2")
    pair(11, "alarm", alarm, wide=True)
    for rr in range(4, 12):
        f.height(rr, 21)
    f.fit("I6:K6", eq.name, min_row_pt=21)
    f.height(12, 8)
    row = 13

    def block(key: str, text: str, n_rows: int) -> None:
        nonlocal row
        f.breakable(row)
        ref_l, ref_v = f"B{row}:C{row + n_rows - 1}", f"D{row}:K{row + n_rows - 1}"
        f.label(ref_l, lab[key], key)
        f.value(ref_v, text, key, v="top")
        f.fit(ref_v, text, min_row_pt=16)
        row += n_rows

    block("symptom", inc.symptom, 2)
    block("first_response", inc.first_response, 2)
    block("investigation", inc.investigation, 3)
    block("cause", _text_or_blank(ctx, "cause", inc.cause), 2)
    block("action", _numbered_action(inc.action), 4)

    # 使用部品（列順が v1 と違う: 部品名 → 型式・品番 → 数量 → 単価）
    n_parts = max(2, len(inc.parts_used))
    f.breakable(row)
    f.label(f"B{row}:C{row + n_parts}", lab["parts"], "parts")
    for ref, text, key in ((f"D{row}:F{row}", "部品名", "name"), (f"G{row}:H{row}", "型式・品番", "part_no"),
                           (f"I{row}", "数量", "qty"), (f"J{row}:K{row}", "単価", "unit_price")):
        f.put(ref, text, size=9, fill=f.fill, h="center")
        ctx.rec.labels[f"parts.{key}"] = text
    f.height(row, 17)
    parts = []
    for n in range(n_parts):
        rr = row + 1 + n
        f.height(rr, 18)
        if n < len(inc.parts_used):
            p, q = inc.parts_used[n]
            price = fm.z(f"@{p.unit_price_yen:,}")
            f.put(f"D{rr}:F{rr}", p.name, size=9)
            f.fit(f"D{rr}:F{rr}", p.name, min_row_pt=18)
            f.put(f"G{rr}:H{rr}", p.part_no, size=9)
            f.put(f"I{rr}", fm.qty(q), size=9, h="center")
            f.put(f"J{rr}:K{rr}", price, size=9, h="right")
            parts.append({"name": p.name, "part_no": p.part_no, "qty": fm.qty(q), "unit_price": price})
        else:
            for ref in (f"D{rr}:F{rr}", f"G{rr}:H{rr}", f"I{rr}", f"J{rr}:K{rr}"):
                f.put(ref, "交換部品なし" if (n == 0 and ref.startswith("D")) else None, size=9)
    ctx.rec.values["parts"] = parts
    row += n_parts + 1

    block("result", _text_or_blank(ctx, "result", inc.result), 2)
    block("prevention", inc.prevention, 2)
    lots = _lots_text(inc)
    f.breakable(row)
    f.label(f"B{row}:C{row}", lab["lots"], "lots")
    f.value(f"D{row}:K{row}", lots, "lots")
    f.fit(f"D{row}:K{row}", lots, min_row_pt=21)
    row += 1

    # 写真
    n_img_rows = 2 if len(ctx.captions) > 2 else 1
    if spec.photos_outside:
        # 枠には「右側に貼付」とだけ書き、写真は印刷範囲外（M列〜）に縦に貼る
        f.breakable(row)
        f.label(f"B{row}:C{row}", lab["photos"], "photos")
        f.put(f"D{row}:K{row}", "別紙のとおり（右側に貼付）", size=9)
        f.height(row, 21)
        ws.column_dimensions["M"].width = 44
        pr = 4
        caps = []
        for n, cap in enumerate(ctx.captions, 1):
            text = f"写真{n}　{cap}"
            ws.cell(pr, 13, text).font = Font(name=FONT_NAME, size=9, bold=True)
            f.image(_photo_bytes(ctx, cap, n, "main"), "M", pr + 1, 280, 210, x_off=4, y_off=2)
            for rr in range(pr + 1, pr + 16):
                if ws.row_dimensions[rr].height is None:
                    f.height(rr, 15)
            caps.append(text)
            pr += 17
        ctx.rec.values["photos"] = caps
        ctx.rec.images[ws.title] = len(ctx.captions)
        ctx.rec.irregularities.append("写真を印刷範囲外（M列）に貼付")
        row += 1
    else:
        start = row
        rows_needed = n_img_rows * (_img_rows(203) + 1)
        f.label(f"B{row}:C{row + rows_needed - 1}", lab["photos"], "photos")
        row = _photo_grid(f, ctx, row, [("D", "G"), ("H", "K")], (270, 203), "main", border_cols=("D", "K"), split_ok=False)
        assert row == start + rows_needed
    block("remarks", _remarks_text(inc), 2)

    # 承認欄（下部右）
    f.breakable(row)
    row += 1
    f.height(row - 1, 10)
    for col, text, key, who in (("I", "承認", "approver", ctx.approver), ("J", "確認", "checker", ctx.checker),
                                ("K", lab["creator"], "creator", ctx.creator)):
        f.label(f"{col}{row}", text, key)
        f.value(f"{col}{row + 1}:{col}{row + 2}", who, key, h="center", size=12, bold=True, color=STAMP_RED)
    f.height(row, 18)
    f.height(row + 1, 20)
    f.height(row + 2, 20)
    row += 3
    f.page_setup(f"A1:L{row}", gridlines=True)
    return f


# ---------------------------------------------------------------------------
# v3（Rev.3）: 見出し行＋値行の表形式、縦結合ラベル
# ---------------------------------------------------------------------------
def _build_v3(ctx: Ctx):
    inc, eq, fm, spec = ctx.inc, ctx.inc.equipment, ctx.fmt, ctx.spec
    ws = ctx.wb.active
    ws.title = ctx.r.choice(["設備修理報告書", "報告書(Rev3)"])
    ctx.rec.main_sheet = ws.title
    widths = {"A": 5.5, **{get_column_letter(c): 8.5 for c in range(2, 13)}}
    f = Form(ws, widths, "F2F2F2", ctx.rec)
    head = PatternFill("solid", fgColor="D9D9D9")
    f.fill = head
    rev, rev_date = REV_INFO["v3"]
    interim = inc.status in ("対応中", "保留")

    f.put("A1:H2", "設備修理報告書" + ("（中間）" if interim else ""), size=18, bold=True, h="left", border=None)
    for col, text, key, who in (("J", "承認", "approver", ctx.approver), ("K", "確認", "checker", ctx.checker),
                                ("L", "作成", "creator", ctx.creator)):
        f.label(f"{col}1", text, key, size=9)
        f.value(f"{col}2:{col}3", who, key, h="center", size=12, bold=True, color=STAMP_RED)
    f.height(1, 16)
    f.height(2, 22)
    f.height(3, 22)
    f.put("A3:H3", f"{FORM_NO} {rev}（{rev_date}）　製造部 設備保全課", size=8, border=None)
    f.height(4, 8)

    # 表形式ヘッダ（見出し行・値行）
    cols6 = [("A", "B"), ("C", "D"), ("E", "F"), ("G", "H"), ("I", "J"), ("K", "L")]
    rd = ctx.blank("report_date", *fm.date_(spec.report_date))
    asg = _text_or_blank(ctx, "assignee", "、".join(_person_full(p) for p in inc.assignees))
    header_rows = [
        [("報告番号", "report_id", (spec.report_id, None, None)), ("報告日", "report_date", rd),
         ("設備番号", "equipment_id", (eq.equipment_id, None, None)), ("設備名", "equipment_name", (eq.name, None, None)),
         ("ライン", "line", (eq.line, None, None)), ("重要度", "severity", (inc.severity, None, None))],
        [("発生日時", "occurred_at", fm.datetime_(inc.occurred_at)), ("復旧日時", "recovered_at", ctx.blank("recovered_at", *fm.recovered(inc))),
         ("停止時間(分)", "downtime_min", ctx.blank("downtime_min", *fm.downtime(inc))), ("報告者", "reporter", (inc.reporter.name, None, None)),
         ("担当者", "assignee", (asg, None, None)), ("故障区分", "category", (inc.category, None, None))],
    ]
    row = 5
    for items in header_rows:
        for (a, b), (text, key, (disp, raw, nf)) in zip(cols6, items):
            f.label(f"{a}{row}:{b}{row}", text, key, size=9)
            f.value(f"{a}{row + 1}:{b}{row + 1}", disp, key, raw=raw, fmt=nf, h="center")
        f.height(row, 17)
        longest = max((it[2][0] for it in items), key=_disp_width)
        f.fit(f"G{row + 1}:H{row + 1}", longest, min_row_pt=26)
        row += 2
    alarm = _alarm_text(inc, "v3")
    for ref, text, key in ((f"A{row}:B{row}", "TR番号", "incident_id"), (f"C{row}:H{row}", "アラーム", "alarm"),
                           (f"I{row}:J{row}", "原因区分", "cause_category"), (f"K{row}:L{row}", "対応状況", "status")):
        f.label(ref, text, key, size=9)
    f.value(f"A{row + 1}:B{row + 1}", inc.incident_id, "incident_id", h="center")
    f.value(f"C{row + 1}:H{row + 1}", alarm, "alarm")
    cc = "" if inc.status == "対応中" else inc.cause_category
    f.value(f"I{row + 1}:J{row + 1}", cc, "cause_category", h="center")
    f.value(f"K{row + 1}:L{row + 1}", inc.status, "status", h="center")
    f.height(row, 17)
    f.fit(f"C{row + 1}:H{row + 1}", alarm, min_row_pt=24)
    row += 2
    f.height(row, 8)
    row += 1

    def block(title: str, key: str, text: str, n_rows: int) -> None:
        nonlocal row
        f.breakable(row)
        ref_l, ref_v = f"A{row}:A{row + n_rows - 1}", f"B{row}:L{row + n_rows - 1}"
        f.label(ref_l, title, key, vertical=True, size=9)
        f.value(ref_v, text, key, v="top")
        f.fit(ref_v, text, min_row_pt=16, min_total_pt=len(title) * 12 + 8)
        row += n_rows

    block("故障内容", "symptom", inc.symptom, 3)
    block("原因調査", "investigation", inc.investigation, 3)
    block("原因", "cause", _text_or_blank(ctx, "cause", inc.cause), 2)
    why = "\n".join(f"なぜ{n}：{w}" for n, w in enumerate(inc.why_why, 1))
    block("なぜなぜ分析", "why_why", why, 3)
    block("修理内容", "action", _numbered_action(inc.action), 4)

    # 使用部品（縦結合ラベル＋表＋部品費計）
    n_parts = max(2, len(inc.parts_used))
    f.breakable(row)
    f.label(f"A{row}:A{row + n_parts + 1}", "使用部品", "parts", vertical=True, size=9)
    for ref, text, key in ((f"B{row}:C{row}", "品番", "part_no"), (f"D{row}:G{row}", "品名", "name"),
                           (f"H{row}:I{row}", "メーカー", "maker"), (f"J{row}", "数量", "qty"), (f"K{row}:L{row}", "金額(円)", "amount")):
        f.put(ref, text, size=9, fill=PatternFill("solid", fgColor="F2F2F2"), h="center")
        ctx.rec.labels[f"parts.{key}"] = text
    f.height(row, 17)
    parts, total = [], 0
    for n in range(n_parts):
        rr = row + 1 + n
        f.height(rr, 18)
        if n < len(inc.parts_used):
            p, q = inc.parts_used[n]
            amount = p.unit_price_yen * q
            total += amount
            f.put(f"B{rr}:C{rr}", p.part_no, size=9)
            f.put(f"D{rr}:G{rr}", p.name, size=9)
            f.put(f"H{rr}:I{rr}", D.MAKER_SHORT.get(p.maker, p.maker), size=9)
            f.fit(f"D{rr}:G{rr}", p.name, min_row_pt=18)
            f.put(f"J{rr}", fm.qty(q), size=9, h="center")
            f.put(f"K{rr}:L{rr}", amount, size=9, h="right", fmt="#,##0")
            parts.append({"part_no": p.part_no, "name": p.name, "maker": D.MAKER_SHORT.get(p.maker, p.maker),
                          "qty": fm.qty(q), "amount": f"{amount:,}"})
        else:
            for ref in (f"B{rr}:C{rr}", f"D{rr}:G{rr}", f"H{rr}:I{rr}", f"J{rr}", f"K{rr}:L{rr}"):
                f.put(ref, None, size=9)
    rr = row + n_parts + 1
    f.put(f"B{rr}:I{rr}", None, size=9, border=Border(top=THIN, bottom=THIN, left=THIN))
    f.put(f"J{rr}", "部品費計", size=8, h="center", fill=PatternFill("solid", fgColor="F2F2F2"))
    f.put(f"K{rr}:L{rr}", total, size=9, h="right", fmt="#,##0", bold=True)
    ctx.rec.labels["parts_total"] = "部品費計"
    ctx.rec.values["parts"] = parts
    ctx.rec.values["parts_total"] = f"{total:,}"
    f.height(rr, 18)
    row = rr + 1

    block("修理結果", "result", inc.result, 2)
    block("再発防止策", "prevention", _text_or_blank(ctx, "prevention", inc.prevention), 2)
    block("水平展開", "horizontal_deployment", inc.horizontal_deployment, 2)

    # 影響ロット・廃棄枚数・修理費用
    lots = "、".join(inc.lots_affected) if inc.lots_affected else "なし"
    f.breakable(row)
    f.label(f"A{row}:A{row + 1}", "影響ロット", "lots", vertical=True, size=8)
    f.value(f"B{row}:F{row + 1}", lots, "lots", v="top")
    f.label(f"G{row}:G{row + 1}", "廃棄枚数", "scrap_wafers", size=9)
    f.value(f"H{row}:H{row + 1}", fm.z(f"{inc.scrap_wafers}枚"), "scrap_wafers", h="center")
    f.label(f"I{row}:J{row + 1}", "修理費用\n(概算)", "cost", size=9)
    disp, raw, nf = fm.yen(inc.cost_yen)
    f.value(f"K{row}:L{row + 1}", disp, "cost", raw=raw, fmt=nf, h="right")
    f.fit(f"B{row}:F{row + 1}", lots, min_row_pt=18, min_total_pt=5 * 11 + 8)
    row += 2

    # 写真
    rows_img = _img_rows(195) + 1
    n_rows = rows_img * (2 if len(ctx.captions) > 2 else 1)
    f.label(f"A{row}:A{row + n_rows - 1}", "写真", "photos", vertical=True, size=9)
    row = _photo_grid(f, ctx, row, [("B", "F"), ("G", "L")], (260, 195), "main", border_cols=("B", "L"), split_ok=False)
    block("備考", "remarks", _remarks_text(inc), 2)
    f.page_setup(f"A1:L{row - 1}")
    return f


# ---------------------------------------------------------------------------
# 追加シート（写真・参考資料・非表示シート）
# ---------------------------------------------------------------------------
def _add_photo_sheet(ctx: Ctx) -> None:
    inc, r = ctx.inc, ctx.r
    ws = ctx.wb.create_sheet("写真")
    f = Form(ws, {"A": 2, **{get_column_letter(c): 12 for c in range(2, 8)}, "H": 2}, "F2F2F2", None)
    part = inc.parts_used[0][0].name if inc.parts_used else None
    pool = ["装置画面（アラーム履歴）", "作業エリアの養生状況", "復旧後の装置画面", "点検口を開けた状態", "処置後の外観"]
    if part:
        pool += [f"交換前の{part}", f"取り外した{part}", f"交換後の{part}"]
    pool += [f"{c}（拡大）" for c in inc.photos]
    caps = r.sample(pool, k=min(len(pool), r.choice([2, 3, 4, 4, 5, 6])))
    f.put("B1:G1", f"写真（{inc.incident_id}　{inc.equipment.equipment_id} {inc.equipment.name}）", size=12, bold=True, border=None)
    f.height(1, 24)
    f.put("B2:G2", f"撮影日：{inc.occurred_at:%Y/%m/%d}　撮影者：{ctx.creator}", size=9, border=None)
    row = _photo_grid(f, ctx, 4, [("B", "D"), ("E", "G")], (240, 180), "sheet", border_cols=("B", "G"),
                      captions=caps, prefix="No.")
    ctx.rec.images["写真"] = len(caps)
    ctx.rec.extra_sheets["写真"] = {"captions": [f"No.{n}：{c}" for n, c in enumerate(caps, 1)]}
    f.page_setup(f"A1:H{row}")


def _add_reference_sheet(ctx: Ctx) -> None:
    inc, r = ctx.inc, ctx.r
    ws = ctx.wb.create_sheet("参考資料")
    f = Form(ws, {"A": 2, "B": 16, "C": 14, "D": 52, "E": 14}, "FFF2CC", None)
    f.put("B1:E1", f"参考資料　{inc.incident_id}（{inc.equipment.equipment_id}）", size=12, bold=True, border=None)
    f.height(1, 24)
    row = 3
    sections = []

    def header(title: str, cols: list[tuple[str, str]]) -> None:
        nonlocal row
        f.breakable(row)
        f.put(f"B{row}:E{row}", title, bold=True, border=None)
        sections.append(title)
        row += 1
        for ref, text in cols:
            f.put(ref.format(r=row), text, size=9, fill=f.fill, h="center")
        row += 1

    def line(cells: list[tuple[str, object]], text_for_fit: str = "", fit_ref: str = "D{r}") -> None:
        nonlocal row
        for ref, val in cells:
            f.put(ref.format(r=row), val, size=9, v="top")
        f.fit(fit_ref.format(r=row), text_for_fit, min_row_pt=17)
        row += 1

    kinds = ["timeline", "why", "history", "cost"]
    use = [k for k in kinds if k != "why" or inc.why_why]
    use = sorted(r.sample(use, k=min(len(use), r.choice([2, 3, 3, 4]))), key=kinds.index)

    if "timeline" in use:
        header(f"{len(sections) + 1}. 時系列", [("B{r}", "日時"), ("C{r}", "区分"), ("D{r}", "内容"), ("E{r}", "担当")])
        steps = [(inc.occurred_at, "発生", inc.symptom, inc.detected_by),
                 (inc.reported_at, "連絡", inc.first_response, D.surname(inc.reporter)),
                 (inc.response_started_at, "保全着手", "現地確認・調査開始", ctx.creator)]
        if inc.completed_at:
            steps.append((inc.completed_at, "復旧", inc.result or "生産へ引渡し", ctx.creator))
        for at, kind, text, who in steps:
            line([("B{r}", at.strftime("%Y/%m/%d %H:%M")), ("C{r}", kind), ("D{r}", text), ("E{r}", who)], text)
        row += 1
    if "why" in use:
        header(f"{len(sections) + 1}. なぜなぜ分析", [("B{r}", "段階"), ("C{r}:E{r}", "内容")])
        for n, w in enumerate(inc.why_why, 1):
            line([("B{r}", f"なぜ{n}"), ("C{r}:E{r}", w)], w, fit_ref="C{r}:E{r}")
        row += 1
    if "history" in use:
        header(f"{len(sections) + 1}. 同設備の過去トラブル（直近180日）", [("B{r}", "発生日"), ("C{r}", "TR No."), ("D{r}", "現象"), ("E{r}", "重要度")])
        past = [i for i in D.standard_incidents() if i.equipment.equipment_id == inc.equipment.equipment_id
                and inc.occurred_at - timedelta(days=180) <= i.occurred_at < inc.occurred_at][-6:]
        if not past:
            line([("B{r}", "―"), ("C{r}", "―"), ("D{r}", "該当なし"), ("E{r}", "―")])
        for i in past:
            line([("B{r}", i.occurred_at.strftime("%Y/%m/%d")), ("C{r}", i.incident_id), ("D{r}", i.symptom), ("E{r}", i.severity)], i.symptom)
        row += 1
    if "cost" in use:
        header(f"{len(sections) + 1}. 費用内訳", [("B{r}:D{r}", "項目"), ("E{r}", "金額(円)")])
        parts_sum = sum(p.unit_price_yen * q for p, q in inc.parts_used)
        labor = int(inc.work_hours * 6500)
        other = inc.cost_yen - parts_sum - labor
        items = [(f"部品費（{len(inc.parts_used)}品目）", parts_sum), (f"社内工数 {inc.work_hours:g}h × 6,500円", labor)]
        if other >= 1000:
            items.append(("その他（FE作業費・廃棄ウェーハ等）", other))
        else:   # 合計は百円単位で丸めているため、端数は工数側に含めて書く
            items[1] = (f"社内工数 {inc.work_hours:g}h（端数調整含む）", inc.cost_yen - parts_sum)
        for text, val in items + [("合計", inc.cost_yen)]:
            f.put(f"B{row}:D{row}", text, size=9, bold=text == "合計")
            f.put(f"E{row}", val, size=9, h="right", fmt="#,##0", bold=text == "合計")
            f.height(row, 17)
            row += 1
    ctx.rec.extra_sheets["参考資料"] = {"sections": sections}
    f.page_setup(f"A1:E{row}")


def _add_hidden_sheet(ctx: Ctx, main_ws, kind: str, eq_cell: str | None, sev_cell: str | None) -> None:
    inc = ctx.inc
    ws = ctx.wb.create_sheet(kind)
    ws.sheet_state = "hidden"
    ctx.rec.hidden_sheets.append(kind)
    if kind == "リスト":
        # 入力規則（ドロップダウン）用のリスト。現場の様式でよくある作り
        eqs = D.equipment_master()
        cols = {"A": ("設備番号", [e.equipment_id for e in eqs]), "B": ("設備名", [e.name for e in eqs]),
                "D": ("故障区分", ["機械", "電気", "制御", "ソフト", "ユーティリティ", "人為", "品質", "外部要因"]),
                "E": ("重要度", ["重大", "大", "中", "小"]),
                "G": ("保全担当", [p.name for p in D.people() if p.section.startswith("設備保全課")])}
        for col, (title, items) in cols.items():
            ws[f"{col}1"] = title
            for n, v in enumerate(items, 2):
                ws[f"{col}{n}"] = v
        if eq_cell:
            dv = DataValidation(type="list", formula1=f"'リスト'!$A$2:$A${len(eqs) + 1}", allow_blank=True)
            main_ws.add_data_validation(dv)
            dv.add(eq_cell)
        if sev_cell:
            dv2 = DataValidation(type="list", formula1="'リスト'!$E$2:$E$5", allow_blank=True)
            main_ws.add_data_validation(dv2)
            dv2.add(sev_cell)
        ctx.rec.irregularities.append("非表示シート「リスト」（入力規則のリスト元）")
    else:
        # 作成者が下書きに使ったメモを非表示にしたまま提出したケース
        memo = [f"{inc.occurred_at:%m/%d %H:%M} 発生（{inc.detected_by}）",
                f"{inc.reported_at:%H:%M} {D.surname(inc.reporter)}さんから連絡",
                f"{inc.response_started_at:%H:%M} 着手"]
        memo += [f"・{x}" for x in inc.first_response.replace(" → ", "→").split("→") if x.strip()][:3]
        if inc.parts_used:
            memo.append("部品：" + "、".join(f"{p.name}×{q}" for p, q in inc.parts_used))
        if inc.completed_at:
            memo.append(f"{inc.completed_at:%m/%d %H:%M} 復旧")
        memo.append("※報告書清書後にこのシートは削除すること")
        ws["A1"] = "作業メモ（下書き）"
        ws["A1"].font = Font(name=FONT_NAME, bold=True)
        for n, m in enumerate(memo, 3):
            ws[f"A{n}"] = m
        ws.column_dimensions["A"].width = 80
        ctx.rec.extra_sheets["作業メモ"] = {"hidden": True, "lines": memo}
        ctx.rec.irregularities.append("非表示シート「作業メモ」（下書きメモが残っている）")


# ---------------------------------------------------------------------------
# ワークブック組み立て・保存
# ---------------------------------------------------------------------------
def _build_workbook(spec: FileSpec) -> tuple[Workbook, Record]:
    r = D.rng(f"f1-repair-report:{spec.inc.incident_id}")
    wb = Workbook()
    rec = Record()
    ctx = Ctx(spec=spec, r=r, fmt=Fmt(spec, r), rec=rec, wb=wb)
    ctx.approver, ctx.checker, ctx.creator = _stamps(spec)
    ctx.captions = _photo_captions(spec.inc, r)

    builder = {"v1": _build_v1, "v2": _build_v2, "v3": _build_v3}[spec.version]
    form = builder(ctx)
    ws = form.ws

    # 版固有の既定イレギュラー
    inc = spec.inc
    if inc.status in ("対応中", "保留"):
        rec.irregularities.append(f"中間報告（状態: {inc.status}）のため原因・結果などが未記入または暫定")
    if not ctx.approver:
        rec.irregularities.append("承認欄が未押印")
    if spec.version == "v1":
        rec.irregularities.append("重要度がチェックボックス表記（■大 など）")
    if spec.version == "v2":
        rec.irregularities.append("項目名のゆれ（v2 同義語）と1セル内「ラベル：値」")
    if spec.version == "v3":
        rec.irregularities.append("文章項目のラベルが縦書き・縦結合")
    if spec.zenkaku:
        rec.irregularities.append("日付・数値欄が全角数字")
    if spec.blank_field:
        label = rec.labels.get(spec.blank_field, spec.blank_field)
        rec.irregularities.append(f"必須項目の空欄: {label}")
    if spec.version == "v1" and _version_of_date(spec.report_date) != "v1":
        rec.irregularities.append(f"改訂後も旧様式（{REV_INFO['v1'][0]}）を使用")
    if spec.version == "v2" and _version_of_date(spec.report_date) == "v3":
        rec.irregularities.append(f"改訂後も旧様式（{REV_INFO['v2'][0]}）を使用")

    if spec.photo_sheet:
        _add_photo_sheet(ctx)
        rec.irregularities.append("追加シート「写真」")
    if spec.ref_sheet:
        _add_reference_sheet(ctx)
        rec.irregularities.append("追加シート「参考資料」")
    if spec.hidden:
        eq_cell = sev_cell = None
        for row in ws.iter_rows():
            for c in row:
                if c.value == inc.equipment.equipment_id and eq_cell is None:
                    eq_cell = c.coordinate
                if c.value == inc.severity and sev_cell is None and spec.version != "v1":
                    sev_cell = c.coordinate
        _add_hidden_sheet(ctx, ws, spec.hidden, eq_cell, sev_cell)

    writer = _writer_person(spec)
    created = datetime.combine(spec.report_date, datetime.min.time()) + timedelta(hours=r.randint(9, 18), minutes=r.randint(0, 59))
    wb.properties.creator = writer.name
    wb.properties.lastModifiedBy = writer.name
    wb.properties.title = f"設備修理報告書 {spec.report_id}"
    wb.properties.created = created
    wb.properties.modified = created + timedelta(minutes=r.randint(5, 240))
    return wb, rec


def _version_of_date(d: date) -> str:
    return "v1" if d < date(2024, 4, 1) else ("v2" if d < date(2025, 7, 1) else "v3")


def _save_deterministic(wb: Workbook, path: Path) -> None:
    """openpyxl の保存結果を、ZIP内タイムスタンプを固定して書き直す（再実行で同一バイトにする）。"""
    buf = BytesIO()
    archive = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, allowZip64=True)
    ExcelWriter(wb, archive).save()   # save_workbook() は modified を現在時刻で上書きするため使わない
    src = zipfile.ZipFile(BytesIO(buf.getvalue()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            zi = zipfile.ZipInfo(info.filename, date_time=ZIP_TIME)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            out.writestr(zi, src.read(info.filename))


# ---------------------------------------------------------------------------
# README
# ---------------------------------------------------------------------------
def _readme(specs: list[FileSpec], records: list[Record]) -> str:
    def files_with(pred) -> str:
        names = [s.filename for s, rc in zip(specs, records) if pred(s, rc)]
        return "、".join(f"`{n}`" for n in names) if names else "（なし）"

    cnt = {v: sum(1 for s in specs if s.version == v) for v in ("v1", "v2", "v3")}
    syn_rows = "\n".join(f"| {k} | {' / '.join(v)} |" for k, v in V2_SYNONYMS.items())
    lines = [
        "# F1 設備修理報告書（サンプル帳票）",
        "",
        "製造部 設備保全課の「設備修理報告書」（" + FORM_NO + "）を想定した Excel 帳票サンプル 30 ファイル。",
        "`scripts/samples/domain.py` の正準トラブル履歴 `standard_incidents()` から、重要度「大」「重大」または部品を使用した",
        "トラブルを期間・設備カテゴリ・状態が偏らないように選んで作成している。",
        "",
        "再生成: `python -m scripts.samples.f1_repair_report`（固定シードのため何度実行しても同一バイトのファイルになる）",
        "",
        "## フォルダ構成",
        "",
        "様式の版ごとにサブフォルダを分けてある。各サブフォルダの中身は同じ版の .xlsx だけなので、",
        "アプリの「帳票取り込み」にフォルダ1つ分をまとめて入れれば、帳票の種類とシートを1回選ぶだけで全ファイルを読み取れる。",
        "フォルダ名は「版の名前＋いつから」。",
        "",
        "```",
        FORM_DIR + "/",
        "├─ _README.md        … この説明",
        "├─ _expected.jsonl   … 抽出結果の正解データ（30行。file は版フォルダからのパス）",
        *[f"{'└─' if v == 'v3' else '├─'} {_version_dir(v)}/   … {REV_INFO[v][0]}（{REV_INFO[v][1]}）の帳票 {cnt[v]} ファイル"
          for v in ("v1", "v2", "v3")],
        "```",
        "",
        "## 様式の版（改訂履歴）",
        "",
        "| 版 | layout_version | 制定・改訂 | フォルダ | ファイル数 | レイアウトの特徴 |",
        "|---|---|---|---|---|---|",
        f"| {REV_INFO['v1'][0]} | v1 | {REV_INFO['v1'][1]} | `{_version_dir('v1')}/` | {cnt['v1']} | A〜H列。左ラベル・右値の2列組み（報告番号/報告日/管理No.、設備番号/設備名 …）。"
        "文章項目（故障内容・原因調査・原因・修理内容・修理結果・再発防止策・備考）は見出し行の下に複数行結合セル。"
        "承認欄（承認/確認/作成）は右上。重要度は「□重大　■大　□中　□小」のチェックボックス表記。日時は日付型セル＋表示形式。 |",
        f"| {REV_INFO['v2'][0]} | v2 | {REV_INFO['v2'][1]} | `{_version_dir('v2')}/` | {cnt['v2']} | A列を狭い余白にした B〜K列。ラベルは B:C結合＋値 D:F、右側は H＋I:K と列位置がズレる。"
        "項目名がファイルごとに揺れる（下表）。トラブルNo./設備番号/設備名の一部は「設備番号：CMP-101」のような1セル内ラベル。"
        "応急処置・影響ロット欄あり。日時・停止時間は文字列（例: R6.5.12、17.2h、17時間12分）。承認欄は下部右（承認/確認/担当 or 作成）。 |",
        f"| {REV_INFO['v3'][0]} | v3 | {REV_INFO['v3'][1]} | `{_version_dir('v3')}/` | {cnt['v3']} | A〜L列。上部は見出し行＋値行の表形式（報告番号・報告日・設備番号・設備名・ライン・重要度 / "
        "発生日時・復旧日時・停止時間(分)・報告者・担当者・故障区分 / TR番号・アラーム・原因区分・対応状況）。"
        "文章項目は A列の縦書き・縦結合ラベル＋B:L結合セル。なぜなぜ分析・水平展開・廃棄枚数・修理費用・部品費計あり。承認欄は右上。 |",
        "",
        "旧様式の使い続け: 改訂後の日付でも古い版のファイルが3件ある（Rev.2期間の Rev.1、Rev.3期間の Rev.1 と Rev.2）。"
        "フォルダ分けは日付ではなく様式の版に合わせているため、この3件も版のフォルダに入っている。",
        "",
        "## 共通の体裁",
        "",
        "- フォント: " + FONT_NAME + "（タイトル16〜18pt、本文10pt、部品表9pt）。ラベルセルは塗りつぶし（v1 淡青 / v2 淡緑 / v3 灰）",
        "- 罫線: 項目枠は細線、v1 タイトル下は太線。結合セルは外周に罫線",
        "- 行の高さ: 文章量（全角2・半角1で見積もり）に合わせて調整。v3 の縦書きラベルは文字数分の高さを確保",
        "- 印刷設定: A4縦・横1ページに合わせる・水平中央・印刷範囲設定・フッタにページ番号。項目ブロックや写真が"
        "ページで分断されないよう、ブロックの切れ目に手動改ページ（2ページ構成が多い）。v1/v3 は枠線非表示",
        "- 押印: 承認欄は赤字の姓（電子印の代わり）。確認者が作成者本人の場合は「／」",
        "- 写真: Pillow で描いた設備スケッチ（赤枠で着目部位を囲み、撮影日時・設備番号・キャプション入り）の PNG を1〜3枚、",
        "  各写真の下のセルに「写真1：キャプション」を記入",
        "- ブックのプロパティ（作成者・作成日時）は報告書の作成者・報告日に合わせてある",
        "",
        "## 報告番号の体系",
        "",
        "- v1: `保全-年度-連番4桁`（例: 保全-2023-0412、4月始まりの年度）",
        "- v2: `R西暦-連番4桁`（例: R2024-0877）",
        "- v3: `MR-YYMM-連番3桁`（例: MR-2508-041）",
        "- どの版もトラブル管理番号（TR-YYYY-NNNNN）を別欄に記載（v1 管理No. / v2 トラブルNo. 等 / v3 TR番号）",
        "",
        "## v2 の項目名ゆれ（ファイルごとに1つ）",
        "",
        "| キー | ラベル候補 |",
        "|---|---|",
        syn_rows,
        "",
        "## 意図的なイレギュラー",
        "",
        f"- 追加シート「写真」（写真を2〜6枚追加）: {files_with(lambda s, rc: s.photo_sheet)}",
        f"- 追加シート「参考資料」（時系列・なぜなぜ分析・同設備の過去トラブル・費用内訳から2〜4ブロック）: {files_with(lambda s, rc: s.ref_sheet)}",
        f"- 非表示シート「リスト」（設備番号・重要度の入力規則のリスト元）: {files_with(lambda s, rc: s.hidden == 'リスト')}",
        f"- 非表示シート「作業メモ」（下書きメモが残ったまま）: {files_with(lambda s, rc: s.hidden == '作業メモ')}",
        f"- 日付・数値欄が全角数字（例: ２０２３年６月８日 ２３:５６、１,２３５分。コロン・カンマは半角のまま）: {files_with(lambda s, rc: s.zenkaku)}",
        f"- 必須項目の空欄: " + "、".join(f"`{s.filename}`（{rc.labels.get(s.blank_field, s.blank_field)}）"
                                         for s, rc in zip(specs, records) if s.blank_field),
        f"- 中間報告（対応中・保留。原因/結果/復旧日時などが未記入・暫定、承認欄未押印）: "
        f"{files_with(lambda s, rc: s.inc.status in ('対応中', '保留'))}",
        f"- 承認欄が未押印: {files_with(lambda s, rc: not rc.values.get('approver'))}",
        f"- 写真を印刷範囲外（M列）に縦に貼付（枠内は「別紙のとおり（右側に貼付）」）: {files_with(lambda s, rc: s.photos_outside)}",
        f"- シート名が「Sheet1」のまま: {files_with(lambda s, rc: s.default_sheet_name)}",
        "- 記入内容そのものの揺れ: 記入者ごとの書き癖（敬体、半角カナ、全角数字、①/(1)/1) の番号、誤変換）は domain.py 由来",
        "- 再発案件（備考に「TR-… の再発」）: " + files_with(lambda s, rc: s.inc.recurrence)
        + "。うち前回トラブルの報告書もこのサンプルに含まれるもの: "
        + files_with(lambda s, rc: s.inc.recurrence and s.inc.related_incident_id in {x.inc.incident_id for x in specs}),
        "- ファイル名の付け方がばらばら（報告番号入り／日付入り／【重大】付き／_中間報告 付き）",
        "",
        "## _expected.jsonl",
        "",
        "1行1ファイルの JSON（UTF-8）。`file` はこのフォルダからの相対パス"
        f"（例: `{_version_dir(specs[0].version)}/{specs[0].filename}`）。",
        "",
        "```",
        '{"file": 版フォルダ名/ファイル名, "layout_version": "v1"|"v2"|"v3", "form_revision": "Rev.1" 等, "main_sheet": 報告書シート名,',
        ' "values": {キー: 人が読むとおりの値（表示形式適用後の文字列。部品は行ごとの dict のリスト、写真はキャプションのリスト）},',
        ' "labels_used": {キー: そのファイルで実際に使われているラベル文字列（部品表の列名は parts.part_no 等）},',
        ' "sheets": 全シート名, "hidden_sheets": [...], "images": {シート名: 枚数}, "extra_sheets": {...},',
        ' "irregularities": [...], "incident_id": 元トラブルID, "status": 状態, "severity": 重要度}',
        "```",
        "",
        "主なキー: report_id, report_date, incident_id, equipment_id, equipment_name, line, maker, occurred_at, recovered_at, "
        "downtime_min, severity, category, reporter, assignee, work_hours, cost, alarm, symptom, first_response, investigation, "
        "cause, cause_category, why_why, action, parts, parts_total, result, prevention, horizontal_deployment, lots, scrap_wafers, "
        "status, report_type, photos, remarks, approver, checker, creator",
        "",
        "- 空欄は空文字 `\"\"`。チェックボックス表記は■（選択）の付いた選択肢の文字で記録（例: 重要度 `大`、v2 の報告区分「■完了　□中間」→ `完了`）",
        "- 日付型セルは Excel での表示（例: `2024/05/12 22:25`、`2023年11月2日`）で記録",
        "",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
def generate(output_root: Path | None = None) -> list[Path]:
    root = Path(output_root) if output_root is not None else D.OUTPUT_ROOT
    out_dir = root / "forms" / FORM_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    # 名前の変わった古い生成物を残さない（このフォルダは本スクリプト専用）。
    # 版フォルダを使う前の直下の .xlsx と、今は作らない版フォルダも消す（_README.md / _expected.jsonl は残す）。
    keep_dirs = {_version_dir(v) for v in REV_INFO}
    for old in out_dir.glob("*.xlsx"):
        old.unlink()
    for sub in sorted(p for p in out_dir.iterdir() if p.is_dir()):
        if sub.name in keep_dirs:
            for old in sub.glob("*.xlsx"):
                old.unlink()
        else:
            shutil.rmtree(sub)
    for name in sorted(keep_dirs):
        (out_dir / name).mkdir(exist_ok=True)

    specs = _file_specs()
    records, paths = [], []
    for spec in specs:
        wb, rec = _build_workbook(spec)
        path = out_dir / _version_dir(spec.version) / spec.filename
        _save_deterministic(wb, path)
        rec_sheets = wb.sheetnames
        records.append((rec, rec_sheets))
        paths.append(path)

    exp_path = out_dir / "_expected.jsonl"
    with exp_path.open("w", encoding="utf-8", newline="\n") as fp:
        for spec, (rec, sheets) in zip(specs, records):
            obj = {
                "file": f"{_version_dir(spec.version)}/{spec.filename}",   # 帳票フォルダからの相対パス
                "layout_version": spec.version,
                "form_revision": REV_INFO[spec.version][0],
                "main_sheet": rec.main_sheet,
                "values": rec.values,
                "labels_used": rec.labels,
                "sheets": sheets,
                "hidden_sheets": rec.hidden_sheets,
                "images": rec.images,
                "extra_sheets": rec.extra_sheets,
                "irregularities": rec.irregularities,
                "incident_id": spec.inc.incident_id,
                "status": spec.inc.status,
                "severity": spec.inc.severity,
            }
            fp.write(json.dumps(obj, ensure_ascii=False) + "\n")
    readme_path = out_dir / "_README.md"
    readme_path.write_text(_readme(specs, [rc for rc, _ in records]), encoding="utf-8", newline="\n")
    return paths + [exp_path, readme_path]


if __name__ == "__main__":
    import time

    t0 = time.time()
    written = generate()
    for p in written:
        print(p)
    print(f"{len(written)} files ({time.time() - t0:.1f}s)")
