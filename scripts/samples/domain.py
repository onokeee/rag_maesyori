"""サンプルデータ共通ドメインモデル（製造部 設備保全課を想定した半導体工場）。

    python -m scripts.samples.domain      # 件数・分布のサマリを表示

設備マスタ・人・部品・アラーム・トラブル履歴・チョコ停を、固定シードで決定的に生成する。
各サンプル生成スクリプトはここから同じデータを取り出し、Excel帳票やCSVに書き出す。

- トラブル（Incident）は設備カテゴリ×サブシステム別の「シナリオ」（_bank_*.py）を核に、
  測定値などのパラメータ・記入者の書き癖・共通フレーズを組み合わせて文章を作る。
- 発生頻度は 設備種別 × 経年 × 初期故障 × 持病（悪い号機）× 季節 × 休日 で重み付けする。
- 運転時間・RF積算・処理枚数などの値は、チャンバー/ヘッド/号車ごとの使用量カウンタ（トラブル対応の部品交換と
  T4 の定期交換でリセット）から決める。影響台数は設備マスタの実台数、MFC等の使用年数は設置日を上限にする。
- 処置に「○時間保持」「慣らし運転○分」などの待ち時間が書かれていれば、停止時間はそれを下回らない。
- 同じ関数を何度呼んでも（プロセスを変えても）同一の結果になる。乱数は rng(name) からのみ取る。
"""
from __future__ import annotations

import bisect
import math
import random
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path

from . import _bank_front, _bank_infra, _bank_mid

# ---------------------------------------------------------------------------
# 基本設定
# ---------------------------------------------------------------------------
SEED = 20230401
OUTPUT_ROOT: Path = Path(__file__).resolve().parents[2] / "samples"
PERIOD_START = date(2023, 4, 1)
PERIOD_END = date(2026, 8, 31)
N_INCIDENTS = 8000          # 瞬低の同時多発分が少し上乗せされる
N_MINOR_STOPS = 25000
RECURRENCE_DAYS = 30        # 同一設備・同一故障モードがこの日数以内に再発したら「再発」扱い


def rng(name: str) -> random.Random:
    """名前ごとに独立した決定的乱数を返す（str シードは sha512 で展開されるためプロセス間でも同一）。"""
    return random.Random(f"{SEED}:{name}")


# ---------------------------------------------------------------------------
# データクラス
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Equipment:
    equipment_id: str
    name: str
    category: str
    maker: str
    model: str
    line: str
    process: str
    area: str
    installed_date: date
    criticality: str  # "A" / "B" / "C"


@dataclass(frozen=True)
class Person:
    employee_id: str
    name: str
    department: str
    section: str
    role: str


@dataclass(frozen=True)
class Part:
    part_no: str
    name: str
    maker: str
    unit_price_yen: int
    subsystem: str
    applicable_categories: tuple


@dataclass(frozen=True)
class AlarmCode:
    code: str
    message: str
    category: str   # 重故障 / 軽故障 / 警告 / インターロック
    subsystem: str


@dataclass
class Incident:
    incident_id: str
    occurred_at: datetime
    detected_by: str
    reported_at: datetime
    response_started_at: datetime
    completed_at: datetime | None
    equipment: Equipment
    shift: str
    reporter: Person
    assignees: list
    category: str
    subsystem: str
    severity: str
    status: str
    symptom: str
    alarm: AlarmCode | None
    first_response: str
    investigation: str
    cause: str
    cause_category: str
    why_why: list
    action: str
    parts_used: list
    downtime_min: int
    work_hours: float
    result: str
    prevention: str
    horizontal_deployment: str
    lots_affected: list
    scrap_wafers: int
    cost_yen: int
    recurrence: bool
    related_incident_id: str | None
    remarks: str
    photos: list = field(default_factory=list)
    scenario_id: str = ""   # 生成元シナリオ（分析・デバッグ用の追加項目）


@dataclass
class MinorStop:
    event_id: str
    occurred_at: datetime
    equipment: Equipment
    alarm: AlarmCode | None
    duration_min: int
    recovered_by: str
    note: str


# ---------------------------------------------------------------------------
# 書式ヘルパー
# ---------------------------------------------------------------------------
_ZEN_TABLE = {c: c + 0xFEE0 for c in range(0x21, 0x7F)}


def to_zenkaku(s: str, digits_only: bool = False) -> str:
    """ASCII英数記号を全角に変換する。digits_only=True なら数字だけ。"""
    if digits_only:
        return s.translate({c: c + 0xFEE0 for c in range(0x30, 0x3A)})
    return s.translate(_ZEN_TABLE)


_HANKANA_WORDS = {"アラーム": "ｱﾗｰﾑ", "リセット": "ﾘｾｯﾄ", "ポンプ": "ﾎﾟﾝﾌﾟ", "センサ": "ｾﾝｻ", "エラー": "ｴﾗｰ",
                  "ロット": "ﾛｯﾄ", "ウェーハ": "ｳｪｰﾊ", "チャンバー": "ﾁｬﾝﾊﾞｰ", "フィルタ": "ﾌｨﾙﾀ"}


def to_hankaku_kana(s: str) -> str:
    """現場でよく見る半角カナ表記（一部の頻出語のみ）に置き換える。"""
    for k, v in _HANKANA_WORDS.items():
        s = s.replace(k, v)
    return s


def _wareki(d: date) -> tuple[str, int]:
    if d >= date(2019, 5, 1):
        return "R", d.year - 2018
    return "H", d.year - 1988


def fmt_date_variants(d: date | datetime) -> list[str]:
    """同じ日付の、現場の帳票でありがちな表記ゆれを返す。"""
    if isinstance(d, datetime):
        d = d.date()
    g, y = _wareki(d)
    gname = "令和" if g == "R" else "平成"
    return [
        f"{d.year}/{d.month:02d}/{d.day:02d}",
        f"{d.year}-{d.month:02d}-{d.day:02d}",
        f"{g}{y}.{d.month}.{d.day}",
        f"{d.year}年{d.month}月{d.day}日",
        f"{d.year}/{d.month}/{d.day}",
        f"{d.year % 100:02d}/{d.month}/{d.day}",
        f"{gname}{y}年{d.month}月{d.day}日",
        f"{d.month}/{d.day}",
        f"{d.year}.{d.month:02d}.{d.day:02d}",
    ]


def fmt_date(d: date | datetime, r: random.Random | None = None, style: int | None = None) -> str:
    """日付を1表記で返す。style 指定がなければ r でランダム選択（r も無ければ YYYY/MM/DD）。"""
    v = fmt_date_variants(d)
    if style is not None:
        return v[style % len(v)]
    if r is None:
        return v[0]
    return r.choices(v, weights=[30, 15, 8, 12, 10, 6, 3, 8, 8])[0]


def fmt_datetime_variants(dt: datetime) -> list[str]:
    hm = f"{dt.hour:02d}:{dt.minute:02d}"
    return [f"{v} {hm}" for v in fmt_date_variants(dt)[:5]] + [
        f"{dt.month}/{dt.day} {dt.hour}時{dt.minute:02d}分", f"{dt.year}/{dt.month:02d}/{dt.day:02d} {dt.hour}:{dt.minute:02d}"]


# ---------------------------------------------------------------------------
# マスタ: 設備
# ---------------------------------------------------------------------------
CATEGORY_OF_KIND = {
    "CMP": "CMP装置", "CVD": "CVD装置", "ETC": "エッチング装置", "LIT": "露光装置", "CLN": "洗浄装置",
    "IMP": "イオン注入装置", "INS": "検査装置", "OHT": "搬送系", "AGV": "搬送系", "ROB": "搬送系", "STK": "搬送系",
    "UPW": "ユーティリティ", "EXH": "ユーティリティ", "SCR": "ユーティリティ", "PCW": "ユーティリティ", "VAC": "ユーティリティ",
}
PROCESS_KINDS = ("CMP", "CVD", "ETC", "LIT", "CLN", "IMP", "INS")
TRANSPORT_KINDS = ("OHT", "AGV", "ROB", "STK")
UTILITY_KINDS = ("UPW", "EXH", "SCR", "PCW", "VAC")

# メーカー正式名 -> 報告書で使う略称
MAKER_SHORT = {
    "扶桑ポリッシュシステム": "扶桑", "大和精機": "大和", "昭栄プラズマテック": "昭栄", "新光ファブテック": "新光",
    "瑞穂エンジニアリング": "瑞穂", "光陽オプトロニクス": "光陽", "北辰光学": "北辰", "清流ウェットテック": "清流",
    "日昇イオンテクノロジー": "日昇", "明光インスペクション": "明光", "大鷹搬送システム": "大鷹", "三峰ロボティクス": "三峰",
    "水晶アクアシステム": "水晶アクア", "東邦空調設備": "東邦空調", "環研エコシステム": "環研", "精和バキューム": "精和",
}

# (ID, 名称, メーカー, 型式, ライン, 工程, エリア, 設置日, 重要度)
_EQUIPMENT_ROWS = [
    ("CMP-101", "酸化膜CMP 1号機", "扶桑ポリッシュシステム", "FPS-300G", "L1", "STI-CMP", "Fab1 2F CR ベイ05", (2009, 6, 12), "A"),
    ("CMP-102", "酸化膜CMP 2号機", "扶桑ポリッシュシステム", "FPS-300G", "L1", "ILD-CMP", "Fab1 2F CR ベイ05", (2010, 2, 3), "A"),
    ("CMP-103", "W-CMP 3号機", "大和精機", "DY-CMP8", "L2", "W-CMP", "Fab1 2F CR ベイ06", (2011, 9, 20), "A"),
    ("CMP-104", "Cu-CMP 4号機", "扶桑ポリッシュシステム", "FPS-300X", "L3", "Cu-CMP", "Fab1 2F CR ベイ07", (2014, 4, 8), "A"),
    ("CMP-105", "Cu-CMP 5号機", "扶桑ポリッシュシステム", "FPS-300X", "L4", "Cu-CMP", "Fab2 3F CR ベイ12", (2018, 7, 2), "A"),
    ("CMP-106", "酸化膜CMP 6号機", "大和精機", "DY-CMP8e", "L5", "ILD-CMP", "Fab2 3F CR ベイ13", (2019, 11, 15), "B"),
    ("CMP-107", "W-CMP 7号機", "扶桑ポリッシュシステム", "FPS-300X", "L5", "W-CMP", "Fab2 3F CR ベイ13", (2021, 3, 1), "B"),
    ("CMP-108", "STI-CMP 8号機", "扶桑ポリッシュシステム", "FPS-300XE", "L6", "STI-CMP", "Fab2 3F CR ベイ16", (2024, 10, 7), "A"),
    ("CVD-201", "P-SiN CVD 1号機", "昭栄プラズマテック", "SPT-Pro300", "L1", "P-SiN成膜", "Fab1 2F CR ベイ08", (2008, 11, 4), "B"),
    ("CVD-202", "TEOS CVD 2号機", "昭栄プラズマテック", "SPT-Pro300", "L1", "TEOS-SiO2成膜", "Fab1 2F CR ベイ08", (2009, 5, 18), "A"),
    ("CVD-203", "W-CVD 3号機", "新光ファブテック", "NKF-W300", "L2", "W-CVD", "Fab1 2F CR ベイ09", (2012, 1, 23), "A"),
    ("CVD-204", "SiON CVD 4号機", "昭栄プラズマテック", "SPT-Lumina", "L3", "SiON成膜", "Fab1 2F CR ベイ09", (2015, 8, 6), "B"),
    ("CVD-205", "P-SiN CVD 5号機", "新光ファブテック", "NKF-C300", "L3", "P-SiN成膜", "Fab1 2F CR ベイ10", (2013, 3, 11), "B"),
    ("CVD-206", "TEOS CVD 6号機", "昭栄プラズマテック", "SPT-Lumina", "L4", "TEOS-SiO2成膜", "Fab2 3F CR ベイ14", (2018, 5, 21), "A"),
    ("CVD-207", "W-CVD 7号機", "新光ファブテック", "NKF-W300", "L5", "W-CVD", "Fab2 3F CR ベイ14", (2020, 2, 17), "A"),
    ("CVD-208", "P-SiN CVD 8号機", "昭栄プラズマテック", "SPT-Lumina", "L5", "パッシベーション成膜", "Fab2 3F CR ベイ15", (2021, 9, 6), "B"),
    ("CVD-209", "SiON CVD 9号機", "昭栄プラズマテック", "SPT-Lumina II", "L6", "SiON成膜", "Fab2 3F CR ベイ17", (2024, 11, 18), "B"),
    ("ETC-301", "Polyエッチャ 1号機", "瑞穂エンジニアリング", "MZ-Etch3000", "L1", "Poly Gate Etch", "Fab1 2F CR ベイ03", (2008, 9, 1), "A"),
    ("ETC-302", "Oxideエッチャ 2号機", "瑞穂エンジニアリング", "MZ-Etch3000", "L2", "Contact Etch", "Fab1 2F CR ベイ03", (2010, 6, 14), "A"),
    ("ETC-303", "Metalエッチャ 3号機", "昭栄プラズマテック", "SPT-Etch M", "L2", "Metal Etch", "Fab1 2F CR ベイ04", (2011, 12, 5), "B"),
    ("ETC-304", "SiNエッチャ 4号機", "瑞穂エンジニアリング", "MZ-Etch Neo", "L3", "SiN Spacer Etch", "Fab1 2F CR ベイ04", (2016, 2, 29), "B"),
    ("ETC-305", "Polyエッチャ 5号機", "瑞穂エンジニアリング", "MZ-Etch Neo", "L4", "Poly Gate Etch", "Fab2 3F CR ベイ10", (2018, 3, 26), "A"),
    ("ETC-306", "Oxideエッチャ 6号機", "瑞穂エンジニアリング", "MZ-Etch Neo", "L4", "Via Etch", "Fab2 3F CR ベイ10", (2018, 10, 1), "A"),
    ("ETC-307", "Metalエッチャ 7号機", "昭栄プラズマテック", "SPT-Etch M2", "L5", "Metal Etch", "Fab2 3F CR ベイ11", (2020, 7, 13), "B"),
    ("ETC-308", "アッシャ 8号機", "昭栄プラズマテック", "SPT-Ash200", "L5", "レジストアッシング", "Fab2 3F CR ベイ11", (2019, 4, 22), "C"),
    ("ETC-309", "Oxideエッチャ 9号機", "瑞穂エンジニアリング", "MZ-Etch Neo+", "L6", "Contact Etch", "Fab2 3F CR ベイ18", (2025, 1, 20), "A"),
    ("LIT-401", "KrFスキャナ 1号機", "光陽オプトロニクス", "KY-KrF300", "L1", "KrF露光", "Fab1 2F CR リソベイ01", (2009, 3, 2), "B"),
    ("LIT-402", "ArFスキャナ 2号機", "光陽オプトロニクス", "KY-ArF550", "L2", "ArF露光(Critical)", "Fab1 2F CR リソベイ01", (2013, 10, 21), "A"),
    ("LIT-403", "i線ステッパ 3号機", "北辰光学", "HS-i12", "L3", "i線露光", "Fab1 2F CR リソベイ02", (2007, 7, 9), "C"),
    ("LIT-404", "ArFスキャナ 4号機", "光陽オプトロニクス", "KY-ArF550", "L4", "ArF露光(Critical)", "Fab2 3F CR リソベイ05", (2018, 8, 27), "A"),
    ("LIT-405", "KrFスキャナ 5号機", "光陽オプトロニクス", "KY-KrF300B", "L5", "KrF露光", "Fab2 3F CR リソベイ05", (2020, 1, 14), "B"),
    ("LIT-406", "ArFスキャナ 6号機", "光陽オプトロニクス", "KY-ArF600", "L6", "ArF露光(Critical)", "Fab2 3F CR リソベイ06", (2025, 3, 10), "A"),
    ("CLN-501", "RCA洗浄 1号機", "清流ウェットテック", "SWT-WB8", "L1", "前洗浄(RCA)", "Fab1 2F CR ベイ01", (2008, 10, 6), "B"),
    ("CLN-502", "枚葉洗浄 2号機", "清流ウェットテック", "SWT-SS300", "L2", "ポストCMP洗浄", "Fab1 2F CR ベイ06", (2012, 6, 25), "B"),
    ("CLN-503", "HF洗浄 3号機", "清流ウェットテック", "SWT-WB8", "L3", "HF洗浄", "Fab1 2F CR ベイ02", (2014, 11, 10), "B"),
    ("CLN-504", "枚葉洗浄 4号機", "清流ウェットテック", "SWT-SS300", "L4", "レジスト剥離後洗浄", "Fab2 3F CR ベイ09", (2018, 4, 16), "A"),
    ("CLN-505", "枚葉洗浄 5号機", "清流ウェットテック", "SWT-SS300", "L5", "ポストCMP洗浄", "Fab2 3F CR ベイ13", (2020, 9, 7), "B"),
    ("CLN-506", "枚葉洗浄 6号機", "清流ウェットテック", "SWT-SS300R", "L6", "前洗浄", "Fab2 3F CR ベイ16", (2024, 9, 2), "B"),
    ("IMP-601", "高電流注入 1号機", "日昇イオンテクノロジー", "NIT-HC3", "L1", "高電流注入(S/D)", "Fab1 2F CR ベイ11", (2009, 1, 19), "A"),
    ("IMP-602", "中電流注入 2号機", "日昇イオンテクノロジー", "NIT-MC5", "L2", "中電流注入(Well)", "Fab1 2F CR ベイ11", (2011, 4, 4), "B"),
    ("IMP-603", "高電流注入 3号機", "日昇イオンテクノロジー", "NIT-HC3", "L4", "高電流注入(S/D)", "Fab2 3F CR ベイ08", (2017, 12, 11), "A"),
    ("IMP-604", "高エネルギー注入 4号機", "日昇イオンテクノロジー", "NIT-HE1", "L5", "高エネルギー注入", "Fab2 3F CR ベイ08", (2019, 8, 5), "B"),
    ("INS-701", "パターン欠陥検査 1号機", "明光インスペクション", "MI-BF500", "L2", "パターン欠陥検査", "Fab1 2F CR 検査ベイ", (2012, 2, 13), "B"),
    ("INS-702", "パーティクル検査 2号機", "明光インスペクション", "MI-DF900", "L3", "パーティクル検査", "Fab1 2F CR 検査ベイ", (2015, 5, 25), "C"),
    ("INS-703", "膜厚測定 3号機", "北辰光学", "HS-TF200", "L4", "膜厚測定", "Fab2 3F CR 検査ベイ", (2018, 6, 18), "C"),
    ("INS-704", "重ね合わせ測定 4号機", "北辰光学", "HS-OVL3", "L5", "重ね合わせ測定", "Fab2 3F CR リソベイ05", (2020, 3, 23), "B"),
    ("INS-705", "パターン欠陥検査 5号機", "明光インスペクション", "MI-BF700", "L6", "パターン欠陥検査", "Fab2 3F CR 検査ベイ", (2025, 2, 3), "B"),
    ("OHT-801", "Fab1 OHTシステム", "大鷹搬送システム", "OT-Track3（車両48台）", "L1-L3", "工程間搬送", "Fab1 2F CR 天井", (2010, 4, 1), "A"),
    ("OHT-802", "Fab2 OHTシステム", "大鷹搬送システム", "OT-Track5（車両40台）", "L4-L6", "工程間搬送", "Fab2 3F CR 天井", (2018, 2, 19), "A"),
    ("AGV-811", "Fab1 AGV（マガジン搬送）", "大鷹搬送システム", "OT-AGV200", "L1-L3", "資材・マガジン搬送", "Fab1 1F 通路", (2016, 10, 3), "C"),
    ("AGV-812", "Fab2 AGV（SLAM型）", "三峰ロボティクス", "MR-NAV30", "L4-L6", "資材・マガジン搬送", "Fab2 1F 通路", (2022, 5, 9), "C"),
    ("ROB-821", "ウェーハソーター 1号機", "三峰ロボティクス", "MR-Sorter25", "L2", "ソーティング", "Fab1 2F CR ベイ12", (2013, 7, 1), "B"),
    ("ROB-822", "ウェーハソーター 2号機", "三峰ロボティクス", "MR-Sorter25N", "L5", "ソーティング", "Fab2 3F CR ベイ19", (2021, 1, 25), "B"),
    ("STK-831", "Fab2 自動倉庫（ストッカ）", "大鷹搬送システム", "OT-Stocker600", "L4-L6", "FOUP保管", "Fab2 3F CR ストッカ室", (2018, 2, 19), "A"),
    ("UPW-901", "超純水製造装置 1系", "水晶アクアシステム", "SA-UPW150", "Fab1共通", "純水供給", "Fab1 B1F ユーティリティ室", (2008, 3, 31), "A"),
    ("UPW-902", "超純水製造装置 2系", "水晶アクアシステム", "SA-UPW200", "Fab2共通", "純水供給", "ユーティリティ棟 1F", (2017, 11, 30), "A"),
    ("EXH-911", "酸排気ファン Fab1", "東邦空調設備", "TKF-EX75", "Fab1共通", "酸排気", "Fab1 屋上", (2008, 3, 31), "A"),
    ("EXH-912", "一般排気ファン Fab2", "東邦空調設備", "TKF-EX90", "Fab2共通", "一般排気", "Fab2 屋上", (2017, 11, 30), "B"),
    ("SCR-921", "燃焼式除害装置 Fab2", "環研エコシステム", "KE-BS400", "Fab2共通", "排ガス除害", "Fab2 1F サブファブ", (2018, 1, 15), "A"),
    ("PCW-931", "冷却水設備（PCW）", "東邦空調設備", "TKF-CT1200", "全Fab共通", "冷却水供給", "ユーティリティ棟 屋外", (2008, 3, 31), "A"),
    ("VAC-941", "ハウス真空ポンプ設備", "精和バキューム", "SV-LR220", "全Fab共通", "ハウス真空", "ユーティリティ棟 1F", (2009, 8, 24), "B"),
]


@lru_cache(maxsize=None)
def _equipment_tuple() -> tuple:
    out = []
    for eid, name, maker, model, line, proc, area, ymd, crit in _EQUIPMENT_ROWS:
        kind = eid.split("-")[0]
        out.append(Equipment(eid, name, CATEGORY_OF_KIND[kind], maker, model, line, proc, area, date(*ymd), crit))
    return tuple(out)


def equipment_master() -> list[Equipment]:
    return list(_equipment_tuple())


def kind_of(eq: Equipment) -> str:
    """設備IDの接頭辞（CMP, OHT, UPW ...）。"""
    return eq.equipment_id.split("-")[0]


def fab_of(eq: Equipment) -> str:
    return "Fab1" if eq.area.startswith("Fab1") or eq.line in ("L1", "L2", "L3", "L1-L3", "Fab1共通") else "Fab2"


def equipment_by_id(eid: str) -> Equipment:
    return {e.equipment_id: e for e in _equipment_tuple()}[eid]


# ---------------------------------------------------------------------------
# マスタ: 人
# ---------------------------------------------------------------------------
_PEOPLE_ROWS = [
    # (社員番号, 氏名, 部, 課・係, 役職)
    ("M10231", "高橋 誠", "製造部", "設備保全課", "課長"),
    ("M10544", "中村 浩二", "製造部", "設備保全課 保全1係", "係長"),
    ("M11302", "小林 大輔", "製造部", "設備保全課 保全1係", "主任"),
    ("M11876", "山口 翔太", "製造部", "設備保全課 保全1係", "主任"),
    ("M12410", "松本 拓也", "製造部", "設備保全課 保全1係", "担当"),
    ("M12655", "井上 亮", "製造部", "設備保全課 保全1係", "担当"),
    ("M13021", "木村 優斗", "製造部", "設備保全課 保全1係", "担当"),
    ("M13388", "林 和也", "製造部", "設備保全課 保全1係", "担当"),
    ("M13702", "清水 彩花", "製造部", "設備保全課 保全1係", "担当"),
    ("M10688", "斎藤 健太郎", "製造部", "設備保全課 保全2係", "係長"),
    ("M11544", "山崎 淳", "製造部", "設備保全課 保全2係", "主任"),
    ("M12127", "森 雄一", "製造部", "設備保全課 保全2係", "担当"),
    ("M12893", "池田 直樹", "製造部", "設備保全課 保全2係", "担当"),
    ("M13155", "橋本 蓮", "製造部", "設備保全課 保全2係", "担当"),
    ("M13520", "阿部 真央", "製造部", "設備保全課 保全2係", "担当"),
    ("M13811", "石川 陸", "製造部", "設備保全課 保全2係", "担当"),
    ("M10902", "前田 修", "製造部", "設備保全課 施設係", "主任"),
    ("M12033", "藤田 剛", "製造部", "設備保全課 施設係", "担当"),
    ("M12978", "岡田 慎吾", "製造部", "設備保全課 施設係", "担当"),
    ("M13640", "後藤 悠真", "製造部", "設備保全課 施設係", "担当"),
    ("M11020", "長谷川 聡", "製造部", "製造課 A班", "班長"),
    ("M12744", "村上 美咲", "製造部", "製造課 A班", "オペレーター"),
    ("M13902", "近藤 大樹", "製造部", "製造課 A班", "オペレーター"),
    ("M11133", "遠藤 隆", "製造部", "製造課 B班", "班長"),
    ("M12810", "青木 由佳", "製造部", "製造課 B班", "オペレーター"),
    ("M13950", "坂本 颯", "製造部", "製造課 B班", "オペレーター"),
    ("M11247", "藤井 和弘", "製造部", "製造課 C班", "班長"),
    ("M12866", "西村 愛", "製造部", "製造課 C班", "オペレーター"),
    ("M14011", "福田 海斗", "製造部", "製造課 C班", "オペレーター"),
    ("M11390", "太田 宏明", "製造部", "製造課 D班", "班長"),
    ("M12921", "三浦 紗希", "製造部", "製造課 D班", "オペレーター"),
    ("M14057", "岡本 陽向", "製造部", "製造課 D班", "オペレーター"),
    ("M10870", "原田 哲也", "製造部", "生産技術課", "主任技師"),
    ("M12302", "藤原 智子", "製造部", "生産技術課", "技師"),
    ("M13266", "中島 啓介", "製造部", "生産技術課", "技師"),
    ("M11655", "小川 恵", "品質保証部", "品質保証課", "主任"),
    ("M13477", "竹内 俊", "品質保証部", "品質保証課", "担当"),
    ("FE-0107", "金子 正樹", "設備メーカー", "扶桑ポリッシュシステム", "FE"),
    ("FE-0213", "上田 光", "設備メーカー", "昭栄プラズマテック", "FE"),
    ("FE-0320", "野口 貴之", "設備メーカー", "瑞穂エンジニアリング", "FE"),
    ("FE-0431", "丸山 司", "設備メーカー", "光陽オプトロニクス", "FE"),
    ("FE-0615", "久保 翼", "設備メーカー", "日昇イオンテクノロジー", "FE"),
]


@lru_cache(maxsize=None)
def _people_tuple() -> tuple:
    return tuple(Person(*row) for row in _PEOPLE_ROWS)


def people() -> list[Person]:
    return list(_people_tuple())


def surname(p: Person) -> str:
    return p.name.split(" ")[0]


# ---------------------------------------------------------------------------
# マスタ: 部品・アラーム
# ---------------------------------------------------------------------------
_SUBSYS_CODE = {}


@lru_cache(maxsize=None)
def _parts_tuple() -> tuple:
    rows = _bank_front.PARTS + _bank_mid.PARTS + _bank_infra.PARTS
    out, seen = [], set()
    r = rng("parts")
    for i, (name, maker, price, sub, cats) in enumerate(rows):
        if name in seen:
            continue
        seen.add(name)
        prefix = {"CMP装置": "PC", "CVD装置": "PV", "エッチング装置": "PE", "露光装置": "PL", "洗浄装置": "PW",
                  "イオン注入装置": "PI", "検査装置": "PM", "搬送系": "PT", "ユーティリティ": "PU"}[cats[0]]
        part_no = f"{prefix}{r.randint(10, 99)}-{1000 + i * 7 + r.randint(0, 6):04d}"
        if len(cats) > 5:
            part_no = f"PX{r.randint(10, 99)}-{1000 + i * 7:04d}"
        out.append(Part(part_no, name, maker, price, sub, tuple(cats)))
    return tuple(out)


def parts_catalog() -> list[Part]:
    return list(_parts_tuple())


@lru_cache(maxsize=None)
def _part_by_name() -> dict:
    return {p.name: p for p in _parts_tuple()}


_ALARM_SOURCES = {**_bank_front.ALARMS, **_bank_mid.ALARMS, **_bank_infra.ALARMS}


@lru_cache(maxsize=None)
def _alarms_tuple() -> tuple:
    out = []
    for kind, rows in _ALARM_SOURCES.items():
        for code, msg, cat, sub in rows:
            out.append(AlarmCode(code, msg, cat, sub))
    return tuple(out)


def alarm_codes() -> list[AlarmCode]:
    return list(_alarms_tuple())


@lru_cache(maxsize=None)
def _alarm_by_code() -> dict:
    return {a.code: a for a in _alarms_tuple()}


def alarm_codes_for(eq: Equipment) -> list[AlarmCode]:
    """設備に表示されうるアラーム（種別固有＋共通）。"""
    k = kind_of(eq)
    codes = [c for c, *_ in _ALARM_SOURCES.get(k, [])]
    if k not in UTILITY_KINDS:
        codes += [c for c, *_ in _ALARM_SOURCES["COMMON"]]
    return [_alarm_by_code()[c] for c in codes]


# ---------------------------------------------------------------------------
# シナリオ・QCテンプレートの参照
# ---------------------------------------------------------------------------
_SCEN = {**_bank_front.SCENARIOS, **_bank_mid.SCENARIOS, **_bank_infra.SCENARIOS}
_MINOR = {**_bank_front.MINOR, **_bank_mid.MINOR, **_bank_infra.MINOR}
_QC = {**_bank_front.QC, **_bank_mid.QC, **_bank_infra.QC}


def _qc_params(kind: str, r: random.Random) -> dict:
    if kind in _bank_front.QC:
        return _bank_front.QC_PARAMS(r)
    if kind in _bank_mid.QC:
        return _bank_mid.qc_params(r)
    return _bank_infra.qc_params(r)


def _scenarios_for(eq: Equipment) -> list[dict]:
    k = kind_of(eq)
    common = _SCEN["COMMON"]
    if k in UTILITY_KINDS:
        common = []          # 共通シナリオはウェーハ処理を前提とした文面のため、ユーティリティには適用しない
    elif k in TRANSPORT_KINDS:
        common = [s for s in common if s["id"] != "COM.misop"]
    return _SCEN[k] + common


# ---------------------------------------------------------------------------
# カレンダー・シフト
# ---------------------------------------------------------------------------
def _is_holiday(d: date) -> bool:
    if d.weekday() >= 5:
        return True
    md = (d.month, d.day)
    return md in {(1, 1), (1, 2), (1, 3), (12, 29), (12, 30), (12, 31), (4, 29), (5, 3), (5, 4), (5, 5),
                  (8, 13), (8, 14), (8, 15), (8, 16), (2, 11), (11, 3), (11, 23)}


def shift_of(dt: datetime) -> str:
    """日勤 8:00-20:00 / 夜勤 20:00-翌8:00。土日・連休は休日。"""
    base = dt.date() if dt.hour >= 8 else dt.date() - timedelta(days=1)
    if _is_holiday(base):
        return "休日"
    return "日勤" if 8 <= dt.hour < 20 else "夜勤"


def crew_of(dt: datetime) -> str:
    """4班2交替: 2日ごとに日勤・夜勤の班が入れ替わる。"""
    base = dt.date() if dt.hour >= 8 else dt.date() - timedelta(days=1)
    k = (base - date(2023, 1, 1)).days // 2
    return "ABCD"[k % 4] if 8 <= dt.hour < 20 else "ABCD"[(k + 2) % 4]


# ---------------------------------------------------------------------------
# テンプレート展開
# ---------------------------------------------------------------------------
class _Params(dict):
    missing: set = set()

    def __missing__(self, key):
        _Params.missing.add(key)
        return ""


def _fill(tpl: str, p: dict) -> str:
    return tpl.format_map(p) if "{" in tpl else tpl


# 英字・ハイフン・ドットに隣接しない数字（設備ID・ロットID・アラームコードは半角のまま残す）
_ZEN_DIGITS_RE = re.compile(r"(?<![A-Za-z0-9\-_.,])\d+(?:[.,]\d+)*(?![A-Za-z0-9\-_.,])")
_TYPO_PAIRS = [("異常", "以上"), ("回収", "改修"), ("規定", "既定"), ("機械", "機会"), ("保証", "保障"), ("確認", "確認確認")]


@lru_cache(maxsize=None)
def _writer_habit(employee_id: str) -> dict:
    """記入者ごとの書き癖（敬体・全角数字・半角カナ・番号の振り方・誤変換率）。"""
    r = rng(f"writer:{employee_id}")
    return dict(
        polite=r.random() < 0.22,
        zen=r.choice([0.0, 0.0, 0.0, 0.0, 0.25, 0.6]),
        hkana=r.random() < 0.12,
        numbering=r.choice(["1. ", "1. ", "①", "(1)", "1)", "・"]),
        typo=r.choice([0.0, 0.0, 0.01, 0.03]),
        terse=r.random() < 0.3,
        arrow=r.choice(["→", "、", "。", " → "]),
    )


_POLITE_RULES = [
    (re.compile(r"(確認|実施|交換|清掃|調整|回収|手配|判断|推定|連絡|測定|依頼|送付|共有|再開|停止|復帰)(。)"), r"\1しました\2"),
    (re.compile(r"あり。"), "ありました。"), (re.compile(r"なし。"), "ありませんでした。"),
    (re.compile(r"られた。"), "られました。"), (re.compile(r"んだ。"), "みました。"), (re.compile(r"していた。"), "していました。"), (re.compile(r"だった。"), "でした。"), (re.compile(r"いた。"), "いました。"),
]


def _season_ok(text: str, month: int) -> bool:
    """季節に依存する文（夏場・冬季など）を、合わない月に使わない。"""
    if any(w in text for w in ("夏場", "夏季", "猛暑", "外気温上昇")):
        return month in (6, 7, 8, 9)
    if any(w in text for w in ("冬季", "冬場")):
        return month in (11, 12, 1, 2, 3)
    return True


def _polite(s: str) -> str:
    for pat, rep in _POLITE_RULES:
        s = pat.sub(rep, s)
    return s


def _noise(s: str, habit: dict, r: random.Random) -> str:
    """記入者の癖に応じて全角数字化・半角カナ・誤変換を混ぜる。"""
    if not s:
        return s
    if habit["zen"] and r.random() < habit["zen"]:
        s = _ZEN_DIGITS_RE.sub(lambda m: to_zenkaku(m.group(0), digits_only=True), s)
    if habit["hkana"] and r.random() < 0.5:
        s = to_hankaku_kana(s)
    if habit["typo"] and r.random() < habit["typo"] * 4:
        a, b = r.choice(_TYPO_PAIRS)
        if a in s:
            s = s.replace(a, b, 1)
    return s


def _numbered(steps: list[str], style: str) -> str:
    out = []
    for i, s in enumerate(steps, 1):
        if style == "①":
            out.append(f"{chr(0x2460 + i - 1)}{s}")
        elif style == "(1)":
            out.append(f"({i}) {s}")
        elif style == "1)":
            out.append(f"{i}) {s}")
        elif style == "・":
            out.append(f"・{s}")
        else:
            out.append(f"{i}. {s}")
    return "\n".join(out)


def _end(s: str) -> str:
    s = s.rstrip()
    return s if s.endswith(("。", "）", ")", "!", "？")) else s + "。"


# ---------------------------------------------------------------------------
# 共通フレーズバンク
# ---------------------------------------------------------------------------
_SYM_PREFIX = {
    "any": ["{time}頃、", "{lot}処理中、", "PM明け{n}ロット目で", "立上げ直後、", "", "", "連休明けの稼働開始時、", "レシピ切替後、"],
    "日勤": ["朝の巡回時、", "日勤帯、", "昼休憩明けに"],
    "夜勤": ["夜勤帯、", "深夜{time}頃、", "明け方、"],
    "休日": ["休日出勤中に", "休日の{time}頃、"],
}
_SYM_SUFFIX = ["。前日も同様のワーニングが{cnt}回出ていた", "。リセット1回では復帰せず", "（同型機{sib}は正常）", "。再現性あり",
               "。現象は断続的", "。{op}さんより保全へ連絡", "。直前にアラーム履歴なし", "。スループットへの影響大",
               "。処理中ウェーハ{n}枚あり", "", "", "", "", ""]
_FIRST = {
    "alarm": ["アラーム内容を確認し装置をDOWN登録", "アラームリセット1回実施、再発したため保全コール", "リセットで一時復帰したが{cnt}分後に再発",
              "画面のアラーム履歴を写真に記録", "装置を停止し保全へPHS連絡", "生産管理へダウン連絡", "アラームリセット不可、保全呼出し",
              "処理中ウェーハの状態をモニタで確認"],
    "quality": ["該当ロット{lot}をHOLD", "品証・生産技術へ連絡", "装置を着工停止（インヒビット設定）", "同時期処理ロットをリストアップ",
                "SPCチャートと処理履歴を照合"],
    "patrol": ["現場で状態を確認し写真撮影", "装置を停止し保全へ連絡", "周辺の立入り制限", "運転状態を記録し班長へ報告"],
    "util": ["中央監視でアラーム内容を確認", "現地へ急行し状態確認", "予備系統への切替を確認", "関係装置の担当へ一斉連絡", "施設係長へ報告"],
    "tail": ["仕掛りを{sib}へ振替", "LOTO実施", "処理中ロット{lot}を退避", "保全{m}が{rmin}分後に到着", "班長へ報告", "引継ぎ簿に記入"],
}
_GENERIC_INV = [
    "アラーム履歴を確認、同アラームは直近{days}日で{cnt}回発生していた。",
    "前回PMは{pm_days}日前に実施、作業記録に特記事項なし。",
    "同型機{sib}のパラメータと比較し、設定値の差異はなし。",
    "FDCトレンドを遡ると{days}日前から兆候が見られた。",
    "装置ログを回収し{fe}へ送付、解析を依頼。",
    "目視点検では外観上の異常は見当たらず、計測で確認した。",
    "作業前に装置停止・LOTOを実施し安全を確保してから点検。",
    "過去の類似トラブル事例を保全DBで検索し、点検箇所を絞り込んだ。",
    "班長への聞き取りでは、発生前に特段の作業はなかった。",
    "発生直前の処理レシピ・ロット情報に異常なし。",
]
_RESULT = {
    "ok": ["{qc}。{done}に生産復帰", "処置後、{qc}。生産へ引渡し", "{qc}。{done}より通常稼働", "復旧後{n}ロット流動し再発なし。{qc}",
           "処置後の確認で異常なし（{qc}）", "{done}復旧。その後{cnt}時間連続稼働し問題なし"],
    "watch": ["暫定処置で{done}に生産復帰。{days}日間トレンド監視とする", "復帰後、再発有無を経過観察中（{days}日間）", "{qc}。ただし原因特定に至らず、再発時はログ取得予定"],
    "hold": ["暫定処置で生産再開。部品入荷待ち（納期{days}日）", "代替運用で生産継続中。{fe}回答待ち", "応急処置のみ実施。恒久処置は次回PM（{days}日後）で実施予定"],
    # 発生から半年以上たっても保留のもの（部品待ち・次回PM待ちでは不自然なので、長期の計画待ちとして書く）
    "hold_long": ["暫定処置で運用継続中。恒久対策は設備更新計画（{fy}年度）で実施予定", "代替条件で運用継続（技術承認済み）。{fe}の改造案の回答待ち",
                  "暫定処置で生産継続。恒久処置は年次の全停止PMで実施予定"],
}
_PREV_GENERIC = ["点検チェックシートに反映", "保全カレンダーへ登録済み", "保全課内の朝会で事例共有", "作業標準書を改訂（改訂番号 Rev.{n}）",
                 "FDC監視項目の追加を生産技術へ依頼", "予備品在庫の見直し"]
_HORIZ = ["同型機（{sibs}）の同部位を点検、異常なし", "{sibs}にも次回PM（{month}月）で展開予定", "本機固有の事象のため展開不要",
          "保全1係・2係の定例会で共有", "他Fabの同型機へ展開済み", "{sibs}は{nday}日までに点検予定（担当 {m}）", "同型機なし。類似構造の設備で点検実施",
          "水平展開要否を検討中", ""]
# 再発防止策が他号機・全号機に及んでいるのに「展開不要」と書くのは矛盾するので、この語があれば展開不要を選ばない
_HORIZ_WIDE_WORDS = ("全号機", "全車両", "水平展開", "同型機", "他号機")
_REMARKS = ["部品在庫残{n}個、発注済み", "FE作業費は保守契約内", "夜勤→日勤へ引継ぎ（{op}）", "報告書記入が{days}日遅れ", "品証へ連絡済み（{qa}）",
            "部品は{sib}の予備品から転用、後日補充", "メーカー見積依頼中", "安全上の問題なし", "写真は共有フォルダ保全\\{year}\\{eq}に保存",
            "暫定対策のため恒久対策は別途報告", "作業時間には待機時間を含む", "チョコ停の段階で予兆あり（チョコ停記録参照）", "月例会議で報告予定"]
# メーカーが関わった案件（担当にFEがいる・処置にFE/メーカー名が出る）でしか書かない備考
_REMARKS_NEED_MAKER = ("FE作業費は保守契約内", "メーカー見積依頼中")
# 部品を使った案件でしか書かない備考
_REMARKS_NEED_PARTS = ("部品在庫残{n}個、発注済み", "部品は{sib}の予備品から転用、後日補充")

# なぜなぜ分析の最後（仕組み・管理面の要因）。原因分類ごとに候補を持ち、1つの言い回しに偏らないようにする
_WHY_TAIL_COMMON = [
    "点検周期が設定されていなかった", "異常を早期に検知する監視項目が無かった", "作業標準書に確認手順が記載されていなかった",
    "前任者からの引継ぎで管理項目が抜け落ちていた", "過去の類似事例が共有されていなかった", "管理基準・点検項目への落とし込みができていなかった",
]
_WHY_TAIL = {
    "摩耗": ["使用時間・処理枚数による交換基準が無かった", "摩耗量の測定方法と判定基準が決まっていなかった",
           "生産優先で交換が先送りされ、その判断基準が無かった", "予備品の在庫が無く交換時期を延ばしていた",
           "メーカー推奨の交換周期を自社の稼働率で見直していなかった", "トレンドデータを定期的に確認する仕組みが無かった",
           "購買品の仕様変更（材質変更）の影響を評価していなかった", "PM延期時の代替確認ルールが無かった"],
    "劣化": ["経年部品の寿命管理台帳が整備されていなかった", "予防交換の計画・予算化がされていなかった", "劣化傾向を監視する項目が無かった",
           "メーカー推奨の交換時期を把握していなかった", "設置からの使用年数を管理していなかった",
           "購買品の仕様変更（ロット変更）の影響を評価していなかった", "トレンドデータを定期的に確認する仕組みが無かった",
           "点検結果の判定基準が担当者任せだった"],
    "汚れ": ["清掃周期が稼働実態（処理枚数）に合っていなかった", "清掃範囲が作業標準書に明記されていなかった",
           "汚れを検知する手段（差圧・光量の監視）が無かった", "清掃後の確認基準が決まっていなかった",
           "レシピ変更による汚れ方の変化を評価していなかった", "PM延期時の代替確認ルールが無かった", "清掃作業の時間が計画に確保されていなかった"],
    "異物": ["異物の侵入経路が特定・管理されていなかった", "部品開梱・取付場所のルールが無かった", "立上げ時の異物確認が手順に無かった",
           "上流側フィルタ・ストレーナの点検周期が無かった", "購買品の梱包仕様変更の影響を確認していなかった",
           "PM作業後の清掃確認が担当者任せだった", "異物混入時の検知手段が無かった"],
    "締結緩み": ["締付けトルクと合いマークの基準が無かった", "PM・工事後の増締め確認が手順に無かった", "振動部位の定期増締めが点検項目に無かった",
             "工事完了時の検査項目に漏れ確認が無かった", "作業者のダブルチェックが形骸化していた", "協力会社作業の品質確認ルールが無かった"],
    "断線": ["屈曲部の配線方法の基準が無かった", "ケーブルの交換周期が設定されていなかった", "導通・絶縁抵抗の定期測定が点検項目に無かった",
           "断線の予兆（抵抗値・電流値）を監視していなかった", "耐屈曲ケーブルへの仕様変更を検討していなかった",
           "経年部品の寿命管理台帳が整備されていなかった", "設置時の配線施工の検査基準が無かった"],
    "調整不良": ["調整値の管理幅が作業標準書に無かった", "調整後の確認者（ダブルチェック）が決まっていなかった",
             "PM後の立上げ確認項目に位置確認が無かった", "調整作業のスキル認定が無かった", "部品交換後の再ティーチング手順が標準化されていなかった",
             "調整記録を残す様式が無かった"],
    "設定ミス": ["パラメータ変更時の承認フローが無かった", "設定値のバックアップと照合の手順が無かった", "変更履歴を確認する仕組みが無かった",
             "共用IDで操作しており変更者が特定できなかった", "変更管理（4M変更）の手続きが形骸化していた",
             "設定変更時の注意点が教育に含まれていなかった"],
    "作業ミス": ["作業手順書が現場の実態と合っていなかった", "ダブルチェックのルールが守られていなかった", "繁忙時の作業人員の配置基準が無かった",
             "当該作業の教育・認定がされていなかった", "作業の引継ぎが口頭のみだった", "ヒヤリハットの情報が共有されていなかった",
             "夜勤帯の確認体制が手薄だった"],
    "施工不良": ["施工業者へ作業基準を提示していなかった", "施工後の立会検査が無かった", "協力会社作業の品質確認ルールが無かった",
             "施工記録（写真）を残す決まりが無かった"],
    "ソフト不具合": ["メーカーの既知不具合情報を入手する仕組みが無かった", "パッチ適用の計画が立てられていなかった",
               "ソフト更新時の受入れ試験項目が不足していた", "ログ領域・リソースの監視をしていなかった", "定期再起動の運用が決まっていなかった"],
    "設計起因": ["導入時の仕様検討で使用条件が考慮されていなかった", "メーカーの設計変更情報を入手していなかった",
             "同型機の不具合情報がメーカーから展開されていなかった", "耐久性評価を購入仕様に含めていなかった"],
    "能力不足": ["設備能力の余裕度を定期的に見直していなかった", "増産時の能力評価をしていなかった", "夏季ピーク時の運用基準が無かった"],
    "前工程起因": ["前工程との異常連絡ルールが無かった", "前工程の変更情報が共有されていなかった", "受入れ時の確認項目に該当特性が無かった"],
    "外部要因": ["瞬低時の復旧手順が整備されていなかった", "UPS対象範囲の見直しがされていなかった"],
}
# 原因が準備された連鎖と合わないときに組み立てる連鎖の3段目（原因のすぐ上の要因）
_WHY_MID = {
    "摩耗": ["交換時期を過ぎて使用していた", "摩耗の進行に気付けなかった", "交換基準が明確でなかった"],
    "劣化": ["劣化の兆候を監視していなかった", "予防交換の対象になっていなかった", "交換基準が明確でなかった"],
    "汚れ": ["点検・清掃の周期が実態に合っていなかった", "汚れの付着に気付けなかった"],
    "異物": ["異物の混入を防げていなかった", "異物の付着に気付けなかった"],
    "締結緩み": ["締結状態を確認していなかった", "緩みの兆候に気付けなかった"],
    "断線": ["屈曲・擦れの状態を点検していなかった", "断線の予兆に気付けなかった"],
    "調整不良": ["調整後の確認が不十分だった", "位置ずれに気付けなかった"],
    "施工不良": ["施工後の確認が不十分だった"],
    "ソフト不具合": ["既知不具合の対策が未適用だった"],
    "能力不足": ["設備能力に余裕が無かった"],
}
_WHY_MID_DEFAULT = ["点検・清掃の周期が実態に合っていなかった", "劣化の兆候を監視していなかった", "交換基準が明確でなかった",
                    "日常点検で異常に気付けなかった"]


# ---------------------------------------------------------------------------
# トラブル生成
# ---------------------------------------------------------------------------
_KIND_RATE = {"CMP": 1.35, "CVD": 1.1, "ETC": 1.2, "LIT": 1.0, "CLN": 0.95, "IMP": 1.05, "INS": 0.55, "OHT": 1.25, "AGV": 0.8,
              "ROB": 0.75, "STK": 0.7, "UPW": 0.75, "EXH": 0.5, "SCR": 0.7, "PCW": 0.7, "VAC": 0.45}
# 持病のある号機: (持病シナリオ, 発生率倍率, 対策完了日)
_BAD_UNITS = {
    "CMP-103": (("CMP.head.membrane", "CMP.slurry.clog"), 2.0, date(2026, 12, 31)),
    "CVD-205": (("CVD.dp.overload",), 1.8, date(2025, 4, 15)),
    "ETC-302": (("ETC.esc.he", "ETC.chiller"), 1.9, date(2026, 12, 31)),
    "LIT-402": (("LIT.stage.servo",), 1.5, date(2025, 9, 30)),
    "OHT-801": (("OHT.wheel", "OHT.hoist.belt"), 1.4, date(2026, 12, 31)),
    "CLN-504": (("CLN.nozzle.drip",), 1.7, date(2024, 12, 20)),
    "IMP-603": (("IMP.hv.arc",), 1.8, date(2026, 12, 31)),
}
_SEVERITIES = ("重大", "大", "中", "小")
_YEAR_RATE = {2023: 0.94, 2024: 1.0, 2025: 1.05, 2026: 1.08}


def _age_years(eq: Equipment, d: date) -> float:
    return (d - eq.installed_date).days / 365.25


def _scenario_weights(eq: Equipment, month: int, age: int, bad_active: bool) -> list[float]:
    ws = []
    bad = _BAD_UNITS.get(eq.equipment_id)
    for s in _scenarios_for(eq):
        w = s["w"]
        if s["season"]:
            w *= s["season"].get(month, 1.0)
        if s["aging"]:
            w *= 1.0 + 0.06 * max(age, 0)
        if bad and bad_active and s["id"] in bad[0]:
            w *= 2.2
        ws.append(w)
    return ws


@lru_cache(maxsize=None)
def _scen_w_cached(eid: str, month: int, age: int, bad_active: bool) -> tuple:
    ws = _scenario_weights(equipment_by_id(eid), month, age, bad_active)
    return tuple(ws), sum(ws)


def _day_eq_weight(eq: Equipment, d: date) -> tuple[float, bool]:
    if d < eq.installed_date:
        return 0.0, False
    age = _age_years(eq, d)
    w = _KIND_RATE[kind_of(eq)] * _YEAR_RATE[d.year]
    w *= 1.0 + 0.025 * min(age, 20)
    if age < 0.5:
        w *= 2.2            # 初期故障期
    elif age < 1.0:
        w *= 1.4
    bad = _BAD_UNITS.get(eq.equipment_id)
    bad_active = bool(bad and d <= bad[2])
    if bad_active:
        w *= bad[1]
    if _is_holiday(d):
        w *= 0.88
    if d.month in (7, 8):
        w *= 1.08
    base_sum = _scen_w_cached(eq.equipment_id, 1, 0, False)[1]
    month_sum = _scen_w_cached(eq.equipment_id, d.month, int(age), bad_active)[1]
    return w * (month_sum / base_sum), bad_active


_HOUR_W = [3, 3, 2.6, 2.4, 2.4, 2.6, 3, 3.6, 5, 5.6, 5.4, 5.0, 4.2, 4.8, 5.2, 5.0, 4.6, 4.2, 3.8, 3.6, 3.8, 3.6, 3.4, 3.2]


def _lot_id(r: random.Random, d: date) -> str:
    prod = r.choice(["MK", "SR", "PX", "TQ", "NB", "HV"])
    wk = d.isocalendar()[1]
    return f"{prod}{d.year % 100:02d}{wk:02d}{r.randint(1, 399):03d}.{r.randint(1, 3)}"


def _sample_severity(r, s, eq) -> str:
    w = s["sev"]
    sev = r.choices(_SEVERITIES, weights=(w[0] * 0.45, w[1] * 0.6, w[2], w[3] * 1.2))[0]
    i = _SEVERITIES.index(sev)
    if eq.criticality == "A" and i > 0 and r.random() < 0.12:
        i -= 1
    elif eq.criticality == "C" and i < 3 and r.random() < 0.2:
        i += 1
    return _SEVERITIES[i]


_DT_MEDIAN = {"小": (35, 0.6, 8, 240), "中": (170, 0.55, 40, 720), "大": (600, 0.5, 180, 2400), "重大": (1900, 0.5, 600, 10080)}
_SEV_Q = {"小": 0.45, "中": 0.9, "大": 1.25, "重大": 1.6}


@dataclass
class _Draft:
    eq: Equipment
    s: dict
    at: datetime
    master: "_Draft | None" = None
    prev: "_Draft | None" = None
    chronic_n: int = 0
    inc: Incident | None = None
    group_n: int = 1          # 瞬低などで同時に記録された件数（親の行に、自分を含めた件数を持つ）
    ntool_fab: int = 0        # 瞬低で工場内で止まった台数（同時停止グループで共通の値。親に持つ）


# ---------------------------------------------------------------------------
# 使用量カウンタ（運転時間・RF積算・処理枚数・充放電サイクル・経過日数）
# ---------------------------------------------------------------------------
# トラブル記録の「TMP運転○h」「メンブレン使用○枚」などは、トラブルごとの乱数ではなく部品ごとの状態から決める。
#   値 = 1日あたりの増分 × (発生日時 − 使用開始)
#   使用開始 = 設置日 / その部品を交換したトラブルの完了日時 / 定期交換（T4 の PM 計画）の日 のうち最新
# 交換の記録が無い間は値が増え続ける（前回より小さい値になるのは、間に交換がある場合だけ）。
# 設置後の交換記録が無い部品は、初めて記録に出るときの値を first の範囲（設置からの経過が上限）で決め、以後はそこから積み上げる。
@dataclass(frozen=True)
class _Counter:
    key: str
    rate: tuple               # 1日あたりの増分の範囲（設備ごとに一様乱数で固定）
    first: tuple              # 交換記録の無い部品が初めて記録に出るときの値の範囲
    parts: tuple = ()         # この語を含む部品を交換したらリセット（トラブルの使用部品・T4 の定期交換）
    scen: tuple = ()          # このシナリオの処置が完了したらリセット（部品を使わないウェットクリーニングなど）
    scope: str = ""           # 値を分ける単位（パラメータ名）: ch チャンバー / head ヘッド / pl プラテン / veh 号車 / gas ガス。空は設備単位
    rate_key: str = ""        # 増分を共有するカウンタ群（同じチャンバーの RF 放電時間など）
    kinds: tuple = ()         # 対象の設備種別


_COUNTERS = {c.key: c for c in [
    # CMP（枚数・時間はプラテン/ヘッド単位。PM は T4 の定期交換）
    _Counter("pou", (1.0, 1.0), (5, 30), parts=("POUフィルタ",), kinds=("CMP",)),                               # 日
    _Counter("pad", (0, 0), (150, 1050), parts=("研磨パッド",), scope="pl", kinds=("CMP",)),                     # 枚（増分は寿命設定から）
    _Counter("dresser", (3.0, 4.5), (60, 380), parts=("ドレッサディスク",), scope="pl", kinds=("CMP",)),           # h
    _Counter("membrane", (70, 95), (900, 5200), parts=("ヘッドメンブレン",), scope="head", kinds=("CMP",)),        # 枚
    _Counter("retainer", (70, 95), (1500, 8000), parts=("リテーナリング",), scope="head", kinds=("CMP",)),         # 枚
    # CVD
    _Counter("cvd_rf", (8, 14), (5000, 25000), parts=("真空コンデンサ",), scope="ch", kinds=("CVD",)),             # h
    _Counter("cvd_clean", (12, 30), (300, 2400), parts=("チャンバーOリング (FKM)",), scen=("CVD.cham.particle",),
             scope="ch", kinds=("CVD",)),                                                                      # μm（ウェットクリーニング後の成膜量）
    _Counter("cvd_oring", (1.0, 1.0), (20, 85), parts=("チャンバーOリング",), scope="ch", kinds=("CVD",)),          # 日
    _Counter("dp", (21.0, 23.5), (9000, 20000), parts=("ドライポンプ",), scope="ch", kinds=("CVD",)),              # h
    _Counter("mfc", (1 / 365.25, 1 / 365.25), (1.0, 15.0), parts=("MFC (",), scope="gas", kinds=("CVD", "ETC")),  # 年
    # ETC（RF 放電時間はチャンバーごとに共通の増分）
    _Counter("esc", (10, 16), (1500, 6000), parts=("静電チャック",), scope="ch", rate_key="etc_rf", kinds=("ETC",)),
    _Counter("etc_match", (10, 16), (1000, 6000), parts=("真空コンデンサ",), scope="ch", rate_key="etc_rf", kinds=("ETC",)),
    _Counter("electrode", (10, 16), (300, 2800), parts=("上部電極",), scope="ch", rate_key="etc_rf", kinds=("ETC",)),
    _Counter("focus_ring", (10, 16), (100, 900), parts=("フォーカスリング",), scope="ch", rate_key="etc_rf", kinds=("ETC",)),
    _Counter("tmp", (21.0, 23.8), (15000, 35000), parts=("ターボ分子ポンプ",), scope="ch", kinds=("ETC",)),         # h
    # 露光・検査・洗浄・注入
    _Counter("laser", (2.0, 3.5), (800, 1900), parts=("レーザチャンバ",), kinds=("LIT",)),                        # 百万ショット
    _Counter("alg_lamp", (20, 24), (800, 4000), parts=("アライメント照明ランプ",), kinds=("LIT",)),                 # h
    _Counter("ins_lamp", (18, 24), (400, 2400), parts=("検査光源ランプ",), kinds=("INS",)),                        # h
    _Counter("claw", (1.0, 1.0), (20, 170), parts=("ロボットチャック爪",), kinds=("CLN",)),                         # 日
    _Counter("filament", (10, 16), (40, 420), parts=("フィラメント",), kinds=("IMP",)),                            # h
    # 搬送・ユーティリティ
    _Counter("hoist_km", (35, 60), (8000, 36000), parts=("ホイストベルト",), scope="veh", kinds=("OHT",)),          # km
    _Counter("agv_battery", (3.0, 5.0), (600, 2400), parts=("AGVバッテリー",), scope="veh", kinds=("AGV",)),        # サイクル
    _Counter("resin", (12 / 365.25, 12 / 365.25), (2, 12), parts=("イオン交換樹脂",), kinds=("UPW",)),              # 月
    _Counter("uv", (23.0, 24.0), (1000, 8000), parts=("UVランプ",), kinds=("UPW",)),                              # h
    _Counter("upw_pump", (12, 20), (18000, 45000), kinds=("UPW",)),                                             # h（ポンプ本体の累計）
]}
# シナリオ -> 文に使うカウンタ
_SCEN_COUNTERS = {
    "CMP.slurry.clog": ("pou",), "CMP.pad.scratch": ("pad",), "CMP.cond.torque": ("dresser",), "CMP.head.membrane": ("membrane",),
    "CMP.head.retainer": ("retainer",), "CVD.mfc.drift": ("mfc",), "ETC.mfc": ("mfc",), "CVD.rf.reflect": ("cvd_rf",),
    "CVD.cham.particle": ("cvd_clean",), "CVD.cham.leak": ("cvd_oring",), "CVD.dp.overload": ("dp",), "ETC.esc.he": ("esc",),
    "ETC.rf.match": ("etc_match",), "ETC.cham.arc": ("electrode",), "ETC.fr.wear": ("focus_ring",), "ETC.tmp": ("tmp",),
    "LIT.laser.energy": ("laser",), "LIT.align.lamp": ("alg_lamp",), "INS.lamp": ("ins_lamp",), "CLN.robot.claw": ("claw",),
    "IMP.filament": ("filament",), "OHT.hoist.belt": ("hoist_km",), "AGV.battery": ("agv_battery",), "UPW.resin": ("resin",),
    "UPW.uv": ("uv",), "UPW.pump": ("upw_pump",),
}
_AGV_FLEET = {"AGV-811": 6, "AGV-812": 8}      # AGV システムごとの車両数


@lru_cache(maxsize=None)
def _pad_life(eid: str) -> int:
    """CMP 研磨パッドの寿命設定（枚）。設備ごとに固定。"""
    return rng(f"pad-life:{eid}").choice([900, 1000, 1100, 1200])


@lru_cache(maxsize=None)
def _laser_life(eid: str) -> int:
    """露光装置のレーザチャンバ寿命目安（億ショット）。設備ごとに固定。"""
    return rng(f"laser-life:{eid}").choice([30, 40, 50])


@lru_cache(maxsize=None)
def _counter_rate(eid: str, key: str) -> float:
    c = _COUNTERS[key]
    r = rng(f"counter-rate:{eid}:{c.rate_key or key}")
    if key == "pad":
        return _pad_life(eid) / 30 * r.uniform(0.9, 1.15)     # 月1回の定期交換で寿命設定前後になる使い方
    return r.uniform(*c.rate)


@lru_cache(maxsize=None)
def _pm_replacements() -> dict:
    """定期交換（PM）の実施日時。T4 保全作業記録の PM 計画を正とし、2022年1月から期間末まで作る。

    戻り値: 設備ID -> [(部品名, datetime), ...]
    """
    from . import t4_maintenance_blocks as t4          # t4 は domain を import するので、使うときに読み込む

    months, y, m = [], 2022, 1
    while (y, m) <= (PERIOD_END.year, PERIOD_END.month):
        months.append((y, m))
        y, m = (y, m + 1) if m < 12 else (y + 1, 1)
    out: dict = defaultdict(list)
    for w in t4._pm_rows(months):
        if w.part is not None:
            out[w.eq.equipment_id].append((w.part.name, datetime(w.work_date.year, w.work_date.month, w.work_date.day, w.hour)))
    return dict(out)


@lru_cache(maxsize=None)
def _pm_times(eid: str, key: str) -> tuple:
    """カウンタ key の部品を定期交換した日時（昇順）。搬送車両の PM 行は号車と対応しないので使わない。"""
    c = _COUNTERS[key]
    if c.scope == "veh" or not c.parts:
        return ()
    return tuple(sorted(t for name, t in _pm_replacements().get(eid, []) if any(w in name for w in c.parts)))


class _CounterBook:
    """使用量カウンタの状態。トラブルを発生日時順に作りながら、部品を交換した日時を積み上げる。"""

    def __init__(self):
        self.resets: dict = defaultdict(list)    # (設備ID, key, sub) -> [交換日時]（昇順）。sub "*" は全チャンバー/全ヘッド
        self.base: dict = {}                     # (設備ID, key) -> 設置後の交換記録が無い部品の使用開始日時

    def value(self, eq: Equipment, key: str, sub: str, at: datetime) -> tuple[float, str]:
        """発生日時 at の値と、使用開始の由来（install=設置時から / carry=期間前に交換済み / pm=定期交換 / repair=トラブル対応で交換）。"""
        c = _COUNTERS[key]
        eid = eq.equipment_id
        inst = datetime(eq.installed_date.year, eq.installed_date.month, eq.installed_date.day, 9)
        start, origin = inst, "install"
        pm = _pm_times(eid, key)
        i = bisect.bisect_right(pm, at)
        if i and pm[i - 1] > start:
            start, origin = pm[i - 1], "pm"
        for s in (sub, "*"):
            lst = self.resets.get((eid, key, s), [])
            j = bisect.bisect_right(lst, at)
            if j and lst[j - 1] > start:
                start, origin = lst[j - 1], "repair"
        rate = _counter_rate(eid, key)
        if origin == "install":
            b = self.base.get((eid, key))
            if b is None:
                ar = rng(f"counter-base:{eid}:{key}")
                span = (at - inst).total_seconds() / 86400
                v0 = ar.uniform(*c.first)
                if key == "mfc" and ar.random() < 0.45:
                    b = inst                                            # 導入時から一度も交換していない
                elif v0 / rate >= span * 0.95:
                    b = inst + (at - inst) * ar.uniform(0.0, 0.4)      # 設置から日が浅い設備は経過の範囲に収める
                else:
                    b = at - timedelta(days=v0 / rate)
                self.base[(eid, key)] = b
            if b > inst:
                start, origin = b, "carry"
        return max(0.0, rate * (at - start).total_seconds() / 86400), origin

    def record(self, eq: Equipment, s: dict, P: dict, parts: list, done_at: datetime | None) -> None:
        """トラブル対応で部品を交換した（または清掃を完了した）ことを記録する。"""
        if done_at is None:
            return
        k = kind_of(eq)
        names = [p.name for p, _ in parts]
        for c in _COUNTERS.values():
            if k not in c.kinds:
                continue
            if not (any(w in n for w in c.parts for n in names) or s["id"] in c.scen):
                continue
            sub = str(P.get(c.scope, "*")) if c.scope else ""
            bisect.insort(self.resets[(eq.equipment_id, c.key, sub)], done_at)


def _fmt_span(days: float) -> str:
    """経過日数を「約」に続けて書く期間にする（10日 / 3週間 / 5ヶ月 / 2年 / 1年半）。"""
    if days < 13:
        return f"{max(1, round(days))}日"
    if days < 56:
        return f"{round(days / 7)}週間"
    if days < 365:
        return f"{max(2, round(days / 30.4))}ヶ月"
    y = days / 365.25
    return f"{int(y)}年半" if y - int(y) >= 0.5 else f"{int(y)}年"


def _counter_params(eq: Equipment, P: dict, vals: dict) -> dict:
    """カウンタ値から文に入るパラメータを作る。値と矛盾しないよう、判定・なぜなぜの文も値に合わせて選ぶ。"""
    v = {k: x for k, (x, _) in vals.items()}
    origin = {k: o for k, (_, o) in vals.items()}
    out: dict = {}
    if "pou" in v:
        fd = max(1, int(v["pou"]))
        out.update(fd=fd, fd_note=("月次交換の予定日を過ぎていた。" if fd > 31 else "月次交換の直前だった。" if fd >= 22
                                   else "前回交換から日が浅く、早期の目詰まり。"))
    if "pad" in v:
        life = _pad_life(eq.equipment_id)
        used = max(20, int(v["pad"]))
        ratio = used / life
        gd0 = float(P.get("gd0") or 0.75)
        worn = ratio >= 0.85
        out.update(life=life, used=used, life2=max(600, int(life * 0.8 / 50) * 50),
                   gd=f"{max(0.10, gd0 * (1 - 0.8 * min(ratio, 1.1))):.2f}",
                   pad_life_why="パッド寿命設定が長すぎた" if worn else f"寿命設定（{life}枚）より早く溝が摩耗した",
                   pad_why="寿命設定を新スラリー導入時に見直していなかった" if worn else "スラリー変更後の溝深さ測定をしていなかった")
    if "dresser" in v:
        out["disk"] = max(10, int(v["dresser"]))
    if "membrane" in v:
        mem = max(50, int(v["membrane"]))
        over = mem >= 5000
        out.update(mem=f"{mem:,}",
                   mem_why1="推奨交換枚数（5,000枚）を超えて使用していた" if over else "推奨枚数に達する前に境界部から亀裂が進行した",
                   mem_why2="PM周期（2ヶ月）が処理枚数の増加に合っていなかった" if over else "境界部の外観点検が点検項目に無かった")
    if "retainer" in v:
        rl = max(50, int(v["retainer"]))
        out.update(rl=f"{rl:,}", rt=f"{max(0.15, 2.0 - 1.85 * rl / 8000):.2f}")     # 新品2.0mm から処理枚数に比例して減る
    if "cvd_rf" in v:
        out["hrs"] = f"{int(v['cvd_rf']):,}"
    if "cvd_clean" in v:
        um = max(30, int(v["cvd_clean"]))
        out.update(rfh=f"{um:,}", cln=f"{max(300, int(um * 0.7 / 100) * 100):,}")
    if "cvd_oring" in v:
        out["oring_age"] = _fmt_span(v["cvd_oring"])
    if "dp" in v:
        out["hrs"] = f"{int(v['dp']):,}"
    if "mfc" in v:
        age = _fmt_span(v["mfc"] * 365.25)
        note = f"導入から約{age}、一度も交換していない" if origin["mfc"] == "install" else f"前回交換から約{age}経過"
        out.update(mfc_age=age, mfc_age_note=note)
    if "esc" in v:
        h = int(v["esc"])
        over = h >= 6000
        out.update(rfh=f"{h:,}", esc_pm=f"{max(2000, min(5500, int(h * 0.8 / 500) * 500)):,}",
                   esc_why1="RF積算が推奨の6,000hを超えていた" if over else "推奨6,000hより早くシール帯の侵食が進んだ",
                   esc_why2="RF積算時間による交換計画が無かった" if over else "高パワーレシピの比率増加を寿命管理に反映していなかった")
    if "etc_match" in v:
        out["hrs"] = f"{int(v['etc_match']):,}"
    if "electrode" in v:
        h = int(v["electrode"])
        over = h >= 3000
        out.update(hrs=f"{h:,}",
                   arc_why1="上部電極のRF積算が交換基準3,000hを超えていた" if over else "交換基準3,000hに達する前に電極表面が荒れた",
                   arc_why2="PM周期（6ヶ月）がRF時間の増加に合っていなかった" if over else "高パワーレシピの増加を交換基準に反映していなかった")
    if "focus_ring" in v:
        h = int(v["focus_ring"])
        out.update(hrs=f"{h:,}", mt=int(6 + 22 * min(1.0, h / 1000)))                 # RF時間に応じた厚み減少（%）
    if "tmp" in v:
        out["hrs"] = f"{int(v['tmp']):,}"
    if "laser" in v:
        shot = int(v["laser"])
        shl = _laser_life(eq.equipment_id)
        near = shot >= shl * 100 * 0.9
        out.update(shot=f"{shot:,}", shl=shl,
                   laser_why1="チャンバのショット数が寿命目安に達していた" if near else "寿命目安より早く放電電極の摩耗が進んだ",
                   laser_why2="寿命到達前の交換計画を立てていなかった" if near else "HV電圧の上昇傾向を監視していなかった")
    if "alg_lamp" in v:
        h = int(v["alg_lamp"])
        out.update(lamp=f"{h:,}", lamp_pm=f"{max(500, int(h * 0.8 / 100) * 100):,}")
    if "ins_lamp" in v:
        out["lamp"] = f"{int(v['ins_lamp']):,}"
    if "claw" in v:
        out["claw_age"] = _fmt_span(v["claw"])
    if "filament" in v:
        out["fil"] = f"{max(12, int(v['filament'])):,}"
    if "hoist_km" in v:
        km = max(100, int(v["hoist_km"]))
        veh = P.get("veh", "")
        note = (f"{veh}号車の走行距離は{km:,}km、交換目安（4万km）を超えていた。" if km >= 40000
                else f"{veh}号車の走行距離は{km:,}km（交換目安4万kmの手前）。")
        out.update(km=f"{km:,}", km_note=note, km_pm=f"{max(10000, min(30000, int(km * 0.8 / 5000) * 5000)):,}")
    if "agv_battery" in v:
        cyc = max(30, int(v["agv_battery"]))
        out.update(cyc=f"{cyc:,}", bat_cause=(f"バッテリーの経年劣化（{cyc:,}サイクル）による容量低下" if cyc >= 1500
                                             else f"バッテリーセルの早期劣化（{cyc:,}サイクル）による容量低下"))
    if "resin" in v:
        mo = max(1, round(v["resin"]))
        out.update(resin=mo, resin_why="交換目安（12ヶ月）を超えて使用していた" if mo >= 12
                   else "原水負荷が高く、交換目安より早く交換容量に達した")
    if "uv" in v:
        h = int(v["uv"])
        out.update(uv=f"{h:,}", uv_why="交換目安（8,000h）を超えて点灯していた" if h >= 8000 else "一部ランプが寿命前に劣化した（個体差）")
    if "upw_pump" in v:
        out["hr"] = f"{int(v['upw_pump']):,}"
    return out


def _consistent_cause(s: dict, ci: int, P: dict) -> int:
    """使用量と矛盾する原因候補（使用枚数が少ないのに「寿命設定が長すぎた」など）を、同じシナリオの別候補に替える。"""
    sid = s["id"]
    if sid == "CMP.pad.scratch" and ci == 2 and P.get("used", 0) < 0.85 * P.get("life", 1):
        return 0
    if sid == "CMP.head.retainer" and ci == 0 and float(P.get("rt") or 0) > 1.0:
        return 1
    return ci


def _affected_tools(eq: Equipment, at: datetime, ar: random.Random) -> dict:
    """ユーティリティ停止で影響を受けた装置台数。その時点で設置済みの、同じFab（全Fab共通なら全体）の実台数を上限にする。"""
    d = at.date()
    shared = eq.line == "全Fab共通"
    inst = [e for e in _equipment_tuple() if e.installed_date <= d and (shared or fab_of(e) == fab_of(eq))]

    def pick(kinds: tuple) -> int:
        n = sum(1 for e in inst if kind_of(e) in kinds)
        return ar.randint(min(2, max(1, n)), max(1, n))

    exh = ("CLN", "ETC", "CVD") if eq.equipment_id == "EXH-911" else PROCESS_KINDS     # 酸排気は薬液・プラズマ系のみ
    return dict(ntool_cln=pick(("CLN",)), ntool_cvd=pick(("CVD",)), ntool_exh=pick(exh),
                ntool_pcw=pick(("CMP", "CVD", "ETC", "LIT", "IMP")), ntool_vac=pick(("CMP", "CLN", "INS", "LIT", "ROB")))


def _power_dip_count(dr: _Draft) -> int:
    """瞬低で工場内（同じFab）で止まった台数。同時に記録された件数以上・設置台数以下で、グループ内は同じ値。"""
    root = dr.master or dr
    if not root.ntool_fab:
        ar = rng(f"power-dip:{root.eq.equipment_id}:{root.at:%Y%m%d%H%M}")
        d = root.at.date()
        n_fab = sum(1 for e in _equipment_tuple() if e.installed_date <= d and (fab_of(e) == fab_of(root.eq) or e.line == "全Fab共通"))
        root.ntool_fab = ar.randint(root.group_n, max(root.group_n, n_fab))
    return root.ntool_fab


def _apply_context(dr: _Draft, s: dict, eq: Equipment, at: datetime, P: dict, book: _CounterBook, iid: str) -> None:
    """設備マスタの実台数・車両数・使用量カウンタなど、トラブル単体の乱数では決められない値で P を上書きする（補助乱数のみ使う）。"""
    k = kind_of(eq)
    ar = rng(f"incident-context:{iid}")
    if k == "OHT":
        m = re.search(r"車両(\d+)台", eq.model)
        P["veh"] = ar.randint(1, int(m.group(1)) if m else 40)
    elif k == "AGV":
        P["veh"] = ar.randint(1, _AGV_FLEET.get(eq.equipment_id, 6))
    if k in UTILITY_KINDS:
        P.update(_affected_tools(eq, at, ar))
    if s["id"] == "COM.power_dip":
        P["ntool_fab"] = _power_dip_count(dr)
    if k == "LIT":
        P["shl"] = _laser_life(eq.equipment_id)
    keys = _SCEN_COUNTERS.get(s["id"], ())
    if keys:
        vals = {}
        for key in keys:
            c = _COUNTERS[key]
            sub = str(P.get(c.scope, "")) if c.scope else ""
            vals[key] = book.value(eq, key, sub, at)
        P.update(_counter_params(eq, P, vals))


# ---------------------------------------------------------------------------
# 明記された待ち時間と停止時間の整合
# ---------------------------------------------------------------------------
_WAIT_WORDS = ("保持", "待ち", "慣らし", "エージング", "フラッシング", "作業約", "放置")
_DUR_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(時間|分)")


def _wait_minutes(steps: list[str]) -> int:
    """処置の文に明記された待ち・保持の時間（加圧試験○時間保持、慣らし運転○分、温度安定待ち○時間など）の合計（分）。"""
    total = 0.0
    for st in steps:
        if not any(w in st for w in _WAIT_WORDS):
            continue
        durs = [float(x) * (60 if u == "時間" else 1) for x, u in _DUR_RE.findall(st)]
        if durs:
            total += max(durs)
        elif "冷却待ち" in st:
            total += 90
        elif "エージング" in st:
            total += 30
    return int(total)


def _min_downtime(steps: list[str], first: str, inv: list[str], rep_delay: int, wait: int, extra: int) -> int:
    """文に明記された時間から見て最低限必要な停止時間（分）。明記が無ければ 0（通常の停止時間の分布を変えない）。

    明記された待ち（保持・慣らし運転・エージング・FE来場）があれば 連絡・着手までの時間＋待ち時間＋作業時間（手順数×12分＋extra）、
    初動に「保全○○が△分後に到着」とだけあれば 連絡までの時間＋到着までの時間＋extra。
    """
    arrive = re.search(r"が(\d+)分後に到着", first)
    arrive_min = int(arrive.group(1)) if arrive else 0
    waits = _wait_minutes(steps)
    if any("FE" in st for st in steps):
        # FE作業を含む処置は FE 来場を待っている
        for t in inv:
            m = re.search(r"(\d+)時間後に来場", t)
            if m:
                waits += int(m.group(1)) * 60
    if waits:
        work = 12 * len([x for x in steps if x != "（継続対応中）"]) + extra
    elif arrive_min:
        work = extra
    else:
        return 0
    return int(math.ceil((rep_delay + max(wait, arrive_min) + waits + work) / 5) * 5)


def _draw_drafts() -> list[_Draft]:
    r = rng("incident-schedule")
    eqs = _equipment_tuple()
    days = [PERIOD_START + timedelta(days=i) for i in range((PERIOD_END - PERIOD_START).days + 1)]
    keys, cum, total = [], [], 0.0
    for d in days:
        for eq in eqs:
            w, bad_active = _day_eq_weight(eq, d)
            if w <= 0:
                continue
            total += w
            keys.append((d, eq, bad_active))
            cum.append(total)
    drafts: list[_Draft] = []
    process_like = [e for e in eqs if kind_of(e) in PROCESS_KINDS + TRANSPORT_KINDS]
    for _ in range(N_INCIDENTS):
        d, eq, bad_active = keys[bisect.bisect_left(cum, r.random() * total)]
        ws, _sum = _scen_w_cached(eq.equipment_id, d.month, int(_age_years(eq, d)), bad_active)
        s = r.choices(_scenarios_for(eq), weights=ws)[0]
        hour = r.choices(range(24), weights=_HOUR_W)[0]
        at = datetime(d.year, d.month, d.day, hour, r.randint(0, 59))
        dr = _Draft(eq, s, at)
        drafts.append(dr)
        if s["id"] == "COM.power_dip":
            # 瞬低は工場内の複数設備で同時に停止する
            others = [e for e in process_like if e is not eq and e.installed_date <= d and fab_of(e) == fab_of(eq)]
            for e2 in r.sample(others, k=min(len(others), r.randint(3, 9))):
                drafts.append(_Draft(e2, s, at + timedelta(minutes=r.randint(0, 3)), master=dr))
    drafts.sort(key=lambda x: (x.at, x.eq.equipment_id))
    # 再発の紐付け（同一設備・同一シナリオが120日以内）
    last: dict = {}
    for dr in drafts:
        key = (dr.eq.equipment_id, dr.s["id"])
        p = last.get(key)
        gap = (dr.at - p.at).total_seconds() / 86400 if p is not None else 999
        # 期間内でも、記入者が「再発」と認識して紐付けるのは近いものほど多い
        if gap <= RECURRENCE_DAYS and dr.master is None and r.random() < 0.85 * (1 - gap / (RECURRENCE_DAYS + 10)):
            dr.prev = p
            dr.chronic_n = p.chronic_n + 1
        last[key] = dr
    return drafts


def _pick_people(r, eq, sev, shift, detected, s, fe_on):
    ppl = _people_tuple()
    k = kind_of(eq)
    if k in UTILITY_KINDS:
        sect = "設備保全課 施設係"
    else:
        sect = "設備保全課 保全1係" if fab_of(eq) == "Fab1" else "設備保全課 保全2係"
    mnt = [p for p in ppl if p.section == sect]
    staff = [p for p in mnt if p.role != "係長"]
    chief = [p for p in mnt if p.role == "係長"] or [p for p in ppl if p.role == "課長"]
    n = {"小": r.choice([1, 1, 2]), "中": r.choice([1, 2, 2]), "大": r.choice([2, 2, 3]), "重大": r.choice([3, 4])}[sev]
    if shift != "日勤":
        n = max(1, n - 1)
    assignees = r.sample(staff, k=min(n, len(staff)))
    if sev == "重大" or (sev == "大" and r.random() < 0.25):
        assignees.append(chief[0])
    if fe_on:
        fe = [p for p in ppl if p.role == "FE" and p.section == eq.maker]
        if fe:
            assignees.append(fe[0])
    if s["cat"] == "品質" and r.random() < 0.5:
        assignees.append(r.choice([p for p in ppl if p.section == "生産技術課"]))
    return assignees


def _reporter(r, eq, at, detected, assignees) -> Person:
    ppl = _people_tuple()
    if detected in ("SPC異常", "FDC監視") or detected.startswith(("SPC", "QC", "インライン", "技術")):
        return r.choice([p for p in ppl if p.section == "生産技術課"])
    if detected in ("後工程からの連絡", "マクロ検査"):
        return r.choice([p for p in ppl if p.section in ("品質保証課", "生産技術課")])
    if detected in ("保全巡回", "中央監視") or kind_of(eq) in UTILITY_KINDS:
        return assignees[0]
    crew = crew_of(at)
    return r.choice([p for p in ppl if p.section == f"製造課 {crew}班"])


def _detected_by(r, s, eq, alarm) -> str:
    k = kind_of(eq)
    if s["det"] and (alarm is None or r.random() < 0.6):
        return r.choice(s["det"])
    if k in UTILITY_KINDS:
        return r.choices(["中央監視", "保全巡回", "使用先装置からの連絡"], weights=[70, 18, 12])[0]
    if alarm is not None:
        return r.choices(["装置アラーム", "オペレーター", "FDC監視", "保全巡回"], weights=[72, 14, 9, 5])[0]
    return r.choices(["オペレーター", "保全巡回", "FDC監視", "SPC異常"], weights=[45, 25, 15, 15])[0]


_INLINE_OPT = re.compile(r"（\?([^）]*)）")
_TOKEN_RE = re.compile(r"[ァ-ヴー]{3,}|[A-Za-z0-9]{2,}|[一-龥]+")


def _tokens(s: str) -> set:
    """原因との関連度判定用の語（カタカナ・トライグラム、英数語、漢字バイグラム）。テンプレートの {key} は除く。"""
    s = re.sub(r"\{[^}]*\}", " ", s)
    out = set()
    for m in _TOKEN_RE.findall(s):
        if m[0] >= "一" and len(m) >= 2:
            out.update(m[i:i + 2] for i in range(len(m) - 1))          # 漢字はバイグラム
        elif "ァ" <= m[0] <= "ヴ" or m[0] == "ー":
            out.update(m[i:i + 3] for i in range(len(m) - 2))          # カタカナ複合語はトライグラム
        else:
            out.add(m)
    return out


class _Affinity:
    """シナリオ内に複数ある原因候補のうち、選ばれた原因に固有の語で文章片の相性を測る。

    score > 0: 選ばれた原因に固有の語を含む / score < 0: 他の原因候補に固有の語だけを含む / 0: 中立。
    """

    def __init__(self, causes: list[str], ci: int):
        mine = _tokens(causes[ci])
        others = set().union(*[_tokens(c) for i, c in enumerate(causes) if i != ci]) if len(causes) > 1 else set()
        self.spec = mine - others
        self.other = others - mine

    def score(self, text: str) -> int:
        t = _tokens(text)
        return len(t & self.spec) - len(t & self.other)


def _pick_parts(r, s, eq, act_in: list[str], act_out: list[str], sev: str, gas: str | None, aff: _Affinity | None = None) -> list:
    pb = _part_by_name()
    out = []
    for name, lo, hi in s["parts"]:
        opt = name.startswith("?")
        name = name.lstrip("?")
        part = pb[name]
        if eq.category not in part.applicable_categories:
            continue
        if aff is not None and opt:
            sc = aff.score(name)
            if sc < 0:
                continue      # 別原因向けの任意部品
            if sc > 0:
                opt = False
        if name.startswith("MFC (") and gas:
            cands = [p for p in pb.values() if p.name.startswith(f"MFC ({gas} ") and eq.category in p.applicable_categories]
            part = cands[0] if cands else pb["MFC (汎用 1slm)"]
        if opt:
            # 部品名の本体と括弧内の略称（ESC, TMP など）で、採用した処置文と対応付ける
            core = re.split(r"[ (（]", name)[0]
            aliases = [core] + ([core[2:]] if len(core) >= 6 else []) + [a for a in re.findall(r"[（(]([^）)]+)[）)]", name) if len(a) >= 3]
            hit_in = [a for a in act_in if any(al in a for al in aliases)]
            if hit_in:
                take = any(w in a for a in hit_in for w in ("交換", "新品", "補充", "更新", "貼替"))
            elif any(al in a for al in aliases for a in act_out):
                take = False
            else:
                take = r.random() < 0.4
            if part.unit_price_yen >= 1_000_000 and sev in ("小", "中") and r.random() < 0.8:
                take = False
            if not take:
                continue
        out.append((part, r.randint(lo, hi)))
    return out


_MID_OPT = re.compile(r"・\?([^、。（）・]+)")        # 「清掃・?交換」のような括弧なしの任意部分
_OPT_MARK = re.compile(r"(^|[・、（(\s])\?")


def _why_tail(cc: str, chain: list, ar: random.Random) -> str:
    """なぜなぜ分析の最後に置く仕組み・管理面の要因。原因分類の候補＋共通候補から、連鎖に無いものを選ぶ。"""
    pool = [x for x in _WHY_TAIL.get(cc, []) + _WHY_TAIL_COMMON if x not in chain]
    return ar.choice(pool)


def _build_incident(r: random.Random, dr: _Draft, iid: str, book: _CounterBook) -> Incident:
    eq, s, at = dr.eq, dr.s, dr.at
    k = kind_of(eq)
    ppl = _people_tuple()
    ar = rng(f"incident-aux:{iid}")      # 追加の選択（整合を取るための選び直し等）はこの補助乱数から取る
    shift = shift_of(at)
    sev = _sample_severity(r, s, eq)
    if dr.master is not None and dr.master.inc is not None and r.random() < 0.6:
        sev = r.choice(["中", "小", sev])

    # --- アラーム・検知 ---
    # 現象文にアラームコードが書かれていればそれに合わせる（現象とアラーム欄の食い違いを防ぐ）
    sym_tpl = r.choice([t for t in s["sym"] if _season_ok(t, at.month)] or s["sym"])
    alarm = None
    in_text = [c for c in s["alm"] if c in sym_tpl]
    if in_text:
        alarm = _alarm_by_code()[in_text[0]]
    elif s["alm"] and r.random() < 0.85 and not any(w in sym_tpl for w in ("SPC", "QC", "検査", "測定")):
        alarm = _alarm_by_code()[r.choice(s["alm"])]
    detected = _detected_by(r, s, eq, alarm)

    # --- パラメータ ---
    sibs = [e.equipment_id for e in _equipment_tuple() if kind_of(e) == k and e is not eq]
    has_sib = bool(sibs)
    if not sibs:
        sibs = ["他号機"]
    ops = [p for p in ppl if p.section.startswith("製造課")]
    fe_prob = s["fe"] if s["fe"] is not None else 0.12
    fe_on = r.random() < fe_prob * {"小": 0.3, "中": 0.7, "大": 1.3, "重大": 1.8}[sev]
    q_prob, q_max = s["q"]
    lots = []
    if r.random() < min(0.95, q_prob * _SEV_Q[sev]):
        lots = [_lot_id(r, at.date()) for _ in range(r.choice([1, 1, 1, 2, 2, 3] + ([4, 6] if sev == "重大" else [])))]
    P = _Params(
        eq=eq.equipment_id, name=eq.name, lot=lots[0] if lots else _lot_id(r, at.date()), cnt=r.randint(2, 9), n=r.randint(1, 5),
        days=r.randint(2, 30), pm_days=r.randint(3, 85), slot=r.randint(1, 25), h=r.randint(2, 12), sib=r.choice(sibs),
        fe=MAKER_SHORT.get(eq.maker, eq.maker), op=surname(r.choice(ops)), time=f"{at.hour}:{at.minute:02d}",
        ch=r.choice(["CH-A", "CH-B", "CH-C"] if eq.maker != "瑞穂エンジニアリング" else ["PM1", "PM2", "PM3", "PM4"]),
        rmin=r.randint(5, 40), m=surname(r.choice([p for p in ppl if p.section.startswith("設備保全課") and p.role in ("主任", "担当")])),
        qa=surname(r.choice([p for p in ppl if p.section == "品質保証課"])), year=at.year, month=(at.month + r.randint(1, 3) - 1) % 12 + 1,
        sibs="、".join(r.sample(sibs, k=min(len(sibs), r.randint(1, 3)))), nday=r.randint(10, 28),
    )
    if s["p"]:
        P.update(s["p"](r))
    _apply_context(dr, s, eq, at, P, book, iid)
    assignees = _pick_people(r, eq, sev, shift, detected, s, fe_on)
    writer = assignees[0]
    hb = _writer_habit(writer.employee_id)
    reporter = _reporter(r, eq, at, detected, assignees)

    # --- 状態 ---
    # 未完了（対応中）は直近1〜2ヶ月に集中する。半年以上前の案件はほぼ完了済みで、残るのは長期の保留・経過観察だけ
    days_left = (datetime(PERIOD_END.year, PERIOD_END.month, PERIOD_END.day, 23, 59) - at).days
    if days_left < 20:
        w_status = [30, 40, 15, 15]
    elif days_left < 60:
        w_status = [62, 18, 10, 10]
    elif days_left < 180:
        w_status = [86, 3, 5, 6]
    else:
        w_status = [93.8, 0.1, 1.3, 4.8]
    status = r.choices(["完了", "対応中", "保留", "経過観察"], weights=w_status)[0]

    # --- 現象 ---
    sym = _fill(sym_tpl, P)
    if r.random() < 0.38:
        pool = [x for x in _SYM_PREFIX["any"] + _SYM_PREFIX[shift]
                if "連休明け" not in x or (_is_holiday(at.date() - timedelta(days=1)) and not _is_holiday(at.date()))]
        pre = _fill(r.choice(pool), P)
        if not (P["lot"] in sym and "{lot}" in pre) and not sym.startswith(("{", pre[:3])):
            sym = pre + sym
    if r.random() < 0.3:
        suf = r.choice(_SYM_SUFFIX if alarm is not None else [x for x in _SYM_SUFFIX if "リセット" not in x and "ワーニング" not in x])
        sym = sym.rstrip("。") + _fill(suf, P)
    if dr.prev is not None and dr.prev.inc is not None and r.random() < 0.55:
        pv = dr.prev.inc
        sym = sym.rstrip("。") + f"。{fmt_date(pv.occurred_at, r)}にも同様の不具合あり（{pv.incident_id}）"
    if dr.master is not None and dr.master.inc is not None:
        sym = sym.rstrip("。") + f"。同時刻に工場内複数設備が停止（{dr.master.inc.incident_id} 参照）"

    # --- 初動 ---
    if k in UTILITY_KINDS:
        fpool = _FIRST["util"]
    elif s["cat"] == "外部要因":
        fpool = [_FIRST["alarm"][0], _FIRST["alarm"][4], _FIRST["alarm"][5], _FIRST["alarm"][7], _FIRST["patrol"][3]]
    elif detected in ("装置アラーム", "FDC監視"):
        fpool = _FIRST["alarm"]
    elif lots or s["cat"] == "品質":
        fpool = _FIRST["quality"]
    else:
        fpool = _FIRST["patrol"] + _FIRST["alarm"][:3]
    fr = r.sample(fpool, k=r.choice([1, 2, 2, 3]))
    if s["fr"]:
        fr.insert(0, r.choice(s["fr"]))
    if r.random() < 0.5:
        fr.append(r.choice(_FIRST["tail"]))
    first = hb["arrow"].join(_fill(x, P) for x in fr)

    # --- 原因（先に決め、調査・処置・部品・なぜなぜを原因に合わせて選ぶ） ---
    ci = _consistent_cause(s, r.randrange(len(s["cause"])), P)
    aff = _Affinity(s["cause"], ci)
    cause = _fill(s["cause"][ci], P)
    cause_base = cause
    cause_cat = s["cc"]
    age = int(_age_years(eq, at.date())) + 1
    plain = s["cat"] in ("外部要因", "人為")
    if r.random() < 0.3 and not plain:
        cause = cause + r.choice(["と推定", "と判断", "（メーカー見解も同様）", "。ただし複合要因の可能性あり", "と思われる"])
    if r.random() < 0.3 and not plain:
        cause = cause.rstrip("。") + _fill(r.choice(["（前回交換から約{cnt}ヶ月）", f"（設置{age}年目）", "（前回PMから{pm_days}日）",
                                                    "（直近{days}日で兆候あり）", f"（{eq.model}で既知の弱点）"]), P)
    if dr.chronic_n >= 2 and eq.equipment_id in _BAD_UNITS and r.random() < 0.25:
        cause += "。同一箇所での再発が続いており、構造上の問題（設計起因）としてメーカーへ改善申し入れ"
        cause_cat = "設計起因"
    if s["why"]:
        # 連鎖の文には使用量に応じた判定（{mem_why1} など）が入るので、埋めた文で原因との相性を測る
        chains = [[_fill(x, P) for x in w] for w in s["why"]]
        scored = [(aff.score(" ".join(w)), i) for i, w in enumerate(chains)]
        best = max(sc for sc, _ in scored)
        why = list(chains[r.choice([i for sc, i in scored if sc == best])])
        if not plain and (best <= -3 or (best < 0 and len(s["why"]) > 1)):
            # 用意された連鎖が別原因向けなら、選んだ原因から組み立てる
            first_why = why[0] if aff.score(why[0]) >= 0 else f"{s['sub']}の異常で停止した"
            why = [first_why, re.sub(r"（[^）]*）$", "", cause_base), ar.choice(_WHY_MID.get(cause_cat, _WHY_MID_DEFAULT))]
            why.append(_why_tail(cause_cat, why, ar))
    else:
        why = []
    if why and r.random() < 0.3 and not plain:
        # 最後の「なぜ」を、記入者が仕組みの問題として書き直すことがある（原因分類ごとの候補から）
        why[-1] = _why_tail(cause_cat, why, ar) if ar.random() < 0.8 else why[-1]
    while why and len(why) < 3:
        why.append(_why_tail(cause_cat, why, ar))
    why = why[:5]

    # --- 調査 ---
    inv_src = [(t, aff.score(t)) for t in s["inv"] if _season_ok(t, at.month)] or [(t, 0) for t in s["inv"][:2]]
    kmin, kmax = (2, 3) if hb["terse"] else (2, 4)
    must = [i for i, (_, sc) in enumerate(inv_src) if sc > 0]
    neutral = [i for i, (_, sc) in enumerate(inv_src) if sc == 0]
    kk = r.randint(kmin, kmax)
    idx = must[:kk]
    if len(idx) < kk and neutral:
        idx += r.sample(neutral, k=min(len(neutral), kk - len(idx)))
    if not idx:
        idx = [r.randrange(len(inv_src))]
    inv = [_fill(inv_src[i][0], P) for i in sorted(idx)]
    n_gen = max(0, r.choice([0, 1, 1, 2]) - (1 if hb["terse"] else 0))
    gen_pool = _GENERIC_INV if s["cat"] != "外部要因" else _GENERIC_INV[5:]   # 瞬低・地震に「兆候」等は書かない
    for g in r.sample(gen_pool, k=n_gen):
        inv.insert(r.randint(0, len(inv)), _fill(g, P))
    if fe_on and r.random() < 0.7:
        inv.append(r.choice([f"{P['fe']}FEへ連絡し、{r.randint(2, 20)}時間後に来場。", f"{P['fe']}のサービス窓口へ技術問合せ、FE来場で共同調査。",
                             f"{P['fe']}FEと電話で状況共有、部品手配を並行で進めた。"]))
    investigation = "".join(_end(x) for x in inv)

    # --- 処置 ---
    opt_in, opt_out, steps = [], [], []
    n_req = sum(1 for a in s["act"] if not a.startswith("?"))
    for a in s["act"]:
        body, opt = a.lstrip("?"), a.startswith("?")
        m_inline = _INLINE_OPT.search(body)
        if m_inline:
            # 「清掃（?交換）」のような文中の任意部分: 採用なら括弧を残し部品と対応付ける
            if r.random() < 0.45:
                body = body[:m_inline.start()] + f"（{m_inline.group(1)}）" + body[m_inline.end():]
                opt_in.append(body)
            else:
                opt_out.append(body[:m_inline.start()] + body[m_inline.end():] + m_inline.group(1))
                body = body[:m_inline.start()] + body[m_inline.end():]
        m_mid = _MID_OPT.search(body)
        if m_mid:
            # 「清掃・?交換」のような括弧なしの任意部分: 採用なら「清掃・交換」、不採用なら「清掃」。
            # 部品は選んだ原因と関係が深いと必ず交換される（_pick_parts）ので、処置の文もそれに合わせる
            if aff.score(body) > 0 or (aff.score(body) == 0 and ar.random() < 0.45):
                body = body[:m_mid.start()] + f"・{m_mid.group(1)}" + body[m_mid.end():]
                opt_in.append(body)
            else:
                opt_out.append(body[:m_mid.start()] + m_mid.group(1) + body[m_mid.end():])
                body = body[:m_mid.start()] + body[m_mid.end():]
        body = _OPT_MARK.sub(r"\1", body)     # 任意ステップの記号「?」は文に残さない
        sc = aff.score(body)
        if sc < 0 and opt:
            opt_out.append(body)          # 別原因向けの処置は採らない
            continue
        if opt and not (sc > 0 or r.random() < 0.5):
            opt_out.append(body)
            continue
        if opt:
            opt_in.append(body)
        steps.append(_fill(body, P))
    if r.random() < 0.25 and k not in UTILITY_KINDS:
        steps.insert(0, r.choice(["装置をMAINTモードへ切替、LOTO実施", "処理中ウェーハを退避", "安全確認（保護具着用）の上で作業開始"]))
    if _QC.get(k) and r.random() < 0.8:
        qc_text = _fill(r.choice(_QC[k]), _Params({**P, **_qc_params(k, r)}))
        steps.append(qc_text)
    else:
        qc_text = "動作確認OK"
    if r.random() < 0.35:
        steps.append(_fill(r.choice(["生産へ引渡し（{op}班長確認）", "製造課へ復旧連絡", "DOWN解除・MES更新", "作業後の工具・部品員数確認"]), P))
    parts = _pick_parts(r, s, eq, [_fill(x, P) for x in opt_in], [_fill(x, P) for x in opt_out], sev, P.get("gas"), aff)

    # 高額部品（100万円超）を交換した案件は小規模トラブルにはならない
    if any(p.unit_price_yen >= 1_000_000 for p, _ in parts) and sev in ("小", "中"):
        sev = "大" if any(p.unit_price_yen >= 5_000_000 for p, _ in parts) else r.choice(["中", "大", "大"])

    # --- 時間 ---
    med, sg, lo, hi = _DT_MEDIAN[sev]
    dt_min = min(hi, max(lo, r.lognormvariate(math.log(med), sg))) * s["dt"]
    if shift != "日勤":
        dt_min *= 1.12
    if any(p.unit_price_yen >= 1_000_000 for p, _ in parts):
        dt_min += r.randint(180, 1200)
    if detected in ("SPC異常", "インライン欠陥検査", "後工程からの連絡", "マクロ検査") or detected.startswith(("SPC", "QC", "パーティクル")):
        rep_delay = r.randint(30, 480)
    else:
        rep_delay = r.randint(1, 25) if shift == "日勤" else r.randint(3, 40)
    reported_at = at + timedelta(minutes=rep_delay)
    wait = r.randint(3, 25) if shift == "日勤" else r.randint(8, 70)
    m_arrive = re.search(r"が(\d+)分後に到着", first)
    if m_arrive:
        wait = max(wait, int(m_arrive.group(1)))     # 初動に書いた到着時刻と対応開始を合わせる
    response_at = reported_at + timedelta(minutes=wait)
    # ダウンタイム = 発生→連絡 + 連絡→着手 + 修理時間（dt_min）
    downtime = int(rep_delay + wait + dt_min)
    if r.random() < 0.6:
        downtime = max(rep_delay + wait + 5, round(downtime / 5) * 5)
    # 処置に「加圧試験6時間保持」「慣らし運転40分」などと書いてあれば、停止時間はそれ以上になる
    work_extra = ar.randint(10, 40)
    downtime = max(downtime, _min_downtime(steps, first, inv, rep_delay, wait, work_extra))
    completed_at = at + timedelta(minutes=downtime)
    # 実作業時間: ダウンタイムには待機（部品・FE・昇温待ち）を含むため一部のみ。1人あたり上限を設ける
    per_person = min(downtime / 60 * r.uniform(0.3, 0.7), 9 + downtime / 1440 * 5)
    internal = [p for p in assignees if p.role != "FE"]
    work_hours = max(0.25, round(per_person * max(1, len(internal)) * 4) / 4)

    # --- 結果・対策・展開 ---
    qc_p = _Params({**P, "qc": qc_text.rstrip("。"), "done": fmt_date(completed_at, r) + f" {completed_at.hour}:{completed_at.minute:02d}"})
    if status == "完了":
        result = _fill(r.choice(_RESULT["ok"]), qc_p)
    elif status == "経過観察":
        result = _fill(r.choice(_RESULT["watch"]), qc_p)
        if r.random() < 0.4:
            cause = r.choice(["特定できず。", "原因不明（再現せず）。"]) + f"{cause}の可能性を疑っている"
            cause_cat = "不明"
    else:
        result = ""
    # 再発防止策も原因との相性が最も良いものから選ぶ
    prev_best = max((aff.score(x) for x in s["prev"]), default=0)
    prev_src = [x for x in s["prev"] if aff.score(x) >= min(0, prev_best)]
    prev = r.sample(prev_src, k=min(len(prev_src), r.choice([1, 1, 2, 2, 3]))) if prev_src else []
    if r.random() < 0.25:
        prev.append(r.choice(_PREV_GENERIC))
    prevention = hb["arrow"].join(_fill(x, P).rstrip("。") for x in prev) if hb["arrow"] != "、" else "／".join(_fill(x, P) for x in prev)
    if sev == "小" and r.random() < 0.15:
        prevention = r.choice(["特になし（突発的事象のため監視継続）", "現状の点検で対応可、追加対策なし"])
    horiz_pool = [x for x in _HORIZ if "同型機なし" not in x] if has_sib else [x for x in _HORIZ if "{sibs}" not in x and "同型機（" not in x]
    horiz_tpl = r.choice(horiz_pool)
    if "展開不要" in horiz_tpl and (any(w in prevention for w in _HORIZ_WIDE_WORDS) or dr.master is not None or dr.group_n > 1):
        # 他号機に及ぶ対策・工場内の同時停止（瞬低）なのに「本機固有」とは書かない
        horiz_tpl = ar.choice([x for x in horiz_pool if x and "展開不要" not in x])
    horiz = _fill(horiz_tpl, P)

    # 対応中・保留は記録が途中
    if status == "対応中":
        completed_at = None
        steps = steps[: r.randint(1, max(1, len(steps) // 2))] + ["（継続対応中）"]
        investigation = "".join(_end(x) for x in inv[: r.randint(1, 2)]) + r.choice(["引き続き調査中。", "メーカー回答待ち。", "原因調査中、次シフトへ引継ぎ。"])
        cause, cause_cat, why, prevention, horiz = "", "不明", [], "", ""
        parts = parts[:1] if r.random() < 0.4 else []
        downtime = rep_delay + wait + r.randint(60, 900)
        downtime = max(downtime, _min_downtime(steps, first, inv, rep_delay, wait, work_extra))
    elif status == "保留":
        completed_at = None
        hold_long = days_left >= 180
        result = _fill(r.choice(_RESULT["hold"]), qc_p)
        if hold_long:
            result = _fill(ar.choice(_RESULT["hold_long"]), _Params({**qc_p, "fy": PERIOD_END.year + 1}))
        cause = cause_base + r.choice(["の可能性が高い（確定は部品交換後）", "と推定（要確認）", ""])
        prevention = r.choice(["恒久対策は部品入荷後に検討", "恒久対策は原因確定後に検討", ""])
        hold_step = _fill(r.choice(["暫定処置で運転再開（監視強化）", "部品手配中（{fe}へ発注済み）", "代替条件で運用継続（技術承認済み）",
                                    "次回PMで本処置予定"]), P)
        if hold_long and "部品手配中" in hold_step:
            hold_step = "暫定処置で運転継続（監視強化）"
        steps = steps[: max(1, len(steps) // 3)] + [hold_step]
        downtime = rep_delay + wait + int(dt_min * r.uniform(0.4, 0.9))
        downtime = max(downtime, _min_downtime(steps, first, inv, rep_delay, wait, work_extra))
        parts = []
    if status in ("対応中", "保留"):
        work_hours = max(0.25, round(min(downtime / 60 * r.uniform(0.3, 0.6), 8) * max(1, len(internal)) * 4) / 4)

    # --- 数量・費用 ---
    scrap = 0
    if lots and q_max and r.random() < 0.55:
        scrap = min(25 * len(lots), max(0, int(r.randint(0, q_max) * _SEV_Q[sev])))
    labor = work_hours * 6500
    fe_cost = r.choice([150_000, 180_000, 240_000]) * max(1, downtime // 1440 + 1) if fe_on else 0
    wafer_cost = r.randint(38, 95) * 1000
    cost = int(round((sum(p.unit_price_yen * q for p, q in parts) + labor + fe_cost + scrap * wafer_cost) / 100) * 100)

    remarks = "" if r.random() < 0.45 else r.choice(_REMARKS)
    if remarks:
        # 保守契約・メーカー見積の備考は、FE が作業した／処置にFE・メーカー名が出る案件にだけ書く。部品在庫・転用の備考は部品を使った案件だけ
        act_text = "".join(steps)
        maker_involved = any(p.role == "FE" for p in assignees) or "FE" in act_text or P["fe"] in act_text
        ng = (() if maker_involved else _REMARKS_NEED_MAKER) + (() if parts else _REMARKS_NEED_PARTS)
        if remarks in ng:
            remarks = ar.choice([x for x in _REMARKS if x not in ng])
    remarks = _fill(remarks, P)
    book.record(eq, s, P, parts, completed_at or (at + timedelta(minutes=downtime) if parts else None))
    if dr.prev is not None and dr.prev.inc is not None and not remarks and r.random() < 0.4:
        remarks = f"再発。前回 {dr.prev.inc.incident_id}（{fmt_date(dr.prev.at, r)}）"
    photos = []
    if s["photo"] and r.random() < 0.72:
        photos = r.sample(s["photo"], k=min(len(s["photo"]), r.randint(1, 3)))

    # --- 記入者の癖 ---
    if hb["polite"]:
        investigation, result = _polite(investigation), _polite(result)
    symptom = _noise(sym, hb, r)
    first = _noise(first, hb, r)
    investigation = _noise(investigation, hb, r)
    cause = _noise(cause, hb, r)
    action = _noise(_numbered(steps, hb["numbering"]), hb, r)
    result = _noise(result, hb, r)

    related = None
    if dr.master is not None and dr.master.inc is not None:
        related = dr.master.inc.incident_id
    elif dr.prev is not None and dr.prev.inc is not None:
        related = dr.prev.inc.incident_id

    return Incident(
        incident_id=iid, occurred_at=at, detected_by=detected, reported_at=reported_at, response_started_at=response_at,
        completed_at=completed_at, equipment=eq, shift=shift, reporter=reporter, assignees=assignees, category=s["cat"],
        subsystem=s["sub"], severity=sev, status=status, symptom=symptom, alarm=alarm, first_response=first,
        investigation=investigation, cause=cause, cause_category=cause_cat, why_why=why, action=action, parts_used=parts,
        downtime_min=int(downtime), work_hours=float(work_hours), result=result, prevention=prevention, horizontal_deployment=horiz,
        lots_affected=lots, scrap_wafers=scrap, cost_yen=cost, recurrence=dr.prev is not None, related_incident_id=related,
        remarks=remarks, photos=photos, scenario_id=s["id"],
    )


@lru_cache(maxsize=None)
def _incidents_tuple() -> tuple:
    drafts = _draw_drafts()
    for dr in drafts:
        if dr.master is not None:
            dr.master.group_n += 1          # 瞬低の同時停止グループの件数
    r = rng("incident-text")
    book = _CounterBook()
    seq: dict[int, int] = {}
    out = []
    for dr in drafts:
        y = dr.at.year
        seq[y] = seq.get(y, 0) + 1
        dr.inc = _build_incident(r, dr, f"TR-{y}-{seq[y]:05d}", book)
        out.append(dr.inc)
    return tuple(out)


def standard_incidents() -> list[Incident]:
    """正準のトラブル履歴（2023-04-01〜2026-08-31、約8,000件、発生日時順）。"""
    return list(_incidents_tuple())


# ---------------------------------------------------------------------------
# チョコ停
# ---------------------------------------------------------------------------
_MINOR_RATE = {"CMP": 1.5, "CVD": 0.9, "ETC": 1.0, "LIT": 1.0, "CLN": 1.2, "IMP": 1.1, "INS": 1.2, "OHT": 3.0, "AGV": 1.4,
               "ROB": 1.5, "STK": 0.8, "UPW": 0.25, "EXH": 0.15, "SCR": 0.3, "PCW": 0.2, "VAC": 0.2}
# チョコ停メモとアラームを対応付けるキーワード（サブシステム -> メモ中の語）
_SUBSYS_KW = {
    "ウェーハ搬送": ["搬送", "受け渡し", "ウェーハ有無", "ロードカップ", "ロボ"], "後洗浄ユニット": ["洗浄部", "DIW"], "コンディショナ": ["ドレッサ"],
    "研磨ヘッド": ["ヘッド"], "MFC/ガス供給": ["ガス"], "真空搬送": ["ロボ", "スライド"], "ドライポンプ": ["DP"], "RF電源": ["Vpp"],
    "終点検出": ["EPD"], "TMP/排気": ["He"], "レチクル搬送": ["レチクル", "BCR"], "アライメント": ["アライメント"], "フォーカス": ["AF"],
    "スピンチャック": ["チャックピン"], "ノズル": ["ノズル"], "薬液供給": ["補充ポンプ"], "ビームライン": ["ビーム"], "イオン源": ["ソース"],
    "高圧電源": ["アーク"], "光学系/光源": ["AF", "光量"], "プリアライナ/搬送": ["ノッチ", "プリアライン"], "画像処理PC": ["ホスト", "レシピ"],
    "走行部": ["障害物", "渋滞"], "グリッパ": ["グリッパ"], "ホイスト": ["ホイスト"], "安全センサ": ["スキャナ"], "バッテリー": ["バッテリー"],
    "移載部": ["移載"], "ハンド/吸着": ["吸着"], "ロードポート": ["マッピング", "ドア", "FOUP ID"], "クレーン": ["クレーン"],
    "入出庫ポート": ["入出庫"], "N2パージ": ["N2パージ"], "RO膜": ["RO"], "UF膜": ["微粒子"], "UV酸化": ["TOC"], "排気ファン": ["振動"],
    "ダクト/ダンパ": ["ダクト"], "循環水": ["pH", "スプレー"], "水質管理": ["導電率"], "循環ポンプ": ["供給圧"], "封水/冷却": ["封水"],
    "真空ポンプ": ["真空圧"],
}
_MINOR_SUFFIX = ["", "", "", "", "（{cnt}回目）", "。本日{cnt}回目", "。様子見", "。次回PMで点検予定", "。保全へ情報共有", "。同ロット継続"]


@lru_cache(maxsize=None)
def _minor_tuple() -> tuple:
    r = rng("minor-stops")
    eqs = _equipment_tuple()
    ppl = _people_tuple()
    days = [PERIOD_START + timedelta(days=i) for i in range((PERIOD_END - PERIOD_START).days + 1)]
    keys, cum, total = [], [], 0.0
    for d in days:
        for eq in eqs:
            if d < eq.installed_date:
                continue
            w = _MINOR_RATE[kind_of(eq)] * (1.0 + 0.02 * min(_age_years(eq, d), 20))
            if eq.equipment_id in _BAD_UNITS and d <= _BAD_UNITS[eq.equipment_id][2]:
                w *= 1.4
            total += w
            keys.append((d, eq))
            cum.append(total)
    abc = _alarm_by_code()
    raw = []
    for _ in range(N_MINOR_STOPS):
        d, eq = keys[bisect.bisect_left(cum, r.random() * total)]
        at = datetime(d.year, d.month, d.day, r.choices(range(24), weights=_HOUR_W)[0], r.randint(0, 59), r.randint(0, 59))
        raw.append((at, eq))
    raw.sort(key=lambda x: (x[0], x[1].equipment_id))
    out, seq = [], {}
    for at, eq in raw:
        k = kind_of(eq)
        codes, notes = _MINOR[k]
        note_t = r.choice(notes)
        note = _fill(note_t, _Params(cnt=r.randint(2, 5)))
        cands = [abc[c] for c in codes if any(kw in note for kw in _SUBSYS_KW.get(abc[c].subsystem, [abc[c].subsystem]))]
        alarm = r.choice(cands) if cands and r.random() < 0.9 else (None if r.random() < 0.6 else abc[r.choice(codes)])
        dur = int(min(30, max(1, r.lognormvariate(math.log(4), 0.7))))
        if "自動" in note or "自然" in note:
            recovered = "自動復帰"
            dur = min(dur, 5)
        elif k in UTILITY_KINDS or r.random() < 0.15:
            sect = "設備保全課 施設係" if k in UTILITY_KINDS else ("設備保全課 保全1係" if fab_of(eq) == "Fab1" else "設備保全課 保全2係")
            recovered = r.choice([p for p in ppl if p.section == sect]).name
        else:
            recovered = r.choice([p for p in ppl if p.section == f"製造課 {crew_of(at)}班"]).name
        if r.random() < 0.25:
            note = note.rstrip("。") + _fill(r.choice(_MINOR_SUFFIX), _Params(cnt=r.randint(2, 6)))
        if r.random() < 0.05:
            note = to_zenkaku(note, digits_only=True)
        seq[at.year] = seq.get(at.year, 0) + 1
        out.append(MinorStop(f"MS-{at.year}-{seq[at.year]:05d}", at, eq, alarm, dur, recovered, note))
    return tuple(out)


def minor_stops() -> list[MinorStop]:
    """チョコ停記録（約25,000件、発生日時順）。"""
    return list(_minor_tuple())


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import collections
    import time

    t0 = time.time()
    inc = standard_incidents()
    ms = minor_stops()
    print(f"設備 {len(equipment_master())} / 人 {len(people())} / 部品 {len(parts_catalog())} / アラーム {len(alarm_codes())}")
    print(f"トラブル {len(inc)} 件 / チョコ停 {len(ms)} 件  ({time.time() - t0:.1f}s)")
    for label, f in [("重大度", lambda i: i.severity), ("故障区分", lambda i: i.category), ("状態", lambda i: i.status),
                     ("年", lambda i: i.occurred_at.year)]:
        print(label, dict(collections.Counter(map(f, inc)).most_common()))
    if _Params.missing:
        print("未定義テンプレートキー:", sorted(_Params.missing))
