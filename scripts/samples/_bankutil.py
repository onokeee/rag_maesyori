"""フレーズバンク記述用の小さなヘルパー。

シナリオ（S）は1つの「故障パターン」を表す dict。キーの意味:
    id     : 一意キー（例 "CMP.slurry.clog"）
    sub    : サブシステム名
    cat    : 故障区分（機械/電気/制御/ソフト/ユーティリティ/人為/品質/外部要因）
    cc     : 原因分類（摩耗/劣化/汚れ/設定ミス/締結緩み/断線/ソフト不具合/異物/設計起因/不明 など）
    w      : 発生しやすさ（相対重み）
    sev    : 重大度の重み (重大, 大, 中, 小)
    dt     : ダウンタイム倍率
    alm    : 関連アラームコードの候補（空ならアラームなしで検知）
    det    : 検知経路の候補（省略時は既定）
    p      : r -> dict  固有パラメータ（測定値など）を返す関数
    sym    : 現象の文テンプレート
    fr     : 初動の固有文（省略可）
    inv    : 調査の文テンプレート（2〜4文を順序を保って採用）
    cause  : 原因の文テンプレート
    why    : なぜなぜ分析の連鎖候補（各要素は3〜5段のリスト）
    act    : 処置ステップ（"?" 始まりは任意ステップ）
    parts  : (部品名, 最小数, 最大数)。部品名の先頭 "?" は任意
    prev   : 再発防止策テンプレート
    q      : (ロット影響確率, 廃棄ウェーハ最大数)
    photo  : 写真キャプション候補
    season : {月: 倍率}（省略時は季節変動なし）
    fe     : メーカーFE手配確率（省略時はカテゴリ既定）
    aging  : True なら経年で発生率が上がる（摩耗・劣化系）
テンプレート内の {キー} は domain 側の共通パラメータ（eq, ch, lot, h, cnt など）と p の戻り値で置換する。
"""
from __future__ import annotations


def S(**kw) -> dict:
    """シナリオ定義に既定値を補う。"""
    kw.setdefault("w", 1.0)
    kw.setdefault("sev", (1, 8, 35, 56))
    kw.setdefault("dt", 1.0)
    kw.setdefault("alm", [])
    kw.setdefault("det", None)
    kw.setdefault("p", None)
    kw.setdefault("fr", [])
    kw.setdefault("why", [])
    kw.setdefault("parts", [])
    kw.setdefault("prev", [])
    kw.setdefault("q", (0.0, 0))
    kw.setdefault("photo", [])
    kw.setdefault("season", None)
    kw.setdefault("fe", None)
    kw.setdefault("aging", kw.get("cc") in ("摩耗", "劣化", "断線"))
    return kw


def fx(r, a: float, b: float, nd: int = 1) -> str:
    """a〜b の一様乱数を小数 nd 桁の文字列で返す。"""
    return f"{r.uniform(a, b):.{nd}f}"


def comma(n: int) -> str:
    return f"{n:,}"


SUMMER = {6: 1.2, 7: 1.7, 8: 1.9, 9: 1.4}
WINTER = {12: 1.4, 1: 1.6, 2: 1.5}
RAINY = {6: 1.6, 7: 1.3}
THUNDER = {7: 2.2, 8: 2.8, 9: 1.8}
