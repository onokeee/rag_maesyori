"""動作確認用のサンプルデータを作る。

    python scripts/make_samples.py                     # 小さな帳票の見本（修理報告書3種・点検記録表）を samples/ に作る（テストでも使う）
    python scripts/make_samples.py --large             # 大きなサンプル（一覧表 T1〜T5 / 帳票 F1〜F5）を samples/ に作る
    python scripts/make_samples.py --large --only T1,F3
    python scripts/make_samples.py --large --out D:/tmp/s

小さな見本は、同じ「設備修理報告書」でもレイアウト・項目名が異なる3種類と、別帳票（点検記録表）。
大きなサンプルの各ジェネレータは scripts/samples/<module>.py の generate(output_root) で、固定シードのため何度実行しても
同じ内容になる。最後に生成結果の一覧（ファイル数・行数・サイズ・秒数）を表示し、出力先に README.md（データセットの目録）を書き出す。
"""
from __future__ import annotations

import argparse
import csv
import importlib
import io
import sys
import time
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
from pathlib import Path

from PIL import Image as PILImage, ImageDraw
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill


# ==== 小さな帳票の見本（テストでも使う） ============================================================
LABEL_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")


def _label(ws, coord: str, text: str, merge: str | None = None):
    ws[coord] = text
    ws[coord].font = Font(bold=True)
    ws[coord].fill = LABEL_FILL
    if merge:
        ws.merge_cells(merge)


def _value(ws, coord: str, value, merge: str | None = None):
    ws[coord] = value
    ws[coord].alignment = WRAP
    if merge:
        ws.merge_cells(merge)


def _photo(ws, anchor: str, text: str = "photo"):
    img = PILImage.new("RGB", (240, 160), "#d0d7de")
    draw = ImageDraw.Draw(img)
    draw.rectangle([70, 50, 170, 110], outline="red", width=4)
    draw.text((10, 10), text, fill="black")
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    ws.add_image(Image(buffer), anchor)


def make_repair_standard(path: Path) -> Path:
    """標準レイアウト: 左にラベル・右に値、文章項目は見出しの下に結合セル。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "修理報告書"
    ws["A1"] = "設備修理報告書"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:H1")

    _label(ws, "A3", "報告番号"); _value(ws, "B3", "R2026-00123", "B3:C3")
    _label(ws, "E3", "発生日"); _value(ws, "F3", datetime(2026, 9, 14), "F3:G3")
    ws["F3"].number_format = "yyyy/mm/dd"
    _label(ws, "A4", "設備番号"); _value(ws, "B4", "EQ-001", "B4:C4")
    _label(ws, "E4", "設備名"); _value(ws, "F4", "CMP装置", "F4:G4")
    _label(ws, "A5", "報告者"); _value(ws, "B5", "山田 太郎", "B5:C5")
    _label(ws, "E5", "作業時間"); _value(ws, "F5", 2.5)

    _label(ws, "A7", "故障内容", "A7:H7")
    _value(ws, "A8", "ウェーハ搬送時に搬送アームが停止した。\n装置画面に「Robot Position Error」が表示された。", "A8:H10")
    _label(ws, "A12", "原因", "A12:H12")
    _value(ws, "A13", "搬送アーム位置センサーのコネクタ接触不良。", "A13:H14")
    _label(ws, "A16", "修理内容", "A16:H16")
    _value(ws, "A17", "1. センサーコネクタを取り外し、端子を清掃した。\n2. コネクタを再接続し、搬送動作を確認した。", "A17:H19")
    _label(ws, "A21", "修理結果", "A21:H21")
    _value(ws, "A22", "搬送動作が正常に復旧した。10回の搬送テストで異常なし。", "A22:H23")
    _photo(ws, "J3", "sensor")

    ref = wb.create_sheet("参考資料")
    ref["A1"] = "センサー取扱説明書 抜粋"
    master = wb.create_sheet("マスタ")
    master["A1"], master["B1"] = "設備番号", "設備名"
    for i, (eq, name) in enumerate([("EQ-001", "CMP装置"), ("EQ-002", "CVD装置"), ("EQ-003", "露光装置")], start=2):
        master[f"A{i}"], master[f"B{i}"] = eq, name

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def make_repair_shifted(path: Path) -> Path:
    """レイアウト違い: 位置がズレている・項目名の揺れ・1セル内ラベル・値が文字列の日付。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "修理報告書(2)"
    ws["B2"] = "修理報告書"
    ws["B2"].font = Font(bold=True, size=16)

    _label(ws, "C4", "報告No"); _value(ws, "F4", "R2026-00124")
    _label(ws, "C5", "装置番号"); _value(ws, "F5", "EQ-002")
    ws["C6"] = "装置名：CVD装置"
    _label(ws, "C7", "発生日"); _value(ws, "F7", "2026/09/10")
    _label(ws, "C8", "作業時間"); _value(ws, "F8", "1.5時間")
    _label(ws, "C9", "作成者"); _value(ws, "F9", "佐藤 花子")

    _label(ws, "C11", "症状"); _value(ws, "D11", "成膜レートが低下した。", "D11:J11")
    _label(ws, "C12", "故障原因"); _value(ws, "D12", "ガス供給ラインのMFC不良。", "D12:J12")
    _label(ws, "C13", "処置内容"); _value(ws, "D13", "MFCを交換した。", "D13:J13")
    _label(ws, "C14", "修理結果"); _value(ws, "D14", "成膜レートが規格内に回復した。", "D14:J14")
    _photo(ws, "L4", "mfc-1")
    _photo(ws, "L14", "mfc-2")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def make_repair_table(path: Path) -> Path:
    """表形式: 1行目に項目名が横並び・2行目に値、文章項目は縦結合ラベルの右。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "修理報告書"
    ws["A1"] = "設備修理報告書"
    ws["A1"].font = Font(bold=True, size=16)

    for col, label in zip("ABCDE", ["報告番号", "設備番号", "設備名", "発生日", "報告者"]):
        _label(ws, f"{col}3", label)
    for col, value in zip("ABCDE", ["R2026-00125", "EQ-003", "露光装置", datetime(2026, 8, 30), "鈴木 一郎"]):
        _value(ws, f"{col}4", value)

    _label(ws, "A6", "不具合内容", "A6:A8"); _value(ws, "B6", "ステージ移動時に異音が発生した。", "B6:H8")
    _label(ws, "A9", "原因", "A9:A10"); _value(ws, "B9", "リニアガイドの潤滑不足。", "B9:H10")
    _label(ws, "A11", "対応内容", "A11:A12"); _value(ws, "B11", "グリスアップを実施した。", "B11:H12")
    _label(ws, "A13", "結果", "A13:A14"); _value(ws, "B13", "異音は解消した。", "B13:H14")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def make_inspection(path: Path) -> Path:
    """別の帳票（点検記録表）。テンプレート判定の確認用。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "点検記録"
    ws["A1"] = "定期点検記録表"
    _label(ws, "A3", "点検日"); _value(ws, "B3", datetime(2026, 9, 1))
    _label(ws, "A4", "点検者"); _value(ws, "B4", "高橋")
    _label(ws, "A5", "設備番号"); _value(ws, "B5", "EQ-001")
    _label(ws, "A7", "点検項目"); _label(ws, "B7", "判定"); _label(ws, "C7", "所見")
    for i, (item, result, note) in enumerate([("異音", "OK", ""), ("振動", "OK", ""), ("油漏れ", "NG", "パッキン劣化")], start=8):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"] = item, result, note

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def make_all(directory: Path) -> list[Path]:
    return [
        make_repair_standard(directory / "修理報告書_標準.xlsx"),
        make_repair_shifted(directory / "修理報告書_レイアウト違い.xlsx"),
        make_repair_table(directory / "修理報告書_表形式.xlsx"),
        make_inspection(directory / "点検記録表.xlsx"),
    ]


# ==== 大きなサンプル（一覧表 T1〜T5 / 帳票 F1〜F5） ==================================================
# 「python scripts/make_samples.py --large」で直接実行した場合もパッケージを import できるようにする
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

DEFAULT_OUT = PROJECT_ROOT / "samples"


@dataclass(frozen=True)
class Generator:
    key: str          # --only で指定する記号（T1 など）
    module: str       # scripts.samples 配下のモジュール名
    kind: str         # "table" / "form"
    title: str        # 日本語の名前


GENERATORS: list[Generator] = [
    Generator("T1", "t1_trouble_list", "table", "トラブル対応一覧（Excel・素直な表）"),
    Generator("T2", "t2_system_csv", "table", "設備故障履歴（システム出力CSV・CP932）"),
    Generator("T3", "t3_crosstab", "table", "月別停止時間集計（クロス集計・数式）"),
    Generator("T4", "t4_maintenance_blocks", "table", "保全作業記録・部品交換（1シート複数表）"),
    Generator("T5", "t5_capa_ledger_csv", "table", "是正処置管理台帳（UTF-8 CSV＋Excel）"),
    Generator("F1", "f1_repair_report", "form", "設備修理報告書"),
    Generator("F2", "f2_trouble_report", "form", "トラブル対応報告書（初動〜恒久対策）"),
    Generator("F3", "f3_8d_report", "form", "8D是正処置報告書"),
    Generator("F4", "f4_inspection_report", "form", "設備点検・保全作業報告書"),
    Generator("F5", "f5_process_abnormality", "form", "工程異常連絡票"),
]


@dataclass
class RunResult:
    gen: Generator
    paths: list[Path] = field(default_factory=list)
    seconds: float = 0.0
    rows: str = ""
    size_bytes: int = 0
    error: str | None = None


# ---------------------------------------------------------------------------
# 行数の数え方（表示用のおおよその規模）
# ---------------------------------------------------------------------------
def _csv_records(path: Path) -> int:
    """CSV のレコード数（セル内改行を考慮。ヘッダ・前置き行も含む物理レコード数）。"""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        return 0
    return sum(1 for _ in csv.reader(io.StringIO(text, newline="")))


def _xlsx_rows(path: Path) -> int:
    """全シートの使用行数（max_row）の合計。"""
    from openpyxl import load_workbook

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(path, read_only=True)
    try:
        return sum(ws.max_row or 0 for ws in wb.worksheets)
    finally:
        wb.close()


def _describe_rows(gen: Generator, paths: list[Path]) -> str:
    if gen.kind == "form":
        n = sum(1 for p in paths if p.suffix == ".xlsx")
        exp = next((p for p in paths if p.name == "_expected.jsonl"), None)
        n_exp = sum(1 for _ in exp.open(encoding="utf-8")) if exp and exp.exists() else 0
        return f"{n}帳票 / 正解{n_exp}行"
    parts = []
    for p in paths:
        if p.suffix == ".csv":
            parts.append(f"{p.stem[:12]}:{_csv_records(p):,}rec")
        elif p.suffix == ".xlsx":
            parts.append(f"{p.stem[:12]}:{_xlsx_rows(p):,}行")
    return " ".join(parts)


# ---------------------------------------------------------------------------
# 実行
# ---------------------------------------------------------------------------
def run_generator(gen: Generator, out_root: Path) -> RunResult:
    result = RunResult(gen)
    t0 = time.perf_counter()
    try:
        module = importlib.import_module(f"scripts.samples.{gen.module}")
        written = module.generate(out_root)
        result.paths = [Path(p) for p in written]
    except Exception as exc:  # 1つ失敗しても残りは続ける
        result.error = f"{type(exc).__name__}: {exc}"
    result.seconds = time.perf_counter() - t0
    if result.error is None:
        result.size_bytes = sum(p.stat().st_size for p in result.paths if p.exists())
        result.rows = _describe_rows(gen, result.paths)
    return result


def _parse_only(text: str | None) -> list[Generator]:
    if not text:
        return list(GENERATORS)
    keys = [k.strip().upper() for k in text.replace("、", ",").split(",") if k.strip()]
    by_key = {g.key: g for g in GENERATORS}
    unknown = [k for k in keys if k not in by_key]
    if unknown:
        raise SystemExit(f"不明な指定: {', '.join(unknown)}（指定できるのは {', '.join(by_key)}）")
    return [g for g in GENERATORS if g.key in keys]   # 指定順ではなく既定の順で実行


def _print_summary(results: list[RunResult], total: float) -> None:
    head = f"{'ID':<3} {'files':>5} {'size(KB)':>10} {'sec':>6}  rows / 内容"
    print()
    print(head)
    print("-" * 78)
    for r in results:
        if r.error:
            print(f"{r.gen.key:<3} {'-':>5} {'-':>10} {r.seconds:>6.1f}  ERROR {r.error}")
            continue
        print(f"{r.gen.key:<3} {len(r.paths):>5} {r.size_bytes / 1024:>10,.0f} {r.seconds:>6.1f}  {r.rows}")
    print("-" * 78)
    n_files = sum(len(r.paths) for r in results)
    size = sum(r.size_bytes for r in results)
    print(f"{'計':<3} {n_files:>5} {size / 1024:>10,.0f} {total:>6.1f}")


# ---------------------------------------------------------------------------
# samples/README.md（目録）
# ---------------------------------------------------------------------------
_DATASET_TEXT: dict[str, str] = {
    "T1": """\
**目的**: 最も素直な一覧表（ベースライン）。1行＝1トラブルの台帳を、列見出し1行＋Excelテーブルで持つ。
**形式**: xlsx（3シート: トラブル一覧 / 設備マスタ / 集計メモ）＋ README（UTF-8）。トラブル一覧は 22列・8,095行。
**意図的な揺れ**: 複数人が1セルに追記した対応内容ログ（日付・記入者の書き方の揺れ、時刻だけの行、同日/翌日、メール転記、新しい順、見出し型、同上）、セル内改行（処置内容の手順、使用部品）、全角英数字・半角ｶﾀｶﾅの混在、手順番号の書き方の違い（1. / (1) / ① / ・）、
IME誤変換・「確認確認」のような重複入力、対応中の行は原因・完了日が空、発生日と発生時刻が別セル、
ライン列の値の型の混在（L1 / L1-L3 / Fab1共通）、瞬低による同時刻の多発、集計メモシートは結合セル＋自由記述＋小さな表。""",
    "T2": """\
**目的**: 保全管理システムからのエクスポートを想定した「コード値だらけ」のCSV。コード表との突合が前提。
**形式**: CSV（CP932 / CRLF）。本体は 45列・33,095レコード（設備故障 8,095＋チョコ停 25,000）、コード表 573行。
**意図的な揺れ**: ヘッダの前に出力条件4行・末尾に「合計件数」行、クォート内のLF改行、カンマ・`""` エスケープを含む項目、
チョコ停行の大量の空欄、旧システム移行データ（全角化・改行→／・費用がカンマ付き文字列・登録者 SYS-MIG）、
社員番号/FE番号/システムユーザーの混在、コード桁数の不統一、Å→U+212B などCP932変換。""",
    "T3": """\
**目的**: 人が作るクロス集計表（設備×年月の停止時間・件数）。数式・結合見出し・小計行を含む。
**形式**: xlsx（年度別4シート＋全期間＋件数）＋ README。数式セル 2,563 個（キャッシュ値あり）。
**意図的な揺れ**: 2段の結合見出し、年度でズレる列位置（重要度列の追加・L6の有無）、小計ラベルの表記揺れ（L1 計 / L1小計 / L1計）、
設置前は「－」（FY2024だけ半角 -）、0 の空欄/数値0 の混在、前年比列の数値と文字の混在、作成日の書式がシートごとに違う、
押印欄、条件付き書式に見える静的な塗り、1シートに2表を縦積み、表の下の※注記。""",
    "T4": """\
**目的**: 月ごとのシートに、ライン別の小さな表が縦に（一部は左右に）並ぶ保全作業記録。表の切り出しが難しいケース。
**形式**: xlsx（記入要領＋月別18シート）＋ README。データ 6,981行（突発 4,899 / 定期・予防 2,082）。
**意図的な揺れ**: 1シート複数表、表ごとに繰り返す見出しの表記揺れ、2段見出し（結合）、非表示の内部コード列、
2025年10月の様式改訂（品番列追加・列移動）、左右に並んだ2表、小計ラベルの揺れ、数値の文字列化（2個 / ¥18,500 / 1.5h）、
全角数字、〃・同上、日付の型の混在（日付セル / 3/12夜 / R7.3.12）、※注記行、該当なしの空表。""",
    "T5": """\
**目的**: 品証・保全が共同で持つ是正処置（CAPA）台帳。長文の複数行テキストと、形式の揃わない日付を多く含む。
**形式**: CSV（UTF-8 BOM付き / LF）5,003行・16列、2026上期分の xlsx（2段見出し）745行、README。
**意図的な揺れ**: 和暦（R5.4.12 / 令和5年…）とISO日付の混在、期限欄の文字（次回PM時 / 至急）、必須項目の空欄、
人名・部署名の書き方の揺れ、設備欄の書き方（ID / ID 名称 / 名称（ID））、全角ID・ハイフン抜けのトラブルNo、
重複起票、全行にセル内改行・カンマ・ダブルクォート、起票者ごとの文体（【見出し】/ ■見出し：/ 箇条書き）。""",
    "F1": """\
**目的**: 1トラブル＝1ファイルの設備修理報告書。3つの版のレイアウト違いに対する項目抽出の評価用。
**形式**: xlsx 30ファイル（v1 / v2 / v3 各10）＋ `_expected.jsonl`（正解値）＋ `_README.md`。写真画像 69枚。
**意図的な揺れ**: 版ごとの配置（左ラベル右値 / B列始まり / 見出し行＋値行）、項目名の同義語（装置No./設備No./設備番号 など）、
「設備番号：IMP-604」形式の1セル内ラベル、改訂後も旧様式を使うファイル、写真・参考資料シート、非表示シート、
必須項目の空欄、日付・数値の全角数字、中間報告、印刷範囲外の写真、日付セルと文字日付（R6.5.14）の混在。""",
    "F2": """\
**目的**: 初動〜恒久対策までを1枚にまとめたトラブル対応報告書。時系列・なぜなぜ分析・恒久対策表など表形式の部品を含む。
**形式**: xlsx 30ファイル（旧様式Rev.3 15 / 新様式Rev.5 15）＋ `_expected.jsonl` ＋ `_README.md`。
**意図的な揺れ**: 新旧で違う項目名（件名/表題、暫定対策/応急処置、真因/根本原因 など）、方眼紙レイアウトと表レイアウト、
発生日時の日付・時刻分割、和暦、停止時間の文字表記（○時間○分）、ラベル内改行（停止\\n時間）、時系列の日付省略、
なぜなぜの途中までの記入、期限の「○月末」、承認なし、ファイル名の付け方の揺れ。""",
    "F3": """\
**目的**: 品質保証の 8D 是正処置報告書（社内・顧客クレーム）。D1〜D8 の節・5W2H・特性要因図・評価表を含む長い帳票。
**形式**: xlsx 30ファイル（5レイアウト、うち顧客クレーム10）＋ `_expected.jsonl` ＋ `_README.md`。
**意図的な揺れ**: 方眼紙2シート / 1シート縦長 / 縦書きラベル列の3系統、D節見出しの揺れ（D1 チーム編成 / D1：チームの結成 / D1）、
項目名の揺れ（管理No./8D No./報告書No.）、5W2H の並べ方3種、日付セル（和暦書式含む）と文字日付の混在、印影画像、
非表示のリストシート、改訂履歴シート、対策前後の写真、全角数字。""",
    "F4": """\
**目的**: 定期点検・予防保全・事後保全の作業報告書。20〜40行のチェックシート（基準値・測定値・判定）を含む。
**形式**: xlsx 30ファイル（Rev.1 A4縦 14 / Rev.2 A4横 16）＋ `_expected.jsonl` ＋ `_README.md`。トレンドシート12（うちグラフ6）。
**意図的な揺れ**: 1表 / 左右2表のチェックシート、表ごとに違う列見出し、〃、固定行数テンプレートの空行、チェックボックス表記（■定期点検 □予防保全）、
非表示列（前回値）、全角数字、測定値空欄で判定○、数値項目に「OK」、○の異体字（◯ / 〇）、承認欄の空欄、旧様式の継続使用。""",
    "F5": """\
**目的**: 製造・生技から保全への工程異常連絡票（発行側と回答側の2部構成）。欠陥マップ画像を含む。
**形式**: xlsx 30ファイル（A4縦 Rev.3 14 / A3横 Rev.2 10 / A3横 Rev.1 6）＋ `_expected.jsonl` ＋ `_README.md`。
**意図的な揺れ**: 回答欄の位置（下 / 右）と版ごとの項目名、回答欄の空欄・一部回答、督促メモ、■/□ のチェックボックス（■と選択肢が別セル）、
日付セルと文字日付（R6.8.25 / 24/8/25）、設備IDと名称が同じセル、数量の文字列化、キャッシュ値のない =SUM()、
ロット一覧の別シート、1セルに「回答日：…　回答部署：…」、縦書きラベル、Sheet1 のままのシート名、記入要領シートの残存。""",
}


def _inventory_rows(out_root: Path) -> list[tuple[str, int, int]]:
    """出力先の tables/ と forms/<様式>/ のファイル一覧（相対パス, ファイル数, バイト数）。

    帳票の .xlsx は様式の版ごとのフォルダに入っているので、版フォルダを1行ずつ出す
    （`forms/<様式>/<版>/*.xlsx`）。`_README.md` と `_expected.jsonl` は様式フォルダの直下。
    """
    rows = []
    tables = out_root / "tables"
    if tables.is_dir():
        for p in sorted(tables.iterdir()):
            if p.is_file():
                rows.append((f"tables/{p.name}", 1, p.stat().st_size))
    forms = out_root / "forms"
    if forms.is_dir():
        for d in sorted(forms.iterdir()):
            if not d.is_dir():
                continue
            for sub in sorted(p for p in d.iterdir() if p.is_dir()):
                xlsx = sorted(sub.glob("*.xlsx"))
                if xlsx:
                    rows.append((f"forms/{d.name}/{sub.name}/*.xlsx", len(xlsx),
                                 sum(p.stat().st_size for p in xlsx)))
            for name in ("_expected.jsonl", "_README.md"):
                if (d / name).exists():
                    rows.append((f"forms/{d.name}/{name}", 1, (d / name).stat().st_size))
    return rows


def write_manifest(out_root: Path) -> Path:
    """samples/README.md を書く。実行時間など毎回変わる値は入れない（再生成で差分が出ないように）。"""
    lines: list[str] = []
    add = lines.append
    add("# サンプルデータ目録（samples/）")
    add("")
    add("RAG前処理アプリ（帳票Excel・一覧表Excel/CSV → Markdown）の動作確認と精度評価に使う、大規模なサンプルデータ。")
    add("半導体工場の **製造部 設備保全課** を想定し、設備・人・部品・アラーム・トラブル履歴を共通のドメインモデル")
    add("（`scripts/samples/domain.py`）から固定シードで生成している。文章は現場の技術者が書いた体裁で、")
    add("略語・書き癖・誤変換・全角半角の混在などを意図的に含む。")
    add("")
    add("## 再生成")
    add("")
    add("```")
    add("python scripts/generate_samples.py                 # 全10種を samples/ に出力")
    add("python scripts/generate_samples.py --only T1,F3    # 一部だけ")
    add("python scripts/generate_samples.py --out <dir>     # 出力先を変える")
    add("python -m scripts.samples.t1_trouble_list          # 個別に実行")
    add("python -m scripts.samples.evaluate_forms           # 帳票の抽出精度を現行アプリで評価")
    add("```")
    add("")
    add("- 乱数はすべて `domain.rng(name)`（固定シード）から取るため、何度実行しても同じ内容になる。")
    add("- 期間: 2023/04/01〜2026/08/31。設備 61台（16種）、トラブル 8,095件、チョコ停 25,000件。")
    add("- `samples/` は .gitignore 済み。直下の `修理報告書_*.xlsx` / `点検記録表.xlsx` は旧来の小さな見本"
        "（`scripts/make_samples.py`）で、本目録の対象外。")
    add("")
    add("## 一覧")
    add("")
    add("| ID | 名前 | 種別 | 生成モジュール |")
    add("|---|---|---|---|")
    for g in GENERATORS:
        add(f"| {g.key} | {g.title} | {'一覧表' if g.kind == 'table' else '帳票'} | `scripts/samples/{g.module}.py` |")
    add("")
    add("### ファイル")
    add("")
    add("| パス | ファイル数 | サイズ(KB) |")
    add("|---|---:|---:|")
    for rel, n, size in _inventory_rows(out_root):
        add(f"| `{rel}` | {n} | {size / 1024:,.0f} |")
    add("")
    add("## 文字コード・形式")
    add("")
    add("| ファイル | 文字コード | 改行 | 備考 |")
    add("|---|---|---|---|")
    add("| T2 の CSV 2本 | CP932（Shift_JIS） | CRLF（セル内はLF） | 前置き4行＋末尾の合計行あり |")
    add("| T5 の CSV | UTF-8（BOM付き） | LF | 全行にセル内改行・カンマ・ダブルクォート |")
    add("| xlsx 全般 | — | — | openpyxl で生成。ZIP内タイムスタンプ・文書プロパティは固定 |")
    add("| README / `_README.md` / `_expected.jsonl` | UTF-8（BOMなし） | LF | 件数などは生成時にデータから計算 |")
    add("")
    add("## 各データセット")
    add("")
    for g in GENERATORS:
        add(f"### {g.key} {g.title}")
        add("")
        for text_line in _DATASET_TEXT[g.key].splitlines():
            add(text_line + "  " if text_line.startswith("**") else text_line)
        add("")
        folder = "tables/" if g.kind == "table" else "forms/<様式>/"
        add(f"詳細は `{folder}` の README を参照。")
        add("")
    add("## 帳票フォルダの構成")
    add("")
    add("帳票（F1〜F5）の .xlsx は**様式の版ごとのフォルダ**に分かれている。版フォルダの中身は .xlsx だけなので、")
    add("フォルダを1つそのまま「帳票取り込み」に置けば、帳票の種類とシートを1回選ぶだけで全部読める。")
    add("`_README.md` と `_expected.jsonl` は様式フォルダの直下に置いたまま。")
    add("")
    add("```")
    add("forms/<様式>/_README.md")
    add("forms/<様式>/_expected.jsonl")
    add("forms/<様式>/<版>/*.xlsx        ← 例: F1_設備修理報告書/Rev1_2019制定/")
    add("```")
    add("")
    add("版フォルダの名前の付け方は F1〜F5 で共通で、「版の名前＋いつから」（例: `Rev1_2019制定`・`Rev5_2025年4月改訂`。")
    add("いちばん古い版など、制定・改訂の時期が帳票に載っていないものは使われていた時期を書く＝`A3横_Rev1_2023まで`）。")
    add("版の名前だけでは区別が付かないものには、区別に要るものだけを足す（F3 は様式番号とシート枚数、F5 は用紙サイズ）。")
    add("版ごとの件数は上の「ファイル」の表と各 `_README.md` を参照。")
    add("")
    add("## 帳票の正解データ（`_expected.jsonl`）")
    add("")
    add("帳票フォルダ（F1〜F5）には1ファイル1行の正解データがある。共通キーは `file`（**版フォルダ名／ファイル名** の相対パス。")
    add("区切りは `/`）、`layout_version`（様式の版）、")
    add("`values`（項目キー → Excelに表示されている値。リストや表は配列）、`labels_used`（そのファイルで実際に使われたラベル文字列）。")
    add("値は原則として Excel の表示どおり（全角数字や和暦もそのまま）。キー名は F1〜F5 で揃えてある")
    add("（`report_id` `equipment_id` `equipment_name` `occurred_at` `reporter` `symptom` `cause` `action` `result` `prevention` など）。")
    add("`scripts/samples/evaluate_forms.py` がこれを使って現行アプリの抽出結果を採点する（結果は `forms/_evaluation.json`）。")
    add("")
    add("## データ間の相互参照")
    add("")
    add("すべてのデータは同じドメインモデルから作っているため、ID で突き合わせられる。")
    add("")
    add("| キー | 形式 | 出てくる場所 |")
    add("|---|---|---|")
    add("| トラブル管理番号 | `TR-YYYY-NNNNN` | T1 管理番号 / T2 管理番号・関連管理番号 / T4 管理No（2025/3〜）/ T5 関連トラブルNo / "
        "F1 管理No.（incident_id）/ F2 管理No（report_id）/ F3 関連報告書No. / F4 関連No.・特記事項 / F5 設備トラブル報告No.・元トラブル |")
    add("| チョコ停番号 | `MS-YYYY-NNNNN` | T2（記録区分=2）。T3 の件数シート下段はこの件数の集計 |")
    add("| 設備番号 | `CMP-101` など | 全データ。T1 設備マスタ / T2 コード表（設備）が台帳 |")
    add("| 社員・担当者 | 氏名 / `M12410` | T1・T4・T5・帳票は氏名（姓のみの場合あり）、T2 は社員番号（コード表で氏名に変換） |")
    add("| 部品 | 品番 `PV55-1270` など | T2 使用部品 / T4 品番（2025/10〜）/ F1・F2・F4 使用部品・交換部品 |")
    add("| 是正処置 | `CA-YYYY-NNNN` | T5 台帳No（本文中の統合先参照も含む） |")
    add("")
    add("- T3 の停止時間は T1/T2 の停止時間（分）とチョコ停の停止時間を、発生月で集計した値（月をまたぐ停止は発生月に全量計上）。")
    add("- T4 の突発行はトラブルの使用部品・処置手順から、定期・予防行は PM 計画から作っている。F4 の定期交換部品と同じ周期。")
    add("- T5 は重要度「中」以上のトラブル＋品質クレームから起票。F3（8D）の関連トラブル30件のうち28件は T5 の関連トラブルNoにも出てくる。")
    add("- 帳票の文章（現象・原因・処置）は T1/T2 の同じトラブル行と同じ出所。ただし帳票側は様式に合わせて敬体化・全角化・要約されていることがある。")
    add("")
    add("## 既知の注意点")
    add("")
    add("- ドメインの文章バンクに由来する「ファラデーカップ清掃・?交換」の `?` が T1/T2 の75件に残っている（F2・F5 は自前で除去）。")
    add("- CMP の EPD 光量シナリオで「約110〜220%」と不自然な値になる文章が47件ある（ドメイン側の既知の不具合）。")
    add("- F1 の写真 PNG は Windows の MS ゴシック（msgothic.ttc）で描画しており、フォントの無い環境では画像のバイト列が変わる。")
    add("")
    path = out_root / "README.md"
    path.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    return path


def generate_large(args) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    targets = _parse_only(args.only)
    out_root = Path(args.out)
    if not out_root.is_absolute():
        out_root = Path.cwd() / out_root
    out_root.mkdir(parents=True, exist_ok=True)

    print(f"出力先: {out_root}")
    t0 = time.perf_counter()
    results = []
    for gen in targets:
        print(f"[{gen.key}] {gen.title} ...", flush=True)
        r = run_generator(gen, out_root)
        print(f"     {'NG ' + r.error if r.error else 'OK'} ({r.seconds:.1f}s)", flush=True)
        results.append(r)
    manifest = write_manifest(out_root)
    total = time.perf_counter() - t0
    _print_summary(results, total)
    print(f"目録: {manifest}")
    return 1 if any(r.error for r in results) else 0


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    parser = argparse.ArgumentParser(description="サンプルデータを生成する（既定: 小さな帳票の見本。--large で T1〜T5 / F1〜F5）")
    parser.add_argument("--large", action="store_true", help="大きなサンプル（一覧表 T1〜T5 / 帳票 F1〜F5）を作る")
    parser.add_argument("--only", help="--large で生成する対象をカンマ区切りで指定（例: T1,F3）")
    parser.add_argument("--out", default=str(DEFAULT_OUT), help="出力先ルート（既定: samples）")
    args = parser.parse_args(argv)
    if args.large:
        return generate_large(args)
    for p in make_all(Path(args.out)):
        print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
