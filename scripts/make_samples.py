"""動作確認用のサンプル帳票Excelを作る。

    python scripts/make_samples.py            # samples/ に出力

同じ「設備修理報告書」でもレイアウト・項目名が異なる3種類と、別帳票（点検記録表）を作る。
"""
from __future__ import annotations

import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image as PILImage, ImageDraw

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


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "samples"
    for p in make_all(out):
        print(p)
