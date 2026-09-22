"""覚えたシート（forms.WorkbookInfo.to_json / database.pattern_books）の突き合わせ: 実物の見本で確かめる。

tests/test_form_books.py（作った側のテスト。生成した小さなブックが中心）に、実物の帳票での確認を足す
（2026-09-22。利用者の指示「再度シートを置かなくても、登録したときにシートのセル番地と文字情報を記憶しておけばだせるはず」
「セル色の情報は不要」）。

ここで確かめること:
  - 結合した見出し・縦に結合した明細表の見出し・日時や数値のセル・画像のあるシート・4枚のシートを持つ実物
    （F1 Rev3 の修理報告書、F3 の2シートの 8D 報告書）で、画面と同じ道すじ（登録 → クリック）を **Excel を送らずに**
    通しても、できる項目も読み取りの結果も、Excel を送ったときと同じ
  - 覚えた JSON に書いてあるのは決めたものだけ（版・1904年基準・シート名・非表示・結合・セル・数式・画像の場所）で、
    色らしいもの（RGB・テーマ色・色の名前）は無い。塗りは有る無しと組の番号だけ
  - 実物の全部の文字セルをクリックしても、開いたブックと覚えたシートで同じ項目になる
  - 開き直した画面に、シートの切り替え・日時の文字・画像の件数がそのまま出る
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict
from pathlib import Path

import pytest
from openpyxl.utils import get_column_letter

from app import core
from app import database as db
from app.forms import WorkbookInfo, click_field, extract_document, load_workbook_info, match_pattern, table_cells
from tests.test_forms import add_field, create_type, panel_html

SAMPLES = Path(__file__).resolve().parents[1] / "samples" / "forms"
# 結合した見出し（A12:A14「症状内容」）・縦に結合した明細表の見出し（A27:A30「使用部品」）・日時（A8）・数値（E8）・
# 写真のシート（画像6件）を持つ修理報告書
F1_REV3 = SAMPLES / "F1_設備修理報告書" / "Rev3_2025改訂" / "MR-2510-151_ETC-309_修理報告書.xlsx"
# 2枚の報告シートと「リスト」「改訂履歴」の4シート。塗りの組が3つ（見出し・区画・列見出し）ある 8D 報告書
F3_TWO_SHEETS = SAMPLES / "F3_8D是正処置報告書" / "QA-F-021_Rev3_2シート_2023改訂" / "8D-2024-029_IMP-604_8D.xlsx"

# 色らしい文字。JSON のどこにも出てはいけない（セルの文字にたまたま含まれる形は、この2つの見本には無い）
COLOUR_LIKE = re.compile(r"rgb|theme|tint|argb|indexed|fgColor|bgColor|#[0-9A-Fa-f]{6}\b|\bFF[0-9A-F]{6}\b", re.I)


@pytest.fixture(autouse=True)
def _clean_cache():
    """テストごとに、メモリに覚えているブックを空にする（プロセスで1つの置き場なので）。"""
    core.clear()
    yield
    core.clear()


def _sample(path: Path) -> Path:
    if not path.exists():
        pytest.skip(f"サンプルがありません: {path}")
    return path


def _fields_of(app, pattern_id: int) -> list[dict]:
    with app.app_context():
        return [asdict(f) for f in db.load_pattern(pattern_id).fields]


def _click_without_excel(client, pattern_id: int, sheet: str, label: str, value: str = "") -> dict:
    """セルのクリックだけを送る（ファイルも合図も無し＝覚えたシートで作る）。"""
    res = client.post(f"/form-types/{pattern_id}/fields", data={"sheet": sheet, "label_cell": label, "value_cell": value})
    assert res.status_code == 200, res.get_data(as_text=True)
    return res.get_json()


def _stored_info(app, pattern_id: int) -> WorkbookInfo:
    with app.app_context():
        return WorkbookInfo.from_json(db.load_pattern_book(pattern_id)["book_json"])


def _extraction(info: WorkbookInfo, pattern) -> dict:
    sheets = match_pattern(info, pattern).sheet_names or info.sheet_names[:1]
    out = extract_document(info, pattern, sheets)
    return {"sheets": sheets, "values": out["values"], "fields": out["fields"], "attachments": out["attachments"]}


# ---- 実物の帳票で、Excel を送らずに同じ項目・同じ読み取りになる --------------------------------------

@pytest.mark.samples
def test_f1_rev3_clicks_without_the_excel_make_the_same_fields_and_values(app, client):
    """修理報告書（結合した見出し・明細表・日時・数値・写真のシート）: Excel を送るときと覚えたシートで、項目も読み取りも同じ。"""
    path = _sample(F1_REV3)
    info = load_workbook_info(path)
    sheet = info.sheet_names[0]
    clicks = [("A5", "A6"),      # 報告番号（横に結合した見出しと値）
              ("A7", "A8"),      # 発生日時（値は datetime。覚えたシートでも型のまま戻る）
              ("E7", "E8"),      # 停止時間(分)（値は int。単位「分」）
              ("A12", "B12"),    # 症状内容（A12:A14 と B12:L14 の結合）
              ("A23", "B23"),    # 修理内容（複数行の文章）
              ("B27", "")]       # 使用部品の明細表（列見出しの1回クリック。A27:A30 が縦に結合した見出し）

    # Excel を毎回送る（登録したときの道すじ）
    fresh = create_type(client, path, "送る")
    for label, value in clicks:
        add_field(client, fresh, sheet, label, value)

    # 同じ Excel で登録だけして、あとは Excel を送らない（開き直したときの道すじ）
    kept = create_type(client, path, "覚え")
    core.clear()
    last = None
    for label, value in clicks:
        last = _click_without_excel(client, kept, sheet, label, value)
    assert "登録したときに覚えたシートを実際に読み取った結果です" in last["html"]

    fresh_fields, kept_fields = _fields_of(app, fresh), _fields_of(app, kept)
    assert [f["field_name"] for f in kept_fields] == ["report_id", "occurred_date", "downtime", "symptom", "repair", "parts"]
    assert kept_fields == fresh_fields
    table = next(f for f in kept_fields if f["field_name"] == "parts")
    assert table["data_type"] == "table" and table["table_columns"][:2] == ["品番", "品名"]
    assert next(f for f in kept_fields if f["field_name"] == "downtime")["unit"] == "分"

    # その設定での読み取り: 開いたブックと覚えたシートで、値・項目・画像の場所が同じ
    with app.app_context():
        pattern = db.load_pattern(kept)
    from_book, from_snapshot = _extraction(info, pattern), _extraction(_stored_info(app, kept), pattern)
    assert from_snapshot == from_book
    assert from_snapshot["values"]["occurred_date"] == info.grids[sheet].cells[(8, 1)].text     # 日時のまま（「2025-10-24 12:13」）
    assert from_snapshot["values"]["downtime"] == 2120
    assert from_snapshot["values"]["parts"]["rows"], "明細表の行が読めているはず"
    assert len(from_snapshot["attachments"]) == len(info.images_in(from_book["sheets"]))
    assert f"{len(clicks)}項目中 <strong>{len(clicks)}</strong>項目が見つかりました" in last["html"]
    assert not list(Path(app.config["UPLOAD_DIR"]).rglob("*"))


@pytest.mark.samples
def test_f3_two_sheet_report_clicks_on_both_sheets_from_the_snapshot(app, client):
    """4シートの 8D 報告書: 2枚の報告シートの欄と明細表を Excel を送らずにクリックしても、送ったときと同じ。"""
    path = _sample(F3_TWO_SHEETS)
    info = load_workbook_info(path)
    first, second = info.sheet_names[:2]
    assert len(info.sheet_names) == 4
    clicks = [(first, "A6", "E6"),      # 8D No.
              (first, "A10", "E10"),    # 発生日時（datetime）
              (first, "Y10", "AC10"),   # 停止時間（int）
              (first, "A14", ""),       # D1 Team の明細表（A14:E14 …）
              (second, "A14", "G14"),   # 2枚目の「選定理由」
              (second, "A10", "")]      # 2枚目の対策案の表（列見出し「No」）

    fresh = create_type(client, path, "送る")
    for sheet, label, value in clicks:
        add_field(client, fresh, sheet, label, value)
    kept = create_type(client, path, "覚え")
    core.clear()
    for sheet, label, value in clicks:
        _click_without_excel(client, kept, sheet, label, value)

    assert _fields_of(app, kept) == _fields_of(app, fresh)
    with app.app_context():
        pattern = db.load_pattern(kept)
        assert [s.sheet_name for s in pattern.sheets] == [first, second]
    from_book, from_snapshot = _extraction(info, pattern), _extraction(_stored_info(app, kept), pattern)
    assert from_snapshot == from_book
    assert from_snapshot["sheets"] == [first, second]
    assert from_snapshot["values"]["occurred_date"] == info.grids[first].cells[(10, 5)].text


# ---- 覚えた JSON に書いてあるもの ------------------------------------------------------------------

@pytest.mark.samples
@pytest.mark.parametrize("path", [F1_REV3, F3_TWO_SHEETS], ids=["F1_Rev3", "F3_2sheets"])
def test_the_snapshot_json_of_a_real_report_has_only_the_agreed_keys_and_no_colour(path):
    """JSON の項目は決めたものだけ。色らしい文字は無く、塗りは 0/1 と小さな組の番号だけ。大きさは数KB〜数十KB。"""
    info = load_workbook_info(_sample(path))
    text = info.to_json()
    assert not COLOUR_LIKE.search(text), COLOUR_LIKE.search(text)
    data = json.loads(text)
    assert set(data) == {"version", "date1904", "sheets", "images"}
    assert data["version"] == 1 and data["date1904"] is False
    groups: set[int] = set()
    for sheet in data["sheets"]:
        assert set(sheet) == {"name", "hidden", "merged", "cells", "uncached"}
        assert all(len(b) == 4 and all(isinstance(n, int) for n in b) for b in sheet["merged"])
        for item in sheet["cells"]:
            assert len(item) == 9
            row, col, max_row, max_col, value, bold, filled, fill, fmt_unit = item
            assert 1 <= row <= max_row and 1 <= col <= max_col
            assert isinstance(value, (str, int, float, bool, dict))
            if isinstance(value, dict):
                assert set(value) == {"t", "v"} and value["t"] in ("datetime", "date", "time", "timedelta")
            assert bold in (0, 1) and filled in (0, 1)
            assert isinstance(fill, int) and 0 <= fill <= 10 and (fill > 0) == (filled == 1)
            assert isinstance(fmt_unit, str)
            groups.add(fill)
    assert groups - {0}, "見出しの塗りが組の番号として残るはず"
    for img in data["images"]:
        assert set(img) == {"type", "sheet", "location", "name"}
    assert 1_000 < len(text.encode("utf-8")) < 64 * 1024


@pytest.mark.samples
@pytest.mark.parametrize("path", [F1_REV3, F3_TWO_SHEETS], ids=["F1_Rev3", "F3_2sheets"])
def test_every_cell_of_a_real_report_clicks_the_same_from_the_snapshot(path):
    """全部の文字セルを見出しとして（値なし・右隣・下・同じセル）クリックしても、開いたブックと覚えたシートで同じ項目。"""
    info = load_workbook_info(_sample(path))
    back = WorkbookInfo.from_json(info.to_json())
    compared = 0
    for name, grid in info.grids.items():
        other = back.grids[name]
        assert table_cells(other) == table_cells(grid), name
        for cell in grid.text_cells():
            label = cell.coord.split(":")[0]
            right = f"{get_column_letter(cell.max_col + 1)}{cell.row}"
            below = f"{get_column_letter(cell.col)}{cell.max_row + 1}"
            for value in ("", right, below, label):
                assert click_field(other, label, value, ()) == click_field(grid, label, value, ()), (name, label, value)
                compared += 1
    assert compared > 100


# ---- 覚えたシートが入れ替わるのは、ファイルを実際に置いたときだけ ----------------------------------------

def test_a_panel_post_without_a_real_file_does_not_replace_the_remembered_sheet(app, client, tmp_path):
    """名前の無い空のファイル部品と合図（book_hash）だけの POST panel では、覚えたシートを入れ替えない。

    クリックに添えた別の Excel はメモリの置き場に残るので、その合図で「別のExcelに替える」の入口を叩いても
    覚えは変わらない（入れ替わるのは、ファイルを実際に置いたときだけ）。
    """
    import io

    from tests.test_form_books import _report

    first = _report(tmp_path / "一番.xlsx", "R-1")
    second = _report(tmp_path / "二番.xlsx", "R-2")
    pattern_id = create_type(client, first, "置き替えない")
    with app.app_context():
        before = db.pattern_book_meta(pattern_id)
    add_field(client, pattern_id, "報告書", "A1", "B1", book=second)     # 二番はメモリの置き場に入る
    panel = panel_html(client, pattern_id, book=second)                 # 二番を置いた＝覚えも二番に替わる
    second_hash = panel.split('data-book="')[1].split('"')[0]
    panel_html(client, pattern_id, book=first)                          # 一番を置き直す＝覚えは一番に戻る
    with app.app_context():
        assert db.pattern_book_meta(pattern_id)["file_hash"] == before["file_hash"]

    # 二番の合図＋空の部品で panel を叩いても、覚えは一番のまま（画面には二番が出る）
    res = client.post(f"/form-types/{pattern_id}/panel",
                      data={"book": (io.BytesIO(b""), ""), "book_hash": second_hash},
                      content_type="multipart/form-data")
    assert res.status_code == 200, res.get_data(as_text=True)
    assert res.get_json()["message"] == ""
    assert "R-2" in res.get_json()["html"]
    with app.app_context():
        row = db.pattern_book_meta(pattern_id)
    assert row["file_hash"] == before["file_hash"] and row["file_name"] == "一番.xlsx"
    core.clear()
    assert "R-1" in panel_html(client, pattern_id) and "R-2" not in panel_html(client, pattern_id)


# ---- 開き直した画面 --------------------------------------------------------------------------------

@pytest.mark.samples
def test_reopening_a_real_report_shows_its_sheets_typed_values_and_image_count(app, client):
    """開き直すと、2枚のシートの切り替え・日時の文字・写真の件数・覚えた旨がそのまま出て、置き場は出ない。"""
    path = _sample(F1_REV3)
    info = load_workbook_info(path)
    pattern_id = create_type(client, path, "修理報告書")
    core.clear()

    panel = panel_html(client, pattern_id)
    for name in info.sheet_names:
        assert f">{name}</button>" in panel                 # シートの切り替え
    assert info.grids[info.sheet_names[0]].cells[(8, 1)].text in panel     # 発生日時（datetime から作る文字）
    assert len(info.images) == 6
    for name in info.sheet_names:                                          # 画像の場所も覚えている（報告書に2件・写真に4件）
        assert f"このシートに画像が{len(info.images_in([name]))}件あります。" in panel
    assert f"登録したときに覚えたシートです（{path.name}、" in panel
    assert 'id="drop-book-edit"' not in panel and "別のExcelに替える" in panel
    with app.app_context():
        meta = db.pattern_book_meta(pattern_id)
    assert meta["file_name"] == path.name and 1_000 < meta["size"] < 64 * 1024
