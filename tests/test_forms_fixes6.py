"""帳票まわりの修正（6巡目の確認で見つかった不具合）の回帰テスト。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from excel.extractor import extract_document
from excel.tables import section_of, sections_of
from excel.workbook import load_workbook_info
from models import database as db
from pattern.forms import _section_value
from pattern.model import FieldDef, PatternDef, SheetDef
from tests.test_forms_fixes import _confirmed_doc
from tests.test_output_folder import _set_folder, out_dir  # noqa: F401  (fixture)

FILL = PatternFill("solid", fgColor="DDDDDD")
BOLD = Font(bold=True)


def _head(ws, coord, text):
    ws[coord] = text
    ws[coord].font = BOLD
    ws[coord].fill = FILL


def _label(ws, coord, text, value_coord, value):
    ws[coord] = text
    ws[coord].fill = FILL
    ws[value_coord] = value


def _pattern(*fields):
    return PatternDef(id=1, name="T", version="v1", description="", status="active", image_processing="none",
                      sheets=[SheetDef("報告書", True)], fields=list(fields))


def _field(ex, name):
    return next(f for f in ex["fields"] if f["field_name"] == name)


# ---- R6F-1: 右上の区画・区画の中の小見出し ------------------------------------------------------

def _side_by_side(path, *, issuer=True):
    """右上に「【回答欄】」、その下の行の左に「■ 発行部署記入欄」がある版。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "報告書"
    ws["A1"] = "工程異常連絡書"
    ws["A1"].font = BOLD
    ws["N1"] = "【回答欄】"
    ws["N1"].font = BOLD
    _head(ws, "A3", "■ 発行部署記入欄")
    if issuer:
        _label(ws, "A6", "処置", "B6", "発行側の処置（仮）")
    _label(ws, "N6", "処置", "O6", "回答側の処置（確定）")
    wb.save(path)
    return load_workbook_info(path)


def test_a_right_hand_section_started_higher_keeps_its_columns(tmp_path):
    info = _side_by_side(tmp_path / "cols.xlsx")
    grid = info.grids["報告書"]
    assert section_of(grid, grid.cells[(6, 14)]) == "回答"
    assert section_of(grid, grid.cells[(6, 1)]) == "発行部署記入"
    ex = extract_document(info, _pattern(FieldDef("action", "処置", ["処置"], section="回答")), ["報告書"])
    assert _field(ex, "action")["value"] == "回答側の処置（確定）"


def test_learning_finds_the_right_hand_answer_section(tmp_path):
    from pattern.builder import _learn_sections

    infos = [_side_by_side(tmp_path / "both.xlsx"), _side_by_side(tmp_path / "answer.xlsx", issuer=False)]
    row = {"use": True, "data_type": "string", "field_name": "action", "display_name": "処置",
           "candidates": "処置", "section": ""}
    _learn_sections(infos, {"報告書"}, [row])
    assert row["section"] == "回答"


def _nested(path):
    """「▼ 回答欄」の下に小見出し「1. 暫定対策」があり、その中に回答側の欄がある版。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "報告書"
    _head(ws, "A1", "■ 発行側")
    _label(ws, "A2", "処置内容", "B2", "発行側の処置")
    _head(ws, "A4", "▼ 回答欄")
    _head(ws, "A5", "1. 暫定対策")
    _label(ws, "A6", "処置内容", "B6", "回答側の処置")
    wb.save(path)
    return load_workbook_info(path)


def test_a_sub_heading_inside_the_answer_section_is_still_the_answer_section(tmp_path):
    info = _nested(tmp_path / "nested.xlsx")
    grid = info.grids["報告書"]
    assert sections_of(grid, grid.cells[(6, 1)]) == ["暫定対策", "回答"]
    assert sections_of(grid, grid.cells[(2, 1)]) == ["発行側"]   # 同じ階層の見出しで上の区画は終わる
    field = FieldDef("action", "処置内容", ["処置内容"], section=_section_value("回答欄"))
    assert _field(extract_document(info, _pattern(field), ["報告書"]), "action")["value"] == "回答側の処置"
    # 外側の区画の名前で決めた項目は、内側の小見出しの中の欄だけを見ない（発行側の欄は区画の外）
    issuer = FieldDef("action", "処置内容", ["処置内容"], section="発行側")
    assert _field(extract_document(info, _pattern(issuer), ["報告書"]), "action")["value"] == "発行側の処置"


# ---- R6F-2: 明細表の項目の区画 -----------------------------------------------------------------

def _two_tables(path, *, answer_label="使用部品"):
    wb = Workbook()
    ws = wb.active
    ws.title = "報告書"
    _head(ws, "A1", "■ 発行側")
    _head(ws, "A3", "使用部品")
    for c, t in zip("ABC", ["品番", "品名", "数量"]):
        _head(ws, f"{c}4", t)
    ws["A5"], ws["B5"], ws["C5"] = "P-1", "発行側の部品", 2
    _head(ws, "A8", "▼ 回答欄")
    _head(ws, "A10", answer_label)
    for c, t in zip("ABC", ["品番", "品名", "数量"]):
        _head(ws, f"{c}11", t)
    ws["A12"], ws["B12"], ws["C12"] = "P-9", "回答側の部品", 5
    wb.save(path)
    return load_workbook_info(path)


def _parts(section):
    return FieldDef("parts", "使用部品", ["使用部品"], data_type="table", table_columns=["品番", "品名", "数量"],
                    section=section)


def test_a_table_field_reads_the_table_in_its_section(tmp_path):
    info = _two_tables(tmp_path / "sec.xlsx")
    parts = _field(extract_document(info, _pattern(_parts(_section_value("▼ 回答欄"))), ["報告書"]), "parts")
    assert parts["value"]["rows"] == [["P-9", "回答側の部品", "5"]]
    # 区画を決めていなければ今までどおり上の表
    parts = _field(extract_document(info, _pattern(_parts("")), ["報告書"]), "parts")
    assert parts["value"]["rows"] == [["P-1", "発行側の部品", "2"]]


def test_a_table_found_by_its_columns_prefers_the_one_in_the_section(tmp_path):
    # 回答欄の表の見出しの書き方が違う版: 列見出しで探すときも区画の中の表を選ぶ
    info = _two_tables(tmp_path / "cols.xlsx", answer_label="交換部品")
    field = _parts(_section_value("▼ 回答欄"))
    field.candidates = ["部品明細"]   # どちらの見出しにも当たらない
    parts = _field(extract_document(info, _pattern(field), ["報告書"]), "parts")
    assert parts["value"]["rows"] == [["P-9", "回答側の部品", "5"]]


# ---- R6F-3: 単位だけ直したときの「変更した項目」 --------------------------------------------------

def test_a_unit_only_change_is_listed_as_a_changed_field():
    from views.forms import _changed_fields

    field = {"field_name": "work_hours", "display_name": "作業時間", "value": 2.5, "unit": ""}
    confirmed = {"fields": [field]}
    working = {"fields": [dict(field, unit="時間")]}
    assert _changed_fields(confirmed, working) == ["作業時間"]
    assert _changed_fields(confirmed, {"fields": [dict(field)]}) == []


# ---- UX6-1: まとまりの1件だけを保存先フォルダに保存したとき ------------------------------------------

def test_saving_one_batch_member_points_to_the_next_pending_form(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    first = _confirmed_doc(app, "1.xlsx", batch_id="B", order=0)
    _confirmed_doc(app, "2.xlsx", batch_id="B", order=1)
    with app.app_context():
        pending = db.create_document("3.xlsx", "0" * 64, "documents/3.xlsx", batch_id="B", batch_order=2)
    # 確認文は、この帳票だけがまとまりから消えることを伝える
    page = client.get(f"/forms/{first}").get_data(as_text=True)
    assert "この帳票だけがまとまりから消えます" in page
    page = client.post(f"/forms/{first}/save-to-folder").get_data(as_text=True)
    assert "未確定の1件は残っています" in page
    assert f'class="btn primary" href="/forms/{pending}/' in page and "次の帳票へ" in page


def test_saving_a_single_form_keeps_the_usual_result(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    doc_id = _confirmed_doc(app, "1.xlsx")
    assert "この帳票だけがまとまりから消えます" not in client.get(f"/forms/{doc_id}").get_data(as_text=True)
    page = client.post(f"/forms/{doc_id}/save-to-folder").get_data(as_text=True)
    assert "このアプリからはデータを消しました" in page and "次の帳票へ" not in page
