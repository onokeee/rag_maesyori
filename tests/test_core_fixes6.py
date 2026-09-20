"""6巡目の修正（core）。

- R6-FUZZ-1: シート全体を指すハイパーリンク・コメントの範囲は、openpyxl で開く前に断る
- R6-SEC-2: 書式だけの空の行（<row ht customHeight/>）が大量にあるブックは、帳票では開く前に断る
- R6-2: imports/<id>/ を消し切れなかったときは中身を 0 バイトにする
"""
import re
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.comments import Comment

from core import purge
from core.files import FORM_MAX_MERGED_CELLS, UploadError, precheck_excel
from tests.conftest import confirmed_import
from tests.test_core_files import upload_error


def _rewrite(src: Path, dest: Path, edit) -> Path:
    """xlsx の部品を edit(名前, 中身) で書き換えて dest に保存する。"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, edit(info.filename, zin.read(info.filename)))
    return dest


def _book(path: Path, *, comment=False) -> Path:
    wb = Workbook()
    wb.active["A1"] = "管理No"
    wb.active["B1"] = "設備"
    if comment:
        wb.active["A1"].comment = Comment("メモ", "作成者")
    wb.save(path)
    return path


def _with_hyperlink(tmp_path, ref: str) -> Path:
    def edit(name, data):
        if name == "xl/worksheets/sheet1.xml":
            link = f'<hyperlinks><hyperlink ref="{ref}" location="A1"/></hyperlinks>'.encode()
            data = data.replace(b"<pageMargins", link + b"<pageMargins", 1)
            assert link in data
        return data
    return _rewrite(_book(tmp_path / "base.xlsx"), tmp_path / f"link_{ref.replace(':', '_')}.xlsx", edit)


def _with_comment(tmp_path, ref: str) -> Path:
    def edit(name, data):
        if re.fullmatch(r"xl/comments/?\w*\.xml", name):
            data, n = re.subn(rb'ref="A1"', f'ref="{ref}"'.encode(), data)
            assert n == 1
        return data
    return _rewrite(_book(tmp_path / "base_c.xlsx", comment=True), tmp_path / f"comment_{ref.replace(':', '_')}.xlsx",
                    edit)


def _with_empty_rows(tmp_path, count: int) -> Path:
    def edit(name, data):
        if name == "xl/worksheets/sheet1.xml":
            extra = "".join(f'<row r="{n}" ht="20" customHeight="1"/>' for n in range(2, count + 2)).encode()
            data = data.replace(b"</row></sheetData>", b"</row>" + extra + b"</sheetData>", 1)
            assert extra[:20] in data
        return data
    return _rewrite(_book(tmp_path / "base_r.xlsx"), tmp_path / f"rows_{count}.xlsx", edit)


# ---- R6-FUZZ-1 -------------------------------------------------------------------------

@pytest.mark.parametrize("make", [_with_hyperlink, _with_comment])
def test_a_link_or_comment_over_the_whole_sheet_is_refused_before_opening(tmp_path, make):
    with pytest.raises(UploadError, match="ハイパーリンクまたはコメントの範囲が大きすぎます"):
        precheck_excel(make(tmp_path, "A1:XFD1048576"))
    with pytest.raises(UploadError, match="ハイパーリンクまたはコメントの範囲が大きすぎます"):
        precheck_excel(make(tmp_path, "A1:XFD50"), max_merged=FORM_MAX_MERGED_CELLS)
    precheck_excel(make(tmp_path, "A1"))            # ふつうのリンク・コメント（1セル）は通す
    precheck_excel(make(tmp_path, "A1:C20"))        # 数セルの範囲も通す


@pytest.mark.parametrize("url", ["/forms/upload", "/tables/upload"])
@pytest.mark.parametrize("make", [_with_hyperlink, _with_comment])
def test_both_upload_routes_refuse_a_whole_sheet_link_or_comment(app, client, tmp_path, url, make):
    data = make(tmp_path, "A1:XFD1048576").read_bytes()
    assert "ハイパーリンクまたはコメントの範囲が大きすぎます" in upload_error(client, url, data, "リンク.xlsx")
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


# ---- R6-SEC-2 --------------------------------------------------------------------------

def test_many_empty_formatted_rows_are_refused_for_forms(tmp_path):
    path = _with_empty_rows(tmp_path, FORM_MAX_MERGED_CELLS + 1)
    assert path.stat().st_size < 2 * 1024 * 1024
    with pytest.raises(UploadError, match="行数が上限"):
        precheck_excel(path, max_merged=FORM_MAX_MERGED_CELLS)
    precheck_excel(path)   # 一覧表（読み取り専用で開く。行の書式を作らない）の上限では通る


def test_row_limit_counts_every_row_exactly(tmp_path):
    path = _with_empty_rows(tmp_path, 99)   # 1行目（値の入った行）＋99行 = 100行
    precheck_excel(path, max_rows=100)
    with pytest.raises(UploadError, match="行数が上限（99 行）"):
        precheck_excel(path, max_rows=99)


def test_forms_upload_refuses_many_empty_formatted_rows(app, client, tmp_path):
    data = _with_empty_rows(tmp_path, FORM_MAX_MERGED_CELLS + 1).read_bytes()
    assert "行数が上限" in upload_error(client, "/forms/upload", data, "行だらけ.xlsx")
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


# ---- R6-2 ------------------------------------------------------------------------------

def _leave_files(monkeypatch):
    """rmtree が掴まれたファイルを消せなかったときと同じく、フォルダとファイルを残す。"""
    monkeypatch.setattr(purge.shutil, "rmtree", lambda *a, **k: None)


def test_leftover_import_files_are_emptied_and_flagged(app, client, monkeypatch):
    import_id = confirmed_import(app, client)
    _leave_files(monkeypatch)
    with app.test_request_context("/"):
        folder = purge.import_dir(import_id)
        assert any(p.stat().st_size for p in folder.rglob("*") if p.is_file())
        assert not purge.purge_incomplete()
        purge.purge_table_import(import_id)
        leftovers = [p for p in folder.rglob("*") if p.is_file()]
        assert leftovers and all(p.stat().st_size == 0 for p in leftovers)
        assert purge.purge_incomplete()


def test_a_clean_purge_is_not_flagged(app, client):
    import_id = confirmed_import(app, client)
    with app.test_request_context("/"):
        purge.purge_table_import(import_id)
        assert not purge.import_dir(import_id).exists()
        assert not purge.purge_incomplete()
