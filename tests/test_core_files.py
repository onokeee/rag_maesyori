"""core/files.py: 保存（分割読み sha256）と Excel の事前チェック。"""
import hashlib
import io
import os
import re
import time
import zipfile
from pathlib import Path

import pytest
from flask import Flask
from openpyxl import Workbook
from werkzeug.datastructures import FileStorage

from core import files
from core.files import UploadError, precheck_excel, remove_upload, save_upload, upload_path


@pytest.fixture
def core_app(tmp_path):
    app = Flask(__name__)
    app.config.update(TESTING=True, UPLOAD_DIR=tmp_path / "uploads")
    return app


def _storage(data: bytes, name: str) -> FileStorage:
    return FileStorage(stream=io.BytesIO(data), filename=name)


def _xlsx(path):
    wb = Workbook()
    wb.active["A1"] = "管理No"
    wb.save(path)
    return path


def test_save_upload_hash_and_remove(core_app, monkeypatch):
    monkeypatch.setattr(files, "CHUNK_SIZE", 7)  # 分割読みを確かめる
    data = "管理No,設備\nTR-1,CMP-101\n".encode("cp932") * 50
    with core_app.app_context():
        stored = save_upload(_storage(data, "C:\\fakepath\\故障履歴.csv"), "tables", {".csv", "xlsx"}, 10_000)
        assert stored.file_name == "故障履歴.csv"
        assert stored.stored_path.startswith("tables/") and stored.stored_path.endswith(".csv")
        assert stored.file_hash == hashlib.sha256(data).hexdigest()
        assert stored.size == len(data)
        path = upload_path(stored.stored_path)
        assert path.read_bytes() == data
        remove_upload(stored.stored_path)
        assert not path.exists()
        remove_upload(stored.stored_path)  # 2回目も例外にしない


def test_save_upload_rejects(core_app):
    with core_app.app_context():
        with pytest.raises(UploadError, match="保存し直して"):
            save_upload(_storage(b"x", "old.xls"), "forms", {".xlsx"}, 100)
        with pytest.raises(UploadError, match="大きすぎます"):
            save_upload(_storage(b"x" * 101, "big.csv"), "tables", {".csv"}, 100)
        with pytest.raises(UploadError, match="空です"):
            save_upload(_storage(b"", "empty.csv"), "tables", {".csv"}, 100)
        with pytest.raises(UploadError, match="選択"):
            save_upload(_storage(b"x", ""), "tables", {".csv"}, 100)
        # 失敗したファイルは残さない
        folder = core_app.config["UPLOAD_DIR"] / "tables"
        assert not folder.exists() or not any(folder.iterdir())
        with pytest.raises(UploadError):
            upload_path("../outside.txt")


def test_precheck_accepts_normal_xlsx(tmp_path):
    precheck_excel(_xlsx(tmp_path / "ok.xlsx"))


def test_precheck_ole_header(tmp_path):
    path = tmp_path / "locked.xlsx"
    path.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 504)
    with pytest.raises(UploadError, match="パスワード付き"):
        precheck_excel(path)


def test_precheck_not_zip(tmp_path):
    path = tmp_path / "text.xlsx"
    path.write_text("管理No,設備", encoding="utf-8")
    with pytest.raises(UploadError, match="読み込めません"):
        precheck_excel(path)


def test_precheck_xlsb(tmp_path):
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("xl/workbook.bin", b"\x83\x01\x00")
    with pytest.raises(UploadError, match="xlsb"):
        precheck_excel(path)


def test_precheck_strict(tmp_path):
    path = tmp_path / "strict.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook '
                    'xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheets/></workbook>')
    with pytest.raises(UploadError, match="Strict"):
        precheck_excel(path)


def test_precheck_zip_bomb_ratio(tmp_path):
    path = tmp_path / "bomb.xlsx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", "<workbook/>")
        zf.writestr("xl/worksheets/sheet1.xml", b"\x00" * (files.ZIP_RATIO_MIN_BYTES + 1024))
    assert path.stat().st_size < 1024 * 1024
    with pytest.raises(UploadError, match="圧縮率"):
        precheck_excel(path)


def test_precheck_zip_limits(tmp_path, monkeypatch):
    path = tmp_path / "large.xlsx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", "<workbook/>")
        zf.writestr("xl/worksheets/sheet1.xml", b"a" * 5000)
        zf.writestr("xl/worksheets/sheet2.xml", b"b" * 5000)
    monkeypatch.setattr(files, "ZIP_MAX_PART", 4000)
    with pytest.raises(UploadError, match="1つの部品"):
        precheck_excel(path)
    monkeypatch.setattr(files, "ZIP_MAX_PART", 6000)
    monkeypatch.setattr(files, "ZIP_MAX_TOTAL", 8000)
    with pytest.raises(UploadError, match="合計"):
        precheck_excel(path)


def test_remove_upload_empties_a_locked_file_and_reports_it(core_app, monkeypatch):
    """他のプロセスに掴まれていて消せないときは、例外を出さずに中身を空にして False を返す。

    「ダウンロードしたらこのPCから消えます」（design.md 3.3）と言い切っているので、ファイル名が残っても
    元の Excel/CSV の中身は残さない（空になったファイルは次の起動時に片付く）。
    """
    with core_app.app_context():
        stored = save_upload(_storage("社外秘の中身".encode("utf-8") * 10, "掴まれたブック.xlsx"),
                             "documents", {".xlsx"}, 10_000)
        path = upload_path(stored.stored_path)

        def locked(self, missing_ok=False):
            raise PermissionError(32, "別のプロセスが使用中です")

        monkeypatch.setattr(Path, "unlink", locked)
        monkeypatch.setattr(files, "REMOVE_RETRY_WAIT", 0)
        assert remove_upload(stored.stored_path) is False   # 例外を出さず、消せなかったことを返す
        assert path.exists() and path.read_bytes() == b""   # 中身は残さない

        monkeypatch.undo()
        assert remove_upload(stored.stored_path) is True
        assert not path.exists()


def test_remove_orphan_uploads(core_app):
    """DB から参照されていない取り込み済みファイルだけを片付ける。"""
    with core_app.app_context():
        used = save_upload(_storage(b"a" * 100, "使用中.xlsx"), "documents", {".xlsx"}, 10_000)
        orphan = save_upload(_storage(b"b" * 100, "残骸.xlsx"), "tables", {".xlsx"}, 10_000)
        fresh = save_upload(_storage(b"c" * 100, "保存したばかり.xlsx"), "documents", {".xlsx"}, 10_000)
        other = upload_path("documents/メモ.txt")
        other.write_text("save_upload が作った名前ではないファイル", encoding="utf-8")
        old = time.time() - 3600
        for stored in (used, orphan):
            os.utime(upload_path(stored.stored_path), (old, old))

        removed = files.remove_orphan_uploads(core_app.config["UPLOAD_DIR"], [used.stored_path])

        assert removed == 1
        assert upload_path(used.stored_path).exists()
        assert not upload_path(orphan.stored_path).exists()
        assert other.exists()
        # できたばかりのファイルは、別に起動しているアプリが DB の行を作る前かもしれないので消さない
        assert upload_path(fresh.stored_path).exists()


# ---- 結合セルの面積（openpyxl で開く前に止める） ------------------------------------------------

def _xlsx_with_merge(path, ref: str, *, pad: int = 0):
    """ふつうの xlsx を作り、シートの XML に <mergeCell ref="..."> を書き足す（pad: その前に入れる空白の量）。"""
    src = _xlsx(path.with_name("src_" + path.name))
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                merge = f'{" " * pad}<mergeCells count="1"><mergeCell ref="{ref}"/></mergeCells>'.encode()
                data = data.replace(b"</sheetData>", b"</sheetData>" + merge, 1)
                if b"<sheetData/>" in data:
                    data = data.replace(b"<sheetData/>", b"<sheetData/>" + merge, 1)
            zout.writestr(info, data)
    return path


def test_precheck_refuses_a_whole_sheet_merge_before_openpyxl(tmp_path):
    import time
    path = _xlsx_with_merge(tmp_path / "merge_full.xlsx", "A1:XFD1048576")
    started = time.monotonic()
    with pytest.raises(UploadError, match="結合セルの範囲が大きすぎます"):
        precheck_excel(path)
    assert time.monotonic() - started < 2


def test_precheck_accepts_whole_row_and_whole_column_merges(tmp_path):
    precheck_excel(_xlsx_with_merge(tmp_path / "row.xlsx", "A1:XFD1"))
    precheck_excel(_xlsx_with_merge(tmp_path / "col.xlsx", "A1:A1048576"))


def test_precheck_finds_a_merge_split_across_read_chunks(tmp_path, monkeypatch):
    monkeypatch.setattr(files, "CHUNK_SIZE", 64)   # タグがチャンクの境目で切れるようにする
    for pad in range(0, 64, 7):
        path = _xlsx_with_merge(tmp_path / f"split{pad}.xlsx", "C1:Z200000", pad=pad)
        with pytest.raises(UploadError, match="結合セル"):
            precheck_excel(path)


def test_precheck_counts_each_merge_once(tmp_path, monkeypatch):
    monkeypatch.setattr(files, "CHUNK_SIZE", 64)
    monkeypatch.setattr(files, "MAX_MERGED_CELLS", 100)
    precheck_excel(_xlsx_with_merge(tmp_path / "exact.xlsx", "A1:J10"))   # ちょうど100セルは通す


def test_uploads_with_a_whole_sheet_merge_are_refused_in_japanese(app, client, tmp_path):
    path = _xlsx_with_merge(tmp_path / "merge_full.xlsx", "A1:XFD1048576")
    for url in ("/forms/upload", "/tables/upload"):
        res = client.post(url, data={"file": (io.BytesIO(path.read_bytes()), "結合.xlsx")},
                          content_type="multipart/form-data", follow_redirects=True)
        assert "結合セルの範囲が大きすぎます" in res.get_data(as_text=True), url
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


def test_precheck_reports_a_damaged_compressed_part(tmp_path):
    """圧縮データが壊れている（zlib.error）ときも UploadError（500 にしない）。"""
    path = _xlsx(tmp_path / "ok.xlsx")
    data = bytearray(path.read_bytes())
    with zipfile.ZipFile(path) as zf:
        info = zf.getinfo("xl/workbook.xml")
    start = info.header_offset + 30 + len(info.filename.encode()) + len(info.extra)
    for i in range(start, start + min(info.compress_size, 64)):
        data[i] ^= 0xFF
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(bytes(data))
    with pytest.raises(UploadError, match="壊れている"):
        precheck_excel(broken)


def test_sheetless_workbook_is_refused_for_both_flows(app, client, tmp_path):
    """シートの無いブックは帳票でも一覧表でも受け付けず、アップロードしたファイルも残さない。"""
    import re

    path = _xlsx(tmp_path / "s.xlsx")
    out = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(out, "w") as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = re.sub(rb"<sheets>.*?</sheets>", b"<sheets/>", data, flags=re.S)
            z.writestr(item, data)
    sheetless = tmp_path / "sheetless.xlsx"
    sheetless.write_bytes(out.getvalue())
    with pytest.raises(UploadError, match="シートがないブック"):
        precheck_excel(sheetless)
    precheck_excel(path)   # ふつうのブックは通る
    for url in ("/forms/upload", "/tables/upload"):
        res = client.post(url, data={"file": (io.BytesIO(out.getvalue()), "シートなし.xlsx")},
                          content_type="multipart/form-data", follow_redirects=True)
        assert "シートがないブックです" in res.get_data(as_text=True), url
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


def _xlsx_with_cells(path, count: int):
    """ふつうの xlsx のシートに、書式だけのセル <c/> を count 個書き足す（圧縮すると小さい）。"""
    src = _xlsx(path.with_name("src_" + path.name))
    rows = b'<row r="9999">' + b'<c s="0"/>' * count + b"</row>"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data, n = re.subn(rb"<sheetData\s*/>", b"<sheetData>" + rows + b"</sheetData>", data, 1)
                if not n:
                    data = data.replace(b"</sheetData>", rows + b"</sheetData>", 1)
            zout.writestr(info, data)
    return path


def test_precheck_refuses_too_many_cells_in_a_small_file(tmp_path):
    """圧縮すると小さいが展開すると大量のセルがあるブックは、openpyxl で開く前に断る。"""
    path = _xlsx_with_cells(tmp_path / "many.xlsx", files.MAX_CELLS + 1)
    assert path.stat().st_size < 200 * 1024
    with pytest.raises(UploadError, match="セル数が上限"):
        precheck_excel(path)
    precheck_excel(path, max_cells=files.MAX_CELLS + 100)   # 上限を上げれば通る


def test_precheck_counts_cells_split_across_read_chunks_once(tmp_path, monkeypatch):
    monkeypatch.setattr(files, "CHUNK_SIZE", 64)
    path = _xlsx_with_cells(tmp_path / "exact.xlsx", 500)
    with zipfile.ZipFile(path) as zf:
        existing = len(re.findall(rb"<(?:\w+:)?c(?=[\s/>])", zf.read("xl/worksheets/sheet1.xml")))
    precheck_excel(path, max_cells=existing)          # ちょうど上限は通す（二重に数えない）
    with pytest.raises(UploadError, match="セル数が上限"):
        precheck_excel(path, max_cells=existing - 1)  # 1つでも多ければ断る（数え落とさない）


def test_save_upload_and_precheck_messages_carry_no_file_name_or_class_name(app, tmp_path):
    """flash に入るメッセージにはファイル名も例外の種類名も入れない（design.md 3.3）。"""
    from werkzeug.datastructures import FileStorage

    with app.app_context():
        for data, allowed in ((b"x", {".csv"}), (b"", {".xlsx"})):
            with pytest.raises(UploadError) as err:
                save_upload(FileStorage(io.BytesIO(data), "秘密の名前.xlsx"), "documents", allowed, 1024)
            assert "秘密の名前" not in str(err.value)
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"PK\x03\x04" + b"\x00" * 64)
    with pytest.raises(UploadError) as err:
        precheck_excel(broken)
    assert "BadZipFile" not in str(err.value) and "壊れている" in str(err.value)
