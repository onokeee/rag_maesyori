"""core/files.py: 保存（分割読み sha256）と Excel の事前チェック。"""
import hashlib
import io
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
        other = upload_path("documents/メモ.txt")
        other.write_text("save_upload が作った名前ではないファイル", encoding="utf-8")

        removed = files.remove_orphan_uploads(core_app.config["UPLOAD_DIR"], [used.stored_path])

        assert removed == 1
        assert upload_path(used.stored_path).exists()
        assert not upload_path(orphan.stored_path).exists()
        assert other.exists()
