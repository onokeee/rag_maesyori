"""アップロードファイルの保存・事前チェック・削除。

保存は分割して読みながら sha256 を計算する（ブック全体を解析してからハッシュを取らない）。
Excel の事前チェックは openpyxl で開く前に、先頭バイトと zip の目次だけで行う。
"""
from __future__ import annotations

import hashlib
import logging
import posixpath
import re
import shutil
import time
import zipfile
import zlib
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from flask import current_app

logger = logging.getLogger(__name__)

CHUNK_SIZE = 1024 * 1024

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"
STRICT_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main"

# zip 展開の上限（値は仮置き）
ZIP_MAX_TOTAL = 500 * 1024 * 1024
ZIP_MAX_PART = 200 * 1024 * 1024
ZIP_MAX_RATIO = 100
# 小さいパーツは圧縮率が高くても害がないので、この大きさ以上だけ圧縮率を見る
ZIP_RATIO_MIN_BYTES = 16 * 1024 * 1024
# 結合セルの面積（ブック全体の合計）の上限。openpyxl は開くときに結合範囲のセルを1つずつ作るので、
# シート全体の結合（A1:XFD1048576 など）が1つあるだけで、上のサイズ制限より前に固まる（数秒〜終わらない）。
# 行全体の結合（A1:XFD1 = 16,384 セル）や列全体の結合（A:A = 約105万セル）は通す。
MAX_MERGED_CELLS = 2_000_000
# セル数の上限（ブック全体の <c> の数）。openpyxl は開くときにセルを1つずつ作るので、
# 圧縮すると小さいが展開すると大量のセルがあるブック（64KB で 100万セルなど）は、上のサイズ制限を通っても固まる。
# 一覧表は tables.excel_source が別に上限（EXCEL_MAX_CELLS）を持ち「CSVで保存」と案内するので、ここは最後の砦の値。
MAX_CELLS = 1_000_000
# 図形（描画）の上限。openpyxl は開くときに、シートが参照する描画部品（xl/drawings/*.xml）をシートごとに読み直し、
# 図形（アンカー）を1つずつ作る。1つの大きな描画を多数のシートから参照させると、圧縮後は小さくても
# 開くたびに（書類の画面を開くたびにも）数十秒〜終わらない。そこで「描画部品の図形数 × 参照するシート数」と
# 「描画部品の展開後の大きさ × 参照するシート数」の合計を数えて上限を設ける。
# 普通の帳票は1シートに数十個程度の図形なので、十分に余裕のある値にしている。
MAX_DRAWING_ANCHORS = 10_000
MAX_DRAWING_BYTES = 20 * 1024 * 1024
# 図形（アンカー）として数える要素の名前（DrawingML の spreadsheetDrawing）
_ANCHOR_NAMES = ("absoluteAnchor", "oneCellAnchor", "twoCellAnchor")
_DRAWING_REL_SUFFIX = "/drawing"
_DRAWING_ERROR = "図形や画像の数が多すぎるため読み込めません。不要な図形・画像を削除して保存し直してください"
# セルとして数える要素の名前空間（SpreadsheetML 本体の <c> だけ。グラフの <c:chart> などは名前空間が違う）
_SHEET_NAMESPACES = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main", STRICT_NS)
# workbook.xml の名前空間判定で読む先頭バイト数
_WORKBOOK_HEAD_BYTES = 64 * 1024
# 掴まれているファイルを消すときの再試行（ウイルス対策のスキャンなどは短時間で終わる）
REMOVE_RETRIES = 3
REMOVE_RETRY_WAIT = 0.05
# アップロードの保存先（UPLOAD_DIR 直下）と save_upload が付けるファイル名の形
UPLOAD_SUBDIRS = ("documents", "samples", "tables")
_STORED_NAME = re.compile(r"[0-9a-f]{32}\.[A-Za-z0-9]+")


class UploadError(Exception):
    """利用者に見せる日本語メッセージを持つ例外。"""


@dataclass
class StoredFile:
    stored_path: str   # UPLOAD_DIR からの相対パス（/ 区切り）
    file_name: str     # 元のファイル名（表示・ダウンロード用）
    file_hash: str     # sha256
    size: int


def original_name(storage) -> str:
    # secure_filename は日本語を消してしまうので、表示用には元のファイル名（パス部分除去）を使う
    return Path((getattr(storage, "filename", None) or "").replace("\\", "/")).name


def _normalize_ext(ext: str) -> str:
    ext = ext.strip().lower()
    return ext if ext.startswith(".") else f".{ext}"


def _format_size(size: int) -> str:
    return f"{size / 1024 / 1024:.0f}MB" if size >= 1024 * 1024 else f"{size / 1024:.0f}KB"


def upload_path(stored_path: str) -> Path:
    """保存パス → 実ファイルのパス。UPLOAD_DIR の外を指すパスは拒否する。"""
    base = Path(current_app.config["UPLOAD_DIR"]).resolve()
    path = (base / stored_path).resolve()
    if path != base and base not in path.parents:
        raise UploadError("保存先のパスが不正です")
    return path


def save_upload(storage, subdir: str, allowed: set[str], max_bytes: int) -> StoredFile:
    """アップロードを UPLOAD_DIR/subdir に保存する。拡張子・サイズ・空ファイルを確認し、失敗時は消して UploadError。"""
    name = original_name(storage)
    if not name:
        raise UploadError("ファイルを選択してください")
    ext = Path(name).suffix.lower()
    allowed_exts = {_normalize_ext(e) for e in allowed}
    if ext not in allowed_exts:
        kinds = " / ".join(sorted(allowed_exts))
        hint = "（.xls はExcelで .xlsx に保存し直してください）" if ext == ".xls" else ""
        raise UploadError(f"{kinds} のファイルを選んでください{hint}")

    stored = f"{subdir.strip('/')}/{uuid4().hex}{ext}"
    dest = upload_path(stored)
    dest.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256()
    size = 0
    stream = getattr(storage, "stream", storage)
    try:
        with open(dest, "wb") as out:
            while True:
                chunk = stream.read(CHUNK_SIZE)
                if not chunk:
                    break
                size += len(chunk)
                if size > max_bytes:
                    raise UploadError(f"ファイルが大きすぎます（上限 {_format_size(max_bytes)}）")
                digest.update(chunk)
                out.write(chunk)
        if size == 0:
            raise UploadError("ファイルが空です")
    except BaseException:
        dest.unlink(missing_ok=True)
        raise
    return StoredFile(stored_path=stored, file_name=name, file_hash=digest.hexdigest(), size=size)


def precheck_excel(path, max_cells: int | None = None) -> None:
    """Excel（.xlsx/.xlsm）として開いてよいかを、中身を展開せずに確かめる。問題があれば UploadError。

    max_cells: セル数の上限（省略時は MAX_CELLS）。
    """
    path = Path(path)
    with open(path, "rb") as f:
        head = f.read(8)
    if head.startswith(OLE_MAGIC[:4]):
        raise UploadError("パスワード付きのブック、または .xls 形式です。パスワードを外して .xlsx 形式で保存し直してください")
    if not head.startswith(ZIP_MAGIC):
        raise UploadError("Excelファイル（.xlsx / .xlsm）として読み込めません。Excelで開いて .xlsx 形式で保存し直してください")
    try:
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
            names = {i.filename for i in infos}
            if "xl/workbook.bin" in names:
                raise UploadError(".xlsb（バイナリブック）形式は対象外です。Excelで .xlsx 形式で保存し直してください")
            _check_zip_limits(infos)
            workbook = "xl/workbook.xml" if "xl/workbook.xml" in names else next(
                (n for n in sorted(names) if n.lower().endswith("workbook.xml")), None)
            if workbook is None:
                raise UploadError("Excelファイル（.xlsx / .xlsm）ではありません（ブックの情報が見つかりません）")
            with zf.open(workbook) as wb:
                text = wb.read(_WORKBOOK_HEAD_BYTES).decode("utf-8", errors="ignore")
            # シートの置き場所と名前（拡張子）は workbook.xml.rels で自由に決められる（sheet1.dat でも開ける）ので、
            # 名前で選ばずに部品をすべて見る（XML でない部品は、読み始めてすぐ読めなくなって終わる）
            _check_sheet_parts(zf, sorted(names), max_cells or MAX_CELLS)
    # zlib.error / EOFError: 圧縮データが壊れている（zipfile はこれらを包まずにそのまま投げる）
    except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError, RuntimeError, NotImplementedError,
            zlib.error, EOFError) as exc:
        # 例外の種類は画面に出さず（利用者には意味がない）、ログにだけ残す
        logger.warning("Excelファイルの事前チェックで読み込めませんでした: %s", exc.__class__.__name__)
        raise UploadError("Excelファイルとして読み込めません（ファイルが壊れている可能性があります）") from exc
    root = re.search(r"<(?:\w+:)?workbook\b[^>]*>", text)
    if STRICT_NS in (root.group(0) if root else text[:2048]):
        raise UploadError("Strict Open XML 形式のブックです。Excelの「名前を付けて保存」で通常の「Excel ブック (.xlsx)」を選んで保存し直してください")
    # シートが1つも無いブックは断る（読み取り先が無い）。先頭だけ読んでいるので、
    # シートの一覧の終わりが読んだ範囲に無いときは判断しない
    sheets_end = re.search(r"<(?:\w+:)?sheets\s*/>|</(?:\w+:)?sheets>", text)
    if sheets_end and not re.search(r"<(?:\w+:)?sheet\b", text[:sheets_end.end()]):
        raise UploadError("シートがないブックです。シートのあるブックを選んでください")


def _merged_area(ref: str | None) -> int:
    from openpyxl.utils.cell import range_boundaries

    try:
        min_col, min_row, max_col, max_row = range_boundaries(str(ref or "").strip().upper())
    except (ValueError, TypeError):
        return 0   # 読めない範囲は openpyxl 側で扱いが決まる（ここでは数えない）
    if None in (min_col, min_row, max_col, max_row):
        return 0
    return (abs(max_col - min_col) + 1) * (abs(max_row - min_row) + 1)


class _PartCounter:
    """XML の部品1つを分割して読み、結合範囲の面積とセル数を数える（xml.parsers.expat）。

    文字列を正規表現で探すのではなく XML として読むので、ref = "..."（= の前後の空白）、文字参照（&#88;）、
    どんな名前空間の接頭辞（<x.y:c> など）でも、openpyxl（xml.etree＝同じ expat）と同じ解釈で数える。
    """

    def __init__(self, merged: int, cells: int, max_cells: int):
        from xml.parsers import expat

        self.merged, self.cells, self.max_cells = merged, cells, max_cells
        self.anchors = 0              # この部品の図形（アンカー）の数
        self.drawing_targets = []     # この部品（.rels）が参照する描画部品の Target
        self.parser = expat.ParserCreate(namespace_separator=" ")
        self.parser.StartElementHandler = self._start
        # DTD で実体を定義して大量に展開させる細工は、正しいブックには無いので断る
        self.parser.EntityDeclHandler = self._entity

    def _start(self, name: str, attrs: dict) -> None:
        namespace, _, local = name.rpartition(" ")
        if local == "mergeCell":
            self.merged += _merged_area(attrs.get("ref"))
            if self.merged > MAX_MERGED_CELLS:
                raise UploadError("結合セルの範囲が大きすぎます（シート全体・列全体の結合など）。"
                                  "不要な結合を解除して保存し直してください")
        elif local == "c" and namespace in _SHEET_NAMESPACES:
            self.cells += 1
            if self.cells > self.max_cells:
                raise UploadError(f"セル数が上限（{self.max_cells:,} セル）を超えています。"
                                  "不要なシート・範囲を削除して保存し直してください")
        elif local in _ANCHOR_NAMES:
            self.anchors += 1
        elif local == "Relationship" and str(attrs.get("Type", "")).endswith(_DRAWING_REL_SUFFIX)                 and attrs.get("TargetMode") != "External":
            self.drawing_targets.append(str(attrs.get("Target", "")))

    def _entity(self, *_args) -> None:
        raise UploadError("Excelファイルとして読み込めません（不正なファイルの可能性があります）")


def _check_sheet_parts(zf: zipfile.ZipFile, parts: list[str], max_cells: int) -> None:
    """部品の XML を分割して読み、結合範囲の面積の合計とセル数を数え、上限を超えたら UploadError。

    openpyxl で開く前に確かめる（開いた時点で結合範囲のセルと、すべてのセルが作られてしまうため）。
    XML として読めなくなった部品は、そこまでの分だけ数える（画像などの XML でない部品はすぐ終わる）。
    openpyxl も同じ expat で読むので、読めない部品のその先の結合・セルは作られない（開くこと自体が失敗する）。
    """
    from xml.parsers import expat

    merged = 0
    cells = 0
    anchors: dict[str, int] = {}       # 部品 → 図形の数
    references: dict[str, int] = {}    # 描画部品 → 参照される回数（シートごとに読み直されるため）
    for part in parts:
        counter = _PartCounter(merged, cells, max_cells)
        with zf.open(part) as f:
            try:
                while True:
                    chunk = f.read(CHUNK_SIZE)
                    counter.parser.Parse(chunk, not chunk)
                    if not chunk:
                        break
            except expat.ExpatError:
                pass
        merged, cells = counter.merged, counter.cells
        anchors[part] = counter.anchors
        for target in counter.drawing_targets:
            drawing = _resolve_rel_target(part, target)
            references[drawing] = references.get(drawing, 0) + 1
    _check_drawings(zf, anchors, references)


def _resolve_rel_target(rels_part: str, target: str) -> str:
    """.rels の Target を zip 内のパスにする（xl/worksheets/_rels/sheet1.xml.rels の ../drawings/d.xml → xl/drawings/d.xml）。"""
    if target.startswith("/"):
        return target.lstrip("/")
    folder = posixpath.dirname(posixpath.dirname(rels_part))   # _rels の1つ上＝参照元の部品があるフォルダ
    return posixpath.normpath(posixpath.join(folder, target))


def _check_drawings(zf: zipfile.ZipFile, anchors: dict[str, int], references: dict[str, int]) -> None:
    """描画部品を読み直す回数（参照するシートの数）を掛けて、図形の数と大きさの合計が上限を超えたら UploadError。"""
    total_anchors = 0
    total_bytes = 0
    for drawing, count in references.items():
        if drawing not in anchors:
            continue   # 存在しない部品は openpyxl も読まない
        total_anchors += anchors[drawing] * count
        total_bytes += zf.getinfo(drawing).file_size * count
        if total_anchors > MAX_DRAWING_ANCHORS or total_bytes > MAX_DRAWING_BYTES:
            raise UploadError(_DRAWING_ERROR)


def _check_zip_limits(infos: list[zipfile.ZipInfo]) -> None:
    total = 0
    for info in infos:
        total += info.file_size
        if info.file_size > ZIP_MAX_PART:
            raise UploadError(f"ブックの中身が大きすぎます（1つの部品が展開後 {_format_size(info.file_size)}、上限 {_format_size(ZIP_MAX_PART)}）")
        if info.file_size >= ZIP_RATIO_MIN_BYTES and info.file_size > ZIP_MAX_RATIO * max(info.compress_size, 1):
            raise UploadError("ブックの圧縮率が異常に高いため読み込みを中止しました（壊れているか、不正なファイルの可能性があります）")
    if total > ZIP_MAX_TOTAL:
        raise UploadError(f"ブックの中身が大きすぎます（展開後の合計 {_format_size(total)}、上限 {_format_size(ZIP_MAX_TOTAL)}）")


def remove_upload(stored_path: str | None) -> bool:
    """アップロードしたファイルを消す。消せたら True。

    Windows では他のプロセス（ウイルス対策のスキャン、Excel で開いたまま、同期ソフト）や、解析に失敗した
    ブックを掴んだままの openpyxl のせいで消せないことがある。「ダウンロードしたら消えます」（design.md 3.3）と
    言い切っている以上、消せないときも中身だけは 0 バイトに切り詰めて、元のデータが残らないようにする。
    例外は投げない（ここで落とすと、日本語のエラー案内の代わりに 500 になる）。
    残った空ファイルは次の起動時に remove_orphan_uploads が片付ける。
    """
    if not stored_path:
        return True
    try:
        path = upload_path(stored_path)
    except UploadError:
        return False
    for attempt in range(REMOVE_RETRIES):
        try:
            path.unlink(missing_ok=True)
            return True
        except OSError:
            time.sleep(REMOVE_RETRY_WAIT * (attempt + 1))
    try:
        with open(path, "r+b") as f:
            f.truncate(0)
    except OSError:
        pass
    return False


def remove_orphan_import_dirs(tables_dir, known_ids) -> int:
    """DB に無い取り込みの imports/<id>/ フォルダを消す。消した件数を返す。

    取り込みを消し損ねたとき（削除の途中で落ちた・Windows でファイルを掴まれていた）に、読み込んだ行や
    作った Markdown が残り続けないように起動時に片付ける。作業中の取り込みは DB に行があるので消さない。
    """
    root = Path(tables_dir) / "imports"
    if not root.is_dir():
        return 0
    known = {int(i) for i in known_ids}
    removed = 0
    for path in sorted(root.iterdir()):
        if not path.is_dir() or not path.name.isdigit() or int(path.name) in known:
            continue
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists():
            removed += 1
    return removed


# 起動時の片付けで、これより新しいファイルは消さない（保存してから DB の行を作るまでの間を守る）。
# core.jobs.STALE_AFTER（2分）と同じ長さ。
ORPHAN_GRACE_SECONDS = 120


def remove_orphan_uploads(base, known_paths) -> int:
    """DB のどこからも参照されていないアップロード済みファイルを消す。消した件数を返す。

    取り込みの途中で失敗し、消し損ねたファイルが残ることがある（画面からは消せない）ので起動時に片付ける。
    save_upload が作った名前（uuid + 拡張子）のファイルだけを対象にする。
    """
    base = Path(base)
    known = {str(p).replace("\\", "/").strip("/") for p in known_paths if p}
    removed = 0
    for subdir in UPLOAD_SUBDIRS:
        for path in sorted((base / subdir).glob("*")):
            if not path.is_file() or not _STORED_NAME.fullmatch(path.name):
                continue
            if f"{subdir}/{path.name}" in known:
                continue
            try:
                if time.time() - path.stat().st_mtime < ORPHAN_GRACE_SECONDS:
                    continue   # できたばかり: 別に起動しているアプリが取り込みの途中（DB の行を作る前）かもしれない
            except OSError:
                continue
            try:
                path.unlink()
            except OSError:
                continue
            removed += 1
    return removed
