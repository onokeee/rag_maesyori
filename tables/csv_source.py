"""CSV/TSV の読み込み。文字コード・区切り文字・前置き行を判定し、1レコードずつ返す。

- 文字コード: BOM → UTF-8(厳密) → CP932 → shift_jis_2004。全件をデコードして確かめる
- 区切り文字: 「最も多い列数に一致する行の割合」が高い候補を採用する
- 引用符内の改行に対応するため csv.reader(newline='') に読ませ、行単位の分割はしない
"""
from __future__ import annotations

import codecs
import csv
import io
import re
import threading
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from tables.source import CellInfo, SheetInfo, SourceRow, UploadError, clean_text, value_kind

SNIFF_BYTES = 1024 * 1024
SNIFF_RECORDS = 200
DELIMITERS = [",", "\t", ";", "|"]
DELIMITER_NAMES = {",": "カンマ", "\t": "タブ", ";": "セミコロン", "|": "縦棒"}
REPLACEMENT_CHAR = "〓"
_CHUNK = 1024 * 1024
_EXCEL_WRAPPED = re.compile(r'^="(.*)"$', re.S)
_TRAILER_WORDS = re.compile(r"(合計|総計|件数|END|EOF|以上)", re.I)

csv.field_size_limit(16 * 1024 * 1024)

# 読めない文字を〓に置き換えるエラーハンドラ。置き換えた回数をスレッドごとに数える
_geta_state = threading.local()


def _geta_handler(err: UnicodeDecodeError):
    _geta_state.count = getattr(_geta_state, "count", 0) + 1
    return REPLACEMENT_CHAR, err.end


codecs.register_error("tables_geta", _geta_handler)


@dataclass
class CsvSniff:
    encoding: str
    bom: bool
    delimiter: str
    preamble_rows: int
    header_row: int
    trailer_rows: int
    confidence: float
    warnings: list[str] = field(default_factory=list)
    decode_error_line: int | None = None  # 全件デコードで失敗した物理行（おおよそ）


def sniff_csv(path) -> CsvSniff:
    path = Path(path)
    with path.open("rb") as f:
        head = f.read(SNIFF_BYTES)
    warnings: list[str] = []
    if not head:
        return CsvSniff("utf-8", False, ",", 0, 1, 0, 0.0, ["ファイルが空です"])

    encoding, bom, error_line = _detect_encoding(path, head)
    if error_line is not None:
        warnings.append(
            f"{error_line}行目付近に、文字コード {encoding} として読めない文字があります。"
            f"文字コードを選び直すか、「読めない文字を{REPLACEMENT_CHAR}に置き換える」を選んでください"
        )

    truncated = path.stat().st_size > len(head)
    text = _decode_head(head, encoding)
    delimiter, ratio, records = _detect_delimiter(text, truncated, path.suffix.lower() == ".tsv")
    if ratio == 0.0:
        warnings.append("区切り文字を判定できませんでした（列が1つだけです）。固定長テキストは対象外です")

    header_row = _guess_header_record(records)
    trailer_rows = _count_trailer_rows(path, encoding, delimiter, _modal_width(records))
    confidence = round(ratio * (1.0 if error_line is None else 0.6), 3)
    return CsvSniff(
        encoding=encoding,
        bom=bom,
        delimiter=delimiter,
        preamble_rows=max(0, header_row - 1),
        header_row=header_row,
        trailer_rows=trailer_rows,
        confidence=confidence,
        warnings=warnings,
        decode_error_line=error_line,
    )


def describe_sniff(sniff: CsvSniff) -> str:
    """画面表示用の要約。「文字コード: CP932 / 区切り: カンマ / 前置き行: 4行」"""
    enc = {"utf-8-sig": "UTF-8（BOM付き）", "utf-8": "UTF-8", "cp932": "CP932（Shift_JIS）",
           "shift_jis_2004": "Shift_JIS-2004", "utf-16": "UTF-16"}.get(sniff.encoding, sniff.encoding)
    name = DELIMITER_NAMES.get(sniff.delimiter, sniff.delimiter)
    return f"文字コード: {enc} / 区切り: {name} / 前置き行: {sniff.preamble_rows}行"


# ---- 文字コード ----

def _detect_encoding(path: Path, head: bytes) -> tuple[str, bool, int | None]:
    if head.startswith(codecs.BOM_UTF8):
        return "utf-8-sig", True, _full_decode_error_line(path, "utf-8-sig")
    if head.startswith(codecs.BOM_UTF16_LE) or head.startswith(codecs.BOM_UTF16_BE):
        return "utf-16", True, _full_decode_error_line(path, "utf-16")
    sample = head[:4096]
    if len(sample) >= 4:
        even_nul = sample[0::2].count(0) / max(1, len(sample[0::2]))
        odd_nul = sample[1::2].count(0) / max(1, len(sample[1::2]))
        if odd_nul > 0.3 and even_nul < 0.05:
            return "utf-16-le", False, _full_decode_error_line(path, "utf-16-le")
        if even_nul > 0.3 and odd_nul < 0.05:
            return "utf-16-be", False, _full_decode_error_line(path, "utf-16-be")

    best: tuple[int, str, int] | None = None  # (失敗位置, 文字コード, 行)
    for enc in ("utf-8", "cp932", "shift_jis_2004"):
        err = _full_decode_error(path, enc)
        if err is None:
            return enc, False, None
        pos, line = err
        if best is None or pos > best[0]:
            best = (pos, enc, line)
    assert best is not None
    return best[1], False, best[2]


def _full_decode_error(path: Path, encoding: str) -> tuple[int, int] | None:
    """全件を分割デコードし、失敗した (バイト位置, 物理行) を返す。成功なら None。"""
    decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
    consumed = 0
    newlines = 0
    with path.open("rb") as f:
        while True:
            chunk = f.read(_CHUNK)
            final = not chunk
            try:
                decoder.decode(chunk, final=final)
            except UnicodeDecodeError as e:
                # チャンク内の位置はデコーダのバッファ分ずれるが、行の目安には十分
                offset = max(0, min(len(chunk), e.start))
                return consumed + offset, newlines + chunk[:offset].count(b"\n") + 1
            if final:
                return None
            consumed += len(chunk)
            newlines += chunk.count(b"\n")


def _full_decode_error_line(path: Path, encoding: str) -> int | None:
    err = _full_decode_error(path, encoding)
    return None if err is None else err[1]


def _decode_head(head: bytes, encoding: str) -> str:
    decoder = codecs.getincrementaldecoder(encoding)(errors="tables_geta")
    return decoder.decode(head, final=False).replace("\x00", "")


# ---- 区切り文字と見出し ----

def _parse_records(text: str, delimiter: str, truncated: bool) -> list[list[str]]:
    records: list[list[str]] = []
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, strict=False)
    try:
        for rec in reader:
            records.append(rec)
            if len(records) > SNIFF_RECORDS:
                break
    except csv.Error:
        pass
    if truncated and len(records) <= SNIFF_RECORDS and records:
        records.pop()  # 途中で切れた最後のレコードは使わない
    return records[:SNIFF_RECORDS]


def _width(rec: list[str]) -> int:
    """末尾の空列を除いた列数。"""
    n = len(rec)
    while n and not rec[n - 1].strip():
        n -= 1
    return n


def _modal_width(records: list[list[str]]) -> int:
    widths = [len(r) for r in records if len(r) >= 2]
    if not widths:
        return 1
    return Counter(widths).most_common(1)[0][0]


def _detect_delimiter(text: str, truncated: bool, prefer_tab: bool) -> tuple[str, float, list[list[str]]]:
    best: tuple[tuple, str, list[list[str]]] | None = None
    for delim in DELIMITERS:
        records = _parse_records(text, delim, truncated)
        nonblank = [r for r in records if _width(r) > 0]
        if not nonblank:
            continue
        mode = _modal_width(nonblank)
        if mode < 2:
            ratio = 0.0
        else:
            ratio = sum(1 for r in nonblank if len(r) == mode) / len(nonblank)
        header_like = _header_likeness(nonblank, mode)
        rank = (round(ratio, 2), header_like, 1 if (prefer_tab and delim == "\t") else 0, mode)
        if best is None or rank > best[0]:
            best = (rank, delim, records)
    if best is None:
        return ("\t" if prefer_tab else ","), 0.0, []
    return best[1], best[0][0], best[2]


def _header_likeness(records: list[list[str]], mode: int) -> float:
    for rec in records[:30]:
        if len(rec) == mode:
            cells = [c.strip() for c in rec if c.strip()]
            if not cells:
                continue
            strings = sum(1 for c in cells if value_kind(c, c) in ("string", "code"))
            return round(strings / len(cells), 2)
    return 0.0


def _guess_header_record(records: list[list[str]]) -> int:
    """最頻の列数に一致し、次の行も同じ列数になる最初のレコード（1始まり）。"""
    mode = _modal_width(records)
    if mode < 2:
        return 1
    for i, rec in enumerate(records):
        if len(rec) != mode:
            continue
        cells = [c.strip() for c in rec if c.strip()]
        if len(cells) < max(2, mode * 0.5):
            continue
        strings = sum(1 for c in cells if value_kind(c, c) in ("string", "code"))
        if strings / len(cells) < 0.7:
            continue
        following = [r for r in records[i + 1:i + 4] if _width(r) > 0]
        if following and all(len(r) != mode for r in following):
            continue
        return i + 1
    return 1


def _count_trailer_rows(path: Path, encoding: str, delimiter: str, mode: int) -> int:
    """末尾の件数行・合計行（列数が合わない短い行）を数える。"""
    size = path.stat().st_size
    if size == 0 or mode < 2:
        return 0
    with path.open("rb") as f:
        f.seek(max(0, size - 65536))
        tail = f.read()
    text = tail.decode(encoding.replace("-sig", ""), errors="tables_geta").replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    if size > 65536:
        lines = lines[1:]  # 先頭は途中から始まるので捨てる
    count = 0
    for line in reversed(lines):
        if not line.strip():
            continue
        if '"' in line or count >= 5:
            break
        cells = next(csv.reader([line], delimiter=delimiter))
        if len(cells) < mode * 0.5 and (_width(cells) <= 3 or _TRAILER_WORDS.search(line)):
            count += 1
            continue
        break
    return count


# ---- 本体 ----

class CsvSource:
    kind = "csv"

    def __init__(self, path, file_name: str | None = None, options: dict | None = None):
        self.path = Path(path)
        self.file_name = file_name or self.path.name
        options = dict(options or {})
        self.sniff: CsvSniff | None = None
        if not options.get("encoding") or not options.get("delimiter"):
            self.sniff = sniff_csv(self.path)
        self.encoding: str = options.get("encoding") or self.sniff.encoding
        self.delimiter: str = options.get("delimiter") or self.sniff.delimiter
        errors = str(options.get("errors") or "strict")
        self.replace_errors = errors.startswith("replace")
        self.replaced_rows: list[int] = []  # 〓に置き換えたレコード番号（直近の rows() 走査分）
        self._stats: tuple[int, int] | None = None

    @property
    def warnings(self) -> list[str]:
        return list(self.sniff.warnings) if self.sniff else []

    def close(self) -> None:
        pass

    def sheets(self) -> list[SheetInfo]:
        if self._stats is None:
            max_row = max_col = 0
            for row in self.rows(self.file_name):
                max_row = row.index
                max_col = max(max_col, len(row.cells))
            self._stats = (max_row, max_col)
        return [SheetInfo(name=self.file_name, hidden=False, max_row=self._stats[0], max_col=self._stats[1], table_ranges=[])]

    def rows(self, sheet: str | None = None, start: int = 1, limit: int | None = None) -> Iterator[SourceRow]:
        errors = "tables_geta" if self.replace_errors else "strict"
        self.replaced_rows = []
        index = 0
        emitted = 0
        try:
            with self.path.open("r", encoding=self.encoding, errors=errors, newline="") as f:
                lines = (line.replace("\x00", "") for line in f)
                reader = csv.reader(lines, delimiter=self.delimiter, strict=False)
                for record in reader:
                    index += 1
                    if self.replace_errors and any(REPLACEMENT_CHAR in v for v in record):
                        self.replaced_rows.append(index)
                    if index < start:
                        continue
                    if limit is not None and emitted >= limit:
                        return
                    emitted += 1
                    yield SourceRow(index=index, cells=[_csv_cell(v) for v in record], hidden=None)
        except UnicodeDecodeError as e:
            raise UploadError(
                f"{index + 1}行目付近で、文字コード {self.encoding} として読めない文字がありました。"
                f"文字コードを選び直すか、「読めない文字を{REPLACEMENT_CHAR}に置き換える」を選んでください"
            ) from e


def _csv_cell(raw: str) -> CellInfo:
    text = clean_text(raw)
    m = _EXCEL_WRAPPED.match(text)
    if m:
        text = m.group(1).strip()
    return CellInfo(value=text if text else None, text=text)
