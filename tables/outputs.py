"""画面で見る問題一覧の CSV と、ダウンロード用 zip（RAG に入れる md だけ）。

投入済みとの差分管理はしない（取り込みごとに、その取り込みの全 md を zip にまとめて渡す）。
"""
from __future__ import annotations

import csv
import io
import unicodedata
import zipfile

_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
_FORMULA_PREFIX = ("=", "+", "-", "@", "\t", "\r")


# ---- 問題一覧の CSV（確認画面から見るためのもの。zip には入れない） ---------------------------------------

def guard_formula(value) -> str:
    """Excel で開いたときに式として実行されないよう、= + - @（全角も）で始まる文字列の先頭に ' を付ける。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return str(value)
    s = str(value)
    # 日本語の Excel は全角の ＝ ＋ － ＠ で始まる値も式として読むので、先頭の1文字は NFKC で比べる
    if s.startswith(_FORMULA_PREFIX) or unicodedata.normalize("NFKC", s[:1]).startswith(_FORMULA_PREFIX):
        return "'" + s
    return s


def csv_bytes(header: list[str], rows) -> bytes:
    """UTF-8（BOM付き）・CRLF の CSV。全セルに式の対策をする。"""
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow([guard_formula(h) for h in header])
    for row in rows:
        writer.writerow([guard_formula(v) for v in row])
    return ("﻿" + buf.getvalue()).encode("utf-8")


def issues_csv(issues) -> bytes:
    from tables.checks import LEVEL_LABELS

    rows = []
    for issue in issues:
        d = issue.to_dict() if hasattr(issue, "to_dict") else dict(issue)
        rows.append([LEVEL_LABELS.get(d.get("level"), d.get("level")), d.get("code", ""), d.get("row") or "",
                     d.get("column") or "", d.get("message", "")])
    return csv_bytes(["種類", "コード", "行", "列", "内容"], rows)


# ---- zip --------------------------------------------------------------------------------

def _zip_write(zf: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=_ZIP_TIME)  # 時刻を固定して、同じ内容なら同じ zip にする
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    zf.writestr(info, data)


def build_zip(md_files: list[tuple[str, bytes]]) -> bytes:
    """RAG に入れる md だけをフォルダ分けせずに入れる（zip を開いて、そのまま LightRAG にドラッグできるように）。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in sorted(md_files):
            _zip_write(zf, name, data)
    return buf.getvalue()
