"""管理用CSV と、ダウンロード用 zip（RAG投入用/ と 管理用_RAGには入れない/）。

投入済みとの差分管理はしない（取り込みごとに、その取り込みの全 md を zip にまとめて渡す）。
"""
from __future__ import annotations

import csv
import io
import zipfile

RAG_DIR = "RAG投入用"
ADMIN_DIR = "管理用_RAGには入れない"
_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
_FORMULA_PREFIX = ("=", "+", "-", "@", "\t", "\r")


# ---- 管理用CSV -----------------------------------------------------------------------------

def guard_formula(value) -> str:
    """Excel で開いたときに式として実行されないよう、= + - @ で始まる文字列の先頭に ' を付ける。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return str(value)
    s = str(value)
    if s.startswith(_FORMULA_PREFIX):
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


def normalized_csv(spec, records: list[dict]) -> bytes:
    """正規化データ.csv: 記録キー、列（表示名）、原本ファイル・シート・行。"""
    columns = [(c.key, c.display) for c in spec.columns]
    header = ["記録キー"] + [d for _k, d in columns] + ["原本ファイル", "シート", "行"]
    rows = []
    for rec in records:
        values = rec.get("values", {}) or {}
        source = rec.get("source", {}) or {}
        rows.append([rec.get("key", "")] + [values.get(k) for k, _d in columns]
                    + [source.get("file", ""), source.get("sheet", ""), source.get("row", "")])
    return csv_bytes(header, rows)


def issues_csv(issues) -> bytes:
    from tables.checks import LEVEL_LABELS

    rows = []
    for issue in issues:
        d = issue.to_dict() if hasattr(issue, "to_dict") else dict(issue)
        rows.append([LEVEL_LABELS.get(d.get("level"), d.get("level")), d.get("code", ""), d.get("row") or "",
                     d.get("column") or "", d.get("message", "")])
    return csv_bytes(["種類", "コード", "行", "列", "内容"], rows)


def report_csv(items: list[tuple[str, object]]) -> bytes:
    return csv_bytes(["項目", "値"], [[k, v] for k, v in items])


def import_report_items(imp: dict, spec, file_count: int) -> list[tuple]:
    """取込レポート.csv の行（件数、除外行と理由、照合結果、エラー・警告の数）。"""
    stats = imp.get("stats") or {}
    counts = stats.get("issue_counts") or {}
    items: list[tuple] = [
        ("取り込みID", imp["id"]), ("ファイル名", imp["file_name"]), ("シート", stats.get("sheet", "")),
        ("取り込み設定", spec.name), ("確定日時", imp.get("confirmed_at") or ""),
        ("読み込んだ行", stats.get("rows_scanned")), ("記録件数", stats.get("records")),
        ("継続行として連結", stats.get("continuation_merged")), ("小計・合計の行", stats.get("subtotal_rows")),
    ]
    for reason, n in sorted((stats.get("excluded") or {}).items()):
        items.append((f"除外: {reason}", n))
    for rc in stats.get("reconcile") or []:
        ok = abs(float(rc.get("expected") or 0) - float(rc.get("actual") or 0)) < 1e-6
        items.append((f"照合: {rc.get('label')}（{rc.get('column')}、{rc.get('row')}行目）",
                      f"{'一致' if ok else '不一致'} 表の値={rc.get('expected')} 行の合計={rc.get('actual')}"))
    items += [("エラー", counts.get("error", 0)), ("警告", counts.get("warning", 0)), ("Markdownファイル数", file_count)]
    return items


# ---- zip --------------------------------------------------------------------------------

def _zip_write(zf: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=_ZIP_TIME)  # 時刻を固定して、同じ内容なら同じ zip にする
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    zf.writestr(info, data)


def build_zip(md_files: list[tuple[str, bytes]], extras: dict[str, bytes] | None = None) -> bytes:
    """RAG投入用/ に全 md、管理用_RAGには入れない/ に正規化データ・問題一覧・取込レポートなど。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in sorted(md_files):
            _zip_write(zf, f"{RAG_DIR}/{name}", data)
        for name, data in sorted((extras or {}).items()):
            _zip_write(zf, f"{ADMIN_DIR}/{name}", data)
    return buf.getvalue()
