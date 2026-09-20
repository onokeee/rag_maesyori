"""取り込み結果のチェック（確定を止めるエラーと警告）。

- エラー: 必須の見出しがない、型エラーの割合が上限超え
- 警告: 型エラー、許可値にない区分、年を補った日付、除外した行、日付が空の行、
        読めない文字の置き換え、使っていない列、値が空の必須列、合計の照合
行ごとの型エラーなどは normalize.read_records が Issue にし、ここでは全体を見たチェックを足す。
"""
from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass

from tables.csv_source import MAX_RECORD_LINES

LEVEL_LABELS = {"error": "エラー", "warning": "警告"}


@dataclass
class Issue:
    level: str  # error/warning
    code: str
    message: str
    row: int | None = None
    column: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


def has_blocking(issues) -> bool:
    return any(_level(i) == "error" for i in issues)


def count_levels(issues) -> dict[str, int]:
    c = Counter(_level(i) for i in issues)
    return {"error": c.get("error", 0), "warning": c.get("warning", 0)}


def _level(issue) -> str:
    return issue.level if isinstance(issue, Issue) else str((issue or {}).get("level"))


def _get(obj, name, default=None):
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _values(record) -> dict:
    return _get(record, "values", {}) or {}


def run_checks(records, spec, stats) -> list[Issue]:
    """全体を見たチェック。stats は ImportStats か、その dict。"""
    st = stats if isinstance(stats, dict) else stats.to_dict()
    issues: list[Issue] = []

    for display in st.get("missing_required") or []:
        issues.append(Issue("error", "required_missing", f"必須の列「{display}」の見出しが見つかりません"))

    rates = (spec.checks or {}).get("type_error_rate") or {}
    warn, block = float(rates.get("warn", 0.02)), float(rates.get("block", 0.10))
    checked = st.get("type_checked") or {}
    for key, errors in sorted((st.get("type_errors") or {}).items()):
        total = checked.get(key) or 0
        if not total or not errors:
            continue
        rate = errors / total
        col = spec.column(key)
        label = col.display if col else key
        if rate > block:
            issues.append(Issue("error", "type_error_rate",
                                f"列「{label}」で値を変換できない行が多すぎます（{errors}/{total}件、{rate:.0%}）。"
                                "列の型か見出しの位置を確認してください", column=label))
        elif rate >= warn:
            issues.append(Issue("warning", "type_error_rate",
                                f"列「{label}」で値を変換できない行があります（{errors}/{total}件、{rate:.1%}）", column=label))

    dups = st.get("duplicate_keys") or {}
    if dups:
        sample = "、".join(f"{k}（{n}件）" for k, n in list(sorted(dups.items()))[:5])
        issues.append(Issue("warning", "duplicate_key",
                            f"記録キーが同じ行が{len(dups)}種類あります: {sample}。出現順の番号で区別します"))

    date_key = spec.date_key
    if date_key:
        undated = sum(1 for rec in records if not _values(rec).get(date_key))
        if undated:
            issues.append(Issue("warning", "undated", f"日付が空の行が{undated}件あります（「日付なし」のファイルに入れます）"))
        issues += _date_outliers(records, date_key)

    # 許可値
    for col in spec.columns:
        if col.allowed:
            allowed = set(col.allowed)
            bad = Counter(str(_values(r).get(col.key)) for r in records
                          if _values(r).get(col.key) not in (None, "") and str(_values(r).get(col.key)) not in allowed)
            if bad:
                detail = "、".join(f"{v}（{n}件）" for v, n in bad.most_common(8))
                issues.append(Issue("warning", "not_allowed",
                                    f"列「{col.display}」に選択肢にない値があります: {detail}", column=col.display))
        if col.required and col.key in (st.get("positions") or {}):
            empty = sum(1 for r in records if _values(r).get(col.key) in (None, ""))
            if empty:
                issues.append(Issue("warning", "required_empty",
                                    f"必須の列「{col.display}」が空の行が{empty}件あります", column=col.display))
    if st.get("year_inferred"):
        issues.append(Issue("warning", "year_inferred",
                            f"年のない日付に、年を補った行が{st['year_inferred']}件あります（{st.get('year_context_label') or '年度'}から）"))
    if st.get("year_missing"):
        issues.append(Issue("warning", "year_missing",
                            f"年のない日付で、年を補えなかった行が{st['year_missing']}件あります。"
                            "タイトル・シート名・ファイル名に年度がないためです"))
    for reason, n in sorted((st.get("excluded") or {}).items()):
        if reason in ("非表示の行", "取り消し線の行"):
            issues.append(Issue("warning", "excluded_rows", f"{reason}を{n}行除外しました"))
    if st.get("included_hidden"):
        issues.append(Issue("warning", "included_rows", f"非表示・取り消し線の行を{st['included_hidden']}行取り込みました"))
    if st.get("error_values"):
        issues.append(Issue("warning", "excel_error", f"Excelのエラー値（#N/A など）を空欄として扱ったセルが{st['error_values']}個あります"))
    if st.get("uncached_formulas"):
        by_col = st["uncached_formulas"]
        names = "、".join(f"「{k}」" for k in list(by_col)[:5])
        issues.append(Issue("warning", "uncached_formula",
                            f"Excelで計算されていない数式のセルが{sum(by_col.values())}個あります（列{names}）。"
                            "値が空欄として読まれています。Excelで開いて保存し直してから取り込んでください"))
    if st.get("unclosed_quote_row"):
        row = st["unclosed_quote_row"]
        issues.append(Issue("error", "unclosed_quote",
                            f"{row}行目の \" が閉じていないため、以降の行が1つの値になっています。"
                            "元のファイルの \" を直してから、もう一度取り込んでください", row=row))
    if st.get("long_record_row"):
        row = st["long_record_row"]
        issues.append(Issue("warning", "long_record",
                            f"{row}行目の値が{MAX_RECORD_LINES}行以上あります。\" の閉じ忘れでないか確認してください", row=row))
    if st.get("replaced_rows"):
        rows = st["replaced_rows"]
        issues.append(Issue("warning", "replaced_chars",
                            f"読めない文字を〓に置き換えた行が{len(rows)}行あります（{', '.join(str(r) for r in rows[:10])}行目など）"))
    if st.get("unused_headers"):
        names = "、".join(st["unused_headers"][:10])
        issues.append(Issue("warning", "unused_headers", f"設定にない列があります（取り込みません）: {names}"))
    for missing in st.get("missing_optional") or []:
        issues.append(Issue("warning", "column_missing", f"列「{missing}」の見出しが見つかりません（空欄として扱います）"))
    tolerance = float((spec.checks or {}).get("reconcile_tolerance") or 0)
    for rc in st.get("reconcile") or []:
        diff = abs(float(rc.get("expected") or 0) - float(rc.get("actual") or 0))
        if diff > tolerance + 1e-6:
            issues.append(Issue("warning", "reconcile",
                                f"{rc.get('label') or '合計'}が合いません（列「{rc.get('column')}」: 表の値 {_num(rc.get('expected'))}、"
                                f"行の合計 {_num(rc.get('actual'))}）", row=rc.get("row"), column=rc.get("column")))
    if not records and not any(i.level == "error" for i in issues):
        # 記録が1件も無いまま確定すると、中身の無い zip を渡してしまう。確定を止めて理由も出す
        excluded = sum((st.get("excluded") or {}).values())
        extra = f"（取り消し線・非表示などで{excluded}行を除外しました）" if excluded else ""
        issues.append(Issue("error", "no_records",
                            f"取り込める行がありません{extra}。表の範囲か元のファイルを見直してください"))
    return issues


DATE_OUTLIER_YEARS = 5  # 記録の年の中央値からこれより離れた日付は、打ち間違いの疑い


def _date_outliers(records, date_key: str) -> list[Issue]:
    """ほかの記録から何年も離れた日付（2025年のデータに 2052年 など）。記録ファイルは年月ごとに分かれるので、
    離れた月に1件だけの記録ファイルができる原因になる。"""
    dated = []
    for rec in records:
        s = str(_values(rec).get(date_key) or "")
        if len(s) >= 4 and s[:4].isdigit():
            dated.append((int(s[:4]), s[:10], (_get(rec, "source") or {}).get("row")))
    if len(dated) < 3:
        return []
    years = sorted(y for y, _d, _r in dated)
    median = years[len(years) // 2]
    far = [(d, r) for y, d, r in dated if abs(y - median) > DATE_OUTLIER_YEARS]
    if not far:
        return []
    sample = "、".join(f"{r}行目（{d}）" if r else d for d, r in far[:5]) + ("など" if len(far) > 5 else "")
    return [Issue("warning", "date_outlier",
                  f"ほかの記録から{DATE_OUTLIER_YEARS}年より離れた日付が{len(far)}件あります: {sample}。"
                  "年の打ち間違いでないか確認してください", row=far[0][1])]


def _num(value) -> str:
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    return f"{int(f):,}" if f.is_integer() else f"{f:,.2f}"
