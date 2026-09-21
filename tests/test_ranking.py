"""帳票の種類の順位付け（② で先に選んでおく種類の決め方）。

まとめて置いたとき、先に選んでおく種類は「見つかった項目の割合」ではなく forms.match_score の点で決める。
割合だけで並べると、クリックで作った 4 項目の種類が 4/4 = 100% になり、33 項目中 32 項目が見つかった
本物の種類（97%）に勝ってしまう。画面で選び直せるので行き止まりではないが、間違った既定は
誰も見ずに受け入れる。ここでは、点の性質・ファイルをまとめた順位・画面の既定・samples/forms の
15 の版フォルダ全部で自分の様式が先に選ばれることを確かめる。
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import replace
from pathlib import Path

import pytest

from app import database as db
from app.forms import (
    FieldDef,
    PatternDef,
    PatternMatch,
    PHANTOM_FIELDS,
    SHEET_WORTH,
    SheetDef,
    load_workbook_info,
    match_pattern,
    match_score,
    rank_batch,
    rank_patterns,
    rows_to_pattern,
    suggest_rows,
)
from app.views import _batch_matches

SAMPLES = Path(__file__).resolve().parents[1] / "samples" / "forms"


# ---- 点（match_score）の性質 ----------------------------------------------------------------

def test_a_full_type_with_many_hits_beats_a_tiny_type_found_in_full():
    """依頼で測られた例: F1 は 20 項目全部、F2 は 33 項目中 32、4 項目の種類は 4 項目全部。
    割合では 小(100%) > F1(100%) > F2(97%) だが、点では F2 > F1 > 小 になる。"""
    assert match_score(32, 33) > match_score(20, 20) > match_score(4, 4)
    # 見つかった数が同じなら、項目の少ない（割合の高い）方が上（20/20 > 20/33）
    assert match_score(20, 20) > match_score(20, 33)
    # 21/60 のような「多いが半分も見つからない」種類は 20/20 に勝てない
    assert match_score(21, 60) < match_score(20, 20)
    # 4/4 は、本物の種類が 18/20 しか見つからなくても勝てない
    assert match_score(4, 4) < match_score(18, 20)


def test_score_is_share_with_phantom_missing_fields():
    """点 = (見つかった数 + シート名の一致 × SHEET_WORTH) ÷ (項目数 + PHANTOM_FIELDS)。
    分母に「見つからなかったことにする項目」を足すので、項目の少ない種類ほど 1 に届かない。"""
    assert match_score(4, 4) == pytest.approx(4 / (4 + PHANTOM_FIELDS))
    assert match_score(32, 33) == pytest.approx(32 / (33 + PHANTOM_FIELDS))
    assert match_score(0, 5) == 0.0 and match_score(0, 0) == 0.0     # 項目の無い種類は 0
    # 見つかった数・割合に対して単調
    assert match_score(10, 20) < match_score(11, 20) < match_score(12, 20)
    assert match_score(10, 12) > match_score(10, 20) > match_score(10, 40)


def test_a_matching_sheet_name_is_worth_a_few_fields_not_the_ranking():
    """シート名が同じなら SHEET_WORTH 個の項目が見つかったのと同じだけ足す。名前が違う版もあるので、
    シート名だけで小さな種類が本物を抜けるほどは重くしない。"""
    assert match_score(20, 20, 1.0) == pytest.approx(match_score(20 + SHEET_WORTH, 20, 0.0))
    assert match_score(20, 20, 0.6) > match_score(20, 20, 0.0)          # 部分一致も少し足す
    assert match_score(4, 4, 1.0) < match_score(32, 33, 0.0)             # シート名が合う 4/4 < 名前の違う 32/33
    assert match_score(20, 20, 1.0) < match_score(32, 33, 0.0)           # 依頼の例はシート名が違っても F2


# ---- ファイルをまとめた順位（rank_batch） --------------------------------------------------

def _pattern(pattern_id: int, name: str, n_fields: int) -> PatternDef:
    fields = [FieldDef(f"f{i}", f"項目{i}", [f"項目{i}"]) for i in range(n_fields)]
    return PatternDef(name=name, id=pattern_id, fields=fields, sheets=[SheetDef("報告書")])


def _match(pattern: PatternDef, found: int, sheet: float = 0.0, sheets=("報告書",)) -> PatternMatch:
    total = len(pattern.fields)
    confidence = round(100 * (0.25 * sheet + 0.75 * found / total))
    return PatternMatch(pattern, confidence, list(sheets), found, total,
                        sheet_score=sheet, score=match_score(found, total, sheet))


def test_rank_batch_puts_the_real_type_first_in_the_measured_scenario():
    """F2 の帳票フォルダを置く: F1(20 項目) は毎回 20/20、F2(33 項目) は 32/33、4 項目の種類は 4/4。"""
    f1, f2, tiny = _pattern(1, "F1", 20), _pattern(2, "F2", 33), _pattern(3, "小さい種類", 4)
    ranked = [{1: _match(f1, 20, 1.0), 2: _match(f2, 32, 1.0), 3: _match(tiny, 4, 1.0)} for _ in range(5)]
    order = rank_batch(ranked)
    assert [b.pattern.name for b in order] == ["F2", "F1", "小さい種類"]
    assert order[0].votes == 5 and order[1].votes == 0 and order[2].votes == 0
    assert order[0].found_min == order[0].found_max == 32 and order[0].total_fields == 33
    assert order[0].score == pytest.approx(match_score(32, 33, 1.0))


def test_rank_batch_counts_agreeing_files_before_the_mean_score():
    """3 件中 2 件で最も合った種類が先。残り 1 件で別の種類の点が飛び抜けていても引きずられない。"""
    a, b = _pattern(1, "A", 20), _pattern(2, "B", 20)
    ranked = [
        {1: _match(a, 15), 2: _match(b, 14)},
        {1: _match(a, 15), 2: _match(b, 14)},
        {1: _match(a, 2), 2: _match(b, 20, 1.0)},    # この 1 件だけ B が圧倒的
    ]
    order = rank_batch(ranked)
    assert [b_.pattern.name for b_ in order] == ["A", "B"]
    assert order[0].votes == 2 and order[1].votes == 1
    # 平均点では B の方が高い（票が同じならこちらで決まる）
    assert order[1].score > order[0].score


def test_rank_batch_breaks_equal_votes_by_mean_score_and_gives_ties_to_both():
    a, b = _pattern(1, "A", 20), _pattern(2, "B", 20)
    ranked = [
        {1: _match(a, 10), 2: _match(b, 10)},   # 同点 → 両方に票
        {1: _match(a, 12), 2: _match(b, 11)},
        {1: _match(a, 11), 2: _match(b, 12)},
    ]
    order = rank_batch(ranked)
    assert order[0].votes == order[1].votes == 2
    assert [b_.pattern.name for b_ in order] == ["A", "B"]       # 平均も同じなら登録順（先頭ファイルの並び）
    ranked[1][2] = _match(b, 14)                                   # B の平均が上回る
    assert [b_.pattern.name for b_ in rank_batch(ranked)] == ["B", "A"]


def test_rank_batch_collects_sheets_and_counts_across_files():
    a = _pattern(1, "A", 6)
    ranked = [{1: _match(a, 4, sheets=("報告書",))}, {1: _match(a, 6, sheets=("報告書", "別紙"))},
              {1: _match(a, 5, sheets=("別紙",))}]
    only, = rank_batch(ranked)
    assert (only.found_min, only.found_max, only.total_fields) == (4, 6, 6)
    assert only.sheet_names == ["報告書", "別紙"] and len(only.matches) == 3 and only.votes == 3
    assert rank_batch([]) == []


def test_batch_matches_label_still_says_what_was_found():
    """画面の「N項目中M〜K項目が見つかりました」は、順位の付け方を変えても実際の数を言う。"""
    f2, tiny = _pattern(2, "F2", 33), _pattern(3, "小さい種類", 4)
    ranked = [{2: _match(f2, 32), 3: _match(tiny, 4)}, {2: _match(f2, 31), 3: _match(tiny, 4)}]
    matches = _batch_matches(ranked)
    assert [m.pattern.name for m in matches] == ["F2", "小さい種類"]
    assert matches[0].found_label == "33項目中31〜32項目が見つかりました"
    assert matches[1].found_label == "4項目中4項目が見つかりました"      # 100% でも先頭にはならない
    assert matches[0].found_fields == 31 and matches[0].total_fields == 33   # 少なめの注意書きは最少の数で判定
    assert matches[0].votes == 2 and matches[0].files == 2


# ---- 実際のブックでの順位（rank_patterns） -------------------------------------------------

def _repair_types(repair_infos) -> tuple[PatternDef, PatternDef]:
    """見本から作った設備修理報告書の種類（全部）と、その先頭 4 項目だけの小さな種類。"""
    sheet_rows, field_rows = suggest_rows(repair_infos)
    meta = {"name": "設備修理報告書", "version": "v1"}
    full = rows_to_pattern(1, meta, sheet_rows, field_rows)
    used = [r for r in field_rows if r["use"]][:4]
    tiny = rows_to_pattern(2, {**meta, "name": "小さい種類"}, sheet_rows, used)
    assert len(full.fields) >= 8 and len(tiny.fields) == 4
    return full, tiny


def test_rank_patterns_puts_the_full_type_above_a_tiny_type_found_in_full(repair_infos):
    full, tiny = _repair_types(repair_infos)
    for info in repair_infos:
        ranked = rank_patterns(info, [tiny, full])
        assert ranked[0].pattern.name == "設備修理報告書", info.path.name
        small = ranked[1]
        assert small.found_fields == small.total_fields == 4          # 罠は本物: 小さな種類は 4/4
        assert small.confidence >= ranked[0].confidence               # 割合（confidence）なら小さな種類が勝っていた
        assert ranked[0].score > small.score


def test_match_pattern_reports_sheet_score_and_score(repair_infos):
    full, _ = _repair_types(repair_infos)
    m = match_pattern(repair_infos[0], full)
    assert m.sheet_score == 1.0 and m.sheet_names == ["修理報告書"]
    assert m.score == pytest.approx(match_score(m.found_fields, m.total_fields, m.sheet_score))
    assert 0 < m.score < 1


# ---- ② の画面（/forms/type）で先に選ばれる種類 ------------------------------------------------

def _register(app, pattern: PatternDef) -> int:
    """種類を使用中で登録する（帳票登録の［使用開始］と同じ状態）。"""
    with app.app_context():
        pattern_id = db.create_pattern(pattern.name)
        db.save_pattern(replace(pattern, id=pattern_id), "active")
    return pattern_id


def test_the_type_step_prechecks_the_full_type_not_the_tiny_one(app, client, repair_infos, sample_dir, tmp_path):
    """同じ帳票を 3 件まとめて置く。4 項目の種類（4/4）が使用中でも、先に選ばれるのは見本から作った種類。
    それぞれの種類の「N項目中M項目が見つかりました」とファイルごとの数は、実際の数のまま。"""
    full, tiny = _repair_types(repair_infos)
    tiny_id = _register(app, tiny)      # 先に登録して、名前順・id 順のどちらでも先頭に来るようにする
    full_id = _register(app, full)

    data = (sample_dir / "standard.xlsx").read_bytes()
    parts = [(io.BytesIO(data), f"修理報告書_{i}.xlsx") for i in range(1, 4)]
    res = client.post("/forms/upload", data={"file": parts}, content_type="multipart/form-data")
    assert res.status_code == 200, res.get_data(as_text=True)
    ids = [d["id"] for d in res.get_json()["docs"]]

    html = client.get("/forms/type", query_string={"ids": ",".join(map(str, ids))}).get_json()["html"]
    checked = re.search(r'name="pattern_id" value="(\d+)" checked', html)
    assert checked and int(checked.group(1)) == full_id
    radios = [int(v) for v in re.findall(r'name="pattern_id" value="(\d+)"', html)]
    assert radios[0] == full_id and tiny_id in radios          # 並びも点の順
    assert "4項目中4項目が見つかりました" in html                 # 小さな種類の数は正直に出す
    full_label = re.search(r"(\d+)項目中(\d+)項目が見つかりました", html.split("設備修理報告書", 1)[1])
    assert full_label and int(full_label.group(1)) == len(full.fields) and int(full_label.group(2)) > 4
    # ファイルごとの数（✕ で外すときの目安）は種類ごとに残る
    counts = json.loads(re.search(r"data-file-counts='([^']+)'", html).group(1))
    assert set(counts) == {str(full_id), str(tiny_id)}
    assert all(counts[str(tiny_id)][str(i)] == 4 for i in ids)
    assert all(counts[str(full_id)][str(i)] > 4 for i in ids)


# ---- samples/forms の 15 の版フォルダ -----------------------------------------------------------

def _sample_types_and_folders():
    """F1〜F5 をそれぞれクリックで種類にし（scripts.samples.evaluate_forms と同じ作り方）、
    さらに F1 の先頭 4 項目だけの小さな種類を足す。戻り値: (種類, {版フォルダ: [WorkbookInfo]})。"""
    from scripts.samples.evaluate_forms import click_rows, pick_samples

    families = sorted(d for d in SAMPLES.iterdir() if (d / "_expected.jsonl").exists()) if SAMPLES.is_dir() else []
    if len(families) < 5:
        pytest.skip("samples/forms がありません")
    patterns, folders, tiny_rows = [], {}, None
    for i, family in enumerate(families, start=1):
        entries = [json.loads(l) for l in (family / "_expected.jsonl").open(encoding="utf-8") if l.strip()]
        infos: dict[str, object] = {}
        by_version: dict[str, list[dict]] = {}
        for e in entries:
            by_version.setdefault(e["file"].split("/")[0], []).append(e)
        picks = pick_samples(entries, 3) + [e for v in by_version.values() for e in v[:3]]   # 版ごとに 3 件見る
        for e in picks:
            infos.setdefault(e["file"], load_workbook_info(family / e["file"]))
        sheet_rows, field_rows = click_rows([infos[e["file"]] for e in pick_samples(entries, 3)], family.name)
        patterns.append(rows_to_pattern(i, {"name": family.name, "version": "eval"}, sheet_rows, field_rows))
        if tiny_rows is None:
            tiny_rows = (sheet_rows, field_rows[:4])
        for version, es in by_version.items():
            folders[f"{family.name}/{version}"] = [infos[e["file"]] for e in es[:3]]
    patterns.append(rows_to_pattern(99, {"name": "小さい種類", "version": "eval"}, *tiny_rows))
    return patterns, folders


@pytest.mark.samples
def test_every_sample_version_folder_prechecks_its_own_family():
    """15 の版フォルダのどれを置いても、先に選ばれるのはその様式の種類（4 項目の種類は使用中でも先頭にならない）。
    見つかった項目の割合の平均で並べていた頃は、この小さな種類が半分近くのフォルダで先頭だった。"""
    patterns, folders = _sample_types_and_folders()
    assert len(folders) == 15
    tiny_perfect, old_rule_wrong = 0, 0
    for name, infos in folders.items():
        family = name.split("/")[0]
        ranked = [{m.pattern.id: m for m in rank_patterns(info, patterns)} for info in infos]
        matches = _batch_matches(ranked)
        assert matches[0].pattern.name == family, f"{name}: {[(m.pattern.name, m.found_label) for m in matches[:3]]}"
        tiny = next(m for m in matches if m.pattern.name == "小さい種類")
        tiny_perfect += tiny.found_fields == tiny.total_fields == 4
        by_share = max(matches, key=lambda m: sum(r[m.pattern.id].confidence for r in ranked))
        old_rule_wrong += by_share.pattern.name != family
    assert tiny_perfect >= 5, "小さな種類が 4/4 になるフォルダが無く、罠を試せていない"
    assert old_rule_wrong >= 1, "割合の平均でも全部正しく、この検査では違いが出ない"
