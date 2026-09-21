"""logproc（「経過の記録」の列のルール処理）の単体テスト。

実例セルは docs/research/対応内容セル分析.json の realistic_examples を使う。
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

from logproc import (
    PeopleIndex, SplitOptions, apply_glossary, glossary_hits, mask_text, parse_log, render_timeline, review_notes,
)
from logproc.dates import parse_when_at, resolve_whens
from logproc.extract import extract_identifiers, extract_plans, extract_quantities
from logproc.text import shadow

CELLS_JSON = Path(__file__).resolve().parents[1] / "docs" / "research" / "対応内容セル分析.json"


@pytest.fixture(scope="module")
def cells() -> list[str]:
    if not CELLS_JSON.exists():
        pytest.skip("対応内容セル分析.json がありません")
    return json.loads(CELLS_JSON.read_text(encoding="utf-8"))["realistic_examples"]


@pytest.fixture(scope="module")
def people() -> PeopleIndex:
    return PeopleIndex(
        [
            {"name": "高橋 圭太", "aliases": ["K.T", "TK"], "org": "保全"},
            {"name": "佐藤 誠", "aliases": ["佐藤(保)"], "org": "保全"},
            {"name": "佐藤 美咲", "aliases": ["佐藤(製)"], "org": "製造"},
        ],
        column_names=["佐藤", "田中/山本"],
    )


def _dates(parse):
    return [s.when.date if s.when else None for s in parse.segments]


def _authors(parse):
    return [(s.author.name if s.author else None) for s in parse.segments]


# ---- 実例セル全体 ----

def test_all_realistic_cells_parse_deterministically(cells, people):
    for text in cells:
        a = parse_log(text, date(2024, 4, 1), people)
        b = parse_log(text, date(2024, 4, 1), people)
        assert a == b
        assert a.kind in ("log", "single", "header_cell")
        ids = [s.id for s in a.segments]
        assert ids == [f"s{i}" for i in range(1, len(ids) + 1)]
        for s in a.segments:
            # 位置は分割に使ったテキスト上で原文と一致し、重ならない
            assert a.text[s.start:s.end] == s.raw
            assert s.raw.strip() == s.raw
        for x, y in zip(a.segments, a.segments[1:]):
            assert x.end <= y.start
        render_timeline(a, "設備")
        review_notes(a)


def test_representative_cell_412(cells, people):
    p = parse_log(cells[0], date(2024, 4, 1), people)
    assert p.kind == "log" and p.order == "asc"
    assert len(p.segments) == 6
    assert _dates(p) == ["2024-04-01", "2024-04-01", "2024-04-02", "2024-04-03", "2024-04-08", "2024-04-19"]
    s1, s2, s3, s4, s5, s6 = p.segments
    assert (s1.when.time, s1.author.name) == ("10:00", "田中")
    assert s1.body.startswith("ライン停止の連絡あり")
    assert s1.identifiers == ["ALM-2031"]
    # 時刻だけの行：日付は直前から、記入者は（推定）
    assert s2.when.time == "10:20" and s2.when.how == "inherited" and not s2.when.estimated
    assert s2.author.name == "田中" and s2.author.estimated
    # 【4/2 夜勤】と行末の（K.T）＝人物一覧の別名
    assert s3.when.shift == "夜勤"
    assert s3.author.name == "高橋 圭太" and s3.author.raw == "K.T"
    assert s3.body == "同エラー再発2回、都度リセット"
    assert s3.quantities == ["2回"]
    # 同姓2人は特定しない
    assert s4.author.name == "佐藤" and "2人" in s4.author.note
    assert s4.identifiers == ["RB-ENC-05M"]
    assert "×1" in s4.quantities and s4.plans == ["納期1週間"]
    # 翌週は範囲＋推定
    assert (s5.when.date, s5.when.date_to, s5.when.estimated) == ("2024-04-08", "2024-04-14", True)
    assert s6.body == "再発なし→クローズ"

    lines = render_timeline(p, "搬送ロボット2号機", types={"s1": ["連絡", "初動"]})
    assert lines[0] == ("1. 2024-04-01 10:00［連絡・初動］田中｜搬送ロボット2号機: "
                        "ライン停止の連絡あり（製造 大野さん）。現場確認、搬送ロボのアーム原点復帰エラー（ALM-2031）。")
    assert lines[1] == "2. 2024-04-01 10:20 田中（推定）｜搬送ロボット2号機: 原点復帰→再起動で復旧。様子見。"
    assert lines[2].startswith("3. 2024-04-02 夜勤 高橋 圭太（K.T）｜搬送ロボット2号機: ")
    assert lines[4].startswith("5. 2024-04-08〜2024-04-14（原文「翌週」、推定） 佐藤｜")
    notes = review_notes(p)
    assert "5の日付は原文「翌週」から推定した（基準は直前の記録の2024-04-03）。" in notes
    assert "2の記入者は原文になく、直前の田中を引き継いだ。" in notes
    assert any("佐藤 誠／佐藤 美咲" in n for n in notes)


def test_glossary_in_render(cells, people):
    p = parse_log(cells[0], date(2024, 4, 1), people)
    lines = render_timeline(p, "搬送ロボット2号機", glossary={"様子見": "経過観察"})
    assert lines[1].endswith("原点復帰→再起動で復旧。経過観察。")


def test_slash_joined_single_line(cells, people):
    p = parse_log(cells[1], None, people)
    assert len(p.segments) == 4
    assert _dates(p) == ["2024-04-01", "2024-04-01", "2024-04-02", "2024-04-08"]
    assert _authors(p) == ["山本", "山本", "鈴木", "山本"]
    s3 = p.segments[2]
    assert s3.body == "部品手配 6205ZZ ×2 納期4/8"
    assert s3.identifiers == ["6205ZZ"] and s3.plans == ["納期4/8"]


def test_wareki_and_group_author(cells, people):
    p = parse_log(cells[2], None, people)
    assert _dates(p) == ["2024-04-01", "2024-04-03", "2024-04-10"]
    assert p.segments[0].author.name == "高橋 圭太"
    assert p.segments[0].quantities == ["0.35MPa", "0.22MPa", "0.33MPa"]
    assert p.segments[2].author.estimated


def test_bracket_shift_cells(cells, people):
    p = parse_log(cells[3], date(2024, 4, 2), people)
    assert [s.when.shift for s in p.segments] == ["夜勤", "日勤", "夜勤", "日勤", None]
    assert _dates(p)[-1] == "2024-04-11"
    assert p.segments[0].author is None
    assert "5〜6回" in p.segments[0].quantities


def test_relative_days(cells, people):
    p = parse_log(cells[4], date(2024, 4, 1), people)
    assert p.order == "asc"
    w = [s.when for s in p.segments]
    assert (w[0].date, w[0].shift) == ("2024-04-01", "夜")
    assert (w[1].date, w[1].estimated) == ("2024-04-02", True)
    assert (w[2].date, w[2].time, w[2].estimated) == ("2024-04-02", "15:00", False)   # 同日は推定にしない
    assert (w[3].date, w[3].estimated) == ("2024-04-05", True)
    assert p.segments[0].body.startswith("夜勤班長より連絡")   # 「夜勤班長」は勤務帯にしない


def test_handoff_arrow(cells, people):
    p = parse_log(cells[5], date(2024, 4, 1), people)
    assert p.kind == "single"
    a = p.segments[0].author
    assert a.name == "田中" and a.raw == "田中→佐藤"
    assert p.segments[0].when.how == "base" and p.segments[0].when.estimated


def test_email_block_is_one_segment(cells, people):
    p = parse_log(cells[6], date(2024, 4, 8), people)
    assert len(p.segments) == 3
    email = p.segments[1]
    assert "email" in email.marks
    assert email.raw.startswith("-----Original Message-----") and email.raw.endswith("-----")
    assert "4/10に弊社FEが伺います" in email.raw
    assert _dates(p) == ["2024-04-08", "2024-04-08", "2024-04-09"]
    assert any("メール" in w for w in p.warnings)


def test_no_anchor_paragraph_sentence_split(cells, people):
    text = cells[7]
    p = parse_log(text, date(2024, 4, 1), people)
    assert p.kind == "single"   # 120字以下は分けない
    p2 = parse_log(text, date(2024, 4, 1), people, SplitOptions(sentence_split_min_chars=40))
    assert len(p2.segments) >= 4
    assert all("sentence_split" in s.marks for s in p2.segments)
    assert "".join(s.raw for s in p2.segments) == text.replace("\n", "")
    assert p2.segments[0].raw.startswith("朝一で立ち上げ時にエラー。リセットで復帰。")   # 短い文は前につなぐ
    long_text = text + "工事は来月の連休に実施予定で、それまでは電源電圧を毎日測定して記録する運用とした。"
    p3 = parse_log(long_text, date(2024, 4, 1), people)
    assert len(p3.segments) > 1 and "工事は来月の連休に実施予定" in p3.segments[-1].raw


def test_desc_order_same_year(cells, people):
    p = parse_log(cells[8], date(2024, 5, 9), people)
    assert p.order == "desc"
    assert _dates(p) == ["2024-05-20", "2024-05-13", "2024-05-12", "2024-05-10", "2024-05-09"]
    lines = render_timeline(p, "クランプ")
    assert lines[0].startswith("1. 2024-05-09 渡辺｜クランプ: クランプ動作遅い")
    assert lines[-1].startswith("5. 2024-05-20 木村｜クランプ: クローズ")
    assert p.segments[2].identifiers == ["CDQ2B32-50D"]
    assert p.segments[3].plans == ["納期5/12"]


def test_desc_order_across_year_end(people):
    text = "1/7 木村 漏れなし確認 完了\n1/6 木村 オイルシール交換\n12/29 木村 応急でシール剤塗布\n12/28 木村 油漏れ発見"
    p = parse_log(text, date(2024, 12, 28), people)
    assert p.order == "desc"
    assert _dates(p) == ["2025-01-07", "2025-01-06", "2024-12-29", "2024-12-28"]
    assert render_timeline(p, "")[0].startswith("1. 2024-12-28 木村: 油漏れ発見")


def test_asc_across_year_end(cells, people):
    p = parse_log(cells[20], date(2024, 12, 28), people)
    assert p.order == "asc"
    assert _dates(p) == ["2024-12-28", "2024-12-29", "2025-01-06", "2025-01-07"]


def test_desc_is_not_pushed_to_next_year(people):
    # 新しい順の 5/20→5/9 を翌年にしない（発生日なし・セル内に年あり）
    text = "5/20 木村 クローズ\n5/13 木村 異常なし\n2024/5/9 渡辺 クランプ動作遅い"
    p = parse_log(text, None, people)
    assert p.order == "desc"
    assert _dates(p) == ["2024-05-20", "2024-05-13", "2024-05-09"]


def test_initials_unregistered_and_registered(cells, people):
    p = parse_log(cells[9], date(2024, 4, 15), people)
    assert _authors(p) == ["高橋 圭太", "高橋 圭太", None, None]
    assert p.segments[2].author.raw == "M.S"
    assert p.segments[0].quantities == ["5.0E-3Pa", "1.0E-4Pa"]
    assert any("M.S" in n for n in review_notes(p))


def test_bullets_and_trailing_paren_author(cells, people):
    p = parse_log(cells[10], date(2024, 4, 3), people)
    assert len(p.segments) == 4
    assert all("bullet" in s.marks for s in p.segments)
    assert _authors(p) == ["小林", "小林", "小林", "加藤"]
    assert p.segments[0].body == "設備停止、非常停止ボタン押下されたまま"


def test_no_base_date_year_unknown(cells, people):
    p = parse_log(cells[11], None, people)
    assert _dates(p) == [None, None, None, None]
    assert any("年を決められない" in w for w in p.warnings)
    assert render_timeline(p, "")[0].startswith("1. 6/3（年不明） 松本: ")
    assert p.segments[2].quantities == ["2.2kW", "4P"]


def test_note_line_joins_previous(cells, people):
    p = parse_log(cells[14], date(2024, 10, 2), people)
    assert len(p.segments) == 3
    last = p.segments[-1]
    assert "note" in last.marks and last.body.endswith("※暫定で手動運転にて生産継続中")
    assert "¥385,000" in p.segments[1].quantities and "納期6週間" in p.segments[1].plans


def test_packed_times_split_on_fullwidth_space(cells, people):
    p = parse_log(cells[15], date(2024, 4, 1), people)
    assert [s.when.time for s in p.segments] == ["10:05", "10:15", "10:40", "11:00"]
    assert [s.body for s in p.segments] == ["停止", "現場着", "ﾁｪｰﾝ外れ　復旧", "生産再開"]


def test_short_single(cells, people):
    p = parse_log(cells[16], date(2024, 4, 1), people)
    assert p.kind == "single" and p.segments[0].body == "再起動で復旧"


def test_group_authors_and_identifiers(cells, people):
    p = parse_log(cells[17], date(2024, 4, 22), people)
    assert _authors(p) == ["保全G", "保全G", "メーカーFE", "保全G"]
    assert p.segments[0].identifiers == ["OC2"]
    assert p.segments[2].identifiers == ["FR-A840-7.5K"]
    assert p.segments[3].quantities == ["3時間"]


def test_arrow_chain_split(cells, people):
    p = parse_log(cells[18], date(2024, 4, 1), people)
    assert [s.body for s in p.segments] == ["停止→リセット復帰", "再発→センサ位置ずれ→調整", "以降OK"]
    assert _dates(p) == ["2024-04-01", "2024-04-02", "2024-04-03"]


def test_time_shift_and_lot(cells, people):
    p = parse_log(cells[19], date(2024, 11, 12), people)
    s1 = p.segments[0]
    assert (s1.when.date, s1.when.time, s1.when.shift, s1.author.name) == ("2024-11-12", "21:30", "夜勤", "渡部")
    assert s1.identifiers == ["L2411-0123"]
    assert "50枚" in s1.quantities
    assert p.segments[1].author.raw == "品証 小川" and p.segments[1].author.name == "小川"


def test_mmdd_and_m_dot_d_at_line_start(cells, people):
    p = parse_log(cells[21], None, people)   # 基準はセル内の 2024-04-04
    assert _dates(p) == ["2024-04-01", "2024-04-02", "2024-04-03", "2024-04-04"]
    assert p.segments[2].identifiers == ["D4N-4120"]
    assert all(s.author is None for s in p.segments)   # 「調査」は記入者にしない


def test_correction_mark(cells, people):
    p = parse_log(cells[22], date(2024, 5, 7), people)
    assert "correction" in p.segments[1].marks


def test_meeting_bullets_and_plan(cells, people):
    p = parse_log(cells[23], date(2024, 6, 10), people)
    assert len(p.segments) == 6
    assert _dates(p)[:3] == ["2024-06-10", "2024-06-10", "2024-06-10"]
    assert p.segments[1].author is None   # 「暫定：」は人名ではない
    assert p.segments[2].plans == ["6月末目処"]
    assert _authors(p)[3:] == ["佐々木", "山田", "村上"]


def test_same_day_and_part_of_day_lines(cells, people):
    p = parse_log(cells[24], date(2024, 8, 20), people)
    assert len(p.segments) == 4
    w = [s.when for s in p.segments]
    assert (w[0].date, w[0].shift) == ("2024-08-20", "午前中")
    assert (w[1].date, w[1].time) == ("2024-08-20", "15:00")
    assert (w[2].date, w[2].shift, w[2].how) == ("2024-08-20", "夕方", "inherited")
    assert p.segments[1].identifiers == ["F3"] and p.segments[1].quantities == ["2A"]


def test_header_cell(cells, people):
    p = parse_log(cells[25], date(2024, 4, 1), people)
    assert p.kind == "header_cell"
    assert [s.label for s in p.segments] == ["現象", "原因", "対応", "再発防止"]
    assert all("header_cell" in s.marks and s.when is None for s in p.segments)
    assert render_timeline(p, "設備") == ["現象: 起動しない", "原因: ブレーカーOFF（清掃時に誤ってOFF）", "対応: ON復帰",
                                          "再発防止: ブレーカーにカバー取付 4/12 済"]
    off = parse_log(cells[25], date(2024, 4, 1), people, SplitOptions(header_cells="off"))
    assert off.kind == "log"


def test_checklist_with_date_in_paren(cells, people):
    p = parse_log(cells[26], date(2024, 4, 1), people)
    assert len(p.segments) == 3
    assert all("checklist" in s.marks for s in p.segments)
    assert _dates(p)[:2] == ["2024-04-03", "2024-04-03"]
    assert p.segments[0].author.name == "西村" and p.segments[0].body == "冷却ファン交換　済"
    assert p.segments[2].raw.endswith("→③は予算化待ち")


def test_reference_and_plans(cells, people):
    p = parse_log(cells[27], date(2024, 4, 16), people)
    assert "明日来場予定" in p.segments[0].plans
    assert "reference" in p.segments[1].marks
    p32 = parse_log(cells[32], date(2024, 4, 1), people)
    assert "reference" in p32.segments[0].marks and p32.segments[0].identifiers == ["TR-24-00398"]


def test_x000d_and_romaji(cells, people):
    p = parse_log(cells[28], date(2024, 4, 1), people)
    assert "_x000D_" not in p.text
    assert p.segments[0].author.name == "Tanaka" and p.segments[0].body == "E-stop tripped. Reset OK."
    assert p.segments[1].author.name == "田中"


def test_time_range_and_alias_with_paren(cells, people):
    p = parse_log(cells[29], date(2024, 9, 3), people)
    assert p.segments[0].when.time == "08:15"
    assert "85分" in p.segments[0].quantities
    p30 = parse_log(cells[30], date(2024, 3, 4), people)
    assert _authors(p30) == ["佐藤 誠", "佐藤 美咲", "佐藤 誠"]


def test_fullwidth_digits(cells, people):
    p = parse_log(cells[31], None, people)
    assert [(s.when.date, s.when.time) for s in p.segments] == [("2024-04-01", "10:00"), ("2024-04-01", "11:30")]
    assert _authors(p) == ["田中", "田中"]


def test_weekday_and_note_after_date(cells, people):
    p = parse_log(cells[33], date(2024, 4, 5), people)
    assert _dates(p) == ["2024-04-05", "2024-04-08", "2024-04-08"]
    assert "週明け確認予定" in p.segments[0].plans


def test_era_change(cells, people):
    p = parse_log(cells[34], date(2019, 4, 26), people)
    assert _dates(p) == ["2019-04-26", "2019-05-07", "2019-05-08"]
    assert p.segments[0].author.name == "小西"
    assert p.segments[0].body == "連休前に油圧作動油交換（46番 200L）"
    assert p.segments[1].body.startswith("連休明け立上げ時")   # 日付の後の「連休明け」は相対日にしない


# ---- 日付の個別ケース ----

@pytest.mark.parametrize("text,expected", [
    ("令和6年4月1日", (2024, 4, 1)), ("R6/4/1", (2024, 4, 1)), ("H31.4.26", (2019, 4, 26)), ("R元.5.1", (2019, 5, 1)),
    ("平成30年12月1日", (2018, 12, 1)), ("2024-04-01", (2024, 4, 1)), ("2024年4月1日", (2024, 4, 1)), ("24/4/1", (2024, 4, 1)),
])
def test_full_dates(text, expected):
    w = parse_when_at(shadow(text), 0)
    assert (w.year, w.month, w.day, w.full) == (*expected, True)


@pytest.mark.parametrize("text", ["1.5mm 隙間あり", "0.35MPa", "納期4/8 手配", "NG 2/50", "4/32 不明", "1時間停止"])
def test_not_event_dates(text):
    w = parse_when_at(shadow(text), 0, not_date_res=[__import__("re").compile(p) for p in SplitOptions().not_date_patterns])
    assert w is None or not w.has_date


def test_plan_date_in_body_is_not_split(people):
    p = parse_log("4/2 鈴木 部品手配 納期4/8\n4/9 鈴木 交換済", date(2024, 4, 2), people)
    assert len(p.segments) == 2 and _dates(p) == ["2024-04-02", "2024-04-09"]


def test_next_week_range_from_friday():
    heads = [parse_when_at(shadow("4/5(金)"), 0), parse_when_at(shadow("翌週"), 0), parse_when_at(shadow("週明け"), 0)]
    whens, order, _ = resolve_whens(heads, ["4/5(金)", "翌週", "週明け"], date(2024, 4, 5))
    assert (whens[1].date, whens[1].date_to, whens[1].estimated) == ("2024-04-08", "2024-04-14", True)
    assert whens[2].date == "2024-04-15"   # 週明けは直前（翌週の開始日）からの次の月曜
    assert order == "asc"


def test_unresolvable_relative(people):
    p = parse_log("昨日 夜勤者より連絡あり\n4/3 田中 確認", date(2024, 4, 3), people)
    assert p.segments[0].when.date is None and p.segments[0].when.how == "unresolved"
    assert any("昨日" in w for w in p.warnings)


def test_order_anomaly_within_tolerance(people):
    p = parse_log("4/1 田中 停止\n4/5 田中 交換\n4/3 田中 追記：部品到着", date(2024, 4, 1), people)
    assert _dates(p) == ["2024-04-01", "2024-04-05", "2024-04-03"]
    assert any("前後" in w for w in p.warnings)


def test_time_only_lines_join_option(people):
    text = "4/1 10:00 田中：停止\n10:20 再起動で復旧"
    assert len(parse_log(text, date(2024, 4, 1), people).segments) == 2
    joined = parse_log(text, date(2024, 4, 1), people, SplitOptions(time_only_lines="join"))
    assert len(joined.segments) == 1


def test_numbered_lines(people):
    p = parse_log("1. マッチャー取外し\n2. （継続対応中）", date(2024, 4, 1), people)
    assert [s.body for s in p.segments] == ["マッチャー取外し", "（継続対応中）"]
    assert all("checklist" in s.marks for s in p.segments)


def test_extra_anchor(people):
    text = "4/1 田中 停止\n◎追記 部品到着"
    assert len(parse_log(text, date(2024, 4, 1), people).segments) == 1
    p = parse_log(text, date(2024, 4, 1), people, SplitOptions(extra_anchors=[r"◎"]))
    assert len(p.segments) == 2


def test_empty_cells(people):
    for t in ("", None, "-", "－", "  ", "なし"):
        p = parse_log(t, date(2024, 4, 1), people)
        assert p.kind == "empty" and p.segments == []
        assert render_timeline(p, "設備") == []


def test_split_options_from_dict():
    o = SplitOptions.from_dict({"order": "desc", "sentence_split": {"min_chars": 80}, "header_cells": "off",
                                "extra_anchors": ["^◎"], "not_date_patterns": []})
    assert (o.order, o.sentence_split_min_chars, o.header_cells, o.extra_anchors, o.not_date_patterns) == \
        ("desc", 80, "off", ["^◎"], [])


def test_forced_order_option(people):
    p = parse_log("4/2 田中 交換\n4/1 田中 停止", date(2024, 4, 1), people, SplitOptions(order="asc"))
    assert p.order == "asc"


# ---- 記入者 ----

def test_people_index_resolution(people):
    assert people.resolve("K.T").name == "高橋 圭太"
    assert people.resolve("高橋").name == "高橋 圭太"
    amb = people.resolve("佐藤")
    assert amb.name == "佐藤" and "佐藤 誠／佐藤 美咲" in amb.note
    assert people.resolve("佐藤(製)").name == "佐藤 美咲"
    assert people.resolve("保全G").name == "保全G"
    assert people.is_known("山本")   # 担当列の「田中/山本」を分けて登録


def test_known_name_without_separator(people):
    p = parse_log("4/3佐藤(保)エンコーダ交換", date(2024, 4, 3), people)
    assert p.segments[0].author.name == "佐藤 誠" and p.segments[0].body == "エンコーダ交換"


def test_unknown_word_is_not_author():
    p = parse_log("4/3 旧品 保管\n4/4 交換 完了（在庫あり）", date(2024, 4, 3), PeopleIndex())
    assert all(s.author is None for s in p.segments)


# ---- 識別子・数量・予定句 ----

def test_identifiers_exclude_dates_units_and_masks():
    text = "R6.4.1 ALM-2031 発生。0.35MPa、198V、2A。No.3 AM10:00 ［電話番号］ FE-100 6205ZZ Ver4.1 ＰＬＣ２"
    assert extract_identifiers(text) == ["ALM-2031", "FE-100", "6205ZZ", "Ver4.1", "PLC2"]


def test_quantities_and_plans():
    text = "ケーブル手配（RB-ENC-05M ×1、納期1週間）。再発2回。見積¥385,000。1時間に5〜6回。6月末目処で対策予定"
    q = extract_quantities(text)
    assert q[:2] == ["×1", "1週間"] and "2回" in q and "¥385,000" in q and "5〜6回" in q
    assert "05M" not in "".join(q)
    assert extract_plans(text) == ["納期1週間", "6月末目処で対策予定"]


# ---- マスク ----

@pytest.mark.parametrize("phone", [
    "090-1234-5678", "09012345678", "03-1234-5678", "(03)1234-5678", "（03）1234-5678", "0312345678",
    "0120-123-456", "0800-123-4567", "０９０－１２３４－５６７８", "+81-90-1234-5678", "内線1234", "内線：567",
    "03 1234 5678", "080ー1234ー5678",
])
def test_mask_phone_variants(phone):
    masked, spans = mask_text(f"担当 上田様 携帯{phone} まで", ["電話番号"])
    assert masked == "担当 上田様 携帯［電話番号］ まで"
    assert spans[0].kind == "phone" and spans[0].text == phone


@pytest.mark.parametrize("text", ["2024-04-01 10:00", "0401 停止", "ロット L2411-0123", "D4N-4120", "0.35MPa", "1234-5678"])
def test_mask_does_not_touch_non_phone_numbers(text):
    assert mask_text(text, ["phone"])[0] == text


def test_mask_email_amount_person(people):
    text = "yamaguchi@example.co.jp から回答。見積 ¥385,000、追加 12,000円。佐藤 誠さん・高橋さんに連絡（K.T）"
    masked, spans = mask_text(text)
    assert masked.startswith("［メール］ から回答")
    assert "385,000" in masked   # 既定は電話・メールだけ
    masked2, spans2 = mask_text(text, ["email", "金額", "人名"], names=people.names())
    assert "［金額］" in masked2 and "385,000" not in masked2 and "12,000円" not in masked2
    assert "佐藤" not in masked2 and "高橋" not in masked2 and "K.T" not in masked2
    for s in spans2:
        assert text[s.start:s.end] == s.text


# ---- 用語集 ----

GLOSSARY = {
    "様子見": "経過観察",
    "チョコ停": "短時間停止（チョコ停）",
    "FE": {"to": "メーカーのフィールドエンジニア（FE）", "ascii_boundary": True},
    "TEL済": "電話連絡済み",
    "取替": "交換",
    "交換済": "交換済み",
}


def test_glossary_protects_identifiers():
    text = "FE来場、FE-100 の基板取替。チョコ停は様子見。TEL済"
    out = apply_glossary(text, GLOSSARY)
    assert out == "メーカーのフィールドエンジニア（FE）来場、FE-100 の基板交換。短時間停止（チョコ停）は経過観察。電話連絡済み"


def test_glossary_protected_spans_explicit_and_boundary():
    text = "XFE-100 と FE-100"
    # 保護範囲を渡さなくても ascii_boundary で前後が英数字・「-」なら置き換えない
    assert apply_glossary(text, {"FE": {"to": "エンジニア", "ascii_boundary": True}}, protected_spans=[]) == text
    # ascii_boundary なしでも識別子の範囲は保護される
    assert apply_glossary("FE-100 FE", {"FE": {"to": "E", "ascii_boundary": False}}) == "FE-100 E"
    assert apply_glossary("FE-100", {"FE": {"to": "E", "ascii_boundary": False}}, protected_spans=[]) == "E-100"


def test_glossary_longest_first_and_idempotent():
    g = {"交換": "取り替え", "交換済": "交換済み"}
    assert apply_glossary("交換済、交換", g) == "交換済み、取り替え"
    once = apply_glossary("チョコ停", GLOSSARY)
    assert apply_glossary(once, GLOSSARY) == once
    hits = glossary_hits("ﾁｮｺ停と様子見", GLOSSARY)   # 半角カナも照合
    assert [(h.term, h.to) for h in hits] == [("ﾁｮｺ停", "短時間停止（チョコ停）"), ("様子見", "経過観察")]


def test_glossary_skips_mask_tokens():
    assert apply_glossary("［電話番号］へTEL済", {"電話": "でんわ", "TEL済": "電話連絡済み"}) == "［電話番号］へ電話連絡済み"


# ---- 速さのための控え（結果は変えない） ----

def test_shadow_matches_per_character_nfkc():
    """shadow は表と控えを使うが、1文字ずつ NFKC をかけた写し（長さを保つ・丸数字は残す）と同じ。"""
    import unicodedata

    def reference(text):
        out = []
        for ch in text:
            code = ord(ch)
            if code < 0x80 or 0x2460 <= code <= 0x24FF or 0x2776 <= code <= 0x2793:
                out.append(ch)
                continue
            n = unicodedata.normalize("NFKC", ch)
            out.append(n if len(n) == 1 else ch)
        return "".join(out)

    sample = "ＡＢＣ１２３　ｶﾀｶﾅ①②❶ ㍻ ﬁ ㌔ Ⅻ ¥１,０００ ＠ｘ．ｊｐ 全角／半角→ “引用” \t\n漢字"
    everything = "".join(chr(c) for c in range(0x80, 0x10000) if not 0xD800 <= c <= 0xDFFF)
    for text in (sample, everything, "", "ascii only 123"):
        assert shadow(text) == reference(text) and len(shadow(text)) == len(text)
        assert shadow(text) == reference(text)          # 2回目（控えから）も同じ


def test_parse_head_memo_does_not_leak_between_cells(people):
    """同じ位置・同じ長さの別のセルを続けて分けても、前のセルの読み取り結果を使わない。"""
    a = parse_log("4/1 田中：ライン停止\n4/2 佐藤：復旧確認", date(2024, 4, 1), people)
    b = parse_log("【現象】ラインの停止\n【原因】センサー汚れ", date(2024, 4, 1), people)
    c = parse_log("4/1 田中：ライン停止\n4/2 佐藤：復旧確認", date(2024, 4, 1), people)
    assert a.kind == "log" and b.kind == "header_cell"
    assert [s.when.date for s in a.segments] == ["2024-04-01", "2024-04-02"] and repr(a) == repr(c)


def test_extractors_are_unchanged_by_the_cached_spans():
    text = "RB-ENC-05M ×1 手配、納期1週間。ALM-2031 表示、5mm ずれ、N2パージ"
    ids = extract_identifiers(text)
    ids.append("書き換え")                               # 戻り値を変えても控えは変わらない
    assert extract_identifiers(text) == ["RB-ENC-05M", "ALM-2031"]
    assert "×1" in extract_quantities(text) and "5mm" in extract_quantities(text)
    assert extract_plans(text) == ["納期1週間"] and extract_plans("") == [] and extract_plans(None) == []


# ---- R6-AI-1: 単位の後ろに英数字・「-」・カタカナが続くときは寸法とみなさず、日付として読む ----------------------
@pytest.mark.parametrize("text", ["4.3 AGV 停止", "4.3 ALM-2031 発生", "4.3 Aライン停止", "4.3 Vベルト交換"])
def test_unit_letter_followed_by_word_is_still_a_date(text):
    w = parse_when_at(shadow(text), 0, not_date_res=[__import__("re").compile(p) for p in SplitOptions().not_date_patterns])
    assert w is not None and w.has_date and (w.month, w.day) == (4, 3)


@pytest.mark.parametrize("text", ["2.5A 流れた", "3.3V", "1.2 A", "2.5Aの電流"])
def test_plain_units_are_still_not_dates(text):
    w = parse_when_at(shadow(text), 0, not_date_res=[__import__("re").compile(p) for p in SplitOptions().not_date_patterns])
    assert w is None or not w.has_date
