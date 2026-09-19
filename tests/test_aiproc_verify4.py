"""AI整形の照合（4巡目）：根拠にない原因・処置、再発・完了の否定形、分数・頻度の数。"""
import pytest

from aiproc import prompts
from aiproc.verify import verify_log_result
from logproc import parse_log
from tests.test_aiproc import STAGE, _good_412, _verify


def _run(cell: str, incident: dict):
    parse = parse_log(cell)
    sent = prompts.build_log_messages(parse, {}, STAGE)[-1]["content"]
    entries = [{"id": f"e{i + 1}", "segs": [s.id], "t": ["メモ"]} for i, s in enumerate(parse.segments)]
    return verify_log_result({"entries": entries, "incident": incident}, parse, sent, spec=STAGE)


def test_confirmed_cause_on_unknown_evidence_is_rejected():
    rep = _run("4/1 佐藤：搬送停止、リセットで復旧。\n4/2 佐藤：原因は不明、調査継続",
               {"root_cause": {"q": "原因は不明", "v": "ベアリング摩耗", "certainty": "確定", "src": ["e2"]}})
    assert "incident.root_cause" in rep.failed_items, rep.issues
    assert any(i.code == "speculation" and "不明" in i.message for i in rep.issues)
    assert "root_cause" not in rep.accepted["incident"]


def test_confirmed_cause_needs_quote():
    result = _good_412()
    result["incident"]["root_cause"]["q"] = None
    rep = _verify(result)
    assert "incident.root_cause" in rep.failed_items and any(i.code == "quote" for i in rep.issues)
    # 「疑い」なら引用なしでも出す（従来どおり）
    result["incident"]["root_cause"]["certainty"] = "疑い"
    assert "incident.root_cause" in _verify(result).ok_items


def test_content_not_in_evidence_is_flagged_not_silently_ok():
    rep = _run("4/1 佐藤：搬送停止。\n4/2 佐藤：様子見",
               {"permanent_actions": [{"v": "安全カバーの設置", "src": ["e2"]}],
                "parts": [{"name": "ベアリング", "src": ["e2"]}]})
    assert rep.status() == "flagged"
    assert {i.path for i in rep.issues if i.code == "content"} == {"incident.permanent_actions[0]", "incident.parts[0]"}
    # 用語集の言い換え（様子見→経過観察）は根拠にある語として扱う
    rep = _run("4/1 佐藤：搬送停止。\n4/2 佐藤：再起動後、様子見",
               {"temporary_actions": [{"v": "経過観察", "src": ["e2"]}]})
    assert not any(i.code == "content" for i in rep.issues), rep.issues
    # 研究例（言い換えを含む）は ok のまま
    assert _verify(_good_412()).status() == "ok"


@pytest.mark.parametrize("line", ["その後、再発はなし", "再度測定し正常、再発は見られない", "再度測定し正常"])
def test_recurrence_yes_rejected_on_negated_evidence(line):
    rep = _run(f"4/1 佐藤：モーター異音。ベアリング交換\n4/2 佐藤：{line}", {"recurrence": {"v": "あり", "src": ["e2"]}})
    assert "incident.recurrence" in rep.failed_items, rep.issues


@pytest.mark.parametrize("line", ["同エラー再発2回", "再度停止したためリセット", "再発2回、以降再発なし"])
def test_recurrence_yes_still_accepted(line):
    rep = _run(f"4/1 佐藤：モーター異音。\n4/2 佐藤：{line}", {"recurrence": {"v": "あり", "src": ["e2"]}})
    assert "incident.recurrence" in rep.ok_items, rep.issues


def test_recurrence_no_accepts_wa_nashi():
    rep = _run("4/1 佐藤：モーター異音。ベアリング交換\n4/2 佐藤：その後、再発はなし", {"recurrence": {"v": "なし", "src": ["e2"]}})
    assert "incident.recurrence" in rep.ok_items, rep.issues


@pytest.mark.parametrize("line,ok", [
    ("来週交換し完了予定", False), ("交換作業は完了していない", False), ("交換は今月中に完了見込み", False),
    ("まだ解決していない", False), ("再発なし→クローズ", True), ("交換完了", True),
])
def test_done_rejects_not_yet_forms(line, ok):
    rep = _run(f"4/1 佐藤：モーター異音。ベアリング発注\n4/2 佐藤：{line}", {"final_state": {"v": "完了", "src": ["e2"]}})
    assert ("incident.final_state" in rep.ok_items) is ok, rep.issues


@pytest.mark.parametrize("cell,v", [
    ("4/1 佐藤：流量不足。\n4/2 佐藤：バルブ開度を1/2に調整し復旧", "バルブ開度を1/2に調整"),
    ("4/1 佐藤：目視漏れ。\n4/2 佐藤：1日1回の目視点検を追加", "1日1回の目視点検を追加"),
    ("4/1 佐藤：ボルト緩み。\n4/2 佐藤：ボルトを1/4回転増し締め", "ボルトを1/4回転増し締め"),
])
def test_fraction_and_frequency_are_not_dates(cell, v):
    rep = _run(cell, {"permanent_actions": [{"v": v, "src": ["e2"]}]})
    assert not any(i.code == "date" for i in rep.issues), rep.issues
    assert "incident.permanent_actions[0]" in rep.ok_items


@pytest.mark.parametrize("v", ["4/1 10:20に原点復帰", "翌日に再起動", "4/8にケーブル手配", "4/1/2024に交換", "4月8日に交換"])
def test_real_dates_still_rejected(v):
    rep = _run("4/1 佐藤：原点復帰エラー。\n4/2 佐藤：再起動、ケーブル手配、交換",
               {"temporary_actions": [{"v": v, "src": ["e2"]}]})
    assert any(i.code == "date" for i in rep.issues), (v, rep.issues)
