"""core/mdtext.py: Markdown テキスト処理。"""
from core.mdtext import escape_md_line, estimate_tokens, join_blocks, md_bullet, nfkc_value


def test_nfkc_value():
    assert nfkc_value("ＣＭＰ－１０１　 研磨") == "CMP-101 研磨"
    assert nfkc_value("フィルター  交換\r\n\r\n  流量　再校正  ") == "フィルター 交換\n\n流量 再校正"
    assert nfkc_value("原点 復帰") == "原点 復帰"   # 日本語文字間の半角空白は残す
    assert nfkc_value(None) == ""
    assert nfkc_value(95) == "95"
    assert nfkc_value("①") == "1"


def test_escape_md_line():
    assert escape_md_line("# 見出し") == "\\# 見出し"
    assert escape_md_line("#123 は番号") == "#123 は番号"
    assert escape_md_line("- 項目") == "\\- 項目"
    assert escape_md_line("-5℃") == "-5℃"
    assert escape_md_line("* 注") == "\\* 注"
    assert escape_md_line("+ 追加") == "\\+ 追加"
    assert escape_md_line("> 引用") == "\\> 引用"
    assert escape_md_line("1. 手順") == "1\\. 手順"
    assert escape_md_line("2) 手順") == "2\\) 手順"
    assert escape_md_line("2026.08.03 対応") == "2026.08.03 対応"
    assert escape_md_line("---") == "\\---"
    assert escape_md_line("= = =") == "\\= = ="
    assert escape_md_line("```python") == "\\`\\`\\`python"
    assert escape_md_line("  # 字下げ") == "  \\# 字下げ"
    assert escape_md_line("普通の文") == "普通の文"
    assert escape_md_line("") == ""


def test_md_bullet():
    assert md_bullet("停止時間", "95分") == ["- 停止時間: 95分"]
    assert md_bullet("停止時間", 95) == ["- 停止時間: 95"]
    assert md_bullet("処置", "フィルター交換\n\n流量再校正") == ["- 処置:", "  フィルター交換", "  流量再校正"]
    assert md_bullet("処置", "# 手順\n1. 交換") == ["- 処置:", "  \\# 手順", "  1\\. 交換"]
    assert md_bullet("時系列", ["1. 2026-08-03 14:20［連絡・初動］田中", "2. 復旧"]) == \
        ["- 時系列:", "  1\\. 2026-08-03 14:20［連絡・初動］田中", "  2\\. 復旧"]
    assert md_bullet("原因", "") == []
    assert md_bullet("原因", None) == []
    assert md_bullet("原因", "\n  \n") == []


def test_estimate_tokens():
    assert estimate_tokens("") == 0
    assert estimate_tokens("故障") == 2
    assert estimate_tokens("abc") == 1
    assert estimate_tokens("abcd") == 2
    assert estimate_tokens("CMP研磨") == 3


def test_join_blocks():
    text = join_blocks([["# タイトル"], [], ["- a: 1", "", "- b: 2\r\n  続き"], ["- 出典: x.xlsx"]])
    assert text == "# タイトル\n\n- a: 1\n- b: 2\n  続き\n\n- 出典: x.xlsx\n"
    assert join_blocks([]) == ""
    assert join_blocks([[""], []]) == ""
