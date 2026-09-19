"""core/mdtext.py: Markdown テキスト処理。"""
from core.mdtext import escape_md_line, estimate_tokens, join_blocks, md_bullet, nfkc_value


def test_nfkc_value():
    assert nfkc_value("ＣＭＰ－１０１　 研磨") == "CMP-101 研磨"
    assert nfkc_value("フィルター  交換\r\n\r\n  流量　再校正  ") == "フィルター 交換\n\n流量 再校正"
    assert nfkc_value("原点 復帰") == "原点 復帰"   # 日本語文字間の半角空白は残す
    assert nfkc_value(None) == ""
    assert nfkc_value(95) == "95"
    # 丸数字は NFKC で囲みが外れると「①破損」が「1破損」になり番号と本文の区切りが消えるので、そのまま残す
    assert nfkc_value("①破損ウェーハ片を回収\n②Head3 メンブレン交換") == "①破損ウェーハ片を回収\n②Head3 メンブレン交換"
    assert nfkc_value("⑳Ⓐ㋐㊤") == "⑳Ⓐ㋐㊤"
    # 区切りが残る表記は今までどおり NFKC で正規化する
    assert nfkc_value("㈱テスト ⑴ ⒈ ﾎﾟﾝﾌﾟ") == "(株)テスト (1) 1. ポンプ"


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
    # 実トークン（tiktoken o200k_base）以上になる見積もり: 非ASCII 1文字=1.1、ASCII 2文字=1
    assert estimate_tokens("") == 0
    assert estimate_tokens("故障") == 3
    assert estimate_tokens("abc") == 2
    assert estimate_tokens("abcd") == 2
    assert estimate_tokens("CMP研磨") == 4
    assert estimate_tokens("あ" * 100) == 110
    # ASCII をまとめて数えても、1文字ずつ数えたときと同じ（U+007F/U+0080 の境目・絵文字・孤立サロゲート）
    assert estimate_tokens("\x7f\x80") == 2
    assert estimate_tokens("a\U0001F600b") == 3
    assert estimate_tokens("\ud800x") == 2


def test_estimate_tokens_is_not_below_the_real_count_for_dates_and_part_numbers():
    """日時・品番・計測値は o200k_base で細かく区切られる。右の数は tiktoken（gpt-4o-mini）で数えた実トークン数。"""
    real = {
        "2023-09-01 09:44": 10,
        "- 発生日時: 2023-09-01 09:44": 16,
        "8/3 10:07 田中：FDC ch3 V=1039V I=2.34A P=1.11kW 0x0B73 R=11.1ohm OK": 47,
        "PN-A12345-B67 SN:0x1F3A9C": 16,
        "1,234,567.89": 7,
        "2026/08/03 10:07〜2026/08/03 11:45": 21,
        "ロット番号 L2308-0412-07 の不良率 0.35%": 20,
    }
    for text, count in real.items():
        assert estimate_tokens(text) >= count, text
    # 数字のまとまりは、前が ASCII でない文字・記号・先頭でも数える（「1年2月」は2つ）
    assert estimate_tokens("1年2月") == estimate_tokens("x年y月") + 1
    # 40行の計測ログ。実 1,661 トークン（チャンク 1,500 を超える）を、旧式は 1,252 と見積もり記録の上限 1,400 を通していた
    log = "\n".join(f"8/3 {10 + i * 7 // 60}:{i * 7 % 60:02d} 田中：FDC ch{i % 8} V={1000 + i * 13}V "
                    f"I={2.31 + i / 100:.2f}A 0x{i * 977:04X} dP={i * 0.013:.3f}kPa OK" for i in range(40))
    assert estimate_tokens(log) >= 1661


def test_join_blocks():
    text = join_blocks([["# タイトル"], [], ["- a: 1", "", "- b: 2\r\n  続き"], ["- 出典: x.xlsx"]])
    assert text == "# タイトル\n\n- a: 1\n- b: 2\n  続き\n\n- 出典: x.xlsx\n"
    assert join_blocks([]) == ""
    assert join_blocks([[""], []]) == ""
