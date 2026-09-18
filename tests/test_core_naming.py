"""core/naming.py: 出力ファイル名。"""
from core.naming import LIGHTRAG_HINT_RECORDS, md_filename, safe_filename_part


def test_safe_filename_part_replaces_unsafe_chars():
    assert safe_filename_part('a\\b/c:d*e?f"g<h>i|j') == "a_b_c_d_e_f_g_h_i_j"
    assert safe_filename_part("CMP研磨装置 1号機") == "CMP研磨装置_1号機"
    assert safe_filename_part("ＣＭＰ－１０１") == "CMP-101"          # NFKC
    assert safe_filename_part("行1\n行2\t\x01") == "行1_行2"
    assert safe_filename_part("[重要] 報告") == "重要_報告"
    assert safe_filename_part("報告.[legacy]") == "報告legacy"       # '.[' は除去
    assert safe_filename_part("..報告書__") == "報告書"
    assert safe_filename_part(None) == ""
    assert safe_filename_part("   ") == ""


def test_safe_filename_part_length_and_reserved():
    assert safe_filename_part("あ" * 100) == "あ" * 60
    assert safe_filename_part("あ" * 10 + "_" + "い" * 10, max_len=11) == "あ" * 10
    assert safe_filename_part("con") == "con_"
    assert safe_filename_part("COM1") == "COM1_"


def test_md_filename():
    assert md_filename(["設備修理報告書", "R2026-00123", "CMP-101"]) == "設備修理報告書_R2026-00123_CMP-101.md"
    assert md_filename(["トラブル対応一覧", "", None, "2026-08"]) == "トラブル対応一覧_2026-08.md"
    assert md_filename(["故障履歴", "2026-08"], hint=LIGHTRAG_HINT_RECORDS) == \
        "故障履歴_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)].md"
    assert md_filename([]) == "無題.md"
