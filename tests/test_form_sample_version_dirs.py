"""帳票サンプル（F1〜F5）の版フォルダ名と、その中身の置き方のテスト。

フォルダ名の付け方は5つの生成スクリプトで共通にしてある。
「版の名前＋いつから」を日本語で書き、区別に要るもの（様式番号・用紙サイズ・シート枚数）だけを足す。
ローマ字は使わず、Windows のフォルダ名に使えない文字も使わない。
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from scripts.samples import f1_repair_report as f1
from scripts.samples import f2_trouble_report as f2
from scripts.samples import f3_8d_report as f3
from scripts.samples import f4_inspection_report as f4
from scripts.samples import f5_process_abnormality as f5

SAMPLES = Path(__file__).resolve().parents[1] / "samples" / "forms"

# 帳票フォルダ名 → そのスクリプトが作る版フォルダ名
FAMILIES: dict[str, list[str]] = {
    f1.FORM_DIR: [f1._version_dir(v) for v in f1.REV_INFO],
    f2.FOLDER: list(f2.VERSION_DIR.values()),
    f3.FOLDER: list(f3.VERSION_DIRS.values()),
    f4.FORM_DIR: [f4._rev_dir(v) for v in f4.REV_INFO],
    f5.FORM_FOLDER: list(f5.VERSION_FOLDERS.values()),
}

NG_CHARS = '.:/\\*?"<>|'
# 名前に出てよい英数字は、様式番号（QA-F-021）・版の名前（Rev3）・用紙サイズ（A3/A4）・年だけ
ALLOWED_TOKENS = re.compile(r"QA-F-\d{3}|Rev\d+|A[34]|\d+|_")


def _families():
    return [(family, name) for family, names in FAMILIES.items() for name in names]


def test_version_dir_names_are_windows_safe():
    """版フォルダ名に Windows で使えない文字・前後の空白・末尾のピリオドが無い。"""
    for family, name in _families():
        assert name, family
        assert not set(name) & set(NG_CHARS), f"{family}/{name}"
        assert name == name.strip() and not name.endswith(("。", "、")), f"{family}/{name}"
        assert len(name) <= 60, f"{family}/{name}"


def test_version_dir_names_have_no_romaji():
    """版フォルダ名にローマ字（hougan・cols・sheet など）が残っていない。"""
    for family, name in _families():
        rest = ALLOWED_TOKENS.sub("", name)
        assert not re.search(r"[A-Za-z]", rest), f"ローマ字が残っている: {family}/{name}"


def test_version_dir_names_say_the_revision_and_when():
    """版フォルダ名は「版の名前＋いつから」。版の名前と年があり、制定／改訂／まで で終わる。"""
    for family, name in _families():
        assert re.search(r"Rev\d+", name), f"版の名前が無い: {family}/{name}"
        assert re.search(r"(19|20)\d{2}", name), f"年が無い: {family}/{name}"
        assert name.endswith(("制定", "改訂", "まで")), f"いつからが無い: {family}/{name}"


def test_version_dir_names_are_unique_in_a_family():
    """同じ帳票の中で版フォルダ名がぶつからない（別レイアウトを同じ名前にしていない）。"""
    for family, names in FAMILIES.items():
        assert len(set(names)) == len(names), family
    assert [len(v) for v in FAMILIES.values()] == [3, 2, 5, 2, 3]


def test_every_layout_version_has_a_folder():
    """生成スクリプトが作りうる版（layout_version）には、必ず入れ先のフォルダがある。"""
    assert set(f1.REV_INFO) == {"v1", "v2", "v3"} and set(f4.REV_INFO) == {"v1", "v2"}
    assert set(f2.VERSION_DIR) == {f2.OLD, f2.NEW}
    assert set(f5.VERSION_FOLDERS) == {"A3横_Rev.1", "A3横_Rev.2", "A4縦_Rev.3"}
    made = {f3._make_style(rank, claim, rank).layout_version for claim in (False, True) for rank in range(20)}
    assert made == set(f3.VERSION_DIRS), made ^ set(f3.VERSION_DIRS)


# ---- samples/forms の実ファイル ----

def _family_dir(name: str) -> Path:
    path = SAMPLES / name
    if not path.exists():
        pytest.skip(f"サンプルがありません: {name}")
    return path


@pytest.mark.samples
def test_sample_tree_has_one_folder_per_version():
    """帳票フォルダの直下は版フォルダと _README.md・_expected.jsonl だけで、版フォルダの中は .xlsx だけ。"""
    for family, names in FAMILIES.items():
        root = _family_dir(family)
        assert sorted(p.name for p in root.iterdir() if p.is_dir()) == sorted(names), family
        assert sorted(p.name for p in root.iterdir() if p.is_file()) == ["_README.md", "_expected.jsonl"], family
        for name in names:
            kids = list((root / name).iterdir())
            assert kids, f"{family}/{name} が空"
            assert all(k.is_file() and k.suffix == ".xlsx" for k in kids), f"{family}/{name}"


@pytest.mark.samples
def test_sample_expected_file_is_the_path_under_the_version_folder():
    """_expected.jsonl の file は「版フォルダ/ファイル名」で、実在する .xlsx を指している。"""
    for family, names in FAMILIES.items():
        root = _family_dir(family)
        rows = [json.loads(line) for line in (root / "_expected.jsonl").read_text(encoding="utf-8").splitlines() if line.strip()]
        assert len(rows) == 30, family
        for row in rows:
            folder, _, filename = row["file"].partition("/")
            assert folder in names, f"{family}: {row['file']}"
            assert filename.endswith(".xlsx") and "/" not in filename, f"{family}: {row['file']}"
            assert (root / folder / filename).exists(), f"{family}: {row['file']}"
        # 版フォルダの .xlsx は全部 _expected.jsonl に載っている
        listed = {row["file"] for row in rows}
        actual = {f"{d.name}/{p.name}" for d in root.iterdir() if d.is_dir() for p in d.iterdir()}
        assert listed == actual, family


@pytest.mark.samples
def test_sample_readme_names_every_version_folder():
    """_README.md がフォルダ名を載せていて、フォルダごとまとめて取り込めると書いてある。"""
    for family, names in FAMILIES.items():
        text = _family_dir(family).joinpath("_README.md").read_text(encoding="utf-8")
        assert "フォルダ" in text.split("\n##")[0] or "フォルダ" in text, family
        for name in names:
            assert name in text, f"{family}: {name} が _README.md に無い"
        assert "帳票取り込み" in text and "まとめて" in text, family
        assert "版の名前＋いつから" in text, family


def test_rerun_removes_the_folders_of_old_names(tmp_path):
    """作り直すと、前の名前の版フォルダと中の .xlsx が残らない（F4 で確かめる）。"""
    out = tmp_path / "forms" / f4.FORM_DIR
    stale = out / "Rev1_2018seitei"          # 前の名前のフォルダ
    stale.mkdir(parents=True)
    (stale / "古い報告書.xlsx").write_bytes(b"old")
    (stale / "メモ.txt").write_text("消えてよい", encoding="utf-8")
    (out / "直下の残り.xlsx").write_bytes(b"old")

    paths = f4.generate(tmp_path)

    assert not stale.exists() and not (out / "直下の残り.xlsx").exists()
    assert sorted(p.name for p in out.iterdir() if p.is_dir()) == sorted(FAMILIES[f4.FORM_DIR])
    assert {p.parent.name for p in paths if p.suffix == ".xlsx"} == set(FAMILIES[f4.FORM_DIR])
    rows = [json.loads(line) for line in (out / "_expected.jsonl").read_text(encoding="utf-8").splitlines()]
    assert all((out / row["file"]).exists() for row in rows)
