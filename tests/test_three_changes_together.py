"""2026-09-21 の3つの直しが、3画面そろって噛み合っているかの通しテスト。

直した3つ:
  1. 帳票サンプルの版フォルダ名を「版の名前＋いつから」にそろえた（scripts/samples/*）
  2. 帳票登録は Excel を預からず、設定だけ残す（views/form_types.py・core/workbook_cache.py）
  3. 「列の対応づけ」の言葉づかい（設備 → 対象、追記ログ → 経過の記録、「知らせ」の列をやめた）

ここで見るのは、3つの担当が別々に直したあとに残りやすい「つなぎ目」:
  - 帳票取り込みの画面が、まだ「見本の Excel から…」と案内していないか（帳票登録はもう預からない）
  - 一覧表の検証メッセージの中で「設備」と「対象」が混ざっていないか
  - 版フォルダを丸ごと置くと、1つの帳票の種類・1組のシートでまとまって読めるか（1ファイル＝1つの .md）
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest

from models import database as db
from pattern.builder import suggest_rows
from pattern.forms import rows_to_pattern
from excel.workbook import load_workbook_info
from tables.spec import spec_from_dict, validate_spec

SAMPLES = Path(__file__).resolve().parents[1] / "samples" / "forms"

# 置いた Excel を預からなくなったので、画面のどこにも出さない言葉（帳票側）
NO_SAMPLE_WORDS = ("見本の Excel", "見本のExcel", "見本ファイル")
# 役割の名前を変えたので、画面のどこにも出さない言葉（一覧表側）
NO_ROLE_WORDS = ("追記ログ", "AI整形の対象", ">設備</option>", ">知らせ<")


# ---- つなぎ目1: 帳票取り込みの画面が「見本」と言わない -------------------------------------

def test_forms_page_does_not_promise_a_kept_sample_excel(client):
    """帳票取り込みの入口（種類がまだ無いとき）が「見本の Excel から」と案内しない。

    帳票登録は Excel を預からなくなったので、「見本」と書くと預かるように読めてしまう。
    """
    html = client.get("/forms/", follow_redirects=True).get_data(as_text=True)
    assert "帳票の種類がまだありません" in html
    for word in NO_SAMPLE_WORDS:
        assert word not in html, word


def test_type_fragment_without_active_types_does_not_say_sample(client, sample_dir):
    """帳票を置いたのに使用中の種類が無いときの案内も「見本」と言わない。"""
    from tests.test_forms_flow import book_part

    res = client.post("/forms/upload", data={"file": book_part(sample_dir / "standard.xlsx")},
                      content_type="multipart/form-data")
    doc_id = res.get_json()["docs"][0]["id"]
    html = client.get(f"/forms/{doc_id}/type").get_json()["html"]
    assert "使用中の帳票の種類がありません" in html
    for word in NO_SAMPLE_WORDS:
        assert word not in html, word


def test_no_screen_says_sample_or_log_words(client, sample_dir):
    """3画面とも、やめた言葉を1つも出さない（帳票登録で種類を1つ作ったあとでも）。"""
    from tests.test_forms_flow import add_field, create_type, panel_html

    pattern_id = create_type(client, sample_dir / "standard.xlsx", name="通し確認")
    add_field(client, pattern_id, "修理報告書", "A4", "B4")
    pages = [client.get(p, follow_redirects=True).get_data(as_text=True)
             for p in ("/forms/", "/tables", "/form-types/")]
    pages.append(panel_html(client, pattern_id, book=sample_dir / "standard.xlsx"))
    pages.append(panel_html(client, pattern_id))          # Excel を置いていない画面
    for html in pages:
        for word in NO_SAMPLE_WORDS + NO_ROLE_WORDS:
            assert word not in html, word


# ---- つなぎ目2: 一覧表の検証メッセージで役割の呼び方をそろえる -----------------------------

def test_spec_errors_use_the_new_role_names_only():
    """④の保存から出る検証メッセージが「設備」と「対象」を混ぜない。

    役割 entity の画面名は「対象（設備・製品・顧客など）」。まとめ方の説明だけ「設備×月」のままだと、
    同じものを2つの名前で呼ぶことになる。
    """
    bad = spec_from_dict({"name": "", "columns": [{"key": "a", "display": "あ", "type": "string"}],
                          "markdown": {"group_by": "entity_month"},
                          "log_stage": {"column": "無い列"}})
    joined = "\n".join(validate_spec(bad))
    assert "対象×月でまとめるには" in joined
    assert "役割「対象（設備・製品・顧客など）」" in joined
    assert "経過の記録の列「無い列」がありません" in joined
    assert "設備×月" not in joined
    assert "役割「entity" not in joined and "AI整形の対象列" not in joined


def test_two_entity_columns_are_refused_with_the_new_name():
    """役割「対象」を2列に付けたときの断り文も新しい名前で言う。"""
    spec = spec_from_dict({"name": "一覧", "columns": [
        {"key": "equipment_id", "display": "設備番号", "type": "code", "role": "entity"},
        {"key": "product_id", "display": "製品番号", "type": "code", "role": "entity"},
    ]})
    joined = "\n".join(validate_spec(spec))
    assert "役割「対象（設備・製品・顧客など）」の列は1つだけにしてください" in joined


# ---- つなぎ目3: 版フォルダを丸ごと置く ---------------------------------------------------

def _version_dirs() -> list[Path]:
    """samples/forms の版フォルダ全部（samples を作っていなければ空）。"""
    if not SAMPLES.is_dir():
        return []
    return sorted(v for family in sorted(SAMPLES.iterdir()) if family.is_dir()
                  for v in family.iterdir() if v.is_dir())


def _register(app, family: Path) -> int:
    """その様式の帳票の種類を作る（版が混ざるように、版ごとに1件ずつ見て候補を作る）。"""
    picks = [sorted(v.glob("*.xlsx"))[0] for v in sorted(p for p in family.iterdir() if p.is_dir())][:3]
    sheet_rows, field_rows = suggest_rows([load_workbook_info(p) for p in picks])
    with app.app_context():
        pattern_id = db.create_pattern(family.name)
        db.save_pattern(rows_to_pattern(pattern_id, {"name": family.name, "version": ""}, sheet_rows, field_rows),
                        "active")
    return pattern_id


@pytest.mark.samples
@pytest.mark.parametrize("family_name", ["F1_設備修理報告書", "F5_工程異常連絡票"])
def test_a_whole_version_folder_reads_as_one_batch(app, client, family_name):
    """版フォルダの .xlsx を丸ごと置くと、1つの帳票の種類・1組のシートで全件読めて、
    zip には1ファイルにつき1つの .md が入る（版ごとにフォルダを分けた狙い）。"""
    family = SAMPLES / family_name
    if not family.is_dir():
        pytest.skip("samples/forms がありません")
    pattern_id = _register(app, family)
    version = sorted(p for p in family.iterdir() if p.is_dir())[0]
    files = sorted(version.glob("*.xlsx"))

    parts = [(io.BytesIO(p.read_bytes()), p.name) for p in files]   # 画面でフォルダを丸ごと置いたのと同じ
    res = client.post("/forms/upload", data={"file": parts}, content_type="multipart/form-data")
    assert res.status_code == 200, res.get_data(as_text=True)
    body = res.get_json()
    ids = [d["id"] for d in body["docs"]]
    assert len(ids) == len(files), f"{version.name}: 置いた {len(files)} 件のうち {len(ids)} 件しか取り込めていない"
    batch_id = body["batch_id"]

    # 帳票の種類とシートは1回だけ選ぶ（置いた全部に同じものを使う）
    html = client.get("/forms/type", query_string={"ids": ",".join(map(str, ids))}).get_json()["html"]
    chosen = re.search(r'name="pattern_id" value="(\d+)" checked', html)
    assert chosen is not None and int(chosen.group(1)) == pattern_id, "その様式の種類が選ばれていない"
    sheets = re.findall(r'name="sheets" value="([^"]+)" checked', html)
    assert sheets, "読み取るシートにチェックが付いていない"

    form = {"ids": ",".join(map(str, ids)), "pattern_id": str(pattern_id), "sheets": sheets}
    assert client.post("/forms/read", data=form).status_code == 200
    for doc_id in ids:
        assert client.post(f"/forms/{doc_id}/confirm", json={}).status_code == 200, doc_id

    res = client.get(f"/forms/batches/{batch_id}/download.zip")
    assert res.status_code == 200
    with zipfile.ZipFile(io.BytesIO(res.get_data())) as zf:
        names = zf.namelist()
    assert len(names) == len(files) == len(set(names)), f"{version.name}: {len(files)}件 → {names}"
    assert all(n.endswith(".md") for n in names)


@pytest.mark.samples
def test_every_version_folder_holds_only_xlsx_so_it_can_be_dropped_whole():
    """どの版フォルダも中身は .xlsx だけ（_README.md などが混ざると丸ごと置けない）。"""
    dirs = _version_dirs()
    if not dirs:
        pytest.skip("samples/forms がありません")
    for version in dirs:
        kids = sorted(version.iterdir())
        assert kids and all(k.is_file() and k.suffix == ".xlsx" for k in kids), \
            f"{version.parent.name}/{version.name}: {[k.name for k in kids]}"
