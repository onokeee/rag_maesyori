"""まとめ取り込み: 同じフォームの帳票を一度に読み取り、読み取り結果を縦に全部並べる。

利用者の指示（2026-09-20）:
  「帳票取り込みですが、同じフォームのExcelをまとめて一気に処理できるようにしたい。
    ③の読み取り結果の表示を②タブ選択で切り替えるのではなく、全部のシートをまとめて表示する
    （スクロールして全部確認していく）イメージで」
"""
import io
import json
import zipfile

from models import database as db
from tests.test_forms_flow import activate, add_field, create_type, finish, upload_forms


def _pattern(client, sample_dir) -> int:
    pattern_id = create_type(client, sample_dir / "standard.xlsx", "設備修理報告書")
    add_field(client, pattern_id, "修理報告書", "A3", "B3")     # 報告番号
    add_field(client, pattern_id, "修理報告書", "E4", "F4")     # 設備名
    activate(client, pattern_id)
    return pattern_id


def _copies(sample_dir, tmp_path, count: int) -> list:
    """同じフォームの帳票を名前だけ変えて count 件（まとめて置くファイル）。"""
    data = (sample_dir / "standard.xlsx").read_bytes()
    paths = []
    for i in range(1, count + 1):
        path = tmp_path / f"修理報告書_{i}.xlsx"
        path.write_bytes(data)
        paths.append(path)
    return paths


def read_all(client, ids, pattern_id, sheets, **extra):
    return client.post("/forms/read", data={"ids": ",".join(str(i) for i in ids),
                                            "pattern_id": pattern_id, "sheets": sheets, **extra})


# ---- ② 帳票の種類とシートは、置かれた分すべてで1回だけ決める ------------------------------------

def test_the_type_and_sheets_are_decided_once_for_the_whole_batch(app, client, sample_dir, tmp_path):
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))

    body = client.get("/forms/type?ids=" + ",".join(str(i) for i in ids)).get_json()
    html = body["html"]
    assert body["has_types"] and [d["id"] for d in body["docs"]] == ids
    # 1組の選択（種類のラジオとシートのチェック）が全部のファイルに効く。タブで1件ずつ選ばない
    assert html.count("data-type-form") == 1 and "data-doc-tabs" not in html
    assert f'value="{pattern_id}" checked' in html and html.count('name="sheets"') == 3
    assert f'name="ids" value="{",".join(str(i) for i in ids)}"' in html
    assert "置かれた<strong>3件</strong>すべてに" in html and "3件をまとめて読み取る" in html
    # 合わないファイルを外せるよう、ファイルごとに見つかった項目数と ✕（外す）を出す
    for doc_id in ids:
        assert f'data-file-count="{doc_id}"' in html and f'data-drop-doc="{doc_id}"' in html


# ---- ③ 読み取り結果は全部の帳票を縦に並べる ---------------------------------------------------

def test_three_forms_are_read_and_confirmed_from_the_stacked_view(app, client, sample_dir, tmp_path):
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))

    body = read_all(client, ids, pattern_id, ["修理報告書"]).get_json()
    html = body["html"]
    assert [d["id"] for d in body["docs"]] == ids
    # 3件とも塊として並ぶ（見出しにファイル名と状態）
    assert html.count('class="review-doc"') == 3
    for no, doc_id in enumerate(ids, start=1):
        assert f'id="doc-{doc_id}"' in html and f'id="reviewForm-{doc_id}"' in html
        assert f"修理報告書_{no}.xlsx" in html
    assert html.count("要確認 <strong") == 3 and "未確定" in html
    # 進み具合の1行（まとめ取り込みでは1件ずつ確定しないので、いま見ている場所を出す）
    assert "3件中1件目を表示中" in html and "data-next-doc" in html

    # 3件とも確定でき、zip には3件ぶんの .md が入る
    for doc_id in ids:
        assert client.post(f"/forms/{doc_id}/confirm", json={}).get_json()["ok"] is True
    state = finish(client, ids)
    assert state["confirmed"] == 3 and state["total"] == 3
    assert "確定済み <strong>3</strong> / 3 件" in state["html"]
    assert "まとめて Markdown をダウンロード（zip）" in state["html"]

    with app.app_context():
        batch_id = db.get_document(ids[0])["batch_id"]
    res = client.get(f"/forms/batches/{batch_id}/download.zip")
    assert res.status_code == 200
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        names = zf.namelist()
    assert len(names) == 3 and all(n.endswith(".md") for n in names)
    with app.app_context():
        assert all(db.get_document(i) is None for i in ids)    # 渡したら消える


def test_the_progress_line_counts_the_confirmed_forms(app, client, sample_dir, tmp_path):
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))
    read_all(client, ids, pattern_id, ["修理報告書"])

    client.post(f"/forms/{ids[0]}/confirm", json={})
    html = read_all(client, ids, pattern_id, ["修理報告書"], acknowledge="on").get_json()["html"]
    assert "3件中1件目を表示中" in html
    assert html.count('class="badge badge-green" data-doc-state>確定済み') == 1
    assert finish(client, ids)["confirmed"] == 1


def test_an_edit_in_the_second_block_is_saved_to_that_document(app, client, sample_dir, tmp_path):
    """2つ目の塊で直した値は、その帳票だけに入る（ほかの帳票は触らない）。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))
    body = read_all(client, ids, pattern_id, ["修理報告書"]).get_json()
    second = body["docs"][1]

    res = client.post(f"/forms/{second['id']}/draft",
                      json={"values": {"equipment_name": "CMP研磨装置（2号機）"}, "version": second["version"]})
    assert res.status_code == 204
    with app.app_context():
        saved = [json.loads(db.get_document(i)["data_json"]) for i in ids]
    values = [s["values"]["equipment_name"] for s in saved]
    assert values[1] == "CMP研磨装置（2号機）"
    assert values[0] != "CMP研磨装置（2号機）" and values[2] != "CMP研磨装置（2号機）"
    assert saved[1]["fields"][1]["edited"] and not saved[0]["fields"][1]["edited"]


def test_only_the_first_sheet_grid_comes_with_the_read(app, client, sample_dir, tmp_path):
    """12件置いても重くならないよう、元のシートの表は先頭だけ作り、残りは後から読み込む。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))
    html = read_all(client, ids, pattern_id, ["修理報告書"]).get_json()["html"]

    assert html.count('data-cell="B3"') == 1          # 表は先頭の帳票だけ
    assert html.count("data-grid-wait") == 2          # 残りは画面に入ってから
    for doc_id in ids:
        assert f'data-grid-url="/forms/{doc_id}/grid"' in html

    grid = client.get(f"/forms/{ids[1]}/grid").get_json()
    assert grid["doc_id"] == ids[1] and 'data-cell="B3"' in grid["html"]
    assert client.get(f"/forms/{ids[1]}/grid", headers={}).status_code == 200


def test_a_file_that_does_not_match_can_be_left_out(app, client, sample_dir, tmp_path):
    """合わないファイルは外して、残りをそのまままとめて読み取れる。"""
    pattern_id = _pattern(client, sample_dir)
    paths = _copies(sample_dir, tmp_path, 3)
    ids = upload_forms(client, *paths)

    assert client.post(f"/forms/{ids[1]}/delete").get_json()["ok"] is True
    rest = [ids[0], ids[2]]
    body = client.get("/forms/type?ids=" + ",".join(str(i) for i in ids)).get_json()
    assert [d["id"] for d in body["docs"]] == rest    # 外したファイルはもう出ない

    html = read_all(client, rest, pattern_id, ["修理報告書"]).get_json()["html"]
    assert html.count('class="review-doc"') == 2 and "2件中1件目を表示中" in html


# ---- 1件だけのときは今までどおり（まとまりの1行も出さない） ---------------------------------------

def test_a_single_file_review_has_one_block_and_no_batch_line(app, client, sample_dir):
    pattern_id = _pattern(client, sample_dir)
    doc_id, = upload_forms(client, sample_dir / "standard.xlsx")

    body = client.get(f"/forms/{doc_id}/type").get_json()
    assert "2項目中2項目が見つかりました" in body["html"] and "置かれた<strong>" not in body["html"]
    assert "この帳票はやめる" in body["html"]

    body = read_all(client, [doc_id], pattern_id, ["修理報告書"]).get_json()
    html = body["html"]
    assert html.count('class="review-doc"') == 1
    assert "data-batch-bar" not in html and "件目を表示中" not in html
    assert 'data-cell="B3"' in html and "data-grid-wait" not in html   # 1件のときは元のシートもすぐ出す
    assert 'name="value-report_id"' in html and body["version"]

    # 確定して .md をダウンロードするところまで今までどおり
    assert client.post(f"/forms/{doc_id}/confirm", json={"version": body["version"]}).get_json()["ok"] is True
    state = finish(client, [doc_id])
    assert "zip" not in state["html"] and f"/forms/{doc_id}/download.md" in state["html"]
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 200


# ---- 確定したあとに直した帳票は、ダウンロードの前に確定し直す ---------------------------------------

def test_a_modified_form_is_confirmed_again_before_the_md_download(app, client, sample_dir):
    """1件だけのときも zip と同じで、直した値が .md に入らずに消えることが無いようにする。

    確定したあとに読み取り結果を直すと状態は「修正中」になる。このとき .md は確定済みの古い版から
    作られるので、ダウンロードのボタンは zip と同じ data-confirm-all を持ち（review.js が先に確定し直す）、
    確認文にもそのことを書く。
    """
    pattern_id = _pattern(client, sample_dir)
    doc_id, = upload_forms(client, sample_dir / "standard.xlsx")
    body = read_all(client, [doc_id], pattern_id, ["修理報告書"]).get_json()

    client.post(f"/forms/{doc_id}/confirm", json={"version": body["version"]})
    state = finish(client, [doc_id])
    assert "data-confirm-all" in state["html"]                       # 確定済みでも同じ通り道を使う
    assert "確定し直してから" not in state["html"]                    # 直していなければ断らない

    # 確定したあとに直す → 修正中。このままでは .md は確定済みの古い版から作られる
    res = client.post(f"/forms/{doc_id}/draft",
                      json={"values": {"equipment_name": "直した設備名"}, "version": body["version"]})
    assert res.status_code == 204
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "modified"
        assert "直した設備名" not in db.get_document(doc_id)["confirmed_json"]

    state = finish(client, [doc_id])
    assert "data-confirm-all" in state["html"] and "確定し直してから、直した値で Markdown を作ります。" in state["html"]

    # review.js が先に確定し直すので、渡す .md には直した値が入る
    client.post(f"/forms/{doc_id}/confirm", json={"version": res.headers["X-Doc-Version"]})
    assert "直した設備名" in client.get(f"/forms/{doc_id}/download.md").get_data(as_text=True)
