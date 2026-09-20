# 統合設計書：RAG用Markdown作成アプリ（帳票・一覧表）

版: 2026-09-15（作り直し v2）。この文書が実装の正本。根拠となった調査・設計は `docs/research/` に置く。

## 0. 前提と範囲

- 利用者は1人。**ログインなし**。起動は `127.0.0.1` のみ（`app.py` が他のアドレスを拒否）。
- 目的：Excel/CSV を読み取り、**LightRAG に手作業で投入しやすい Markdown(.md) を作ってダウンロードする**（または決めたフォルダ＝LightRAG の `INPUT_DIR` などに保存する。2.8）まで。LightRAG への送信・アプリ内のSQL照会は作らない。
- 取り込みは2系統に分ける。
  - **帳票**（1ファイル＝1件。例：設備修理報告書）… 既存の「ラベル探索」方式を改良。
  - **一覧表**（Excel/CSV、1行＝1件。例：トラブル対応一覧、故障履歴CSV、月別停止時間のクロス集計）… 新規。
- 一覧表の文章列（例：複数人の追記ログが1セルに入った「対応内容」）は、**ルールで時系列に分解**し、任意で **AI整形（keep モード、原文照合つき）** を重ねる。
- AI は OpenAI 互換 API（既存 `services/llm.py`。ヘッダーでモデル選択）。AI なしでも全フローが完結すること。
- 対象 LightRAG: 1.5.7 を主対象（1.4.x の差分は案内ページに注記）。
- **範囲の見直し（2026-09-16、利用者の判断）**: 一覧表の更新管理（期間の置き換え・投入済みとの差分・投入済みにする・出力画面・取り消し）、クロス集計、名寄せ辞書は作らない（8章「後回し」）。一覧表は**取り込みごとに、その取り込みの記録だけから全 md を作り、全ファイルを zip で渡す**。以下の章でこれらに触れている箇所は8章が優先する。

## 1. 用語（画面の文言はこの表に従う）

| 概念 | 画面の語 | 使わない語 |
|---|---|---|
| アプリ名 | RAG用Markdown作成 | Excel帳票AI前処理 |
| 帳票の取り込み | 帳票を取り込む（1ファイル＝1件） | 帳票Excelのアップロード |
| 一覧表の取り込み | 一覧表を取り込む（Excel/CSV・1行＝1件） | 表形式 |
| 帳票の読み取り設定（旧 pattern/テンプレート） | 帳票の種類 | テンプレート、パターン |
| 一覧表の設定（table template） | 一覧表の取り込み設定 | 表テンプレート |
| 見本（sample） | 見本ファイル | サンプルExcel |
| 抽出 | 読み取り | 解析、前処理、抽出 |
| 帳票の確認画面 | 読み取り結果の確認・修正 | 前処理結果、抽出結果 |
| 登録 | 確定（ボタン: 確定してMarkdownを作成） | 登録 |
| 帳票の状態 | 読み取り前 / 確認中 / 確定済み / 修正中 | 解析待ち / 確認待ち / 登録済み |
| 帳票の種類の状態 | 作成中 / 使用中 / 停止中 | 下書き / 有効 / 無効 |
| 一致度 | 「10項目中9項目が見つかりました」 | 信頼度%、確信度 |
| 値の出どころ | 自動で読み取り / AIが入力（要確認） / 手で修正 | AI補完 |
| 原本 | 元のファイル（ボタン: 元のファイルをダウンロード） | 原本を開く |
| AI の列処理 | AI整形 | 前処理 |
| ホームに残っているものの一覧 | 作業中 / ダウンロード待ち | 取り込み履歴、登録文書、文書一覧 |
| まとめて選んだ複数の帳票 | まとめ取り込み（例:「まとめ取り込み 3/12件目」）| バッチ、一括処理 |

ボタンは動詞。1画面に主ボタンは1つ。ステップ名＝画面見出し＝次へ進むボタンの動詞。

## 2. 画面構成とルート

### 2.1 ナビゲーション（base.html）

`ホーム` / `帳票を取り込む` / `一覧表を取り込む` / `設定 ▼`（帳票の種類・一覧表の取り込み設定・AI接続・保存先フォルダ・LightRAGへの入れ方）。右端にヘッダーのAIモデル選択（既存）。

取り込み履歴の画面は無い（3.3 のとおりデータを残さないため、履歴に出せるものが無い）。

### 2.2 ホーム `GET /`
- 大きな入口カード2枚（図＋1文）：「帳票を取り込む（1ファイル＝1件の報告書・記録票）」「一覧表を取り込む（Excel/CSV、1行＝1件の一覧・台帳・集計表）」。
- データを残さないこと（3.3）の案内文。
- 「作業中の帳票」（読み取り前・確認中）「作業中の一覧表」（読み込み前・読み込み中・確認中・確定処理中・失敗）：続きを開く／削除。
- 「ダウンロード待ちの帳票」「ダウンロード待ちの一覧表」（確定済みでまだダウンロードしていないもの）：ダウンロード（押すと消える確認つき）。
  まとめ取り込みは**まとまり1行**にまとめ、zip のボタン（全件確定なら全部、未確定が残っていれば「確定済みN件だけ」）と
  中の帳票を出す。1件だけの `.md` には「この帳票だけがまとまりから消えます」の確認を付ける。
  保存先フォルダ（2.8）が設定されていれば、各行に［フォルダに保存］（POST・確認つき）も出す。行ごとのボタンは小さい普通のボタンにする（1画面に主ボタンは1つ）。
- ダウンロードすると消えるので、この画面が残っているものの一覧そのものになる（取り込み履歴の代わり）。
- 初回（帳票の種類も一覧表の取り込み設定も0件）は「はじめに」3ステップ（見本から種類/設定を作る → 取り込む → Markdownをダウンロード）。

### 2.3 帳票フロー（blueprint `forms`、prefix `/forms`）

ステップ表示：`1 ファイルを選ぶ` → `2 帳票の種類とシートを確認` → `3 読み取り結果の確認・修正` → `4 完了（Markdownをダウンロード）`

| ルート | 内容 |
|---|---|
| `GET /forms/new` | ファイル選択（ドラッグ＆ドロップ可、**複数選択可**）。.xlsx/.xlsm。帳票の種類が0件なら作成へ案内 |
| `POST /forms/upload` | 1ファイル: 保存→ `/forms/<id>/type`。複数ファイル（50件まで）: 1つの**取り込みのまとまり**（`documents.batch_id`）にして先頭の帳票へ。読めなかったファイルはメッセージにして他は進める |
| `GET /forms/<id>/type` | 候補の種類（「10項目中9項目」表示）、シート選択、[AIで種類を推定]。一覧表らしい（見出し行の下に同形行が10行以上）なら「一覧表の取り込みへ」を提案 |
| `POST /forms/<id>/read` | 読み取り→ `/forms/<id>/review` |
| `GET /forms/<id>/review` | 左：元のシートのHTMLプレビュー（セルグリッド、項目にフォーカスで見出しセル・値セルをハイライト）。右：項目（入力欄、状態タグ、警告）。上部チップ「要確認 N / AIが入力 N / 手で修正 N」。[AIで空欄を探す]。Markdownプレビュー（タブ、入力に連動して更新）。主ボタン [確定してMarkdownを作成]。出口は非破壊の「ホームに戻る（作業内容は保存されています）」 |
| `POST /forms/<id>/draft` | 入力内容の途中保存（JSON、fetch。PRGなし・204） |
| `POST /forms/<id>/ai-fill` | 空欄をAIで探す→review へ |
| `POST /forms/<id>/confirm` | 必須欠落があれば止める（「空欄のまま確定」チェックで許可）。confirmed に保存→ `/forms/<id>/done` |
| `GET /forms/<id>/done` | 完了。主ボタン [この帳票のMarkdownをダウンロード（.md）]（押すと消える確認つき）、[次の帳票を取り込む]、[ホームへ]。まとまりの帳票なら進み具合（3/12）・一覧・[次の帳票へ]、全件確定後は [まとめてMarkdownをダウンロード（zip）]、未確定が残っていれば [確定済みN件だけをダウンロード（zip）] |
| `GET /forms/<id>` | 詳細（確定内容、Markdown、元ファイル）。[修正する]→review（修正中） |
| `POST /forms/<id>/discard-changes` | 修正中の変更を破棄し確定済みの版に戻す |
| `GET /forms/<id>/download.md` | ダウンロード。md は **確定済みデータから毎回生成**。**渡したあとその帳票のデータを消す**（3.3） |
| `GET /forms/batches/<batch_id>/download.zip` | まとまりの全 md を zip で渡し、**まとまり全体を消す**。未確定が残っていれば断ってその帳票へ戻す。`?confirmed_only=1` なら確定済みの分だけを zip にして**その分だけ**消す（読めない帳票が1件混ざっても作業が止まらないように） |
| `POST /forms/<id>/save-to-folder` | ダウンロードの代わりに、同じ .md を保存先フォルダ（2.8）に書き、書き終えたらダウンロードと同じくその帳票のデータを消す |
| `POST /forms/batches/<batch_id>/save-to-folder` | まとまりの .md を zip にせず1ファイルずつ保存先フォルダに書き、zip のダウンロードと同じ分を消す（フォームの `confirmed_only=1` で確定済みの分だけ） |
| `GET /forms/<id>/download.json` / `original` | 確認用のダウンロード（消さない。ボタンにも「（消えません）」と書く）。JSON の `source` はファイル名・ハッシュ・シートだけで、`original` への URL は書かない（md をダウンロードすると消えて 404 になるため） |
| `POST /forms/<id>/delete` | 削除（確認文「元のファイルと読み取り結果を削除します。元に戻せません」）。`next=` で戻り先を指定（ホームから使う） |
| `POST /forms/<id>/reread` | 種類/シートを選び直して再読み取り（確定済み・修正中なら「手修正とAI入力が失われます」確認必須） |

### 2.4 一覧表フロー（blueprint `tables`、prefix `/tables`）

ステップ表示：`1 ファイルを選ぶ` → `2 取り込み設定を選ぶ` → `3 表の範囲と見出しを確認` → `4 列の対応づけ`（設定を作る/変える時のみ） → `5 AI整形（任意）` → `6 内容とファイルの確認` → `7 完了（ダウンロード）`

| ルート | 内容 |
|---|---|
| `GET /tables/new` | ファイル選択（.xlsx/.xlsm/.csv/.tsv/.txt） |
| `POST /tables/upload` | 事前チェック（形式・zip上限・文字コード判定）→ import 作成 → `/tables/imports/<id>/source` |
| `GET /tables/imports/<id>/source` | CSV：文字コード/区切り/前置き行の判定結果（変更可）。Excel：シート一覧。取り込み設定の候補（「必須列 5/5 一致」）と「新しい取り込み設定を作る」 |
| `POST /tables/imports/<id>/source` | 選択を保存→ layout |
| `GET /tables/imports/<id>/layout` | 先頭60行の色分けグリッド（見出し=青、データ=白、小計・合計=橙、注記=緑、継続行=水色、除外=灰、非表示/取り消し線=斜線）。見出し行・データ終了行を変更（JSで即時再判定 `POST .../layout/detect` JSON）。クロス集計は縦持ち変換の先頭20行も表示 |
| `POST /tables/imports/<id>/layout` | 範囲を保存→（新規設定なら）columns / （既存設定なら）見出し変更がなければ ai、変更があれば columns（見出しの変更を確認） |
| `GET/POST /tables/imports/<id>/columns` | 列の対応づけ表（元の見出し/値の例/推定型と型エラー率/空欄率/標準キー/表示名/説明/型/単位/役割/空欄＝上と同じ/mdでの扱い/AI整形の対象）。保存は JSON（fetch）。設定名・期間の単位・ファイルのまとめ方もここ |
| `GET /tables/imports/<id>/ai` | AI整形（log 役割の列がある場合のみ。なければスキップして preview）。分割プレビュー（AIなし）→ 試し実行（10行）→ 全件実行（範囲・同時実行数・見積もり・外部送信の確認）→ 進捗（一時停止/再開/中止）。「AI整形をしないで進む」 |
| `POST /tables/imports/<id>/ai/...` | `split-preview`（JSON）, `trial`（1行ずつ, JSON）, `run`, `pause`, `resume`, `cancel` |
| `GET /tables/imports/<id>/preview` | 概要（読込件数・除外件数と理由・エラー/警告・合計の照合）、**作られる md の一覧と中身**（下書きはジョブで作り、待ち画面を出す）、問題一覧（CSV）、データ（100行ずつ）。ダウンロードで消えることの一文。期間の置き換え・投入済みとの差分は持たない（8.1） |
| `POST /tables/imports/<id>/confirm` | エラーが残っていれば止める。ジョブで全 md を作る→ done |
| `GET /tables/imports/<id>/done` | [Markdownをまとめてダウンロード（zip）]（押すと消える確認つき）[正規化CSVをダウンロード]。zip は渡したあと**その取り込みのデータを消す**（3.3）。正規化CSVは zip の「管理用_RAGには入れない」フォルダにも入る |
| `POST /tables/imports/<id>/save-to-folder` | zip の代わりに、`RAG投入用/` と同じ md を保存先フォルダ（2.8）の直下に1ファイルずつ書く。設定でオンなら `管理用_RAGには入れない/` と同じファイルを `<保存先>/_管理用_RAGには入れない/<取り込み名>_<日時>/` に書く。書き終えたら zip のダウンロードと同じくその取り込みのデータを消す |
| `GET /tables/imports/<id>/normalized.csv` | 正規化CSV（消さない。zip にも同じものが入る） |
| `POST /tables/imports/<id>/delete` | 取り込みを削除（`core.purge.purge_table_import`）。戻り先はホーム |
| `GET /api/jobs/<job_id>` | ジョブ進捗（JSON） |

### 2.5 取り込み履歴（作らない）
データを残さない方針（3.3）にしたため、`/history` の画面・`views/history.py`・`templates/history.html` は削除した。
残っているもの（作業中・ダウンロード待ち）の一覧はホーム（2.2）が持ち、確定済み帳票のまとめ zip は
「取り込みのまとまり」（`GET /forms/batches/<batch_id>/download.zip`）が引き継ぐ。

### 2.6 設定（blueprint `settings`、prefix `/settings`）
| ルート | 内容 |
|---|---|
| `/settings/form-types` | 帳票の種類の一覧・作成（見本ファイル→候補確認→読み取りテスト→**使用開始**）・編集（保存しても使用中にしない。作成中のものは読み取りテスト後に[使用開始]）・停止/再開・削除（影響：確定済み帳票N件は残る） |
| `/settings/table-templates` | 一覧表の取り込み設定の一覧・編集・JSON書き出し/読み込み・削除 |
| `/settings/ai` | AI接続（既存）＋[接続テスト]（models.list と 1回の短い chat） |
| `/settings/output` | 保存先フォルダ（2.8）: md の保存先フォルダ（絶対パス。空欄＝使わない）、管理用ファイルもサブフォルダに保存するか、同じ名前があるとき（止める＝既定／上書き）。保存のときに確かめる。[フォルダを確認する]（`POST /settings/output/check`、JSON）で確かめ直す |
| `/settings/lightrag` | LightRAGへの入れ方（サーバー設定の推奨、ファイルの入れ方・入れ替え手順、エンティティ種別YAMLのダウンロード） |

旧ルート（`/documents/*`, `/patterns/*`, `/settings/models`, `/settings/aliases`（名寄せ辞書は 8.1）, `/history/*`（2.5））は削除。`/api/models` は維持。

エラー画面（`templates/errors/`）はすべて日本語で、ホームへのボタンを置く。403 は他サイトからの書き込みを断ったとき（`app._refuse_cross_site_write`）で、画面の JSON 送信（`Accept: application/json`）には同じ理由を JSON で返す。404 は「ダウンロード済みか削除されたか URL 違い」、500 は起動した画面のメッセージを見るよう案内する。

### 2.7 UI共通部品（templates/components/_ui.html のマクロ）
`steps(items, current)`、`card`, `badge(state)`, `chips`, `empty_state`, `confirm_delete_form`, `save_form`（保存先フォルダに保存する POST フォーム）, `progress(job)`, `data_grid(rows)`（セルグリッド）, `file_drop(name, accept)`。CSS はデザイントークン（色・余白・角丸）を `:root` 変数で統一。ライト基調、配色は現行を踏襲しつつ整理。

### 2.8 保存先フォルダ（2026-09-19 追加。`services/output_folder.py`・`views/folder_save.py`）
ダウンロードしたファイルをエクスプローラーで探して LightRAG に入れ直す手間をなくすため、作った md を**利用者が決めたフォルダに直接保存**できる。
ふつうは LightRAG サーバーの `INPUT_DIR`（`WORKSPACE` を使うなら `INPUT_DIR/<ワークスペース名>`）を指定し、保存したあと LightRAG の［スキャン］（`POST /documents/scan`）で取り込む。LightRAG への送信はしない（0章のまま。書くのはファイルだけ）。
- **設定**（`data/output_settings.yaml`。AI接続と同じく `services/settings_store` で書く。設定なので残す）：`folder`（絶対パス。空＝使わない）、`save_admin`（管理用ファイルも保存）、`on_conflict`（`stop`＝保存を止めて知らせる〈既定〉／`overwrite`＝上書き）。
  保存時と［フォルダを確認する］で、絶対パスか・あるか・フォルダか・書き込めるか（一時ファイルを作って消す）を確かめる。エクスプローラーの「パスのコピー」が付ける `"` は外す。アプリ自身のデータの置き場所（`UPLOAD_DIR`・`DATA_DIR`・`TABLES_DIR`）の中は断る（起動時の片付けで消されるため）。
- **ボタン**：フォルダが設定されているときだけ、ダウンロードがある所（帳票の完了・詳細の .md、まとまりの zip と「確定済みN件だけ」、一覧表の完了の zip、ホームのダウンロード待ちの行）に［保存先フォルダに保存］を出す。完了・詳細画面ではこれを主ボタンにし、ダウンロードは次の選択肢として残す。POST なので他サイトからの書き込みを断る仕組み（`app._refuse_cross_site_write`）がそのまま効く。確認文はダウンロードと同じく「このPCのアプリから消える／もう一度保存・ダウンロードできない」（修正中の帳票は変更が入らないことを先に書く）。
- **書くもの**：ダウンロードと**同じ中身**（帳票は `download.md`／まとまりの zip の中身と同じ名前・バイト列。一覧表は `tables/pipeline.build_download_files` を zip と共有）を zip にせず1ファイルずつ。md は保存先の直下。一覧表の管理用ファイル（正規化データ・問題一覧・取込レポート）は `save_admin` がオンのときだけ `<保存先>/_管理用_RAGには入れない/<取り込み名>_<日時>/` に置く。
- **LightRAG がサブフォルダを読まないことの根拠**（LightRAG 1.5.7 `lightrag/api/routers/document_routes.py`）：`POST /documents/scan`（4409〜4412行）は `run_scanning_process` の中で `doc_manager.iter_new_files()`（3785行）だけから候補を集める。`iter_new_files`（1502〜1547行）は `os.scandir(self.input_dir)` で `INPUT_DIR` の**直下だけ**を見て（1530行）、拡張子が読める種類にない物と**ファイルでない物（フォルダ）は飛ばす**（1533〜1536行「`__parsed__` and any other directory is skipped」）。再帰はしない。だからサブフォルダ `_管理用_RAGには入れない/` の CSV は取り込まれない（兄弟フォルダに分ける必要は無い）。
- **書き方**：各ファイルは同じフォルダの一時ファイル `.rag_<乱数>.tmp`（`.tmp` は LightRAG の対応拡張子に無いので、書き途中をスキャンが拾わない）に書いて `fsync` し、`os.replace` で本来の名前にする。名前は `core/naming` で作ったものだが、保存先の直下を指すか（区切り文字・`..`・`:` を含まない、`resolve()` した親が保存先）を必ず確かめる。
- **同じ名前**：LightRAG 1.5.x は文書をファイル名で見分ける（doc_id＝名前の MD5。末尾の分割ヒント `.[…]` は外して比べる: `lightrag/parser/routing.py` 62行・1085〜1099行）。取り込み終えたファイルは `INPUT_DIR/__parsed__/` に移り、そこに同名があれば `_001` などを足す（`lightrag/utils.py` 917〜942行、`document_routes.py` 164行）。同じ名前のまま**スキャン**すると、`classify_scan_file`（`document_routes.py` 3040〜3135行）が既存の行を見て分ける：`PROCESSED` なら取り込まず `__parsed__/<名前>_001.md` などに移すだけ（3101行・3843〜3850行「Skipping already processed file」。文書の一覧には何も出ない）、取り込み待ち・取り込み中の行で同じ物理名なら `RESUME_SAME_PHYSICAL_SOURCE`（3126〜3129行・3897〜3904行）で**前の内容（full_docs）のまま**処理を続け、上書きした中身は使われない。どちらも LightRAG の中身は古いまま。（アップロード・insert の経路では `[DUPLICATE:filename]` になる: `lightrag/pipeline.py` 1340〜1409行。）
  **削除と `delete_file`**：LightRAG の文書の削除は、既定ではファイルを消さない（`DeleteDocRequest.delete_file` の既定は False: 1065〜1068行。WebUI の「アップロードされたファイルも削除」も既定はオフ: `lightrag_webui/src/components/documents/DeleteDocumentsDialog.tsx` 45行）。オンにすると `delete_file_variants_by_file_path`（2153〜2200行、4232〜4236行から呼ぶ）が `INPUT_DIR` と `INPUT_DIR/__parsed__` の**両方**から同じ名前（`__parsed__` では `_001` などを外して比べる）のファイルを消す。
  そこで利用者への案内は、**保存の前に**削除するなら `delete_file=true`（`__parsed__` の古いファイルも消え、アプリの「止める」に引っかからない）、**保存のあとで**削除するなら `delete_file=false`（true だと、いま保存したファイルも消える。アプリにはもうデータが無い）とする（保存の結果の画面・LightRAGへの入れ方 6・設定の画面・README）。
  そこで保存の前に、保存先の直下と `__parsed__/` にある「LightRAG で同じ文書になる名前」（ヒントを外し、大文字・小文字をそろえて比べる。`__parsed__` では末尾の `_001` などが LightRAG の番号か名前の一部（`No._123` など）か分からないので、番号を外した名前と外さない名前の**どちらか**が合えば同じとみなす。`output_folder.lightrag_key`・`find_clashes`）を探す。
  `stop` なら**何も書かずに**名前の一覧と「LightRAG で先に削除してから入れ直す」説明を出す（409）。`overwrite` なら同じ名前のファイルは上書きし、結果の画面に上書きした名前と、LightRAG で同じ文書とみなされる別のファイル（ヒント違い・`__parsed__`）を出す。
- **失敗したとき**：1つでも書けなければ、この回に書いたファイルを消し、上書きしたファイルは退避しておいた元の中身に戻し、作ったサブフォルダも消す。データは消さない（もう一度押せる）。失敗の画面には書けなかったファイルの名前を出す。
  消すときは `core.files.remove_upload` と同じく少し待って何度か試し（読み取り専用なら外してから）、それでも消せなかったファイル（ウイルス対策・同期ソフトが掴んでいたなど）は「消した」とは言わず、失敗の画面にパスを並べて「スキャンの前に消す」よう案内する（画面に出すだけで flash には入れない）。上書きで退避した `.rag_….bak` が消せなかったときも結果の画面に出す。
- **パスの長さ**：Windows の長いパスが無効（`LongPathsEnabled`=0）なら、パス全体で259文字（フォルダを作るときは247文字）まで。保存の前にすべての名前（一時ファイル `.rag_<16桁>.tmp` の長さも含む）を確かめ、超える名前があれば何も書かずに 400 で知らせる。設定の画面と［フォルダを確認する］では、名前に使える文字数が120文字を切るフォルダに注意を出し、確かめ用の一時ファイルも書けない長さなら「パスが長すぎます」と出す（権限の問題とは言わない）。
- **書けるかの確かめ**：一時ファイルを `os.open(O_CREAT|O_EXCL)` で1回だけ作って消す。`tempfile.mkstemp` は使わない（Windows でアクセス権で断られたフォルダだと、`PermissionError` を名前の衝突とみなして試し続け、戻ってこない。保存はロックの中で確かめるので、ほかの保存まで止まる）。
- **管理用のサブフォルダ**：`_管理用_RAGには入れない` がジャンクション・シンボリックリンクなら（あるいは解決した先が保存先の直下でなければ）何も書かずに断る。作った日時のフォルダも、解決した先が保存先の中かを確かめてから書く。
- **消す**：すべて書けたあとにだけ、ダウンロードを渡し終えたときと同じ関数（`purge_documents` / `purge_batch` / `purge_table_import`）で消す（3.3）。同じものを2回押しても重ならないよう、保存と削除はプロセス内のロックで1つずつ行う（2回目は「もう無い」404）。
- **結果の画面**：POST の応答でそのまま出す（リダイレクトしない・`Cache-Control: no-store`）。保存先、md の件数、ファイル名（スクロールする一覧）、上書きした名前、次に LightRAG で行うこと（［スキャン］か `POST /documents/scan`、入れ替えなら削除してから。**この時点の削除は「アップロードされたファイルも削除」をオフで**。上書きした名前か同じ文書とみなす別のファイルがあれば、その注意を目立たせて出す）。ファイル名は flash（セッションクッキー）に入れない（3.3）。
  同じ名前で止めた画面（409）は、`__parsed__` にあるものについては「削除のときに『アップロードされたファイルも削除』を入れる（または `__parsed__` の古いファイルを手で消す）」と案内し、「上書きする」への切り替えは直下にあるときだけ勧める（`__parsed__` の重なりは上書きしても LightRAG に入らないため）。
- **ボタン（ホームのまとまりの中）**：まとまりの中の1件ずつにある「.mdだけ」の横にも、フォルダが設定されていれば「フォルダに保存」を出す（その1件だけがまとまりから消える）。

## 3. データモデル（SQLite `instance/app.db`）

`models/database.py`：接続時に `PRAGMA foreign_keys=ON; journal_mode=WAL; busy_timeout=5000`。`PRAGMA user_version` による連番マイグレーション（`MIGRATIONS: list[Callable[[sqlite3.Connection], None]]`）。バックグラウンドスレッド用に `connect()`（`g` を使わない接続）を公開。

### 3.1 帳票（既存を拡張）
- `patterns`：既存列＋`title_fields TEXT DEFAULT '[]'`（タイトル・ファイル名に使う field_name の配列）、`md_options TEXT DEFAULT '{}'`（`{"domain_context_fields": [...], "omit_person_fields": true}` 等）、`version_no INTEGER DEFAULT 1`（保存ごとに+1）。`status` の値は draft/active/inactive のまま（表示語だけ変更）。
- `pattern_fields`：＋`unit TEXT DEFAULT ''`、`rag_output TEXT DEFAULT 'show'`（show/omit）。`data_type` は string/text/date/number/table（明細表）。`extraction_rule` は `{"direction": "auto"}`、明細表は＋`"columns": [見本で見た列見出し]`（見出しの書き方が違う帳票で、列見出しが似た表を探すのに使う）。探す区画がある項目は＋`"section": "回答"`（区切りの見出しの名前。その区画の中だけでラベルを探す。8.0）。
- 明細表の値（`data_json` の各項目の `value`）：`{"columns": ["品番", "品名", "数量"], "rows": [["PW48-1591", "ベアリング", "2"], ...]}`。行が無ければ null。連番だけの No 列と「なし」だけの行は読まない。
- `documents`：＋`confirmed_json TEXT`、`confirmed_at TEXT`、`title TEXT DEFAULT ''`（一覧の見出し。タイトル項目の値を連結）、`batch_id TEXT DEFAULT ''`・`batch_order INTEGER DEFAULT 0`（まとめ取り込み。同じ `batch_id` の帳票を順に確認し、zip でまとめて渡して一緒に消す）。`data_json` は作業中の値。状態は導出：`data_json IS NULL`→読み取り前、`confirmed_json IS NULL`→確認中、`confirmed_json != data_json`→修正中、それ以外→確定済み。`markdown`・`registered_at`・`status` 列は使わない（互換のため残す。マイグレーションで `registered_at`→`confirmed_at`、`status='registered'` の `data_json`→`confirmed_json` にコピー）。

### 3.2 一覧表
```
table_templates(id PK, name UNIQUE, description, current_version_id, created_at, updated_at)
table_template_versions(id PK, template_id FK, version INTEGER, spec_json TEXT, spec_hash TEXT, used INTEGER DEFAULT 0, note, created_at, UNIQUE(template_id, version))
table_imports(id PK, template_id FK NULL, template_version_id FK NULL, file_name, file_hash, stored_path,
              source_json TEXT,   -- {"kind":"csv|excel","encoding","delimiter","preamble_rows","sheet","header_row","header_rows","data_end_row"}
              period_json TEXT,   -- {"grain":"month|fiscal_year|all","start":"2026-08","end":"2026-08"}
              status TEXT,        -- uploaded / reading / preview / confirming / confirmed / discarded / failed
              stats_json TEXT, issues_path TEXT, rows_path TEXT, job_id INTEGER, created_at, updated_at, confirmed_at)
jobs(id PK, kind TEXT, ref_type TEXT, ref_id INTEGER, status TEXT,  -- queued/running/paused/done/failed/cancelled/interrupted
     params_json, progress_json, message, cancel_requested INTEGER DEFAULT 0, pause_requested INTEGER DEFAULT 0,
     heartbeat_at, created_at, updated_at)
llm_calls(cache_key PK, raw_text, parsed_json, model, params_json, structured_mode, finish_reason, tokens_in, tokens_out, latency_ms, created_at,
          import_id)   -- この応答を払った取り込み（マイグレーション _m7_llm_calls_owner。古いDBの行は NULL）
ai_items(id PK, template_id, import_id, stage_id, row_key, template_version_id, source_hash, context_hash, segments_hash, cache_key,
         status TEXT,   -- pending/ok/flagged/rule_only/error/skipped/outdated/excluded
         result_json, checks_json, override TEXT, attempts INTEGER, error TEXT, job_id, updated_at,
         UNIQUE(import_id, template_id, stage_id, row_key))   -- 取り込みごと（マイグレーション _m6_ai_items_per_import）
```
一覧表の行データはDBに持たない。取り込み中だけ `data/tables/imports/<import_id>/`（`rows.jsonl.gz`・`issues.csv`・`md/`・`preview_md/`・`source_cache.json`）に置き、ダウンロードしたらフォルダごと消す（3.3）。確定済み全行の保存（`state/current.jsonl.gz`）は作らない（8.1）。


### 3.3 データを残さない（保存方針）

サーバー（このPC）の容量を使わないため、**取り込んだデータはダウンロードが終わった時点で消す**。再ダウンロードはできない。

| | 消すもの | いつ |
|---|---|---|
| 帳票 | `uploads/documents/<保存名>`、`documents` の行（と `document_id` を持つ表の行） | `GET /forms/<id>/download.md` を**送り終えたあと**。まとまりは `GET /forms/batches/<batch_id>/download.zip` でまとまり全体（`?confirmed_only=1` なら確定済みの分だけ）。保存先フォルダに保存したとき（`POST …/save-to-folder`）は、**全ファイルを書き終えたあと**に同じ範囲 |
| 一覧表 | `uploads/tables/<保存名>`、`TABLES_DIR/imports/<id>/`（rows.jsonl.gz・issues・控え・md・preview_md）、`table_imports` の行、`jobs`（`ref_type='table_import'`）、`ai_items`（`import_id` がこの取り込みの分）、どの `ai_items` からも参照されなくなった `llm_calls` | `GET /tables/imports/<id>/download.zip` を**送り終えたあと**。保存先フォルダに保存したとき（`POST /tables/imports/<id>/save-to-folder`）は**全ファイルを書き終えたあと** |

- **保存先フォルダへの保存も「渡した」とみなす**（2.8）：保存先フォルダに書いたものはアプリの外のファイルなので、ダウンロードと同じくアプリのデータを消す。
  消すのは全ファイルを書き終えたあとだけ。1つでも書けなかったとき・同じ名前があって止めたときは、その回に書いたファイルを消して（上書きしたものは戻して）、データは残す。
  ダウンロードと違って「受け取られる前に消える」ことは無い（書き終えたことをアプリ自身が確かめてから消す）。保存したファイルはアプリの管理の外で、消すのは利用者（LightRAG が取り込むと `__parsed__/` に移す）。
  残るのは設定 `data/output_settings.yaml`（フォルダのパスと2つの選択。取り込んだデータは入らない）。

- **残すもの**：帳票の種類（`patterns` 系）・一覧表の取り込み設定（`table_templates` 系）・AI接続の設定（`data/model_settings.yaml`・`data/prefs.yaml`）・保存先フォルダの設定（`data/output_settings.yaml`）。これらは設定であってデータではない。
- **履歴は持たない**：何をいつ取り込んだかの1行も残さない。だから取り込み履歴の画面が無い（2.5）。
  番号の続き（`sqlite_sequence`。AUTOINCREMENT の表が「これまでに使った一番大きい番号」＝取り込んだ件数を覚えている）も、
  取り込みの表（`documents`・`table_imports`・`jobs`・`ai_items`）が空になったとき乱数（2^31〜2^52）に置き換える
  （`core/purge.forget_id_counters`。消すたびの後始末 `_shrink` と起動時の片付けで行う）。AUTOINCREMENT は外さない:
  外すと消した番号がすぐ使い回され、開いたままの古い画面（途中保存・ダウンロードのボタン）が新しい取り込みを指して
  書き換えたり消したりする。乱数にすると最初の番号（1〜）とも前回の番号とも重ならない（重なる確率は1回あたり
  件数 / 約4.5×10^15）ので、古い画面の操作は 404 になる。作業中の取り込みが残っている間は番号の続きはそのまま。
  新しいDBは 1 から始まる。
- **消し方**：`core/purge.py`（`purge_documents` / `purge_batch` / `purge_table_import`）。消す表は名前で決め打ちせず、その取り込みを指す列（`document_id` / `import_id`）を持つ表を `sqlite_master` から探す（表が増えても消し残さない）。
- **欠けないダウンロード**：md も zip も全体をメモリに作ってから消す。作れなかったとき（未確定・失敗）は何も消さない。
  消すのは**本文を最後まで送り終えたあと**（`core/purge.purge_after_send`）。通信が切れた・ブラウザを閉じたなどで
  本文を渡しきれなかったときは消さないので、もう一度ダウンロードできる。
  「送り終えた」は、サーバー（waitress）の送信の溜めに本文の最後を入れ終えたことで、ブラウザが受け取ったことではない。
  溜めの上限は `app.OUTBUF_HIGH_WATERMARK = 16KB`（waitress の既定 16MB のままだと、16MB 未満の zip は受け取られる前に
  消えていた）。この上限と OS の送受信の溜めの分（2026-09-19 の測定で合わせて約48KB）は受け取られる前に送り終えた
  ことになるので、次のとおりになる。
  - 約48KB より大きい md・zip（一覧表の zip など）: 最後の約48KB より前で切れたら消えない。
  - 約48KB 以下の md・zip（帳票1件の md はほぼこれ）と、大きいファイルの最後の約48KB で切れた場合: 消える。
    取り戻すには、手元の元の Excel をもう一度アップロードして読み取り・確認し直す（一覧表は取り込み設定が残るので
    同じ設定で取り込み直せる。AI整形の結果は消えるので、AI を使う列は再度課金される）。
  ブラウザが受け取ったことをサーバーは知る方法が無い（画面からの受け取り確認は作っていない。8.0）。
  Range 付きの要求（ダウンロードマネージャー・途中からの再開）にも全体を 200 で返し（`send_file(conditional=False)`）、
  `purge_after_send` は 200 以外（206・304）の応答では消さない（一部だけ渡して消すことが無いように）。
- **消した中身を残さない**：接続時に `PRAGMA secure_delete = ON`（消した行の中身をその場でゼロ埋め）、消したあとに
  `PRAGMA wal_checkpoint(TRUNCATE)` と `VACUUM`。これが無いと、行を消しても `instance/app.db` の解放ページや
  `app.db-wal` から帳票の値・設備名・人名が平文で読めてしまう。
- **消せなかったファイル**：`remove_upload` は他のプロセスに掴まれていて消せないとき（ウイルス対策のスキャン、Excel で
  開いたまま、同期ソフト）、中身を 0 バイトに切り詰めて `False` を返す。名前は残っても元のデータは残さない。
- **AI整形の控え**：`ai_items`・`llm_calls` も取り込んだ内容そのものなので一緒に消す。消すのは `ai_items.import_id` が
  その取り込みの分と、その取り込みが払った `llm_calls`（`llm_calls.import_id`。どの `ai_items` からも参照されていないもの）
  だけで、同じ設定で作業中の別の取り込みの結果と応答キャッシュは残す（巻き添えの再課金を避ける）。
  同じ表を取り込み直すと AI に再度課金される（承知のうえ）。
  取り込み設定を消したときは、その設定の `ai_items` と、それで参照されなくなった `llm_calls` をその場で消す
  （`tables/store.delete_template` → `core/purge.sweep_orphan_ai`。起動時の片付けでも同じ掃除をする）。
- **控えを書き戻さない**：`tables/source_cache.ImportSource` はフォルダを `__init__` でだけ作る。消したあとに
  別のタブの画面処理が控えを書いても、`imports/<id>/` は復活しない。
- **消し損ね**：起動時に、DBから参照されていないアップロードファイルと `imports/<id>/` フォルダ、もう無い取り込みを指す `ai_items`（と、それで参照されなくなった `llm_calls`）を片付ける（`app._cleanup_leftovers` → `core/files.remove_orphan_uploads` / `remove_orphan_import_dirs`、`core/purge.delete_orphan_ai_items`）。`purge_table_import` も同じ掃除をする。作業中のものは DB に行があるので消さない。`purge_table_import` で `imports/<id>/` を消し切れなかったとき（Windows で掴まれていた）は、残ったファイルの中身を 0 バイトに切り詰め（`remove_upload` と同じ。名前は残っても読み込んだ行・md は残さない）、ログに残し、次の起動時に片付く。アップロードのファイルや `imports/<id>/` を消し切れなかったことは `core/purge.purge_incomplete()` で分かり、保存先フォルダへの保存の結果の画面は「消しました」と言わずに「消し切れませんでした」と知らせる。帳票・一覧表の［削除］の案内も「一部のファイルは使用中で消し切れず、中身を空にしました。次の起動時に片付きます」にする。
- **消したあとの AI整形**：動いている AI整形のジョブは、取り込みの行が消えたら新しい呼び出しを出さず、結果も書かずに止まる（`aiproc/runner._import_gone`。ジョブの行が消えたときは `JobContext` が中止扱いにする）。一覧表の削除・確定・zip のダウンロードは AI整形の実行中は受け付けない。
- **他サイトからのダウンロード（＝削除）を断る**：消すダウンロード（`forms.download_md` / `forms.download_batch` / `tables.download_zip`）は GET でも、`Sec-Fetch-Site` が same-origin / none 以外、他サイトの `Origin`、（`Sec-Fetch-Site` が無いときは）他サイトの `Referer` なら 403（`views.PURGING_ENDPOINTS`）。
- **読み取れないアップロード**：事前チェックや読み込みで思わぬ例外が出ても、アップロードしたファイルは消してから落とす（帳票・一覧表とも）。
- **画面の知らせ**：ダウンロードのボタンには確認ダイアログ（`data-confirm`）、完了・確認画面には「ダウンロードするとこのPCから消える／もう一度ダウンロードできない」の一文を出す。
  保存先フォルダが設定されているときは［保存先フォルダに保存］にも同じ意味の確認（`views.forms.save_confirm` / `batch_save_confirm`、`views.tables.SAVE_TO_FOLDER_CONFIRM`）を付け、完了・詳細画面とホームの案内に「保存したときも同じ」を足す。
  修正中（確定済みの版あり）の帳票は、確定し直していない変更が入らずに消えることを確認文の先頭に書く（確認・完了画面とホームの .md / zip ボタン。`views.forms.delete_confirm` / `batch_zip_confirm`）。
  消えないボタン（帳票の `download.json`・`original`）には「（消えません）」と書く。ホームではまとめ取り込みを1行にまとめ、
  1件だけダウンロードするとまとまりが崩れることを押す前に知らせる。
- **画面のメッセージにファイル名を出さない**：`flash` は署名付きセッションクッキーとしてブラウザに残るので、
  取引先名や「社外秘」を含みうるファイル名は載せない（サーバー側を消してもブラウザに残るため）。
- **残ると分かっていて残すもの**：帳票の種類の見本ファイル（`uploads/samples/<uuid>.xlsx`＋`pattern_samples`。読み取りテストと
  項目の見直しに使うため。画面に「このPCに残り続けます」と書き、1件ずつ削除できる）、一覧表の取り込み設定の名前
  （初期値はファイル名にしない。空欄＋プレースホルダ）。
- **新しい表**：取り込みを指す列は必ず `document_id` / `import_id` という名前にする（purge の探索に乗せるため）。
  使わない表は置かない（`table_outputs`・`table_downloads`・`table_template_samples`・`alias_entries` はマイグレーション `_m5` で削除）。

## 4. モジュール構成と担当（並行実装の境界）

凡例：WP＝作業パッケージ。**各WPは自分の担当パス以外を編集しない**。他WPが必要な変更は完了報告に「統合メモ」として書く。

| WP | 担当パス | 内容 |
|---|---|---|
| WP-core | `models/database.py`, `core/__init__.py`, `core/jobs.py`, `core/files.py`, `core/naming.py`, `core/mdtext.py`, `tests/test_core_*.py` | DB（全スキーマ・マイグレーション・WAL）、ジョブ実行、アップロード保存と事前チェック、安全なファイル名、Markdown テキスト処理 |
| WP-read | `tables/__init__.py`, `tables/source.py`, `tables/csv_source.py`, `tables/excel_source.py`, `tables/detect.py`, `tables/dictionary.py`, `tables/mapping.py`, `tests/test_tables_read*.py` | 表ソース（CSV/Excel）、見出し帯・行分類・種類判定、標準キー辞書、列の対応づけ候補 |
| WP-pipe | `tables/spec.py`, `tables/normalize.py`, `tables/checks.py`, `tables/markdown.py`, `tables/summaries.py`, `tables/outputs.py`, `tables/pipeline.py`, `tables/store.py`, `tests/test_tables_pipe*.py` | 取り込み設定の仕様、正規化、チェック、md生成（記録・集計・データセット説明）、zip、全体の実行関数、DBアクセス（`tables/state.py` は 8.1 で外した） |
| WP-log | `logproc/*.py`, `tests/test_logproc*.py` | 追記ログの分割、日時解決、記入者、識別子・数量・予定句、マスク、用語集、時系列の描画 |
| WP-ai | `aiproc/*.py`, `services/llm.py`（ジョブ用呼び出し口の追加のみ。既存関数の挙動は変えない）, `tests/test_aiproc*.py`, `tests/fake_servers.py`（拡張のみ） | 構造化出力の方式判定、プロンプト生成、照合、キャッシュ、AIジョブ、custom 段 |
| WP-forms | `excel/*`, `pattern/*`, `export/formats.py`, `services/ai_assist.py`, `tests/test_extraction.py`, `tests/test_forms_md.py` | 帳票の md 改善（タイトル・ファイル名・定型文削減・値の NFKC・単位・出さない項目）、種類定義の拡張、一覧表らしさ判定関数 |
| WP-shell | `app.py`, `config.py`, `templates/base.html`, `templates/components/*`, `templates/home.html`, `templates/settings/*`, `templates/errors/*`, `static/style.css`, `static/app.js`, `views/__init__.py`, `views/home.py`, `views/settings.py`, `views/files.py`（削除し core/files へ移行）| レイアウト・デザイン・ナビ・ホーム（作業中／ダウンロード待ち）・設定（AI接続・名寄せ辞書・LightRAG案内・一覧表設定の一覧）・エラー画面、blueprint 登録 |
| WP-formsui | `views/forms.py`, `views/form_types.py`, `templates/forms/*`, `templates/form_types/*`, `static/review.js`, `tests/test_forms_flow.py` | 帳票フロー画面と帳票の種類の管理画面 |
| WP-tablesui | `views/tables.py`, `templates/tables/*`, `static/tables.js`, `tests/test_tables_flow.py` | 一覧表フロー画面・取り込み設定の編集画面・出力画面 |
| WP-samples | `scripts/samples/*`（追記のみ）, `samples/`（生成物） | T1 に「対応内容」追記ログ列を追加（書き方の揺れを再現）など |

依存：WP-core・WP-read・WP-log・WP-forms は並行（第1波）。WP-pipe・WP-ai は第1波の後（第2波）。WP-shell は第1波と並行可（テンプレートのみ）。WP-formsui・WP-tablesui は第2波の後（第3波）。

## 5. 主要インターフェース

### 5.1 core
```python
# core/files.py
class UploadError(Exception): ...
@dataclass class StoredFile: stored_path: str; file_name: str; file_hash: str; size: int
def save_upload(storage, subdir: str, allowed: set[str], max_bytes: int) -> StoredFile   # 分割読みで sha256
def precheck_excel(path, max_cells=None, max_merged=None, max_rows=None) -> None   # OLE(D0CF11E0)=パスワード付き/xls、xl/workbook.bin=xlsb、Strict名前空間、zip展開上限(合計500MB/1パーツ200MB/圧縮率100。圧縮率はブック全体でも見る＝小さいパーツを並べて展開させない)、
                                      # 結合セルの面積の合計(200万セル。帳票・見本は FORM_MAX_MERGED_CELLS=20万。通常モードで開くと結合1つごとに全セルをたどるため)、セル数（<c> の数。帳票・見本は EXCEL_MAX_CELLS=50万、省略時 100万）、
                                      # ハイパーリンク・コメントの範囲（ref）の面積の合計(ブック全体で MAX_LINKED_CELLS=5万。A1:XFD1048576 で固まらないように)、
                                      # <row> の数(max_rows。省略時は max_merged と同じ。帳票・見本は20万行、一覧表は200万行)、
                                      # 図形(描画パーツの図形数・大きさ×参照するシート数)、シートが1つも無いブック、壊れた圧縮データ(zlib.error/EOFError) → UploadError(日本語。ファイル名・例外の種類名は入れない)
                                      # 上の結合・セル・行・リンクの数え方: 1つのパーツを複数の <sheet> が指していると openpyxl はその回数だけ読み直すので、参照される回数を掛けて数える
def upload_path(stored_path) -> Path;  def remove_upload(stored_path) -> bool   # 消せたら True。掴まれて消せないときは中身を0バイトにして False
def remove_orphan_uploads(upload_dir, known: set[str]) -> int;  def remove_orphan_import_dirs(tables_dir, import_ids: set[int]) -> int

# core/purge.py（3.3 データを残さない）
def purge_documents(doc_ids) -> int;  def purge_batch(batch_id: str) -> int;  def purge_table_import(import_id: int) -> int   # 戻り値は消した行数
def purge_after_send(response, fn, *args) -> response   # 本文を最後まで送り終えたときだけ消す（途中で切れたら消さない）
# 消す表は sqlite_master から document_id / import_id 列を持つ表を探して決める（新しい表はこの列名にする）
# 消したあと: PRAGMA wal_checkpoint(TRUNCATE) と VACUUM（中身のゼロ埋めは接続時の PRAGMA secure_delete = ON）
def forget_id_counters(db) -> int   # 空になった取り込みの表の sqlite_sequence を乱数にする（番号の続きを履歴にしない）

# core/naming.py
def safe_filename_part(text: str, max_len: int = 60) -> str   # NFKC、\ / : * ? " < > | 制御文字 空白 '[' ']' を _ に、'.[' を除去、前後の . _ を除去
def md_filename(parts: list[str], hint: str | None = None) -> str   # "_".join(safe parts) + (f".[{hint}]" if hint) + ".md"
def hint_chunk_tokens(hint: str | None = None) -> int   # ヒントの chunk_ts（記録の大きさの上限の元）
LIGHTRAG_HINT_RECORDS = "legacy-R(chunk_ts=1500,chunk_ol=0)"

# core/mdtext.py
def nfkc_value(text) -> str            # NFKC＋連続空白の畳み込み（改行は保持）。日本語文字間の半角空白は残す
def escape_md_line(line: str) -> str   # 行頭 # - * + > 数字. と、行全体が --- === の行、``` をエスケープ
def md_bullet(label: str, value) -> list[str]   # 単一行: ["- label: value"]、複数行: ["- label:", "  行1", "  行2"]（空行は出さない）
def estimate_tokens(text: str) -> int  # 非ASCII 1文字=1.1、ASCII の記号 1文字=1、数字のまとまり1つ=1＋3桁ごとに1、ほかの ASCII 2文字=1（実トークン以上になる見積もり）
def join_blocks(blocks: list[list[str]]) -> str   # ブロック間に空行1つ、末尾改行1つ、LF

# core/jobs.py
class JobContext: job_id: int; def progress(self, **kw); def heartbeat(self); def should_stop(self) -> bool; def wait_if_paused(self) -> bool; def check_cancel(self) -> None
class JobError(Exception): ...         # 日本語のメッセージをそのまま画面に出す（aiproc.runner.AIJobError と tables.pipeline.PipelineError が継承）
def error_message(exc) -> str          # 想定外の例外は Python の例外名を出さず日本語の案内にする（内容はログへ）
def start_job(kind: str, ref_type: str, ref_id: int, fn: Callable[[JobContext], dict | None], params: dict | None = None) -> int
def get_job(job_id) -> dict | None;  def request_pause(job_id); def request_resume(job_id); def request_cancel(job_id)
def recover_interrupted() -> None      # 起動時: running/queued で heartbeat が2分以上古い → interrupted
# 実装: 単一ワーカースレッド＋キュー。fn 内で DB を使うときは database.connect()。app_context は start_job 時に app を捕まえて with app.app_context() で実行
```

### 5.2 tables（読み取り）
```python
# tables/source.py
@dataclass class CellInfo: value: object; text: str; number_format: str | None = None; bold: bool = False; strike: bool = False; fill: bool = False; merged_anchor: tuple[int,int] | None = None
@dataclass class SourceRow: index: int (1始まり); cells: list[CellInfo]; hidden: bool | None = None
@dataclass class SheetInfo: name: str; hidden: bool; max_row: int; max_col: int; table_ranges: list[str]; date1904: bool = False
class TableSource(Protocol):
    kind: str   # "csv" | "excel"
    def sheets(self) -> list[SheetInfo]
    def rows(self, sheet: str, start: int = 1, limit: int | None = None) -> Iterator[SourceRow]
def open_source(path: Path, file_name: str, options: dict | None = None) -> TableSource   # 拡張子で振り分け

# tables/csv_source.py
@dataclass class CsvSniff: encoding: str; bom: bool; delimiter: str; preamble_rows: int; header_row: int; trailer_rows: int; confidence: float; warnings: list[str]
def sniff_csv(path) -> CsvSniff          # BOM→utf-8厳密→cp932→shift_jis_2004→utf-16。区切りは「最頻列数に一致する行の割合」優先
class CsvSource: ...                     # options: encoding, delimiter, errors("strict"|"replace〓")。csv.reader(newline='')、NUL除去、="00123" の中身取り出し

# tables/excel_source.py
class ExcelSource: ...                   # openpyxl data_only=True（通常モード）。結合セルは merged_anchor。行 hidden、font.strike、bold、fill。セル数上限（既定50万）で UploadError

# tables/detect.py
@dataclass class RowClass: index: int; kind: str  # title/header/data/continuation/subtotal/note/blank/excluded
                    reason: str = ""
@dataclass class LayoutGuess: sheet: str; table_kind: str  # list/crosstab/form_like/unknown
    header_rows: list[int]; data_start: int; data_end: int; headers: list[str]  # 2段は "上_下" で結合、結合セルは右へ埋める
    row_classes: list[RowClass]  # 先頭60行＋データ範囲の要約
    confidence: float; warnings: list[str]
def guess_layout(source, sheet, anchors: list[str] | None = None, header_row: int | None = None, data_end: int | None = None) -> LayoutGuess
def classify_rows(source, sheet, layout, key_columns: list[int] | None = None) -> Iterator[tuple[SourceRow, RowClass]]
def looks_like_list(source, sheet) -> bool   # 帳票フローから「一覧表らしい」判定にも使う
def split_header_unit(header: str) -> tuple[str, str]   # "停止時間(分)" → ("停止時間","分")

# tables/dictionary.py  … 表用の標準キー辞書（帳票の pattern/dictionary.py は変更しない）
STANDARD_COLUMNS: list[StdColumn(key, display, type, role, synonyms)]
# 例: record_no(code,key) occurred_at(datetime,date) equipment_id(code,entity) equipment_name(string,entity_label) line process
#     failure_category(enum,category) severity symptom(text) cause(text) action(text,log候補) response_log(text,log) downtime(number,measure,分) work_hours cost status worker(string,person) part_name quantity unit_price

# tables/mapping.py
@dataclass class ColumnSuggestion: index: int; header: str; unit: str; key: str | None; display: str; type: str; role: str; examples: list[str]; type_error_rate: float; blank_rate: float; md: str  # body/attribute/omit
    fill_down_blank: bool = False; log: bool = False; matched_by: str = ""  # template/dictionary/similar/none
def suggest_columns(headers, sample_rows, template_spec=None) -> list[ColumnSuggestion]
def match_templates(headers: list[str], sheet_or_file_name: str, specs: list[TableSpec]) -> list[tuple[TableSpec, int, int]]  # (spec, 一致した必須列数, 必須列数)
```

### 5.3 tables（パイプライン）
```python
# tables/spec.py … 取り込み設定（JSON で保存。dataclass ⇔ dict、validate、spec_hash）
@dataclass class ColumnSpec: key: str; display: str; headers: list[str]; type: str   # code/string/text/date/datetime/time/number/enum/status
    role: str = "attribute"   # key/date/entity/entity_label/category/measure/text/log/person/attribute
    unit: str = ""; unit_conversions: dict = {}; required: bool = False; md: str = "attribute"  # body/attribute/omit
    fill_down_blank: bool = False; normalize: list[str] = ["nfkc"]; alias_dictionary: str | None = None
    allowed: list[str] = []; value_map: dict = {}; description: str = ""
@dataclass class LogStageSpec: column: str; enabled_ai: bool = False; context_columns: list[str]; people: list[dict]; groups: list[str]; glossary: dict; entry_types: list[str]; instruction: str = ""; incident: bool = True; run_if: dict; limits: dict
@dataclass class CustomStageSpec: id: str; inputs: list[str]; prompt: str; output_type: str  # text/choice
    choices: list[str] = []; max_chars: int = 80; fallback: str = "不明"; target_key: str = ""; quote_required: bool = True
@dataclass class SummarySpec: id: str  # entity_fiscal_year / month
    metrics: list[str] = ["count"]  # count, sum:<key>, avg:<key>, max:<key>
    top_n: int = 5
@dataclass class TableSpec:
    name: str; description: str = ""; file_types: list[str]; name_patterns: list[str]
    kind: str = "list"   # list / crosstab
    header: {anchors, rows(1|2)}   # 見出しを探す行数・データの終わりの条件・集計行の語は detect.py が決める（設定では変えない）
    exclude: {hidden_rows:"exclude_with_warning", strike_rows:"exclude_with_warning"}
    continuation_rows: str = "merge_into_previous"   # or "keep"
    na_tokens: list[str]; fiscal_year_start_month: int = 4
    columns: list[ColumnSpec]
    crosstab: {id_columns: [...], value_key, value_display, unit, ignore_headers:[年計,合計,平均], blank_as:"no_data"} | None
    record: {key: [..], fallback_key: [occurred_at, equipment_id, "symptom:20"]}
    period: {grain: "month", date_column: "occurred_at"}
    log_stage: LogStageSpec | None; custom_stages: list[CustomStageSpec]
    markdown: {file_prefix, group_by: "month" | "entity_month", max_records_per_file: 300, lightrag_hint: True, dataset_card: True,
               records: True, dedupe_timeline: True, summaries: [SummarySpec], title_columns: [...], omit_person: True}
    # lightrag_hint の既定はオン（新しい設定だけ。保存済みの設定は JSON の値をそのまま使う）。dedupe_timeline は 6.2 参照
    checks: {type_error_rate: {warn: 0.02, block: 0.10}}
def spec_from_dict(d) -> TableSpec; def spec_to_dict(spec) -> dict; def spec_hash(spec) -> str; def validate_spec(spec) -> list[str]

# tables/normalize.py
@dataclass class RecordRow: key: str; values: dict[str, object]; originals: dict[str, str]; source: {"file","sheet","row"}; warnings: list[str]
def read_records(source, source_opts: dict, layout: LayoutGuess, spec: TableSpec, aliases: AliasLookup, on_progress=None) -> tuple[list[RecordRow], list[Issue], ImportStats]
# 型変換は excel/text.py を拡張して使う（和暦 令和/平成/昭和、年なし M/D を年度で補完、8桁日付、6桁時刻、△▲・末尾マイナス、桁区切り、%、timedelta→分、単位換算）

# tables/checks.py
@dataclass class Issue: level: str  # error/warning
    code: str; message: str; row: int | None = None; column: str | None = None
def run_checks(records, spec, stats, period) -> list[Issue]

# tables/state.py は作っていない（期間の置き換え・取り消しは 8.1 で外した）

# tables/markdown.py
@dataclass class MdFile: name: str; text: str; kind: str  # dataset/records/summary
def render_all(spec, records: list[dict], ai_results: dict[str, dict], meta: dict) -> list[MdFile]   # 決定的（同じ入力→同じバイト列）
# tables/summaries.py: entity_fiscal_year, month の集計（コードで計算）
# tables/outputs.py（投入済みとの差分は持たない。8.1）
def normalized_csv(spec, records) -> bytes; def issues_csv(issues) -> bytes; def report_csv(items) -> bytes
def build_zip(md_files: list[tuple[str, bytes]], extras: dict[str, bytes] | None) -> bytes   # RAG投入用/ と 管理用_RAGには入れない/（正規化データ.csv, 問題一覧.csv, 取込レポート.csv）
# tables/pipeline.py … ジョブ本体
def run_read(ctx, import_id) -> dict      # 読込→rows.jsonl.gz, issues.csv, stats
def run_preview(ctx, import_id) -> dict   # 確認画面用の md の下書き（preview_md/）
def run_render(ctx, import_id) -> dict    # 確定→その取り込みの全 md を md/ に作る
def build_download(import_id, imp, spec) -> bytes   # zip 全体をバイト列で返す（遅延生成にしない。3.3）
```

### 5.4 logproc（ルールのみ・純粋関数）
```python
@dataclass class Segment: id: str  # s1..
    raw: str; body: str; start: int; end: int
    when: WhenInfo | None; author: AuthorInfo | None
    identifiers: list[str]; quantities: list[str]; plans: list[str]; marks: list[str]  # "signature","header_cell","email" 等
@dataclass class WhenInfo: text: str; date: str | None; time: str | None; date_to: str | None; shift: str | None; estimated: bool; note: str
@dataclass class AuthorInfo: raw: str; name: str | None; estimated: bool; note: str
@dataclass class LogParse: segments: list[Segment]; order: str  # asc/desc/unknown
    kind: str  # log / header_cell / single / empty
    warnings: list[str]
def parse_log(text: str, base_date: date | None, people: PeopleIndex, options: SplitOptions) -> LogParse
def mask_text(text, rules: list[str]) -> tuple[str, list[MaskSpan]]           # 電話番号・メール（任意で人名・金額）
def apply_glossary(text, glossary: dict, protected_spans) -> str
def render_timeline(parse: LogParse, entity_label: str, types: dict[str, list[str]] | None = None) -> list[str]   # "1. 2024-04-01 10:00［連絡・初動］田中｜搬送ロボット2号機: 本文"
```

### 5.5 aiproc
```python
# services/llm.py に追加（既存 ask_json は維持）
def job_client_settings() -> dict   # base_url, api_key, model, 固定（ジョブ開始時）。fingerprint（キー除く）
def chat_raw(settings, messages, response_format=None, max_tokens=None, timeout=None) -> ChatResult(text, finish_reason, tokens_in, tokens_out, latency_ms, headers)  # timeout は1回ごとの上限（行の残り時間）。クライアントは設定のタイムアウトで作ったものを使い回す
def detect_structured_mode(settings) -> str   # json_schema / json_object / prompt_only（メモリキャッシュ）

# aiproc/prompts.py: build_log_messages(parse, context, spec) -> list[dict]; build_custom_messages(stage, inputs) -> list[dict]
# aiproc/verify.py: verify_log_result(result, parse, sent_text) -> VerifyReport(ok_items, failed_items, issues[level, path, message])
#   チェック: 構造/選択肢、セグメントIDを1回ずつ（隣接結合のみ）、引用(q, *_q)が根拠セグメントの部分文字列、識別子の完全一致、数量・回数、
#            出力に日付・時刻・相対日がない（引用内を除く）、人名（人物一覧・〇〇さん）がない、完了・否定語（済/未/予定/待ち/手配/なし/OK/NG/完了）の反転・欠落、推量語の消失
# aiproc/cache.py: cache_key(messages, model, params, schema) ; get/put（即時コミット）
# aiproc/runner.py: run_ai_job(ctx, import_id, scope, concurrency) ; trial_row(import_id, row_key) -> dict（1行同期）
# aiproc/estimate.py: estimate(import_id, trial_stats) -> {"rows", "calls", "tokens_in", "tokens_out", "minutes", "duration_text"}  # duration_text: 「1分未満」「約N分」
```
AI の出力スキーマ（keep）：`{"entries":[{"id","segs":[...],"t":[種別...]}], "incident":{"root_cause":{"q","v","certainty","src"}, "temporary_actions":[{"v","src"}], "permanent_actions":[{"v","src"}], "parts":[{"name","model","qty_q","src"}], "recurrence":{"v","count_q","src"}, "final_state":{"v","src"}}, "summary":[...](長いレコードのみ)}`。日付・時刻・人名・数字は書かせない（数量は `_q` 引用）。照合に落ちた項目は出さない（項目単位）。1回だけ再依頼。

## 6. Markdown 出力仕様（LightRAG 調査に基づく）

共通規則（`core/mdtext.py` で実装）：
- UTF-8（BOMなし）・LF・本文に生成日時や内部IDを書かない（決定的）。
- **1レコード（帳票1件／表1行）の本文だけで意味が通る**：種別、識別番号、設備名（設備番号）、日付（ISO＋「2026年8月」）を本文に書く。
- **レコード内に空行を入れない**。レコード（見出しブロック）間は空行1つ。複数行の値は2文字下げの連続行。
- **パイプ表は使わない**。`- 項目: 値`。
- 値は NFKC＋空白の畳み込み。ただし**囲み文字（丸数字 ①、丸英字 Ⓐ、丸カナ ㋐ など）は原文どおり残す**（`core/mdtext.nfkc_keep_enclosed`。①→1 だと番号と本文の区切りが消える）。日付・数値の解析、NA 判定、コードの突き合わせなど**比較用の正規化は従来どおり NFKC のみ**。設備番号・設備名は名寄せ辞書で正式表記に。数値は単位付き（`停止時間: 95分`）。
- 全ファイル共通の定型文を入れない。出典は末尾1行 `- 出典: 元ファイル名（識別番号）`。
- 人名（person 役割・担当者）は既定で出さない（設定で出せる）。
- ファイル名は `core/naming.md_filename`。論理文書に対して安定・一意。`.[` `]` は除去。LightRAG ヒント（`.[legacy-R(chunk_ts=1500,chunk_ol=0)]`）は一覧表の取り込み設定で付ける（**新しい設定の既定はオン**・画面では推奨と表示。保存済みの設定は値をそのまま使う。帳票には付けない）。
- 推定トークン数（`core/mdtext.estimate_tokens`）は**実トークン以上**になる式（非ASCII 1文字=1.1、ASCII の記号 1文字=1、数字のまとまり1つ=1＋3桁ごとに1、ほかの ASCII 2文字=1。o200k_base は数字を3桁ずつ・記号を1文字ずつ区切るため、日時・品番の多い記録も実トークンを下回らない。まれな漢字（髙・﨑 など）だけは1文字あたり実トークンの方が多い）。記録の上限・見出しへの識別子付与の判定に使う。

### 6.1 帳票
```
# 設備修理報告書 R2026-00123｜CMP研磨装置1号機（CMP-101）｜2026-09-14

- 帳票の種類: 設備修理報告書
- 報告番号: R2026-00123
- 設備: CMP研磨装置1号機（CMP-101）
- 発生日: 2026-09-14（2026年9月）
- 作業時間: 2.5時間

## 故障内容（R2026-00123／CMP-101 CMP研磨装置1号機／2026-09-14）
ウェーハ搬送時に搬送アームが停止した。
装置画面に「Robot Position Error」が表示された。

## 原因（R2026-00123／CMP-101 CMP研磨装置1号機／2026-09-14）
...

- 添付画像: 2枚
- 出典: 修理報告書_標準.xlsx（報告番号 R2026-00123）
```
- タイトル＝種類名＋タイトル項目（`patterns.title_fields`、未設定なら識別らしい項目：report_id/equipment/occurred_date の辞書キー順）の値。
- 長文項目の見出しに識別子を入れるのは、推定トークン数が 1,200（LightRAG の既定の固定窓＝1帳票が2断片に割れうる大きさ）を超える帳票だけ（それ以外は `## 故障内容`）。
- ファイル名：`{種類名}_{タイトル項目値...}.md`。タイトル項目に報告番号が入らないときは末尾に `{file_hash先頭8}` を足す（タイトル項目が空なら `{種類名}_{file_hash先頭8}.md`）。
- タイトル項目に報告番号も設備も入らないときは、元ファイル名（拡張子なし）をタイトルの末尾に足す。
- 値（正規化後20文字以内・1行）が、その種類の他の欄の候補ラベル・表示名・明細表の列見出しと一致する項目は、読み取り誤りとみなして Markdown に出さない（タイトル・ファイル名にも使わない。JSON には残す。手修正した値は対象外）。
- AI入力の値には（AI入力）、手修正は印なし（確定時に人が確認済みのため）。
- 画像のセル座標・種類の版・DBの文書IDは出さない（JSON側に残す）。
- 明細表（型「明細表」の項目）は長文項目と同じく `## 見出し` の下に1行1明細で書く：`- 品番: PW48-1591／品名: ベアリング／数量: 2`。空のセルは書かない。合計行は `- 合計: 投入数: 50枚／…`。パイプ表は使わない。連番だけの `No` 列は書かない（読み取った表・手で入れた行のどちらも）。
  - 見出しに識別子を入れる帳票では、明細表1節の推定トークン数が 800 を超えたところで `## {見出し}（続き）（{識別子}）` に分ける（分けないと、明細表の行だけで埋まった断片に識別番号も設備名も入らない）。
  - 読み取り：探す見出し（「■ 交換部品」「使用部品」など）の下2行以内、または縦に結合した見出しの右にある列見出しの行から、空行・列見出しと同じ色のセル・表の左端の色付きの項目欄・合計行まで。縦結合のデータセルは各行に同じ値。見出しが見つからなければ、保存した列見出しと半分以上同じ表を探す。
  - 見本からの候補：列見出しの上（または左）に見出しのある表で、行を足せる形（2行以上／No だけの空き行／「■」「1.」の見出し／縦結合の見出しに余りの行）のものを明細表の項目にし、その列見出しは1つの値の項目にしない。「影響｜停止時間｜影響ロット」のような見出しの行＋値の行1つは項目の並びとして読む。列見出しが半分以上同じ表は、見本ごとに見出しの書き方が違っても1つの候補にまとめる。
  - **積み重なった列見出し（2026-09-19 の利用者の判断）**：読み取った範囲のすぐ下（2行以内）に、同じ列位置・同じ塗りつぶし色の列見出しの行が続くときは、その組も同じ明細表として読み、1つの値にまとめる（`excel/tables.stacked_tables` → `merge_table_values`）。組の数は N 組まで同じ扱い（特定の帳票に合わせた作りにしない）。
    列見出しは出てきた順に並べ、同じ見出しは同じ列にそろえる。組ごとに見出しが変わる特性要因図（人｜機械／材料｜方法／測定｜環境）は「人｜機械｜材料｜方法｜測定｜環境」の6列になり、各行は自分の組の列だけ埋まる。見出しが同じまま繰り返される様式（ページごとに見出しを書く表）は行が増えるだけになる。
    Markdown は今までどおり1行1明細で、**その行の空でないセルの列見出しだけを書く**：`- 人: ①日常点検での見落とし／機械: 軸受の摩耗` の次の行が `- 方法: 点検手順に記載なし／環境: 室温の変動`。
    こうする理由：1行が「どの見出しの、どの値か」を自分の中に持つので、LightRAG が行の途中で切っても意味が通る（レコード自体で意味が通る・パイプ表を使わない、という6章の決まりを変えずに済む）。
    確認画面には組の数を出し、まとめた表をそのまま直せるようにする（行によって空の欄があることも書く）。
  - 確認画面では表の形で修正できる（行の追加・削除）。AIで空欄を探す対象にはしない。承認欄（区分×作成/確認/承認）・なぜなぜ分析・5W2H は明細表にしない（項目として読む）。
- 1つの値の項目の探し方（`excel/extractor.py`）：ラベル候補に一致するセルを上→下・左→右に見て、その右、無ければ下の値を取る。順番は
  「候補どおりのラベル」→「表記だけ違うラベル」（末尾の括弧書き・「内容」「欄」「日時」の違い。例「発生原因（なぜ起きたか）」「応急処置内容」「復旧完了」。
  括弧書きどうしが違う「原因（推定）／原因（確定）」は別項目として扱い、表記違いの照合は塗りつぶし・太字・「ラベル：値」のセルだけ）、
  それぞれで「すぐ隣に値があるラベル」を先に見る（押印欄の縦書き「発信部署」の2行下にある日付を値にしない）。最後に明細表の列見出しにあるラベルを見る。
  - 値にしないもの：行が並ぶ明細表の列見出し（「設備No｜設備名」の下に3行以上）、表の連番の列見出し（No）、区切りの見出し（「▼ 回答欄」）、欄外の様式番号（「様式MT-031 Rev.1」）。
    押印欄（承認｜確認｜作成の下に印と日付の2行）は明細表とみなさず、今までどおり項目として読む。
  - 「対象設備／使用設備／設備」のように設備番号と設備名を1つのセルに書く欄（「CMP-108　STI-CMP 8号機」「ROB-821（ウェーハソーター 1号機）」）は、
    番号と名前に分けて設備番号・設備名の項目に入れる（分けられない値ならその欄は使わない）。
  - 「3時間40分」は項目の単位（分・時間）に換算する。単位が決まっていない項目では分にする。
- **数値項目の単位（2026-09-19 の利用者の判断）**：単位は「帳票の種類の設定」→「書かれた値（「390分」「2.5h」「1,032分」）」の順に決める（`excel/text.numeric_unit`）。既定の単位は辞書に持たない（勝手に決めない）。
  - 両方から決まらない項目のうち、**単位で意味が変わると辞書が知っているもの**（時間・工数・金額・寸法・重量・温度・圧力・流量・電流電圧＝`pattern/dictionary.AMBIGUOUS_UNIT_HINTS`）は要確認にし、「単位が書かれていません（分か時間かで意味が変わります）。値に単位を付けて入力してください（例: 1456分）。これから読み取る帳票のためには「帳票の種類」の画面でこの項目の単位を決めてください（読み取り済みの帳票には反映されません）」と出す（種類の単位を変えても読み取り済みの帳票は読み直さないので、その場で直せる方法を先に書く）。
    件数・回数・人数・枚数・率など、単位が無いのが普通の項目は警告しない（確認画面が警告で埋まらないようにするため。単位の無い数値をすべて警告にはしない）。
  - 種類の単位と書かれた単位が違うとき（「停止時間（分）」の欄に「14.9h」）は、**書かれたとおりの単位で出して**要確認にする（勝手に換算しない。「14.9分」と書くと誤った事実になる）。
  - セルの値が数値だけで、表示形式に単位が書かれている（`#,##0"分"` で画面には「3,095分」）ときは、その単位を値に書かれた単位として扱う（`excel/text.format_unit`）。
  - 「2:45」や時刻・[h]:mm のセルは時:分の時間として項目の単位（分・時間。無ければ分）に換算する。範囲（「10～20分」）で単位が項目の単位と違うときは数値にしない。数値の部分だけを読んだ値・数値として読めない値は、手で入力した値でも要確認のままにする。
  - Excel のエラー値（「#REF!」「#N/A」「#DIV/0!」）は値にせず、空にして要確認にする（一覧表の側と同じ）。
  - 「390分」のように数値の後ろが単位だけのときは、単位として取り込むので「数値の部分だけを読み取りました」の警告は出さない。
  - この判断は `excel/extractor.number_unit` 1か所で行い、読み取り（`_apply_number_unit`）と [AIで空欄を探す]（`services/ai_assist.fill_missing`）の両方から呼ぶ（経路で単位の付き方が変わらないようにするため）。
- **年の無い日付（2026-09-19 の利用者の判断）**：「2/12 3時17分」のような年の無い日付は**年を補わない**。値は書かれたままの文字列で残し、要確認にして「年が書かれていません。元のファイルを確かめて、2026-02-12 のように年から書いてください」と出す（`excel/text.to_date`）。
- **日付のあとの時刻・続く文字**：「2024年7月29日 13:41」「R5.11.16 12:11」「2023-07-10T23:08」「2023年7月10日 23時08分」の時刻は残す（年月日・和暦・曜日「(月)」のあとでも）。日付と時刻のあとに別の文字が続く（「2023/7/10～7/12」）ときは日付だけを読み、「日付のあとに「～7/12」が続いています」と要確認にする（範囲の終わりを黙って落とさない）。和暦は令和（R）と平成（H）を読む。
  同じ帳票の別の項目やファイル名から年を推すと、外れたときに Markdown に誤った日付を書くことになるため。「24/8/25」のような2桁の年は年の欄が空とは限らないので、今までどおり「日付として解釈できません」。

### 6.2 一覧表（記録ファイル）
ファイル単位の既定：**発生年月ごと**（`{prefix}_{YYYY-MM}.md`）。1ファイル300件を超えたら `_part2`…（キー順で分割。分割位置は件数で決まる）。設定で「設備×月」も選択可（`{prefix}_{設備番号}_{YYYY-MM}.md`）。日付が空の行は `{prefix}_日付なし.md`。
```
# トラブル対応一覧 2026年8月の記録（1/2）

- データ種別: トラブル対応一覧（1行＝1件）の記録
- 対象期間: 2026-08-01〜2026-08-31
- このファイルの記録: 300件（2026年8月の全 342件のうち 1〜300件目）

## 【TR-2026-00831】CMP研磨装置1号機（CMP-101）スラリー流量低下アラームで研磨停止｜2026-08-03
- 管理No: TR-2026-00831
- 発生日時: 2026-08-03 14:20（2026年8月）
- 設備: CMP研磨装置1号機（CMP-101）
- ライン: L1
- 故障区分: ユーティリティ
- 現象: スラリー流量低下アラーム（ALM-2031）で研磨停止
- 原因: スラリー供給ラインのフィルター目詰まり
- 処置:
  フィルター交換
  流量再校正
- 停止時間: 95分
- 対応の要点（AI抽出）:
  原因: エンコーダケーブルのコネクタ緩み（確定）
  恒久処置: コネクタの増し締め、エンコーダケーブル交換
  使用部品: エンコーダケーブル RB-ENC-05M ×1
- 対応の時系列:
  1. 2026-08-03 14:20［連絡・初動］田中｜CMP研磨装置1号機: ライン停止の連絡あり。…
  2. 2026-08-03 14:40［暫定処置］田中（推定）｜CMP研磨装置1号機: 原点復帰→再起動で復旧。経過観察。
- 出典: T1_トラブル対応一覧_2023-2026.xlsx（管理No TR-2026-00831）

## 【TR-2026-00877】...
```
- 見出しは `title_columns`（既定: record_no, entity_label(entity), symptom 先頭40字）＋日付。
- 「対応の時系列」はルール出力（AI の種別があれば［］に付与）。「対応の要点」は AI 照合に通った項目のみ。
- 時系列の本文から、**同じレコードの他の列（原因・処置内容・使用部品など text/string 列）と同じ文**（「。」「、」改行で区切り、番号・空白を除いて完全一致、6文字以上）を省く。全部が重複なら「（処置内容と同じ）」に縮める。記入が1件だけのログも同じ。比べる前に行頭の番号（「1.」「(1)」「①」など）と「・」を外す。見出しごとの形（「対応内容（見出しごと）」として出すもの）は時系列ではないので省かない。取り込み設定の `markdown.dedupe_timeline`（既定 True。画面は「対応の時系列から、他の列と同じ内容の文を省く」）で切り替えられる。
- 1レコードの推定トークンの上限は、LightRAG ヒントの `chunk_ts` − 100 ＝ **1,400**（`tables/markdown.RECORD_TOKEN_BUDGET`）。超える場合は「対応の要点」も含めた合計で判定し、時系列を設定の件数（既定20件）→収まる件数まで切って「（以降N件は管理用の正規化CSVに収録）」と書く。
- 設備名の列がない code 型の entity 列では、「ETC-302(OXIDEエッチャ 2号機)」「CVD-203 W-CVD 3号機」を設備番号と名前に分けてから、集計・ファイル分け・表示に使う（番号は英字と数字を含むものだけ。名前の側も番号だけなら分けない）。
- 見出しが「コード・区分・フラグ」を含み、値が英数字1〜4文字・種類20以下で、辞書にも取り込み設定にも無い列は、列の候補で既定 `md=omit` にする（`- 状態コード: 9` のような意味のない行を出さない。画面で戻せる）。

### 6.3 一覧表（集計・説明）
- `{prefix}_00_データセット説明.md`：取り込み範囲、件数、ファイル構成、列の意味（description）、出していない列の名前、数値の注意（集計ファイルの単位でのみ確定）、答えられる/答えられない質問の例。
- `{prefix}_集計_月次_{YYYY-MM}.md`：件数、measure 合計、上位5 entity（measure 合計・件数・主な category）、category 内訳。
- `{prefix}_集計_設備別_{entity}_{FY}年度.md`：件数・measure 合計/平均、月別（0件の月も明記）、category 内訳。entity の記録がある年度のみ。
- クロス集計の設定は記録ファイルを作らず、設備別年度集計（月別値の列挙）のみ。
- 冒頭に「集計対象: 期間（取り込み範囲）」を必ず書く。数値はすべてコード計算（AI 不使用）。取り込み範囲・集計対象の期間は、最初と最後の月については記録の実際の最初・最後の日付で書く（「2026-08-03〜2026-08-20」。月の初日・末日に広げない）。記録ファイルの「対象期間」は月のまとまりを表すので月全体のまま。
- 0件の月は、記録のある月どうしの間が12か月以内のときだけ埋める（`tables/summaries.MAX_EMPTY_GAP_MONTHS`）。それより離れている（「2052年」のような入力ミス）ときは間を埋めず、確認画面で、記録の年の中央値から5年を超えて離れた日付を行番号つきで知らせる（警告 `date_outlier`）。

### 6.4 LightRAG 案内ページ（/settings/lightrag）の内容
1. サーバーのバージョン確認（`GET /health`）。1.5.x は doc_id＝ファイル名の MD5、同名は 409 → 変更ファイルは**削除してから**再投入。1.4.x は doc_id＝内容の MD5。
2. `.env` 推奨：`SUMMARY_LANGUAGE=Japanese`、`ENTITY_TYPE_PROMPT_FILE`（1.5.x。`ENTITY_TYPES` は起動失敗）＝アプリからダウンロードできる設備保全向け YAML、埋め込みモデルは最初に固定、数万件なら PostgreSQL 等。
3. 分割：サーバー設定ごとの可否を表で示す。`LIGHTRAG_PARSER` 未設定（1,200トークン固定窓）では記録ファイルの 2〜6 割が途中で切れ文字化けも出るので**使えない**、`env.example` のまま（native-P 2,000）は可、サーバー全体の `legacy-R(chunk_ts=800)` は帳票が割れるので不可。一覧表の記録ファイルは**ヒントを付ける**（推奨）。ヒントの有無で doc_id は変わらないので入れ直す前に削除する。帳票にはヒントを付けない（短い帳票は1断片。明細表を読む帳票は 1,200トークンを超えて分かれることがあるが、長文項目の見出しに識別子が入る）。
4. 件数・ランキング・推移は集計ファイルで答える。記録ファイルだけでは数えられない。
5. 更新手順：このアプリは投入済みファイルとの差分を管理しない（8.1）。取り込みごとに、その取り込みの全ファイルを zip で渡す。完了画面で zip をダウンロード→同じファイル名のものが LightRAG に入っていれば LightRAG 側で先に削除（1.5.x は同名で 409）→idle を待つ→zip の `RAG投入用/` を投入（`管理用_RAGには入れない/` は入れない）。帳票は確定し直したものだけを「削除してから入れ直す」で更新する。
6. 保存先フォルダを `INPUT_DIR` にする使い方（2.8）：LightRAG で同じ名前の文書を先に削除（このときは「アップロードされたファイルも削除」をオン＝`delete_file=true`。オフだと `__parsed__` に古いファイルが残り、アプリの「止める」で保存が止まる）→idle を待つ→アプリで［保存先フォルダに保存］→LightRAG で［スキャン］（`POST /documents/scan`）。保存の**あとで**削除するときは逆にオフにする（オンだと保存したファイルも消える）。スキャンは直下のファイルだけを読み、`_管理用_RAGには入れない/` は読まない。取り込み終えたファイルは `__parsed__/` に移る。

## 7. テスト方針
- 既存テストは新ルートに合わせて書き換える（test_extraction は維持）。
- WPごとの単体テスト（純粋関数中心）。samples/ の実ファイルを使うテストは `@pytest.mark.samples`（samples が無ければ skip）。
- ゴールデン：同じ入力→同じ md バイト列（ハッシュ比較）。
- AI：`tests/fake_servers.py` を拡張し、リクエスト内容で応答を変える（壊れたJSON、原文にない型番、429、遅延）。
- 画面：Flask test client で主要フロー（帳票: upload→type→read→review→confirm→done→download、一覧表: CSV upload→source→layout→columns→preview→confirm→done→zip）。
- ブラウザ確認（統合時）：ホーム（作業中／ダウンロード待ち）、両フロー（帳票は1件・まとめ取り込みの両方）、設定、エラー画面（403/404/500）。
- 帳票サンプルでの精度測定：`python -m scripts.samples.evaluate_forms`（見本ファイル数は `--samples`、見本の選び方をずらすのは `--offset`）。
  見本を変えても結果が同じ傾向か（特定の見本への当て込みでないか）を `--samples 2` や `--offset 3` で確かめる。

## 8. 実装しないもの（第2段階以降）
大きいExcel用の1パス読み込み、数式XMLの解析（小計はキーワード＋太字で判定）、同じ構造の複数シート一括、複数表の自動検出（範囲の手動指定で対応）、帳票内の明細表の高度な形（2段の列見出し、承認欄のような行見出し付きの格子、日付が変わった行だけ書く時系列の日付の補完）、full プロファイル、巨大セルの分割AI処理、抜き取り確認の統計、修正内容からのルール提案、設備台帳、帳票と一覧表の紐付け、Batch API。

### 8.0 残っている不一致・開いている点
測定値（`python -m scripts.samples.evaluate_forms`、見本各3件・対象各30件、2026-09-19）：
F1 98.5% ／ F2 97.7% ／ F3 96.4% ／ F4 100.0% ／ F5 93.9%、
全体で 1つの値の項目 2,941/3,036 = 96.9%（正解が空なのに読んだ 8）、明細表の行 1,970/1,995 = 98.7%。
明細表の行は、評価スクリプトが正解の `*_norm`（ISO の日時）も照合に使い、真偽値のチェック欄（`checked: true/false`）を
☑ / □ とみなすようにしたあとの値（以前の数え方では 88.8%。F2 の時系列 311/314・横展開先 61/61 が照合できていなかった）。
残っているのは次のもの。

**読み取りの不一致**
- 発行側と回答側で同じ意味の欄が並ぶ帳票は、項目の「探す区画」（▼ 回答欄 など区切りの見出しの名前）で読む側を決める。
  見本の中で見出しが2つ以上の区画にあり、どの見本にも共通する名前付きの区画が1つだけの項目は、帳票の種類を作るときに
  自動で決まる（F5 の処置・原因＝回答欄。未回答なら空。`action` 24/24）。区画は `pattern_fields.extraction_rule` の
  `section` に保存する。押印欄の「確認」のように区画の外にも共通してある項目は決めない（F5 の `checker` 16/25。
  見本からはどちら側か決められないので、必要なら画面の「探す区画」で決める）。
- 見本に無い版だけにあるラベル（辞書に無いもの）は読めない（確認画面でラベルを足す運用）。
- 見出しと列見出しの間に「展開区分｜☑同型機」の1行がある明細表は、上の「■」「1.」の見出しをアンカーにして読む
  （F2 横展開先 61/61 行）。
- F5 の明細表の「処置」列（0/6）と数量列（投入数・不良数・保留数 11/16・12/14）は、評価スクリプトが正解の合計値・
  チェック値を明細表の列をつないだ値と比べているための差で、読み取りの誤りではない。

**明細表の形**
- 同じ形の列見出しが縦に積み重なる明細表（F3 の特性要因図＝人｜機械・材料｜方法・測定｜環境）は、2026-09-19 の判断で**全部の組を1つの表にまとめて読む**ようにした（6.1「積み重なった列見出し」。F3 の6カテゴリとも 18/30 ファイルで一致。残り12件は見出し（アンカー）自体が見つからず空欄のまま）。
- 時系列の明細表で日付列と時刻列が分かれている様式は、行の日付を補完・結合しない（8. の「日付が変わった行だけ書く時系列の日付の補完」）。

**値の書き方（2026-09-19 に利用者が決定。実装済み。詳しくは 6.1）**
- 年の無い日付（F1 の「2/12 3時17分」）は年を補わず、書かれたままの文字列を残して要確認にする（見本で 9 件）。
- 数値項目の既定単位は辞書に持たない。単位は「種類の設定 → 書かれた値」の順に決め、どちらからも決まらない項目のうち単位で意味が変わるもの（時間・金額など）だけを要確認にする（見本で 43 件。件数・回数などは警告しない）。種類の単位と書かれた単位が違うときは書かれたとおりに出して要確認にする（F1 の「14.9h」4 件）。
- 丸数字（①②③）は囲みを外さずそのまま出す（帳票の出力経路 `export/formats.py` は `core/mdtext.nfkc_value`＝`nfkc_keep_enclosed` を使う）。一覧表（`tables/normalize.nfkc_text`）も同じ扱い。

**見本の選び方で決まるもの（読み取りの仕組みの限界）**
- 見本に無い書き方の見出しは読めない。F3 の2シート版（「要因分析（6M）／Man（人）／Machine（設備）…」）は、
  「特性要因図／人／機械…」版だけを見本にした種類では「根本原因の追究」が空になる（画面には「ラベルが見つかりません」と出る）。
  別名辞書（8.1）は作らない前提なので、運用（両方の版を見本に入れる）で対応する。未一致の見出し候補を確認画面で知らせる案は未実装。
- 見出し語が明細表の列見出しと同じとき（F3 の D6「実施内容」）に、表を1行下から読んで先頭の明細と一部の列を落とす問題は
  2026-09-19 に直した（`excel/extractor.locate_table` が、見本で見た列見出しと重ならない表を読んだときは `find_table_by_columns` を優先する）。

**データを残さないことの残り（3.3）**
- ダウンロードが途中で切れても、本文の最後の約48KB（waitress の溜め 16KB＋OS の溜め）がまだ届いていないだけなら
  サーバー側は「送り終えた」と判断して消す。帳票1件の md（数KB〜十数KB）はほぼ全部がこれに当たり、途中で切れても消える
  （3.3「欠けないダウンロード」。手元の元の Excel から取り込み直す）。ブラウザが保存し終えたことを画面から知らせて
  から消す形（受け取り確認）にすれば防げるが、ダウンロードを2段階にする変更になるため未実施。
  応答を作れなかったとき・206/304・1バイトも送れなかったときは消さない。
- 帳票の種類の見本ファイルはこのPCに残り続ける（3.3「残ると分かっていて残すもの」）。本物の報告書を見本にする運用では、
  使い終わったら画面から削除する必要がある。自動で消すには「読み取りテストのたびに見本を選び直す」形に変える必要があり、未実施。
- `instance/app.db` のファイル自体は残る（中身は `secure_delete` + `VACUUM` で消える）。`.flask_secret`・`data/model_settings.yaml`
  （APIキーを平文で持つ）・`env` も残る。
- AI の応答のキャッシュ（`llm_calls`）は、払った取り込み（`import_id`）を持つ。どの行の結果にも結び付いていないもの
  （作り直しで置き換わった最初の応答、一時停止の直前に保存された応答）も、その取り込みを消すときに一緒に消える。
  持ち主の分からない行（`import_id` が NULL。この列を足す前の古いDB）だけは、AI整形の作業（`ai_items` の行か
  `ai_format` のジョブ）が残っている取り込みが1つでもある間は消さない（別の取り込みの再開で使うかもしれないため。
  `core/purge._ai_work_alive`）。同じ文面の行で別の取り込みが同じ応答を使っているときは、その応答は使っている側が
  消えるまで残る（消すと巻き添えの再課金になるため。持ち主は「持ち主なし」に戻す）。
- ジョブが待機中・実行中・一時停止中の間は `VACUUM` を省く（`secure_delete` で空いたページは上書き済み。ファイルの縮小は
  次にジョブが無いときの削除で行う）。

**承知のうえで残している小さな点（2026-09-19 の確認）**
- 一覧表の読み込み・md 作成のジョブを始めるとき、状態を先に書いてからジョブを登録するので、その2つの書き込みの間
  （マイクロ秒）に開いた画面は、新しい状態と前回の終わったジョブを見ることがある。画面は POST → リダイレクト → GET の順なので実害は無い。
- CSV の列数の上限（2,000列）は、自動判定（1件目が判定用の読み込み量より長いときは1行目の区切り文字の数から見積もる）と、
  読み進めるとき（`CsvSource.rows` が上限を超える行で「N行目の列数が上限…」を返す）の両方で確かめる。
- AI接続のタイムアウト（接続時を含む）は「混んでいるだけ」として再試行する。止めるのは接続拒否・名前解決の失敗だけ。
- 帳票の1つの値の項目で、日付のすぐ右のセルに時刻だけが書かれていれば日付につなぐ（「2023-05-23」｜「12:07」、
  時刻のセル、「9時5分」。日付だけ読めて警告なし、すき間なし・同じ高さ、その右が空か見出し欄のときだけ。範囲・「頃」付き・
  すでに時刻のある日付はつながない。値のセルは「F9:R9」のように2つのセルの範囲になる）。F2 旧様式の発生・完了日時 60 件が
  これで時刻まで一致する。
- 一覧表で、空行のすぐ後にあるキー列が空の文字だけの行は、前の記録の続きにせず1件の記録として読む（「日付なし」「必須列が空」の
  警告で見える）。空行をはさまない続きの行は従来どおり前の記録につなぐ。
- 確認画面の途中保存は画面ごとの目印（`page_token`）を送り、今の版がその画面自身の保存でできたものなら、古い版を送ってきた
  画面を離れるときの保存（beacon）も受け付ける。目印はサーバーのメモリにだけ持つ（アプリの再起動をまたいだ送信は、従来どおり版だけで判断する）。
  帳票を消したとき（ダウンロード・削除）はその帳票の目印も捨てる（`core/purge.on_documents_purged`）。
- AI整形の1行は、再試行と待ち時間を含めて「タイムアウト×2」（ローカル 600秒・クラウド 240秒、`aiproc/runner.ROW_DEADLINE_FACTOR`）
  で打ち切り、その行をエラーにして次へ進む（「エラーだけ再実行」でやり直せる）。一時停止・中止は、応答待ちの呼び出しを
  2秒（`runner.STOP_GRACE`）待って見捨てる。その行は未処理に戻し、あとから届いた応答は捨てる（キャッシュにも入れないので、
  消したあとに書き込むことはない）。クラウドでは一時停止のたびに実行中の呼び出し1件分（同時実行数ぶん）のトークンが無駄になりうる。
- アップロードの事前確認（`core/files.precheck_excel`）は、DTD の実体宣言（`<!ENTITY`）を含むファイルを受け付けない（Excel は書かない）。

**4巡目の不具合探しで変えた動き（2026-09-19）**
- 「上の行と同じ」の記号（セル全体が 〃・″（NFKC で ′′）・同上・仝、一覧表では 々 も）は、上の行の同じ列の値に置き換える。
  帳票の明細表は `excel/tables._resolve_ditto`（先頭行や上が空のときは書かれたまま）。一覧表は記録番号・日付・設備・設備名・分類・属性・担当者の役割の列だけで（`tables/normalize.DITTO_ROLES`。6巡目に記録番号・担当者を追加）、
  「fill_down_blank」の設定に関係なく前のデータ行の値で埋め、まとめて1件の警告（`ditto_filled`）にする。上に値が無ければ書かれたまま残して
  行ごとの問題（`ditto_unfilled`）にし、設備としては扱わない。
- 合計行とみなす見出しは「合計・小計・総計・総合計・合計数・計、〜合計・〜小計、〜費計・工数計・個数計・件数計・台数計・本数計・枚数計・金額計・額計」
  だけにする（`excel.tables.TOTAL_LABEL_RE`。読み取りと md の両方で使う）。温度計・圧力計・設計などの行で明細表の読み取りを止めず、md からも消さない。
  一覧表の1列目の「pH計」「電力量累計」のような行も、その行に他の値（日付など）があればデータとして読む。
- 見えない文字（ゼロ幅スペース・BOM・向き指定・ソフトハイフン）はセルの文字列・ラベル・シート名から消す（帳票 `excel/text.strip_invisible`、
  一覧表 `tables/source.clean_text`）。「EQ-01」と「EQ-01＋ゼロ幅文字」は同じ設備になる。
- 帳票でラベルは見つかったのに値が空で、すぐ右か下のセルが計算結果の保存されていない数式なら、「数式の計算結果が保存されていません…」の警告にする。
- アップロードの事前確認は、図形・画像のアンカー数（参照するシートの数を掛けたもの）が 10,000、drawing の大きさ（同）が 20MB を超えるブックを受け付けない。
- 一覧表の取り込み設定: 確定前の取り込みが使っている版は上書きせず新しい版を作る（画面の「次の取り込みから使われます」のとおり）。今回のファイルに無い列は
  設定から消さずに残す（列の対応づけ画面に「今回のファイルにない列」として出る）。「ファイル名の先頭」が空か設定名と同じなら '' で保存し、設定名に従う。
  JSON の読み込み・保存では、伏せ字の規則・人名・分割の正規表現（20個まで・各200文字まで・入れ子の繰り返しは不可）などの形を確かめて日本語で断る。
  深く入れ子の JSON も「取り込み設定のJSONではありません」で断る。検証より前に保存された設定の書けない正規表現は、分割のときに使わずに進む。
- 一覧表で「担当者名を出さない」をオンにしたときの時系列の記入者は、ログに書かれたとおりの表記で出す（前の段落から引き継いだものは「（推定）」付き、
  書かれていなければ出さない）。担当者列の氏名に置き換えない。
- 見出し行がデータのように見える（見出しの半分以上が数値・日付、または40文字を超えるセルがある）表は、範囲の画面で警告する。見出し行の無い表には対応しない。
- AI整形の照合: 原因の「確定」は原文の語句（q）が必須で、根拠に「不明・調査中」などがあればエラー。原因・処置・部品名の語が根拠に1つも無ければ警告。
  「再発はなし」「完了予定」「まだ解決していない」を「再発あり」「完了」の根拠にしない。「1/2に調整」「1日1回」を日付とみなさない。
  修正のための2回目の呼び出しが失敗したときは、1回目の結果を残す。試し実行の途中で取り込みが消されたら結果を書かずに止める。
- 開いている点: プロンプト（`aiproc/prompts.py`）には「確定のときは q を必ず入れる」をまだ書いていない（書くとキャッシュのキーが全部変わり、
  処理済みの行にもう一度課金されるため）。今は修正の呼び出しで伝えている。

**5巡目の不具合探しで変えた動き（2026-09-19）**
- 帳票: 区画の名前は「欄」「内容」などの語尾を外し切るまで繰り返して作る（「■ 処置内容欄」「処置内容」「処置」はどれも「処置」。
  学習・保存を通しても区画が変わらない）。確認画面で入力欄に Enter を押しても、［AIで空欄を探す］や［確定］は押されない。
  種類の項目が設備だけで識別子にならないときの「- 出典:」の行は、H1 と同じ項目（番号の項目、無ければ日付）を添える。
- 帳票・帳票の見本のアップロードは、結合セルの面積の合計が 20万セル（`core/files.FORM_MAX_MERGED_CELLS`）を超えるブックを断る
  （行全体の結合 13個以上・列全体の結合など。一覧表は読み取り専用で開き結合を展開しないので、従来の 200万セルのまま）。
- 一覧表: 大文字・小文字だけ違う設備のファイル名には `_2` を付ける（Windows のフォルダで上書きし合わない）。256列を超えるシートと、
  「行×列」がセル数の上限（`tables/excel_source.DEFAULT_MAX_CELLS` 50万。中身のほとんどが空）を超えるシートは、行を最終列まで埋めない
  （CSV と同じく行ごとに長さが違う。読む側は範囲外を空として扱う）。遠くの行のセル1つ（`IV1048576` のメモなど）で
  全行に何十万個のセルを作って読み込みが何十秒も止まるのを防ぐため。
  表から20列以上離れた、下にデータの無い見出しセル1つ（XFD1 のメモなど）は表に入れず警告にする。
  時間の単位の数値列では「1:30」「1時間30分」の文字も分として読む。確定の処理の途中で消えた取り込みのフォルダを作り直さない。
  設定名を変えたとき、手で入れた「ファイル名の先頭」が前の名前と同じでも消さない。
- 一覧表の md: 時系列の重複の削除で「(1)」「①」「⑴」「・」の行頭も外して比べ、1件だけのログも対象にする。見出しの1行目がタグと日付だけ
  なら次の行を使う。トークン数の見積もり（`core/mdtext.estimate_tokens`）は記号と数字を多めに数える（実測より少なく見積もらない）。
- 利用者が設定で書いた正規表現（分割の目印・日付ではない書き方）は、選択肢（`|`）を含む繰り返しのグループも断り、照合には行の先頭
  200文字（`logproc/dates.USER_PATTERN_WINDOW`）だけを渡す。取り込み設定の JSON の確かめで思わぬ例外が出たときも「読み込めません」で断る。
- 保存先フォルダ: 管理用ファイルを保存しない設定（既定）では、一覧表の保存の確認文（完了画面・ホーム）と結果の画面でそう知らせる。
  `__parsed__`・`_管理用_RAGには入れない` の中や、アプリのデータのフォルダの別の書き方（`\\?\`・`\\localhost\C$`）は断る。
  確定済みの分だけ保存したときは残った未確定の件数と［次の帳票へ］を出す。消し損ねたときは「消しました」と言わない。
- 設定ファイル（`services/settings_store`）は UTF-16・CP932 でも読み、読めなければ空とみなす。読み書きは同じロックの中で行い、
  書けないときは画面に日本語で知らせる（ヘッダーのモデル選択も）。すべての応答に `X-Frame-Options: DENY` と
  `frame-ancestors 'none'` を付ける。
- AI: 帳票の AI 呼び出しとモデル一覧は明示のタイムアウト（ローカル300秒・クラウド120秒・一覧30秒）で、SDK の自動再試行はしない。
  構造化出力の確かめは推論モデルでも切れない長さで行う。429 で失敗した行が3行続いたら、エラーにせず未処理に戻して一時停止する
  （「混み合っています…再開を押してください」）。一時停止を押してから止まるまでの間も［再開］を出す。
  2つ目の起動が、1つ目のアプリの待機中のジョブを「中断」にしない（待機中のジョブの生存時刻も更新する）。

**6巡目の不具合探しで変えた動き（2026-09-19）**
- 帳票: 見出しの区画は、右の上の行から始まる見出し（右上の【回答欄】など）の列で右への広がりを止め、その見出しは下の行で別の見出しが
  その列に来るまで自分の列を持つ（右上の押印欄の見出しも同じく下の列を持つ。F1〜F5 の見本には無い形）。入れ子の区画（「暫定対策」の中の
  「回答」など）は `excel/tables.sections_of` で内側から外側まで返し、「探す区画」がどれかに当たればその区画の中とみなす。
  明細表の項目も「探す区画」の中を先に探す（無ければシート全体）。確認画面の「変更した項目」は単位だけの変更も出す。
  まとまりの1件を保存先フォルダに保存したあとは、残りの未確定の件数と［次の帳票へ］を出す。
- 一覧表: 手で決めた見出し行が先頭80行より下でも読む。行番号の入力（「1-3, 5」）は全角も含めてサーバーと画面で同じ規則で読む
  （範囲は10行まで）。役割「追記ログ」は1列だけに付ける。日付の役割が変わらなければ期間の日付列を保つ。キーの重複を断る。
  確定済みの取り込みでは試し実行をしない。読み込み直しが始まって渡せなかったときは 404 ではなく画面に戻して知らせる。
  Excel は 16,384 列・1,048,576 行を超えるシートを断り、値のある列が 2,000 列を超えるシートも CSV と同じ文で断る。
  md のフォルダのパスが長すぎるときは「設定名」「ファイル名の先頭」を短くするよう知らせる。
- 一覧表の md: 月別・設備別の要約の設備名は、取り込み全体でいちばん多い名前にする（グループの先頭行ではなく）。平均は小数1桁に丸める。
  「10:00以降」「4/3から」のように範囲の語が続く日時は見出しから外さない。CSV の数式よけは全角の ＝＋－＠ で始まる値にも付ける。
- 分割（`logproc`）: 「1.5mm」「2.5A」のような寸法・電気量を日付にしない規則（`DEFAULT_NOT_DATE_PATTERNS`）は、単位の後ろに英数字・「-」・
  カタカナが続く「4.3 AGV」「4.3 ALM-2031」「4.3 Aライン」「4.3 Vベルト」には当てず、日付として読む（保存済みの取り込み設定の規則はそのまま）。
- AI: 429 のあとで行の時間切れになったときも 429 として数え、続けば一時停止にする。応答のキャッシュの読み書きがロック中なら少し待って
  再試行し、だめなら使わずに（書かずに）進む。ジョブの進捗・メッセージ・生存時刻の書き込みもロック中なら少し待ち、だめならその1回を見送る
  （`core/jobs.JobContext._soft_update`。状態の書き込みは見送らない）。
- アップロード: ハイパーリンク・コメントの範囲の合計が5万セルを超えるブック、帳票では 20万行を超えるブックを開く前に断る（5.1）。
  JSON の 404 にも「保存先フォルダに保存済み」を出す。帳票の取り込みの画面と一覧表の確認の画面にも、保存先フォルダに保存したときも消えることを出す。

**6巡目の見直しで変えた動き（2026-09-20）**
- 帳票: 値が「2024年7月28日 22:46」のように日付だけで書かれていれば、ラベルに日付の語が無くても（「発生」など）日付の項目にする
  （`pattern/builder._date_value`。値の全体が元号・年月日（＋曜日・時刻）で、`to_date` が警告なしで読めるときだけ。
  「2/12」「12:30」「2026-09-14 に復旧」は文字列のまま）。まとめ取り込みが1件だけになったら、まとまりとして扱わない
  （確認文から「残りの帳票はまとまりに残ります」が消え、ホームにも1件の帳票として出る）。
- 一覧表: 記録番号・担当者の列の「〃」「同上」も直前のデータ行の値で補う（`ditto_filled` の警告にまとめる。
  重複キーの警告が「〃」ではなく本当の伝票番号を指すようになる）。取り込み設定の記録キーは「列:文字数」の書き方も受け取り、
  手で書いた `record.fallback_key` の列も確かめる。右側の別表・遠くの見出しの警告は、無い「列の範囲指定」ではなく
  「Excel で分ける／右隣に移す」を案内する。データの行が0行のときは、見出し行ではなく指定した「終わりの行」を理由として知らせる。
  取り込み設定の保存が UNIQUE で断られたとき、名前の重複でなければ「別の名前にしてください」と言わない（版番号は INSERT の中で数える）。
  中止は、押す直前に読み込みが終わっていれば状態を巻き戻さない。取り込みの［削除］の確認文と、処理中は出さない規則は
  `templates/components/_ui.html` の `table_import_delete_form` に1か所化した。
- AI: 対応していない引数の学習は「値を固定する」ではなく「名前を置き換える」（`max_tokens` → `max_completion_tokens`）。
  接続テストの 20 トークンがそのあとの全行に残らない。学習は（接続先URL, モデル）ごとで、AI接続を保存すると忘れる。
  見積もりは DB 接続を1本で通す（8,000行で約24秒短縮）。基準日の決め方は `tables/spec.base_date_from` の1か所にまとめ、
  分割プレビュー・試し実行・送る文面・できる md で同じ日付になる。AI接続が外れていても、動いている AI整形は［中止］できる。
- AI整形の試し実行の片付けは、まだある別の取り込みが払った応答を消さない（`core/purge.NO_LIVE_OWNER` を `aiproc/runner._ensure_trial_import`
  でも使う。消すと払った取り込みの再開・再実行で再課金になる）。
- 画面: 列の対応づけの保存が複数の理由で断られたら全部並べる（`static/app.js` の `postJson` が `errors` を例外に載せる）。
  AI整形の［見積もる］は答えが返るまで押せない。待ち画面のポーリングは、ジョブが消えた（404）ら止まって画面を読み直す。
  ホームの「作業中の一覧表」の副題は、そこに出る状態（`views.TABLE_IMPORT_ACTIVE`）をすべて言う。

**Markdown の断片**
- 明細表を読むようになったため、帳票の md が 1,200トークン（LightRAG の既定の固定窓）を超えて2つ以上の断片に分かれる様式がある（サンプルでは F2・F3・F4）。長文項目の見出しには識別子が入る。明細表の行だけで埋まった断片には識別番号も設備名も出なかったので、識別子を入れる帳票では明細表1節が 800トークンを超えたところで `## {見出し}（続き）（{識別子}）` に分けるようにした（6.1）。明細表の各行への設備番号の付与・帳票へのヒント付与は、出力仕様の変更になるため入れていない（`docs/research/LightRAGオフライン評価.md` 8.5/8.7）。

### 8.1 後回し（2026-09-16 の範囲の見直しで外したもの）
統合時に、次の機能のコード・ルート・画面・テストを削除した。必要になったら 2.4・2.6・3.2・5.3・6.2〜6.4 の記述をもとに作り直す。
- **一覧表の更新管理**: 期間の置き換え（`tables/state.py` の current/previous、`replace_period`）、投入済みとの差分（`table_outputs.delivered_hash`、差分zip、「削除すべき旧ファイル」）、[投入済みにする]（`table_downloads`）、`/tables/templates/<tid>/outputs` 画面、直前の確定の取り消し。現状は「取り込みごとに全 md を作り、全ファイルを zip で渡す」（`tables/pipeline.py` の `run_read` / `run_render` / `build_download`）。md は `TABLES_DIR/imports/<id>/md/` に保存。
- **クロス集計**（縦持ち変換、`TableSpec.kind="crosstab"`、年月見出しの解釈、設備×月の値の集計ファイル）。表の形の判定（`tables/detect.py`）はクロス集計を見分けるが、範囲確認の画面で「対応していません」と止める。
- **名寄せ辞書**（`/settings/aliases`、`alias_entries` の参照、`ColumnSpec.alias_dictionary`）。値は NFKC と空白の畳み込みだけで揃える。
- **取り込み設定の版の履歴画面**（版は内部で保持: 確定に使った版は上書きせず次の版を作る）。
- **構築後のレビュー・評価フェーズ、LightRAG へのオフライン評価**（利用者が後で行う）。
- **表の読み取り方を取り込み設定で変える項目**（`header.search_rows`、`data_end`（`blank_rows` / `stop_first_col` / `stop_prefix`）、
  `exclude.aggregate_keywords`）。判定は `tables/detect.py` が持つ（合計・小計の語は「設計」「稼働時間累計」などと混ざらないよう
  細かく調整してあり、設定で差し替えると誤判定に戻る）。古い JSON にこれらの項目があっても読み捨てる（`tables/spec.py` の `RETIRED_KEYS`）。
  見出しの行・データの終わりの行は、取り込みごとに［範囲の確認］の画面で指定する。

DB のスキーマ（3.2）にあった `table_outputs`・`table_downloads`・`alias_entries`・`table_template_samples` は、取り込みを指す列
（`document_id` / `import_id`）を持たず purge の探索から漏れるため、マイグレーション `_m5` で削除した（どこからも書いていなかった）。
これらの機能を作り直すときは、取り込み単位の行を持つ表に `import_id` 列を付けてから作る。`table_imports.period_json` は書かない（常に全期間）。
