# 統合設計書：RAG用Markdown作成アプリ（帳票・一覧表）

版: 2026-09-15（作り直し v2）。この文書が実装の正本。根拠となった調査・設計は `docs/research/` に置く。

## 0. 前提と範囲

- 利用者は1人。**ログインなし**。起動は `127.0.0.1` のみ（`app.py` が他のアドレスを拒否）。
- 目的：Excel/CSV を読み取り、**LightRAG に手作業で投入しやすい Markdown(.md) を作ってダウンロードする**まで。LightRAG への送信・アプリ内のSQL照会は作らない。
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
| 取り込み一覧 | 取り込み履歴 | 登録文書、文書一覧 |
| RAG への投入済み印 | 投入済みにする | 送信 |

ボタンは動詞。1画面に主ボタンは1つ。ステップ名＝画面見出し＝次へ進むボタンの動詞。

## 2. 画面構成とルート

### 2.1 ナビゲーション（base.html）

`ホーム` / `帳票を取り込む` / `一覧表を取り込む` / `取り込み履歴` / `設定 ▼`（帳票の種類・一覧表の取り込み設定・名寄せ辞書・AI接続・LightRAGへの入れ方）。右端にヘッダーのAIモデル選択（既存）。

### 2.2 ホーム `GET /`
- 大きな入口カード2枚（図＋1文）：「帳票を取り込む（1ファイル＝1件の報告書・記録票）」「一覧表を取り込む（Excel/CSV、1行＝1件の一覧・台帳・集計表）」。
- 作業中（帳票: 確認中・修正中、一覧表: 処理中・プレビュー）の一覧（続きを開くリンク）。
- 最近の確定（帳票）と最近の一覧表出力。
- 初回（帳票の種類も一覧表の取り込み設定も0件）は「はじめに」3ステップ（見本から種類/設定を作る → 取り込む → Markdownをダウンロード）。

### 2.3 帳票フロー（blueprint `forms`、prefix `/forms`）

ステップ表示：`1 ファイルを選ぶ` → `2 帳票の種類とシートを確認` → `3 読み取り結果の確認・修正` → `4 完了（Markdownをダウンロード）`

| ルート | 内容 |
|---|---|
| `GET /forms/new` | ファイル選択（ドラッグ＆ドロップ可）。.xlsx/.xlsm。帳票の種類が0件なら作成へ案内 |
| `POST /forms/upload` | 保存→ `/forms/<id>/type` |
| `GET /forms/<id>/type` | 候補の種類（「10項目中9項目」表示）、シート選択、[AIで種類を推定]。一覧表らしい（見出し行の下に同形行が10行以上）なら「一覧表の取り込みへ」を提案 |
| `POST /forms/<id>/read` | 読み取り→ `/forms/<id>/review` |
| `GET /forms/<id>/review` | 左：元のシートのHTMLプレビュー（セルグリッド、項目にフォーカスで見出しセル・値セルをハイライト）。右：項目（入力欄、状態タグ、警告）。上部チップ「要確認 N / AIが入力 N / 手で修正 N」。[AIで空欄を探す]。Markdownプレビュー（タブ、入力に連動して更新）。主ボタン [確定してMarkdownを作成]。出口は非破壊の「取り込み履歴に戻る（作業内容は保存されています）」 |
| `POST /forms/<id>/draft` | 入力内容の途中保存（JSON、fetch。PRGなし・204） |
| `POST /forms/<id>/ai-fill` | 空欄をAIで探す→review へ |
| `POST /forms/<id>/confirm` | 必須欠落があれば止める（「空欄のまま確定」チェックで許可）。confirmed に保存→ `/forms/<id>/done` |
| `GET /forms/<id>/done` | 完了。主ボタン [Markdownをダウンロード（.md）]、[次の帳票を取り込む]、[取り込み履歴へ] |
| `GET /forms/<id>` | 詳細（確定内容、Markdown、元ファイル）。[修正する]→review（修正中） |
| `POST /forms/<id>/discard-changes` | 修正中の変更を破棄し確定済みの版に戻す |
| `GET /forms/<id>/download.md` / `.json` / `original` | ダウンロード。md は **確定済みデータから毎回生成** |
| `POST /forms/<id>/delete` | 削除（確認文「元のファイルと読み取り結果を削除します。元に戻せません」） |
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
| `GET /tables/imports/<id>/preview` | 置き換える期間の確認（必須）、概要（読込件数・除外件数と理由・エラー/警告・合計の照合・期間外の行）、**mdファイルの差分**（投入済みと比べた新規/変更/削除、選んだファイルの中身）、問題一覧（CSV）、データ（100行ずつ） |
| `POST /tables/imports/<id>/confirm` | ジョブで確定（状態置換→全md再生成→出力管理更新）→ done |
| `GET /tables/imports/<id>/done` | [差分のみzip（新規N・変更N）] [全ファイルzip] [正規化CSV]、削除すべき旧ファイル名一覧、[投入済みにする] |
| `GET /tables/templates/<tid>/outputs` | その設定の現在の全md（一覧・再ダウンロード・投入済み状態・直前の確定の取り消し） |
| `GET /api/jobs/<job_id>` | ジョブ進捗（JSON） |

### 2.5 取り込み履歴 `GET /history`
- タブ：帳票 / 一覧表。
- 帳票：種類・状態・期間（取り込み日）・キーワード（ファイル名＋タイトル項目の値）で絞り込み。列：No.、タイトル（タイトル項目の値）、種類、状態、取り込み日時、操作（状態に応じて「続きを確認」「Markdownをダウンロード」）。チェックして「選択した○件をダウンロード（zip）」。ページング（50件）。
- 一覧表：取り込み（設定名・ファイル名・件数・状態・日時）。

### 2.6 設定（blueprint `settings`、prefix `/settings`）
| ルート | 内容 |
|---|---|
| `/settings/form-types` | 帳票の種類の一覧・作成（見本ファイル→候補確認→読み取りテスト→**使用開始**）・編集（保存しても使用中にしない。作成中のものは読み取りテスト後に[使用開始]）・停止/再開・削除（影響：確定済み帳票N件は残る） |
| `/settings/table-templates` | 一覧表の取り込み設定の一覧・版・JSON書き出し/読み込み・削除・出力状態へのリンク |
| `/settings/aliases` | 名寄せ辞書（辞書名=設備番号 など、別表記→正式な値・表示名）。一覧表の未登録値から追加も可 |
| `/settings/ai` | AI接続（既存）＋[接続テスト]（models.list と 1回の短い chat） |
| `/settings/lightrag` | LightRAGへの入れ方（サーバー設定の推奨、ファイルの入れ方・入れ替え手順、エンティティ種別YAMLのダウンロード） |

旧ルート（`/documents/*`, `/patterns/*`, `/settings/models`）は削除。`/api/models` は維持。

### 2.7 UI共通部品（templates/components/_ui.html のマクロ）
`steps(items, current)`、`card`, `badge(state)`, `chips`, `empty_state`, `confirm_delete_form`, `progress(job)`, `data_grid(rows)`（セルグリッド）, `file_drop(name, accept)`。CSS はデザイントークン（色・余白・角丸）を `:root` 変数で統一。ライト基調、配色は現行を踏襲しつつ整理。

## 3. データモデル（SQLite `instance/app.db`）

`models/database.py`：接続時に `PRAGMA foreign_keys=ON; journal_mode=WAL; busy_timeout=5000`。`PRAGMA user_version` による連番マイグレーション（`MIGRATIONS: list[Callable[[sqlite3.Connection], None]]`）。バックグラウンドスレッド用に `connect()`（`g` を使わない接続）を公開。

### 3.1 帳票（既存を拡張）
- `patterns`：既存列＋`title_fields TEXT DEFAULT '[]'`（タイトル・ファイル名に使う field_name の配列）、`md_options TEXT DEFAULT '{}'`（`{"domain_context_fields": [...], "omit_person_fields": true}` 等）、`version_no INTEGER DEFAULT 1`（保存ごとに+1）。`status` の値は draft/active/inactive のまま（表示語だけ変更）。
- `pattern_fields`：＋`unit TEXT DEFAULT ''`、`rag_output TEXT DEFAULT 'show'`（show/omit）。
- `documents`：＋`confirmed_json TEXT`、`confirmed_at TEXT`、`title TEXT DEFAULT ''`（検索用。タイトル項目の値を連結）。`data_json` は作業中の値。状態は導出：`data_json IS NULL`→読み取り前、`confirmed_json IS NULL`→確認中、`confirmed_json != data_json`→修正中、それ以外→確定済み。`markdown`・`registered_at`・`status` 列は使わない（互換のため残す。マイグレーションで `registered_at`→`confirmed_at`、`status='registered'` の `data_json`→`confirmed_json` にコピー）。

### 3.2 一覧表
```
table_templates(id PK, name UNIQUE, description, current_version_id, created_at, updated_at)
table_template_versions(id PK, template_id FK, version INTEGER, spec_json TEXT, spec_hash TEXT, used INTEGER DEFAULT 0, note, created_at, UNIQUE(template_id, version))
table_template_samples(id PK, template_id FK, file_name, stored_path, file_hash, created_at)
table_imports(id PK, template_id FK NULL, template_version_id FK NULL, file_name, file_hash, stored_path,
              source_json TEXT,   -- {"kind":"csv|excel","encoding","delimiter","preamble_rows","sheet","header_row","header_rows","data_end_row"}
              period_json TEXT,   -- {"grain":"month|fiscal_year|all","start":"2026-08","end":"2026-08"}
              status TEXT,        -- uploaded / reading / preview / confirming / confirmed / discarded / failed
              stats_json TEXT, issues_path TEXT, rows_path TEXT, job_id INTEGER, created_at, updated_at, confirmed_at)
table_outputs(id PK, template_id FK, file_name, content_hash, delivered_hash, delivered_at, removed INTEGER DEFAULT 0, UNIQUE(template_id, file_name))
table_downloads(id PK, template_id FK, kind TEXT, files_json TEXT, created_at, delivered_marked INTEGER DEFAULT 0)
alias_entries(id PK, dictionary TEXT, alias_norm TEXT, canonical TEXT, display TEXT, created_at, UNIQUE(dictionary, alias_norm))
jobs(id PK, kind TEXT, ref_type TEXT, ref_id INTEGER, status TEXT,  -- queued/running/paused/done/failed/cancelled/interrupted
     params_json, progress_json, message, cancel_requested INTEGER DEFAULT 0, pause_requested INTEGER DEFAULT 0,
     heartbeat_at, created_at, updated_at)
llm_calls(cache_key PK, raw_text, parsed_json, model, params_json, structured_mode, finish_reason, tokens_in, tokens_out, latency_ms, created_at)
ai_items(id PK, template_id, stage_id, row_key, template_version_id, source_hash, context_hash, segments_hash, cache_key,
         status TEXT,   -- pending/ok/flagged/rule_only/error/skipped/outdated/excluded
         result_json, checks_json, override TEXT, attempts INTEGER, error TEXT, job_id, updated_at,
         UNIQUE(template_id, stage_id, row_key))
```
一覧表の行データはDBに持たない。`data/tables/<template_id>/state/current.jsonl.gz`・`previous.jsonl.gz`（確定済み全行）、取り込み中は `data/tables/imports/<import_id>/rows.jsonl.gz`・`issues.csv`。

## 4. モジュール構成と担当（並行実装の境界）

凡例：WP＝作業パッケージ。**各WPは自分の担当パス以外を編集しない**。他WPが必要な変更は完了報告に「統合メモ」として書く。

| WP | 担当パス | 内容 |
|---|---|---|
| WP-core | `models/database.py`, `core/__init__.py`, `core/jobs.py`, `core/files.py`, `core/naming.py`, `core/mdtext.py`, `tests/test_core_*.py` | DB（全スキーマ・マイグレーション・WAL）、ジョブ実行、アップロード保存と事前チェック、安全なファイル名、Markdown テキスト処理 |
| WP-read | `tables/__init__.py`, `tables/source.py`, `tables/csv_source.py`, `tables/excel_source.py`, `tables/detect.py`, `tables/dictionary.py`, `tables/mapping.py`, `tests/test_tables_read*.py` | 表ソース（CSV/Excel）、見出し帯・行分類・種類判定、標準キー辞書、列の対応づけ候補 |
| WP-pipe | `tables/spec.py`, `tables/normalize.py`, `tables/checks.py`, `tables/state.py`, `tables/markdown.py`, `tables/summaries.py`, `tables/outputs.py`, `tables/pipeline.py`, `tables/store.py`, `tests/test_tables_pipe*.py` | 取り込み設定の仕様、正規化、チェック、状態（期間置換・差分・取り消し）、md生成（記録・集計・データセット説明）、出力管理・zip、全体の実行関数、DBアクセス |
| WP-log | `logproc/*.py`, `tests/test_logproc*.py` | 追記ログの分割、日時解決、記入者、識別子・数量・予定句、マスク、用語集、時系列の描画 |
| WP-ai | `aiproc/*.py`, `services/llm.py`（ジョブ用呼び出し口の追加のみ。既存関数の挙動は変えない）, `tests/test_aiproc*.py`, `tests/fake_servers.py`（拡張のみ） | 構造化出力の方式判定、プロンプト生成、照合、キャッシュ、AIジョブ、custom 段 |
| WP-forms | `excel/*`, `pattern/*`, `export/formats.py`, `services/ai_assist.py`, `tests/test_extraction.py`, `tests/test_forms_md.py` | 帳票の md 改善（タイトル・ファイル名・定型文削減・値の NFKC・単位・出さない項目）、種類定義の拡張、一覧表らしさ判定関数 |
| WP-shell | `app.py`, `config.py`, `templates/base.html`, `templates/components/*`, `templates/home.html`, `templates/settings/*`, `templates/history.html`, `static/style.css`, `static/app.js`, `views/__init__.py`, `views/home.py`, `views/settings.py`, `views/history.py`, `views/files.py`（削除し core/files へ移行）| レイアウト・デザイン・ナビ・ホーム・設定（AI接続・名寄せ辞書・LightRAG案内・一覧表設定の一覧）・履歴、blueprint 登録 |
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
def precheck_excel(path) -> None      # OLE(D0CF11E0)=パスワード付き/xls、xl/workbook.bin=xlsb、Strict名前空間、zip展開上限(合計500MB/1パーツ200MB/圧縮率100) → UploadError(日本語)
def upload_path(stored_path) -> Path;  def remove_upload(stored_path) -> None

# core/naming.py
def safe_filename_part(text: str, max_len: int = 60) -> str   # NFKC、\ / : * ? " < > | 制御文字 空白 '[' ']' を _ に、'.[' を除去、前後の . _ を除去
def md_filename(parts: list[str], hint: str | None = None) -> str   # "_".join(safe parts) + (f".[{hint}]" if hint) + ".md"
LIGHTRAG_HINT_RECORDS = "legacy-R(chunk_ts=800,chunk_ol=0)"

# core/mdtext.py
def nfkc_value(text) -> str            # NFKC＋連続空白の畳み込み（改行は保持）。日本語文字間の半角空白は残す
def escape_md_line(line: str) -> str   # 行頭 # - * + > 数字. と、行全体が --- === の行、``` をエスケープ
def md_bullet(label: str, value) -> list[str]   # 単一行: ["- label: value"]、複数行: ["- label:", "  行1", "  行2"]（空行は出さない）
def estimate_tokens(text: str) -> int  # 非ASCII 1文字=1.0、ASCII 3文字=1（保守的）
def join_blocks(blocks: list[list[str]]) -> str   # ブロック間に空行1つ、末尾改行1つ、LF

# core/jobs.py
class JobContext: job_id: int; def progress(self, **kw); def heartbeat(self); def should_stop(self) -> bool; def wait_if_paused(self) -> bool; def check_cancel(self) -> None
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
    header: {anchors, rows(1|2), search_rows} ; data_end: {blank_rows:3, stop_first_col:[合計,総計], stop_prefix:[※,注]}
    exclude: {aggregate_keywords:[小計,計,合計,平均], hidden_rows:"exclude_with_warning", strike_rows:"exclude_with_warning"}
    continuation_rows: str = "merge_into_previous"   # or "keep"
    na_tokens: list[str]; fiscal_year_start_month: int = 4
    columns: list[ColumnSpec]
    crosstab: {id_columns: [...], value_key, value_display, unit, ignore_headers:[年計,合計,平均], blank_as:"no_data"} | None
    record: {key: [..], fallback_key: [occurred_at, equipment_id, "symptom:20"]}
    period: {grain: "month", date_column: "occurred_at"}
    log_stage: LogStageSpec | None; custom_stages: list[CustomStageSpec]
    markdown: {file_prefix, group_by: "month" | "entity_month", max_records_per_file: 300, lightrag_hint: False, dataset_card: True,
               records: True, summaries: [SummarySpec], title_columns: [...], omit_person: True}
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

# tables/state.py
def load_state(template_id) -> list[dict]; def replace_period(current, new_records, period, date_key) -> tuple[list[dict], DiffStats]
def save_state(template_id, records) -> None  # current→previous に回してから書く（一時ファイル→置換）
def undo_last(template_id) -> bool

# tables/markdown.py
@dataclass class MdFile: name: str; text: str; kind: str  # dataset/records/summary
def render_all(spec, records: list[dict], ai_results: dict[str, dict], meta: dict) -> list[MdFile]   # 決定的（同じ入力→同じバイト列）
# tables/summaries.py: entity_fiscal_year, month の集計（コードで計算）
# tables/outputs.py
def diff_against_delivered(template_id, files: list[MdFile]) -> {"new":[...], "changed":[...], "removed":[...], "same":[...]}
def build_zip(template_id, files, mode: "diff"|"all", extras: dict[str, bytes]) -> bytes   # RAG投入用/ と 管理用_RAGには入れない/（変更一覧.csv, 削除すべき旧ファイル.txt, 正規化データ.csv, 問題一覧.csv, 取込レポート.csv）
def mark_delivered(template_id, download_id) -> None
# tables/pipeline.py … ジョブ本体
def run_read(ctx, import_id) -> dict      # 読込→rows.jsonl.gz, issues.csv, stats
def run_confirm(ctx, import_id) -> dict   # 期間置換→md全再生成→table_outputs 更新
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
def chat_raw(settings, messages, response_format=None, max_tokens=None, timeout=120) -> ChatResult(text, finish_reason, tokens_in, tokens_out, latency_ms, headers)
def detect_structured_mode(settings) -> str   # json_schema / json_object / prompt_only（メモリキャッシュ）

# aiproc/prompts.py: build_log_messages(parse, context, spec) -> list[dict]; build_custom_messages(stage, inputs) -> list[dict]
# aiproc/verify.py: verify_log_result(result, parse, sent_text) -> VerifyReport(ok_items, failed_items, issues[level, path, message])
#   チェック: 構造/選択肢、セグメントIDを1回ずつ（隣接結合のみ）、引用(q, *_q)が根拠セグメントの部分文字列、識別子の完全一致、数量・回数、
#            出力に日付・時刻・相対日がない（引用内を除く）、人名（人物一覧・〇〇さん）がない、完了・否定語（済/未/予定/待ち/手配/なし/OK/NG/完了）の反転・欠落、推量語の消失
# aiproc/cache.py: cache_key(messages, model, params, schema) ; get/put（即時コミット）
# aiproc/runner.py: run_ai_job(ctx, import_id, scope, concurrency) ; trial_row(import_id, row_key) -> dict（1行同期）
# aiproc/estimate.py: estimate(import_id, trial_stats) -> {"rows", "calls", "tokens_in", "tokens_out", "minutes"}
```
AI の出力スキーマ（keep）：`{"entries":[{"id","segs":[...],"t":[種別...]}], "incident":{"root_cause":{"q","v","certainty","src"}, "temporary_actions":[{"v","src"}], "permanent_actions":[{"v","src"}], "parts":[{"name","model","qty_q","src"}], "recurrence":{"v","count_q","src"}, "final_state":{"v","src"}}, "summary":[...](長いレコードのみ)}`。日付・時刻・人名・数字は書かせない（数量は `_q` 引用）。照合に落ちた項目は出さない（項目単位）。1回だけ再依頼。

## 6. Markdown 出力仕様（LightRAG 調査に基づく）

共通規則（`core/mdtext.py` で実装）：
- UTF-8（BOMなし）・LF・本文に生成日時や内部IDを書かない（決定的）。
- **1レコード（帳票1件／表1行）の本文だけで意味が通る**：種別、識別番号、設備名（設備番号）、日付（ISO＋「2026年8月」）を本文に書く。
- **レコード内に空行を入れない**。レコード（見出しブロック）間は空行1つ。複数行の値は2文字下げの連続行。
- **パイプ表は使わない**。`- 項目: 値`。
- 値は NFKC＋空白の畳み込み。設備番号・設備名は名寄せ辞書で正式表記に。数値は単位付き（`停止時間: 95分`）。
- 全ファイル共通の定型文を入れない。出典は末尾1行 `- 出典: 元ファイル名（識別番号）`。
- 人名（person 役割・担当者）は既定で出さない（設定で出せる）。
- ファイル名は `core/naming.md_filename`。論理文書に対して安定・一意。`.[` `]` は除去。LightRAG ヒント（`.[legacy-R(chunk_ts=800,chunk_ol=0)]`）は設定で付けられる（既定オフ）。

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
- 長文項目の見出しに識別子を入れるのは、推定トークン数が 1,000 を超える帳票だけ（それ以外は `## 故障内容`）。
- ファイル名：`{種類名}_{タイトル項目値...}.md`。タイトル項目が空なら `{種類名}_{file_hash先頭8}.md`。
- AI入力の値には（AI入力）、手修正は印なし（確定時に人が確認済みのため）。
- 画像のセル座標・種類の版・DBの文書IDは出さない（JSON側に残す）。

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
- 1レコードの推定トークンが 1,500 を超える場合は時系列を先頭20件に切り、「（以降N件は管理用の正規化CSVに収録）」と書く。

### 6.3 一覧表（集計・説明）
- `{prefix}_00_データセット説明.md`：取り込み範囲、件数、ファイル構成、列の意味（description）、数値の注意（集計ファイルの単位でのみ確定）、答えられる/答えられない質問の例。
- `{prefix}_集計_月次_{YYYY-MM}.md`：件数、measure 合計、上位5 entity（measure 合計・件数・主な category）、category 内訳。
- `{prefix}_集計_設備別_{entity}_{FY}年度.md`：件数・measure 合計/平均、月別（0件の月も明記）、category 内訳。entity の記録がある年度のみ。
- クロス集計の設定は記録ファイルを作らず、設備別年度集計（月別値の列挙）のみ。
- 冒頭に「集計対象: 期間（取り込み範囲）」を必ず書く。数値はすべてコード計算（AI 不使用）。

### 6.4 LightRAG 案内ページ（/settings/lightrag）の内容
1. サーバーのバージョン確認（`GET /health`）。1.5.x は doc_id＝ファイル名の MD5、同名は 409 → 変更ファイルは**削除してから**再投入。1.4.x は doc_id＝内容の MD5。
2. `.env` 推奨：`SUMMARY_LANGUAGE=Japanese`、`ENTITY_TYPE_PROMPT_FILE`（1.5.x。`ENTITY_TYPES` は起動失敗）＝アプリからダウンロードできる設備保全向け YAML、埋め込みモデルは最初に固定、数万件なら PostgreSQL 等。
3. 分割：`LIGHTRAG_PARSER` 未設定なら 1,200トークン固定窓（日本語で途中切れあり）。一覧表の記録ファイルは「ファイル名ヒントを付ける」設定か、サーバー側 `LIGHTRAG_PARSER` を legacy-R にする。帳票は1チャンクに収まるので既定でよい。
4. 件数・ランキング・推移は集計ファイルで答える。記録ファイルだけでは数えられない。
5. 更新手順：差分zipの「削除すべき旧ファイル」を LightRAG で削除→idle を待つ→新規・変更を投入→アプリで[投入済みにする]。

## 7. テスト方針
- 既存テストは新ルートに合わせて書き換える（test_extraction は維持）。
- WPごとの単体テスト（純粋関数中心）。samples/ の実ファイルを使うテストは `@pytest.mark.samples`（samples が無ければ skip）。
- ゴールデン：同じ入力→同じ md バイト列（ハッシュ比較）。
- AI：`tests/fake_servers.py` を拡張し、リクエスト内容で応答を変える（壊れたJSON、原文にない型番、429、遅延）。
- 画面：Flask test client で主要フロー（帳票: upload→type→read→review→confirm→done→download、一覧表: CSV upload→source→layout→columns→preview→confirm→done→zip）。
- ブラウザ確認（統合時）：ホーム、両フロー、設定、履歴。

## 8. 実装しないもの（第2段階以降）
大きいExcel用の1パス読み込み、数式XMLの解析（小計はキーワード＋太字で判定）、同じ構造の複数シート一括、複数表の自動検出（範囲の手動指定で対応）、帳票内の明細表、full プロファイル、巨大セルの分割AI処理、抜き取り確認の統計、修正内容からのルール提案、設備台帳、帳票と一覧表の紐付け、Batch API。

### 8.1 後回し（2026-09-16 の範囲の見直しで外したもの）
統合時に、次の機能のコード・ルート・画面・テストを削除した。必要になったら 2.4・2.6・3.2・5.3・6.2〜6.4 の記述をもとに作り直す。
- **一覧表の更新管理**: 期間の置き換え（`tables/state.py` の current/previous、`replace_period`）、投入済みとの差分（`table_outputs.delivered_hash`、差分zip、「削除すべき旧ファイル」）、[投入済みにする]（`table_downloads`）、`/tables/templates/<tid>/outputs` 画面、直前の確定の取り消し。現状は「取り込みごとに全 md を作り、全ファイルを zip で渡す」（`tables/pipeline.py` の `run_read` / `run_render` / `build_download`）。md は `TABLES_DIR/imports/<id>/md/` に保存。
- **クロス集計**（縦持ち変換、`TableSpec.kind="crosstab"`、年月見出しの解釈、設備×月の値の集計ファイル）。表の形の判定（`tables/detect.py`）はクロス集計を見分けるが、範囲確認の画面で「対応していません」と止める。
- **名寄せ辞書**（`/settings/aliases`、`alias_entries` の参照、`ColumnSpec.alias_dictionary`）。値は NFKC と空白の畳み込みだけで揃える。
- **取り込み設定の版の履歴画面**（版は内部で保持: 確定に使った版は上書きせず次の版を作る）。
- **構築後のレビュー・評価フェーズ、LightRAG へのオフライン評価**（利用者が後で行う）。

DB のスキーマ（3.2）の `table_outputs`・`table_downloads`・`alias_entries`・`table_template_samples` はマイグレーション済みの既存DBとの互換のため残しているが、アプリからは使っていない。`table_imports.period_json` も書かない（常に全期間）。
