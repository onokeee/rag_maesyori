# 統合設計書：RAG用Markdown作成アプリ（帳票・一覧表）

版: 2026-09-15（作り直し v2）。この文書が実装の正本。根拠となった調査・設計は `docs/research.md`（元データの json は `docs/` 直下）に置く。

## 0. 前提と範囲

- **社内LANのサーバーで動かし、数人が同時に使う**（2026-09-20 の運用変更。それまでは1人・`127.0.0.1` のみ）。
  運用者はサーバーの JupyterLab のターミナルから `HOST=0.0.0.0 PORT=5000 python run.py` のように起動し、
  利用者は自分のPCのブラウザで `http://<サーバーのアドレス>:5000/` を開く（README「社内LANのサーバーで動かす」）。
- **ログインなし**（社内LANなので制限しない、という利用者の判断）。**アドレスを知っている人は誰でも使え、
  いま取り込み中のデータも開かれれば見える**。起動時にこの2点を必ず案内する（`app.startup_notice`）。
  待ち受け先は環境変数 `HOST`（既定 `127.0.0.1`）・`PORT`（既定 5000）で決める。
  受け付ける宛先の名前（Host ヘッダー）は、ループバック・サーバー自身の名前とアドレス・`ALLOWED_HOSTS` に
  挙げたものだけ（`app.allowed_hosts`。DNSリバインディング対策。`*` ですべて受け付ける）。
  デバッガ（`FLASK_DEBUG=1` / `--debug`、LAN のアドレスでの `DEBUG = True`）では起動しない。
- **誰の作業かはブラウザのクッキーで分ける**（`views.current_session_id()`。3.3「誰の作業か」）。
  そのため署名鍵 `SECRET_KEY` が必須（`app/__init__.py` の `create_app` が `FLASK_SECRET_KEY` か `.flask_secret` から入れる）。
  鍵が変わると全員のクッキーが無効になり、取り込み中のものが自分のものだと分からなくなる（時間切れで捨てられる）。
- **同時に動かす処理の本数**は `JOB_WORKERS`（既定 3・1〜8）。1本だと誰かの長い読み込みで全員が待たされる（5.1 `core/jobs.py`）。
- 目的：Excel/CSV を読み取り、**LightRAG に自分で投げ込みやすい Markdown(.md) を作ってダウンロードする**まで。LightRAG への送信・アプリ内検索・SQL と解析は作らない。
- 取り込みは2系統に分かれる。
  - **帳票**（1ファイル＝1件。例：設備修理報告書）… 中身は「ラベル探し」だけで読む。
  - **一覧表**（Excel/CSV、1行＝1件。例：トラブル対応一覧、故障履歴CSV、部位停止時間のクロス集計）。
- 一覧表の文章列（例：日付ごとに書き足した経過の記録が1セルに入った「対応内容」。画面の役割名は「経過の記録」、コードでは `log`／`logproc`）は、**手元で式系列に分解**し、任意で **AI整形（keep モード、原文と差分）** を重ねる。
- AI は OpenAI 互換 API（実装 `services/llm.py`）。AI なしでも全フローが成立すること。
- 対象 LightRAG: 1.5.7 だけ対象。

### 0.1 画面の作り直し（2026-09-20、利用者の指示）

利用者の言葉そのまま:

> 「UIにほしい機能は、"帳票取り込み","表の取り込み","帳票登録"の3つのみ」
> 「AIは、mdファイル作成時に、特定のバリューの値に対して、文章の構成を行う部分のみでよい」
> 「作業中の表示も要らない（閉じたら消えてよい）」
> 「①→②→③→④と進むのではなく、一画面で全部できるようにしてほしい」

これにより次を決めた。

- 画面は **3つだけ**：`/forms`（帳票取り込み）・`/tables`（表の取り込み）・`/form-types`（帳票登録）。`/` は帳票取り込みへ転送する。
- **ホーム画面は無い**。**設定の画面も無い**（AI接続はどの画面でもヘッダー右上の「AI接続」で開くパネル（2.1・2.6。2026-09-21 に AI整形の段の中から移した）、帳票の種類は「帳票登録」画面）。**一覧表の取り込み設定は保存しない**（取り込みごとに列の対応づけを決める。利用者の指示 2026-09-20）。**取り込み履歴も無い**。
- **①→②→③と画面を移らない**。1つの画面に「段」（`.step`）が下へ増えていき、進むほど次の段が現れる。送信はすべて `fetch`（`window.ragFetch`）で、画面のURLは変わらなくてよい。時間のかかる処理（大きな表の読み込み・Markdown作成）も、同じ画面の中に進み具合を出す。
- **作業中のものは捨てる**。画面を閉じたらその人の分を捨て、起動時にダウンロードしていない帳票・一覧表をすべて捨て、動いている間も2時間さわられていないものを捨てる（3.3）。閉じた画面の続きを開く入口は持たない。
- **AI は「Markdown を作るときに、決めた1つの値の文章を構成する」ところだけ**。帳票側の AI 補助（AIで種類を推定・AIで空欄を探す、`services/ai_assist.py`）は削除した。
- 読み取りの中身（ラベル探し・セクション・チェックボックス・積み上げ明細表・型と単位の推定・レイアウト判定・文字コード判定・正規化・Markdown のルール・保持期間）は**一切変えていない**。これは画面の作り直しであって、読み取りの作り直しではない。

- **範囲の見直し（2026-09-16、利用者の判断）**: 一覧表の更新管理（期間の置き換え・既存済みとの差分・取込済みにする印・出力画面・履歴）、クロス集計、名寄せ辞書は作らない（8章の「作らない」）。一覧表は**取り込みごとに、その取り込みの記録だけから全 md を作り、全ファイルを zip で渡す**。以下の章でここに触れている箇所は8章を優先する。
- **出すファイルの見直し（2026-09-20、利用者の判断）**: 一覧表で渡すのは**RAG に入れる記録の Markdown だけ**。件数・順位・推移のような**定量的な集計は RAG の仕組みにそもそも向いていない**ので、月次集計・設備別年度集計は作らない。**データセット説明も作らない**。管理用の CSV（正規化データ・問題一覧・取込レポート）も渡さない（問題一覧は確認の段の画面と CSV で見る）。zip は**フォルダ分けをしない**。

## 1. 用語（画面の文言はこの表に従う）

| 概念 | 画面の語 | 使わない語 |
|---|---|---|
| アプリ名 | RAG用Markdown作成 | Excel帳票AI前処理 |
| 帳票の取り込み | 帳票取り込み（1ファイル＝1件） | 帳票Excelのアップロード |
| 一覧表の取り込み | 表の取り込み（Excel/CSV・1行＝1件） | 表形式 |
| 帳票の読み取り設定（旧 pattern/テンプレート） | 帳票の種類 | テンプレート、パターン |
| 帳票の種類を作る画面 | 帳票登録 | パターン編集、設定 |
| 一覧表の設定（table spec） | 取り込み設定（その取り込みだけのもの。保存しない） | 表テンプレート |
| 帳票登録で置く Excel（サーバーに残さない） | この帳票のExcel | 見本、見本ファイル、サンプルExcel |
| 一覧表の、1つのセルに日付ごとに書き足した列 | 経過の記録（役割名。コードは `log`・`logproc`） | 追記ログ、AI整形の対象 |
| 一覧表の、記録が何についてのものかを表す列 | 対象（設備・製品・顧客など。コードは `entity`） | 設備 |
| 抽出 | 読み取り | 解析、前処理、抽出 |
| 帳票の確認画面 | 読み取り結果の確認・修正 | 前処理結果、抽出結果 |
| 登録 | 確定（ボタン: 確定してダウンロード） | 登録 |
| 値の出どころ | 自動で読み取り / 手で修正 | AI補完 |
| 見本 | 元のファイル（ボタン: 元のファイルをダウンロード） | 見本を開く |
| AI の列処理 | AI整形 | 前処理 |
| 1画面の中の区切り | 段（.step） | ステップ、ウィザード、手順 |
| まとめて選んだ複数の帳票 | まとめ取り込み（例:「12件中5件を確定しました」）| バッチ、一括処理 |

ボタンは動詞。1画面に主ボタンは1つ。**「次へ」は無い**（画面は移らない）。

## 2. 画面構成とルート

### 2.1 ヘッダー（base.html）

`帳票取り込み` / `表の取り込み` / `帳票登録` の3つだけ。設定・履歴・ホーム・AIモデル選択はヘッダーに置かない。

**右上の「AI接続」**（2026-09-21、利用者の指示「AI接続の設定は、もともとの位置ヘッダーの画面右上『AI接続』に移動させる」
「AI接続はヘッダー上で、接続中 か 未接続 一目で分かるように」）は行き先（画面）ではなく、その場で開くパネル
（`<dialog id="aiDialog">`、`static/app.js` の `window.ragAiHeader`）。ナビは3つのまま。ボタンに状態の灯り
「● 接続中／つながりません／未接続／未確認」と「最終確認 9/21 10:12」を出す（`views.ai_header_ctx` が
`llm.connection_status()` の**覚えている結果**を渡すだけで、画面を開くたびに AI へ確かめには行かない）。
状態の意味・確かめに行く時機・ルートは 2.6。

### 2.2 1画面の中の「段」の作り

3画面とも同じ作りにする。

- 段のしるしは `<section class="step" data-step="<名前>">` で、中に `.step-head`（番号・見出し・要約・開閉）と `.step-body`（中身）を持つ。`templates/components/_ui.html` の `ui.step(id, no, title, open=..., done=...)` マクロで書く。
- 見た目は3つ：`.step`（まだ開いていない＝灰色）、`.step.is-open`（いま作業する段）、`.step.is-done`（済み。見出しに ✓ と要約が出て畳まれ、クリックで開き直せる）。CSS は `static/style.css`。
- 開け閉ては `window.ragSections`（`static/app.js`）：`open(id)` / `done(id, 要約, 次の段)` / `close(id)`（その段を「まだ」に戻す）/ `note(id, 要約)` / `body(id)`（中身の要素）/ `working(id, 文言)`（段の中に「処理中」を出す・消す）。
- 送信は `window.ragFetch(url, options)`：`{json: {...}}`（JSONをPOST）、`{form: formまたはFormData}`（ファイルも送れる）、`{html: true}`（画面の一部のHTMLを受け取る）。サーバが JSON を返せばそのオブジェクト、HTML を返せば `{html: "…"}`。失敗はここでトーストに出してから例外を投げる（`{quiet: true}` で自分で出す）。
- 長い処理はジョブ（`core/jobs.py`）で走らせ、`GET /api/jobs/<job_id>` を同じ画面から読んで進み具合を出す。**別の画面には移らない**。

### 2.3 帳票取り込み `GET /forms`（blueprint `forms`）

1画面。上から順に段が現れる。

1. **ファイル**：帳票の Excel（.xlsx/.xlsm）をドロップ（複数可）。
2. **帳票の種類とシート**：候補（「10項目中9項目が見つかりました」）から選び、使うシートを選ぶ。
3. **読み取った値**：左に元のシート、右に項目。値を直すと Markdown のプレビューも変わる。
4. **確定してダウンロード**：`.md`（まとめ取り込みは zip）。**渡し終わるとその帳票のデータは消える**（3.3）。

読み取る前の「確定してダウンロード」は灰色のままで、見出しに「読み取りが終わると確定できます」と出る。

**まとめ取り込み（同じフォームの帳票を一気に）**（利用者の指示 2026-09-20「③の読み取り結果の表示を②タブ選択で
切り替えるのではなく、全部のシートをまとめて表示する（スクロールして全部確認していく）イメージで」）。

- ②は**置かれた分すべてで1回だけ**決める（`GET /forms/type?ids=…`）。ファイル名を丸いラベルで並べ、その種類で
  見つかった項目数を1つずつ出す（種類を選び直すと数も出し直す）。合わないファイルは ✕ で外す（その場で消える）。
  候補の並びと先に選んでおく種類は、見つかった項目の**割合ではなく点**で決める（2026-09-21）: `forms.match_score` ＝
  （見つかった項目数 ＋ シート名の一致 × `SHEET_WORTH`(2)）÷（種類の項目数 ＋ `PHANTOM_FIELDS`(10)）。分母に「見つからなかった
  ことにする項目」を 10 個足すので、4項目中4項目（100%）の小さな種類が 33項目中32項目（97%）の本物に勝てない
  （4/4 → 0.29、20/20 → 0.67、32/33 → 0.74）。まとめて置いたときは `forms.rank_batch` で「その種類が最も合ったファイルの数（票）
  → 点の平均」の順に並べる（`views._batch_matches`）。割合の平均で並べていたときは、samples/forms の 15 の版フォルダのうち
  7 つでクリックで作った 4 項目の種類が先に選ばれていた（点にして 0。`tests/test_ranking.py`）。`PatternMatch.confidence`
  （0〜100 の割合）は画面には出さず、評価の JSON（`_evaluation.json`）のために残す。「N項目中M〜K項目が見つかりました」は
  今までどおり実際の数を言う。
  読み取りは `POST /forms/read`（`ids` と種類・シート）。1件だけのときの `GET /forms/<id>/type`・
  `POST /forms/<id>/read` も同じ道すじを通る。
- ③は**帳票の数だけ塊（`.review-doc`）を縦に並べる**。塊の見出し（固定表示）にファイル名と状態
  （要確認 n件 / 未確定 / 確定済み / 修正中）、中は今までどおり左が元のシート・右が読み取った値で、その場で直すと
  その帳票だけ自動保存される。段の上に進み具合の1行（「12件中5件を確定しました」）と、まだ手の要る帳票へ飛ぶボタン。
- **重くしない**：読み取りの応答に入れる元のシートの表は**先頭の1件だけ**。残りはその塊が画面に近づいた時点で
  `GET /forms/<id>/grid` を呼んで読み込む（スクロール・窓の大きさが変わるたびに、塊の位置を自分で測る。
  画面の上下 600px 手前から読み込み、入力欄を選んだ・シートのタブを押したときもその場で読み込む）。
  `IntersectionObserver` と `requestAnimationFrame` は使わない。画面に出していない窓では動かないことがあり、
  それだと元のシートが出ないまま止まってしまうため（2026-09-20 の実測）。
  12件・1シート211行×30列の実測で、読み取りの応答は 305KB / 1.9秒（全部の表を入れると 3.0MB / 12件ぶん）。
- ④は**ボタン1つ**（「残りn件も確定して、まとめてダウンロード（zip）」）。押すと未確定の帳票をその場で全部確定してから
  zip を渡す。要確認が残っていても確定できる（要確認は目印であって止めるものではない）。1件だけのときは今までどおり
  `.md` のダウンロードで、こちらも直したまま確定していなければ先に確定し直してから渡す（3.3）。

AI による種類の推定・空欄探しは無い（0.1）。

### 2.4 表の取り込み `GET /tables`（blueprint `tables`）

1画面。上から順に段が現れる。

1. **ファイル**：Excel/CSV をドロップ。
2. **読み取り方**：CSV は文字コード・区切り・前置行、Excel はシート。変えるたびに自動で保存し、下の「表の範囲」を作り直す（取り込み設定は保存しないので、選ぶものは無い）。
3. **表の範囲と見出し行**：先頭60行の色つきグリッドで確かめ、必要なら直す。
4. **列の対応づけ**：表の上に「表の名前」（ファイル名から入れておく）を１つだけ置く。この段で決められるのは **「使う（＝Markdown に出す）」と「役割（識別番号 / 日付 / 対象（設備・製品・顧客など）/ 経過の記録（1つのセルに日付ごとに書き足した列）/ その他）」だけ**で、キー・型・単位・md での扱い・空欄＝上と同じは、見出しと値から候補づくり（`tables/mapping.suggest_columns`）が決める。保存するとその取り込みの設定（`table_imports.spec_json`）になり、読み込みが始まる。
   - **役割の名前はどんな表にも当てはまる言い方にする**（利用者の問い 2026-09-21「役割プルダウンに『設備』があるのはなぜ？」「『AI整形の対象(追記ログ)』の追記ログの意味が分からない」）。このアプリは設備の記録だけを読むものではないため、`entity` は「設備」→「対象（設備・製品・顧客など）」、`log` は「AI整形の対象（追記ログ）」→「経過の記録（1つのセルに日付ごとに書き足した列）」にした。**中で使う名前（`key` / `date` / `entity` / `log` / `attribute`）は変えていない**ので、前に保存した取り込み設定（`spec_json`）はそのまま読める。要約・上の行・エラー文では短縮名「対象」「経過の記録」を使う（但し書きまで入れると1行が読めない長さになるため。但し書きが付くのはプルダウンだけ）。
   - 表と一緒に**役割の説明（`data-role-help`）**を出す（`templates/tables/_p_columns.html`）。要約1行のときは表ごと隠れ、［変更する］で表と一緒に出る（役割を選ぶ人だけが読めばよいため）。対象＝その記録が何についてのものかを表す列で、見出しにも分割後の各かたまりにも書く。経過の記録＝例「対応内容・対応履歴・経過・対応メモ」。日付ごとに切り分けて時系列で出す（AI を使わなくても出る）。1つの表で1列だけ。
   - **決まっていれば要約1行だけ**にする（利用者の問い 2026-09-20「列の対応付けを行う意味は？」）。決めることが無いのに20行を超える表を出しても意味がないため。要約は見つけた四つの役割を名指しし、出す列と出さない列の数も書く（例:「カルテNo＝識別番号、発生日時＝日付、設備番号＝対象、対応内容＝経過の記録として読み取ります。20列のうち20列を Markdown に出します（出さない列はありません）。」）。対象や経過の記録の列が無い表では「対象の列はありません。」のようにそのまま書く。［変更する］（`static/tables.js` の `openColumnsTable`）で表が開く。
   - **「決まっている」の決まり**（`views/tables.py` の `_columns_todo`・`UNSURE_BLANK_RATE`）。次をすべて満たすときだけ要約にする。
     1. 識別番号の列がちょうど１つで、見出しが標準キー辞書と完全一致（`matched_by == "dictionary"`）
     2. 日付の列がちょうど１つで、同じく完全一致
     3. 対象の列は０か１つ。１つなら完全一致（０なら要約に「対象の列はありません」と書く）
     4. 経過の記録の列は０か１つ。１つなら完全一致（０なら要約にそう書く）
     5. 出す列のどれにも、出すかどうかを決め直す理由が無い＝読み取れない値がある（`type_error_rate > 0`）／ほとんど空欄（`blank_rate >= 0.9`）
     似た語で当たっただけ（`matched_by == "similar"`）や値の並びから当てた（`"none"`）列が四つの役割に付いていると 1〜4 で外れる。役割が合っているかは人にしか決められないため。半分くらい空欄なのは決め直す理由にしない（出しても困らないため。ここに出さなければ画面のどこにも出ない）。
   - **決まっていないときは今までどおり表**を開き、上に決めてもらうことを１行ずつ並べる（「識別番号の列が決まっていません。1つ選んでください」「列「アラーム」に読み取れない値があります（5.0%）。出すかどうか決めてください」など）。
   - **表の列は「使う」「見出し（そのまま項目名になる。読むだけ）」「役割」「値の例」の4つだけ**。5列目の「知らせ」は 2026-09-21 にやめた（ほとんどの行で空のうえ、書くことが表の上の行と重なっていたため）。型エラー％・空欄％は表の上の行に出し、**はじめから「使わない」にしてある列の理由だけ**を見出しの下に小さく出す（`views/tables.py` の `_unused_note`。「空欄だけなので、はじめから使わない設定にしています」／「記録に不要な管理用の列らしいので、…」）。理由が読めないと、チェックの外れている列を入れ直してよいのか分からないため。「値は「日付」らしい」（推定した型と決めた型の食い違い）は廃止した（日付の役割は日付として読める列にしか出ないので、読んでも直せない）。
   - 要約のときも**列の表は隠して DOM に残す**ので、［変更する］を押しても押さなくても保存で送る中身（＝できる Markdown）は変わらない（`tests/test_tables_columns_summary.py`）。
5. **AI整形（任意）**：④で「経過の記録」にした1つの列の文章を構成する。**AI接続の設定はこの段には無い**（ヘッダー右上の「AI接続」のパネル。2.1・2.6。2026-09-21 までは段の中の `details` だった）。段の中には「AI接続: 接続中（最終確認 …）／AI接続は未設定です… 設定と接続の確認は、画面右上の［AI接続］で行います」の1行だけ置く。
   **分割プレビュー**は左に原文（セルのまま。改行もそのまま）、右に「順番1、順番2…」のかたまりを並べ、同じ色で対応させる（利用者の指示 2026-09-21「順番1、順番2の表示にしてほしい。また、原文と並べて見れるようにしたい」）。`s1, s2…` の id は AI との受け渡し用で画面には出さない。`POST …/ai/split-preview` の各 segment に `no`（1始まり）と原文の中の位置 `start`/`end` を付け、`static/app.js`（表の取り込み）の `splitCompare` が描く。どちら側をクリック（Enter）しても相手側を強調して見せ、900px より狭いと縦に積む。オレンジは「推定」の印なので、対応の色には使わない。1行目は「N つに分かれました（日付や記入者がオレンジのものは、書き方から推定したものです）。AIに送るか: …」。その列が無い取り込みでは灰色のままで、見出しに「経過の記録の列がないので、この取り込みでは使いません」と1行だけ出す。段の1行目には、ルールで日付・記入者ごとに分けるところまでは AI なしで Markdown に出ること、AI整形はそのうえに原文と照合できた「対応の要点」と記録の種別を足すだけであることを書く。
6. **できるものの確認**：作られる md の一覧と中身、警告、合計の照合。
7. **確定してダウンロード（zip）**：渡し終わるとその取り込みのデータは消える（3.3）。

### 2.5 帳票登録 `GET /form-types`（blueprint `form_types`）

1画面。登録済みの帳票の種類の一覧も同じ画面に出す。

**置いた Excel はサーバーに残さない**（利用者の指示 2026-09-21）。残るのは読み取りの設定だけ（3.3「帳票登録は Excel を
1つも残さない」）。そのぶん、シートを出すには**ブラウザが持っているファイルを毎回送り直す**必要がある。

- ブラウザ（`static/form_types.js`）は選ばれた `File` を `bookFile` に持ち続け、`withBook()` がセルのクリック・項目の削除・
  見出しの手直し・使用開始のたびに `book` という部品で一緒に送る。別の種類を開いたら持ち越さない。
- サーバーは受け取った Excel を読み取り、**読み取った結果（`WorkbookInfo`）だけ**を `core/workbook_cache.py` が短い間だけ
  メモリに覚える（鍵＝作業場所の id ＋ファイルの sha256、30分・8件・64MB まで、古いものから落とす、ディスクには書かない、
  アプリを終えれば消える）。2回目からは開き直さずに済む（結合セル・図形のあるブックは開くのに数秒かかるため）。
  ファイルの代わりに `book_hash`（sha256）だけを送って、覚えているブックを指すこともできる。
- `POST /form-types/<id>/panel`（`GET` は設定だけ）が「同じ帳票の Excel をもう一度置く」入口。**種類は増えない**。
  書き方の違う版に**置き替えて**同じ欄をクリックすると、その項目の探す見出しに足される（前の「見本を何枚も預かる」の代わり）。
- 受け取る大きさの上限は `views.form_types.BOOK_MAX_BYTES = 50MB`（`MAX_CONTENT_LENGTH` とは別。保存せずメモリで読むため）。

1. **この帳票のExcel**：帳票の Excel をドロップ（1ファイル。サーバーには残らない）。
2. **名前**：ファイル名から入れておく。直せる。
3. **項目を決める**：シートのグリッドで、見出しセル → 値セル の順にクリックすると項目が1つ増える。
   「読み取る項目」の見出しはその場で手打ちで直せる（`POST /form-types/<id>/fields/<field_name>/label`）。
   直るのは **Markdown に書き出す名前（キー）だけ**で、探す見出し（候補ラベル）と読み取るセルはクリックしたときのまま。
   もとの見出しは「探す見出し: …」として行の下に残る。ほかの項目と同じ名前には直せない（重なると断る）。
4. **読み取りテスト**：その場で読み取り結果を出す（Excel を置いていないときは「できません。設定は残っています」と断る）。
5. **使用開始**：何も消さない（もともと Excel を預かっていない）。押したあとも画面はそのまま続けて使える。

保存した種類を開き直すと（`GET /form-types/<id>/panel`）、登録のときと同じ左右の画面になる。Excel を預かって
いないので左のシートの欄は置き場（`ui.file_drop`、id は `drop-book-edit` で「新しく登録する」の置き場と分ける）になり、
同じ帳票を置くと `POST /form-types/<id>/panel` でシートが出て、登録のときと同じ操作で直せる（利用者の指示 2026-09-22）。
置くまでは項目の一覧と見出しの手直しだけができ、値の欄は「—（左にExcelを置くと、この設定で読んだ値が出ます）」。
シートが出ているときは見出しの横の「別のExcelに替える」で、書き方の違う同じ帳票に替えられる。

### 2.6 設定の画面（無い）・AI接続のパネル

設定の画面も、設定の blueprint（`views/settings.py`）も無い。AI接続（接続先・APIキー・使うモデル・接続の確認）は
**ヘッダー右上の「AI接続」で開くパネル**（2.1。2026-09-21 に「表の取り込み」画面の AI整形の段の中から移した）で、
次のルートが受ける（URL は昔のまま `/tables/` の下。画面には出ず、`base.html` が `data-*` で JS に渡す）。

| ルート | 内容 |
|---|---|
| `POST /tables/ai-connection` | AI接続の保存（そのブラウザの分だけ。保存したあと、その設定で確かめて結果も覚える）。JSON |
| `POST /tables/ai-connection/test` | 接続の確認（models.list ＋ 1回の短い chat。結果を覚える）。JSON |
| `POST /tables/ai-connection/models` | API からモデル一覧を取り、「使うモデル」の入力欄の候補（datalist）として覚える。JSON |
| `POST /tables/ai-connection/clear` | ［キーを消す］（共有PC向け。そのブラウザの APIキーと確認の結果だけ消す。接続先・モデルは残す）。JSON |

- **ブラウザごと**（利用者の指示「全部空欄にしておいて cookie でユーザー毎に登録内容をずっと保持」）。保存先は
  `ai_connections` 表（持ち主は `views.current_session_id()`。3.3「誰の作業か」）。入力欄に入るのは**そのブラウザが保存した
  値だけ**で、サーバー共通の値（env・前の版の `data/model_settings.yaml`）は placeholder と注意書きで知らせるだけ。
  APIキーは画面に返さない（保存済みかどうかだけ）。パネルには「APIキーはこのサーバーに、あなたのブラウザのIDと結びつけて
  保存します（暗号化はしていません）。共有PCでは使い終わりに［キーを消す］を押してください」と書く。
- **設定の出どころ**は項目ごとに「そのブラウザの値 → 前の版の `data/model_settings.yaml`（サーバー共通・読むだけ。もう書かない）
  → env」（`llm._pick`）。ジョブのスレッドには要求（クッキー）が無いので、AI整形は開始時に `llm.job_client_settings()` で
  固めた設定を `aiproc.run_ai_job(..., settings=)` に渡す。
- **状態の灯り**（`llm.connection_status()`）: `ok` 接続中＝最後の確認がつながった／`ng` つながりません＝設定はあるが最後の
  確認がつながらなかった（パネルに理由が出る）／`off` 未接続＝キーか接続先が無い／`unchecked` 未確認＝サーバー共通の値
  （yaml・env）だけがあり、このブラウザではまだ確かめていない。確かめに行くのは**保存したとき・パネルを開いたとき・AI整形を
  始めるとき**だけ（`views.ai_run` は始める前に `llm.check_connection()` を通し、つながらなければジョブを作らず
  400「AIにつながりません（…）」を返す）。画面を開くだけでは行かない（お金がかかり、画面が待たされる）。
  試し実行で接続先に届かなかった（fatal）ときも「つながりません」にする。

旧ルート（`/documents/*`, `/patterns/*`, `/settings/*` すべて, `/history/*`, `/api/models`）は削除。
`/settings/…` を開くと 404（エラー画面に「帳票取り込みへ」が出る）。取り込み設定を保存する仕組み（一覧・選び直し・名前の変更・削除・JSON の書き出し・読み込み、`/tables/templates/*`）も無い。

エラー画面（`templates/error.html`。`code` で題と本文を切り替える1枚）はすべて日本語で、帳票取り込みへ戻るボタンを置く。403 は他サイトからの書き込みを断ったとき（`app._refuse_cross_site_write`）で、画面の JSON 送信（`Accept: application/json`）には同じ理由を JSON で返す。404 は「ダウンロード済みか、しばらく置いたままで捨てられたか、URL違い」、500 は起動した画面のメッセージを見るよう案内する。

### 2.7 UI共通部品（templates/ui.html のマクロ）
`step(id, no, title, open=)`（1画面の中の段。3画面ともこれで書く）、`card(title, subtitle, level=)`, `badge(state)`, `work_line(name)`, `progress(job)`, `data_grid(rows)`（範囲確認のセルグリッド）, `sheet_grid(grids, click_cells=)`（元の／見本のシート）, `grid_legend()`, `file_drop(name, accept)`。CSS はデザイントークン（色・余白・角丸）を `:root` 変数で統一。PC専用（`body { min-width: 1200px }`）。

### 2.8 保存先フォルダ（2026-09-19 追加 → 2026-09-20 廃止）
作った md を利用者が決めたフォルダへ直接保存する機能は削除した。受け取り口はダウンロードだけ（`services/output_folder.py`・`views/folder_save.py`・`/settings/output`・`data/output_settings.yaml` は無い）。

### 2.9 LightRAGへの取り込み案内ページ（2026-09-20 廃止）
`/settings/lightrag` と設備保全向けエンティティ種別 YAML の配布は削除した。作る Markdown は**どのサーバー設定でも安全**でなければならず、ファイル名の文字ヒント（`.[legacy-R(...)]`）には頼らない。確認の段では「zip に何が入っているか（RAG に入れる .md だけ）」を1〜2文で伝える。

## 3. データモデル（SQLite `instance/app.db`）

`models/database.py`：接続時に `PRAGMA foreign_keys=ON; journal_mode=WAL; busy_timeout=5000`。`PRAGMA user_version` による連番マイグレーション（`MIGRATIONS: list[Callable[[sqlite3.Connection], None]]`）。バックグラウンドスレッド用に `connect()`（`g` を使わない接続）を公開。

### 3.1 帳票（既存を拡張）
- `patterns`：既存列＋`title_fields TEXT DEFAULT '[]'`（タイトル・ファイル名に使う field_name の配列）、`md_options TEXT DEFAULT '{}'`（`{"domain_context_fields": [...]}` 等。入力を削る設定は置かない＝人名の項目も必ず出す）、`version_no INTEGER DEFAULT 1`（保存ごとに+1）。`status` の値は draft/active/inactive のまま（表示語だけ変更）。
- `pattern_fields`：＋`unit TEXT DEFAULT ''`、`rag_output TEXT DEFAULT 'show'`（show/omit）。`data_type` は string/text/date/number/table（明細表）。`extraction_rule` は `{"direction": "auto"}`、明細表は＋`"columns": [見本で見た列見出し]`（見出しの書き方が違う帳票で、列見出しが似た表を探すのに使う）。探す区画がある項目は＋`"section": "回答"`（区切りの見出しの名前。その区画の中だけでラベルを探す。8.0）。
- 明細表の値（`data_json` の各項目の `value`）：`{"columns": ["品番", "品名", "数量"], "rows": [["PW48-1591", "ベアリング", "2"], ...]}`。行が無ければ null。連番だけの No 列と「なし」だけの行は読まない。
- `documents`：＋`confirmed_json TEXT`、`confirmed_at TEXT`、`updated_at TEXT`（最後にさわった時刻。取り込み・読み取り・途中保存・確定で入れ直す。見回り（3.3）が「2時間さわられていないもの」を選ぶのに使う。古い行は NULL なので `COALESCE(confirmed_at, updated_at, created_at)` で見る）、`title TEXT DEFAULT ''`（一覧の見出し。タイトル項目の値を連結）、`batch_id TEXT DEFAULT ''`・`batch_order INTEGER DEFAULT 0`（まとめ取り込み。同じ `batch_id` の帳票をまとめて読み取り、縦に並べて確認し、zip でまとめて渡して一緒に消す）。`data_json` は作業中の値。状態は導出：`data_json IS NULL`→読み取り前、`confirmed_json IS NULL`→確認中、`confirmed_json != data_json`→修正中、それ以外→確定済み。`markdown`・`registered_at`・`status` 列は使わない（互換のため残す。マイグレーションで `registered_at`→`confirmed_at`、`status='registered'` の `data_json`→`confirmed_json` にコピー）。

### 3.2 一覧表
```
table_imports(id PK, template_id, template_version_id,   -- どちらも = id（設定は保存しない。下の注）
              file_name, file_hash, stored_path, session_id TEXT,
              spec_json TEXT, spec_hash TEXT,   -- この取り込みが使う取り込み設定（tables.spec.TableSpec の JSON）
              source_json TEXT,   -- {"kind":"csv|excel","encoding","delimiter","preamble_rows","sheet","header_rows","data_end_row"}
              status TEXT,        -- uploaded / reading / preview / confirming / confirmed / failed
              stats_json TEXT, issues_path TEXT, rows_path TEXT, job_id INTEGER, created_at, updated_at, confirmed_at)
jobs(id PK, kind TEXT, ref_type TEXT, ref_id INTEGER, status TEXT,  -- queued/running/paused/done/failed/cancelled/interrupted
     params_json, progress_json, message, cancel_requested INTEGER DEFAULT 0, pause_requested INTEGER DEFAULT 0,
     heartbeat_at, created_at, updated_at)
llm_calls(cache_key PK, raw_text, parsed_json, model, params_json, structured_mode, finish_reason, tokens_in, tokens_out, latency_ms, created_at,
          import_id)   -- この応答を払った取り込み（マイグレーション _m7_llm_calls_owner。古いDBの行は NULL）
ai_items(id PK, template_id, import_id, stage_id, row_key, template_version_id, source_hash, context_hash, segments_hash, cache_key,
         status TEXT,   -- pending/ok/flagged/rule_only/error/skipped/outdated/excluded
         result_json, checks_json, attempts INTEGER, error TEXT, job_id, updated_at,
         UNIQUE(import_id, template_id, stage_id, row_key))   -- 取り込みごと（マイグレーション _m6_ai_items_per_import）
```
**取り込み設定の表は無い**（`table_templates` ・ `table_template_versions` はマイグレーション `_m9_import_spec` で落とした。
利用者の指示 2026-09-20：「表の方には、取り込み設定を保持しておく機能はいらない」）。設定は取り込みの行（`spec_json`）が持ち、
取り込みを消せば設定も一緒に消える。`template_id` と `template_version_id` は取り込み自身の `id` にそろえて残してある:
AI整形の控え（`ai_items`）と `core/purge.py` がこの番号で取り込みを束ねているため。
`aiproc/runner.load_spec` は取り込みの行（`table_imports.spec_json`）から設定を読む。
`_m9` が一時的に置いた同名の互換ビューは `_m10_drop_unused` で落とした（同じマイグレーションで使わない `period_json` 列も落とす）。
**見本の Excel の控え（`pattern_samples`）も無い**（マイグレーション `_m12_drop_pattern_samples` で落とした。
利用者の指示 2026-09-21：「帳票登録で、見本のExcelは置かずに、設定だけ保持するようにしてほしい」）。
帳票の種類が持つのは `patterns` / `pattern_sheets` / `pattern_fields` の設定だけ。

一覧表の行データはDBに持たない。取り込み中だけ `data/tables/imports/<import_id>/`（`rows.jsonl.gz`・`issues.csv`・`md/`・`preview_md/`・`source_cache.json`）に置き、ダウンロードしたらフォルダごと消す（3.3）。確定済み全行の保存（`state/current.jsonl.gz`）は作らない（8.1）。


### 3.3 データを残さない（保存方針）

サーバーの容量を使わず、**他の利用者の目にも残さない**ため、**取り込んだデータはダウンロードが終わった時点で消す**。再ダウンロードはできない。
**その場でダウンロードしなかったものは、その場ですぐ捨てる**（2026-09-20 の利用者の判断）。数人が同じサーバーを
使うので、誰かの途中のデータが残り続けることが無いようにする（下の「途中のものは捨てる」）。

| | 消すもの | いつ |
|---|---|---|
| 帳票 | `uploads/documents/<保存名>`、`documents` の行（と `document_id` を持つ表の行） | `GET /forms/<id>/download.md` を**送り終えたあと**。まとまりは `GET /forms/batches/<batch_id>/download.zip`。zip に入るのは**確定済みの帳票だけ**で、消えるのもその分だけ（まだ確定していない帳票はサーバーに残り、同じ画面で続きを読める） |
| 一覧表 | `uploads/tables/<保存名>`、`TABLES_DIR/imports/<id>/`（rows.jsonl.gz・issues・控え・md・preview_md）、`table_imports` の行、`jobs`（`ref_type='table_import'`）、`ai_items`（`import_id` がこの取り込みの分）、どの `ai_items` からも参照されなくなった `llm_calls` | `GET /tables/imports/<id>/download.zip` を**送り終えたあと** |


- **残すもの**：帳票の種類（`patterns` 系）・ブラウザごとの AI接続（`ai_connections`。`core.SETTINGS_TABLES`。前の版の `data/model_settings.yaml` はもう書かない）。これらは設定であってデータではない。AI接続だけは、クッキーの寿命（約1年）より長くさわられていない行を `core.sweep_stale_ai_connections`（`AI_CONNECTION_KEEP_DAYS` = 400日。`sweep_stale` のたびに呼ぶ）が消す（もう戻って来ないブラウザの APIキーを残さない）。一覧表の取り込み設定は残さない（取り込みの行にあるので、取り込みと一緒に消える）。
- **履歴は持たない**：何をいつ取り込んだかの1行も残さない。だから取り込み履歴の画面が無い（0.1）。
  番号の続き（`sqlite_sequence`。AUTOINCREMENT の表が「これまでに使った一番大きい番号」＝取り込んだ件数を覚えている）も、
  取り込みの表（`documents`・`table_imports`・`jobs`・`ai_items`）が空になったとき乱数（2^31〜2^52）に置き換える
  （`core/purge.forget_id_counters`。消すたびの後始末 `_shrink` と起動時の片付けで行う）。AUTOINCREMENT は外さない:
  外すと消した番号がすぐ使い回され、開いたままの古い画面（途中保存・ダウンロードのボタン）が新しい取り込みを指して
  書き換えたり消したりする。乱数にすると最初の番号（1〜）とも前回の番号とも重ならない（重なる確率は1回あたり
  件数 / 約4.5×10^15）ので、古い画面の操作は 404 になる。作業中の取り込みが残っている間は番号の続きはそのまま。
  新しいDBは 1 から始まる。
- **誰の作業か（2026-09-20）**：数人が同時に使うので、取り込み中のものには「どの画面のものか」を持たせる。
  Flask のセッションクッキーに持つ番号（`views.current_session_id()`。最初に使うときに作る。クッキーは約1年もたせる＝
  `views.SESSION_LIFETIME` を `PERMANENT_SESSION_LIFETIME` に入れ、要求のたびに延びる。AI接続をブラウザごとに
  「ずっと保持」するため。2026-09-21）を `documents`・`table_imports`・`ai_connections` の `session_id` 列に入れ、
  画面と API は自分の番号のものだけを見る。
  ログインではないので**秘密ではない**（番号を知られれば開ける）。起動時の案内でもそう伝える（0）。
  その画面の分だけまとめて捨てるのが `core/purge.purge_session(session_id)`。
  `session_id` が空の行は「持ち主不明」（この仕組みより前のDBの行・テストで直接作った行）で、
  これまでどおり誰からでも扱える。セッション単位の片付けは拾わず、起動時の一括片付けだけが拾う。
  ほかの人のものを番号で指しても **404**（403 にすると「その番号はある」ことが分かるため）。
  `/api/jobs/<id>` も同じで、ほかの人のジョブ・取り込みごと消えたジョブは 404 を返し、画面は静かに止まる。
  分けないもの（みんなで共有する設定）：帳票の種類（`patterns` 系）。AI接続は**分ける**（`ai_connections`。同じ番号が持ち主。
  ほかの人のキーは使えず、見えない。取り込みを捨てる片付け（`purge_session`・`sweep_stale`・`purge_all_pending`）では消えない）。
  一覧表の取り込み設定は取り込みの行にあるので、最初からそのブラウザのものだけ（ほかの人には名前も見えない）。
- **途中のものは捨てる（2026-09-20）**：作業中の一覧もダウンロード待ちの一覧も持たない（0.1）ので、
  閉じた画面の続きを開く入口が無い。残しておく意味が無いので、**その場でダウンロードしなかったものはその場で捨てる**。
  捨てる機会は次の4つ。どれも `purge_documents` / `purge_table_import` を通る。
  1. **画面を閉じた・隠した**：ブラウザが `POST /forms/discard` / `POST /tables/discard` に
     `navigator.sendBeacon` で `{doc_ids:[…]}` / `{import_ids:[…]}` を送る（`static/app.js` の `ragDiscard`）。
     受け側は `core/purge.discard_documents` / `discard_table_imports`（番号が空なら `purge_session`）。
     閉じた（`pagehide`）ならすぐ、隠れただけ（`visibilitychange`）なら**5分待ってから**送る。
     5分の猶予があるので、LightRAG の WebUI を見に行って戻ってくる分には消えない。
     合図は届かないことがある（強制終了・LANの切断）ので、これだけに頼らず 3. と 4. で拾う。
     応答は読めないので、サーバーはいつでも 204 を返す（もう無い番号・ほかの人の番号・処理中のものは黙って外す）。
     **帳票の合図で一覧表の分まで捨てない**（同じブラウザの別のタブで作業していることがある。
     `purge_session(sid, tables=False)` / `purge_session(sid, documents=False)`）。
  2. **別の取り込みを始めた**：新しいファイルを置く・［別のファイルにする］を押すと、前の分を先に捨てる
     （画面から 1. と同じ宛先へ）。
  3. **しばらくさわられていない**：`core.purge.IDLE_HOURS`（2時間。旧名 `STALE_HOURS`）さわられていないものを
     `app.SWEEP_INTERVAL_SECONDS`（最長10分）ごとに捨てる（`core/purge.sweep_stale`、`app._start_sweeper` の
     daemon スレッド。見回りの間隔は `IDLE_HOURS` に追従し、最短60秒）。帳票も一覧表も
     `COALESCE(confirmed_at, updated_at, created_at)` で切る（取り込んだ時刻ではない。途中保存を
     続けている帳票を消さないため）。**まとめ取り込みは、そのまとまりのどれか1件でもさわられていれば
     まとまりごと残す**（50件を上から順に見ていくと、まだ手が届いていない帳票だけが画面から消え、
     zip が欠けてしまうため）。
  4. **起動時**：ダウンロードしていない帳票・一覧表を**すべて**捨てる（`core/purge.purge_all_pending`、
     `app._purge_pending`）。セッションでは分けない。**数人で使っているときにアプリを再起動すると、
     そのとき作業中だった全員の分が消える**（「その場でダウンロードしない限りその場で捨てる」方針どおりだが、
     再起動は利用者のいない時間に行う）。

  3. と 4. は**動いているジョブ（待機中・実行中・一時停止中）が付いているものを捨てない**。大きな表の読み込みや
  AI整形は2時間を超えることがあり、途中で消すとジョブが「もう無い行」を書きに行って失敗する。ジョブが終われば
  `updated_at` がその時刻になるので、次の回以降に改めて対象になる。
  1. と 2.（`discard_*` / `purge_session`）も同じく処理中のものは捨てない。

  決めたこと（2026-09-20 の統合時）:
  - **ファイルを置いたときにサーバー側で勝手に前の分を捨てない**。「その人の分をぜんぶ」捨てると、
    同じ人が別のタブで開いている作業まで消えてしまう。捨てる番号は画面が指す（上の 2.）。
    タブを再読み込みしたあとなど、画面が番号を忘れた分は 3. の時間切れで片付く。
  - **帳票登録は Excel を預からない**（2026-09-21 の利用者の指示）ので、捨てる対象がそもそも無い。
    置かれた Excel は要求の中で読み取って捨て、`uploads/` にも DB にも残さない（下の「登録中だけ残すもの」）。
  - どちらも `purge_documents` / `purge_table_import` を通るので、消えるものは上の表のとおり。
    設定（帳票の種類・AI接続）は捨てない（一覧表の取り込み設定は取り込みの行にあるので一緒に捨てる）。
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
    取り戻すには、手元の元の Excel をもう一度アップロードして読み取り・確認し直す（一覧表は列の対応づけももう一度決める。
    候補は同じ見出し・同じ値なら同じになる。AI整形の結果は消えるので、AI を使う列は再度課金される）。
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
  `ai_items` が消えたのに生の応答（`llm_calls`）が残る場面は、`core/purge.sweep_orphan_ai` がその場で消す
  （起動時の片付けでも同じ掃除をする）。
- **控えを書き戻さない**：`tables/source_cache.ImportSource` はフォルダを `__init__` でだけ作る。消したあとに
  別のタブの画面処理が控えを書いても、`imports/<id>/` は復活しない。
- **消し損ね**：起動時に、DBから参照されていないアップロードファイルと `imports/<id>/` フォルダ、もう無い取り込みを指す `ai_items`（と、それで参照されなくなった `llm_calls`）を片付ける（`app._cleanup_leftovers` → `core/files.remove_orphan_uploads` / `remove_orphan_import_dirs`、`core/purge.delete_orphan_ai_items`）。`purge_table_import` も同じ掃除をする。作業中のものは DB に行があるので消さない。`purge_table_import` で `imports/<id>/` を消し切れなかったとき（Windows で掴まれていた）は、残ったファイルの中身を 0 バイトに切り詰め（`remove_upload` と同じ。名前は残っても読み込んだ行・md は残さない）、ログに残し、次の起動時に片付く。アップロードのファイルや `imports/<id>/` を消し切れなかったことは `core/purge.purge_incomplete()` で分かる。帳票・一覧表の［削除］の案内も「一部のファイルは使用中で消し切れず、中身を空にしました。次の起動時に片付きます」にする。
- **消したあとの AI整形**：動いている AI整形のジョブは、取り込みの行が消えたら新しい呼び出しを出さず、結果も書かずに止まる（`aiproc/runner._import_gone`。ジョブの行が消えたときは `JobContext` が中止扱いにする）。一覧表の削除・確定・zip のダウンロードは AI整形の実行中は受け付けない。
- **他サイトからのダウンロード（＝削除）を断る**：消すダウンロード（`forms.download_md` / `forms.download_batch` / `tables.download_zip`）は GET でも、`Sec-Fetch-Site` が same-origin / none 以外、他サイトの `Origin`、（`Sec-Fetch-Site` が無いときは）他サイトの `Referer` なら 403（`views.PURGING_ENDPOINTS`）。
- **読み取れないアップロード**：事前チェックや読み込みで思わぬ例外が出ても、アップロードしたファイルは消してから落とす（帳票・一覧表とも）。
- **画面の知らせ**：ダウンロードのボタンには確認ダイアログ（`data-confirm`）、完了・確認画面には「ダウンロードするとサーバーからデータが消える／もう一度ダウンロードできない」の一文を出す。
  確定していない帳票（未確定・修正中）は、**ダウンロードのボタンを押した時点で確定し直してから**渡す。直した値が
  ファイルに入らずに消えることが無いようにするため（`static/review.js` の `data-confirm-all`。zip も1件の .md も同じ）。
  確認文にもそのことを書く（`views.forms._batch_zip_confirm` / `MODIFIED_DOWNLOAD_CONFIRM`）。
  消えないボタン（帳票の `original`）には「（消えません）」と書く。
- **画面のメッセージにファイル名を出さない**：`flash` は署名付きセッションクッキーとしてブラウザに残るので、
  取引先名や「社外秘」を含みうるファイル名は載せない（サーバー側を消してもブラウザに残るため）。
- **帳票登録は Excel を1つも残さない**（利用者の指示 2026-09-21「帳票登録で、見本のExcelは置かずに、設定だけ保持する
  ようにしてほしい」）。前の版は見本を `uploads/samples/<uuid>.xlsx`＋`pattern_samples` に預かり、［使用開始］で消していた。
  いまは**預からない**：置かれた Excel はその要求の中で `core/files.read_upload` がメモリに読み、`excel.workbook.load_workbook_info`
  が `BytesIO` から開いて、中身（bytes）はそのまま捨てる。`uploads/` には何も作らない（`UPLOAD_SUBDIRS` から `samples` を外した）。
  `pattern_samples` 表はマイグレーション `_m12_drop_pattern_samples` で落とす。［使用開始］は何も消さない（もともと何も無い）。
  起動時の `core/purge.purge_old_sample_files()`（`core.files.remove_sample_dir`。前の `purge_all_samples` と入れ替え）は、
  **前の版が残した `uploads/samples` フォルダだけ**を片付ける。
  帳票の種類が持つのは設定だけ：シート名・見出しのセル・値のセル・読み取る向き・項目名。
  一覧表の「表の名前」は取り込みの行にしか無く、取り込みと一緒に消えるので、ファイル名を初期値にしてよい。
  - werkzeug の受け取り方だけは残る穴：アップロードの1つの部品が 500KB を超えると、超えた分を OS の名前の無い一時ファイルに
    逃がし、要求が終わると消す（Windows では `FILE_FLAG_DELETE_ON_CLOSE`）。このアプリのアップロードは前からすべてそうで、
    こちらのコードは何も書かない。実物の帳票は 20〜110KB なので実際には逃げない（`views.form_types._read_book` の説明に書いてある）。
- **新しい表**：取り込みを指す列は必ず `document_id` / `import_id` という名前にする（purge の探索に乗せるため）。
  使わない表はそもそも作らない（`table_outputs`・`table_downloads`・`table_template_samples`・`alias_entries` はスキーマから外し、
  古いDBのためにマイグレーション `_m5` で落とす）。

## 4. モジュール構成と担当（並行実装の境界）

> **2026-09-21 ファイル数の最小化**：Python のパッケージは領域ごとに1ファイルにまとめた（`core/*` → `core.py`、`models/database.py` → `database.py`、
> `excel/*`・`pattern/*`・`export/formats.py` → `forms.py`、`tables/*` → `tables.py`、`logproc/*` → `logproc.py`、`aiproc/*` → `aiproc.py`、
> `services/*` → `llm.py`、`views/*` → `views.py`、`config.py` → `app.py`）。アプリ本体はさらに `app/` フォルダにまとめ（`app.py` は `app/__init__.py`、
> `templates/`・`static/` も `app/` の中）、ルートは起動用の `run.py`・テスト（`tests/`）・サンプルの生成器（`scripts/`）・設計書（`docs/`）だけにした。1ファイルの中は「元 core/jobs.py」の見出しで旧モジュールごとに区切ってあり、
> 以下の表と 5 章の `core/jobs.py` のような名前は、その見出し（ファイルの中の節）を指す。画面は `templates/<画面>.html` の1ファイル
> （段の中身の断片はその中のマクロ `part_*`。`views.render_part` が描く）、JS は `static/app.js` の1ファイル（画面ごとに「// ==== 」で区切る）、
> テストは `tests/test_forms.py`・`test_tables.py`・`test_ai.py`・`test_core.py`・`test_cross.py` と `conftest.py`（偽サーバー・表の操作を含む）。

凡例：WP＝作業パッケージ。**各WPは自分の担当パス以外を編集しない**。他WPが必要な変更は完了報告に「統合メモ」として書く。

| WP | 担当パス | 内容 |
|---|---|---|
| WP-core | `models/database.py`, `core/__init__.py`, `core/jobs.py`, `core/files.py`, `core/naming.py`, `core/mdtext.py`, `core/workbook_cache.py`, `tests/test_core_*.py` | DB（全スキーマ・マイグレーション・WAL）、ジョブ実行、アップロード保存と事前チェック、帳票登録が置いた Excel の一時的な覚え、安全なファイル名、Markdown テキスト処理 |
| WP-read | `tables/__init__.py`, `tables/source.py`, `tables/csv_source.py`, `tables/excel_source.py`, `tables/detect.py`, `tables/dictionary.py`, `tables/mapping.py`, `tests/test_tables_read*.py` | 表ソース（CSV/Excel）、見出し帯・行分類・種類判定、標準キー辞書、列の対応づけ候補 |
| WP-pipe | `tables/spec.py`, `tables/normalize.py`, `tables/checks.py`, `tables/markdown.py`, `tables/records.py`, `tables/outputs.py`, `tables/pipeline.py`, `tables/store.py`, `tests/test_tables_pipe*.py` | 取り込み設定の仕様、正規化、チェック、md生成（記録ファイルだけ）、zip、全体の実行関数、DBアクセス（`tables/state.py` は 8.1 で外した。`tables/summaries.py` は 6.3 で外した） |
| WP-log | `logproc/*.py`, `tests/test_logproc*.py` | 「経過の記録」の列（コードでは `log`／`logproc`。画面の役割名は 2.4 ④）の分割、日時解決、記入者、識別子・数量・予定句、マスク、用語集、時系列の描画 |
| WP-ai | `aiproc/*.py`, `services/llm.py`（ジョブ用呼び出し口の追加のみ。既存関数の挙動は変えない）, `tests/test_aiproc*.py`, `tests/fake_servers.py`（拡張のみ） | 構造化出力の方式判定、プロンプト生成、照合、キャッシュ、AIジョブ、custom 段 |
| WP-forms | `excel/*`, `pattern/*`, `export/formats.py`, `tests/test_extraction.py`, `tests/test_forms_md.py` | 帳票の md 改善（タイトル・ファイル名・定型文削減・値の NFKC・単位・出さない項目）、種類定義の拡張、一覧表らしさ判定関数 |
| WP-shell | `app.py`, `config.py`, `templates/base.html`, `templates/components/_ui.html`, `templates/errors/*`, `static/style.css`, `static/app.js`, `views/__init__.py` | レイアウト・デザイン・ヘッダー（3画面だけ）・段（`.step` と `ragSections`）・エラー画面、blueprint 登録。ホーム画面・設定画面・`views/home.py`・`views/settings.py`・`templates/home.html`・`templates/settings/*` は削除済み |
| WP-formsui | `views/forms.py`, `views/form_types.py`, `templates/forms/*`, `templates/form_types/*`, `static/review.js`, `tests/test_forms_flow.py` | 帳票フロー画面と帳票の種類の管理画面 |
| WP-tablesui | `views/tables.py`, `templates/tables/*`, `static/tables.js`, `tests/test_tables_flow.py`, `tests/test_tables_columns_summary.py` | 一覧表の1画面（読み取り方・範囲・列の対応づけ・AI整形・確認・ダウンロード） |
| WP-samples | `scripts/samples/*`（追記のみ）, `samples/`（生成物） | T1 に「対応内容」（経過の記録）の列を追加（書き方の揺れを再現）など |

依存：WP-core・WP-read・WP-log・WP-forms は並行（第1波）。WP-pipe・WP-ai は第1波の後（第2波）。WP-shell は第1波と並行可（テンプレートのみ）。WP-formsui・WP-tablesui は第2波の後（第3波）。

## 5. 主要インターフェース

### 5.1 core
```python
# core/files.py
class UploadError(Exception): ...
@dataclass class StoredFile: stored_path: str; file_name: str; file_hash: str; size: int
@dataclass class MemoryFile: data: bytes; file_name: str; file_hash: str   # size は len(data)
def save_upload(storage, subdir: str, allowed: set[str], max_bytes: int) -> StoredFile   # 分割読みで sha256
def read_upload(storage, allowed: set[str], max_bytes: int) -> MemoryFile   # 保存せずメモリへ（帳票登録。3.3「帳票登録は Excel を1つも残さない」）
def precheck_excel(source, max_cells=None, max_merged=None, max_rows=None) -> None   # source はパスまたはブックの中身(bytes)。 OLE(D0CF11E0)=パスワード付き/xls、xl/workbook.bin=xlsb、Strict名前空間、zip展開上限(合計500MB/1パーツ200MB/圧縮率100。圧縮率はブック全体でも見る＝小さいパーツを並べて展開させない)、
                                      # 結合セルの面積の合計(200万セル。帳票と帳票登録で置く Excel は FORM_MAX_MERGED_CELLS=20万。通常モードで開くと結合1つごとに全セルをたどるため)、セル数（<c> の数。帳票と帳票登録で置く Excel は EXCEL_MAX_CELLS=50万、省略時 100万）、
                                      # ハイパーリンク・コメントの範囲（ref）の面積の合計(ブック全体で MAX_LINKED_CELLS=5万。A1:XFD1048576 で固まらないように)、
                                      # <row> の数(max_rows。省略時は max_merged と同じ。帳票と帳票登録で置く Excel は20万行、一覧表は200万行)、
                                      # 図形(描画パーツの図形数・大きさ×参照するシート数)、シートが1つも無いブック、壊れた圧縮データ(zlib.error/EOFError) → UploadError(日本語。ファイル名・例外の種類名は入れない)
                                      # 上の結合・セル・行・リンクの数え方: 1つのパーツを複数の <sheet> が指していると openpyxl はその回数だけ読み直すので、参照される回数を掛けて数える
def upload_path(stored_path) -> Path;  def remove_upload(stored_path) -> bool   # 消せたら True。掴まれて消せないときは中身を0バイトにして False
def remove_orphan_uploads(upload_dir, known: set[str]) -> int;  def remove_orphan_import_dirs(tables_dir, import_ids: set[int]) -> int
def remove_sample_dir(base) -> int   # 前の版が置いた uploads/samples をフォルダごと消す（起動時の片付けだけに使う）

# core/purge.py（3.3 データを残さない）
def purge_documents(doc_ids) -> int;  def purge_batch(batch_id: str) -> int;  def purge_table_import(import_id: int) -> int   # 戻り値は消した行数
def purge_after_send(response, fn, *args) -> response   # 本文を最後まで送り終えたときだけ消す（途中で切れたら消さない）
# 消す表は sqlite_master から document_id / import_id 列を持つ表を探して決める（新しい表はこの列名にする）
# 消したあと: PRAGMA wal_checkpoint(TRUNCATE) と VACUUM（中身のゼロ埋めは接続時の PRAGMA secure_delete = ON）
def forget_id_counters(db) -> int   # 空になった取り込みの表の sqlite_sequence を乱数にする（番号の続きを履歴にしない）
# 使っている人ごとに捨てる（3.3 の 1.・2.）。処理中のもの・ほかの人のもの・もう無い番号は黙って外す（何度呼んでも安全）
def discard_documents(doc_ids, session_id=None) -> int;  def discard_table_imports(import_ids, session_id=None) -> int
def purge_session(session_id, *, include_busy=False, documents=True, tables=True) -> tuple[int, int]
IDLE_HOURS = 2   # 旧名 STALE_HOURS（別名として残す）。def sweep_stale(hours=IDLE_HOURS) -> tuple[int, int]
def purge_old_sample_files() -> int   # 前の版が置いた uploads/samples を起動時に片付ける（旧 purge_all_samples）

# core/workbook_cache.py（帳票登録。置いた Excel を「読み取った形」だけメモリに覚える。ディスクには書かない）
@dataclass class Book: file_name: str; file_hash: str; size: int; info: WorkbookInfo; used_at: float
def put(session_id, book) -> Book;  def get(session_id, file_hash) -> Book | None;  def clear() -> None;  def count() -> int
TTL_SECONDS = 30*60;  MAX_ENTRIES = 8;  MAX_BYTES = 64MB   # 古いもの→入れた順に落とす。出し入れは Lock の中（waitress は8スレッド）

# core/naming.py
def safe_filename_part(text: str, max_len: int = 60) -> str   # NFKC、\ / : * ? " < > | 制御文字 空白 '[' ']' を _ に、'.[' を除去、前後の . _ を除去
def md_filename(parts: list[str]) -> str   # "_".join(safe parts) + ".md"（LightRAG のヒントは付けない）

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
# 実装: 列（lane）ごとに JOB_WORKERS 本（app.config["JOB_WORKERS"]、既定 3・1〜8 に丸める）のワーカースレッド＋キュー。
#   数人が同時に使うので1本だと誰かの3万行の読み込みでほかの人が待たされる。同じ取り込みのジョブは本数が増えても同時に動かさない。
#   本数はプロセス内でその列を最初に使ったときに決まり、あとから減らない。
#   fn 内で DB を使うときは database.connect()。app_context は start_job 時に app を捕まえて with app.app_context() で実行
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
def list_kind(source, sheet) -> str   # 行が並ぶ表か。"list" / "crosstab" / ""（どちらでもない）
def kind_from_layout(layout) -> str   # 上の判定のうち「表の形の見立てから決める」部分。控えを使う ImportSource.list_kind と共通
def split_header_unit(header: str) -> tuple[str, str]   # "停止時間(分)" → ("停止時間","分")

# tables/dictionary.py  … 表用の標準キー辞書（帳票の pattern/dictionary.py は変更しない）
STANDARD_COLUMNS: list[StdColumn(key, display, type, role, synonyms)]
# 例: record_no(code,key) occurred_at(datetime,date) equipment_id(code,entity) equipment_name(string,entity_label) line process
#     failure_category(enum,category) severity symptom(text) cause(text) action(text,log候補) response_log(text,log) downtime(number,measure,分) work_hours cost status worker(string,person) part_name quantity unit_price

# tables/mapping.py
@dataclass class ColumnSuggestion: index: int; header: str; unit: str; key: str | None; display: str; type: str; role: str; examples: list[str]; type_error_rate: float; blank_rate: float; md: str  # body/attribute/omit
    fill_down_blank: bool = False; log: bool = False; matched_by: str = ""  # dictionary/similar/none
def suggest_columns(headers, sample_rows) -> list[ColumnSuggestion]
```

### 5.3 tables（パイプライン）
```python
# tables/spec.py … 取り込み設定（取り込みの行に JSON で持つ。dataclass ⇔ dict、validate、spec_hash）
@dataclass class ColumnSpec: key: str; display: str; headers: list[str]; type: str   # code/string/text/date/datetime/time/number/enum/status
    role: str = "attribute"   # key/date/entity/entity_label/category/measure/text/log/person/attribute
    unit: str = ""; unit_conversions: dict = {}; required: bool = False; md: str = "attribute"  # body/attribute/omit
    fill_down_blank: bool = False; normalize: list[str] = ["nfkc"]; alias_dictionary: str | None = None
    allowed: list[str] = []; value_map: dict = {}; description: str = ""
@dataclass class LogStageSpec: column: str; enabled_ai: bool = False; context_columns: list[str]; people: list[dict]; groups: list[str]; glossary: dict; entry_types: list[str]; instruction: str = ""; incident: bool = True; run_if: dict; limits: dict
@dataclass class CustomStageSpec: id: str; inputs: list[str]; prompt: str; output_type: str  # text/choice
    choices: list[str] = []; max_chars: int = 80; fallback: str = "不明"; target_key: str = ""; quote_required: bool = True
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
    markdown: {file_prefix, group_by: "month" | "entity_month", records: True, title_columns: [...]}
    # 2026-09-20 に外した項目（保存済みの JSON にあっても読み飛ばす。RETIRED_KEYS["markdown"]）:
    #   lightrag_hint（ファイル名のヒントは付けない）、dedupe_timeline・omit_person（内容を削らない）、
    #   max_records_per_file（記録ファイルは月ごと・件数では分けない）、
    #   dataset_card・summaries（出すのは記録ファイルだけ。集計・説明は作らない。6.3）
    checks: {type_error_rate: {warn: 0.02, block: 0.10}}
def spec_from_dict(d) -> TableSpec; def spec_to_dict(spec) -> dict; def spec_hash(spec) -> str; def validate_spec(spec) -> list[str]

# tables/normalize.py
@dataclass class RecordRow: key: str; values: dict[str, object]; originals: dict[str, str]; source: {"file","sheet","row"}; warnings: list[str]
def read_records(source, source_opts: dict | None, layout: LayoutGuess, spec: TableSpec, on_progress=None) -> tuple[list[RecordRow], list[Issue], ImportStats]
# 型変換は excel/text.py を拡張して使う（和暦 令和/平成/昭和、年なし M/D を年度で補完、8桁日付、6桁時刻、△▲・末尾マイナス、桁区切り、%、timedelta→分、単位換算）

# tables/checks.py
@dataclass class Issue: level: str  # error/warning
    code: str; message: str; row: int | None = None; column: str | None = None
def run_checks(records, spec, stats) -> list[Issue]

# tables/state.py は作っていない（期間の置き換え・取り消しは 8.1 で外した）

# tables/markdown.py
@dataclass class MdFile: name: str; text: str; kind: str = "records"  # 作るのは記録ファイルだけ
def render_all(spec, records: list[dict], ai_results: dict[str, dict]) -> list[MdFile]   # 決定的（同じ入力→同じバイト列）
def record_blocks(record, spec, ai_results=None, people=None) -> list[list[str]]   # 1件分。大きければ「（続きn/m）」に分ける（6.2）
# tables/records.py: 記録の値の見方（設備の列 entity_value/entity_display/split_entity_code、月、数値の書き方 fmt_number）
# tables/outputs.py（投入済みとの差分は持たない。8.1）
def issues_csv(issues) -> bytes   # 確認の段の画面から見る問題一覧（zip には入れない）
def build_zip(md_files: list[tuple[str, bytes]]) -> bytes   # RAG に入れる .md だけ（フォルダ分けなし）
# tables/pipeline.py … ジョブ本体
def run_read(ctx, import_id) -> dict      # 読込→rows.jsonl.gz, issues.csv, stats
def run_preview(ctx, import_id) -> dict   # 確認画面用の md の下書き（preview_md/）
def run_render(ctx, import_id) -> dict    # 確定→その取り込みの全 md を md/ に作る
def build_download(import_id) -> bytes   # zip 全体をバイト列で返す（遅延生成にしない。3.3）
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
# services/llm.py（設定の解決は項目ごとに「そのブラウザ（ai_connections） → 前の版の yaml → env」。2.6）
def connection_status() -> dict     # ヘッダーとパネルに出す状態（state ok/ng/off/unchecked・state_label・sub_text・入力欄の値・effective・fallback）。キーの値は返さない
def save_browser(session_id, data) -> dict ; clear_browser_key(session_id) -> dict   # そのブラウザの分だけ保存／キーを消す
def check_connection() -> (ok, steps) ; record_check(session_id, ok, steps)          # models.list ＋ 短い chat、結果を覚える
def job_client_settings() -> dict   # base_url, api_key, model, 固定（ジョブ開始時）。fingerprint（キー除く）。ジョブのスレッドには要求が無いので開始時に渡す
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
- **入力の内容は削らない・書き換えない**（2026-09-20 の利用者の指示）。出す列・項目の中身はそのまま書く：人名（person 役割・担当者・押印欄）もコードの列（`- 状態コード: 9`）も出す。出さないのは、画面で「出さない」にした列・項目と、値が空欄だけの列だけ。組み替え（1行＝1レコード、`- 項目: 値`、長い経過の記録を1語も落とさず時系列に分ける）と表記の正規化（和暦→西暦・単位・NFKC）は続ける。
- ファイル名は `core/naming.md_filename`。論理文書に対して安定・一意。`.[` `]` は除去。**LightRAG のファイル名ヒント（`.[legacy-R(...)]`）は付けない**：ヒントはサーバー側の取り込み設定より優先され、そのサーバーが知らない書き方だと取り込みが HTTP 400 で断られる。チャンクへの耐性は本文の作り方（6.2 の記録の分割）だけで成り立たせる。
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
  - 読み取り結果では表の形で修正できる（行の追加・削除）。承認欄（区分×作成/確認/承認）・なぜなぜ分析・5W2H は明細表にしない（項目として読む）。
- 1つの値の項目の探し方（`excel/extractor.py`）：ラベル候補に一致するセルを上→下・左→右に見て、その右、無ければ下の値を取る。順番は
  「候補どおりのラベル」→「表記だけ違うラベル」（末尾の括弧書き・「内容」「欄」「日時」の違い。例「発生原因（なぜ起きたか）」「応急処置内容」「復旧完了」。
  括弧書きどうしが違う「原因（推定）／原因（確定）」は別項目として扱い、表記違いの照合は塗りつぶし・太字・「ラベル：値」のセルだけ）、
  それぞれで「すぐ隣に値があるラベル」を先に見る（押印欄の縦書き「発信部署」の2行下にある日付を値にしない）。最後に明細表の列見出しにあるラベルを見る。
  - 値にしないもの：行が並ぶ明細表の列見出し（「設備No｜設備名」の下に3行以上）、表の連番の列見出し（No）、区切りの見出し（「▼ 回答欄」）、欄外の様式番号（「様式MT-031 Rev.1」）。
    押印欄（承認｜確認｜作成の下に印と日付の2行）は明細表とみなさず、今までどおり項目として読む。
  - 「対象設備／使用設備／設備」のように設備番号と設備名を1つのセルに書く欄（「CMP-108　STI-CMP 8号機」「ROB-821（ウェーハソーター 1号機）」）は、
    番号と名前に分けて設備番号・設備名の項目に入れる（分けられない値ならその欄は使わない）。クリックでもこの欄は1回のクリックで
    設備番号・設備名の2項目になる（`pattern/clicks.split_rows`。文字は消さず、同じセルの読む場所を分けるだけ）。
  - **版によって書き方が変わる帳票**（同じ種類の帳票で、見出しの位置やセル内の書き方が違うもの）の読み方：
    - 見本でクリックした向き（右・下）を先に見て、無ければもう一方も見る。どちらにも無く、見出しのセルが「設備番号　：IMP-603」のように
      1つのセルに見出しと値を持つときは、そのセル内の値を読む（`excel/extractor._value_at`）。
    - 見出しがどうしても見つからない帳票では、見本でクリックしたセルの**番地**を控えとして読む（`_extract_at_cell`）。ただし様式が違うと
      同じ番地に別の欄が来るので、**その番地のセルが他の欄の見出しのとき（`stop_labels`）と、その型として読めないときは値にしない**。
      読めないときは空のままにする（当て推量で埋めるより、空欄のほうが直しやすい）。見出しの無い「値だけ」の項目は、最初から番地で読む。
  - 「3時間40分」は項目の単位（分・時間）に換算する。単位が決まっていない項目では分にする。
- **数値項目の単位（2026-09-19 の利用者の判断）**：単位は「帳票の種類の設定」→「書かれた値（「390分」「2.5h」「1,032分」）」の順に決める（`excel/text.numeric_unit`）。既定の単位は辞書に持たない（勝手に決めない）。
  - 両方から決まらない項目のうち、**単位で意味が変わると辞書が知っているもの**（時間・工数・金額・寸法・重量・温度・圧力・流量・電流電圧＝`pattern/dictionary.AMBIGUOUS_UNIT_HINTS`）は要確認にし、「単位が書かれていません（分か時間かで意味が変わります）。値に単位を付けて入力してください（例: 1456分）。これから読み取る帳票のためには「帳票の種類」の画面でこの項目の単位を決めてください（読み取り済みの帳票には反映されません）」と出す（種類の単位を変えても読み取り済みの帳票は読み直さないので、その場で直せる方法を先に書く）。
    件数・回数・人数・枚数・率など、単位が無いのが普通の項目は警告しない（確認画面が警告で埋まらないようにするため。単位の無い数値をすべて警告にはしない）。
  - 種類の単位と書かれた単位が違うとき（「停止時間（分）」の欄に「14.9h」）は、**書かれたとおりの単位で出して**要確認にする（勝手に換算しない。「14.9分」と書くと誤った事実になる）。
  - セルの値が数値だけで、表示形式に単位が書かれている（`#,##0"分"` で画面には「3,095分」）ときは、その単位を値に書かれた単位として扱う（`excel/text.format_unit`）。
  - 「2:45」や時刻・[h]:mm のセルは時:分の時間として項目の単位（分・時間。無ければ分）に換算する。範囲（「10～20分」）で単位が項目の単位と違うときは数値にしない。数値の部分だけを読んだ値・数値として読めない値は、手で入力した値でも要確認のままにする。
  - Excel のエラー値（「#REF!」「#N/A」「#DIV/0!」）は値にせず、空にして要確認にする（一覧表の側と同じ）。
  - 「390分」のように数値の後ろが単位だけのときは、単位として取り込むので「数値の部分だけを読み取りました」の警告は出さない。
  - この判断は `excel/extractor.number_unit` 1か所で行う（読み取り `_apply_number_unit` から呼ぶ。経路で単位の付き方が変わらないようにするため。
    帳票の AI 補助＝[AIで空欄を探す]（`services/ai_assist.py`）は 2026-09-20 に削除したので、呼び口はこの1つだけになった）。
- **年の無い日付（2026-09-19 の利用者の判断）**：「2/12 3時17分」のような年の無い日付は**年を補わない**。値は書かれたままの文字列で残し、要確認にして「年が書かれていません。元のファイルを確かめて、2026-02-12 のように年から書いてください」と出す（`excel/text.to_date`）。
- **日付のあとの時刻・続く文字**：「2024年7月29日 13:41」「R5.11.16 12:11」「2023-07-10T23:08」「2023年7月10日 23時08分」の時刻は残す（年月日・和暦・曜日「(月)」のあとでも）。日付と時刻のあとに別の文字が続く（「2023/7/10～7/12」）ときは日付だけを読み、「日付のあとに「～7/12」が続いています」と要確認にする（範囲の終わりを黙って落とさない）。和暦は令和（R）と平成（H）を読む。
  同じ帳票の別の項目やファイル名から年を推すと、外れたときに Markdown に誤った日付を書くことになるため。「24/8/25」のような2桁の年は年の欄が空とは限らないので、今までどおり「日付として解釈できません」。

### 6.2 一覧表（記録ファイル）
ファイル単位の既定：**発生年月ごと**（`{prefix}_{YYYY-MM}.md`）。**件数では分けない**（その月の記録は1ファイルに全件。LightRAG は1ファイルを丸ごと読んでからチャンクに分けるので、ファイルを分けても記録の切られ方は変わらない）。設定で「対象×月」（`markdown.group_by = "entity_month"`）も選択可（`{prefix}_{対象の値}_{YYYY-MM}.md`）。画面からは選べない。日付が空の行は `{prefix}_日付なし.md`。
```
# トラブル対応一覧 2026年8月の記録

- データ種別: トラブル対応一覧（1行＝1件）の記録
- 対象期間: 2026-08-01〜2026-08-31
- このファイルの記録: 342件（2026年8月の全件）

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
- 時系列の文は1つも省かない（他の列と同じ文でも、件数が多くても、書かれたまま出す）。
- **記録1件の上限は推定トークン 400**（`tables/markdown.RECORD_TOKEN_BUDGET`）。超えたら「（続きn/m）」に分ける（6.5 の実測で決めた値）。
  - 見出しは1つ目が `## …（1/m）`、2つ目以降が `## …（続きn/m）`。2つ目以降には**管理No・設備・日付の行を書き直す**（切られてもその部分だけで身元が分かる）。
  - 分け目は箇条書きの切れ目。1つの箇条書きが上限を超えるときは、複数行の値なら `- 対応の時系列（続き）:` と見出し行を繰り返し、1行の値なら「。」の後ろで切って `- 項目（続き）: …` にする。**文字は1つも消さない**（1行が丸ごと上限を超えるときは、その1行だけで1つの部分にする）。
- 設備名の列がない code 型の entity 列では、「ETC-302(OXIDEエッチャ 2号機)」「CVD-203 W-CVD 3号機」を設備番号と名前に分けてから、ファイル分け・表示に使う（番号は英字と数字を含むものだけ。名前の側も番号だけなら分けない。`tables/records.split_entity_code`）。

### 6.3 一覧表で作るファイル（2026-09-20: 記録ファイルだけにした）
- 作るのは `{prefix}_{YYYY-MM}.md`（対象×月の設定なら `{prefix}_{entity}_{YYYY-MM}.md`）だけ。**集計ファイル（月次・設備別年度）とデータセット説明は作らない**。
  - 理由（利用者の判断）: 「件数」「順位」「推移」には答えなくてよい。**定量的な集計や計算は RAG の仕組みにそもそも向いていない**。データセット説明も要らない。
  - `tables/summaries.py`（month_summaries・entity_fiscal_year_summaries・dataset_counts と、その計算だけに使っていた resolve_metrics・fmt_average・coverage_months・measure_columns・category_column・fmt_measure・fiscal_year_of など）と `tables/markdown.render_dataset_card` は削除した。記録ファイルにも使う処理（設備の値・月・数値の書き方）は `tables/records.py` に残している。
  - 取り込み設定の `markdown.dataset_card` / `markdown.summaries`（`SummarySpec`）も外した（保存済みの JSON にあっても読み飛ばす）。
- 渡すのは zip 1つだけで、中身は **RAG に入れる .md だけ**（`RAG投入用/`・`管理用_RAGには入れない/` のフォルダ分けはしない）。**正規化データ.csv・問題一覧.csv・取込レポート.csv は作らない**し、画面にも単独の CSV ダウンロードは置かない（確認の段の「問題一覧（CSV）」だけは、200件を超える問題を確かめるために残す）。
- 「内容の確認」の段は今までどおり、件数・除外した行・問題一覧・作られるファイルの一覧と中身の下見・データ100行ずつを出す。
- 遠い年の日付（「2052年」のような入力ミス）は、確認画面で、記録の年の中央値から5年を超えて離れた日付を行番号つきで知らせる（警告 `date_outlier`）。記録ファイルは記録のある月にだけできるので、空の月のファイルは増えない。

### 6.4 LightRAG への入れ方（2026-09-20: 案内ページは廃止。2.9）
このアプリは LightRAG に送信しない。利用者は完了画面で md／zip をダウンロードし、LightRAG の WebUI にドラッグして入れる。
zip の中身は RAG に入れる .md だけなので、開いてそのままドラッグすればよい。
利用者はサーバー設定を見ることも変えることもできないので、Markdown は**どのサーバー設定でも安全**に読めるよう本文の作り方（記録の分割）だけで成り立たせ、ファイル名の分割ヒントには頼らない（付けない）。

### 6.5 記録の上限を決めた実測（2026-09-20、`scripts/eval/lightrag_offline_eval.py`）
見本 T1・T2・T5（46,233件）の月ごと md を、LightRAG のチャンク分割に**オフラインで**かけた（LLM・API は呼ばない）。
経路は **F1200/100**（`LIGHTRAG_PARSER` 未設定時の既定の固定窓）、**F600/50**（管理者が窓を狭めた場合の想定）、**P2000**（native パーサ＋段落）。
「before」＝ 2026-09-19 版（人名・コードの列を落とし、時系列の重複を省き、ファイル名にヒントを付けていた版）。

| 版（記録1件の上限） | 総トークン | ブロック数（1件あたり） | 途中切断 F1200/100 | 途中切断 F600/50 | 管理Noの無いチャンク（F600/50） | 設備・日付の無いチャンク（F600/50） |
|---|---|---|---|---|---|---|
| before（1,400・時系列を切る） | 21.4M | 46,233（1.0） | 32% | 61% | **2,950** | 1,932 |
| 300 | 47.8M | 231,810（5.0） | 9% | 28% | 0 | 137 |
| **400（採用）** | 39.3M | 151,603（3.3） | 14% | 37% | 4 | 138 |
| 600 | 32.3M | 86,863（1.9） | 24% | 57% | 1 | 148 |

- 一番効いたのは上限の値ではなく**分けた部分に管理No・設備・日付を書き直すこと**。狭い窓（600）でも身元の分からないチャンクが 2,950 → ほぼ0 になった。P2000 ではどの版も記録を途中で切らない。
- 400 を選んだ理由：狭い窓の1チャンクに新しく入るのは 600−50＝550トークンなので、**1つの部分が必ず1チャンクに収まる大きさ**にする必要がある（600 だと収まらず、切断が 57% に増える）。推定式は実トークン以上に出るので、550 に対して 400 なら余裕がある。300 にすると切断は9ポイント減るが、トークンが 22%・ブロック数が 53% 増える（＝取り込み側の LLM 呼び出しが増える）。
- ファイル数は月ごとのままで 321〜323／表（件数では分けない。分けても切られ方は変わらない）。ファイル名にヒント記法（`.[...]`）は0件。
- 総トークンの内訳: before 21.4M → **27.7M**（内容を削らないようにした分。人名・コードの列・時系列の全文）→ **39.3M**（分けた部分に管理No・設備・日付を書き直す分。1つ増えるごとに約110トークン）。どちらも利用者の指示（入力を勝手に加工しない／どのサーバー設定でも安全）を優先した結果。

## 7. テスト方針
- 既存テストは新ルートに合わせて書き換える（test_extraction は維持）。
- WPごとの単体テスト（純粋関数中心）。samples/ の実ファイルを使うテストは `@pytest.mark.samples`（samples が無ければ skip）。
- ゴールデン：同じ入力→同じ md バイト列（ハッシュ比較）。
- AI：`tests/fake_servers.py` を拡張し、リクエスト内容で応答を変える（壊れたJSON、原文にない型番、429、遅延）。
- 画面：Flask test client で主要フロー（帳票: upload→type→read→review→confirm→done→download、一覧表: CSV upload→source→layout→columns→preview→confirm→done→zip）。
- ブラウザ確認（統合時）：3つの画面（帳票取り込みは1件・まとめ取り込みの両方）、エラー画面（400/403/404/413/500）。
- 帳票サンプルでの精度測定：`python -m scripts.samples.evaluate_forms`（見本ファイル数は `--samples`、見本の選び方をずらすのは `--offset`）。
  見本を変えても結果が同じ傾向か（特定の見本への当て込みでないか）を `--samples 2` や `--offset 3` で確かめる。

## 8. 実装しないもの（第2段階以降）
大きいExcel用の1パス読み込み、数式XMLの解析（小計はキーワード＋太字で判定）、同じ構造の複数シート一括、複数表の自動検出（範囲の手動指定で対応）、帳票内の明細表の高度な形（2段の列見出し、承認欄のような行見出し付きの格子、日付が変わった行だけ書く時系列の日付の補完）、full プロファイル、巨大セルの分割AI処理、抜き取り確認の統計、修正内容からのルール提案、設備台帳、帳票と一覧表の紐付け、Batch API。

### 8.0 残っている不一致・開いている点
測定値（`python -m scripts.samples.evaluate_forms`、見本各3件・対象各30件、2026-09-20）：
F1 98.5% ／ F2 97.7% ／ F3 95.0% ／ F4 100.0% ／ F5 95.8%、
全体で 1つの値の項目 2,825/2,910 = 97.1%（正解が空なのに読んだ 29）、明細表の行 1,972/1,995 = 98.8%。
明細表の行は、評価スクリプトが正解の `*_norm`（ISO の日時）も照合に使い、真偽値のチェック欄（`checked: true/false`）を
☑ / □ とみなすようにしたあとの値（以前の数え方では 88.8%。F2 の時系列 311/314・横展開先 61/61 が照合できていなかった）。
残っているのは次のもの。

**読み取りの不一致**
- 発行側と回答側で同じ意味の欄が並ぶ帳票は、項目の「探す区画」（▼ 回答欄 など区切りの見出しの名前）で読む側を決める。
  見本の中で見出しが2つ以上の区画にあり、どの見本にも共通する名前付きの区画が1つだけの項目は、帳票の種類を作るときに
  自動で決まる（F5 の処置・原因＝回答欄。未回答なら空。`action` 18/24）。区画は `pattern_fields.extraction_rule` の
  `section` に保存する。押印欄の「確認」のように区画の外にも共通してある項目は決めない（F5 の `checker` 16/25。
  見本からはどちら側か決められないので、必要なら画面の「探す区画」で決める）。
- 見本に無い版だけにあるラベル（辞書に無いもの）は読めない（確認画面でラベルを足す運用）。
- 見出しと列見出しの間に「展開区分｜☑同型機」の1行がある明細表は、上の「■」「1.」の見出しをアンカーにして読む
  （F2 横展開先 61/61 行）。
- F5 の明細表の数量列（品番 10/12、投入数・不良数・保留数 いずれも 12/14）は、評価スクリプトが正解の合計値・
  チェック値を明細表の列をつないだ値と比べているための差で、読み取りの誤りではない。

**明細表の形**
- 同じ形の列見出しが縦に積み重なる明細表（F3 の特性要因図＝人｜機械・材料｜方法・測定｜環境）は、2026-09-19 の判断で**全部の組を1つの表にまとめて読む**ようにした（6.1「積み重なった列見出し」。F3 の6カテゴリとも、特性要因図のある見本 11 件すべてで一致。この表を持たない版では空欄のまま）。
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
- 帳票登録で置いた Excel は**そもそも保存しない**（2026-09-21。読み取った結果だけを `core/workbook_cache.py` が
  メモリに短い間だけ覚える）。［使用開始］は何も消さない。あとから項目を直すときは、同じ帳票の Excel をもう一度置く。
- `instance/app.db` のファイル自体は残る（中身は `secure_delete` + `VACUUM` で消える。ブラウザごとの AI接続 `ai_connections` は
  設定なので、APIキーを平文のまま残す）。`.flask_secret`・`env`・前の版が書いた `data/model_settings.yaml`（あれば。もう書かない）も残る。
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
- 一覧表の取り込み設定: 保存しない（取り込みごとに列の対応づけを決め、`table_imports.spec_json` に入れる。利用者の指示 2026-09-20）。
  画面で決めるのは「表の名前」と、列ごとの「使う」「役割」だけ。md の名前は常に「表の名前」を使う
  （`markdown.file_prefix` は '' のまま。`TableSpec.file_prefix` が名前に落とす）。記録ファイルのまとめ方は常に月ごと。
  設定の検査（`validate_spec`）は残す：伏せ字の規則・人名・分割の正規表現（20個まで・各200文字まで・入れ子の繰り返しは不可）などの
  形を確かめて日本語で断る。設定の JSON 書き出し・読み込みは無い。
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
  学習・保存を通しても区画が変わらない）。読み取り結果で入力欄に Enter を押しても、［確定］は押されない
  （当時あった［AIで空欄を探す］も押されなかった。この帳票の AI 補助は 2026-09-20 に削除した）。
  種類の項目が設備だけで識別子にならないときの「- 出典:」の行は、H1 と同じ項目（番号の項目、無ければ日付）を添える。
- 帳票取り込みと帳票登録で置く Excel は、結合セルの面積の合計が 20万セル（`core/files.FORM_MAX_MERGED_CELLS`）を超えるブックを断る
  （行全体の結合 13個以上・列全体の結合など。一覧表は読み取り専用で開き結合を展開しないので、従来の 200万セルのまま）。
- 一覧表: 大文字・小文字だけ違う設備のファイル名には `_2` を付ける（Windows のフォルダで上書きし合わない）。256列を超えるシートと、
  「行×列」がセル数の上限（`tables/excel_source.DEFAULT_MAX_CELLS` 50万。中身のほとんどが空）を超えるシートは、行を最終列まで埋めない
  （CSV と同じく行ごとに長さが違う。読む側は範囲外を空として扱う）。遠くの行のセル1つ（`IV1048576` のメモなど）で
  全行に何十万個のセルを作って読み込みが何十秒も止まるのを防ぐため。
  表から20列以上離れた、下にデータの無い見出しセル1つ（XFD1 のメモなど）は表に入れず警告にする。
  時間の単位の数値列では「1:30」「1時間30分」の文字も分として読む。確定の処理の途中で消えた取り込みのフォルダを作り直さない。
- 一覧表の md: 時系列の重複の削除で「(1)」「①」「⑴」「・」の行頭も外して比べ、1件だけのログも対象にする。見出しの1行目がタグと日付だけ
  なら次の行を使う。トークン数の見積もり（`core/mdtext.estimate_tokens`）は記号と数字を多めに数える（実測より少なく見積もらない）。
- 利用者が設定で書いた正規表現（分割の目印・日付ではない書き方）は、選択肢（`|`）を含む繰り返しのグループも断り、照合には行の先頭
  200文字（`logproc/dates.USER_PATTERN_WINDOW`）だけを渡す。取り込み設定の JSON の確かめで思わぬ例外が出たときも「読み込めません」で断る。
- 設定ファイル（`services/settings_store`）は UTF-16・CP932 でも読み、読めなければ空とみなす。読み書きは同じロックの中で行い、
  書けないときは画面に日本語で知らせる。すべての応答に `X-Frame-Options: DENY` と
  `frame-ancestors 'none'` を付ける。
- AI: AI の呼び出しとモデル一覧は明示のタイムアウト（ローカル300秒・クラウド120秒・一覧30秒）で、SDK の自動再試行はしない。
  構造化出力の確かめは推論モデルでも切れない長さで行う。429 で失敗した行が3行続いたら、エラーにせず未処理に戻して一時停止する
  （「混み合っています…再開を押してください」）。一時停止を押してから止まるまでの間も［再開］を出す。
  2つ目の起動が、1つ目のアプリの待機中のジョブを「中断」にしない（待機中のジョブの生存時刻も更新する）。

**6巡目の不具合探しで変えた動き（2026-09-19）**
- 帳票: 見出しの区画は、右の上の行から始まる見出し（右上の【回答欄】など）の列で右への広がりを止め、その見出しは下の行で別の見出しが
  その列に来るまで自分の列を持つ（右上の押印欄の見出しも同じく下の列を持つ。F1〜F5 の見本には無い形）。入れ子の区画（「暫定対策」の中の
  「回答」など）は `excel/tables.sections_of` で内側から外側まで返し、「探す区画」がどれかに当たればその区画の中とみなす。
  明細表の項目も「探す区画」の中を先に探す（無ければシート全体）。確認画面の「変更した項目」は単位だけの変更も出す。
- 一覧表: 手で決めた見出し行が先頭80行より下でも読む。行番号の入力（「1-3, 5」）は全角も含めてサーバーと画面で同じ規則で読む
  （範囲は10行まで）。役割「経過の記録」（`log`）は1列だけに付ける。日付の役割が変わらなければ期間の日付列を保つ。キーの重複を断る。
  確定済みの取り込みでは試し実行をしない。読み込み直しが始まって渡せなかったときは 404 ではなく画面に戻して知らせる。
  Excel は 16,384 列・1,048,576 行を超えるシートを断り、値のある列が 2,000 列を超えるシートも CSV と同じ文で断る。
  md のフォルダのパスが長すぎるときは「表の名前」を短くするよう知らせる。
- 一覧表の md: 月別・設備別の要約の設備名は、取り込み全体でいちばん多い名前にする（グループの先頭行ではなく）。平均は小数1桁に丸める。
  「10:00以降」「4/3から」のように範囲の語が続く日時は見出しから外さない。CSV の数式よけは全角の ＝＋－＠ で始まる値にも付ける。
- 分割（`logproc`）: 「1.5mm」「2.5A」のような寸法・電気量を日付にしない規則（`DEFAULT_NOT_DATE_PATTERNS`）は、単位の後ろに英数字・「-」・
  カタカナが続く「4.3 AGV」「4.3 ALM-2031」「4.3 Aライン」「4.3 Vベルト」には当てず、日付として読む（保存済みの取り込み設定の規則はそのまま）。
- AI: 429 のあとで行の時間切れになったときも 429 として数え、続けば一時停止にする。応答のキャッシュの読み書きがロック中なら少し待って
  再試行し、だめなら使わずに（書かずに）進む。ジョブの進捗・メッセージ・生存時刻の書き込みもロック中なら少し待ち、だめならその1回を見送る
  （`core/jobs.JobContext._soft_update`。状態の書き込みは見送らない）。
- アップロード: ハイパーリンク・コメントの範囲の合計が5万セルを超えるブック、帳票では 20万行を超えるブックを開く前に断る（5.1）。

**6巡目の見直しで変えた動き（2026-09-20）**
- 帳票: 値が「2024年7月28日 22:46」のように日付だけで書かれていれば、ラベルに日付の語が無くても（「発生」など）日付の項目にする
  （`pattern/builder._date_value`。値の全体が元号・年月日（＋曜日・時刻）で、`to_date` が警告なしで読めるときだけ。
  「2/12」「12:30」「2026-09-14 に復旧」は文字列のまま）。まとめ取り込みが1件だけになったら、まとまりとして扱わない
  （確認文から「残りの帳票はまとまりに残ります」が消える）。
- 一覧表: 記録番号・担当者の列の「〃」「同上」も直前のデータ行の値で補う（`ditto_filled` の警告にまとめる。
  重複キーの警告が「〃」ではなく本当の伝票番号を指すようになる）。取り込み設定の記録キーは「列:文字数」の書き方も受け取り、
  手で書いた `record.fallback_key` の列も確かめる。右側の別表・遠くの見出しの警告は、無い「列の範囲指定」ではなく
  「Excel で分ける／右隣に移す」を案内する。データの行が0行のときは、見出し行ではなく指定した「終わりの行」を理由として知らせる。
  中止は、押す直前に読み込みが終わっていれば状態を巻き戻さない。取り込みの［削除］の確認文と、処理中は出さない規則は
  `templates/components/_ui.html` の `table_import_delete_form` に1か所化した。
- AI: 対応していない引数の学習は「値を固定する」ではなく「名前を置き換える」（`max_tokens` → `max_completion_tokens`）。
  接続テストの 20 トークンがそのあとの全行に残らない。学習は（接続先URL, モデル）ごとで、AI接続を保存すると忘れる。
  見積もりは DB 接続を1本で通す（8,000行で約24秒短縮）。基準日の決め方は `tables/spec.base_date_from` の1か所にまとめ、
  分割プレビュー・試し実行・送る文面・できる md で同じ日付になる。AI接続が外れていても、動いている AI整形は［中止］できる。
- AI整形の試し実行の片付けは、まだある別の取り込みが払った応答を消さない（`core/purge.NO_LIVE_OWNER` を `aiproc/runner._ensure_trial_import`
  でも使う。消すと払った取り込みの再開・再実行で再課金になる）。
- 画面: 列の対応づけの保存が複数の理由で断られたら全部並べる（`static/app.js` の `ragFetch` が `errors` を例外に載せる）。
  AI整形の［見積もる］は答えが返るまで押せない。待ち画面のポーリングは、ジョブが消えた（404）ら止まって画面を読み直す。

**Markdown の断片**
- 明細表を読むようになったため、帳票の md が 1,200トークン（LightRAG の既定の固定窓）を超えて2つ以上の断片に分かれる様式がある（サンプルでは F2・F3・F4）。長文項目の見出しには識別子が入る。明細表の行だけで埋まった断片には識別番号も設備名も出なかったので、識別子を入れる帳票では明細表1節が 800トークンを超えたところで `## {見出し}（続き）（{識別子}）` に分けるようにした（6.1）。明細表の各行への設備番号の付与・帳票へのヒント付与は、出力仕様の変更になるため入れていない（`docs/research.md` 2章（LightRAG オフライン評価）8.5/8.7）。

### 8.1 後回し（2026-09-16 の範囲の見直しで外したもの）
統合時に、次の機能のコード・ルート・画面・テストを削除した。必要になったら 2.4・2.6・3.2・5.3・6.2〜6.4 の記述をもとに作り直す。
- **一覧表の更新管理**: 期間の置き換え（`tables/state.py` の current/previous、`replace_period`）、投入済みとの差分（`table_outputs.delivered_hash`、差分zip、「削除すべき旧ファイル」）、[投入済みにする]（`table_downloads`）、`/tables/templates/<tid>/outputs` 画面、直前の確定の取り消し。現状は「取り込みごとに全 md を作り、全ファイルを zip で渡す」（`tables/pipeline.py` の `run_read` / `run_render` / `build_download`）。md は `TABLES_DIR/imports/<id>/md/` に保存。
- **クロス集計**（縦持ち変換、`TableSpec.kind="crosstab"`、年月見出しの解釈、設備×月の値の集計ファイル）。表の形の判定（`tables/detect.py`）はクロス集計を見分けるが、範囲確認の画面で「対応していません」と止める。
- **名寄せ辞書**（`/settings/aliases`、`alias_entries` の参照、`ColumnSpec.alias_dictionary`）。値は NFKC と空白の畳み込みだけで揃える。
- **取り込み設定を保存する仕組み**（設定の一覧・名前での選び直し・必須列の一致数での候補提示・版の履歴・削除・JSON の書き出し読み込み、
  `table_templates` 系の表と `tables/mapping.match_templates`）。取り込みごとに列の対応づけを決める（利用者の指示 2026-09-20）。
- **列の対応づけでキー・表示名・型・単位・説明・md での扱い・空欄＝上と同じを手で直すこと**（見出しと値から自動で決める。同上）。
- **構築後のレビュー・評価フェーズ、LightRAG へのオフライン評価**（利用者が後で行う）。
- **表の読み取り方を取り込み設定で変える項目**（`header.search_rows`、`data_end`（`blank_rows` / `stop_first_col` / `stop_prefix`）、
  `exclude.aggregate_keywords`）。判定は `tables/detect.py` が持つ（合計・小計の語は「設計」「稼働時間累計」などと混ざらないよう
  細かく調整してあり、設定で差し替えると誤判定に戻る）。古い JSON にこれらの項目があっても読み捨てる（`tables/spec.py` の `RETIRED_KEYS`）。
  見出しの行・データの終わりの行は、取り込みごとに［範囲の確認］の画面で指定する。

DB のスキーマ（3.2）にあった `table_outputs`・`table_downloads`・`alias_entries`・`table_template_samples` は、取り込みを指す列
（`document_id` / `import_id`）を持たず purge の探索から漏れるため、マイグレーション `_m5` で削除した（どこからも書いていなかった）。
これらの機能を作り直すときは、取り込み単位の行を持つ表に `import_id` 列を付けてから作る。期間の列（`period_json`）も落とした（常に全期間）。
