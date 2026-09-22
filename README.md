# RAG用Markdown作成

Excel/CSV の帳票・一覧表を読み取り、LightRAG に手作業で投入しやすい Markdown(.md) を作ってダウンロードする Flask アプリです。社内LANのサーバーで動かし、数人が自分のPCのブラウザから同時に使います。ログインはありません。LightRAG への送信はしません（ダウンロードまで）。

取り込んだデータはサーバーに残しません。Markdown をダウンロードした時点で、元のファイル・読み取った内容・作ったファイル・DBの行を消します。残るのは帳票の種類と、ブラウザごとの AI接続の設定だけです。

## 起動

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

http://127.0.0.1:5000 を開きます（既定はサーバーの中からだけ開けます）。初回は `instance/app.db` を作り、古い DB は起動時に自動で移行します。

社内LANの他のPCから開くときは、`env.example` を `env`（ドット無し）にコピーして `HOST="0.0.0.0"` を有効にしてください。AI（OpenAI互換API）の接続先・APIキーも同じファイルか、画面右上の［AI接続］で設定します。AI なしでも最後まで使えます。

## 画面

| 画面 | アドレス | できること |
|---|---|---|
| 帳票取り込み | `/forms`（`/` はここへ） | 帳票の Excel を読み取って Markdown にする |
| 表の取り込み | `/tables` | Excel/CSV の一覧表を読み取って Markdown（zip）にする |
| 帳票登録 | `/form-types` | 帳票の Excel を置いて読み取る欄を決め、帳票の種類を作る |
| 解説 | `/guide` | Markdown がどう作られるか・帳票登録でのクリックのしかた |

## 構成

```
run.py                 起動（waitress で create_app() を動かす）
app/                   アプリ本体（views・forms・tables・aiproc・llm・core・database・logproc）
app/templates/         画面（HTML）
app/static/            app.js・style.css
requirements.txt       必要なパッケージ
env.example            環境変数の見本
```
