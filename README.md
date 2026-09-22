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

## 設定（env ファイル）

社内LANの他のPCから開くときや、AI（OpenAI互換API）を使うときは、リポジトリ直下に `env`（ドット無し）という名前でファイルを作り、下をコピーして値を入れてください。AI の接続先・APIキーは画面右上の［AI接続］でブラウザごとに設定することもできます（そちらが優先されます）。**AI なしでも最後まで使えます。**

```sh
# このファイルを "env"（ドット無し）という名前でコピーして値を入れてください。
# 画面右上の「AI接続」でブラウザごとに保存した値がある場合は、そちらが優先されます（空欄の項目だけ、ここの値を使います）。

# OpenAI互換APIの接続先（Ollama等なら http://127.0.0.1:11434/v1）
export OPENAI_BASE_URL="https://api.openai.com/v1"
export OPENAI_API_KEY="<API_KEY>"

# 画面で候補を登録していないときの選択肢（; または , 区切り）と既定モデル
export OPENAI_MODELS="gpt-5.6-sol;gpt-5.6-luna;gpt-4.1;gpt-4o-mini"
export OPENAI_MODEL="gpt-5.6-sol"

# 社内LANのサーバーで動かすとき（JupyterLab のターミナルから起動する）
# export HOST="0.0.0.0"          # 既定 127.0.0.1（このサーバーの中からしか開けない）
# export PORT="5000"
# export ALLOWED_HOSTS="rag-server;rag.example.local"   # 社内DNSの別名で開くとき（; か , 区切り。* で全許可）
# export JOB_WORKERS="3"         # 読み込み・AI整形を同時に動かす本数（1〜8）。数人で使うなら 3 前後

# 任意
# export OPENAI_TEMPERATURE="0"
# export OPENAI_TOP_P=""
# export OPENAI_MAX_TOKENS=""
# export LLM_RATE_LIMIT_RETRIES="3"
# export LLM_RATE_LIMIT_MAX_WAIT="20"
```

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
```
