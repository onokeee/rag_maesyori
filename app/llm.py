"""OpenAI 互換 API の接続と、ヘッダーの「AI接続」からブラウザごとに保存する接続の設定。

設定の出どころは3つで、この順に見る（項目ごと）:
  1. そのブラウザが「AI接続」で保存したもの（database.ai_connections。持ち主は views.current_session_id）
  2. 前の版の画面が書いた data/model_settings.yaml（サーバー共通。もう書かない。あれば読むだけ）
  3. env ファイル（OPENAI_BASE_URL / OPENAI_API_KEY / OPENAI_MODEL / OPENAI_MODELS。手で組んだサーバー向け）
画面の入力欄には 1 の値しか入れない（2・3 は「サーバー共通の設定」として、あることだけ知らせる）。
"""
from __future__ import annotations

import hashlib
import ipaddress
import json
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit

import yaml
from flask import current_app, g, has_request_context, session
from openai import OpenAI

from app.core import REMOVE_RETRIES, REMOVE_RETRY_WAIT



# ====================================================================================================
# 元 services/settings_store.py
# data/ 以下の設定ファイルの読み込み（スレッドセーフ）。
#
# 前の版は AI接続をここに書いていた（data/model_settings.yaml。全員で1つ）。今は書かない。
# 残っているファイルは「サーバー共通の設定」として読むだけ（llm._read_admin）。
# ウイルス対策・同期ソフトが少しの間ファイルを掴んでいることもあるので、core.files.remove_upload と同じく
# 少し待って何度か試す。
# ====================================================================================================

_settings_lock = threading.RLock()

# 手で編集して UTF-8 以外で保存された設定ファイルも読めるように試す文字コード（BOM 付き UTF-8 もここで読める）
_FALLBACK_ENCODINGS = ("utf-8-sig", "cp932")


def data_path(name: str) -> Path:
    return Path(current_app.config["DATA_DIR"]) / name


def _retry(action):
    """PermissionError（ほかのプログラムが掴んでいる）なら、少し待って何度か試す。最後の失敗はそのまま出す。"""
    for attempt in range(REMOVE_RETRIES):
        try:
            return action()
        except PermissionError:
            if attempt == REMOVE_RETRIES - 1:
                raise
            time.sleep(REMOVE_RETRY_WAIT * (attempt + 1))


def _decode(raw: bytes) -> str:
    """設定ファイルの中身を文字にする（UTF-16 は BOM で見分け、それ以外は UTF-8 → CP932 の順に試す）。"""
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in _FALLBACK_ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("utf-8", raw, 0, 1, "UTF-8・CP932 のどちらでも読めません")


def read_yaml(name: str) -> dict:
    path = data_path(name)
    try:
        with _settings_lock:
            if not path.exists():
                return {}
            raw = _retry(path.read_bytes)
        data = yaml.safe_load(_decode(raw))
    except (OSError, ValueError, yaml.YAMLError) as exc:   # UnicodeDecodeError は ValueError の仲間
        print(f"[settings] {path} を読めませんでした（{exc}）")
        return {}
    return data if isinstance(data, dict) else {}


# ====================================================================================================
# 元 services/llm.py
# OpenAI互換APIへの接続とモデル選択。
#
# aiagent_minimal_rag_tougou の llm.py / models.py と同じ仕様:
#   - 接続先とAPIキーは env（OPENAI_BASE_URL / OPENAI_API_KEY）が既定。
#     ヘッダーの「AI接続」でそのブラウザが保存した値（database.ai_connections）があればそちらを優先する。
#     前の版の画面が書いた data/model_settings.yaml（サーバー共通）は、その間（ブラウザ → yaml → env）。
#   - URL はフルパス2本（…/chat/completions と …/models）で持つ。
#   - 選べるモデルは「ブラウザが取得した候補 → yaml の候補 → env の OPENAI_MODELS」の順。使うモデルは必ず候補に入る。
#   - モデルが受け付けない引数（temperature / max_tokens / reasoning_effort）はエラー文を見て直して投げ直す。
#   - 429 はサーバの指示した時間だけ待って投げ直す。
# ====================================================================================================

SETTINGS_FILE = "model_settings.yaml"
ADMIN_KEYS =("models", "default", "api_key", "chat_url", "models_url")
CATALOG_TTL = 300
_MAX_FIX = 4
# ブラウザの作業場所の id のクッキーの鍵（views.SESSION_ID_KEY と同じ。views は画面の層なので、ここからは読まない）
SESSION_ID_KEY = "sid"

_clients: dict[tuple[str, str, float], OpenAI] = {}
_catalog_cache: dict[tuple[str, str], tuple[float, list[str]]] = {}
# モデルが受け付けない引数の覚え書き。接続先×モデルごと（同じモデル名でも別のサーバーなら別の癖）
_QUIRKS: dict[tuple[str, str], dict] = {}
# ジョブの並列呼び出しから _clients / _QUIRKS / 方式判定を同時に更新するためのロック
_lock = threading.RLock()


class LLMNotConfigured(Exception):
    pass


def _cfg(name: str):
    return current_app.config[name]


# ---- 設定の解決 ----------------------------------------------------------------
# 項目ごとに「このブラウザが保存した値 → 前の版の yaml（サーバー共通） → env」の順に見る。
# ブラウザの値は要求（request）の中でしか分からない。ジョブのスレッドには無いので、AI整形は開始時に
# job_client_settings() で設定を固めて渡す（aiproc.start_ai_job）。

def _read_admin() -> dict:
    """前の版が書いた data/model_settings.yaml（サーバー共通・読むだけ）。無ければ空。要求の中では1回だけ読む。"""
    if has_request_context():
        cached = getattr(g, "_ai_admin", None)
        if cached is not None:
            return cached
    data = read_yaml(SETTINGS_FILE)
    admin = {k: v for k, v in data.items() if k in ADMIN_KEYS}
    if has_request_context():
        g._ai_admin = admin
    return admin


def _session_id() -> str:
    """いまの要求のブラウザの作業場所の id。要求の外（ジョブのスレッド）や、まだ配っていないときは空。"""
    if not has_request_context():
        return ""
    sid = session.get(SESSION_ID_KEY)
    return sid if isinstance(sid, str) and len(sid) == 32 else ""


def _browser() -> dict:
    """このブラウザが「AI接続」で保存したもの（要求ごとに1回だけ DB を読む）。無ければ空。"""
    sid = _session_id()
    if not sid:
        return {}
    cached = getattr(g, "_ai_connection", None)
    if cached is None or cached[0] != sid:
        from app import database

        cached = (sid, database.get_ai_connection(sid) or {})
        g._ai_connection = cached
    return cached[1]


def forget_browser() -> None:
    """保存・確認のあと、同じ要求の中で読み直せるようにする。"""
    if has_request_context():
        g.pop("_ai_connection", None)


def _env_chat_url() -> str:
    return f"{_cfg('OPENAI_BASE_URL')}/chat/completions" if _cfg("OPENAI_BASE_URL") else ""


def _env_models_url() -> str:
    return f"{_cfg('OPENAI_BASE_URL')}/models" if _cfg("OPENAI_BASE_URL") else ""


def _pick(browser_key: str, yaml_key: str, env_value: str) -> tuple[str, str]:
    """設定1項目の値と出どころ（"browser" / "server" / "env" / ""）。"""
    value = str(_browser().get(browser_key) or "").strip()
    if value:
        return value, "browser"
    value = str(_read_admin().get(yaml_key) or "").strip()
    if value:
        return value, "server"
    value = str(env_value or "").strip()
    return (value, "env") if value else ("", "")


def llm_api_key() -> str:
    return _pick("api_key", "api_key", _cfg("OPENAI_API_KEY"))[0]


def llm_api_key_source() -> str:
    """"browser"（このブラウザで保存）/ "server"（前の版の yaml）/ "env" / ""（未設定）"""
    return _pick("api_key", "api_key", _cfg("OPENAI_API_KEY"))[1]


def llm_chat_url() -> str:
    return _pick("chat_url", "chat_url", _env_chat_url())[0]


def llm_models_url() -> str:
    return _pick("models_url", "models_url", _env_models_url())[0]


def is_configured() -> bool:
    return bool(llm_chat_url() and llm_api_key())


def default_model() -> str:
    return _pick("model", "default", _cfg("OPENAI_MODEL"))[0]


def available() -> list[str]:
    browser = [str(m).strip() for m in (_browser().get("models") or []) if str(m).strip()]
    admin = [str(m).strip() for m in (_read_admin().get("models") or []) if str(m).strip()]
    names = browser or admin or list(_cfg("OPENAI_MODELS"))
    d = default_model()
    if d and d not in names:
        names.insert(0, d)
    return names


def current_model() -> str:
    """いま使うモデル（画面のモデル選択は無くなったので既定モデル）。"""
    return default_model()


def server_fallback() -> dict:
    """ブラウザが空欄にした項目を埋める「サーバー共通の設定」があるか（前の版の yaml と env）。画面に知らせる。"""
    admin = _read_admin()
    yaml_key = bool(str(admin.get("api_key") or "").strip())
    yaml_url = bool(str(admin.get("chat_url") or "").strip())
    env_key = bool(_cfg("OPENAI_API_KEY"))
    return {
        "yaml": yaml_key or yaml_url or bool(admin.get("models")) or bool(admin.get("default")),
        "yaml_file": str(data_path(SETTINGS_FILE)),
        "env_key": env_key,
        "env_url": _env_chat_url(),
        "any_key": yaml_key or env_key,
    }


# ---- クライアント ---------------------------------------------------------------

def _derived_base(full_url: str, suffix: str) -> str:
    u = str(full_url or "").strip().rstrip("/")
    return u[: -len(suffix)] if u.endswith(suffix) else ""


MODELS_TIMEOUT = 30.0


def _client_for(base: str, timeout: float) -> OpenAI:
    """明示タイムアウト・SDK の再試行なしのクライアント。
    （SDK の既定は読み取り600秒×再試行2回で、応答しない接続先だと画面が最大30分ほど待たされる。429 は _create が待って投げ直す）"""
    key = (base, llm_api_key(), float(timeout))
    with _lock:
        if key not in _clients:
            _clients[key] = OpenAI(base_url=base or None, api_key=key[1] or "not-set", timeout=key[2], max_retries=0)
        return _clients[key]


def models_client() -> OpenAI:
    return _client_for(_derived_base(llm_models_url(), "/models"), MODELS_TIMEOUT)


def reset_llm_client() -> None:
    with _lock:
        _clients.clear()
        _job_clients.clear()
        _catalog_cache.clear()
        _QUIRKS.clear()   # 接続先・キーを変えたら引数の覚え書きも捨てる（再起動なしで効かせる）


def model_catalog(refresh: bool = False) -> list[str]:
    """APIの models.list() とenvの候補を合わせた一覧（300秒キャッシュ）。"""
    return sorted(set(_cfg("OPENAI_MODELS")) | set(fetch_api_models(refresh)))


def fetch_api_models(refresh: bool = False) -> list[str]:
    cache_key = (llm_models_url(), llm_api_key())
    cached = _catalog_cache.get(cache_key)
    if cached and not refresh and time.time() - cached[0] < CATALOG_TTL:
        return cached[1]
    if not llm_api_key():
        raise LLMNotConfigured("APIキーが未設定です。")
    got = sorted(m.id for m in models_client().models.list().data)
    _catalog_cache[cache_key] = (time.time(), got)
    return got


# ---- ヘッダーの「AI接続」（ブラウザごとの保存・状態・確認） ------------------------------------
# 利用者の指示（2026-09-21）:「AI接続はヘッダー上で、接続中 か 未接続 一目で分かるように」。
# ヘッダーに出す状態は4つ:
#   ok        ● 接続中        最後の確認がつながった
#   ng        ● つながりません 設定はあるが最後の確認がつながらなかった（パネルを開くと理由が出る）
#   off       ● 未接続        まだ何も無い（キーか接続先が無い）
#   unchecked ● 未確認        サーバー共通の設定（yaml / env）だけがあり、このブラウザではまだ確かめていない
# 確認しに行くのは「保存したとき」「パネルを開いたとき」「AI整形を始めるとき」だけ（check_connection）。
# 画面を開くだけでは行かない（お金がかかり、画面が待たされる）。結果は database.ai_connections に覚えて表示する。

STATE_LABELS = {"ok": "接続中", "ng": "つながりません", "off": "未接続", "unchecked": "未確認"}
_CHECK_PROMPT = "接続テストです。「OK」とだけ返してください。"


def _check_time_text(iso: str | None) -> str:
    """「9/21 10:12」の形（ヘッダーの「最終確認」）。"""
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(str(iso))
    except ValueError:
        return str(iso)
    return f"{dt.month}/{dt.day} {dt.hour:02d}:{dt.minute:02d}"


def connection_status() -> dict:
    """ヘッダーとパネルに出す、このブラウザの AI接続の状態。APIキーの値は返さない（有無と出どころだけ）。"""
    row = _browser()
    ready = is_configured()
    check = row.get("last_check_ok")
    if check == 1 and ready:
        state = "ok"
    elif check == 0 and ready:
        state = "ng"
    elif ready:
        state = "unchecked"
    else:
        state = "off"
    chat_url = llm_chat_url()
    checked_text = _check_time_text(row.get("last_check_at"))
    return {
        "state": state,
        "state_label": STATE_LABELS[state],
        "ready": ready,
        "checked_at": row.get("last_check_at") or "",
        "checked_text": checked_text,
        # ヘッダーの2行目（「最終確認 9/21 10:12」）。未設定なら開き方、未確認ならそのことを書く
        "sub_text": (f"最終確認 {checked_text}" if checked_text
                     else ("クリックして設定" if state == "off" else "まだ確かめていません")),
        # パネルの中の同じ行（開き方はもう要らない）
        "panel_sub_text": (f"最終確認 {checked_text}" if checked_text
                           else ("" if state == "off" else "まだ確かめていません")),
        "check_detail": row.get("last_check_detail") or "",
        # 入力欄に入れる値は、このブラウザが保存したものだけ（サーバー共通の値で埋めない）
        "chat_url": str(row.get("chat_url") or ""),
        "models_url": str(row.get("models_url") or ""),
        "model": str(row.get("model") or ""),
        "models": list(row.get("models") or []),
        "api_key_saved": bool(str(row.get("api_key") or "").strip()),
        "api_key_source": llm_api_key_source(),
        # いま実際に使われる値（出どころ込み。パネルの「いまの設定」に出す）
        "effective": {
            "chat_url": chat_url, "chat_url_source": _pick("chat_url", "chat_url", _env_chat_url())[1],
            "models_url": llm_models_url(), "model": default_model(),
            "model_source": _pick("model", "default", _cfg("OPENAI_MODEL"))[1],
            "external": bool(chat_url) and not is_local_endpoint(chat_url),
        },
        "fallback": server_fallback(),
        "placeholders": {"chat_url": "https://api.openai.com/v1/chat/completions",
                         "models_url": "https://api.openai.com/v1/models", "model": _cfg("OPENAI_MODEL") or "gpt-5.6-sol"},
    }


def _clean_url(value, suffix: str, label: str) -> str:
    u = str(value or "").strip().rstrip("/")
    if not u:
        return ""
    if any(c.isspace() for c in u):
        raise ValueError(f"{label}のURLに空白が入っています。")
    if not u.startswith(("http://", "https://")):
        raise ValueError(f"{label}のURLは http:// か https:// で始めてください。")
    if not u.endswith(suffix):
        raise ValueError(f"{label}のURLは {suffix} で終わるフルパスで入力してください"
                         f"（例: https://api.openai.com/v1{suffix}）。")
    if len(u) > 500:
        raise ValueError(f"{label}のURLが長すぎます。")
    return u


def save_browser(session_id: str, data: dict) -> dict:
    """ヘッダーの「AI接続」からの保存（そのブラウザの分だけ）。値の問題は ValueError。

    空欄は「サーバー共通の値（yaml / env）に任せる」の意味で、そのまま空で保存する。
    APIキーは値が来たときだけ置き換える（入力欄には出さないので、空のまま保存しても消えない）。
    """
    if not session_id:
        raise ValueError("ブラウザの作業場所が分かりません。画面を開き直してください。")
    chat_url = _clean_url(data.get("chat_url"), "/chat/completions", "チャット")
    models_url = _clean_url(data.get("models_url"), "/models", "モデル一覧")
    model = str(data.get("model") or "").strip()
    if len(model) > 120:
        raise ValueError(f"モデル名が長すぎます: {model[:40]}…")
    if any(c.isspace() for c in model):
        raise ValueError("モデル名に空白が入っています。")
    models = None
    if isinstance(data.get("models"), list):
        models = list(dict.fromkeys(str(m).strip() for m in data["models"] if str(m).strip()))[:200]
        for m in models:
            if len(m) > 120:
                raise ValueError(f"モデル名が長すぎます: {m[:40]}…")

    key_in = data.get("api_key")
    key_new = str(key_in).strip() if isinstance(key_in, str) else ""
    if key_new:
        if any(c.isspace() for c in key_new):
            raise ValueError("APIキーに空白や改行が入っています。コピーし直してください。")
        if not (8 <= len(key_new) <= 500):
            raise ValueError("APIキーの長さが不自然です。値を確かめてください。")

    from app import database

    database.save_ai_connection(session_id, api_key=key_new or None, chat_url=chat_url, models_url=models_url,
                                model=model, models=models)
    forget_browser()
    reset_llm_client()   # 次のAI呼び出しから新しい接続先・キーを使う（再起動不要）
    print("[models] AI接続を保存しました（ブラウザごと）" + (" / APIキーを更新" if key_new else ""))
    return connection_status()


def clear_browser_key(session_id: str) -> dict:
    """［キーを消す］（共有PC）。そのブラウザの APIキーと確認の結果だけ消す。"""
    from app import database

    if session_id:
        database.clear_ai_connection_key(session_id)
        forget_browser()
        reset_llm_client()
    return connection_status()


def check_connection() -> tuple[bool, list[dict]]:
    """接続の確認: モデル一覧の取得と、短いチャットを1回。戻り値 (つながったか, 手順ごとの結果)。

    チャットが通れば AI整形は使える（モデル一覧の API が無い互換サーバーもある）ので、
    つながったかどうかはチャットで決める。
    """
    if not is_configured():
        return False, [{"name": "設定", "ok": False, "detail": "APIキーまたは接続先が未設定です。"}]
    steps = []
    started = time.monotonic()
    try:
        names = fetch_api_models(refresh=True)
        steps.append({"name": "モデル一覧の取得", "ok": True,
                      "detail": f"{len(names)}件のモデルが見つかりました（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": "モデル一覧の取得", "ok": False, "detail": friendly_error(exc)})
    model = current_model()
    started = time.monotonic()
    try:
        # AI整形のジョブと同じ呼び出し口（再試行なし・明示のタイムアウト）で確かめる
        result = chat_raw(job_client_settings(), [{"role": "user", "content": _CHECK_PROMPT}], max_tokens=20)
        reply = (result.text or "").strip()
        steps.append({"name": f"チャット（{model}）", "ok": True,
                      "detail": f"応答あり: {reply[:40] or '（本文なし）'}（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": f"チャット（{model}）", "ok": False, "detail": friendly_error(exc)})
    return steps[-1]["ok"], steps


def record_check(session_id: str, ok: bool, steps: list[dict]) -> None:
    """確認の結果をそのブラウザの行に覚える（ヘッダーの表示のもと）。設定が無いときは覚えない。"""
    if not session_id or not is_configured():
        return
    from app import database

    failed = [s for s in steps if not s.get("ok")]
    if ok and not failed:
        detail = ""
    elif len({s["detail"] for s in failed}) == 1:
        detail = failed[0]["detail"]   # 2つの手順が同じ理由で失敗（接続先に届かない等）なら1回だけ書く
    else:
        detail = "／".join(f"{s['name']}: {s['detail']}" for s in failed)
    database.set_ai_connection_check(session_id, ok, detail)
    forget_browser()


def _ms(started: float) -> int:
    return int((time.monotonic() - started) * 1000)


# ---- 呼び出し ------------------------------------------------------------------

def _fix_for(message: str, kwargs: dict) -> tuple | None:
    """400エラーの文面から、引数の直し方 (set, drop, rename) を決める。

    rename は「引数の名前だけを付け替える」（値は毎回の呼び出しの値をそのまま使う）。
    max_tokens → max_completion_tokens をこの形で覚えないと、最初の1回の値が以後ずっと固定されてしまう。
    """
    low = message.lower()
    if "reasoning_effort" in low and "does not support" in low:
        return ({"reasoning_effort": "none"}, None, None)
    if "reasoning_effort" in low and "unrecognized" in low:
        return (None, "reasoning_effort", None)
    for name in ("temperature", "top_p"):
        if f"'{name}'" in low and ("does not support" in low or "only the default" in low or "unsupported" in low):
            return (None, name, None)
    if "max_tokens" in low and "max_completion_tokens" in low:
        if kwargs.get("max_tokens") is not None:
            return (None, None, {"max_tokens": "max_completion_tokens"})
    return None


def _quirk_key(endpoint: str, model: str) -> tuple[str, str]:
    return (str(endpoint or "").strip().rstrip("/"), str(model or ""))


def _learn(key: tuple[str, str], set_: dict | None = None, drop: str | None = None,
           rename: dict | None = None) -> None:
    with _lock:
        quirk = _QUIRKS.setdefault(key, {"set": {}, "drop": set(), "rename": {}})
        if set_:
            quirk["set"].update(set_)
        if drop:
            quirk["drop"].add(drop)
            quirk["set"].pop(drop, None)
            quirk["rename"].pop(drop, None)
        if rename:
            quirk["rename"].update(rename)


def _no_progress(attempt: dict, set_: dict | None, drop: str | None, rename: dict | None) -> bool:
    """この直し方ではもう変わらない（＝投げ直しても同じ）か。"""
    if set_ and all(attempt.get(k) == v for k, v in set_.items()):
        return True
    if drop and drop not in attempt:
        return True
    if rename and all(old not in attempt or new in attempt for old, new in rename.items()):
        return True
    return False


def _apply_quirks(kwargs: dict, endpoint: str = "") -> dict:
    attempt = dict(kwargs)
    with _lock:
        quirk = _QUIRKS.get(_quirk_key(endpoint, str(kwargs.get("model") or "")))
        if quirk:
            for old, new in quirk.get("rename", {}).items():
                if old in attempt:
                    attempt[new] = attempt.pop(old)
            for name in quirk["drop"]:
                attempt.pop(name, None)
            attempt.update(quirk["set"])
    return attempt


_RETRY_IN = re.compile(r"try again in\s+([\d.]+)\s*(ms|s|m)\b", re.IGNORECASE)


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, (LLMNotConfigured, ValueError, LLMCallError)):
        return str(exc)
    from openai import APIConnectionError, APITimeoutError

    if isinstance(exc, (APIConnectionError, APITimeoutError)):
        # SDK の英語の文（Connection error. など）ではなく、ジョブと同じ日本語の説明を出す
        return str(classify_error(exc, current_model()))
    status = getattr(exc, "status_code", None)
    if status in (401, 403):
        return "APIキーが拒否されました。「AI接続」でキーを確認してください。"
    if status == 404:
        return f"モデルまたは接続先が見つかりません（{current_model()}）。「AI接続」を確認してください。"
    text = str(exc)
    return f"AI呼び出しに失敗しました: {text[:160]}"


# ---- ジョブ用の呼び出し口（一覧表のAI整形） -------------------------------------------
# 設定はジョブ開始時に固定し、専用クライアント（明示タイムアウト・SDK再試行なし）で呼ぶ。
# ワーカースレッドから呼ぶため、ここの関数は current_app を使わない（job_client_settings だけは app_context 内で呼ぶ）。

CLOUD_TIMEOUT = 120.0
LOCAL_TIMEOUT = 300.0
_PROBE_SCHEMA = {
    "type": "object",
    "properties": {"ok": {"type": "boolean"}},
    "required": ["ok"],
    "additionalProperties": False,
}

_job_clients: dict[tuple[str, str, float], OpenAI] = {}
_MODES: dict[tuple[str, str], str] = {}
# 方式判定の出力上限。推論モデルは上限を考える途中で使い切ることがある（本文が空・finish_reason=length）ので
# 小さくしすぎない。打ち切られたら次の値でもう一度だけ試す
_PROBE_BUDGETS = (1000, 4000)
_THINK_RE = re.compile(r"<think>.*?</think>\s*", re.DOTALL | re.IGNORECASE)


@dataclass
class ChatResult:
    text: str                       # 応答本文（<think>…</think> を除いたもの）
    finish_reason: str | None
    tokens_in: int | None
    tokens_out: int | None
    latency_ms: int
    headers: dict = field(default_factory=dict)   # レート制限ヘッダーなど（小文字キー）
    params: dict = field(default_factory=dict)    # 実際に送った引数（messages を除く）


class LLMCallError(Exception):
    """ジョブ用呼び出しの失敗。kind でジョブの扱いを決める。

    fatal: ジョブを止める（401/403、モデルなし、残高不足、接続拒否）
    retry: 待って再試行（429、5xx、タイムアウト）。retry_after は秒（サーバの指示があれば）
    row:   その行だけエラー（400 など）
    """

    def __init__(self, kind: str, message: str, status: int | None = None, retry_after: float | None = None,
                 model: str = ""):
        super().__init__(message)
        self.kind = kind
        self.status = status
        self.retry_after = retry_after
        self.model = model


def is_local_endpoint(url: str) -> bool:
    """localhost・プライベートIP なら True（外部送信の確認とタイムアウトの既定に使う）。"""
    host = (urlsplit(str(url or "")).hostname or "").lower()
    if not host:
        return False
    if host in ("localhost", "host.docker.internal") or host.endswith((".local", ".localhost")):
        return True
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        return False
    return ip.is_loopback or ip.is_private or ip.is_link_local


def settings_fingerprint(settings: dict, structured_mode: str | None = None) -> str:
    """接続先・モデル・パラメータ（・方式）のハッシュ。APIキーは含めない。"""
    payload = {
        "chat_url": str(settings.get("chat_url") or ""),
        "model": str(settings.get("model") or ""),
        "params": settings.get("params") or {},
    }
    if structured_mode:
        payload["structured_mode"] = structured_mode
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


def job_client_settings(model: str | None = None, params: dict | None = None) -> dict:
    """ジョブ開始時に固定する接続設定（app_context 内で呼ぶ）。

    返す dict: chat_url, base_url, api_key, model, params, timeout, local, fingerprint。
    fingerprint にキーは含めない。ジョブの params に保存するときは public_settings() でキーを外す。
    """
    if not is_configured():
        raise LLMNotConfigured("AIの接続先が未設定です。画面右上の「AI接続」でAPIキーと接続先を設定してください。")
    chat_url = llm_chat_url()
    local = is_local_endpoint(chat_url)
    if params is None:
        params = {"temperature": _cfg("OPENAI_TEMPERATURE")}
        if _cfg("OPENAI_TOP_P") is not None:
            params["top_p"] = _cfg("OPENAI_TOP_P")
    settings = {
        "chat_url": chat_url,
        "base_url": _derived_base(chat_url, "/chat/completions"),
        "api_key": llm_api_key(),
        "model": str(model or current_model()),
        "params": dict(params),
        "timeout": LOCAL_TIMEOUT if local else CLOUD_TIMEOUT,
        "local": local,
    }
    settings["fingerprint"] = settings_fingerprint(settings)
    return settings


def public_settings(settings: dict) -> dict:
    """ジョブの params や画面に出してよい部分（APIキーを除く）。"""
    return {k: v for k, v in settings.items() if k != "api_key"}


def _job_client(settings: dict, timeout: float) -> OpenAI:
    key = (str(settings.get("base_url") or ""), str(settings.get("api_key") or ""), float(timeout))
    with _lock:
        if key not in _job_clients:
            _job_clients[key] = OpenAI(base_url=key[0] or None, api_key=key[1] or "not-set",
                                       timeout=key[2], max_retries=0)
        return _job_clients[key]


def _retry_after(exc: Exception) -> float | None:
    headers = getattr(getattr(exc, "response", None), "headers", None)
    for name in ("retry-after-ms", "retry-after"):
        try:
            raw = headers.get(name) if headers is not None else None
        except Exception:
            raw = None
        if raw:
            try:
                sec = float(raw)
            except ValueError:
                continue
            return sec / 1000 if name.endswith("-ms") else sec
    m = _RETRY_IN.search(str(exc))
    if m:
        value, unit = float(m.group(1)), m.group(2).lower()
        return value / 1000 if unit == "ms" else (value * 60 if unit == "m" else value)
    return None


_CONNECT_PHASE_ERRORS = {"ConnectError", "ConnectTimeout"}
_MID_REQUEST_ERRORS = {"RemoteProtocolError", "ReadError", "WriteError", "ConnectionResetError",
                       "ConnectionAbortedError", "BrokenPipeError", "IncompleteRead"}


def _dropped_mid_request(exc: BaseException) -> bool:
    """接続エラーの原因をたどり、接続後に切れたもの（再試行でよい）かを返す。
    httpx / httpx2 のどちらでも効くようにクラス名で見る。原因が分からなければ False（従来どおり止める）。"""
    seen, cur, mid = set(), exc.__cause__ or exc.__context__, False
    while cur is not None and id(cur) not in seen:
        seen.add(id(cur))
        names = {c.__name__ for c in type(cur).__mro__}
        if names & _CONNECT_PHASE_ERRORS:
            return False
        if names & _MID_REQUEST_ERRORS:
            mid = True
        cur = cur.__cause__ or cur.__context__
    return mid


def classify_error(exc: Exception, model: str = "") -> LLMCallError:
    """SDK の例外をジョブでの扱い（fatal / retry / row）に分類する。"""
    if isinstance(exc, LLMCallError):
        return exc
    from openai import APIConnectionError, APITimeoutError

    text = str(exc)
    low = text.lower()
    status = getattr(exc, "status_code", None)
    if isinstance(exc, APITimeoutError) or "timed out" in low:
        return LLMCallError("retry", "AIの応答が時間内に返りませんでした。", None, None, model)
    if isinstance(exc, APIConnectionError):
        # 要求を送った後に切れた（サーバー側の切断・読み書きの失敗）は一時的なので待って再試行する。
        # 接続そのものができない（接続拒否・名前解決・接続タイムアウト）ときだけジョブを止める
        if _dropped_mid_request(exc):
            return LLMCallError("retry", "AIとの接続が途中で切れました。", None, None, model)
        return LLMCallError("fatal", "AIの接続先に接続できませんでした。接続先URLとサーバーの起動を確認してください。",
                            None, None, model)
    if status in (401, 403):
        return LLMCallError("fatal", "APIキーが拒否されました。「AI接続」でキーを確認してください。", status, None, model)
    if status == 402 or "insufficient_quota" in low:
        return LLMCallError("fatal", "APIの残高・利用枠が不足しています。", status, None, model)
    if status == 404 or ("model" in low and ("not found" in low or "does not exist" in low)):
        return LLMCallError("fatal", f"モデルまたは接続先が見つかりません（{model}）。「AI接続」を確認してください。",
                            status, None, model)
    if status == 429 or "rate_limit" in low:
        return LLMCallError("retry", "混み合っています（レート制限）。", status, _retry_after(exc), model)
    if status is not None and status >= 500:
        return LLMCallError("retry", f"AIサーバーでエラーが発生しました（{status}）。", status, _retry_after(exc), model)
    return LLMCallError("row", f"AI呼び出しに失敗しました（{model}）: {text[:160]}", status, None, model)


def _job_fix_for(message: str, kwargs: dict) -> tuple | None:
    fix = _fix_for(message, kwargs)
    if fix is not None:
        return fix
    low = message.lower()
    if "seed" in low and ("unsupported" in low or "unrecognized" in low or "not support" in low):
        return (None, "seed", None)
    return None


def chat_raw(settings: dict, messages: list[dict], response_format: dict | None = None,
             max_tokens: int | None = None, timeout: float | None = None) -> ChatResult:
    """1回の chat 呼び出し（SDK の再試行なし）。受け付けない引数だけはエラー文から直して投げ直す。

    失敗は LLMCallError（kind=fatal/retry/row）。レート制限の待機と再試行は呼び出し側（ジョブ）で行う。
    timeout を渡すとこの1回だけその秒数で打ち切る（クライアントは設定のタイムアウトのものを使い回す。
    行ごとの残り時間で毎回違う値になっても、クライアントを作り増やさない）。
    """
    model = str(settings.get("model") or "")
    base_timeout = float(settings.get("timeout") or CLOUD_TIMEOUT)
    timeout = float(timeout or base_timeout)
    per_request = {"timeout": timeout} if timeout != base_timeout else {}
    kwargs = dict(settings.get("params") or {})
    kwargs.update(model=model, messages=messages)
    if response_format:
        kwargs["response_format"] = response_format
    if max_tokens:
        kwargs["max_tokens"] = int(max_tokens)
    cli = _job_client(settings, base_timeout)
    endpoint = str(settings.get("chat_url") or settings.get("base_url") or "")
    quirk_key = _quirk_key(endpoint, model)
    fixes = 0
    while True:
        attempt = _apply_quirks(kwargs, endpoint)
        started = time.monotonic()
        try:
            raw = cli.chat.completions.with_raw_response.create(**attempt, **per_request)
            resp = raw.parse()
        except Exception as e:
            err = classify_error(e, model)
            fix = _job_fix_for(str(e), attempt) if err.kind == "row" and fixes < _MAX_FIX else None
            if fix is None:
                raise err from e
            set_, drop, rename = fix
            if _no_progress(attempt, set_, drop, rename):
                raise err from e
            fixes += 1
            _learn(quirk_key, set_=set_, drop=drop, rename=rename)
            continue
        latency = int((time.monotonic() - started) * 1000)
        if not hasattr(resp, "choices"):
            # 200 でも HTML（プロキシのブロック画面・接続先URLの誤り）などは SDK が文字列のまま返す
            raise LLMCallError("fatal", "AIの接続先がAPIの形式で応答しませんでした（接続先URLやプロキシを確認してください）。",
                               None, None, model)
        choice = resp.choices[0] if resp.choices else None
        text = (choice.message.content if choice and choice.message else "") or ""
        usage = getattr(resp, "usage", None)
        headers = {k.lower(): v for k, v in raw.headers.items()
                   if k.lower().startswith(("x-ratelimit", "retry-after", "x-request-id"))}
        return ChatResult(
            text=_THINK_RE.sub("", text).strip(),
            finish_reason=getattr(choice, "finish_reason", None),
            tokens_in=getattr(usage, "prompt_tokens", None),
            tokens_out=getattr(usage, "completion_tokens", None),
            latency_ms=latency,
            headers=headers,
            params={k: v for k, v in attempt.items() if k != "messages"},
        )


def response_format_for(mode: str, schema: dict | None, name: str = "result") -> dict | None:
    """方式に応じた response_format（prompt_only は None）。"""
    if mode == "json_schema" and schema:
        return {"type": "json_schema", "json_schema": {"name": name, "schema": schema, "strict": False}}
    if mode in ("json_schema", "json_object"):
        return {"type": "json_object"}
    return None


def parse_json_text(text: str) -> dict:
    """応答から JSON オブジェクトを取り出す（``` 囲み・前後の文も可）。取れなければ ValueError。"""
    text = _THINK_RE.sub("", text or "").strip()
    try:
        data = json.loads(text)
    except ValueError:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if not m:
            raise ValueError("AIの応答をJSONとして読めませんでした。")
        try:
            data = json.loads(m.group(0))
        except ValueError as e:
            raise ValueError("AIの応答をJSONとして読めませんでした。") from e
    if not isinstance(data, dict):
        raise ValueError("AIの応答がJSONオブジェクトではありません。")
    return data


def detect_structured_mode(settings: dict, refresh: bool = False) -> str:
    """構造化出力の方式を json_schema → json_object → prompt_only の順に試して決める。

    エンドポイント×モデルごとにメモリへ保存する（再起動で判定し直す）。
    fatal / retry の失敗は LLMCallError のまま返す（判定結果は保存しない）。
    """
    key = (str(settings.get("chat_url") or settings.get("base_url") or ""), str(settings.get("model") or ""))
    with _lock:
        if not refresh and key in _MODES:
            return _MODES[key]
    probe = [
        {"role": "system", "content": "JSONだけを出力してください。"},
        {"role": "user", "content": '次のJSONをそのまま返してください: {"ok": true}'},
    ]
    mode = "prompt_only"
    settled = True
    for candidate in ("json_schema", "json_object"):
        rf = response_format_for(candidate, _PROBE_SCHEMA, "probe")
        outcome = "rejected"
        for budget in _PROBE_BUDGETS:
            try:
                res = chat_raw(settings, probe, response_format=rf, max_tokens=budget)
            except LLMCallError as e:
                if e.kind == "row":
                    break                # response_format を受け付けない（400 など）→ 次の方式
                raise
            try:
                parse_json_text(res.text)
                outcome = "ok"
                break
            except ValueError:
                if res.finish_reason != "length":
                    break                # 打ち切りではないのに JSON でない → 次の方式
                outcome = "truncated"    # 推論モデルが上限を考える途中で使い切った → 上限を増やしてもう一度
        if outcome == "rejected":
            continue
        # 打ち切りのまま（判定しきれない）でも、指定そのものは受け付けたのでこの方式を使う。覚えずに次回また判定する
        mode, settled = candidate, outcome == "ok"
        break
    if settled:
        with _lock:
            _MODES[key] = mode
    return mode


def forget_structured_modes() -> None:
    with _lock:
        _MODES.clear()
