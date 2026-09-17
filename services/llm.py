"""OpenAI互換APIへの接続とモデル選択。

aiagent_minimal_rag_tougou の llm.py / models.py と同じ仕様:
  - 接続先とAPIキーは env（OPENAI_BASE_URL / OPENAI_API_KEY）が既定。
    画面（AI設定）で保存した data/model_settings.yaml の値があればそちらを優先する。
  - URL はフルパス2本（…/chat/completions と …/models）で持つ。
  - 選べるモデルは「画面で登録した候補 → env の OPENAI_MODELS」の順。既定モデルは必ず候補に入る。
  - モデルが受け付けない引数（temperature / max_tokens / reasoning_effort）はエラー文を見て直して投げ直す。
  - 429 はサーバの指示した時間だけ待って投げ直す。
"""
from __future__ import annotations

import hashlib
import ipaddress
import json
import re
import threading
import time
from dataclasses import dataclass, field
from urllib.parse import urlsplit

from flask import current_app
from openai import OpenAI

from services import settings_store

SETTINGS_FILE = "model_settings.yaml"
PREFS_FILE = "prefs.yaml"
ADMIN_KEYS = ("models", "default", "api_key", "chat_url", "models_url")
CATALOG_TTL = 300
_MAX_FIX = 4

_clients: dict[tuple[str, str], OpenAI] = {}
_catalog_cache: dict[tuple[str, str], tuple[float, list[str]]] = {}
_QUIRKS: dict[str, dict] = {}
# ジョブの並列呼び出しから _clients / _QUIRKS / 方式判定を同時に更新するためのロック
_lock = threading.RLock()


class LLMNotConfigured(Exception):
    pass


class RateLimited(Exception):
    """待って投げ直しても解消しなかったレート制限。"""


def _cfg(name: str):
    return current_app.config[name]


# ---- 設定の解決 ----------------------------------------------------------------

def _read_admin() -> dict:
    data = settings_store.read_yaml(SETTINGS_FILE)
    return {k: v for k, v in data.items() if k in ADMIN_KEYS}


def _env_chat_url() -> str:
    return f"{_cfg('OPENAI_BASE_URL')}/chat/completions" if _cfg("OPENAI_BASE_URL") else ""


def _env_models_url() -> str:
    return f"{_cfg('OPENAI_BASE_URL')}/models" if _cfg("OPENAI_BASE_URL") else ""


def llm_api_key() -> str:
    return str(_read_admin().get("api_key") or "").strip() or _cfg("OPENAI_API_KEY")


def llm_api_key_source() -> str:
    """"screen"（画面で保存）/ "env" / ""（未設定）"""
    if str(_read_admin().get("api_key") or "").strip():
        return "screen"
    return "env" if _cfg("OPENAI_API_KEY") else ""


def llm_chat_url() -> str:
    return str(_read_admin().get("chat_url") or "").strip() or _env_chat_url()


def llm_models_url() -> str:
    return str(_read_admin().get("models_url") or "").strip() or _env_models_url()


def is_configured() -> bool:
    return bool(llm_chat_url() and llm_api_key())


def default_model() -> str:
    return str(_read_admin().get("default") or "").strip() or _cfg("OPENAI_MODEL")


def available() -> list[str]:
    admin = [str(m).strip() for m in (_read_admin().get("models") or []) if str(m).strip()]
    names = admin or list(_cfg("OPENAI_MODELS"))
    d = default_model()
    if d and d not in names:
        names.insert(0, d)
    return names


def current_model() -> str:
    """ヘッダーのモデル選択で選んだモデル（候補から外れていれば既定モデル）。"""
    chosen = str(get_pref("model") or "").strip()
    return chosen if chosen and chosen in available() else default_model()


def choose_model(model: str) -> str:
    model = str(model or "").strip()
    if not model:
        raise ValueError("モデル名が空です。")
    if model not in available():
        raise ValueError(f"{model} は選べません。AI設定で候補に追加してください。")
    set_pref("model", model)
    return model


def get_pref(key: str, default=None):
    return settings_store.read_yaml(PREFS_FILE).get(key, default)


def set_pref(key: str, value) -> None:
    with settings_store.lock():
        prefs = settings_store.read_yaml(PREFS_FILE)
        prefs[key] = value
        settings_store.write_yaml(PREFS_FILE, prefs)


# ---- クライアント ---------------------------------------------------------------

def _derived_base(full_url: str, suffix: str) -> str:
    u = str(full_url or "").strip().rstrip("/")
    return u[: -len(suffix)] if u.endswith(suffix) else ""


def _client_for(base: str) -> OpenAI:
    key = (base, llm_api_key())
    with _lock:
        if key not in _clients:
            _clients[key] = OpenAI(base_url=base or None, api_key=key[1] or "not-set")
        return _clients[key]


def client() -> OpenAI:
    return _client_for(_derived_base(llm_chat_url(), "/chat/completions"))


def models_client() -> OpenAI:
    return _client_for(_derived_base(llm_models_url(), "/models"))


def reset_llm_client() -> None:
    with _lock:
        _clients.clear()
        _job_clients.clear()
        _catalog_cache.clear()


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


# ---- 画面からの保存 --------------------------------------------------------------

def admin_status() -> dict:
    admin = _read_admin()
    return {
        "models": available(),
        "default": default_model(),
        "current": current_model(),
        "from_env": not admin.get("models"),
        "env_models": list(_cfg("OPENAI_MODELS")),
        "env_default": _cfg("OPENAI_MODEL"),
        "settings_file": str(settings_store.data_path(SETTINGS_FILE)),
        "llm_ready": is_configured(),
        # APIキーの値は返さない。設定済みかどうかと出所だけ
        "api_key_set": bool(llm_api_key()),
        "api_key_source": llm_api_key_source(),
        "chat_url": llm_chat_url(),
        "models_url": llm_models_url(),
        "chat_url_source": "screen" if str(admin.get("chat_url") or "").strip() else "env",
        "models_url_source": "screen" if str(admin.get("models_url") or "").strip() else "env",
        "env_chat_url": _env_chat_url(),
        "env_models_url": _env_models_url(),
    }


def save_admin(data: dict) -> dict:
    """「AI設定」画面からの保存。"""
    models = [str(m).strip() for m in (data.get("models") or []) if str(m).strip()]
    if not models:
        raise ValueError("選択できるモデルを1つ以上残してください。")
    if len(models) != len(set(models)):
        raise ValueError("同じモデルが重複しています。")
    for m in models:
        if len(m) > 120:
            raise ValueError(f"モデル名が長すぎます: {m[:40]}…")

    default = str(data.get("default") or "").strip() or models[0]
    if default not in models:
        raise ValueError(f"既定のモデル {default} が候補に入っていません。")

    # APIキー。値が来たときだけ更新する。応答にもログにもキーの値は出さない
    key_in = data.get("api_key")
    key_new = str(key_in).strip() if isinstance(key_in, str) else ""
    key_clear = bool(data.get("api_key_clear"))
    if key_new:
        if any(c.isspace() for c in key_new):
            raise ValueError("APIキーに空白や改行が入っています。コピーし直してください。")
        if not (8 <= len(key_new) <= 500):
            raise ValueError("APIキーの長さが不自然です。値を確かめてください。")

    with settings_store.lock():
        keep = settings_store.read_yaml(SETTINGS_FILE)  # 保存済みのAPIキーを巻き添えで消さないため上書きで重ねる
        keep.update({"models": models, "default": default})

        # 接続先URL（フルパス2本）。空欄で保存すると env の値に戻る。env と同じ値なら上書きとして持たない
        url_changed = False
        for field, suffix, envval, label in (
                ("chat_url", "/chat/completions", _env_chat_url(), "チャット"),
                ("models_url", "/models", _env_models_url(), "モデル一覧")):
            if field not in data:
                continue
            u = str(data.get(field) or "").strip().rstrip("/")
            if u:
                if any(c.isspace() for c in u):
                    raise ValueError(f"{label}のURLに空白が入っています。")
                if not u.startswith(("http://", "https://")):
                    raise ValueError(f"{label}のURLは http:// か https:// で始めてください。")
                if not u.endswith(suffix):
                    raise ValueError(f"{label}のURLは {suffix} で終わるフルパスで入力してください"
                                     f"（例: https://api.openai.com/v1{suffix}）。")
            old = str(keep.get(field) or "").strip()
            new = "" if u == envval else u
            if new:
                keep[field] = new
            else:
                keep.pop(field, None)
            url_changed = url_changed or old != new

        if key_clear:
            keep.pop("api_key", None)
        elif key_new:
            keep["api_key"] = key_new
        settings_store.write_yaml(SETTINGS_FILE, keep)

    if key_clear or key_new or url_changed:
        reset_llm_client()  # 次のAI呼び出しから新しい接続先・キーを使う（再起動不要）
    print(f"[models] モデル設定を更新しました: 候補{len(models)}件 / 既定={default}"
          + (" / APIキーを更新" if key_new else "")
          + (" / APIキーをenvに戻した" if key_clear else "")
          + (" / 接続先URLを変更" if url_changed else ""))
    return admin_status()


# ---- 呼び出し ------------------------------------------------------------------

def _fix_for(message: str, kwargs: dict) -> tuple | None:
    """400エラーの文面から、引数の直し方 (set, drop) を決める。"""
    low = message.lower()
    if "reasoning_effort" in low and "does not support" in low:
        return ({"reasoning_effort": "none"}, None)
    if "reasoning_effort" in low and "unrecognized" in low:
        return (None, "reasoning_effort")
    for name in ("temperature", "top_p"):
        if f"'{name}'" in low and ("does not support" in low or "only the default" in low or "unsupported" in low):
            return (None, name)
    if "max_tokens" in low and "max_completion_tokens" in low:
        v = kwargs.get("max_tokens")
        if v is not None:
            return ({"max_completion_tokens": v}, "max_tokens")
    return None


def _learn(model: str, set_: dict | None = None, drop: str | None = None) -> None:
    with _lock:
        quirk = _QUIRKS.setdefault(model, {"set": {}, "drop": set()})
        if set_:
            quirk["set"].update(set_)
        if drop:
            quirk["drop"].add(drop)
            quirk["set"].pop(drop, None)


def _apply_quirks(kwargs: dict) -> dict:
    attempt = dict(kwargs)
    with _lock:
        quirk = _QUIRKS.get(str(kwargs.get("model") or ""))
        if quirk:
            for name in quirk["drop"]:
                attempt.pop(name, None)
            attempt.update(quirk["set"])
    return attempt


_RETRY_IN = re.compile(r"try again in\s+([\d.]+)\s*(ms|s|m)\b", re.IGNORECASE)


def _is_rate_limit(e: Exception) -> bool:
    return getattr(e, "status_code", None) == 429 or "rate_limit" in str(e).lower()


def _rate_limit_wait(e: Exception, attempt: int) -> float:
    max_wait = _cfg("LLM_RATE_LIMIT_MAX_WAIT")
    try:
        raw = e.response.headers.get("retry-after")
        if raw:
            return min(float(raw), max_wait)
    except Exception:
        pass
    m = _RETRY_IN.search(str(e))
    if m:
        value, unit = float(m.group(1)), m.group(2).lower()
        sec = value / 1000 if unit == "ms" else (value * 60 if unit == "m" else value)
        return min(sec + 0.5, max_wait)
    return min(2.0 ** attempt, max_wait)


def _create(**kwargs):
    """chat.completions.create の呼び出し口。受け付けない引数はエラーを見て直しながら投げ直す。"""
    model = str(kwargs.get("model") or "")
    attempt = _apply_quirks(kwargs)
    fixes = waits = 0
    waited_total = 0.0
    while fixes < _MAX_FIX:
        try:
            return client().chat.completions.create(**attempt)
        except Exception as e:
            if _is_rate_limit(e):
                if waits >= _cfg("LLM_RATE_LIMIT_RETRIES"):
                    raise RateLimited(f"混み合っています（レート制限）。{waits}回・合計{waited_total:.1f}秒待ちましたが"
                                      "解消しませんでした。少し時間をおいてから再実行してください。") from e
                waits += 1
                sec = _rate_limit_wait(e, waits)
                waited_total += sec
                print(f"[llm] レート制限。{sec:.1f}秒待って投げ直します（{waits}/{_cfg('LLM_RATE_LIMIT_RETRIES')}回目）")
                time.sleep(sec)
                continue
            fixes += 1
            fix = _fix_for(str(e), attempt)
            if fix is None:
                raise
            set_, drop = fix
            if set_ and all(attempt.get(k) == v for k, v in set_.items()):
                raise
            if drop and drop not in attempt:
                raise
            _learn(model, set_=set_, drop=drop)
            attempt = _apply_quirks(kwargs)
    return client().chat.completions.create(**attempt)


def ask_json(system: str, user: str, what: str = "AIの応答", model: str | None = None) -> dict:
    """AIに聞いて、応答からJSONオブジェクトを取り出す（``` で囲まれていても可）。"""
    if not is_configured():
        raise LLMNotConfigured("AIの接続先が未設定です。「AI設定」でAPIキーと接続先を設定してください。")
    kwargs = dict(
        model=model or current_model(),
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        temperature=_cfg("OPENAI_TEMPERATURE"),
    )
    if _cfg("OPENAI_TOP_P") is not None:
        kwargs["top_p"] = _cfg("OPENAI_TOP_P")
    if _cfg("OPENAI_MAX_TOKENS") is not None:
        kwargs["max_tokens"] = _cfg("OPENAI_MAX_TOKENS")
    resp = _create(**kwargs)
    content = resp.choices[0].message.content or ""
    m = re.search(r"\{.*\}", content, re.DOTALL)
    if not m:
        raise ValueError(f"{what}をJSONとして解析できませんでした: {content[:200]}")
    data = json.loads(m.group(0))
    if not isinstance(data, dict):
        raise ValueError(f"{what}が想定した形式ではありません。")
    return data


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, (LLMNotConfigured, RateLimited, ValueError)):
        return str(exc)
    status = getattr(exc, "status_code", None)
    if status in (401, 403):
        return "APIキーが拒否されました。「AI設定」でキーを確認してください。"
    if status == 404:
        return f"モデルまたは接続先が見つかりません（{current_model()}）。「AI設定」を確認してください。"
    text = str(exc)
    return f"AI呼び出しに失敗しました: {text[:160]}"


# ---- ジョブ用の呼び出し口（一覧表のAI整形） -------------------------------------------
# ask_json（帳票側）とは別に持つ。設定はジョブ開始時に固定し、専用クライアント（明示タイムアウト・SDK再試行なし）で呼ぶ。
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
        raise LLMNotConfigured("AIの接続先が未設定です。「AI設定」でAPIキーと接続先を設定してください。")
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
        return LLMCallError("fatal", "AIの接続先に接続できませんでした。接続先URLとサーバーの起動を確認してください。",
                            None, None, model)
    if status in (401, 403):
        return LLMCallError("fatal", "APIキーが拒否されました。「AI設定」でキーを確認してください。", status, None, model)
    if status == 402 or "insufficient_quota" in low:
        return LLMCallError("fatal", "APIの残高・利用枠が不足しています。", status, None, model)
    if status == 404 or ("model" in low and ("not found" in low or "does not exist" in low)):
        return LLMCallError("fatal", f"モデルまたは接続先が見つかりません（{model}）。「AI設定」を確認してください。",
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
        return (None, "seed")
    return None


def chat_raw(settings: dict, messages: list[dict], response_format: dict | None = None,
             max_tokens: int | None = None, timeout: float | None = None) -> ChatResult:
    """1回の chat 呼び出し（SDK の再試行なし）。受け付けない引数だけはエラー文から直して投げ直す。

    失敗は LLMCallError（kind=fatal/retry/row）。レート制限の待機と再試行は呼び出し側（ジョブ）で行う。
    """
    model = str(settings.get("model") or "")
    timeout = float(timeout or settings.get("timeout") or CLOUD_TIMEOUT)
    kwargs = dict(settings.get("params") or {})
    kwargs.update(model=model, messages=messages)
    if response_format:
        kwargs["response_format"] = response_format
    if max_tokens:
        kwargs["max_tokens"] = int(max_tokens)
    cli = _job_client(settings, timeout)
    fixes = 0
    while True:
        attempt = _apply_quirks(kwargs)
        started = time.monotonic()
        try:
            raw = cli.chat.completions.with_raw_response.create(**attempt)
            resp = raw.parse()
        except Exception as e:
            err = classify_error(e, model)
            fix = _job_fix_for(str(e), attempt) if err.kind == "row" and fixes < _MAX_FIX else None
            if fix is None:
                raise err from e
            set_, drop = fix
            if (set_ and all(attempt.get(k) == v for k, v in set_.items())) or (drop and drop not in attempt):
                raise err from e
            fixes += 1
            _learn(model, set_=set_, drop=drop)
            continue
        latency = int((time.monotonic() - started) * 1000)
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
    for candidate in ("json_schema", "json_object"):
        try:
            res = chat_raw(settings, probe, response_format=response_format_for(candidate, _PROBE_SCHEMA, "probe"),
                           max_tokens=50)
        except LLMCallError as e:
            if e.kind == "row":
                continue
            raise
        try:
            parse_json_text(res.text)
        except ValueError:
            continue
        mode = candidate
        break
    with _lock:
        _MODES[key] = mode
    return mode


def forget_structured_modes() -> None:
    with _lock:
        _MODES.clear()
