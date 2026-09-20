"""OpenAI互換APIへの接続とモデル選択。

aiagent_minimal_rag_tougou の llm.py / models.py と同じ仕様:
  - 接続先とAPIキーは env（OPENAI_BASE_URL / OPENAI_API_KEY）が既定。
    画面（AI接続）で保存した data/model_settings.yaml の値があればそちらを優先する。
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

_clients: dict[tuple[str, str, float], OpenAI] = {}
_catalog_cache: dict[tuple[str, str], tuple[float, list[str]]] = {}
# モデルが受け付けない引数の覚え書き。接続先×モデルごと（同じモデル名でも別のサーバーなら別の癖）
_QUIRKS: dict[tuple[str, str], dict] = {}
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
        raise ValueError(f"{model} は選べません。「AI接続」で候補に追加してください。")
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


MODELS_TIMEOUT = 30.0


def _client_for(base: str, timeout: float) -> OpenAI:
    """明示タイムアウト・SDK の再試行なしのクライアント。
    （SDK の既定は読み取り600秒×再試行2回で、応答しない接続先だと画面が最大30分ほど待たされる。429 は _create が待って投げ直す）"""
    key = (base, llm_api_key(), float(timeout))
    with _lock:
        if key not in _clients:
            _clients[key] = OpenAI(base_url=base or None, api_key=key[1] or "not-set", timeout=key[2], max_retries=0)
        return _clients[key]


def client() -> OpenAI:
    url = llm_chat_url()
    return _client_for(_derived_base(url, "/chat/completions"),
                       LOCAL_TIMEOUT if is_local_endpoint(url) else CLOUD_TIMEOUT)


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
    """「AI接続」画面からの保存。"""
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
    endpoint = llm_chat_url()
    key = _quirk_key(endpoint, str(kwargs.get("model") or ""))
    attempt = _apply_quirks(kwargs, endpoint)
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
            set_, drop, rename = fix
            if _no_progress(attempt, set_, drop, rename):
                raise
            _learn(key, set_=set_, drop=drop, rename=rename)
            attempt = _apply_quirks(kwargs, endpoint)
    return client().chat.completions.create(**attempt)


def ask_json(system: str, user: str, what: str = "AIの応答", model: str | None = None) -> dict:
    """AIに聞いて、応答からJSONオブジェクトを取り出す（``` で囲まれていても可）。"""
    if not is_configured():
        raise LLMNotConfigured("AIの接続先が未設定です。「AI接続」でAPIキーと接続先を設定してください。")
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
    if not hasattr(resp, "choices"):
        # 200 でも HTML（プロキシのブロック画面・接続先URLの誤り）などは SDK が文字列のまま返す
        raise ValueError("AIの接続先がAPIの形式で応答しませんでした（接続先URLやプロキシを確認してください）。")
    if not resp.choices or not resp.choices[0].message:
        raise ValueError("AIの応答が空でした。")
    content = resp.choices[0].message.content or ""
    try:
        # <think>…</think>（推論の途中に書かれた { } を含む）を除いてから取り出す。失敗は日本語の ValueError
        return parse_json_text(content)
    except ValueError as e:
        shown = _THINK_RE.sub("", content).strip()[:200]
        raise ValueError(f"{what}をJSONとして解析できませんでした（{str(e).rstrip('。')}）: {shown}") from e


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, (LLMNotConfigured, RateLimited, ValueError, LLMCallError)):
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
        raise LLMNotConfigured("AIの接続先が未設定です。「AI接続」でAPIキーと接続先を設定してください。")
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
