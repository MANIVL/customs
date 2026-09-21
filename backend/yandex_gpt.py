"""YandexGPT client for ТаможенФормат.

Credentials from environment (or .env next to main.py):
  YANDEX_API_KEY   — API-ключ сервисного аккаунта
  YANDEX_FOLDER_ID — идентификатор каталога
  YANDEX_MODEL     — опционально: yandexgpt-lite | yandexgpt (по умолчанию lite)
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from pathlib import Path

COMPLETION_URL = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"


def _load_dotenv() -> None:
    """Load KEY=VALUE from desktop/.env if present (does not override existing env)."""
    candidates = [
        Path(__file__).resolve().parent.parent / ".env",
        Path(__file__).resolve().parent / ".env",
    ]
    for path in candidates:
        if not path.is_file():
            continue
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
        break


_load_dotenv()


def is_configured() -> bool:
    return bool(os.environ.get("YANDEX_API_KEY") and os.environ.get("YANDEX_FOLDER_ID"))


def model_uri() -> str:
    folder = os.environ["YANDEX_FOLDER_ID"]
    model = os.environ.get("YANDEX_MODEL", "yandexgpt-lite")
    return f"gpt://{folder}/{model}"


def complete(system: str, user: str, *, temperature: float = 0.1, max_tokens: int = 3000) -> str:
    """Call YandexGPT completion API; return assistant text."""
    if not is_configured():
        raise RuntimeError("YandexGPT не настроен: задайте YANDEX_API_KEY и YANDEX_FOLDER_ID")

    timeout = int(os.environ.get("YANDEX_TIMEOUT", "20"))
    payload = {
        "modelUri": model_uri(),
        "completionOptions": {
            "stream": False,
            "temperature": temperature,
            "maxTokens": max_tokens,
        },
        "messages": [
            {"role": "system", "text": system},
            {"role": "user", "text": user},
        ],
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        COMPLETION_URL,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Api-Key {os.environ['YANDEX_API_KEY']}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"YandexGPT HTTP {e.code}: {detail}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"YandexGPT не ответил за {timeout} с ({e.reason})") from e
    except TimeoutError as e:
        raise RuntimeError(f"YandexGPT не ответил за {timeout} с") from e

    alternatives = body.get("result", {}).get("alternatives") or []
    if not alternatives:
        raise RuntimeError(f"Пустой ответ YandexGPT: {json.dumps(body, ensure_ascii=False)[:400]}")
    return alternatives[0].get("message", {}).get("text", "").strip()
