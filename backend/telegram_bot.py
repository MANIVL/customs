#!/usr/bin/env python3
"""Telegram bot for ТаможенФормат.

Same Excel flow as the website, without SVH search:
upload files, choose AI or keyword fill, receive the workbook and an error protocol.
"""
from __future__ import annotations

import asyncio
import json
import os
import traceback
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass, field

import ai_mapper
import engine
import yandex_gpt
from fastapi import Request
from fastapi.responses import JSONResponse
from starlette.background import BackgroundTasks

_sessions: dict[int, "Session"] = {}
_locks: dict[int, asyncio.Lock] = {}
_mounted = False

WELCOME = (
    "ТаможенФормат\n\n"
    "Пришлите один или несколько Excel-файлов (.xlsx, .xls): "
    "инвойс, спецификацию, упаковочный лист.\n\n"
    "Когда все файлы будут здесь, нажмите «Файлы загружены» "
    "и выберите обработку — с ИИ или без ИИ.\n\n"
    "Бот пришлёт готовый файл: откройте его и выберите, куда сохранить. "
    "Следом придёт протокол ошибок.\n\n"
    "Эталонный шаблон уже на сервере. Заменить его: /template\n"
    "Страна и единица, если их нет в файле: /settings"
)

SAVE_CAPTION = (
    "Готовый файл. Нажмите на него и выберите, куда сохранить: "
    "в загрузки, на диск или в другое место."
)


class TelegramError(RuntimeError):
    pass


@dataclass
class Session:
    files: list[tuple[str, bytes]] = field(default_factory=list)
    country: str = "CN"
    unit: str = "шт"
    awaiting: str | None = None
    busy: bool = False
    result: bytes | None = None
    result_name: str = "result.xlsx"


def token() -> str:
    return os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()


def require_token() -> str:
    value = token()
    if not value:
        raise TelegramError("TELEGRAM_BOT_TOKEN не задан")
    return value


def _session(chat_id: int) -> Session:
    session = _sessions.get(chat_id)
    if session is None:
        session = Session()
        _sessions[chat_id] = session
    return session


def _lock(chat_id: int) -> asyncio.Lock:
    lock = _locks.get(chat_id)
    if lock is None:
        lock = asyncio.Lock()
        _locks[chat_id] = lock
    return lock


def _api(method: str, payload: dict | None = None, *, timeout: int = 60):
    url = f"https://api.telegram.org/bot{require_token()}/{method}"
    data = json.dumps(payload or {}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")[:500]
        raise TelegramError(f"Telegram {method}: {detail}") from e
    if not body.get("ok"):
        raise TelegramError(body.get("description") or method)
    return body.get("result")


def _chunks(text: str, limit: int = 3900) -> list[str]:
    text = (text or "").strip() or "Протокол пуст."
    if len(text) <= limit:
        return [text]
    parts: list[str] = []
    buf = ""
    for line in text.split("\n"):
        while len(line) > limit:
            if buf:
                parts.append(buf)
                buf = ""
            parts.append(line[:limit])
            line = line[limit:]
        if buf and len(buf) + 1 + len(line) > limit:
            parts.append(buf)
            buf = line
        else:
            buf = f"{buf}\n{line}" if buf else line
    if buf:
        parts.append(buf)
    return parts or ["Протокол пуст."]


def _is_general_warning(message: str) -> bool:
    return (
        message.startswith("ИИ недоступен")
        or message.startswith("Часть файлов")
        or message.startswith("Лист «")
    )


def format_protocol(protocol: dict | None) -> str:
    protocol = protocol or {}
    lines = ["Протокол обработки", ""]
    lines.append(protocol.get("summary") or "Обработка завершена.")
    lines.append("")

    notes = [str(m) for m in (protocol.get("notes") or []) if m]
    errors = [str(m) for m in (protocol.get("errors") or []) if m]
    warnings = [str(m) for m in (protocol.get("warnings") or []) if m]
    unknown = [str(m) for m in (protocol.get("unknown") or []) if m]
    findings = protocol.get("findings") or []

    if notes:
        lines.append("Заметки:")
        lines.extend(f"• {m}" for m in notes)
        lines.append("")

    if findings:
        general = [m for m in warnings if _is_general_warning(m)]
        if general:
            lines.append("Замечания:")
            lines.extend(f"• {m}" for m in general)
            lines.append("")
        lines.append("По позициям:")
        for finding in findings:
            message = (finding or {}).get("message") or ""
            if message:
                lines.append(f"• {message}")
        lines.append("")
    else:
        if errors:
            lines.append("Ошибки:")
            lines.extend(f"• {m}" for m in errors)
            lines.append("")
        if warnings:
            lines.append("Замечания:")
            lines.extend(f"• {m}" for m in warnings)
            lines.append("")

    if unknown:
        lines.append("Неизвестные элементы:")
        lines.extend(f"• {m}" for m in unknown)
        lines.append("")

    if not errors and not warnings and not findings and not unknown:
        lines.append("Ошибок не обнаружено, файл обработан корректно.")

    return "\n".join(lines).strip()


def _file_list(session: Session) -> str:
    return "\n".join(f"• {name}" for name, _content in session.files)


def kb_collect(session: Session) -> dict | None:
    if not session.files:
        return None
    return {
        "inline_keyboard": [
            [{"text": "Файлы загружены", "callback_data": "files_done"}],
            [{"text": "Очистить список", "callback_data": "files_clear"}],
        ]
    }


def kb_mode() -> dict:
    return {
        "inline_keyboard": [
            [
                {"text": "С ИИ", "callback_data": "mode_ai"},
                {"text": "Без ИИ", "callback_data": "mode_keyword"},
            ]
        ]
    }


def kb_after() -> dict:
    return {
        "inline_keyboard": [
            [{"text": "Отправить файл ещё раз", "callback_data": "resend"}],
            [
                {"text": "С ИИ", "callback_data": "mode_ai"},
                {"text": "Без ИИ", "callback_data": "mode_keyword"},
            ],
            [{"text": "Новая партия", "callback_data": "reset"}],
        ]
    }


def choose_text(session: Session) -> str:
    return (
        "Файлы загружены. Как их обработать?\n\n"
        f"{_file_list(session)}\n\n"
        f"Если в файле нет страны, подставится {session.country}. "
        f"Если нет единицы — {session.unit}."
    )


def send_message(chat_id: int, text: str, reply_markup: dict | None = None) -> None:
    payload: dict = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    _api("sendMessage", payload)


def edit_message(chat_id: int, message_id: int, text: str, reply_markup: dict | None = None) -> None:
    payload: dict = {"chat_id": chat_id, "message_id": message_id, "text": text}
    if reply_markup is not None:
        payload["reply_markup"] = reply_markup
    _api("editMessageText", payload)


def answer_callback(callback_id: str, text: str = "") -> None:
    payload: dict = {"callback_query_id": callback_id}
    if text:
        payload["text"] = text[:180]
    try:
        _api("answerCallbackQuery", payload)
    except TelegramError:
        traceback.print_exc()


def send_document(chat_id: int, filename: str, content: bytes, caption: str | None = None) -> None:
    safe = (filename or "result.xlsx").replace("\r", "").replace("\n", "").replace('"', "")
    boundary = f"----TfBot{uuid.uuid4().hex}"

    def field(name: str, value: str) -> bytes:
        return (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
            f"{value}\r\n"
        ).encode("utf-8")

    quoted = urllib.parse.quote(safe)
    head = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="document"; filename="{safe}"; '
        f"filename*=UTF-8''{quoted}\r\n"
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ).encode("utf-8")
    body = field("chat_id", str(chat_id))
    if caption:
        body += field("caption", caption[:1024])
    body += head + content + f"\r\n--{boundary}--\r\n".encode("utf-8")
    url = f"https://api.telegram.org/bot{require_token()}/sendDocument"
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")[:500]
        raise TelegramError(f"Не удалось отправить файл: {detail}") from e
    if not result.get("ok"):
        raise TelegramError(result.get("description") or "sendDocument")


def download_telegram_file(file_id: str) -> bytes:
    meta = _api("getFile", {"file_id": file_id}, timeout=60)
    path = (meta or {}).get("file_path")
    if not path:
        raise TelegramError("Telegram не отдал файл")
    url = f"https://api.telegram.org/file/bot{require_token()}/{path}"
    with urllib.request.urlopen(url, timeout=180) as resp:
        return resp.read()


def process_excel(mode: str, files: list[tuple[str, bytes]], settings: dict) -> tuple[bytes, dict]:
    import server

    template = server.load_template_bytes()
    contents = [content for _name, content in files]
    names = [name for name, _content in files]
    if mode == "ai":
        if not yandex_gpt.is_configured():
            raise RuntimeError(
                "YandexGPT не настроен. Добавьте YANDEX_API_KEY и YANDEX_FOLDER_ID "
                "или выберите обработку без ИИ."
            )
        try:
            return ai_mapper.process_with_ai(template, contents, settings, names)
        except Exception as e:
            try:
                result_bytes, report = engine.process(template, contents, settings)
            except Exception as e2:
                raise RuntimeError(
                    f"{server._public_error(e2)} (после ошибки ИИ: {server._public_error(e)})"
                ) from e2
            report = {**report, "mode": "keyword_fallback", "ai_error": str(e)}
            proto = dict(report.get("protocol") or {})
            warns = list(proto.get("warnings") or [])
            warns.insert(0, "ИИ недоступен или не ответил — файл собран по заголовкам.")
            proto["warnings"] = warns
            summary = proto.get("summary") or ""
            if summary and "замечан" not in summary and "ошибк" not in summary:
                proto["summary"] = summary.rstrip(".") + ". Есть замечания."
            report["protocol"] = proto
            return result_bytes, report
    return engine.process(template, contents, settings)


def reset_batch(session: Session) -> None:
    session.files.clear()
    session.result = None
    session.result_name = "result.xlsx"
    session.awaiting = None
    session.busy = False


def _document_name(document: dict) -> str:
    name = (document.get("file_name") or "").strip()
    if name:
        return name
    mime = (document.get("mime_type") or "").lower()
    if mime == "application/vnd.ms-excel":
        return "input.xls"
    return "input.xlsx"


def _user_error(exc: BaseException) -> str:
    if isinstance(exc, (RuntimeError, ValueError, FileNotFoundError, TelegramError)) and str(exc).strip():
        return str(exc).strip()
    import server

    return server._public_error(exc)


async def _send_text(chat_id: int, text: str, reply_markup: dict | None = None) -> None:
    parts = _chunks(text)
    for index, part in enumerate(parts):
        markup = reply_markup if index == len(parts) - 1 else None
        await asyncio.to_thread(send_message, chat_id, part, markup)


async def _deliver_result(chat_id: int, session: Session, report: dict) -> None:
    await asyncio.to_thread(
        send_document,
        chat_id,
        session.result_name,
        session.result or b"",
        SAVE_CAPTION,
    )
    await _send_text(chat_id, format_protocol(report.get("protocol")), kb_after())


async def _run_mode(chat_id: int, session: Session, mode: str, status_message_id: int | None) -> None:
    if session.busy:
        await _send_text(chat_id, "Уже обрабатываю предыдущий запрос.")
        return
    if not session.files:
        await _send_text(chat_id, "Сначала пришлите хотя бы один Excel-файл.", kb_collect(session))
        return

    session.busy = True
    session.awaiting = None
    label = "с ИИ" if mode == "ai" else "без ИИ"
    status = f"Обрабатываю файлы {label}. Это может занять минуту."
    try:
        if status_message_id:
            try:
                await asyncio.to_thread(
                    edit_message,
                    chat_id,
                    status_message_id,
                    status,
                    {"inline_keyboard": []},
                )
            except TelegramError:
                await _send_text(chat_id, status)
        else:
            await _send_text(chat_id, status)
        try:
            await asyncio.to_thread(_api, "sendChatAction", {"chat_id": chat_id, "action": "upload_document"})
        except TelegramError:
            pass
        settings = {"country": session.country, "unit": session.unit}
        result_bytes, report = await asyncio.to_thread(process_excel, mode, list(session.files), settings)
        session.result = result_bytes
        session.result_name = "result.xlsx"
        await _deliver_result(chat_id, session, report)
    except Exception as e:
        traceback.print_exc()
        await _send_text(
            chat_id,
            "Не удалось обработать файлы.\n\n" + _user_error(e),
            kb_mode() if session.files else None,
        )
    finally:
        session.busy = False


async def _add_input_file(chat_id: int, session: Session, document: dict) -> None:
    import server

    if session.busy:
        await _send_text(chat_id, "Дождитесь окончания обработки, затем пришлите файлы ещё раз.")
        return
    if len(session.files) >= server.MAX_FILES:
        await _send_text(chat_id, f"Слишком много файлов (максимум {server.MAX_FILES}).")
        return
    size = int(document.get("file_size") or 0)
    if size > server.MAX_FILE_SIZE:
        await _send_text(chat_id, "Файл слишком большой (максимум 20 МБ).")
        return
    name = _document_name(document)
    try:
        content = await asyncio.to_thread(download_telegram_file, document["file_id"])
        server._validate_excel(name, content)
    except Exception as e:
        traceback.print_exc()
        await _send_text(chat_id, _user_error(e))
        return
    session.files.append((name, content))
    session.result = None
    session.awaiting = None
    await _send_text(
        chat_id,
        (
            f"Добавлен файл «{name}».\n"
            f"Сейчас файлов: {len(session.files)}.\n\n"
            f"{_file_list(session)}\n\n"
            "Пришлите ещё файлы или нажмите «Файлы загружены»."
        ),
        kb_collect(session),
    )


async def _save_template(chat_id: int, session: Session, document: dict) -> None:
    import server

    name = _document_name(document)
    try:
        content = await asyncio.to_thread(download_telegram_file, document["file_id"])
        meta = server.store_template(name, content)
    except Exception as e:
        traceback.print_exc()
        await _send_text(chat_id, "Шаблон не сохранён.\n\n" + _user_error(e))
        return
    session.awaiting = None
    await _send_text(
        chat_id,
        f"Шаблон «{meta.get('filename') or name}» сохранён. Он используется и на сайте, и в боте.\n\n"
        "Пришлите входные Excel-файлы.",
        kb_collect(session),
    )


async def _on_message(message: dict) -> None:
    chat_id = (message.get("chat") or {}).get("id")
    if chat_id is None:
        return
    session = _session(int(chat_id))
    text = (message.get("text") or "").strip()
    command = text.split()[0].lower() if text.startswith("/") else ""
    command = command.split("@", 1)[0]

    if command in ("/start", "/help"):
        reset_batch(session)
        await _send_text(chat_id, WELCOME)
        return
    if command == "/cancel":
        reset_batch(session)
        await _send_text(chat_id, "Сброшено. Пришлите Excel-файлы заново.")
        return
    if command == "/template":
        session.awaiting = "template"
        await _send_text(
            chat_id,
            "Пришлите эталонный шаблон в формате .xlsx. "
            "Он заменит текущий шаблон и на сайте, и в боте.",
        )
        return
    if command == "/settings":
        session.awaiting = "settings"
        await _send_text(
            chat_id,
            f"Сейчас: страна {session.country}, единица {session.unit}.\n"
            "Пришлите новые значения через пробел, например: CN шт",
        )
        return

    document = message.get("document")
    if document:
        if session.awaiting == "template":
            await _save_template(chat_id, session, document)
        else:
            await _add_input_file(chat_id, session, document)
        return

    if message.get("photo") or message.get("video") or message.get("audio"):
        await _send_text(chat_id, "Нужен Excel-файл (.xlsx или .xls), отправленный как документ.")
        return

    if session.awaiting == "settings" and text:
        parts = text.split(None, 1)
        if len(parts) < 2 or len(parts[0]) > 40 or len(parts[1]) > 40:
            await _send_text(chat_id, "Напишите страну и единицу через пробел, например: IT шт")
            return
        session.country = parts[0].strip()
        session.unit = parts[1].strip()
        session.awaiting = None
        await _send_text(
            chat_id,
            f"Сохранено: страна {session.country}, единица {session.unit}.",
            kb_collect(session),
        )
        return

    if text:
        await _send_text(
            chat_id,
            "Пришлите Excel-файл или нажмите «Файлы загружены», если файлы уже в чате.",
            kb_collect(session),
        )


async def _on_callback(callback: dict) -> None:
    message = callback.get("message") or {}
    chat_id = (message.get("chat") or {}).get("id")
    if chat_id is None:
        return
    session = _session(int(chat_id))
    data = callback.get("data") or ""
    message_id = message.get("message_id")

    if data == "files_clear":
        session.files.clear()
        session.result = None
        await _send_text(chat_id, "Список очищен. Пришлите Excel-файлы заново.")
        return
    if data == "reset":
        reset_batch(session)
        await _send_text(chat_id, "Новая партия. Пришлите Excel-файлы.")
        return
    if data == "files_done":
        if not session.files:
            await _send_text(chat_id, "Сначала пришлите хотя бы один Excel-файл.")
            return
        session.awaiting = None
        text = choose_text(session)
        if message_id:
            try:
                await asyncio.to_thread(edit_message, chat_id, message_id, text, kb_mode())
                return
            except TelegramError:
                traceback.print_exc()
        await _send_text(chat_id, text, kb_mode())
        return
    if data in ("mode_ai", "mode_keyword"):
        mode = "ai" if data == "mode_ai" else "keyword"
        await _run_mode(chat_id, session, mode, message_id)
        return
    if data == "resend":
        if not session.result:
            await _send_text(chat_id, "Сначала обработайте файлы.", kb_mode() if session.files else None)
            return
        await asyncio.to_thread(
            send_document,
            chat_id,
            session.result_name,
            session.result,
            SAVE_CAPTION,
        )
        return


async def handle_update(update: dict) -> None:
    try:
        if not token():
            return
        callback = update.get("callback_query")
        message = update.get("message")
        if callback and callback.get("id"):
            await asyncio.to_thread(answer_callback, callback["id"])
        chat_id = None
        if callback:
            chat_id = ((callback.get("message") or {}).get("chat") or {}).get("id")
        elif message:
            chat_id = (message.get("chat") or {}).get("id")
        if chat_id is None:
            return
        async with _lock(int(chat_id)):
            if callback:
                await _on_callback(callback)
            elif message:
                await _on_message(message)
    except Exception:
        traceback.print_exc()


def configure_webhook() -> None:
    if not token():
        print("Telegram bot: TELEGRAM_BOT_TOKEN не задан, бот выключен.")
        return
    base = (
        os.environ.get("TELEGRAM_WEBHOOK_URL")
        or os.environ.get("RENDER_EXTERNAL_URL")
        or ""
    ).strip().rstrip("/")
    if not base:
        print("Telegram bot: публичный адрес не задан, webhook не устанавливается. Локально запустите python bot.py")
        return
    url = base if base.endswith("/api/telegram/webhook") else base + "/api/telegram/webhook"
    payload: dict = {
        "url": url,
        "allowed_updates": ["message", "callback_query"],
    }
    secret = os.environ.get("TELEGRAM_WEBHOOK_SECRET", "").strip()
    if secret:
        payload["secret_token"] = secret
    try:
        _api("setWebhook", payload)
        print(f"Telegram webhook: {url}")
    except Exception:
        traceback.print_exc()


def mount(app) -> None:
    """Attach the webhook route to the existing FastAPI app."""
    global _mounted
    if _mounted:
        return
    _mounted = True

    @app.post("/api/telegram/webhook")
    async def telegram_webhook(request: Request, background_tasks: BackgroundTasks):
        if not token():
            return JSONResponse({"ok": False, "error": "bot disabled"}, status_code=503)
        secret = os.environ.get("TELEGRAM_WEBHOOK_SECRET", "").strip()
        if secret and request.headers.get("x-telegram-bot-api-secret-token") != secret:
            return JSONResponse({"ok": False}, status_code=401)
        update = await request.json()
        background_tasks.add_task(handle_update, update)
        return {"ok": True}

    @app.on_event("startup")
    async def _telegram_startup() -> None:
        await asyncio.to_thread(configure_webhook)


def run_polling() -> None:
    """Local mode. Takes the bot off the Render webhook until the web app starts again."""
    require_token()
    try:
        _api("deleteWebhook", {"drop_pending_updates": False})
    except TelegramError:
        traceback.print_exc()
    print("Telegram bot: long polling. Остановка — Ctrl+C.")

    async def _loop() -> None:
        offset = None
        while True:
            payload: dict = {
                "timeout": 25,
                "allowed_updates": ["message", "callback_query"],
            }
            if offset is not None:
                payload["offset"] = offset
            try:
                updates = await asyncio.to_thread(_api, "getUpdates", payload, timeout=40)
            except Exception:
                traceback.print_exc()
                await asyncio.sleep(3)
                continue
            for update in updates or []:
                offset = int(update["update_id"]) + 1
                await handle_update(update)

    try:
        asyncio.run(_loop())
    except KeyboardInterrupt:
        print("Telegram bot остановлен.")
