#!/usr/bin/env python3
"""FastAPI backend for the web build of ТаможенФормат.

Serves both the API and the static frontend from a single process.
"""
import os
import sys
import json
import base64
import shutil
import traceback
import datetime
import sqlite3
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
import engine
import ai_mapper
import yandex_gpt
import articles

from fastapi import FastAPI, UploadFile, File, Form, Query, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

# Bundled defaults live next to the code (shipped in the image / repo).
BUNDLE_DATA_DIR = os.path.join(BASE_DIR, "data")
DEFAULT_TEMPLATE_PATH = os.path.join(BUNDLE_DATA_DIR, "default_template.xlsx")

# Mutable runtime data (template upload, article catalog). On Render point DATA_DIR
# at a persistent disk mount so uploads survive restarts/redeploys.
DATA_DIR = os.environ.get("DATA_DIR") or BUNDLE_DATA_DIR
os.makedirs(DATA_DIR, exist_ok=True)
TEMPLATE_PATH = os.path.join(DATA_DIR, "current_template.xlsx")
META_PATH = os.path.join(DATA_DIR, "template_meta.json")
APP_DB_PATH = os.path.join(DATA_DIR, "app_state.db")


def _init_template_storage() -> None:
    """Ensure current template exists; restore from SQLite backup or bundled default."""
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(TEMPLATE_PATH):
        # Backfill durable blob if file exists but DB is empty (first boot on new disk).
        try:
            if _load_template_blob() is None:
                meta = {}
                if os.path.exists(META_PATH):
                    with open(META_PATH, encoding="utf-8") as f:
                        meta = json.load(f)
                with open(TEMPLATE_PATH, "rb") as f:
                    content = f.read()
                _save_template_blob(
                    content,
                    meta.get("filename") or "template.xlsx",
                    meta.get("updated_at")
                    or datetime.datetime.now(datetime.timezone.utc).isoformat(),
                )
        except Exception:
            traceback.print_exc()
        return
    # Prefer durable SQLite copy (same DATA_DIR / disk)
    try:
        restored = _load_template_blob()
        if restored:
            with open(TEMPLATE_PATH, "wb") as f:
                f.write(restored["content"])
            if not os.path.exists(META_PATH):
                with open(META_PATH, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "filename": restored.get("filename") or "template.xlsx",
                            "updated_at": restored.get("updated_at")
                            or datetime.datetime.now(datetime.timezone.utc).isoformat(),
                        },
                        f,
                        ensure_ascii=False,
                    )
            return
    except Exception:
        traceback.print_exc()
    if os.path.exists(DEFAULT_TEMPLATE_PATH):
        shutil.copy(DEFAULT_TEMPLATE_PATH, TEMPLATE_PATH)
        updated_at = datetime.datetime.now(datetime.timezone.utc).isoformat()
        meta = {
            "filename": "etalonnyi-shablon.xlsx",
            "updated_at": updated_at,
        }
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False)
        try:
            with open(TEMPLATE_PATH, "rb") as f:
                _save_template_blob(f.read(), meta["filename"], updated_at)
        except Exception:
            traceback.print_exc()


def _get_app_db() -> sqlite3.Connection:
    conn = sqlite3.connect(APP_DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def _init_app_db() -> None:
    with _get_app_db() as db:
        db.execute(
            """
            CREATE TABLE IF NOT EXISTS template_blob (
                id          INTEGER PRIMARY KEY CHECK (id = 1),
                filename    TEXT NOT NULL,
                content     BLOB NOT NULL,
                updated_at  TEXT NOT NULL
            )
            """
        )
        db.commit()


def _save_template_blob(content: bytes, filename: str, updated_at: str) -> None:
    with _get_app_db() as db:
        db.execute(
            """
            INSERT INTO template_blob (id, filename, content, updated_at)
            VALUES (1, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                filename=excluded.filename,
                content=excluded.content,
                updated_at=excluded.updated_at
            """,
            (filename, content, updated_at),
        )


def _load_template_blob() -> dict | None:
    if not os.path.exists(APP_DB_PATH):
        return None
    with _get_app_db() as db:
        row = db.execute(
            "SELECT filename, content, updated_at FROM template_blob WHERE id = 1"
        ).fetchone()
    if not row:
        return None
    return {
        "filename": row["filename"],
        "content": bytes(row["content"]),
        "updated_at": row["updated_at"],
    }


_init_app_db()
_init_template_storage()

app = FastAPI(title="ТаможенФормат")
APP_VERSION = "2026-09-22.1"


def _public_error(exc: BaseException) -> str:
    """Human-readable error; empty AssertionError becomes a file:line hint."""
    msg = (str(exc) or "").strip() or exc.__class__.__name__
    frames = traceback.extract_tb(exc.__traceback__) if exc.__traceback__ else []
    if not frames:
        return msg
    last = frames[-1]
    loc = f"{os.path.basename(last.filename)}:{last.lineno}"
    if not (str(exc) or "").strip():
        return f"{exc.__class__.__name__} в {loc} ({last.name})"
    return f"{msg} [{exc.__class__.__name__} в {loc}]"


@app.exception_handler(RequestValidationError)
async def _validation_error(_request: Request, exc: RequestValidationError):
    parts = []
    for err in exc.errors():
        loc = ".".join(str(x) for x in (err.get("loc") or ()) if x != "body")
        msg = err.get("msg") or "некорректное поле"
        parts.append(f"{loc}: {msg}" if loc else msg)
    return JSONResponse(
        {"error": "; ".join(parts) or "Некорректный запрос"},
        status_code=422,
    )


@app.middleware("http")
async def _frontend_no_cache(request, call_next):
    """Keep browsers from serving a stale frontend after a deploy."""
    response = await call_next(request)
    path = request.url.path or "/"
    if (
        path == "/"
        or path.endswith((".html", ".js", ".css"))
        or path in ("/index.html", "/app.js", "/style.css")
    ):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB per file
MAX_FILES = 30


def _safe_filename(name: str | None, fallback: str = "template.xlsx") -> str:
    if not name:
        return fallback
    cleaned = Path(name).name.strip()
    return cleaned or fallback


def _validate_excel(filename: str | None, content: bytes) -> None:
    if not filename:
        raise ValueError("Ожидается файл с расширением .xlsx или .xls")
    lower = filename.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xls")):
        raise ValueError("Ожидается файл с расширением .xlsx или .xls")
    if len(content) > MAX_FILE_SIZE:
        raise ValueError("Файл слишком большой (максимум 20 МБ)")
    if lower.endswith(".xlsx"):
        if not content.startswith(b"PK"):
            raise ValueError("Файл не является корректным .xlsx")
    elif lower.endswith(".xls"):
        if not content.startswith(b"\xd0\xcf\x11\xe0"):
            raise ValueError("Файл не является корректным .xls")


def _read_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"filename": None, "updated_at": None}


def load_template_bytes() -> bytes:
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            return f.read()
    raise FileNotFoundError("Шаблон не найден: загрузите эталонный файл")


def store_template(filename: str, content: bytes) -> dict:
    """Save the shared reference template used by the site and the Telegram bot."""
    if not filename or not str(filename).lower().endswith(".xlsx"):
        raise ValueError("Эталонный шаблон должен быть в формате .xlsx")
    if not content.startswith(b"PK"):
        raise ValueError("Файл не является корректным .xlsx")
    if len(content) > MAX_FILE_SIZE:
        raise ValueError("Файл слишком большой (максимум 20 МБ)")
    safe = _safe_filename(filename)
    with open(TEMPLATE_PATH, "wb") as f:
        f.write(content)
    meta = {
        "filename": safe,
        "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    try:
        _save_template_blob(content, meta["filename"], meta["updated_at"])
    except Exception:
        traceback.print_exc()
    return meta



@app.get("/api/health")
def health():
    return {
        "ok": True,
        "version": APP_VERSION,
        "data_dir": DATA_DIR,
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "persistent_data": os.path.normpath(DATA_DIR) != os.path.normpath(BUNDLE_DATA_DIR),
        "telegram": bool(os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()),
    }


def _article_body_filters(body: dict) -> dict:
    filters = body.get("filters") or {}
    if not isinstance(filters, dict):
        raise ValueError("Фильтр должен быть объектом")
    return filters


@app.get("/api/articles")
def list_articles(
    q: str = "",
    page: int = Query(1, ge=1),
    per_page: int = Query(40, ge=1, le=200),
):
    return articles.search(q, page, per_page)


@app.post("/api/articles/query")
async def query_articles(request: Request):
    body = await request.json()
    if not isinstance(body, dict):
        return JSONResponse({"error": "Некорректный запрос"}, status_code=422)
    try:
        return articles.search(
            str(body.get("q") or ""),
            int(body.get("page") or 1),
            int(body.get("per_page") or 40),
            _article_body_filters(body),
        )
    except (ValueError, TypeError) as e:
        return JSONResponse({"error": str(e)}, status_code=422)


@app.post("/api/articles/values")
async def article_values(request: Request):
    body = await request.json()
    if not isinstance(body, dict):
        return JSONResponse({"error": "Некорректный запрос"}, status_code=422)
    try:
        values = articles.distinct_values(
            str(body.get("field") or ""),
            str(body.get("q") or ""),
            _article_body_filters(body),
        )
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    return {"values": values}


@app.get("/api/articles/{article_id}")
def get_article(article_id: int):
    row = articles.get_article(article_id)
    if row is None:
        return JSONResponse({"error": "Артикул не найден"}, status_code=404)
    return row


@app.post("/api/articles")
async def create_article(request: Request):
    try:
        return articles.create_article(await request.json())
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)


@app.put("/api/articles/{article_id}")
async def update_article(article_id: int, request: Request):
    try:
        return articles.update_article(article_id, await request.json())
    except LookupError as e:
        return JSONResponse({"error": str(e)}, status_code=404)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)


# ── AI Excel column mapping (YandexGPT) ───────────────────────────────────

@app.get("/api/ai/status")
def ai_status():
    return {"configured": yandex_gpt.is_configured()}


@app.post("/api/ai/analyze")
async def ai_analyze(
    inputs: list[UploadFile] = File(...),
    template: UploadFile | None = File(None),
    country: str = Form("CN"),
    unit: str = Form("шт"),
):
    """YandexGPT сопоставляет столбцы по смыслу и заполняет шаблон."""
    try:
        template_bytes = await _resolve_template(template)
        input_bytes, input_names = await _read_all(inputs)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    except FileNotFoundError as e:
        return JSONResponse({"error": str(e)}, status_code=422)

    settings = {"country": country, "unit": unit}

    if not yandex_gpt.is_configured():
        return JSONResponse(
            {
                "error": (
                    "YandexGPT не настроен. "
                    "На Render: Dashboard → tamozhenformat → Environment → "
                    "добавьте YANDEX_API_KEY и YANDEX_FOLDER_ID (из локального .env), затем Redeploy. "
                    "Локально: задайте те же переменные в файле .env рядом с main.py."
                )
            },
            status_code=503,
        )

    try:
        result_bytes, report = await run_in_threadpool(
            ai_mapper.process_with_ai,
            template_bytes,
            input_bytes,
            settings,
            input_names,
        )
    except Exception as e:
        traceback.print_exc()
        # Fallback to keyword engine
        try:
            result_bytes, report = await run_in_threadpool(
                engine.process, template_bytes, input_bytes, settings
            )
            report = {**report, "mode": "keyword_fallback", "ai_error": str(e)}
            proto = dict(report.get("protocol") or {})
            warns = list(proto.get("warnings") or [])
            warns.insert(0, "ИИ недоступен или не ответил — файл собран по заголовкам.")
            proto["warnings"] = warns
            summary = proto.get("summary") or ""
            if summary and "замечан" not in summary and "ошибк" not in summary:
                proto["summary"] = summary.rstrip(".") + ". Есть замечания."
            report["protocol"] = proto
        except Exception as e2:
            traceback.print_exc()
            return JSONResponse(
                {"error": f"{_public_error(e2)} (после ошибки ИИ: {_public_error(e)})"},
                status_code=500,
            )

    try:
        return _xlsx_result_response(result_bytes, report)
    except Exception as e3:
        traceback.print_exc()
        return JSONResponse({"error": _public_error(e3)}, status_code=500)


@app.get("/api/template")
def get_template_info():
    meta = _read_meta()
    exists = os.path.exists(TEMPLATE_PATH)
    return {"exists": exists, **meta}


@app.get("/api/template/download")
def download_template():
    if not os.path.exists(TEMPLATE_PATH):
        return JSONResponse({"error": "no template stored"}, status_code=404)
    meta = _read_meta()
    name = _safe_filename(meta.get("filename"), "template.xlsx")
    return FileResponse(TEMPLATE_PATH, filename=name)


@app.post("/api/template")
async def upload_template(file: UploadFile = File(...)):
    content = await file.read()
    try:
        meta = store_template(file.filename or "", content)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    return {"exists": True, **meta}


async def _read_all(files: list[UploadFile]) -> tuple[list[bytes], list[str]]:
    if len(files) > MAX_FILES:
        raise ValueError(f"Слишком много файлов (максимум {MAX_FILES})")
    out: list[bytes] = []
    names: list[str] = []
    for f in files:
        content = await f.read()
        _validate_excel(f.filename, content)
        out.append(content)
        names.append(f.filename or "input.xlsx")
    return out, names


async def _resolve_template(template: UploadFile | None) -> bytes:
    if template is not None:
        content = await template.read()
        if content:
            _validate_excel(template.filename, content)
            return content
    return load_template_bytes()


def _xlsx_result_response(result_bytes: bytes, report: dict) -> JSONResponse:
    protocol = report.get("protocol") or {}
    return JSONResponse(
        {
            "filename": "result.xlsx",
            "items_found": report.get("items_found", 0),
            "mode": report.get("mode") or "",
            "protocol": protocol,
            "file_b64": base64.b64encode(result_bytes).decode("ascii"),
        }
    )


@app.post("/api/preview")
async def preview(
    inputs: list[UploadFile] = File(...),
    template: UploadFile | None = File(None),
    country: str = Form("CN"),
    unit: str = Form("шт"),
):
    try:
        template_bytes = await _resolve_template(template)
        input_bytes, _names = await _read_all(inputs)
        settings = {"country": country, "unit": unit}
        _, report = await run_in_threadpool(engine.process, template_bytes, input_bytes, settings)
        return JSONResponse(report)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": _public_error(e)}, status_code=500)


@app.post("/api/process")
async def process_files(
    inputs: list[UploadFile] = File(...),
    template: UploadFile | None = File(None),
    country: str = Form("CN"),
    unit: str = Form("шт"),
):
    try:
        template_bytes = await _resolve_template(template)
        input_bytes, _names = await _read_all(inputs)
        settings = {"country": country, "unit": unit}
        result_bytes, report = await run_in_threadpool(
            engine.process, template_bytes, input_bytes, settings
        )
        return _xlsx_result_response(result_bytes, report)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": _public_error(e)}, status_code=500)


import telegram_bot

telegram_bot.mount(app)

FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("APP_HOST", "127.0.0.1")
    port = int(os.environ.get("APP_PORT", "8756"))
    uvicorn.run(app, host=host, port=port)
