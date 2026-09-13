#!/usr/bin/env python3
"""FastAPI backend for the web build of ТаможенФормат.

Serves both the API and the static frontend from a single process.

Endpoints:
  GET  /api/health             -> {"ok": true}
  GET  /api/template           -> info about the currently stored template
  GET  /api/template/download  -> download the currently stored template
  POST /api/template           -> upload/replace the stored template
  POST /api/preview            -> JSON report of what would be filled in
  POST /api/process            -> filled .xlsx as a download
"""
import io
import os
import sys
import json
import shutil
import traceback
import datetime
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
import engine

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
TEMPLATE_PATH = os.path.join(DATA_DIR, "current_template.xlsx")
DEFAULT_TEMPLATE_PATH = os.path.join(DATA_DIR, "default_template.xlsx")
META_PATH = os.path.join(DATA_DIR, "template_meta.json")

if not os.path.exists(TEMPLATE_PATH) and os.path.exists(DEFAULT_TEMPLATE_PATH):
    shutil.copy(DEFAULT_TEMPLATE_PATH, TEMPLATE_PATH)
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "filename": "etalonnyi-shablon.xlsx",
                "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            },
            f,
            ensure_ascii=False,
        )

app = FastAPI(title="ТаможенФормат")

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB per file
MAX_FILES = 30


def _safe_filename(name: str | None, fallback: str = "template.xlsx") -> str:
    if not name:
        return fallback
    cleaned = Path(name).name.strip()
    return cleaned or fallback


def _validate_xlsx(filename: str | None, content: bytes) -> None:
    if not filename or not filename.lower().endswith(".xlsx"):
        raise ValueError("Ожидается файл с расширением .xlsx")
    if len(content) > MAX_FILE_SIZE:
        raise ValueError("Файл слишком большой (максимум 20 МБ)")
    if not content.startswith(b"PK"):
        raise ValueError("Файл не является корректным .xlsx")


def _read_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"filename": None, "updated_at": None}


@app.get("/api/health")
def health():
    return {"ok": True}


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
        _validate_xlsx(file.filename, content)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    with open(TEMPLATE_PATH, "wb") as f:
        f.write(content)
    meta = {
        "filename": _safe_filename(file.filename),
        "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    return {"exists": True, **meta}


async def _read_all(files: list[UploadFile]) -> list[bytes]:
    if len(files) > MAX_FILES:
        raise ValueError(f"Слишком много файлов (максимум {MAX_FILES})")
    out = []
    for f in files:
        content = await f.read()
        _validate_xlsx(f.filename, content)
        out.append(content)
    return out


async def _resolve_template(template: UploadFile | None) -> bytes:
    if template is not None:
        content = await template.read()
        if content:
            _validate_xlsx(template.filename, content)
            return content
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            return f.read()
    raise FileNotFoundError("Шаблон не найден: загрузите эталонный файл")


@app.post("/api/preview")
async def preview(
    inputs: list[UploadFile] = File(...),
    template: UploadFile | None = File(None),
    country: str = Form("CN"),
    unit: str = Form("шт"),
):
    try:
        template_bytes = await _resolve_template(template)
        input_bytes = await _read_all(inputs)
        settings = {"country": country, "unit": unit}
        _, report = await run_in_threadpool(engine.process, template_bytes, input_bytes, settings)
        return JSONResponse(report)
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)


@app.post("/api/process")
async def process_files(
    inputs: list[UploadFile] = File(...),
    template: UploadFile | None = File(None),
    country: str = Form("CN"),
    unit: str = Form("шт"),
):
    try:
        template_bytes = await _resolve_template(template)
        input_bytes = await _read_all(inputs)
        settings = {"country": country, "unit": unit}
        result_bytes, report = await run_in_threadpool(
            engine.process, template_bytes, input_bytes, settings
        )
        headers = {
            "Content-Disposition": 'attachment; filename="result.xlsx"',
            "X-Items-Found": str(report["items_found"]),
        }
        return StreamingResponse(
            io.BytesIO(result_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)


FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("APP_HOST", "127.0.0.1")
    port = int(os.environ.get("APP_PORT", "8756"))
    uvicorn.run(app, host=host, port=port)
