#!/usr/bin/env python3
"""FastAPI backend for the web build of ТаможенФормат.

Serves both the API and the static frontend from a single process.
"""
import io
import os
import sys
import json
import base64
import shutil
import re
import traceback
import datetime
import time
import sqlite3
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from contextlib import contextmanager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
import engine
import ai_mapper
import yandex_gpt

from fastapi import FastAPI, UploadFile, File, Form, Query, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

# Bundled defaults live next to the code (shipped in the image / repo).
BUNDLE_DATA_DIR = os.path.join(BASE_DIR, "data")
DEFAULT_TEMPLATE_PATH = os.path.join(BUNDLE_DATA_DIR, "default_template.xlsx")

# Mutable runtime data (template upload, SVH cache). On Render point DATA_DIR
# at a persistent disk mount so uploads survive restarts/redeploys.
DATA_DIR = os.environ.get("DATA_DIR") or BUNDLE_DATA_DIR
os.makedirs(DATA_DIR, exist_ok=True)
TEMPLATE_PATH = os.path.join(DATA_DIR, "current_template.xlsx")
META_PATH = os.path.join(DATA_DIR, "template_meta.json")
SVH_DB_PATH = os.path.join(DATA_DIR, "svh_cache.db")
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
APP_VERSION = "2026-09-21.3"


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
    """Prevent browsers from keeping a stale app.js that breaks SVH on the public URL."""
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
ALTA_SVH_SEARCH_URL = "https://www.alta.ru/svh/search/"
ALTA_SVH_DETAIL_URL = "https://www.alta.ru/svh/"

# ── SQLite cache ──────────────────────────────────────────────────────────

def _get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(SVH_DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def _init_db():
    with _get_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS svh_search (
                query_key   TEXT PRIMARY KEY,
                results     TEXT NOT NULL,
                fetched_at  REAL NOT NULL
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS svh_detail (
                license_key TEXT PRIMARY KEY,
                data        TEXT NOT NULL,
                fetched_at  REAL NOT NULL
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS svh_page (
                page_num    INTEGER PRIMARY KEY,
                html        BLOB NOT NULL,
                fetched_at  REAL NOT NULL
            )
        """)
        db.commit()

_init_db()

@contextmanager
def _db():
    conn = _get_db()
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()

SVH_SEARCH_CACHE_TTL = 6 * 3600  # 6 hours — fewer live hits to Alta from Render
SVH_DETAIL_CACHE_TTL = 12 * 3600
SVH_HTTP_CONNECT_TIMEOUT = 8
SVH_HTTP_READ_TIMEOUT = 18
SVH_HTTP_RETRIES = 2
SVH_HTTP_BUDGET = 28  # keep under Render request limits

def _cache_get_search(query_key: str, *, allow_stale: bool = False) -> list[dict] | None:
    with _db() as db:
        row = db.execute("SELECT results, fetched_at FROM svh_search WHERE query_key = ?", (query_key,)).fetchone()
    if not row:
        return None
    age = time.time() - row["fetched_at"]
    if allow_stale or age < SVH_SEARCH_CACHE_TTL:
        return json.loads(row["results"])
    return None

def _cache_set_search(query_key: str, results: list[dict]):
    with _db() as db:
        db.execute(
            "INSERT OR REPLACE INTO svh_search (query_key, results, fetched_at) VALUES (?, ?, ?)",
            (query_key, json.dumps(results, ensure_ascii=False), time.time()),
        )

def _cache_get_detail(license_key: str, *, allow_stale: bool = False) -> dict | None:
    with _db() as db:
        row = db.execute("SELECT data, fetched_at FROM svh_detail WHERE license_key = ?", (license_key,)).fetchone()
    if not row:
        return None
    age = time.time() - row["fetched_at"]
    if allow_stale or age < SVH_DETAIL_CACHE_TTL:
        return json.loads(row["data"])
    return None

def _cache_set_detail(license_key: str, data: dict):
    with _db() as db:
        db.execute(
            "INSERT OR REPLACE INTO svh_detail (license_key, data, fetched_at) VALUES (?, ?, ?)",
            (license_key, json.dumps(data, ensure_ascii=False), time.time()),
        )

def _cache_clear():
    with _db() as db:
        db.execute("DELETE FROM svh_search")
        db.execute("DELETE FROM svh_detail")
        db.execute("DELETE FROM svh_page")

def _cache_stats() -> dict:
    with _db() as db:
        s = db.execute("SELECT COUNT(*) AS c FROM svh_search").fetchone()["c"]
        d = db.execute("SELECT COUNT(*) AS c FROM svh_detail").fetchone()["c"]
        p = db.execute("SELECT COUNT(*) AS c FROM svh_page").fetchone()["c"]
    return {"search_entries": s, "detail_entries": d, "page_entries": p, "total": s + d + p}

# ── In-memory short cache (fallback for fast repeated queries) ──
_svh_mem_cache: dict[str, tuple[float, list[dict]]] = {}


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


class _AltaSvhParser(HTMLParser):
    """Extract result cards from Alta's public SVH/TS search page."""

    def __init__(self):
        super().__init__()
        self.cards: list[dict] = []
        self._in_results_col = False
        self._results_col_depth = 0
        self._in_card = False
        self._card_depth = 0
        self._card: dict | None = None
        self._collecting = False
        self._collect_target: str | None = None
        self._collect_buf: list[str] = []
        self._in_phone = False
        self._in_email = False
        self._address_col_seen = False
        self._right_col_seen = False

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        classes = attributes.get("class", "").split()

        # Detect results column: col-75
        if not self._in_results_col and "col-75" in classes:
            self._in_results_col = True
            self._results_col_depth = 1
            return
        if self._in_results_col:
            self._results_col_depth += 1

        # Only parse cards inside results column
        if not self._in_results_col:
            return

        # Detect card container: boxSubstrate boxSubstrate-offset-0 p-10 mb10
        if not self._in_card and "boxSubstrate" in classes and "boxSubstrate-offset-0" in classes and "p-10" in classes:
            self._in_card = True
            self._card = {
                "type": "", "name": "", "url": "", "address": "",
                "license": "", "customs": "", "transport": "",
                "phone": "", "email": "", "inn": "", "owner_address": "",
            }
            self._card_depth = 1
            self._address_col_seen = False
            self._right_col_seen = False
            return

        if not self._in_card:
            return

        self._card_depth += 1

        # Link to detail page
        if tag == "a" and "/svh/" in attributes.get("href", "") and not self._card["url"]:
            href = attributes["href"]
            if href.startswith("/"):
                href = "https://www.alta.ru" + href
            self._card["url"] = href

        # h3 contains the name
        if tag == "h3" and not self._card["name"]:
            self._collect_target = "name"
            self._collect_buf = []
            self._collecting = True

        # Type is the text before h3 (first text in card)
        if not self._card["type"]:
            self._collect_target = "type"
            self._collect_buf = []
            self._collecting = True

        # Address: pSvh_fieldColumn-list pSvh_fieldColumn-left lightgray
        if "pSvh_fieldColumn-list" in classes and "pSvh_fieldColumn-left" in classes and "lightgray" in classes:
            self._collect_target = "address"
            self._collect_buf = []
            self._collecting = True
            self._address_col_seen = True

        # License: div inside pSvh_fieldColumn-right
        if "pSvh_fieldColumn-right" in classes and tag == "div":
            if not self._address_col_seen:
                # First right column = license (before customs row)
                self._collect_target = "license"
                self._collect_buf = []
                self._collecting = True
            else:
                self._right_col_seen = True
                # Second right column = transport
                self._collect_target = "transport"
                self._collect_buf = []
                self._collecting = True

        # Customs: pSvh_fieldColumn-left (second occurrence after address)
        if "pSvh_fieldColumn-list" in classes and "pSvh_fieldColumn-left" in classes and not "lightgray" in classes:
            if self._address_col_seen:
                self._collect_target = "customs"
                self._collect_buf = []
                self._collecting = True

        # Phone link
        if tag == "a" and attributes.get("href", "").startswith("tel:"):
            self._in_phone = True
            self._collect_buf = []
        # Email link
        if tag == "a" and attributes.get("href", "").startswith("mailto:"):
            self._in_email = True
            self._collect_buf = []

    def handle_endtag(self, tag):
        # Close results column
        if not self._in_results_col:
            return
        self._results_col_depth -= 1
        if self._results_col_depth <= 0:
            self._in_results_col = False

        # Close card
        if not self._in_card:
            return

        # Close phone / email
        if self._in_phone and tag == "a":
            val = "".join(self._collect_buf).strip()
            if val:
                self._card["phone"] = val
            self._in_phone = False
            self._collect_buf = []
        if self._in_email and tag == "a":
            val = "".join(self._collect_buf).strip()
            if val:
                self._card["email"] = val
            self._in_email = False
            self._collect_buf = []

        # Close collection
        if self._collecting and tag in ("br", "div", "a", "span", "p", "h3", "strong", "em"):
            if self._collect_target and self._card:
                val = " ".join("".join(self._collect_buf).split())
                if val and not self._card.get(self._collect_target):
                    self._card[self._collect_target] = val
            self._collecting = False
            self._collect_target = None
            self._collect_buf = []

        self._card_depth -= 1
        if self._card_depth <= 0:
            if self._card and self._card.get("name"):
                self.cards.append(self._card)
            self._in_card = False
            self._card = None

    def handle_data(self, data):
        if not self._in_results_col:
            return

        if self._in_phone or self._in_email:
            self._collect_buf.append(data)
            return

        if self._collecting and self._collect_target:
            self._collect_buf.append(data)


# ── Detail-page parser ────────────────────────────────────────────────────

class _AltaSvhDetailParser(HTMLParser):
    """Extract detailed card from a single SVH/TS page (e.g. /svh/svh-10001070223200072/)."""

    def __init__(self):
        super().__init__()
        self.title = ""
        self.in_title = False
        self.title_depth = 0
        self.current_label = None
        self.current_data: list[str] = []
        self.result: dict = {
            "name": "",
            "address": "",
            "owner_address": "",
            "phone": "",
            "email": "",
            "inn": "",
            "license": "",
            "type": "",
            "customs": "",
            "transport": "",
        }
        self.label_pattern = re.compile(r"^(Склад временного хранения расположен по адресу:|Адрес владельца:|Телефон:|E-mail:|ИНН:|№ лицензии:|Тип склада:|Таможня:)\s*$", re.UNICODE)
        self._current_key = None

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        classes = attributes.get("class", "").split()

        if tag == "h1" and not self.in_title:
            self.in_title = True
            self.title_depth = 1

        # Phone link
        if tag == "a" and attributes.get("href", "").startswith("tel:"):
            self._current_key = "phone"
            self.current_data = []
        # Email link
        if tag == "a" and attributes.get("href", "").startswith("mailto:"):
            self._current_key = "email"
            self.current_data = []

    def handle_endtag(self, tag):
        if self.in_title and tag == "h1":
            self.in_title = False
            self.title = " ".join(self.current_data).strip()
            self.current_data = []

        if self._current_key and tag == "a":
            val = "".join(self.current_data).strip()
            if val:
                self.result[self._current_key] = val
            self._current_key = None
            self.current_data = []

    def handle_data(self, data):
        if self.in_title:
            self.current_data.append(data)

        if self._current_key:
            self.current_data.append(data)

        # Detect label text (bold / strong before colon)
        stripped = data.strip()
        if not stripped:
            return
        # Check if this data matches a known label pattern
        for label, key in [
            ("Склад временного хранения расположен по адресу:", "address"),
            ("Адрес владельца:", "owner_address"),
            ("Телефон:", "phone"),
            ("E-mail:", "email"),
            ("ИНН:", "inn"),
            ("№ лицензии:", "license"),
            ("Тип склада:", "type"),
            ("Таможня:", "customs"),
        ]:
            if stripped.startswith(label):
                self._current_key = key
                self.current_data = [stripped[len(label):].strip()]
                return


def _alta_headers(purpose: str) -> dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        # Avoid brotli/odd encodings; identity is safest for urllib on Render.
        "Accept-Encoding": "identity",
        "Connection": "close",
        "Cache-Control": "no-cache",
        "Referer": "https://www.alta.ru/svh/",
        "X-Purpose": purpose,
    }


def _http_get_text(url: str, purpose: str) -> str:
    """GET Alta.ru HTML with retries; raises the last error on total failure."""
    last_err: Exception | None = None
    started = time.time()
    for attempt in range(1, SVH_HTTP_RETRIES + 1):
        remaining = SVH_HTTP_BUDGET - (time.time() - started)
        if remaining < 3:
            break
        # urllib timeout must be a single number (not a connect/read tuple).
        timeout = min(
            SVH_HTTP_CONNECT_TIMEOUT + SVH_HTTP_READ_TIMEOUT,
            max(5.0, remaining - 1),
        )
        try:
            request = Request(url, headers=_alta_headers(purpose))
            with urlopen(request, timeout=timeout) as response:
                raw = response.read()
                ctype = (response.headers.get("Content-Type") or "").lower()
                charset = "utf-8"
                if "charset=" in ctype:
                    charset = ctype.split("charset=", 1)[1].split(";")[0].strip() or "utf-8"
                try:
                    return raw.decode(charset, errors="replace")
                except LookupError:
                    return raw.decode("utf-8", errors="replace")
        except Exception as e:
            last_err = e
            print(f"[SVH] {purpose} attempt {attempt}/{SVH_HTTP_RETRIES} failed: {e}")
            if attempt < SVH_HTTP_RETRIES and (time.time() - started) < SVH_HTTP_BUDGET - 3:
                time.sleep(min(1.2 * attempt, 2.0))
    assert last_err is not None
    raise last_err


def _fetch_svh(query: dict[str, str]) -> list[dict]:
    query_string = urlencode({key: value for key, value in query.items() if value}, encoding="utf-8")
    # Check in-memory cache first
    mem = _svh_mem_cache.get(query_string)
    if mem and time.time() - mem[0] < 120:
        return mem[1]
    # Check SQLite cache
    cached = _cache_get_search(query_string)
    if cached is not None:
        _svh_mem_cache[query_string] = (time.time(), cached)
        return cached
    try:
        html = _http_get_text(f"{ALTA_SVH_SEARCH_URL}?{query_string}", "SVH search")
    except Exception:
        stale = _cache_get_search(query_string, allow_stale=True)
        if stale is not None:
            print(f"[SVH] serving stale cache for query={query_string[:60]}")
            _svh_mem_cache[query_string] = (time.time(), stale)
            return stale
        raise

    # Detect soft-block / challenge pages from datacenter IPs
    low = html.lower()
    blocked = any(
        m in low
        for m in (
            "access denied",
            "just a moment",
            "cf-browser-verification",
            "ddos-guard",
            "checking your browser",
        )
    )
    results = _parse_svh_cards_regex(html)

    box_count = len(re.findall(r"boxSubstrate boxSubstrate-offset-0", html))
    print(
        f"[SVH] DEBUG: query={query_string[:60]}, cards={len(results)}, "
        f"boxSubstrate-offset-0={box_count}, blocked={blocked}, html_len={len(html)}"
    )

    if blocked and not results:
        stale = _cache_get_search(query_string, allow_stale=True)
        if stale is not None:
            return stale
        raise RuntimeError("Alta.ru временно блокирует запросы с сервера. Повторите через минуту.")

    # If the page clearly has cards but parsing failed — try not to poison cache
    if not results and box_count > 0:
        results = _parse_svh_cards_regex(html.replace("p-10 ", "").replace(" p-10", ""))

    if results:
        _cache_set_search(query_string, results)
        _svh_mem_cache[query_string] = (time.time(), results)
    else:
        # Do not cache empty miss for long — may be a transient Alta glitch
        _svh_mem_cache[query_string] = (time.time(), results)
    return results


def _parse_svh_cards_regex(html: str) -> list[dict]:
    """Parse SVH cards from Alta.ru search results using div-counting."""
    cards = []

    # Find the results column (col-75) to limit search scope
    col75_start = html.find('class="col-75')
    if col75_start == -1:
        search_html = html
    else:
        # Find the next major section after results
        next_footer = html.find('<footer', col75_start)
        next_main = html.find('</main>', col75_start)
        end_pos = min(
            [p for p in [next_footer, next_main] if p != -1]
            or [len(html)]
        )
        search_html = html[col75_start:end_pos]

    # Find all card opening tags (Alta class list varies: with/without p-10 etc.)
    card_open_pattern = re.compile(
        r'class="boxSubstrate boxSubstrate-offset-0[^"]*"[^>]*>'
    )
    for card_open in card_open_pattern.finditer(search_html):
        start = card_open.end()

        # Count divs to find the matching closing tag
        depth = 1
        pos = start
        card_close = -1
        while depth > 0 and pos < len(search_html):
            next_open = search_html.find('<div', pos)
            next_close = search_html.find('</div>', pos)
            if next_close == -1:
                break
            if next_open != -1 and next_open < next_close:
                # Check it's not a closing tag
                tag_text = search_html[next_open:next_open+6]
                if not tag_text.startswith('</d'):
                    depth += 1
                    pos = next_open + 4
                else:
                    depth -= 1
                    if depth == 0:
                        card_close = next_close
                        break
            else:
                depth -= 1
                if depth == 0:
                    card_close = next_close
                    break
                pos = next_close + 6

        if card_close == -1:
            continue

        card_html = search_html[start:card_close].strip()
        if not card_html:
            continue

        card = {
            "type": "", "name": "", "url": "", "address": "",
            "license": "", "customs": "", "transport": "",
            "phone": "", "email": "", "inn": "", "owner_address": "",
        }

        # Type: first text before <br>
        type_match = re.search(r'^(.*?)(?:<br|<div)', card_html, re.DOTALL)
        if type_match:
            card["type"] = re.sub(r'<[^>]+>', '', type_match.group(1)).strip()
            card["type"] = re.sub(r'\s+', ' ', card["type"])

        # Name + URL: inside <div class="h3"><a href="...">Name</a></div>
        name_match = re.search(
            r'class="h3">.*?<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>',
            card_html, re.DOTALL
        )
        if name_match:
            card["url"] = "https://www.alta.ru" + name_match.group(1)
            card["name"] = re.sub(r'<[^>]+>', '', name_match.group(2)).strip()

        # Address: pSvh_fieldColumn-left lightgray
        addr_match = re.search(
            r'pSvh_fieldColumn-left lightgray">(.*?)</div>',
            card_html, re.DOTALL
        )
        if addr_match:
            card["address"] = re.sub(r'<[^>]+>', '', addr_match.group(1)).strip()

        # License: first pSvh_fieldColumn-right div content
        lic_match = re.search(
            r'pSvh_fieldColumn-right">(.*?)</div>',
            card_html, re.DOTALL
        )
        if lic_match:
            card["license"] = re.sub(r'<[^>]+>', '', lic_match.group(1)).strip()

        # Customs: second pSvh_fieldColumn-left div (without lightgray)
        all_left = re.findall(r'pSvh_fieldColumn-left[^>]*">(.*?)</div>', card_html, re.DOTALL)
        if len(all_left) >= 2:
            card["customs"] = re.sub(r'<[^>]+>', '', all_left[1]).strip()

        # Transport: second pSvh_fieldColumn-right div
        all_right = re.findall(r'pSvh_fieldColumn-right">(.*?)</div>', card_html, re.DOTALL)
        if len(all_right) >= 2:
            card["transport"] = re.sub(r'<[^>]+>', '', all_right[1]).strip()

        if card["name"]:
            cards.append(card)

    return cards


def _fetch_svh_detail(license_url: str) -> dict | None:
    # Extract license key from URL like https://www.alta.ru/svh/svh-10001070223200072/
    m = re.search(r"/svh/([\w\-]+)", license_url)
    if not m:
        return None
    license_key = m.group(1)
    # Check cache
    cached = _cache_get_detail(license_key)
    if cached is not None:
        return cached
    url = license_url if license_url.startswith("http") else ALTA_SVH_DETAIL_URL + license_key + "/"
    try:
        html = _http_get_text(url, "SVH detail")
    except Exception:
        stale = _cache_get_detail(license_key, allow_stale=True)
        if stale is not None:
            print(f"[SVH] serving stale detail cache for {license_key}")
            return stale
        raise
    parser = _AltaSvhDetailParser()
    parser.feed(html)
    result = parser.result
    if result.get("name"):
        result["url"] = url
        _cache_set_detail(license_key, result)
    return result if result.get("name") else None


def _fetch_all_pages(max_pages: int = 82) -> list[dict]:
    """Fetch all pages of the SVH registry and return merged list of cards."""
    all_cards: list[dict] = []
    seen_urls: set[str] = set()
    for page in range(1, max_pages + 1):
        cache_key = f"_page_{page}"
        cached_page = _cache_get_search(cache_key)
        if cached_page:
            all_cards.extend(cached_page)
            continue
        url = f"{ALTA_SVH_SEARCH_URL}" if page == 1 else f"{ALTA_SVH_SEARCH_URL}page_{page}/"
        try:
            html = _http_get_text(url, "SVH full registry")
            parser = _AltaSvhParser()
            parser.feed(html)
            page_cards = parser.cards
            _cache_set_search(cache_key, page_cards)
            for card in page_cards:
                if card.get("url") and card["url"] not in seen_urls:
                    seen_urls.add(card["url"])
                    all_cards.append(card)
        except Exception:
            break  # Stop at first error (no more pages)
    return all_cards


@app.get("/api/health")
def health():
    return {
        "ok": True,
        "version": APP_VERSION,
        "data_dir": DATA_DIR,
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "persistent_data": os.path.normpath(DATA_DIR) != os.path.normpath(BUNDLE_DATA_DIR),
    }


@app.get("/api/svh/search")
async def search_svh(
    s_adres: str = "",
    s_name: str = "",
    s_tam: str = "",
    s_tam_name: str = "",
    s_nlic: str = "",
    s_vidtrans: str = "",
):
    query = {
        "s_adres": s_adres.strip(),
        "s_name": s_name.strip(),
        "s_tam": s_tam.strip(),
        "s_tam_name": s_tam_name.strip(),
        "s_nlic": s_nlic.strip(),
        "s_vidtrans": s_vidtrans.strip(),
    }
    if not any(query.values()):
        return JSONResponse({"error": "Укажите хотя бы один параметр поиска"}, status_code=422)
    try:
        results = await run_in_threadpool(_fetch_svh, query)
        return {
            "source": "Alta.ru",
            "source_url": ALTA_SVH_SEARCH_URL,
            "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "count": len(results),
            "results": results,
        }
    except Exception as e:
        traceback.print_exc()
        stale_note = ""
        try:
            # Last-ditch: any stale cache for this exact query
            qs = urlencode({k: v for k, v in query.items() if v}, encoding="utf-8")
            stale = _cache_get_search(qs, allow_stale=True)
            if stale:
                return {
                    "source": "Alta.ru (кэш)",
                    "source_url": ALTA_SVH_SEARCH_URL,
                    "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                    "count": len(stale),
                    "results": stale,
                    "warning": f"Живой запрос к Alta.ru не удался ({e}); показан сохранённый результат.",
                }
        except Exception:
            pass
        return JSONResponse(
            {
                "error": (
                    "Не удалось получить данные Alta.ru. "
                    "Сервер не достучался до реестра вовремя — повторите через 15–30 секунд. "
                    f"({e})"
                )
            },
            status_code=502,
        )


@app.get("/api/svh/detail/{license_key}")
async def detail_svh(license_key: str):
    """Get detailed information about a specific SVH/TS by license key."""
    try:
        result = await run_in_threadpool(_fetch_svh_detail, license_key)
        if result:
            return {"source": "Alta.ru", "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(), "result": result}
        return JSONResponse({"error": f"СВХ с ключом '{license_key}' не найден"}, status_code=404)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": f"Ошибка получения данных: {e}"}, status_code=502)


@app.get("/api/svh/all")
async def all_svh(page: int = Query(1, ge=1), per_page: int = Query(100, ge=1, le=500)):
    """Get the full SVH registry with pagination. Fetches from Alta.ru if cache is empty."""
    try:
        all_cards = await run_in_threadpool(_fetch_all_pages)
        total = len(all_cards)
        start = (page - 1) * per_page
        end = start + per_page
        return {
            "source": "Alta.ru",
            "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "total": total,
            "page": page,
            "per_page": per_page,
            "total_pages": (total + per_page - 1) // per_page if total else 0,
            "results": all_cards[start:end],
        }
    except Exception as e:
        traceback.print_exc()
        return JSONResponse({"error": f"Ошибка получения реестра: {e}"}, status_code=502)


@app.get("/api/svh/cache/stats")
async def svh_cache_stats():
    """Get cache statistics."""
    return _cache_stats()


@app.delete("/api/svh/cache")
async def clear_svh_cache():
    """Clear all SVH cache data."""
    _cache_clear()
    _svh_mem_cache.clear()
    return {"ok": True, "message": "Кэш очищен"}


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
            return JSONResponse({"error": str(e2) or e2.__class__.__name__}, status_code=500)

    try:
        return _xlsx_result_response(result_bytes, report)
    except Exception as e3:
        traceback.print_exc()
        return JSONResponse({"error": str(e3) or e3.__class__.__name__}, status_code=500)


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
        if not file.filename or not file.filename.lower().endswith(".xlsx"):
            raise ValueError("Эталонный шаблон должен быть в формате .xlsx")
        if not content.startswith(b"PK"):
            raise ValueError("Файл не является корректным .xlsx")
        if len(content) > MAX_FILE_SIZE:
            raise ValueError("Файл слишком большой (максимум 20 МБ)")
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
    try:
        _save_template_blob(content, meta["filename"], meta["updated_at"])
    except Exception:
        traceback.print_exc()
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
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            return f.read()
    raise FileNotFoundError("Шаблон не найден: загрузите эталонный файл")


def _xlsx_result_response(result_bytes: bytes, report: dict) -> StreamingResponse:
    headers = {
        "Content-Disposition": 'attachment; filename="result.xlsx"',
        "X-Items-Found": str(report.get("items_found", 0)),
        "X-AI-Mode": str(report.get("mode") or ""),
        "Access-Control-Expose-Headers": "X-Items-Found, X-AI-Mode, X-Protocol",
    }
    protocol = report.get("protocol")
    if protocol:
        try:
            encoded = base64.b64encode(
                json.dumps(protocol, ensure_ascii=False).encode("utf-8")
            ).decode("ascii")
            # Proxies reject oversized response headers; keep protocol short.
            if len(encoded) <= 3500:
                headers["X-Protocol"] = encoded
        except Exception:
            traceback.print_exc()
    return StreamingResponse(
        io.BytesIO(result_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
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
        return JSONResponse({"error": str(e)}, status_code=500)


FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("APP_HOST", "127.0.0.1")
    port = int(os.environ.get("APP_PORT", "8756"))
    uvicorn.run(app, host=host, port=port)
