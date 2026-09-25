"""Article catalog seeded from the VPR workbook and edited in the web UI."""
from __future__ import annotations

import os
import sqlite3
import datetime
from typing import Any

import xlrd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_DIR = os.environ.get("DATA_DIR") or BUNDLE_DATA_DIR
DB_PATH = os.path.join(DATA_DIR, "articles.db")
SOURCE_XLS = os.path.join(BUNDLE_DATA_DIR, "articles_source.xls")

# (field, Russian label, header needles in priority order)
FIELDS: list[tuple[str, str, tuple[str, ...]]] = [
    ("article", "Артикул", ("артикул",)),
    ("description", "Описание товара", ("описание товара",)),
    ("origin_code", "Код страны происхождения", ("код страны происхождения",)),
    ("hs_code", "Код товара", ("код товара",)),
    ("group_description", "Описание группы", ("описание группы",)),
    ("manufacturer", "Наименование фирмы-изготовителя", ("фирмы-изготовителя", "изготовителя")),
    ("brand", "Марка", ("марка",)),
    ("model", "Модель", ("модель",)),
    (
        "extra_code",
        "Доп. код",
        ("дополнительной таможенной", "классификатору"),
    ),
]
FIELD_KEYS = [key for key, _label, _needles in FIELDS]


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def _connect() -> sqlite3.Connection:
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _init_schema(db: sqlite3.Connection) -> None:
    columns = ",\n".join(f"{key} TEXT NOT NULL DEFAULT ''" for key in FIELD_KEYS if key != "article")
    db.execute(
        f"""
        CREATE TABLE IF NOT EXISTS articles (
            id INTEGER PRIMARY KEY,
            article TEXT NOT NULL COLLATE NOCASE UNIQUE,
            {columns},
            updated_at TEXT NOT NULL
        )
        """
    )
    db.execute("CREATE INDEX IF NOT EXISTS articles_article_idx ON articles(article)")
    columns = {row[1] for row in db.execute("PRAGMA table_info(articles)")}
    if "trademark_note" in columns:
        db.execute("ALTER TABLE articles DROP COLUMN trademark_note")


def _cell_text(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    return str(value).strip().upper()


def _norm_header(value: Any) -> str:
    return " ".join(_cell_text(value).lower().split())


def _find_sheet(book: xlrd.book.Book):
    for name in book.sheet_names():
        if "ОСНОВНАЯ" in name.upper():
            return book.sheet_by_name(name)
    return book.sheet_by_index(0)


def _header_map(sheet) -> tuple[int, dict[str, int]]:
    for row_idx in range(min(8, sheet.nrows)):
        headers = [_norm_header(sheet.cell_value(row_idx, col)) for col in range(sheet.ncols)]
        if not any("артикул" == text or text.startswith("артикул") for text in headers):
            continue
        mapping: dict[str, int] = {}
        for key, _label, needles in FIELDS:
            for col, text in enumerate(headers):
                if not text or col in mapping.values():
                    continue
                if any(needle in text for needle in needles):
                    # «код товара» must not take the additional-classifier column
                    if key == "hs_code" and "классификатор" in text:
                        continue
                    mapping[key] = col
                    break
        if "article" in mapping and "description" in mapping:
            return row_idx, mapping
    raise ValueError("В файле не найдена вкладка с колонками артикула и описания товара")


def _load_source_rows() -> list[dict[str, str]]:
    if not os.path.isfile(SOURCE_XLS):
        return []
    book = xlrd.open_workbook(SOURCE_XLS)
    sheet = _find_sheet(book)
    header_row, mapping = _header_map(sheet)
    by_article: dict[str, dict[str, str]] = {}
    for row_idx in range(header_row + 1, sheet.nrows):
        record = {key: "" for key in FIELD_KEYS}
        for key, col in mapping.items():
            record[key] = _cell_text(sheet.cell_value(row_idx, col))
        article = record["article"]
        if not article:
            continue
        by_article[article.casefold()] = record
    return list(by_article.values())


def _seed_if_empty(db: sqlite3.Connection) -> None:
    count = db.execute("SELECT COUNT(*) AS c FROM articles").fetchone()["c"]
    if count:
        return
    rows = _load_source_rows()
    if not rows:
        return
    stamp = "2020-01-01T00:00:00+00:00"
    placeholders = ", ".join("?" for _ in FIELD_KEYS)
    columns = ", ".join(FIELD_KEYS)
    db.executemany(
        f"INSERT INTO articles ({columns}, updated_at) VALUES ({placeholders}, ?)",
        [tuple(row[key] for key in FIELD_KEYS) + (stamp,) for row in rows],
    )


def init() -> None:
    with _connect() as db:
        _init_schema(db)
        _seed_if_empty(db)
        db.commit()


def field_meta() -> list[dict[str, str]]:
    return [{"key": key, "label": label} for key, label, _needles in FIELDS]


def _row_dict(row: sqlite3.Row) -> dict[str, Any]:
    data = {"id": row["id"], "updated_at": row["updated_at"]}
    for key in FIELD_KEYS:
        data[key] = row[key] or ""
    return data


def _clean_payload(payload: dict[str, Any], *, require_article: bool) -> dict[str, str]:
    if not isinstance(payload, dict):
        raise ValueError("Ожидается объект с полями артикула")
    cleaned = {key: str(payload.get(key) or "").strip().upper() for key in FIELD_KEYS}
    if require_article and not cleaned["article"]:
        raise ValueError("Укажите артикул")
    if len(cleaned["article"]) > 200:
        raise ValueError("Артикул слишком длинный")
    return cleaned


def _find_conflict(db: sqlite3.Connection, article: str, exclude_id: int | None) -> sqlite3.Row | None:
    if exclude_id is None:
        return db.execute(
            "SELECT id FROM articles WHERE article = ? COLLATE NOCASE",
            (article,),
        ).fetchone()
    return db.execute(
        "SELECT id FROM articles WHERE article = ? COLLATE NOCASE AND id != ?",
        (article, exclude_id),
    ).fetchone()


def _normalize_filters(filters: dict[str, Any] | None, skip_field: str | None = None) -> dict[str, list[str]]:
    if not filters:
        return {}
    if not isinstance(filters, dict):
        raise ValueError("Фильтр должен быть объектом")
    cleaned: dict[str, list[str]] = {}
    for key, values in filters.items():
        if key not in FIELD_KEYS or key == skip_field:
            continue
        if not isinstance(values, list):
            raise ValueError("Фильтр столбца должен быть списком значений")
        cleaned[key] = [str(value) for value in values]
    return cleaned


def _where(query: str, filters: dict[str, Any] | None, skip_field: str | None = None) -> tuple[str, list[Any]]:
    parts: list[str] = []
    params: list[Any] = []
    text = (query or "").strip()
    if text:
        like = f"%{text.replace('%', '').replace('_', '')}%"
        parts.append("(" + " OR ".join(f"{key} LIKE ? COLLATE NOCASE" for key in FIELD_KEYS) + ")")
        params.extend([like] * len(FIELD_KEYS))
    for key, values in _normalize_filters(filters, skip_field).items():
        if not values:
            parts.append("0")
            continue
        blanks = any(not str(value).strip() for value in values)
        concrete = [value for value in values if str(value).strip()]
        clauses: list[str] = []
        if blanks:
            clauses.append(f"TRIM(COALESCE({key}, '')) = ''")
        if concrete:
            placeholders = ", ".join("?" for _ in concrete)
            clauses.append(f"{key} IN ({placeholders})")
            params.extend(concrete)
        parts.append("(" + " OR ".join(clauses) + ")")
    if not parts:
        return "", []
    return "WHERE " + " AND ".join(parts), params


def search(query: str, page: int, per_page: int, filters: dict[str, Any] | None = None) -> dict[str, Any]:
    page = max(1, int(page or 1))
    per_page = min(max(1, int(per_page or 40)), 200)
    where, params = _where(query, filters)
    with _connect() as db:
        total = db.execute(f"SELECT COUNT(*) AS c FROM articles {where}", params).fetchone()["c"]
        rows = db.execute(
            f"""
            SELECT * FROM articles
            {where}
            ORDER BY article COLLATE NOCASE
            LIMIT ? OFFSET ?
            """,
            [*params, per_page, (page - 1) * per_page],
        ).fetchall()
    return {
        "fields": field_meta(),
        "items": [_row_dict(row) for row in rows],
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": max(1, (total + per_page - 1) // per_page) if total else 1,
    }


def distinct_values(field: str, query: str, filters: dict[str, Any] | None = None) -> list[str]:
    if field not in FIELD_KEYS:
        raise ValueError("Неизвестный столбец")
    where, params = _where(query, filters, skip_field=field)
    with _connect() as db:
        rows = db.execute(
            f"""
            SELECT DISTINCT CASE
                WHEN TRIM(COALESCE({field}, '')) = '' THEN ''
                ELSE {field}
            END AS value
            FROM articles
            {where}
            ORDER BY CASE WHEN value = '' THEN 0 ELSE 1 END, value COLLATE NOCASE
            """,
            params,
        ).fetchall()
    values = [row["value"] for row in rows]
    if "" not in values:
        values.insert(0, "")
    return values


def most_common_description(hs_code: str) -> dict[str, Any]:
    code = str(hs_code or "").strip()
    if not code:
        return {"description": "", "count": 0}
    with _connect() as db:
        row = db.execute(
            """
            SELECT description, COUNT(*) AS n
            FROM articles
            WHERE hs_code = ? AND TRIM(description) != ''
            GROUP BY description
            ORDER BY n DESC, description COLLATE NOCASE
            LIMIT 1
            """,
            (code,),
        ).fetchone()
    if row is None:
        return {"description": "", "count": 0}
    return {"description": row["description"], "count": int(row["n"])}


def get_article(article_id: int) -> dict[str, Any] | None:
    with _connect() as db:
        row = db.execute("SELECT * FROM articles WHERE id = ?", (article_id,)).fetchone()
    return _row_dict(row) if row else None


def create_article(payload: dict[str, Any]) -> dict[str, Any]:
    cleaned = _clean_payload(payload, require_article=True)
    with _connect() as db:
        if _find_conflict(db, cleaned["article"], None):
            raise ValueError("Такой артикул уже есть в базе")
        cur = db.execute(
            f"INSERT INTO articles ({', '.join(FIELD_KEYS)}, updated_at) VALUES ({', '.join('?' for _ in FIELD_KEYS)}, ?)",
            tuple(cleaned[key] for key in FIELD_KEYS) + (_now(),),
        )
        db.commit()
        article_id = int(cur.lastrowid)
    created = get_article(article_id)
    if created is None:
        raise RuntimeError("Не удалось сохранить артикул")
    return created


def export_all() -> list[dict[str, Any]]:
    with _connect() as db:
        rows = db.execute("SELECT * FROM articles ORDER BY article COLLATE NOCASE").fetchall()
    return [_row_dict(row) for row in rows]


def apply_remote(items: list[dict[str, Any]]) -> dict[str, int]:
    """Keep the newer copy of each article. Used by both sides of the sync."""
    inserted = updated = 0
    with _connect() as db:
        for item in items:
            if not isinstance(item, dict):
                continue
            cleaned = _clean_payload(item, require_article=False)
            if not cleaned["article"]:
                continue
            remote_stamp = str(item.get("updated_at") or "")
            existing = _find_conflict(db, cleaned["article"], None)
            if existing is None:
                db.execute(
                    f"INSERT INTO articles ({', '.join(FIELD_KEYS)}, updated_at) VALUES ({', '.join('?' for _ in FIELD_KEYS)}, ?)",
                    tuple(cleaned[key] for key in FIELD_KEYS) + (remote_stamp or _now(),),
                )
                inserted += 1
                continue
            current = db.execute("SELECT * FROM articles WHERE id = ?", (existing["id"],)).fetchone()
            local_stamp = current["updated_at"] or ""
            if remote_stamp and local_stamp and remote_stamp <= local_stamp:
                continue
            same = all((current[key] or "") == cleaned[key] for key in FIELD_KEYS)
            if same:
                continue
            assignments = ", ".join(f"{key} = ?" for key in FIELD_KEYS)
            db.execute(
                f"UPDATE articles SET {assignments}, updated_at = ? WHERE id = ?",
                tuple(cleaned[key] for key in FIELD_KEYS) + (remote_stamp or _now(), existing["id"]),
            )
            updated += 1
        db.commit()
    return {"inserted": inserted, "updated": updated}


def update_article(article_id: int, payload: dict[str, Any]) -> dict[str, Any]:
    cleaned = _clean_payload(payload, require_article=True)
    with _connect() as db:
        existing = db.execute("SELECT id FROM articles WHERE id = ?", (article_id,)).fetchone()
        if not existing:
            raise LookupError("Артикул не найден")
        if _find_conflict(db, cleaned["article"], article_id):
            raise ValueError("Такой артикул уже есть в базе")
        assignments = ", ".join(f"{key} = ?" for key in FIELD_KEYS)
        db.execute(
            f"UPDATE articles SET {assignments}, updated_at = ? WHERE id = ?",
            tuple(cleaned[key] for key in FIELD_KEYS) + (_now(), article_id),
        )
        db.commit()
    updated = get_article(article_id)
    if updated is None:
        raise RuntimeError("Не удалось сохранить артикул")
    return updated


init()
