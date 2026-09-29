"""Build a catalog-shaped table from article codes found in an arbitrary Excel file."""
from __future__ import annotations

import io
from typing import Any

import openpyxl
from openpyxl.styles import Font

import articles
import excel_io

# Headers that mark an article column even when none of its codes are in the catalog yet.
_HEADER_EXACT = {
    "артикул",
    "артикул товара",
    "код артикула",
    "article",
    "article no",
    "article number",
    "sku",
    "item",
    "item no",
    "item number",
    "part no",
    "part number",
    "product code",
}
_HEADER_SCAN_ROWS = 25


def _token(value: Any) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip().upper()
    if isinstance(value, int):
        return str(value)
    return " ".join(str(value).strip().upper().split())


def _is_article_header(text: str) -> bool:
    norm = " ".join(text.casefold().split())
    if not norm:
        return False
    if norm in _HEADER_EXACT or norm.startswith("артикул"):
        return True
    return False


def _grids(content: bytes, filename: str | None) -> list[tuple[str, list[list[str]]]]:
    book = excel_io.load_workbook(content, data_only=True, filename=filename)
    grids: list[tuple[str, list[list[str]]]] = []
    for sheet in book.worksheets:
        rows = [[_token(cell) for cell in row] for row in sheet.iter_rows(values_only=True)]
        grids.append((sheet.title, rows))
    return grids


def _looks_like_article(token: str) -> bool:
    if _is_article_header(token) or len(token) < 3:
        return False
    return any(ch.isdigit() for ch in token) or "-" in token or "/" in token


def _column_values(rows: list[list[str]], col: int, start: int, index: dict[str, dict[str, Any]]) -> list[str]:
    values: list[str] = []
    for row in rows[start:]:
        if col >= len(row):
            continue
        token = row[col]
        if not token or _is_article_header(token):
            continue
        if token.casefold() in index or _looks_like_article(token):
            values.append(token)
    return values


def _best_column(
    grids: list[tuple[str, list[list[str]]]],
    index: dict[str, dict[str, Any]],
) -> tuple[int, int, int, bool] | None:
    """Return (sheet, column, data start, header was recognized) for the article column."""
    header_choices: list[tuple[int, int, int, int, int]] = []
    plain_choices: list[tuple[int, int, int, int]] = []
    for sheet_i, (_title, rows) in enumerate(grids):
        width = max((len(row) for row in rows), default=0)
        header_at: dict[int, int] = {}
        for row_i, row in enumerate(rows[:_HEADER_SCAN_ROWS]):
            for col, token in enumerate(row):
                if col not in header_at and _is_article_header(token):
                    header_at[col] = row_i
        for col in range(width):
            start = header_at[col] + 1 if col in header_at else 0
            values = _column_values(rows, col, start, index)
            hits = sum(1 for value in values if value.casefold() in index)
            if col in header_at:
                header_choices.append((hits, len(values), sheet_i, col, start))
            elif hits:
                plain_choices.append((hits, sheet_i, col, hits))
    if header_choices:
        header_choices.sort(key=lambda item: (item[0], item[1]), reverse=True)
        _hits, _count, sheet_i, col, start = header_choices[0]
        return sheet_i, col, start, True
    if not plain_choices:
        return None
    plain_choices.sort(reverse=True)
    best_hits, sheet_i, col, _hits = plain_choices[0]
    second = plain_choices[1][0] if len(plain_choices) > 1 else 0
    if best_hits < 1 or best_hits == second:
        return None
    return sheet_i, col, 0, False


_COMPARE_KEYS = ("hs_code", "origin_code", "group_description")


def _compare_kind(token: str) -> str | None:
    norm = " ".join(token.casefold().split())
    if not norm or "описание товара" in norm:
        return None
    if "описание группы" in norm or norm == "описание":
        return "group_description"
    if "происхожд" in norm or norm in {"origin", "origin code", "country", "country of origin"}:
        return "origin_code"
    if "код страны" in norm:
        return "origin_code"
    if "код товара" in norm or norm in {"hs code", "hs", "коды", "hscode", "hs-code"}:
        return "hs_code"
    if "hs" in norm and "code" in norm:
        return "hs_code"
    return None


def _compare_map(header: list[str]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col, token in enumerate(header):
        kind = _compare_kind(token)
        if kind and kind not in mapping:
            mapping[kind] = col
    return mapping


def _blank_item(code: str) -> dict[str, str]:
    item = {key: "" for key in articles.FIELD_KEYS}
    item["article"] = code
    return item


def _catalog_item(row_data: dict[str, Any]) -> dict[str, str]:
    return {key: row_data[key] for key in articles.FIELD_KEYS}


def build_load_table(content: bytes, filename: str | None = None) -> dict[str, Any]:
    if not content:
        raise ValueError("Файл пустой")
    grids = _grids(content, filename)
    if not grids or not any(rows for _title, rows in grids):
        raise ValueError("В файле нет данных")
    index = articles.index_by_article()
    chosen = _best_column(grids, index)
    items: list[dict[str, str]] = []
    diffs: list[dict[str, Any]] = []
    seen_diffs: set[str] = set()
    compared: dict[str, int] = {}
    if chosen is None:
        mode = "scan"
        for _title, rows in grids:
            for row in rows:
                for token in row:
                    if not token or _is_article_header(token):
                        continue
                    row_data = index.get(token.casefold())
                    if row_data is None:
                        continue
                    items.append(_catalog_item(row_data))
    else:
        mode = "column"
        sheet_i, col, start, _header = chosen
        rows = grids[sheet_i][1]
        compared = _compare_map(rows[start - 1] if start else [])
        for row in rows[start:]:
            code = row[col] if col < len(row) else ""
            if not code or _is_article_header(code):
                continue
            if code.casefold() not in index and not _looks_like_article(code):
                continue
            row_data = index.get(code.casefold())
            items.append(_catalog_item(row_data) if row_data else _blank_item(code))
            if row_data is None or code.casefold() in seen_diffs:
                continue
            file_values = {
                key: (row[pos] if pos < len(row) else "")
                for key, pos in compared.items()
            }
            changed = [
                key
                for key in _COMPARE_KEYS
                if file_values.get(key) and file_values.get(key) != (row_data[key] or "")
            ]
            if not changed:
                continue
            seen_diffs.add(code.casefold())
            diffs.append(
                {
                    "id": row_data["id"],
                    "article": row_data["article"],
                    "db": _catalog_item(row_data),
                    "file": {key: file_values.get(key, "") for key in _COMPARE_KEYS},
                    "fields": changed,
                }
            )
    found_count = sum(1 for item in items if any(item[key] for key in articles.FIELD_KEYS if key != "article"))
    return {
        "mode": mode,
        "sheet": grids[chosen[0]][0] if chosen else "",
        "total": len(items),
        "found_count": found_count,
        "items": items,
        "fields": articles.field_meta(),
        "diffs": diffs,
        "compared_fields": list(compared),
    }


def load_table_xlsx(result: dict[str, Any]) -> bytes:
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Загрузочная таблица"
    fields = result.get("fields") or articles.field_meta()
    headers = [field["label"] for field in fields]
    keys = [field["key"] for field in fields]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for item in result.get("items") or []:
        sheet.append([item.get(key) or "" for key in keys])
    for column in sheet.columns:
        letter = column[0].column_letter
        width = max(12, min(48, max(len(str(cell.value or "")) for cell in column) + 2))
        sheet.column_dimensions[letter].width = width
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()
