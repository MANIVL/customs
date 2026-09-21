"""AI-assisted column mapping for Excel → template fields via YandexGPT."""
from __future__ import annotations

import json
import re
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

import engine
import excel_io
import yandex_gpt
import countries
import units

CANONICAL_FIELDS = [
    "no", "tariff_code", "name", "article", "marks", "manufacturer",
    "country", "unit", "qty", "price", "amount", "net_weight",
    "gross_weight", "cll", "mnr", "date_from", "date_to", "mnr_code",
]

FIELD_HINTS = {
    "no": "порядковый номер позиции (№, No.) — обычно 1,2,3…",
    "tariff_code": "код ТН ВЭД / HS (длинный числовой код)",
    "name": "наименование / описание товара — длинный текст, слова, не артикул",
    "article": "артикул / SKU / model — короткий код вроде E2903-00, без пробелов",
    "marks": "торговая марка / brand",
    "manufacturer": "производитель / manufacturer",
    "country": "страна происхождения",
    "unit": "единица измерения (шт, kg…)",
    "qty": "количество — число",
    "price": "цена за единицу — число",
    "amount": "стоимость / итого строки — число",
    "net_weight": "вес нетто — число",
    "gross_weight": "вес брутто / gross — число",
    "cll": "количество мест / паллет",
    "mnr": "номер сертификата/декларации",
    "date_from": "дата начала действия сертификата",
    "date_to": "дата окончания действия сертификата",
    "mnr_code": "код вида документа МНР",
}

# Exact-ish header → field (checked before AI). Article before generic "описание".
HEADER_EXACT = [
    ("артикул", "article"),
    ("article", "article"),
    ("item no", "article"),
    ("item no.", "article"),
    ("item#", "article"),
    ("item number", "article"),
    ("sku", "article"),
    ("model no", "article"),
    ("model", "article"),
    ("наименование товара", "name"),
    ("name of product", "name"),
    ("наименование", "name"),
    ("описание товара", "name"),  # group title — may be overridden by samples
    ("описание", "name"),
    ("description", "name"),
    ("код изделия", "article"),
    ("код товара", "tariff_code"),
    ("таможенный код", "tariff_code"),
    ("тн вэд", "tariff_code"),
    ("hs code", "tariff_code"),
    ("tariff", "tariff_code"),
    ("кол-во мест", "cll"),
    ("количество мест", "cll"),
    ("number of units", "cll"),
    ("cartons", "cll"),
    ("кол-во кор", "cll"),
    ("кол во кор", "cll"),
    ("кол-во паллет", "cll"),
    ("количество паллет", "cll"),
    ("кол-во шт на паллете", "_skip"),
    ("qty per pallet", "_skip"),
    ("вес нетто/паллет", "_skip"),
    ("вес гросс/паллет", "_skip"),
    ("нетто/паллет", "_skip"),
    ("гросс/паллет", "_skip"),
    ("общий вес нетто", "net_weight"),
    ("общий вес гросс", "gross_weight"),
    ("общий вес брутто", "gross_weight"),
    ("total net weight", "net_weight"),
    ("total gross weight", "gross_weight"),
    ("всего шт", "qty"),
    ("всего, шт", "qty"),
    ("кол-во", "qty"),
    ("количество", "qty"),
    ("q-ty", "qty"),
    ("qty", "qty"),
    ("нетто вес", "net_weight"),
    ("вес нетто", "net_weight"),
    ("нетто", "net_weight"),
    ("net weight", "net_weight"),
    ("гросс вес", "gross_weight"),
    ("вес брутто", "gross_weight"),
    ("брутто", "gross_weight"),
    ("gross", "gross_weight"),
    ("цена", "price"),
    ("price", "price"),
    ("стоимость", "amount"),
    ("итого", "amount"),
    ("сумма", "amount"),
    ("amount", "amount"),
    ("единица измерения", "unit"),
    ("unit of measure", "unit"),
    ("unit /", "unit"),
    ("uom", "unit"),
    ("unit", "unit"),
    ("страна происхождения товара", "country"),
    ("страна происхождения", "country"),
    ("country of origin", "country"),
    ("страна-производитель", "country"),
    ("страна производитель", "country"),
    ("origin country", "country"),
    ("place of origin", "country"),
    ("made in", "country"),
    ("происхождение", "country"),
    ("страна", "country"),
    ("country", "country"),
    ("原产国", "country"),
    ("原产地", "country"),
    ("производитель", "manufacturer"),
    ("manufacturer", "manufacturer"),
    ("торговая марка", "marks"),
    ("brand", "marks"),
    # Short tokens last — matched as whole words only (see _header_phrase_match)
    ("кор", "cll"),
    ("ctn", "cll"),
    ("гросс", "gross_weight"),
]

ARTICLE_RE = re.compile(r"^[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._\-/]{1,40}$")
CURRENCY_RE = re.compile(
    r"^(USD|EUR|CNY|RMB|RUR|RUB|GBP|JPY|CHF|HKD|\$|€|¥)$",
    re.IGNORECASE,
)
UNIT_TOKENS = frozenset({
    "шт", "шт.", "pcs", "pc", "pcs.", "set", "sets", "kg", "кг", "кg",
    "кор", "кор.", "ctn", "carton", "cartons",
})
HEADER_LIKE_ARTICLES = frozenset({
    "артикул", "article", "sku", "описание", "description", "наименование",
    "name", "model", "модель",
})
SKIP_ROW_MARKERS = (
    "всего", "итого", "total", "сумма:", "subtotal", "grand total",
    "sum", "the seller", "seller:", "buyer", "consignee", "паллет всего",
    "всес паллет", "всего паллет", "境内货源地", "shipping marks",
    "packing:", "measurement:", "order no", "контейнер", "container",
)
SUM_FIELDS = frozenset({"qty", "amount", "net_weight", "gross_weight", "cll"})
MAX_EXTRACT_ROWS = 1500
_CATALOG_NAME_MARKERS = (
    "прайс", "каталог", "справочник", "pricelist", "price list",
    "catalog", "nomenclature", "номенклатур",
)

SYSTEM_PROMPT = """Ты помощник по таможенным Excel-документам.
Тебе дают листы: заголовки с номерами столбцов (col), подзаголовки и примеры ячеек {col: значение}.

Сопоставь столбцы с каноническими полями. Ключи полей:
""" + "\n".join(f"- {k}: {FIELD_HINTS[k]}" for k in CANONICAL_FIELDS) + """

ЖЁСТКИЕ ПРАВИЛА:
1) «Артикул» / Article / SKU → ТОЛЬКО article. Туда короткие коды (EDAY132RU-00), НЕ описание товара.
2) «Наименование» / Name / Description / описание → ТОЛЬКО name. Туда длинный текст, НЕ артикул.
3) Часто бывает ДВУХСТРОЧНЫЙ заголовок: сверху «Описание товара / Кол-во / Цена», снизу «артикул | описание».
   Тогда article и name бери из нижней строки; qty/price/amount — из верхней или по числам в примерах.
4) Нельзя назначить один столбец двум полям.
5) Столбцы только с «шт», «кг», «кор», «RMB», «CNY», «USD» — это unit/валюта, НЕ price/amount/qty.
   Если рядом с валютой есть число (RMB | 345.6) — price/amount = столбец с ЧИСЛОМ.
6) На packing list мапь net_weight / gross_weight / cll (места/кор) и обязательно article+name, если они есть в примерах.
7) data_start_row — первая строка с реальным артикулом (не «артикул», не категория, не контейнер).
8) PO/NO, Order No, order_32_26 — это номер заказа, НЕ article. Article = Item NO / SKU (E36124-CP).
9) Кол-во (qty) обычно целые 3, 60, 102; Цена (price) — с копейками 618.42; amount ≈ qty × price. Не путай эти столбцы.
10) «Price Terms:», «Payment Terms:», «Country of Origin:», «Currency:» — поля шапки документа, НЕ заголовки таблицы.
11) «Общий гросс» / total gross — не net_weight и не price; строковый гросс/нетто бери из «Гросс вес» / «Нетто вес».
12) Ответ — ТОЛЬКО JSON-массив (без markdown):
[{"sheet":"имя","header_row":N,"data_start_row":M,"mapping":{"article":2,"name":3,"qty":4,"price":9,"amount":12}}]
col в mapping — 1-based номер столбца Excel.
"""


def _cell_str(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).replace("\t", " ").replace("\n", " ").strip()


def _looks_like_article(val: Any) -> bool:
    s = _cell_str(val)
    if not s or " " in s or len(s) > 40:
        return False
    if s.lower() in HEADER_LIKE_ARTICLES:
        return False
    return bool(ARTICLE_RE.match(s)) and not s.replace(".", "").isdigit()


def _looks_like_order_ref(val: Any) -> bool:
    """PO / order number, not a product SKU."""
    s = _cell_str(val).lower()
    if not s:
        return False
    return bool(re.match(r"^(order[_\-\s]|po[_\-\s/#]|po/no)", s))


def _looks_like_sku(val: Any) -> bool:
    """Product code like E36124-CP, not order_32_26."""
    if _looks_like_order_ref(val) or not _looks_like_article(val):
        return False
    s = _cell_str(val)
    return bool(re.search(r"[A-Za-z]", s) and re.search(r"\d", s))


def _is_metadata_label(text: str) -> bool:
    """Document header fields like «Price Terms:», not table column titles."""
    s = _cell_str(text)
    if not s:
        return False
    if s.endswith(":") or s.endswith("："):
        return True
    n = engine.normalize(s)
    prefixes = (
        "payment terms", "price terms", "invoice no", "b/l", "messrs",
        "ship from", "ship via", "delivery to", "currency", "contact",
        "container no", "seal no", "to port",
    )
    return n.startswith(prefixes)


def _looks_like_name(val: Any) -> bool:
    s = _cell_str(val)
    if not s:
        return False
    if s.lower() in HEADER_LIKE_ARTICLES:
        return False
    return (" " in s and len(s) >= 8) or len(s) >= 20


def _looks_like_currency(val: Any) -> bool:
    return bool(CURRENCY_RE.match(_cell_str(val)))


def _looks_like_unit(val: Any) -> bool:
    s = _cell_str(val).lower().rstrip(".")
    return s in UNIT_TOKENS


def _looks_like_hs_code(val: Any) -> bool:
    s = _cell_str(val).replace(" ", "").replace(".", "")
    return bool(re.fullmatch(r"\d{6,12}", s))


def _sheet_is_catalog(summary: dict) -> bool:
    """Skip price-list / nomenclature sheets (thousands of SKUs, not the shipment)."""
    name = engine.normalize(summary.get("sheet") or "")
    compact = re.sub(r"[^a-zа-яё0-9]+", "", name)
    if compact in {"прайс", "price", "catalog", "каталог"}:
        return True
    if any(m in name for m in _CATALOG_NAME_MARKERS):
        return True
    max_row = int(summary.get("max_row") or 0)
    if max_row < 400:
        return False
    km = summary.get("keyword_mapping") or {}
    has_trade = any(km.get(k) for k in ("price", "amount", "net_weight", "gross_weight"))
    return not has_trade


def _find_candidate_header_rows(ws: Worksheet, max_scan: int = 50) -> list[int]:
    max_row = min(ws.max_row or 1, max_scan)
    max_col = min(ws.max_column or 1, 30)
    scored: list[tuple[int, int, int]] = []
    keyword_bags = [engine.HEADER_KEYWORDS[k] for k in (
        "article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "no", "tariff_code", "cll"
    )]
    for r in range(1, max_row + 1):
        texts = []
        hits = 0
        has_article = False
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str) or not v.strip():
                continue
            text = v.strip()
            if _is_metadata_label(text):
                continue
            texts.append(text)
            norm = engine.normalize(text)
            if any(_header_phrase_match(p, norm) for p, _f in HEADER_EXACT):
                hits += 1
            elif any(any(kw in norm for kw in bag) for bag in keyword_bags):
                hits += 1
            if "артикул" in norm or "item no" in norm or norm in {"article", "sku"}:
                has_article = True
        if len(texts) < 2 and not has_article:
            continue
        score = hits * 10 + len(texts) + (25 if has_article else 0)
        if hits >= 2 or len(texts) >= 4 or has_article:
            scored.append((score, len(texts), r))
    scored.sort(reverse=True)
    return [r for _s, _n, r in scored[:8]] or [1]


def _headers_from_rows(ws: Worksheet, rows: list[int], max_col: int) -> list[dict]:
    """Collect labeled cells from one or more header/subheader rows."""
    by_col: dict[int, str] = {}
    for r in rows:
        for c in range(1, max_col + 1):
            text = _cell_str(ws.cell(r, c).value)
            if not text or _is_metadata_label(text):
                continue
            prev = by_col.get(c, "")
            prev_n = engine.normalize(prev)
            cur_n = engine.normalize(text)
            if not prev:
                by_col[c] = text
            elif ("артикул" in cur_n or "item no" in cur_n or cur_n == "article") and "артикул" not in prev_n:
                by_col[c] = text
            elif ("описание" in cur_n or "наименование" in cur_n) and len(text) <= len(prev):
                by_col[c] = text
            elif _is_metadata_label(prev) and not _is_metadata_label(text):
                by_col[c] = text
    return [{"col": c, "text": by_col[c]} for c in sorted(by_col)]


def _header_phrase_match(phrase: str, norm: str) -> bool:
    """Word-boundary match for single tokens; substring for multi-word labels."""
    return engine.header_keyword_matches(phrase, norm)


def _keyword_mapping_from_headers(headers: list[dict]) -> dict[str, int]:
    """Deterministic mapping from clear header labels; article wins over name on same col."""
    field_cols: dict[str, int] = {}
    col_field: dict[int, str] = {}
    priority = {
        "article": 100, "name": 90, "qty": 80, "price": 70, "amount": 70,
        "net_weight": 75, "gross_weight": 75, "cll": 65, "no": 95, "tariff_code": 85,
        "marks": 40, "country": 55, "manufacturer": 40, "unit": 30,
    }

    def assign(field: str, col: int) -> None:
        if field not in CANONICAL_FIELDS:
            return
        old_field = col_field.get(col)
        if old_field and priority.get(old_field, 0) > priority.get(field, 0):
            return
        if old_field and old_field in field_cols and field_cols[old_field] == col:
            del field_cols[old_field]
        if field in field_cols and field_cols[field] != col:
            return
        field_cols[field] = col
        col_field[col] = field

    for h in headers:
        if _is_metadata_label(h["text"]):
            continue
        norm = engine.normalize(h["text"])
        if not norm:
            continue
        # Skip long free-text / codes mistaken for headers
        if len(norm) > 40 or _looks_like_article(h["text"]) or _looks_like_name(h["text"]):
            if not any(_header_phrase_match(p, norm) for p, _f in HEADER_EXACT):
                continue
        for phrase, field in HEADER_EXACT:
            if not _header_phrase_match(phrase, norm):
                continue
            # PO/Order No is not the product SKU (Item NO / артикул is).
            if field == "article" and any(x in norm for x in ("po/no", "po no", "order no", "номер заказа")) and "item" not in norm:
                continue
            # «Общий гросс» without «вес» is an extras total, not line weight.
            # Keep «общий вес нетто/гросс» — that is the SKU line total.
            if field in ("gross_weight", "net_weight", "amount") and (
                "общ" in norm or "total" in norm or "grand" in norm
            ) and "вес" not in norm and "weight" not in norm:
                continue
            assign(field, h["col"])
            break
    return field_cols


def _wide_samples(ws: Worksheet, start_row: int, max_col: int, limit: int = 8) -> list[dict]:
    samples = []
    max_row = ws.max_row or start_row
    for r in range(start_row, min(start_row + 12, max_row + 1)):
        cells = {}
        for c in range(1, max_col + 1):
            val = _cell_str(ws.cell(r, c).value)
            if val:
                cells[str(c)] = val[:120]
        if len(cells) >= 2:
            samples.append({"row": r, "cells": cells})
        if len(samples) >= limit:
            break
    return samples


def summarize_sheet(ws: Worksheet) -> dict:
    max_col = min(ws.max_column or 1, 25)
    header_candidates = _find_candidate_header_rows(ws)
    best = header_candidates[0]
    # Include nearby label rows only (skip numeric data rows mistaken as candidates)
    header_rows = [best]

    def row_table_label_score(r: int) -> int:
        score = 0
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            if _is_metadata_label(v):
                continue
            norm = engine.normalize(v)
            if not norm or _to_number(v) is not None:
                continue
            if any(_header_phrase_match(p, norm) for p, _f in HEADER_EXACT):
                score += 2
        return score

    for r in range(max(1, best - 1), min((ws.max_row or best), best + 1) + 1):
        if r == best:
            continue
        if row_table_label_score(r) >= 2:
            header_rows.append(r)
    header_rows = sorted(set(header_rows))
    headers = _headers_from_rows(ws, header_rows, max_col)
    samples = _wide_samples(ws, best + 1, max_col)
    keyword_mapping = _keyword_mapping_from_headers(headers)
    keyword_mapping = _refine_mapping_with_samples(keyword_mapping, {"samples": samples})
    keyword_mapping = _infer_missing_from_samples(keyword_mapping, samples)

    return {
        "sheet": ws.title,
        "max_row": ws.max_row or 0,
        "header_candidates": header_candidates,
        "suggested_header_row": best,
        "header_rows": header_rows,
        "headers": headers,
        "samples": samples,
        "keyword_mapping": keyword_mapping,
    }


def _parse_json_reply(text: str):
    raw = text.strip()
    if "```" in raw:
        for part in raw.split("```"):
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("{") or part.startswith("["):
                raw = part
                break
    start_obj = raw.find("{")
    start_arr = raw.find("[")
    if start_obj < 0 and start_arr < 0:
        raise ValueError(f"JSON не найден в ответе ИИ: {raw[:200]}")
    start = start_arr if start_arr >= 0 and (start_obj < 0 or start_arr < start_obj) else start_obj
    data, _end = json.JSONDecoder().raw_decode(raw[start:])
    return data


def _as_sheet_list(data) -> list[dict]:
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        if isinstance(data.get("sheets"), list):
            return [x for x in data["sheets"] if isinstance(x, dict)]
        if "mapping" in data or "sheet" in data:
            return [data]
    return []


def _normalize_mapping(raw_mapping: dict) -> dict[str, int]:
    mapping: dict[str, int] = {}
    used: set[int] = set()
    for key, col in (raw_mapping or {}).items():
        try:
            col_i = int(col)
        except (TypeError, ValueError):
            continue
        if key in CANONICAL_FIELDS and col_i > 0 and col_i not in used:
            mapping[key] = col_i
            used.add(col_i)
    return mapping


def _refine_mapping_with_samples(mapping: dict[str, int], summary: dict) -> dict[str, int]:
    """Fix article/name swaps and currency-mapped price/amount using sample values."""
    mapping = dict(mapping)
    samples = summary.get("samples") or []
    if not samples:
        return mapping

    def sample_vals(col: int) -> list[str]:
        out = []
        for s in samples:
            v = (s.get("cells") or {}).get(str(col))
            if v:
                out.append(v)
        return out

    def col_numeric_score(col: int) -> int:
        vals = sample_vals(col)
        return sum(1 for v in vals if _to_number(v) is not None)

    def col_currency_score(col: int) -> int:
        return sum(1 for v in sample_vals(col) if _looks_like_currency(v))

    art_col = mapping.get("article")
    name_col = mapping.get("name")

    if art_col and name_col:
        art_vals = sample_vals(art_col)
        name_vals = sample_vals(name_col)
        art_as_code = sum(_looks_like_article(v) for v in art_vals)
        art_as_name = sum(_looks_like_name(v) for v in art_vals)
        name_as_code = sum(_looks_like_article(v) for v in name_vals)
        name_as_name = sum(_looks_like_name(v) for v in name_vals)
        if art_as_name > art_as_code and name_as_code > name_as_name:
            mapping["article"], mapping["name"] = name_col, art_col

    elif art_col and not name_col:
        vals = sample_vals(art_col)
        if vals and sum(_looks_like_name(v) for v in vals) > sum(_looks_like_article(v) for v in vals):
            mapping["name"] = art_col
            del mapping["article"]

    elif name_col and not art_col:
        vals = sample_vals(name_col)
        if vals and sum(_looks_like_article(v) for v in vals) > sum(_looks_like_name(v) for v in vals):
            mapping["article"] = name_col
            del mapping["name"]

    # Prefer SKU column (E36124-CP) over PO/order (order_32_26)
    art_col = mapping.get("article")
    if art_col:
        used_now = set(mapping.values())
        art_vals = sample_vals(art_col)
        if art_vals and sum(_looks_like_order_ref(v) for v in art_vals) >= 2:
            best_c, best_s = None, 0
            cols = sorted({int(c) for s in samples for c in (s.get("cells") or {})})
            for col in cols:
                if col == art_col:
                    continue
                if col in used_now and col != mapping.get("name"):
                    continue
                sku_hits = sum(_looks_like_sku(v) for v in sample_vals(col))
                if sku_hits > best_s:
                    best_s, best_c = sku_hits, col
            if best_c and best_s >= 2:
                if mapping.get("name") == best_c:
                    del mapping["name"]
                mapping["article"] = best_c

    # Qty is usually integers; unit price often has decimals. Swap if clearly reversed.
    qty_col, price_col = mapping.get("qty"), mapping.get("price")
    if qty_col and price_col:
        def _nums(col: int) -> list[float]:
            out = []
            for v in sample_vals(col):
                n = _to_number(v)
                if n is not None:
                    out.append(n)
            return out

        qn, pn = _nums(qty_col), _nums(price_col)
        if qn and pn:
            q_int = sum(1 for x in qn if float(x).is_integer())
            p_int = sum(1 for x in pn if float(x).is_integer())
            q_dec = len(qn) - q_int
            if q_dec >= 2 and p_int >= 2 and q_int <= p_int:
                mapping["qty"], mapping["price"] = price_col, qty_col

    qty_only = mapping.get("qty")
    if qty_only:
        qvals = sample_vals(qty_only)
        if qvals and sum(_looks_like_hs_code(v) for v in qvals) >= max(2, (len(qvals) + 1) // 2):
            del mapping["qty"]

    # «Общий гросс» is a total, not the line gross weight.
    headers_by_col = {
        h["col"]: engine.normalize(h.get("text") or "")
        for h in (summary.get("headers") or [])
    }
    gw = mapping.get("gross_weight")
    if gw:
        gw_txt = headers_by_col.get(gw, "")
        if ("общ" in gw_txt or "total" in gw_txt) and "вес" not in gw_txt and "weight" not in gw_txt:
            used_now = set(mapping.values())
            for col, text in headers_by_col.items():
                if col == gw or col in used_now:
                    continue
                if any(k in text for k in ("гросс", "брутто", "gross")) and "общ" not in text and "total" not in text:
                    mapping["gross_weight"] = col
                    break

    # Price/amount headers often sit on the currency sub-column (RMB); shift to numeric neighbor.
    used = set(mapping.values())
    for field in ("price", "amount"):
        col = mapping.get(field)
        if not col:
            continue
        if col_currency_score(col) > col_numeric_score(col):
            for cand in (col + 1, col + 2, col - 1):
                if cand < 1 or cand in used:
                    continue
                if col_numeric_score(cand) >= 2 and col_currency_score(cand) == 0:
                    mapping[field] = cand
                    used.discard(col)
                    used.add(cand)
                    break

    # Drop unit/currency mistaken as qty etc.
    for field in ("qty", "price", "amount", "net_weight", "gross_weight", "cll"):
        col = mapping.get(field)
        if not col:
            continue
        vals = sample_vals(col)
        if vals and all(_looks_like_currency(v) or _looks_like_unit(v) for v in vals):
            del mapping[field]

    return mapping


def _infer_missing_from_samples(mapping: dict[str, int], samples: list[dict]) -> dict[str, int]:
    """Infer article/name/price/amount/cll from value shapes when headers were incomplete."""
    if not samples:
        return mapping
    mapping = dict(mapping)
    used = set(mapping.values())

    # Collect columns that appear in samples
    cols = sorted({int(c) for s in samples for c in (s.get("cells") or {})})

    def vals(col: int) -> list[str]:
        return [s["cells"][str(col)] for s in samples if str(col) in (s.get("cells") or {})]

    def score_article(col: int) -> int:
        vs = vals(col)
        sku = sum(_looks_like_sku(v) for v in vs)
        art = sum(_looks_like_article(v) for v in vs)
        order = sum(_looks_like_order_ref(v) for v in vs)
        return sku * 5 + art - order * 5

    def score_name(col: int) -> int:
        return sum(_looks_like_name(v) for v in vals(col))

    def score_num(col: int) -> int:
        return sum(1 for v in vals(col) if _to_number(v) is not None)

    if "article" not in mapping:
        ranked = sorted(cols, key=score_article, reverse=True)
        for col in ranked:
            if col in used:
                continue
            if score_article(col) >= 2:
                mapping["article"] = col
                used.add(col)
                break

    if "name" not in mapping:
        art = mapping.get("article")
        # Prefer column immediately right of article
        candidates = []
        if art:
            candidates.append(art + 1)
        candidates.extend(cols)
        for col in candidates:
            if col in used or col < 1:
                continue
            if score_name(col) >= 2:
                mapping["name"] = col
                used.add(col)
                break

    # Numeric fields still missing — only qty from leftover number columns
    for field in ("qty",):
        if field in mapping:
            continue
        best_col = None
        best_score = 0
        for col in cols:
            if col in used:
                continue
            if score_article(col) or score_name(col):
                continue
            vs = vals(col)
            if not vs:
                continue
            if field == "qty" and sum(_looks_like_hs_code(v) for v in vs) >= 2:
                continue
            if any(_looks_like_name(v) or _looks_like_article(v) for v in vs):
                continue
            if vs and all(_looks_like_currency(v) or _looks_like_unit(v) for v in vs):
                continue
            sc = score_num(col)
            if sc > best_score:
                best_score = sc
                best_col = col
        if best_col is not None and best_score >= 2:
            mapping[field] = best_col
            used.add(best_col)

    for field in ("price", "amount"):
        if field in mapping:
            continue
        for col in cols:
            if col in used:
                continue
            # Prefer numeric column whose left neighbor is currency
            left = vals(col - 1) if col > 1 else []
            if score_num(col) >= 2 and any(_looks_like_currency(v) for v in left):
                mapping[field] = col
                used.add(col)
                break

    return mapping


def _guess_data_start_row(ws: Worksheet, header_row: int, mapping: dict[str, int]) -> int:
    art_col = mapping.get("article")
    name_col = mapping.get("name")
    max_row = min((ws.max_row or header_row) + 1, header_row + 25)
    for r in range(header_row + 1, max_row):
        art = ws.cell(r, art_col).value if art_col else None
        name = ws.cell(r, name_col).value if name_col else None
        if art_col and _looks_like_article(art):
            return r
        if name_col and _looks_like_name(name) and not _looks_like_article(art):
            # category lines often have name-like text without article
            continue
    return header_row + 1


def _merge_mappings(keyword: dict[str, int], ai: dict[str, int], summary: dict) -> dict[str, int]:
    """Keyword mapping wins for article/name/qty/…; AI fills gaps; then sample refine."""
    merged = dict(ai)
    for key in ("article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "cll", "no", "tariff_code"):
        if key in keyword:
            ai_col = merged.get(key)
            kw_col = keyword[key]
            if ai_col and ai_col != kw_col:
                for k, c in list(merged.items()):
                    if c == kw_col and k != key:
                        del merged[k]
            merged[key] = kw_col
    for key, col in keyword.items():
        merged.setdefault(key, col)
    final: dict[str, int] = {}
    used: set[int] = set()
    for key in ("article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "cll", "no", "tariff_code"):
        col = merged.get(key)
        if col and col not in used:
            final[key] = col
            used.add(col)
    for key, col in merged.items():
        if key not in final and col not in used:
            final[key] = col
            used.add(col)
    final = _refine_mapping_with_samples(final, summary)
    final = _infer_missing_from_samples(final, summary.get("samples") or [])
    return final


def _sheet_looks_useful(summary: dict) -> bool:
    if _sheet_is_catalog(summary):
        return False
    if len(summary["headers"]) < 2 and not (summary.get("keyword_mapping") or {}).get("article"):
        # Still useful if wide samples have article-like codes
        samples = summary.get("samples") or []
        art_hits = 0
        for s in samples:
            for v in (s.get("cells") or {}).values():
                if _looks_like_article(v):
                    art_hits += 1
        if art_hits < 2:
            return False
    texts = " ".join(h["text"] for h in summary["headers"]).lower()
    markers = (
        "артикул", "наименование", "описание", "кол-во", "цена", "нетто", "брутто",
        "гросс", "итого", "article", "qty", "price", "gross", "net",
    )
    if any(m in texts for m in markers):
        return True
    km = summary.get("keyword_mapping") or {}
    return "article" in km or "qty" in km or "net_weight" in km


def map_all_sheets(summaries: list[dict]) -> list[dict]:
    usable = [s for s in summaries if _sheet_looks_useful(s)]
    if not usable:
        return []

    # If keyword mapping already covers article+name+qty for all sheets, skip AI
    def mapping_is_strong(s: dict) -> bool:
        km = s.get("keyword_mapping") or {}
        if not (
            {"article", "name"}.issubset(km)
            or {"article", "qty"}.issubset(km)
        ):
            return False
        art_col = km.get("article")
        samples = s.get("samples") or []
        if not art_col or not samples:
            return True
        vs = [(row.get("cells") or {}).get(str(art_col)) for row in samples]
        vs = [v for v in vs if v]
        if not vs:
            return True
        if sum(_looks_like_order_ref(v) for v in vs) > sum(_looks_like_sku(v) for v in vs):
            return False
        return True

    strong = all(mapping_is_strong(s) for s in usable)

    ai_by_sheet: dict[str, dict] = {}
    if not strong and yandex_gpt.is_configured():
        payload = {
            "sheets": [
                {
                    "sheet": s["sheet"],
                    "suggested_header_row": s["suggested_header_row"],
                    "header_rows": s.get("header_rows") or [s["suggested_header_row"]],
                    "headers": s["headers"],
                    "sample_cells": s["samples"],
                    "hint_keyword_mapping": s.get("keyword_mapping") or {},
                }
                for s in usable
            ]
        }
        reply = yandex_gpt.complete(
            SYSTEM_PROMPT,
            json.dumps(payload, ensure_ascii=False, indent=2),
            temperature=0.0,
            max_tokens=2000,
        )
        data = _parse_json_reply(reply)
        for item in _as_sheet_list(data):
            name = item.get("sheet")
            if name:
                ai_by_sheet[name] = item

    plans: list[dict] = []
    for summary in usable:
        item = ai_by_sheet.get(summary["sheet"], {})
        ai_map = _normalize_mapping(item.get("mapping") or {})
        mapping = _merge_mappings(summary.get("keyword_mapping") or {}, ai_map, summary)
        header_row = int(item.get("header_row") or summary["suggested_header_row"])
        data_start = int(item.get("data_start_row") or 0)
        if data_start <= header_row:
            data_start = 0
        plans.append({
            "sheet": summary["sheet"],
            "header_row": header_row,
            "data_start_row": data_start,  # resolved later with worksheet
            "mapping": mapping,
            "source": "keyword+ai" if ai_map else "keyword",
            "_summary": summary,
        })
    return plans


def _sheet_kind(name: str) -> str:
    norm = engine.normalize(name)
    compact = re.sub(r"[^a-zа-яё0-9]+", "", norm)
    if compact in {"pl", "pk", "packing", "packinglist", "пакинг", "упаковка", "упаковочныйлист"}:
        return "packing"
    if compact in {"inv", "invoice", "инвойс", "инв"}:
        return "invoice"
    if any(k in norm for k in ("pack", "пак", "пакинг", "pl(", "pl ")):
        return "packing"
    if any(k in norm for k in ("invoice", "инвойс", "инв", "in (", "in(")):
        return "invoice"
    return "other"


def _is_skip_row(rec: dict) -> bool:
    art = rec.get("article")
    name = rec.get("name")
    art_s = _cell_str(art).lower()
    name_s = _cell_str(name).lower()
    if art_s in HEADER_LIKE_ARTICLES or name_s in HEADER_LIKE_ARTICLES:
        return True
    if engine.is_summary_item(art, name):
        return True
    if art_s.startswith("the ") or "seller" in art_s:
        return True
    if "shipping marks" in art_s or "order no" in art_s:
        return True
    art_n = engine.normalize(art)
    if art_n in {"amount", "net to pay", "паллетизация", "palletization"}:
        return True
    if not art_s and name_s and not re.search(r"[A-Za-zА-Яа-яЁё0-9]", name_s):
        return True
    return False


def _to_number(val: Any) -> float | None:
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = _cell_str(val).replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _merge_field(target: dict, key: str, value: Any) -> None:
    if value in (None, ""):
        return
    if key in SUM_FIELDS:
        left = _to_number(target.get(key))
        right = _to_number(value)
        if left is not None and right is not None:
            total = left + right
            if abs(total - round(total)) < 1e-9:
                target[key] = int(round(total))
            else:
                target[key] = round(total, 4)
            return
        if target.get(key) in (None, ""):
            target[key] = value
        return
    if target.get(key) in (None, ""):
        target[key] = value


def _absorb_detail_rows(ws: Worksheet, item_row: int, rec: dict, mapping: dict[str, int]) -> None:
    """Invoice lines often put name / HS / country on the rows under the SKU."""
    art_col = mapping.get("article")
    qty_col = mapping.get("qty")
    max_c = min(ws.max_column or 12, 20)
    last = ws.max_row or item_row
    for r2 in range(item_row + 1, min(item_row + 6, last + 1)):
        art2 = ws.cell(r2, art_col).value if art_col else None
        qty2 = ws.cell(r2, qty_col).value if qty_col else None
        if (
            art2 not in (None, "")
            and (_looks_like_sku(art2) or (_looks_like_article(art2) and not _looks_like_name(art2)))
            and _to_number(qty2) is not None
        ):
            break
        if engine.is_summary_item(art2, None):
            break
        for c in range(1, max_c + 1):
            raw = ws.cell(r2, c).value
            if raw in (None, ""):
                continue
            text = _cell_str(raw)
            norm = engine.normalize(text)
            nxt = ws.cell(r2, c + 1).value if c < max_c else None
            if any(k in norm for k in ("таможенный код", "тн вэд", "hs code", "код тн", "тариф")):
                if rec.get("tariff_code") in (None, "") and nxt not in (None, ""):
                    rec["tariff_code"] = nxt
                continue
            if any(k in norm for k in ("страна происхождения", "country of origin", "origin country")):
                if rec.get("country") in (None, ""):
                    code = countries.resolve_country(nxt) or countries.resolve_country(text)
                    if code:
                        rec["country"] = code
                continue
            if rec.get("name") in (None, "") and _looks_like_name(raw) and not engine.is_summary_item(raw, None):
                rec["name"] = raw if isinstance(raw, str) else text


def extract_items_from_sheet(ws: Worksheet, plan: dict) -> dict[int, dict]:
    mapping: dict[str, int] = plan["mapping"]
    if not mapping or "article" not in mapping:
        # Without article column we cannot reliably merge shipment lines.
        return {}

    header_row = int(plan.get("header_row") or 1)
    start = int(plan.get("data_start_row") or 0)
    if start <= header_row:
        start = _guess_data_start_row(ws, header_row, mapping)
    max_row = ws.max_row or start
    no_col = mapping.get("no")
    items: dict[int, dict] = {}
    blank_streak = 0
    seq = 0

    for r in range(start, max_row + 1):
        rec: dict = {}
        non_empty = False
        for field, col in mapping.items():
            if field == "no":
                continue
            val = ws.cell(r, col).value
            if val not in (None, ""):
                # Prefer numeric coercion for numeric fields
                if field in SUM_FIELDS or field in {"price", "qty"}:
                    num = _to_number(val)
                    rec[field] = num if num is not None else val
                else:
                    rec[field] = val
                non_empty = True

        if rec.get("country") in (None, ""):
            found = countries.find_country_in_row(
                ws, r, skip_cols=mapping.values()
            )
            if found:
                rec["country"] = found
                non_empty = True
        if rec.get("unit") in (None, ""):
            found_u = units.find_unit_in_row(
                ws, r, skip_cols=mapping.values()
            )
            if found_u:
                rec["unit"] = found_u
                non_empty = True

        item_no = None
        if no_col:
            raw_no = ws.cell(r, no_col).value
            if isinstance(raw_no, (int, float)) and float(raw_no).is_integer() and raw_no > 0:
                item_no = int(raw_no)

        if not non_empty and item_no is None:
            blank_streak += 1
            if blank_streak >= 5:
                break
            continue
        blank_streak = 0

        if _is_skip_row(rec):
            continue

        art = rec.get("article")
        name = rec.get("name")
        if art in (None, "") and name in (None, ""):
            continue
        if art in (None, "") and isinstance(name, (int, float)):
            continue
        if art in (None, "") and _cell_str(name).isdigit():
            continue
        if mapping.get("qty") and rec.get("qty") in (None, "") and rec.get("price") in (None, ""):
            if rec.get("net_weight") in (None, "") and rec.get("gross_weight") in (None, ""):
                continue
        # Require a real article code when the column is mapped
        if art not in (None, "") and not _looks_like_article(art):
            # Allow if name exists and art looks like a code with spaces stripped oddly
            if not _looks_like_name(name):
                continue

        if item_no is None:
            seq += 1
            item_no = seq

        # Final per-row swap if values clearly inverted
        if art is not None and name is not None:
            if _looks_like_name(art) and _looks_like_article(name):
                rec["article"], rec["name"] = name, art

        if not any(rec.get(k) for k in engine.CERT_FIELDS):
            cert = engine.find_cert_in_row(ws, r)
            if cert:
                rec.update(cert)

        _absorb_detail_rows(ws, r, rec, mapping)
        if rec.get("amount") in (None, "") and rec.get("qty") not in (None, "") and rec.get("price") not in (None, ""):
            try:
                rec["amount"] = round(float(rec["qty"]) * float(rec["price"]), 4)
            except (TypeError, ValueError):
                pass

        items[item_no] = rec
        if len(items) >= MAX_EXTRACT_ROWS:
            break

    return items


def _merge_by_article_or_no(items_list: list[dict[int, dict]], *, sum_numeric: bool = True) -> dict[int, dict]:
    by_article: dict[str, dict] = {}
    order: list[str] = []
    no_article: list[dict] = []

    for items in items_list:
        for _no, rec in items.items():
            if _is_skip_row(rec):
                continue
            art = rec.get("article")
            key = engine.normalize(art) if art not in (None, "") else ""
            if key:
                if key not in by_article:
                    by_article[key] = dict(rec)
                    order.append(key)
                else:
                    target = by_article[key]
                    for k, v in rec.items():
                        if sum_numeric:
                            _merge_field(target, k, v)
                        elif target.get(k) in (None, ""):
                            target[k] = v
            else:
                no_article.append(dict(rec))

    merged: dict[int, dict] = {}
    i = 1
    for key in order:
        merged[i] = by_article[key]
        i += 1
    for rec in no_article:
        merged[i] = rec
        i += 1
    return merged


def _combine_invoice_and_packing(invoice: dict[int, dict], packing: dict[int, dict]) -> dict[int, dict]:
    """Attach packing weights to invoice lines; keep invoice qty/amount."""
    by_art: dict[str, dict] = {}
    order: list[str] = []
    for _no, rec in invoice.items():
        art = engine.normalize(rec.get("article"))
        if not art:
            continue
        by_art[art] = dict(rec)
        order.append(art)

    weight_fields = ("net_weight", "gross_weight", "cll")
    for _no, rec in packing.items():
        art = engine.normalize(rec.get("article"))
        if not art:
            continue
        if art in by_art:
            target = by_art[art]
            for f in weight_fields:
                if rec.get(f) not in (None, ""):
                    target[f] = rec[f]
            for f in engine.CERT_FIELDS:
                if target.get(f) in (None, "") and rec.get(f) not in (None, ""):
                    target[f] = rec[f]
            if target.get("name") in (None, "") and rec.get("name") not in (None, ""):
                target["name"] = rec["name"]
            for f in ("country", "unit", "manufacturer", "marks"):
                if target.get(f) in (None, "") and rec.get(f) not in (None, ""):
                    target[f] = rec[f]
        else:
            by_art[art] = dict(rec)
            order.append(art)

    return {i + 1: by_art[k] for i, k in enumerate(order)}


def _has_price(rec: dict) -> bool:
    return rec.get("price") not in (None, "") or rec.get("amount") not in (None, "")


def _has_weight(rec: dict) -> bool:
    return any(rec.get(k) not in (None, "") for k in ("net_weight", "gross_weight", "cll"))


def _partition_invoice_packing(items: dict[int, dict]) -> tuple[dict[int, dict], dict[int, dict]]:
    invoice: dict[int, dict] = {}
    packing: dict[int, dict] = {}
    i_inv = i_pk = 1
    for rec in items.values():
        rec = dict(rec)
        if _has_price(rec):
            invoice[i_inv] = rec
            i_inv += 1
        elif _has_weight(rec):
            packing[i_pk] = rec
            i_pk += 1
        else:
            invoice[i_inv] = rec
            i_inv += 1
    return invoice, packing


def _merge_ai_workbooks(all_items: list[dict[int, dict]]) -> dict[int, dict]:
    """Invoice files stay primary; packing only attaches weights (and leftover SKUs)."""
    inv_list: list[dict[int, dict]] = []
    pk_list: list[dict[int, dict]] = []
    for items in all_items:
        inv, pk = _partition_invoice_packing(items)
        if inv:
            inv_list.append(inv)
        if pk:
            pk_list.append(pk)
    invoice = _merge_by_article_or_no(inv_list, sum_numeric=True)
    packing = _merge_by_article_or_no(pk_list, sum_numeric=True)
    if invoice or packing:
        return _combine_invoice_and_packing(invoice, packing)
    return {}


def _fill_missing_countries(items: dict[int, dict]) -> None:
    codes = []
    for rec in items.values():
        c = rec.get("country")
        if c in (None, "") or isinstance(c, (list, dict, tuple)):
            continue
        codes.append(str(c))
    if not codes:
        return
    top = max(set(codes), key=codes.count)
    for rec in items.values():
        if rec.get("country") in (None, ""):
            rec["country"] = top


def extract_items_from_workbook_ai(content: bytes, filename: str | None = None) -> tuple[dict[int, dict], list[dict]]:
    wb = excel_io.load_workbook(content, data_only=True, filename=filename)
    summaries = [summarize_sheet(ws) for ws in wb.worksheets]
    try:
        plans = map_all_sheets(summaries)
    except Exception as e:
        plans = []
        for s in summaries:
            if not _sheet_looks_useful(s):
                continue
            mapping = _refine_mapping_with_samples(dict(s.get("keyword_mapping") or {}), s)
            mapping = _infer_missing_from_samples(mapping, s.get("samples") or [])
            plans.append({
                "sheet": s["sheet"],
                "header_row": s["suggested_header_row"],
                "data_start_row": 0,
                "mapping": mapping,
                "error": str(e),
                "source": "keyword_fallback",
            })

    # Resolve data_start_row against actual sheets; strip internal keys from report
    resolved_plans: list[dict] = []
    for plan in plans:
        p = {k: v for k, v in plan.items() if not k.startswith("_")}
        if p["sheet"] in wb.sheetnames:
            if not p.get("data_start_row"):
                p["data_start_row"] = _guess_data_start_row(
                    wb[p["sheet"]], int(p.get("header_row") or 1), p.get("mapping") or {}
                )
        resolved_plans.append(p)
    plans = resolved_plans

    invoice_list: list[dict[int, dict]] = []
    packing_list: list[dict[int, dict]] = []
    other_list: list[dict[int, dict]] = []

    for plan in plans:
        if not plan.get("mapping") or "article" not in plan["mapping"]:
            continue
        if plan["sheet"] not in wb.sheetnames:
            continue
        items = extract_items_from_sheet(wb[plan["sheet"]], plan)
        kind = _sheet_kind(plan["sheet"])
        # Content fallback: sheet with weights but no prices → packing
        if kind == "other":
            m = plan["mapping"]
            if ("net_weight" in m or "gross_weight" in m) and "price" not in m and "amount" not in m:
                kind = "packing"
            elif "price" in m or "amount" in m:
                kind = "invoice"
        if kind == "packing":
            packing_list.append(items)
        elif kind == "invoice":
            invoice_list.append(items)
        else:
            other_list.append(items)

    invoice = _merge_by_article_or_no(invoice_list, sum_numeric=True)
    packing = _merge_by_article_or_no(packing_list, sum_numeric=True)
    other = _merge_by_article_or_no(other_list, sum_numeric=True)

    if invoice or packing:
        merged = _combine_invoice_and_packing(invoice, packing)
    else:
        merged = other
        other = {}

    if other:
        # merge remaining sheets without double-counting invoice qty
        extra = _merge_by_article_or_no([merged, other], sum_numeric=False)
        merged = extra if extra else merged

    doc_country = countries.find_document_country(wb)
    if doc_country:
        for rec in merged.values():
            if rec.get("country") in (None, ""):
                rec["country"] = doc_country

    return merged, plans


def process_with_ai(
    template_bytes: bytes,
    input_files: list[bytes],
    settings: dict | None = None,
    filenames: list[str] | None = None,
) -> tuple[bytes, dict]:
    settings = settings or {}
    filenames = filenames or [None] * len(input_files)

    all_items: list[dict[int, dict]] = []
    all_plans: list[dict] = []
    for content, fname in zip(input_files, filenames):
        try:
            items, plans = extract_items_from_workbook_ai(content, filename=fname)
            all_items.append(items)
            all_plans.extend(plans)
        except Exception as e:
            all_plans.append({
                "sheet": fname or "file",
                "mapping": {},
                "error": str(e).strip() or e.__class__.__name__,
                "source": "load",
            })

    merged = _merge_ai_workbooks(all_items)
    if not merged:
        errors = [p.get("error") for p in all_plans if p.get("error")]
        detail = errors[0] if errors else "ИИ не смог сопоставить столбцы"
        raise RuntimeError(detail)

    _fill_missing_countries(merged)

    try:
        result_bytes = engine.fill_template(template_bytes, merged, settings)
    except Exception as e:
        raise RuntimeError(
            "Не удалось заполнить шаблон: " + (str(e).strip() or e.__class__.__name__)
        ) from e
    mode = "ai"
    try:
        protocol = engine.build_fill_protocol(
            merged, settings=settings, mappings=all_plans, mode=mode
        )
    except Exception:
        protocol = {
            "summary": f"В шаблон перенесено {len(merged)} позиций.",
            "errors": [],
            "warnings": [],
            "notes": [],
            "items_found": len(merged),
        }
    report = {
        "mode": mode,
        "items_found": len(merged),
        "protocol": protocol,
        "mappings": [
            {
                "sheet": p.get("sheet"),
                "header_row": p.get("header_row"),
                "mapping": p.get("mapping", {}),
                "source": p.get("source"),
                "error": p.get("error"),
            }
            for p in all_plans
        ],
        "items": {
            no: {k: (str(v) if hasattr(v, "isoformat") else v) for k, v in rec.items()}
            for no, rec in sorted(merged.items())
        },
    }
    return result_bytes, report
