"""AI-assisted column mapping for Excel → template fields via YandexGPT."""
from __future__ import annotations

import json
import re
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

import engine
import excel_io
import yandex_gpt

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

# Exact-ish header → field (checked before AI). Longer phrases first.
HEADER_EXACT = [
    ("наименование товара", "name"),
    ("name of product", "name"),
    ("наименование", "name"),
    ("описание", "name"),
    ("description", "name"),
    ("артикул", "article"),
    ("article", "article"),
    ("sku", "article"),
    ("код товара", "tariff_code"),
    ("тн вэд", "tariff_code"),
    ("tariff", "tariff_code"),
    ("кол-во мест", "cll"),
    ("количество мест", "cll"),
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
    ("производитель", "manufacturer"),
    ("manufacturer", "manufacturer"),
    ("торговая марка", "marks"),
    ("brand", "marks"),
]

ARTICLE_RE = re.compile(r"^[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._\-/]{1,40}$")
SKIP_ROW_MARKERS = (
    "всего", "итого", "total", "сумма:", "subtotal", "grand total",
    "sum", "the seller", "seller:", "buyer", "consignee", "паллет всего",
    "всес паллет", "всего паллет", "境内货源地",
)
SUM_FIELDS = frozenset({"qty", "amount", "net_weight", "gross_weight", "cll"})

SYSTEM_PROMPT = """Ты помощник по таможенным Excel-документам.
Тебе дают листы: заголовки с номерами столбцов (col) и примеры ячеек {col: значение}.

Сопоставь столбцы с каноническими полями. Ключи полей:
""" + "\n".join(f"- {k}: {FIELD_HINTS[k]}" for k in CANONICAL_FIELDS) + """

ЖЁСТКИЕ ПРАВИЛА:
1) «Артикул» / Article / SKU → ТОЛЬКО article. Туда короткие коды (E2903-00), НЕ описание товара.
2) «Наименование» / Name / Description → ТОЛЬКО name. Туда длинный текст («ванна …»), НЕ артикул.
3) Нельзя назначить один столбец двум полям.
4) Смотри примеры значений: если в столбце коды без пробелов — это article; если фразы — name.
5) Столбцы только с «шт», «кг», «CNY», «USD» не мапь.
6) Ответ — ТОЛЬКО JSON-массив (без markdown):
[{"sheet":"имя","header_row":N,"data_start_row":M,"mapping":{"article":2,"name":3,"qty":4}}]
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
    return bool(ARTICLE_RE.match(s)) and not s.replace(".", "").isdigit()


def _looks_like_name(val: Any) -> bool:
    s = _cell_str(val)
    if not s:
        return False
    return (" " in s and len(s) >= 8) or len(s) >= 20


def _find_candidate_header_rows(ws: Worksheet, max_scan: int = 50) -> list[int]:
    max_row = min(ws.max_row or 1, max_scan)
    max_col = min(ws.max_column or 1, 30)
    scored: list[tuple[int, int, int]] = []
    keyword_bags = [engine.HEADER_KEYWORDS[k] for k in (
        "article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "no", "tariff_code"
    )]
    for r in range(1, max_row + 1):
        texts = []
        hits = 0
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str) or not v.strip():
                continue
            text = v.strip()
            texts.append(text)
            norm = engine.normalize(text)
            if any(any(kw in norm for kw in bag) for bag in keyword_bags):
                hits += 1
        if len(texts) < 2:
            continue
        score = hits * 10 + len(texts)
        if hits >= 2 or len(texts) >= 4:
            scored.append((score, len(texts), r))
    scored.sort(reverse=True)
    return [r for _s, _n, r in scored[:5]] or [1]


def _keyword_mapping_from_headers(headers: list[dict]) -> dict[str, int]:
    """Deterministic mapping from clear header labels."""
    mapping: dict[str, int] = {}
    used_cols: set[int] = set()
    for h in headers:
        norm = engine.normalize(h["text"])
        if not norm:
            continue
        for phrase, field in HEADER_EXACT:
            if phrase in norm and field not in mapping and h["col"] not in used_cols:
                # Avoid matching bare "марка" inside unrelated words if needed — phrase list is ordered
                mapping[field] = h["col"]
                used_cols.add(h["col"])
                break
    return mapping


def summarize_sheet(ws: Worksheet) -> dict:
    max_col = min(ws.max_column or 1, 25)
    header_candidates = _find_candidate_header_rows(ws)
    best = header_candidates[0]
    headers = []
    for c in range(1, max_col + 1):
        text = _cell_str(ws.cell(best, c).value)
        if text:
            headers.append({"col": c, "text": text})

    samples = []
    for r in range(best + 1, min(best + 8, (ws.max_row or best) + 1)):
        cells = {}
        for h in headers:
            val = _cell_str(ws.cell(r, h["col"]).value)
            if val:
                cells[str(h["col"])] = val
        if cells:
            samples.append({"row": r, "cells": cells})

    return {
        "sheet": ws.title,
        "header_candidates": header_candidates,
        "suggested_header_row": best,
        "headers": headers,
        "samples": samples,
        "keyword_mapping": _keyword_mapping_from_headers(headers),
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
    """Fix article/name swaps using sample cell values."""
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

    art_col = mapping.get("article")
    name_col = mapping.get("name")

    if art_col and name_col:
        art_vals = sample_vals(art_col)
        name_vals = sample_vals(name_col)
        art_as_code = sum(_looks_like_article(v) for v in art_vals)
        art_as_name = sum(_looks_like_name(v) for v in art_vals)
        name_as_code = sum(_looks_like_article(v) for v in name_vals)
        name_as_name = sum(_looks_like_name(v) for v in name_vals)
        # Swapped: article column holds phrases, name column holds codes
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

    return mapping


def _merge_mappings(keyword: dict[str, int], ai: dict[str, int], summary: dict) -> dict[str, int]:
    """Keyword mapping wins for article/name/qty/…; AI fills gaps."""
    merged = dict(ai)
    # Prefer deterministic headers for critical identity fields
    for key in ("article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "no", "tariff_code"):
        if key in keyword:
            # free AI col if conflict
            ai_col = merged.get(key)
            kw_col = keyword[key]
            if ai_col and ai_col != kw_col:
                for k, c in list(merged.items()):
                    if c == kw_col and k != key:
                        del merged[k]
            merged[key] = kw_col
    for key, col in keyword.items():
        merged.setdefault(key, col)
    # ensure unique columns
    final: dict[str, int] = {}
    used: set[int] = set()
    # critical fields first
    for key in ("article", "name", "qty", "price", "amount", "net_weight", "gross_weight", "no", "tariff_code"):
        col = merged.get(key)
        if col and col not in used:
            final[key] = col
            used.add(col)
    for key, col in merged.items():
        if key not in final and col not in used:
            final[key] = col
            used.add(col)
    return _refine_mapping_with_samples(final, summary)


def _sheet_looks_useful(summary: dict) -> bool:
    if len(summary["headers"]) < 2:
        return False
    texts = " ".join(h["text"] for h in summary["headers"]).lower()
    markers = ("артикул", "наименование", "кол-во", "цена", "нетто", "брутто", "гросс", "итого", "article", "qty", "price")
    return any(m in texts for m in markers)


def map_all_sheets(summaries: list[dict]) -> list[dict]:
    usable = [s for s in summaries if _sheet_looks_useful(s)]
    if not usable:
        return []

    # If keyword mapping already covers article+name+qty for all sheets, skip AI
    strong = all(
        {"article", "name"}.issubset(s.get("keyword_mapping") or {})
        or {"article", "qty"}.issubset(s.get("keyword_mapping") or {})
        for s in usable
    )

    ai_by_sheet: dict[str, dict] = {}
    if not strong and yandex_gpt.is_configured():
        payload = {
            "sheets": [
                {
                    "sheet": s["sheet"],
                    "suggested_header_row": s["suggested_header_row"],
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
            max_tokens=4000,
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
        plans.append({
            "sheet": summary["sheet"],
            "header_row": int(item.get("header_row") or summary["suggested_header_row"]),
            "data_start_row": int(
                item.get("data_start_row")
                or (summary["suggested_header_row"] + 1)
            ),
            "mapping": mapping,
            "source": "keyword+ai" if ai_map else "keyword",
        })
    return plans


def _sheet_kind(name: str) -> str:
    norm = engine.normalize(name)
    if any(k in norm for k in ("pack", "пак", "пакинг", "pl(", "pl ")):
        return "packing"
    if any(k in norm for k in ("invoice", "инвойс", "инв", "in (", "in(")):
        return "invoice"
    return "other"


def _is_skip_row(rec: dict) -> bool:
    art = _cell_str(rec.get("article")).lower()
    name = _cell_str(rec.get("name")).lower()
    if art in {"sum", "total"} or name in {"sum", "total"}:
        return True
    if art.startswith("the ") or "seller" in art:
        return True
    if "паллет" in art or "паллет" in name:
        return True
    if "вес паллет" in art or "вес паллет" in name:
        return True
    blob = " ".join(_cell_str(v).lower() for v in rec.values())
    if any(m in blob for m in SKIP_ROW_MARKERS):
        return True
    if not art and name and not re.search(r"[A-Za-zА-Яа-яЁё0-9]", name):
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
            target[key] = int(total) if total.is_integer() else total
            return
        if target.get(key) in (None, ""):
            target[key] = value
        return
    if target.get(key) in (None, ""):
        target[key] = value


def extract_items_from_sheet(ws: Worksheet, plan: dict) -> dict[int, dict]:
    mapping: dict[str, int] = plan["mapping"]
    if not mapping or "article" not in mapping:
        # Without article column we cannot reliably merge shipment lines.
        return {}

    start = plan["data_start_row"]
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
                rec[field] = val
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

        items[item_no] = rec

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
                    # packing already summed across containers
                    target[f] = rec[f]
            if target.get("name") in (None, "") and rec.get("name") not in (None, ""):
                target["name"] = rec["name"]
        else:
            by_art[art] = dict(rec)
            order.append(art)

    return {i + 1: by_art[k] for i, k in enumerate(order)}


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
            plans.append({
                "sheet": s["sheet"],
                "header_row": s["suggested_header_row"],
                "data_start_row": s["suggested_header_row"] + 1,
                "mapping": mapping,
                "error": str(e),
                "source": "keyword_fallback",
            })

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
        if kind == "packing":
            packing_list.append(items)
        elif kind == "invoice":
            invoice_list.append(items)
        else:
            other_list.append(items)

    invoice = _merge_by_article_or_no(invoice_list, sum_numeric=True)
    packing = _merge_by_article_or_no(packing_list, sum_numeric=True)
    other = _merge_by_article_or_no(other_list, sum_numeric=True)

    merged = _combine_invoice_and_packing(invoice, packing)
    if other:
        # merge remaining sheets without double-counting invoice qty
        extra = _merge_by_article_or_no([merged, other], sum_numeric=False)
        merged = extra if extra else merged

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
        items, plans = extract_items_from_workbook_ai(content, filename=fname)
        all_items.append(items)
        all_plans.extend(plans)

    merged = _merge_by_article_or_no(all_items)
    if not merged:
        errors = [p.get("error") for p in all_plans if p.get("error")]
        detail = errors[0] if errors else "ИИ не смог сопоставить столбцы"
        raise RuntimeError(detail)

    result_bytes = engine.fill_template(template_bytes, merged, settings)
    report = {
        "mode": "ai",
        "items_found": len(merged),
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
