"""AI-assisted column mapping for Excel → template fields via YandexGPT."""
from __future__ import annotations

import io
import json
import re
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

import engine
import yandex_gpt

CANONICAL_FIELDS = [
    "no", "tariff_code", "name", "article", "marks", "manufacturer",
    "country", "unit", "qty", "price", "amount", "net_weight",
    "gross_weight", "cll", "mnr", "date_from", "date_to", "mnr_code",
]

FIELD_HINTS = {
    "no": "порядковый номер позиции (№, No.)",
    "tariff_code": "код ТН ВЭД / HS / tariff code",
    "name": "наименование товара / description",
    "article": "артикул / article / SKU",
    "marks": "торговая марка / brand",
    "manufacturer": "производитель / manufacturer",
    "country": "страна происхождения",
    "unit": "единица измерения (шт, kg…)",
    "qty": "количество в единицах измерения",
    "price": "цена за единицу",
    "amount": "стоимость / итого / сумма строки",
    "net_weight": "вес нетто",
    "gross_weight": "вес брутто / gross",
    "cll": "количество мест / паллет / packages",
    "mnr": "номер сертификата/декларации",
    "date_from": "дата начала действия сертификата",
    "date_to": "дата окончания действия сертификата",
    "mnr_code": "код вида документа МНР",
}

SYSTEM_PROMPT = """Ты помощник по таможенным Excel-документам.
Тебе дают несколько листов: заголовки столбцов и примеры строк.
Для КАЖДОГО листа сопоставь столбцы с каноническими полями шаблона.

Канонические поля (используй только эти ключи):
""" + "\n".join(f"- {k}: {FIELD_HINTS[k]}" for k in CANONICAL_FIELDS) + """

Правила:
- Сопоставляй по смыслу, даже если названия отличаются (например «Нетто вес» → net_weight, «Итого» → amount, «Кол-во» → qty, «Артикул» → article).
- Если столбец — только единица валюты/измерения (CNY, шт, кг) без данных — не включай.
- Если подходящего столбца нет — не указывай поле.
- header_row и data_start_row — 1-based номера строк.
- Ответь ТОЛЬКО JSON без markdown:
{"sheets":[{"sheet":"имя листа","header_row":N,"data_start_row":M,"mapping":{"article":2,"name":3,"qty":4}}]}
где значения mapping — 1-based номера столбцов Excel.
"""


def _cell_str(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).replace("\t", " ").replace("\n", " ").strip()


def _find_candidate_header_rows(ws: Worksheet, max_scan: int = 50) -> list[int]:
    """Rank rows that look like product-table headers (keyword hits preferred)."""
    max_row = min(ws.max_row or 1, max_scan)
    max_col = min(ws.max_column or 1, 30)
    scored: list[tuple[int, int, int]] = []  # score, non_empty, row
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
        # Prefer keyword-rich rows strongly over sparse metadata rows
        score = hits * 10 + len(texts)
        if hits >= 2 or len(texts) >= 4:
            scored.append((score, len(texts), r))
    scored.sort(reverse=True)
    rows = [r for _s, _n, r in scored[:5]]
    return rows or [1]


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
    for r in range(best + 1, min(best + 6, (ws.max_row or best) + 1)):
        row = [_cell_str(ws.cell(r, c).value) for c in range(1, max_col + 1)]
        if any(row):
            samples.append({"row": r, "values": row})
    return {
        "sheet": ws.title,
        "header_candidates": header_candidates,
        "suggested_header_row": best,
        "headers": headers,
        "samples": samples,
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
    if start_arr >= 0 and (start_obj < 0 or start_arr < start_obj):
        start = start_arr
    else:
        start = start_obj
    decoder = json.JSONDecoder()
    data, _end = decoder.raw_decode(raw[start:])
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
    for key, col in (raw_mapping or {}).items():
        if key in CANONICAL_FIELDS and isinstance(col, int) and col > 0:
            mapping[key] = col
    return mapping


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

    payload = {
        "sheets": [
            {
                "sheet": s["sheet"],
                "header_candidates": s["header_candidates"],
                "headers": s["headers"],
                "sample_rows": s["samples"],
            }
            for s in usable
        ]
    }
    reply = yandex_gpt.complete(
        SYSTEM_PROMPT,
        json.dumps(payload, ensure_ascii=False, indent=2),
        temperature=0.05,
        max_tokens=4000,
    )
    data = _parse_json_reply(reply)
    by_name = {s["sheet"]: s for s in usable}
    plans: list[dict] = []
    for item in _as_sheet_list(data):
        name = item.get("sheet")
        summary = by_name.get(name)
        if not summary:
            continue
        plans.append({
            "sheet": name,
            "header_row": int(item.get("header_row") or summary["suggested_header_row"]),
            "data_start_row": int(item.get("data_start_row") or (summary["suggested_header_row"] + 1)),
            "mapping": _normalize_mapping(item.get("mapping") or {}),
        })
    return plans


def extract_items_from_sheet(ws: Worksheet, plan: dict) -> dict[int, dict]:
    mapping: dict[str, int] = plan["mapping"]
    if not mapping:
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
            if blank_streak >= 3:
                break
            continue
        blank_streak = 0

        if item_no is None:
            if not (rec.get("article") or rec.get("name") or rec.get("qty") or rec.get("amount")):
                continue
            seq += 1
            item_no = seq

        if not any(rec.get(k) for k in engine.CERT_FIELDS):
            cert = engine.find_cert_in_row(ws, r)
            if cert:
                rec.update(cert)

        items[item_no] = rec

    return items


def _merge_by_article_or_no(items_list: list[dict[int, dict]]) -> dict[int, dict]:
    by_article: dict[str, dict] = {}
    order: list[str] = []
    no_article: list[dict] = []

    for items in items_list:
        for _no, rec in items.items():
            art = rec.get("article")
            key = engine.normalize(art) if art not in (None, "") else ""
            if key:
                if key not in by_article:
                    by_article[key] = dict(rec)
                    order.append(key)
                else:
                    target = by_article[key]
                    for k, v in rec.items():
                        if target.get(k) in (None, "") and v not in (None, ""):
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
    return merged if merged else engine.merge_workbooks(items_list)


def extract_items_from_workbook_ai(content: bytes) -> tuple[dict[int, dict], list[dict]]:
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    summaries = [summarize_sheet(ws) for ws in wb.worksheets]
    try:
        plans = map_all_sheets(summaries)
    except Exception as e:
        plans = [{"sheet": s["sheet"], "error": str(e), "mapping": {}} for s in summaries]

    items_list: list[dict[int, dict]] = []
    for plan in plans:
        if not plan.get("mapping") or plan.get("error"):
            continue
        if plan["sheet"] not in wb.sheetnames:
            continue
        items_list.append(extract_items_from_sheet(wb[plan["sheet"]], plan))

    return _merge_by_article_or_no(items_list), plans


def process_with_ai(
    template_bytes: bytes,
    input_files: list[bytes],
    settings: dict | None = None,
) -> tuple[bytes, dict]:
    settings = settings or {}
    if not yandex_gpt.is_configured():
        raise RuntimeError("YandexGPT не настроен")

    all_items: list[dict[int, dict]] = []
    all_plans: list[dict] = []
    for content in input_files:
        items, plans = extract_items_from_workbook_ai(content)
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
