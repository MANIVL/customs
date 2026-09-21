"""
Excel mapping engine.

Reads one or more "input" Excel workbooks (Invoice / Packing list / Specification,
with sheet names and header positions that may vary between shipments) and fills a
fixed "template" workbook whose column headers must not be changed.

Core ideas:
- Headers are located by keyword search (Russian + English), not by fixed
  row/column numbers, so the input layout can shift between files.
- Item rows are found by locating the "No. / №" column and reading a
  contiguous run of positive integers below the header row.
- Data for a given item number is merged across all sheets that mention it
  (e.g. Net weight might live in "Инвойс" and/or "Спецификация" and/or
  "Пакинг" - first sheet in priority order that has a non-empty value wins).
- Certificate/declaration text ("... от DD.MM.YYYY действует до DD.MM.YYYY")
  is located anywhere in an item's row (any sheet, any column) and parsed
  into number / date-from / date-to / code fields.
- Country of origin is read from the files (column, row, or header area) and
  mapped to ISO 3166-1 alpha-2 via the OКСМ classifier (Китай → CN).
"""

from __future__ import annotations

import re
import io
from dataclasses import dataclass, field
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment
import datetime
import excel_io
import countries
import units


# --------------------------------------------------------------------------
# Canonical field definitions
# --------------------------------------------------------------------------

# Keyword fragments (lower-cased, whitespace-normalized) used to recognize a
# header cell for each canonical field. A header cell matches a field if any
# of its keywords is a substring of the normalized cell text.
HEADER_KEYWORDS: dict[str, list[str]] = {
    "no": ["№", "no.", "no /"],
    "tariff_code": ["tariff code", "код товара", "тн вэд", "hs code"],
    "name": ["name of product", "наименование товара", "наименование", "описание", "description"],
    "article": ["article", "артикул", "sku"],
    "marks": ["marks /", "торговая марка", "бренд", "brand"],
    "country": [
        "country of origin", "страна происхождения товара", "страна происхождения",
        "страна-производитель", "страна производитель", "origin country",
        "place of origin", "made in", "происхождение", "страна", "country",
        "原产国", "原产地",
    ],
    "manufacturer": ["manufacturer", "производитель"],
    "unit": [
        "unit /", "единица измерения", "единицы измерения", "unit of measure",
        "uom",
    ],
    "qty": ["q-ty", "кол-во в ед", "кол-во", "количество", "qty"],
    "price": ["price usd", "цена долл", "цена", "price"],
    "amount": ["amount, usd", "стоимость, долл", "стоимость", "итого", "сумма", "amount"],
    "net_weight": ["net weight", "вес нетто", "вес  нетто", "нетто вес", "нетто"],
    "gross_weight": ["gross weight", "вес брутто", "вес  брутто", "гросс вес", "брутто", "gross"],
    "cll": ["number of units load", "кол-во мест", "паллет", "пал"],
}

# Fields that get written into the template's certificate/declaration block.
# A row can carry up to 2 certificates (мнр / мнр 2 columns in the template).
CERT_FIELDS = ["mnr", "date_from", "date_to", "mnr_code", "mnr2", "date_from2", "date_to2", "mnr_code2"]

# Order in which sheets are consulted when the same field appears in more
# than one place. Matched case-insensitively / by substring against the
# actual sheet name.
DEFAULT_SHEET_PRIORITY = ["инвойс", "invoice", "спецификация", "specification", "пакинг", "packing"]

# Hard-coded certificate-number-prefix -> "код мнр" rules.
MNR_CODE_RULES: list[tuple[str, str]] = [
    ("еаэс n ru д-", "01402"),
    ("еаэс ru с-", "01401"),
    ("росс ru д-", "01408"),
]

CERT_RE = re.compile(
    r"(?P<num>[A-ZА-ЯЁ0-9№][A-ZА-ЯЁ0-9№\-./ ]*?)\s*от\s*(?P<from>\d{2}\.\d{2}\.\d{4})\s*"
    r"действ\w*\s*до\s*(?P<to>\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE | re.UNICODE,
)

# Fallback pattern for certificate-like numbers that only carry an issue date
# and no expiry (e.g. "свидетельство о гос. регистрации": "KG.11.01.09.012.R.006517.10.25 от 29.10.2025").
# Only tried when CERT_RE (with "действует до") does not match a given cell.
CERT_RE_SIMPLE = re.compile(
    r"(?P<num>[A-ZА-ЯЁ0-9№][A-ZА-ЯЁ0-9№\-./ ]*?)\s*от\s*(?P<from>\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE | re.UNICODE,
)


def normalize(text: Any) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def header_keyword_matches(keyword: str, norm: str) -> bool:
    """Match a header keyword in normalized cell text.

    Short prefixes (≤3 chars) stay substring matches («пал» → «паллет»).
    Longer single tokens use word boundaries so «article» does not hit
    «ARTICLES CO., LTD» in a data cell mistaken for a header.
    """
    if not keyword or not norm:
        return False
    if len(keyword) <= 3 or any(ch in keyword for ch in " /."):
        return keyword in norm
    return re.search(
        r"(?<![0-9a-zа-яё])" + re.escape(keyword) + r"(?![0-9a-zа-яё])",
        norm,
    ) is not None


def any_header_keyword(keywords: list[str], norm: str) -> bool:
    return any(header_keyword_matches(kw, norm) for kw in keywords)


_SUMMARY_HEADS = (
    "всего", "итого", "total", "sum", "subtotal", "grand total",
    "паллет", "паллет всего", "всего паллет",
)


def is_summary_label(text: Any) -> bool:
    """True for a totals cell, not a product description that mentions «всего 12 кв.м»."""
    s = normalize(text).strip(" .:;-")
    if not s:
        return False
    if s in _SUMMARY_HEADS:
        return True
    if len(s) > 48:
        return False
    return s.startswith(_SUMMARY_HEADS)


def is_summary_item(article: Any, name: Any) -> bool:
    """Skip invoice total/pallet lines; keep goods whose description contains «всего»."""
    art = str(article or "").strip()
    art_l = art.lower()
    if is_summary_label(art):
        return True
    if art_l.startswith(("shipping", "packing", "order no", "measurement")):
        return True
    if art_l.startswith("container") or (len(art_l) < 40 and "контейнер" in art_l):
        return True
    if art_l in {"артикул", "article", "sku", "шт", "шт."}:
        return True
    name_s = str(name or "").strip()
    if name_s and len(name_s) < 48 and is_summary_label(name_s):
        return True
    return False


def parse_date(text: str) -> datetime.datetime | None:
    try:
        return datetime.datetime.strptime(text.strip(), "%d.%m.%Y")
    except Exception:
        return None


def mnr_code_for(cert_number: str) -> str | None:
    norm = normalize(cert_number)
    for prefix, code in MNR_CODE_RULES:
        if norm.startswith(prefix):
            return code
    # Anything that isn't a ЕАЭС/РОСС certificate or declaration (e.g. a
    # свидетельство о гос. регистрации number) is classified as 01206.
    if not norm.startswith("еаэс") and not norm.startswith("росс"):
        return "01206"
    return None


# --------------------------------------------------------------------------
# Header / item-table detection
# --------------------------------------------------------------------------

@dataclass
class SheetTable:
    sheet_name: str
    header_row: int
    columns: dict[str, int]  # canonical field -> 1-based column index
    item_rows: dict[int, int]  # item number -> 1-based row index


def find_no_header(ws: Worksheet, max_scan_rows: int = 60, max_scan_cols: int = 40):
    """Locate the header row/col of the item-number ("No. / №") column by
    finding a header-like cell followed by a run of ascending integers."""
    max_row = min(ws.max_row, max_scan_rows)
    max_col = min(ws.max_column, max_scan_cols)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = ws.cell(row=r, column=c).value
            norm = normalize(val)
            if not norm:
                continue
            if any(k in norm for k in HEADER_KEYWORDS["no"]):
                # check next 1-3 rows for integer 1 (or any small positive int)
                for probe in range(1, 4):
                    below = ws.cell(row=r + probe, column=c).value
                    if isinstance(below, (int, float)) and float(below).is_integer() and below > 0:
                        return r, c
    return None, None


def build_sheet_table(ws: Worksheet) -> SheetTable | None:
    header_row, no_col = find_no_header(ws)
    if header_row is None:
        # Fallback: locate a header row by known field keywords (no № column).
        header_row, no_col = _find_header_without_no(ws)
    if header_row is None:
        return None

    max_col = ws.max_column
    columns: dict[str, int] = {}
    if no_col:
        columns["no"] = no_col

    def _row_looks_like_items(r: int) -> bool:
        """True when the row already has an item number — do not treat as a header."""
        if not no_col:
            return False
        val = ws.cell(row=r, column=no_col).value
        return isinstance(val, (int, float)) and float(val).is_integer() and val > 0

    def absorb_header_row(r: int) -> None:
        for c in range(1, max_col + 1):
            norm = normalize(ws.cell(row=r, column=c).value)
            if not norm:
                continue
            for fkey, keywords in HEADER_KEYWORDS.items():
                if not any_header_keyword(keywords, norm):
                    continue
                # «Price USD / unit / …» is a price header, not the unit column
                if fkey == "unit" and any_header_keyword(HEADER_KEYWORDS["price"], norm):
                    continue
                # Prefer article over a generic "описание товара" on the same column
                if c in columns.values() and fkey != "article":
                    continue
                if fkey in columns and fkey != "article":
                    continue
                if fkey == "article":
                    # displace name if it occupied this column
                    for k, col in list(columns.items()):
                        if col == c and k != "article":
                            del columns[k]
                    columns["article"] = c
                elif fkey not in columns:
                    columns[fkey] = c
                break

    absorb_header_row(header_row)
    # Two-line headers: «Описание товара / Кол-во / Цена» above «артикул | описание»
    for probe in (header_row - 2, header_row - 1, header_row + 1, header_row + 2):
        if probe < 1 or probe > (ws.max_row or probe):
            continue
        if _row_looks_like_items(probe):
            continue
        absorb_header_row(probe)
    if "article" in columns and columns.get("name") == columns.get("article"):
        for probe in (header_row, header_row + 1, header_row + 2):
            if probe > (ws.max_row or probe):
                break
            if _row_looks_like_items(probe):
                continue
            for c in range(1, max_col + 1):
                norm = normalize(ws.cell(row=probe, column=c).value)
                if any_header_keyword(HEADER_KEYWORDS["name"], norm) and c != columns["article"]:
                    columns["name"] = c
                    break
            if columns.get("name") != columns.get("article"):
                break

    # «Цена» / «Итого» often label the currency sub-column; shift to numeric neighbor.
    def _col_is_currency(c: int, probe_from: int) -> bool:
        hits = 0
        nums = 0
        for rr in range(probe_from, min(probe_from + 8, (ws.max_row or probe_from) + 1)):
            val = ws.cell(row=rr, column=c).value
            if val in (None, ""):
                continue
            s = str(val).strip().upper()
            if s in {"USD", "EUR", "CNY", "RMB", "RUR", "RUB", "GBP", "$", "€", "¥"}:
                hits += 1
            else:
                try:
                    float(str(val).replace(",", ".").replace(" ", ""))
                    nums += 1
                except Exception:
                    pass
        return hits > nums and hits > 0

    data_probe = header_row + 1
    for field in ("price", "amount"):
        col = columns.get(field)
        if not col or not _col_is_currency(col, data_probe):
            continue
        for cand in (col + 1, col + 2, col - 1):
            if cand < 1 or cand in columns.values():
                continue
            if not _col_is_currency(cand, data_probe):
                # ensure it has numbers
                has_num = False
                for rr in range(data_probe, min(data_probe + 8, (ws.max_row or data_probe) + 1)):
                    val = ws.cell(row=rr, column=cand).value
                    if isinstance(val, (int, float)):
                        has_num = True
                        break
                if has_num:
                    columns[field] = cand
                    break

    # Infer article/name columns from values when headers omit them (common on PL).
    if "article" not in columns:
        best_c = None
        best_hits = 0
        unit_like = {"шт", "шт.", "pcs", "pc", "kg", "кг", "кор", "кор.", "ctn", "rmb", "usd", "cny"}
        for c in range(1, min(max_col, 15) + 1):
            owner = next((k for k, col in columns.items() if col == c), None)
            if owner and owner not in {"name"}:  # may displace generic name group-header
                continue
            hits = 0
            for rr in range(header_row + 1, min(header_row + 12, (ws.max_row or header_row) + 1)):
                val = ws.cell(row=rr, column=c).value
                if not isinstance(val, str):
                    continue
                s = val.strip()
                if not s or " " in s or len(s) < 4 or len(s) > 40:
                    continue
                if s.lower().rstrip(".") in unit_like:
                    continue
                if s.lower() in {"артикул", "article", "sku"}:
                    continue
                if re.match(r"^[A-Za-zА-Яа-яЁё0-9][A-Za-zА-Яа-яЁё0-9._\-/]{2,40}$", s) and not s.replace(".", "").isdigit():
                    hits += 1
            if hits > best_hits:
                best_hits = hits
                best_c = c
        if best_c and best_hits >= 2:
            for k, col in list(columns.items()):
                if col == best_c and k != "article":
                    del columns[k]
            columns["article"] = best_c
    if "name" not in columns and "article" in columns:
        cand = columns["article"] + 1
        if cand not in columns.values():
            name_hits = 0
            for rr in range(header_row + 1, min(header_row + 12, (ws.max_row or header_row) + 1)):
                val = ws.cell(row=rr, column=cand).value
                if isinstance(val, str) and " " in val.strip() and len(val.strip()) >= 8:
                    name_hits += 1
            if name_hits >= 2:
                columns["name"] = cand
    # If name was wrongly assigned to article-like col 2 before article inference
    if "article" in columns and "name" in columns and columns["name"] < columns["article"]:
        # swap if values suggest it
        art_c, name_c = columns["article"], columns["name"]
        art_hits = name_hits = 0
        for rr in range(header_row + 1, min(header_row + 8, (ws.max_row or header_row) + 1)):
            a = ws.cell(row=rr, column=art_c).value
            n = ws.cell(row=rr, column=name_c).value
            if isinstance(a, str) and " " not in a.strip() and len(a.strip()) >= 4:
                art_hits += 1
            if isinstance(n, str) and " " in n.strip():
                name_hits += 1
        if art_hits < 2 and name_hits >= 2:
            columns["article"], columns["name"] = name_c, art_c

    # Walk down collecting item rows.
    item_rows: dict[int, int] = {}
    blank_streak = 0
    r = header_row + 1
    max_row = ws.max_row
    seq = 0
    # Prefer № column; else use article/name presence.
    anchor_col = columns.get("no") or columns.get("article") or columns.get("name")
    if not anchor_col:
        return None

    # Skip category / sub-header lines until first article-like cell
    if "article" in columns:
        while r <= max_row:
            art = ws.cell(row=r, column=columns["article"]).value
            if isinstance(art, str) and art.strip() and art.strip().lower() not in {
                "артикул", "article", "sku", "описание", "description"
            }:
                # require code-like or keep and let filters handle
                break
            if art not in (None, "") and not isinstance(art, str):
                break
            # blank or header-ish
            if isinstance(art, str) and art.strip().lower() in {
                "артикул", "article", "sku", "описание", "description"
            }:
                r += 1
                continue
            # category row without article code — skip a few
            r += 1
            if r > header_row + 5:
                r = header_row + 1
                break

    start_r = r
    r = start_r
    while r <= max_row:
        if "no" in columns:
            val = ws.cell(row=r, column=columns["no"]).value
            if isinstance(val, (int, float)) and float(val).is_integer() and val > 0:
                item_rows[int(val)] = r
                blank_streak = 0
            else:
                blank_streak += 1
                if blank_streak >= 3:
                    break
        else:
            art = ws.cell(row=r, column=anchor_col).value
            if art not in (None, ""):
                art_s = str(art).strip().lower()
                if art_s not in {"артикул", "article", "sku", "описание", "description"}:
                    seq += 1
                    item_rows[seq] = r
                    blank_streak = 0
                else:
                    blank_streak += 1
            else:
                blank_streak += 1
                if blank_streak >= 3:
                    break
        r += 1

    if not item_rows:
        return None

    return SheetTable(sheet_name=ws.title, header_row=header_row, columns=columns, item_rows=item_rows)


def _find_header_without_no(ws: Worksheet, max_scan_rows: int = 60, max_scan_cols: int = 40):
    """Find a header row that contains article/name/qty-like labels without a № column."""
    max_row = min(ws.max_row, max_scan_rows)
    max_col = min(ws.max_column, max_scan_cols)
    keys = ("article", "name", "qty", "price", "amount", "net_weight", "gross_weight")
    best = None
    best_score = 0
    for r in range(1, max_row + 1):
        hits = 0
        has_article = False
        for c in range(1, max_col + 1):
            norm = normalize(ws.cell(row=r, column=c).value)
            if not norm:
                continue
            for fkey in keys:
                if any_header_keyword(HEADER_KEYWORDS[fkey], norm):
                    hits += 1
                    if fkey == "article":
                        has_article = True
                    break
        score = hits + (3 if has_article else 0)
        if score > best_score:
            best_score = score
            best = r
    if best_score >= 2:
        return best, None
    return None, None


def find_cert_in_row(ws: Worksheet, row: int, max_col: int = 40) -> dict | None:
    """Scan a row (any sheet, any column) for certificate/declaration-like
    text and return up to 2 matches, in the order the columns appear.

    Most rows carry either 0 or 1 certificate, but some carry 2 (e.g. a
    conformity certificate AND a separate свидетельство о гос.
    регистрации in an adjacent column) - those get written into the
    template's "мнр" and "мнр 2" blocks respectively.
    """
    matches: list[dict] = []
    for c in range(1, min(ws.max_column or max_col, max_col) + 1):
        raw = ws.cell(row=row, column=c).value
        if raw is None:
            continue
        val = raw if isinstance(raw, str) else str(raw)
        if not val.strip():
            continue
        m = CERT_RE.search(val)
        if not m:
            m = CERT_RE_SIMPLE.search(val)
        if not m:
            continue
        num = m.group("num").strip().rstrip(",;")
        d_from = parse_date(m.group("from"))
        d_to = parse_date(m.group("to")) if "to" in m.groupdict() else None
        matches.append({
            "mnr": num,
            "date_from": d_from,
            "date_to": d_to,
            "mnr_code": mnr_code_for(num),
        })
        if len(matches) >= 2:
            break

    if not matches:
        return None

    result = dict(matches[0])
    if len(matches) >= 2:
        result["mnr2"] = matches[1]["mnr"]
        result["date_from2"] = matches[1]["date_from"]
        result["date_to2"] = matches[1]["date_to"]
        result["mnr_code2"] = matches[1]["mnr_code"]
    return result


def sheet_priority_key(name: str) -> int:
    norm = normalize(name)
    for i, key in enumerate(DEFAULT_SHEET_PRIORITY):
        if key in norm:
            return i
    return len(DEFAULT_SHEET_PRIORITY)


# --------------------------------------------------------------------------
# Extraction from one or more input workbooks
# --------------------------------------------------------------------------

def extract_items_from_workbook(wb) -> dict[int, dict]:
    """Extract items; sum qty on invoice sheets, weights on packing sheets.

    Certificates are scanned on every sheet that has an article column
    (invoice / packing / specification / unnamed), then merged onto the
    matching article. Spec and other sheets also fill empty non-weight fields.
    """
    tables: list[SheetTable] = []
    for ws in wb.worksheets:
        t = build_sheet_table(ws)
        if t and "article" in t.columns:
            tables.append(t)

    sum_fields = {"qty", "amount", "net_weight", "gross_weight", "cll"}
    weight_fields = {"net_weight", "gross_weight", "cll"}

    def add_num(target: dict, key: str, value) -> None:
        if value in (None, ""):
            return
        if key in sum_fields:
            try:
                left = float(target[key]) if target.get(key) not in (None, "") else None
                right = float(value)
            except (TypeError, ValueError):
                left = right = None
            if left is not None and right is not None:
                total = left + right
                target[key] = int(total) if float(total).is_integer() else total
                return
        if target.get(key) in (None, ""):
            target[key] = value

    def sheet_kind(name: str) -> str:
        n = normalize(name)
        compact = re.sub(r"[^a-zа-яё0-9]+", "", n)
        if compact in {"pl", "packing", "packinglist", "пакинг", "упаковка"}:
            return "packing"
        if compact in {"inv", "invoice", "инвойс", "инв"}:
            return "invoice"
        if any(k in n for k in ("pack", "пак", "pl(", "pl ")):
            return "packing"
        if any(k in n for k in ("invoice", "инвойс", "in (", "in(")):
            return "invoice"
        return "other"

    def merge_missing(target: dict, source: dict, keys: list[str] | None = None) -> None:
        for k, v in source.items():
            if keys is not None and k not in keys:
                continue
            if v in (None, ""):
                continue
            if target.get(k) in (None, ""):
                target[k] = v

    def collect(kind_filter: str | None) -> dict[str, dict]:
        by_article: dict[str, dict] = {}
        for t in tables:
            if kind_filter is not None and sheet_kind(t.sheet_name) != kind_filter:
                continue
            ws = wb[t.sheet_name]
            for _no, row in t.item_rows.items():
                rec: dict = {}
                for field_key, col in t.columns.items():
                    if field_key == "no":
                        continue
                    val = ws.cell(row=row, column=col).value
                    if val not in (None, ""):
                        rec[field_key] = val
                if rec.get("country") in (None, ""):
                    found = countries.find_country_in_row(
                        ws, row, skip_cols=t.columns.values()
                    )
                    if found:
                        rec["country"] = found
                if rec.get("unit") in (None, ""):
                    found_u = units.find_unit_in_row(
                        ws, row, skip_cols=t.columns.values()
                    )
                    if found_u:
                        rec["unit"] = found_u
                art = rec.get("article")
                if art in (None, ""):
                    continue
                if is_summary_item(art, rec.get("name")):
                    continue
                # Category section titles without qty
                if rec.get("qty") in (None, "") and rec.get("price") in (None, "") and " " in str(art):
                    continue
                key = normalize(art)
                if key not in by_article:
                    by_article[key] = dict(rec)
                else:
                    for k, v in rec.items():
                        add_num(by_article[key], k, v)
                if not any(by_article[key].get(c) for c in CERT_FIELDS):
                    cert = find_cert_in_row(ws, row)
                    if cert:
                        for ck, cv in cert.items():
                            if by_article[key].get(ck) in (None, ""):
                                by_article[key][ck] = cv
        return by_article

    invoice = collect("invoice")
    packing = collect("packing")
    other = collect("other")

    # Unnamed / specification-only workbooks: treat "other" as the main source.
    if not invoice and not packing:
        invoice = other
        other = {}

    order = list(invoice.keys())
    for key, rec in packing.items():
        if key in invoice:
            for f in weight_fields:
                if rec.get(f) not in (None, ""):
                    invoice[key][f] = rec[f]
            merge_missing(invoice[key], rec, CERT_FIELDS + ["name", "country", "unit", "manufacturer", "marks"])
        else:
            invoice[key] = dict(rec)
            order.append(key)

    # Spec and other sheets often hold certificates and extra fields.
    for key, rec in other.items():
        if key in invoice:
            merge_missing(invoice[key], rec)
        else:
            invoice[key] = dict(rec)
            order.append(key)

    result = {i + 1: invoice[k] for i, k in enumerate(order)}
    doc_country = countries.find_document_country(wb)
    if doc_country:
        for rec in result.values():
            if rec.get("country") in (None, ""):
                rec["country"] = doc_country
    return result


def merge_workbooks(items_list: list[dict[int, dict]]) -> dict[int, dict]:
    """Merge items from multiple files; same article → sum qty/weights."""
    by_article: dict[str, dict] = {}
    order: list[str] = []
    no_article: list[dict] = []
    sum_fields = {"qty", "amount", "net_weight", "gross_weight", "cll"}

    def add_num(target: dict, key: str, value) -> None:
        if value in (None, ""):
            return
        if key in sum_fields:
            try:
                left = float(target[key]) if target.get(key) not in (None, "") else None
                right = float(value)
            except (TypeError, ValueError):
                left = right = None
            if left is not None and right is not None:
                total = left + right
                target[key] = int(total) if float(total).is_integer() else total
                return
        if target.get(key) in (None, ""):
            target[key] = value

    for items in items_list:
        for _no, rec in items.items():
            art = rec.get("article")
            key = normalize(art) if art not in (None, "") else ""
            if key:
                if key not in by_article:
                    by_article[key] = dict(rec)
                    order.append(key)
                else:
                    for k, v in rec.items():
                        add_num(by_article[key], k, v)
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


# --------------------------------------------------------------------------
# Writing into the template
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# Unified styling for cells written by fill_template (single font/color and
# alignment for the whole filled area, plus per-field number formats that
# override any leftover/inconsistent formatting baked into the template).
# --------------------------------------------------------------------------

DATA_FONT = Font(name="Cambria", size=11, bold=False, color="FF000000")
DATA_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=False)

DATA_NUMBER_FORMATS: dict[str, str] = {
    "no": "General",
    "tariff_code": "General",
    "name": "General",
    "article": "General",
    "marks": "General",
    "manufacturer": "General",
    "country": "General",
    "unit": "General",
    "qty": "General",
    "price": "0.00",
    "amount": "0.00",
    "net_weight": "0.00",
    "gross_weight": "0.00",
    "cll": "General",
    "mnr": "@",
    "date_from": "DD.MM.YYYY",
    "date_to": "DD.MM.YYYY",
    "mnr_code": "@",
    "mnr2": "@",
    "date_from2": "DD.MM.YYYY",
    "date_to2": "DD.MM.YYYY",
    "mnr_code2": "@",
}

TEMPLATE_FIELD_KEYWORDS = dict(HEADER_KEYWORDS)
# Base keyword -> (primary field key, secondary/"2" field key). The template
# has two certificate blocks: "мнр/от/до/код мнр" (1st cert) and
# "мнр 2/от 2/до 2/код мнр 2" (2nd cert, e.g. свидетельство о гос. регистрации).
TEMPLATE_CERT_KEYWORDS: dict[str, tuple[list[str], str, str]] = {
    "mnr": (["мнр"], "mnr", "mnr2"),
    "date_from": (["от"], "date_from", "date_from2"),
    "date_to": (["до"], "date_to", "date_to2"),
    "mnr_code": (["код мнр"], "mnr_code", "mnr_code2"),
}


def find_template_table(ws: Worksheet, max_scan_rows: int = 20):
    """Template header row is usually row 1, but search a little further
    just in case. Also locates the (мнр/от/до/код мнр) certificate columns,
    matched in header order to disambiguate от/до from от2/до2 duplicates."""
    header_row, no_col = find_no_header(ws, max_scan_rows=max_scan_rows)
    if header_row is None:
        header_row = 1

    max_col = ws.max_column
    columns: dict[str, int] = {}
    if no_col:
        columns["no"] = no_col

    cert_candidates: dict[str, list[int]] = {k: [] for k in TEMPLATE_CERT_KEYWORDS}

    for c in range(1, max_col + 1):
        raw = ws.cell(row=header_row, column=c).value
        norm = normalize(raw)
        if not norm:
            continue
        for fkey, keywords in TEMPLATE_FIELD_KEYWORDS.items():
            if fkey in columns:
                continue
            if not any_header_keyword(keywords, norm):
                continue
            if fkey == "unit" and any_header_keyword(TEMPLATE_FIELD_KEYWORDS["price"], norm):
                continue
            columns[fkey] = c
            break
        for cfkey, (keywords, _primary_key, _secondary_key) in TEMPLATE_CERT_KEYWORDS.items():
            # exact-ish match: header text equals the keyword, ignoring a
            # trailing " 2" duplicate-suffix (мнр 2 / от 2 / до 2 / код мнр 2)
            base = norm.replace(" 2", "").strip()
            if base in keywords:
                cert_candidates[cfkey].append(c)

    for cfkey, cols in cert_candidates.items():
        _keywords, primary_key, secondary_key = TEMPLATE_CERT_KEYWORDS[cfkey]
        if len(cols) >= 1:
            columns[primary_key] = cols[0]  # 1st certificate block (мнр/от/до/код мнр)
        if len(cols) >= 2:
            columns[secondary_key] = cols[1]  # 2nd certificate block (мнр 2/от 2/до 2/код мнр 2)

    return header_row, columns


def fill_template(template_bytes: bytes, items: dict[int, dict], settings: dict) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.worksheets[0]

    header_row, columns = find_template_table(ws)
    data_start = header_row + 1

    countries.apply_country_codes(items, fallback=settings.get("country"))
    units.apply_units(items, fallback=settings.get("unit"))

    for i, item_no in enumerate(sorted(items.keys())):
        rec = items[item_no]
        row = data_start + i

        def put(field_key, value):
            col = columns.get(field_key)
            if not col:
                return
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = DATA_FONT
            cell.alignment = DATA_ALIGNMENT
            numfmt = DATA_NUMBER_FORMATS.get(field_key)
            if numfmt:
                cell.number_format = numfmt

        put("no", item_no)
        put("tariff_code", rec.get("tariff_code"))
        put("name", rec.get("name"))
        put("article", rec.get("article"))
        put("marks", rec.get("marks"))
        put("manufacturer", rec.get("manufacturer"))
        put("country", rec.get("country"))
        put("unit", rec.get("unit"))
        put("qty", rec.get("qty"))
        put("price", rec.get("price"))
        put("amount", rec.get("amount"))
        put("net_weight", rec.get("net_weight"))
        put("gross_weight", rec.get("gross_weight"))
        put("cll", rec.get("cll"))
        put("mnr", rec.get("mnr"))
        put("date_from", rec.get("date_from"))
        put("date_to", rec.get("date_to"))
        put("mnr_code", rec.get("mnr_code"))
        put("mnr2", rec.get("mnr2"))
        put("date_from2", rec.get("date_from2"))
        put("date_to2", rec.get("date_to2"))
        put("mnr_code2", rec.get("mnr_code2"))

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------
# High-level entry point
# --------------------------------------------------------------------------

def process(template_bytes: bytes, input_files: list[bytes], settings: dict | None = None) -> tuple[bytes, dict]:
    settings = settings or {}
    items_list = []
    for content in input_files:
        wb = excel_io.load_workbook(content, data_only=True)
        items_list.append(extract_items_from_workbook(wb))

    merged = merge_workbooks(items_list)
    result_bytes = fill_template(template_bytes, merged, settings)

    report = {
        "mode": "keyword",
        "items_found": len(merged),
        "items": {no: {k: (str(v) if isinstance(v, datetime.datetime) else v) for k, v in rec.items()}
                  for no, rec in sorted(merged.items())},
    }
    return result_bytes, report
