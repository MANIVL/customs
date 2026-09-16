"""Классификатор стран мира (ISO 3166-1 / ОКСМ) → код alpha-2."""
from __future__ import annotations

import json
import os
import re
import unicodedata
from typing import Any, Iterable

_DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "countries.json")

_LOADED = False
_CODES: set[str] = set()
_ALPHA3: dict[str, str] = {}
_NUMERIC: dict[str, str] = {}
_ALIAS: dict[str, str] = {}
_ALIAS_SORTED: list[tuple[str, str]] = []  # longest alias first

# Labels that mark a document-level origin cell.
_ORIGIN_LABELS = (
    "страна происхождения товара",
    "страна происхождения",
    "country of origin",
    "place of origin",
    "origin country",
    "страна-производитель",
    "страна производитель",
    "происхождение товара",
    "made in",
    "country/origin",
    "原产国",
    "原产地",
)

_PUNCT_RE = re.compile(r"[«»\"'`´‘’“”]")
_SEP_RE = re.compile(r"[\s_/\\,;:|()\[\]{}·•.–—-]+")
_DOT_RE = re.compile(r"\.")


def _fold(text: str) -> str:
    nfd = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")


def normalize_country_text(value: Any) -> str:
    if value is None:
        return ""
    s = _fold(str(value)).strip().lower().replace("ё", "е")
    s = _PUNCT_RE.sub("", s)
    s = _DOT_RE.sub("", s)
    s = _SEP_RE.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def _add_alias(alias: str, code: str) -> None:
    key = normalize_country_text(alias)
    if not key:
        return
    _ALIAS.setdefault(key, code)
    compact = key.replace(" ", "")
    if compact != key:
        _ALIAS.setdefault(compact, code)


def load_classifier(path: str | None = None) -> None:
    """Load (or reload) the countries JSON into lookup tables."""
    global _LOADED, _ALIAS_SORTED
    src = path or _DATA_PATH
    with open(src, encoding="utf-8") as f:
        payload = json.load(f)

    _CODES.clear()
    _ALPHA3.clear()
    _NUMERIC.clear()
    _ALIAS.clear()

    for rec in payload.get("countries") or []:
        code = str(rec.get("alpha2") or "").strip().upper()
        if len(code) != 2:
            continue
        _CODES.add(code)
        _add_alias(code, code)
        alpha3 = str(rec.get("alpha3") or "").strip().upper()
        if len(alpha3) == 3:
            _ALPHA3[alpha3] = code
        numeric = str(rec.get("numeric") or "").strip().zfill(3)
        if numeric.isdigit():
            _NUMERIC[numeric] = code
        for name in (rec.get("name_ru"), rec.get("name_en")):
            if name:
                _add_alias(str(name), code)
        for alias in rec.get("aliases") or []:
            _add_alias(str(alias), code)

    _ALIAS_SORTED = sorted(_ALIAS.items(), key=lambda kv: len(kv[0]), reverse=True)
    _LOADED = True


def _ensure_loaded() -> None:
    if not _LOADED:
        load_classifier()


def known_codes() -> set[str]:
    _ensure_loaded()
    return set(_CODES)


def _word_in_text(alias: str, text: str) -> bool:
    if not alias or alias not in text:
        return False
    pattern = r"(?<![0-9a-zа-я])" + re.escape(alias) + r"(?![0-9a-zа-я])"
    return re.search(pattern, text) is not None


def resolve_country(value: Any, *, allow_numeric: bool = True) -> str | None:
    """Map a cell/header value to ISO 3166-1 alpha-2, or None if unknown."""
    if value is None or value == "":
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    raw = str(value).strip()
    if not raw:
        return None

    _ensure_loaded()

    letters = re.sub(r"[^A-Za-zА-Яа-яЁё]", "", raw)
    if len(letters) == 2 and letters.isascii() and letters.isalpha():
        code = letters.upper()
        if code in _CODES:
            return code
    if len(letters) == 3 and letters.isascii() and letters.isalpha():
        mapped = _ALPHA3.get(letters.upper())
        if mapped:
            return mapped

    if allow_numeric:
        digits = re.sub(r"\D", "", raw)
        if digits and normalize_country_text(raw) in {digits, digits.zfill(3)} and 1 <= len(digits) <= 3:
            mapped = _NUMERIC.get(digits.zfill(3))
            if mapped:
                return mapped

    norm = normalize_country_text(raw)
    if not norm:
        return None
    if norm in _ALIAS:
        return _ALIAS[norm]
    compact = norm.replace(" ", "")
    if compact in _ALIAS:
        return _ALIAS[compact]

    # Phrase inside a longer cell: «Country of origin: China», «Китай / China».
    # Skip 1–2 character aliases here (IN, NO, TO, BE…) — too many false hits.
    for alias, code in _ALIAS_SORTED:
        if len(alias) < 3:
            continue
        if _word_in_text(alias, norm):
            return code
    return None


def find_country_in_row(ws, row: int, *, skip_cols: Iterable[int | None] | None = None, max_col: int = 40) -> str | None:
    """Scan a data row for a country name or ISO code."""
    skip = {c for c in (skip_cols or []) if c}
    last = min(ws.max_column or max_col, max_col)
    best_short: str | None = None
    for c in range(1, last + 1):
        if c in skip:
            continue
        val = ws.cell(row=row, column=c).value
        if val in (None, ""):
            continue
        if isinstance(val, (int, float)):
            continue
        text = str(val).strip()
        if not text or len(text) > 80:
            continue
        if re.fullmatch(r"[\d\s.,]+", text):
            continue
        code = resolve_country(text, allow_numeric=False)
        if not code:
            continue
        # Prefer a cell that is itself just a country / code.
        if len(text) <= 24:
            return code
        if best_short is None:
            best_short = code
    return best_short


def find_document_country(wb) -> str | None:
    """Look for origin in the header area of any sheet (outside the item table)."""
    _ensure_loaded()
    for ws in wb.worksheets:
        max_r = min(ws.max_row or 1, 50)
        max_c = min(ws.max_column or 1, 24)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = ws.cell(row=r, column=c).value
                if not isinstance(val, str) and not isinstance(val, (int, float)):
                    continue
                if val in (None, ""):
                    continue
                text = str(val).strip()
                if not text or len(text) > 120:
                    continue
                norm = normalize_country_text(text)
                if not norm:
                    continue
                labeled = any(normalize_country_text(label) in norm for label in _ORIGIN_LABELS)
                if labeled:
                    after = ""
                    if ":" in text:
                        after = text.split(":", 1)[1]
                    elif "：" in text:
                        after = text.split("：", 1)[1]
                    code = resolve_country(after) if after.strip() else None
                    if not code:
                        code = resolve_country(text)
                    if code:
                        return code
                    for dr, dc in ((0, 1), (0, 2), (1, 0), (1, 1)):
                        nr, nc = r + dr, c + dc
                        if nr > max_r or nc > max_c:
                            continue
                        neighbor = ws.cell(row=nr, column=nc).value
                        code = resolve_country(neighbor)
                        if code:
                            return code
    return None


def apply_country_codes(items: dict[int, dict], *, fallback: Any = None, document_country: str | None = None) -> None:
    """Replace raw country text on items with alpha-2 codes (in place)."""
    fb = resolve_country(fallback) if fallback not in (None, "") else None
    for rec in items.values():
        raw = rec.get("country")
        if raw not in (None, ""):
            code = resolve_country(raw)
            if not code:
                letters = re.sub(r"[^A-Za-z]", "", str(raw).strip())
                if len(letters) == 2:
                    code = letters.upper()
            rec["country"] = code
            continue
        rec["country"] = document_country or fb
