"""Классификатор единиц измерения (ОКЕИ / инвойсы) → короткий код: ШТ, ПАР, М…"""
from __future__ import annotations

import json
import os
import re
from typing import Any, Iterable

_DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "units.json")

_LOADED = False
_CODES: set[str] = set()
_ALIAS: dict[str, str] = {}
_ALIAS_SORTED: list[tuple[str, str]] = []

_CURRENCY = frozenset({
    "usd", "eur", "cny", "rmb", "rur", "rub", "gbp", "jpy", "chf", "hkd", "$", "€", "¥",
})

_PUNCT_RE = re.compile(r"[«»\"'`´]")


def _norm(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower().replace("ё", "е")
    s = _PUNCT_RE.sub("", s)
    s = s.replace("²", "2").replace("³", "3")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_classifier(path: str | None = None) -> None:
    global _LOADED, _ALIAS_SORTED
    src = path or _DATA_PATH
    with open(src, encoding="utf-8") as f:
        payload = json.load(f)
    _CODES.clear()
    _ALIAS.clear()
    for rec in payload.get("units") or []:
        code = str(rec.get("code") or "").strip().upper()
        if not code:
            continue
        _CODES.add(code)
        _ALIAS.setdefault(_norm(code), code)
        _ALIAS.setdefault(code.lower(), code)
        for alias in rec.get("aliases") or []:
            key = _norm(alias)
            if key:
                _ALIAS.setdefault(key, code)
            compact = key.replace(" ", "").replace(".", "")
            if compact and compact not in _ALIAS:
                _ALIAS[compact] = code
    _ALIAS_SORTED = sorted(
        ((k, v) for k, v in _ALIAS.items() if len(k) >= 3),
        key=lambda kv: len(kv[0]),
        reverse=True,
    )
    _LOADED = True


def _ensure_loaded() -> None:
    if not _LOADED:
        load_classifier()


def known_tokens() -> set[str]:
    _ensure_loaded()
    return set(_ALIAS.keys())


def resolve_unit(value: Any) -> str | None:
    """Map a cell value to a canonical unit code, or None."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return None
    raw = str(value).strip()
    if not raw or len(raw) > 40:
        return None
    _ensure_loaded()
    key = _norm(raw)
    if not key or key in _CURRENCY:
        return None
    if re.fullmatch(r"[\d\s.,]+", key):
        return None
    if key in _ALIAS:
        return _ALIAS[key]
    compact = key.replace(" ", "").replace(".", "")
    if compact in _ALIAS:
        return _ALIAS[compact]
    # «pcs.», «шт,», «пар/компл»
    key_nosym = re.sub(r"[.,;:/]+$", "", key).strip()
    if key_nosym in _ALIAS:
        return _ALIAS[key_nosym]
    # Phrase: «единица: шт», «Unit: pair»
    for alias, code in _ALIAS_SORTED:
        if re.search(r"(?<![0-9a-zа-яё])" + re.escape(alias) + r"(?![0-9a-zа-яё])", key):
            return code
    return None


def find_unit_in_row(ws, row: int, *, skip_cols: Iterable[int | None] | None = None, max_col: int = 40) -> str | None:
    """Scan a data row for a short unit token in an unmapped column."""
    skip = {c for c in (skip_cols or []) if c}
    last = min(ws.max_column or max_col, max_col)
    for c in range(1, last + 1):
        if c in skip:
            continue
        val = ws.cell(row=row, column=c).value
        if val in (None, "") or isinstance(val, (int, float)):
            continue
        text = str(val).strip()
        if not text or len(text) > 16:
            continue
        code = resolve_unit(text)
        if code:
            return code
    return None


def apply_units(items: dict[int, dict], *, fallback: Any = None) -> None:
    """Replace raw unit text with canonical codes; empty → settings fallback."""
    fb = resolve_unit(fallback) if fallback not in (None, "") else None
    if not fb and fallback not in (None, ""):
        fb = str(fallback).strip().upper() or None
    for rec in items.values():
        raw = rec.get("unit")
        if raw not in (None, ""):
            rec["unit"] = resolve_unit(raw) or str(raw).strip()
            continue
        rec["unit"] = fb
