"""Load .xlsx / .xls workbooks into openpyxl Workbook objects."""
from __future__ import annotations

import io
import re

import openpyxl
from openpyxl.workbook import Workbook


XLSX_MAGIC = b"PK"
XLS_MAGIC = b"\xd0\xcf\x11\xe0"  # OLE Compound Document
_ILLEGAL_CELL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_BAD_TITLE_RE = re.compile(r'[:\\/?*\[\]]')


def sniff_excel_kind(content: bytes, filename: str | None = None) -> str:
    name = (filename or "").lower()
    if content.startswith(XLSX_MAGIC) or name.endswith(".xlsx"):
        return "xlsx"
    if content.startswith(XLS_MAGIC) or name.endswith(".xls"):
        return "xls"
    raise ValueError("Ожидается файл Excel (.xlsx или .xls)")


def load_workbook(content: bytes, *, data_only: bool = True, filename: str | None = None) -> Workbook:
    kind = sniff_excel_kind(content, filename)
    label = filename or kind
    try:
        if kind == "xlsx":
            return openpyxl.load_workbook(io.BytesIO(content), data_only=data_only)
        return _load_xls_as_workbook(content)
    except AssertionError as e:
        raise RuntimeError(
            f"Не удалось прочитать «{label}»: файл Excel повреждён или в неподдерживаемом формате"
        ) from e


def _safe_sheet_title(name: str | None, used: set[str]) -> str:
    title = _BAD_TITLE_RE.sub("_", (name or "Sheet")).strip()[:31] or "Sheet"
    base = title
    n = 2
    while title.lower() in used:
        suffix = f"_{n}"
        title = (base[: 31 - len(suffix)] + suffix) or f"Sheet{n}"
        n += 1
    used.add(title.lower())
    return title


def _safe_cell_value(value):
    if value is None:
        return None
    if isinstance(value, (int, float, bool)):
        return value
    if not isinstance(value, str):
        try:
            value = str(value)
        except Exception:
            return None
    value = _ILLEGAL_CELL_RE.sub("", value)
    if not value:
        return None
    return value[:32767]


def _load_xls_as_workbook(content: bytes) -> Workbook:
    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError(
            "Для .xls установите зависимость xlrd: pip install xlrd==1.2.0"
        ) from e

    try:
        book = xlrd.open_workbook(file_contents=content)
    except AssertionError as e:
        raise RuntimeError(
            "Не удалось прочитать .xls: файл в старом или повреждённом формате Excel"
        ) from e

    out = Workbook()
    out.remove(out.active)
    used_titles: set[str] = set()

    for sheet in book.sheets():
        ws = out.create_sheet(title=_safe_sheet_title(sheet.name, used_titles))
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                value = cell.value
                # xlrd dates come as floats with type XL_CELL_DATE
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate_as_datetime(cell.value, book.datemode)
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(cell.value)
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    value = None
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    value = None
                try:
                    ws.cell(row=r + 1, column=c + 1, value=_safe_cell_value(value))
                except Exception:
                    continue
    if not out.worksheets:
        out.create_sheet("Sheet1")
    return out
