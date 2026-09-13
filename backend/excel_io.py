"""Load .xlsx / .xls workbooks into openpyxl Workbook objects."""
from __future__ import annotations

import io
from typing import BinaryIO

import openpyxl
from openpyxl.workbook import Workbook


XLSX_MAGIC = b"PK"
XLS_MAGIC = b"\xd0\xcf\x11\xe0"  # OLE Compound Document


def sniff_excel_kind(content: bytes, filename: str | None = None) -> str:
    name = (filename or "").lower()
    if content.startswith(XLSX_MAGIC) or name.endswith(".xlsx"):
        return "xlsx"
    if content.startswith(XLS_MAGIC) or name.endswith(".xls"):
        return "xls"
    raise ValueError("Ожидается файл Excel (.xlsx или .xls)")


def load_workbook(content: bytes, *, data_only: bool = True, filename: str | None = None) -> Workbook:
    kind = sniff_excel_kind(content, filename)
    if kind == "xlsx":
        return openpyxl.load_workbook(io.BytesIO(content), data_only=data_only)
    return _load_xls_as_workbook(content)


def _load_xls_as_workbook(content: bytes) -> Workbook:
    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError(
            "Для .xls установите зависимость xlrd: pip install xlrd==1.2.0"
        ) from e

    book = xlrd.open_workbook(file_contents=content)
    out = Workbook()
    out.remove(out.active)

    for sheet in book.sheets():
        ws = out.create_sheet(title=(sheet.name or "Sheet")[:31])
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
                ws.cell(row=r + 1, column=c + 1, value=value)
    if not out.worksheets:
        out.create_sheet("Sheet1")
    return out
