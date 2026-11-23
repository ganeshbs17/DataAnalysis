"""Read xlsx / xls / csv cells as raw text plus original cell type (FR-02).

Account numbers must never pass through ``float`` formatting. Integral numeric
cells are rendered as integer text; floats are converted via ``Decimal(repr())``
so a value such as 7692500102064500.0 becomes ``7692500102064500``, never
``7.6925E+15`` or ``...0.0``.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from ..core.models import CellType, RawCell

SCI_RE = re.compile(r"^\s*[+-]?\d+(?:\.\d+)?[eE][+-]?\d+\s*$")


@dataclass
class SheetData:
    name: str
    rows: list[list[RawCell]] = field(default_factory=list)

    @property
    def non_empty_rows(self) -> int:
        return sum(1 for r in self.rows if any(not c.is_empty for c in r))


class UnsupportedFile(Exception):
    pass


def number_cell(v: int | float) -> RawCell:
    """Convert a numeric cell value to integer/decimal text without float formatting."""
    if isinstance(v, bool):
        return RawCell(text=str(v).upper(), ctype=CellType.BOOL, value=v)
    if isinstance(v, int):
        s = str(v)
        return RawCell(text=s, ctype=CellType.INT, digits=len(s.lstrip("-")))
    d = Decimal(repr(float(v)))
    if d == d.to_integral_value():
        s = str(int(d))
        return RawCell(text=s, ctype=CellType.INT, digits=len(s.lstrip("-")))
    s = format(d.normalize(), "f")
    return RawCell(text=s, ctype=CellType.FLOAT, digits=len(s.split(".")[0].lstrip("-")))


def _datetime_cell(v) -> RawCell:
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0, 0):
            return RawCell(text=v.date().isoformat(), ctype=CellType.DATETIME, value=v.date())
        return RawCell(text=v.isoformat(sep=" "), ctype=CellType.DATETIME, value=v)
    if isinstance(v, dt.date):
        return RawCell(text=v.isoformat(), ctype=CellType.DATE, value=v)
    if isinstance(v, dt.time):
        return RawCell(text=v.strftime("%H:%M:%S"), ctype=CellType.TIME, value=v)
    if isinstance(v, dt.timedelta):  # openpyxl may return durations for [h]:mm
        t = (dt.datetime.min + v).time()
        return RawCell(text=t.strftime("%H:%M:%S"), ctype=CellType.TIME, value=t)
    raise TypeError(v)


def _python_value_cell(v) -> RawCell:
    if v is None:
        return RawCell()
    if isinstance(v, (bool, int, float)):
        return number_cell(v)
    if isinstance(v, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return _datetime_cell(v)
    s = str(v)
    return RawCell(text=s, ctype=CellType.TEXT if s.strip() else CellType.EMPTY)


# ---------------------------------------------------------------- xlsx
def _read_xlsx(path: Path) -> list[SheetData]:
    import openpyxl

    wb_vals = openpyxl.load_workbook(path, data_only=True)
    wb_form = openpyxl.load_workbook(path, data_only=False)
    sheets = []
    for ws in wb_vals.worksheets:
        wf = wb_form[ws.title]
        sd = SheetData(ws.title)
        for r_idx, row in enumerate(ws.iter_rows(), start=1):
            cells = []
            for c in row:
                if c.value is None:
                    fc = wf.cell(row=r_idx, column=c.column)
                    if fc.data_type == "f":
                        cells.append(RawCell(text="", ctype=CellType.FORMULA_NO_CACHE))
                        continue
                if c.data_type == "e":
                    cells.append(RawCell(text=str(c.value), ctype=CellType.ERROR))
                    continue
                cells.append(_python_value_cell(c.value))
            sd.rows.append(cells)
        sheets.append(sd)
    return sheets


# ---------------------------------------------------------------- xls
def _read_xls(path: Path) -> list[SheetData]:
    import xlrd

    book = xlrd.open_workbook(str(path))
    sheets = []
    for sh in book.sheets():
        sd = SheetData(sh.name)
        for r in range(sh.nrows):
            cells = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                t = cell.ctype
                if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    cells.append(RawCell())
                elif t == xlrd.XL_CELL_NUMBER:
                    cells.append(number_cell(cell.value))
                elif t == xlrd.XL_CELL_DATE:
                    if cell.value < 1:
                        tval = xlrd.xldate_as_datetime(cell.value, book.datemode).time()
                        cells.append(_datetime_cell(tval))
                    else:
                        cells.append(_datetime_cell(xlrd.xldate_as_datetime(cell.value, book.datemode)))
                elif t == xlrd.XL_CELL_BOOLEAN:
                    cells.append(RawCell(text="TRUE" if cell.value else "FALSE", ctype=CellType.BOOL))
                elif t == xlrd.XL_CELL_ERROR:
                    cells.append(RawCell(text="#ERR", ctype=CellType.ERROR))
                else:
                    cells.append(_python_value_cell(cell.value))
            sd.rows.append(cells)
        sheets.append(sd)
    return sheets


# ---------------------------------------------------------------- csv
def _read_csv(path: Path) -> list[SheetData]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:  # pragma: no cover - cp1252 decodes nearly anything
        text = raw.decode("latin-1")
    sd = SheetData(path.stem)
    for row in csv.reader(io.StringIO(text, newline="")):
        cells = []
        for v in row:
            if not v.strip():
                cells.append(RawCell(text=v, ctype=CellType.EMPTY))
            else:
                cells.append(RawCell(text=v, ctype=CellType.TEXT, scientific=bool(SCI_RE.match(v))))
        sd.rows.append(cells)
    return [sd]


def read_workbook(path: str | Path) -> list[SheetData]:
    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(p)
    if ext == ".xls":
        return _read_xls(p)
    if ext == ".csv":
        return _read_csv(p)
    raise UnsupportedFile(f"Unsupported file type '{ext}'. Use .xlsx, .xls or .csv.")
