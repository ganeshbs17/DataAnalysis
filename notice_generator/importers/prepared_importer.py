"""Phase 1 importer: Prepared Notice Data, one row = one notice (SRS §3.1)."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from ..core.models import CellType, RawCell, Record, SourceRef
from .cell_reader import SheetData, read_workbook
from .column_mapper import ColumnMapping, detect_header_row, map_columns


class ImportError_(Exception):
    """Plain-language import failure (NFR-09)."""


@dataclass
class ImportResult:
    path: Path
    sheets: list[tuple[str, int]]          # (name, non-empty row count)
    sheet: str
    mapping: ColumnMapping
    records: list[Record]
    skipped_empty: int
    warnings: list[str] = field(default_factory=list)


def list_sheets(path: str | Path) -> list[tuple[str, int]]:
    return [(s.name, s.non_empty_rows) for s in read_workbook(path)]


def import_prepared(path: str | Path, sheet: str | None = None, header_row: int | None = None,
                    overrides: dict[str, int] | None = None) -> ImportResult:
    """Import a prepared-notice workbook. The source file is opened read-only."""
    p = Path(path)
    sheets = read_workbook(p)
    if not sheets:
        raise ImportError_("The file contains no worksheets.")

    sd: SheetData | None
    if sheet is not None:
        sd = next((s for s in sheets if s.name == sheet), None)
        if sd is None:
            raise ImportError_(f"Worksheet '{sheet}' was not found in the file.")
    else:
        # Pick the first sheet that has a recognisable header row.
        sd = next((s for s in sheets if detect_header_row(s.rows)), None)
        if sd is None:
            raise ImportError_("No worksheet has a header row with the expected column names. "
                               "Choose the sheet and header row manually.")

    hr = header_row or detect_header_row(sd.rows)
    if not hr:
        raise ImportError_(f"No header row found in worksheet '{sd.name}'.")
    headers = [c.text.strip() for c in sd.rows[hr - 1]]
    mapping = map_columns(headers, hr)
    if overrides:
        for f, col in overrides.items():
            mapping.field_to_col[f] = col
            mapping.how[f] = "manual"

    warnings: list[str] = []
    for f in mapping.unmapped_fields:
        warnings.append(f"Column '{f}' was not found in the file.")

    records: list[Record] = []
    skipped = 0
    for r_idx in range(hr, len(sd.rows)):
        row = sd.rows[r_idx]
        if all(c.is_empty for c in row):
            skipped += 1
            continue
        src: dict[str, RawCell] = {}
        for f, col in list(mapping.field_to_col.items()) + list(mapping.extra_cols.items()):
            src[f] = row[col] if col < len(row) else RawCell()
        for f, c in src.items():
            if c.ctype == CellType.FORMULA_NO_CACHE:
                warnings.append(f"Row {r_idx + 1}: '{f}' is a formula with no saved value. "
                                "Open and save the file in Excel, or enter the value.")
        records.append(Record(serial=len(records) + 1,
                              source_ref=SourceRef(p.name, sd.name, r_idx + 1),
                              source=src))
    return ImportResult(path=p, sheets=[(s.name, s.non_empty_rows) for s in sheets], sheet=sd.name,
                        mapping=mapping, records=records, skipped_empty=skipped, warnings=warnings)
