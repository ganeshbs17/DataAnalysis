"""Header detection and column auto-mapping (FR-02 steps 3-4).

Mapping order: exact name → case/space-insensitive name → synonym list.
Worksheet names and column positions are never hardcoded.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..core.models import FIELDS, RawCell

DEFAULT_SYNONYMS: dict[str, list[str]] = {
    "SuspectName": ["name", "recipient name", "beneficiary name", "accused name", "suspect"],
    "Address": ["recipient address", "suspect address", "beneficiary address"],
    "AmountDeposited": ["amount", "amount credited", "total amount", "deposited amount"],
    "SuspectBankName": ["recipient bank", "beneficiary bank", "suspect bank"],
    "SuspectAccountNumber": ["recipient account", "beneficiary account", "suspect account",
                             "suspect a/c no", "account no", "account number"],
    "BankAccountName": ["victim bank", "originating bank", "complainant bank"],
    "BankAccountNumber": ["victim account", "originating account", "complainant account"],
    "DateOfTransfer": ["transfer date", "date of transaction", "transaction date", "dates"],
    "OfficerName": ["officer", "investigating officer", "io name"],
    "AppearanceDate": ["date of appearance"],
    "AppearanceTime": ["time of appearance"],
}


def norm_header(s: str) -> str:
    return re.sub(r"[\s_]+", "", s or "").lower()


@dataclass
class ColumnMapping:
    header_row: int                               # 1-based
    headers: list[str]
    field_to_col: dict[str, int] = field(default_factory=dict)
    how: dict[str, str] = field(default_factory=dict)   # field -> exact|normalised|synonym|manual
    extra_cols: dict[str, int] = field(default_factory=dict)
    fields: list[str] = field(default_factory=lambda: list(FIELDS))

    @property
    def unmapped_fields(self) -> list[str]:
        return [f for f in self.fields if f not in self.field_to_col]


def _known_lookup(synonyms: dict[str, list[str]], fields=None) -> dict[str, str]:
    lookup = {}
    flds = list(fields or FIELDS)
    for f in flds:
        lookup[norm_header(f)] = f
    for f, syns in synonyms.items():
        if f not in flds:
            continue
        for s in syns:
            lookup.setdefault(norm_header(s), f)
    return lookup


def detect_header_row(rows: list[list[RawCell]], synonyms=None, max_scan: int = 30,
                      fields: list[str] | None = None) -> int | None:
    """Return the 1-based index of the first row matching ≥3 known field names
    (or all of them, when the template has fewer than 3 fields)."""
    lookup = _known_lookup(synonyms or DEFAULT_SYNONYMS, fields)
    need = min(3, len(set(lookup.values()))) or 1
    for i, row in enumerate(rows[:max_scan], start=1):
        hits = {lookup[norm_header(c.text)] for c in row if norm_header(c.text) in lookup}
        if len(hits) >= need:
            return i
    return None


def map_columns(headers: list[str], header_row: int, synonyms=None,
                fields: list[str] | None = None) -> ColumnMapping:
    synonyms = synonyms or DEFAULT_SYNONYMS
    flds = list(fields or FIELDS)  # the template's data fields
    m = ColumnMapping(header_row=header_row, headers=headers, fields=flds)
    used: set[int] = set()

    def assign(f: str, col: int, how: str):
        m.field_to_col[f] = col
        m.how[f] = how
        used.add(col)

    # 1. exact
    for f in flds:
        for i, h in enumerate(headers):
            if i not in used and h == f:
                assign(f, i, "exact")
                break
    # 2. case/space-insensitive
    for f in flds:
        if f in m.field_to_col:
            continue
        for i, h in enumerate(headers):
            if i not in used and norm_header(h) == norm_header(f):
                assign(f, i, "normalised")
                break
    # 3. synonyms
    for f in flds:
        if f in m.field_to_col:
            continue
        syns = {norm_header(s) for s in synonyms.get(f, [])}
        for i, h in enumerate(headers):
            if i not in used and norm_header(h) in syns:
                assign(f, i, "synonym")
                break
    # Extra columns are kept under their own (cleaned) header names.
    for i, h in enumerate(headers):
        if i not in used and h.strip():
            m.extra_cols[re.sub(r"\W+", "", h.strip())] = i
    return m
