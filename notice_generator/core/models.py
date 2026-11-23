"""Internal data model (SRS §8).

Account numbers are always held as ``str``; nothing in this module (or anywhere
in the pipeline) converts them to ``float`` (NFR-12).
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

# Prepared Notice Data columns (SRS §3.1). Header names match placeholder names.
FIELDS: list[str] = [
    "SuspectName",
    "Address",
    "AmountDeposited",
    "SuspectBankName",
    "SuspectAccountNumber",
    "BankAccountName",
    "BankAccountNumber",
    "DateOfTransfer",
    "OfficerName",
    "AppearanceDate",
    "AppearanceTime",
]
REQUIRED_FIELDS: list[str] = FIELDS[:8]
ACCOUNT_FIELDS: tuple[str, ...] = ("SuspectAccountNumber", "BankAccountNumber")
GUI_FIELDS: tuple[str, ...] = ("OfficerName", "AppearanceDate", "AppearanceTime")

# Internal meaning of each placeholder (SRS §8). Kept for documentation and
# for UI labels; the placeholder name is the key used throughout the PoC.
INTERNAL_NAMES: dict[str, str] = {
    "SuspectName": "recipient_name",
    "Address": "recipient_address",
    "SuspectBankName": "recipient_bank",
    "SuspectAccountNumber": "recipient_account",
    "BankAccountName": "victim_bank",
    "BankAccountNumber": "victim_account",
    "AmountDeposited": "amount",
    "DateOfTransfer": "transfer_dates",
    "OfficerName": "officer_name",
    "OfficerDesignation": "officer_designation",
    "AppearanceDate": "appearance_date",
    "AppearanceTime": "appearance_time",
    "NoticeDate": "notice_date",
    "CaseRef": "case_ref",
    "CaseNumber": "case_ref",
    "AmountInWords": "amount_in_words",
}


class CellType(str, Enum):
    EMPTY = "empty"
    TEXT = "text"
    INT = "int"          # numeric cell holding an integral value
    FLOAT = "float"      # numeric cell with a fractional part
    DATE = "date"
    DATETIME = "datetime"
    TIME = "time"
    BOOL = "bool"
    FORMULA_NO_CACHE = "formula_no_cache"
    ERROR = "error"


@dataclass
class RawCell:
    """A cell read as raw text plus its original type (FR-02 step 5)."""

    text: str = ""
    ctype: CellType = CellType.EMPTY
    value: Any = None          # date/time object for date-like cells
    digits: int = 0            # integer digit count for numeric cells
    scientific: bool = False   # CSV text in scientific notation

    @property
    def is_empty(self) -> bool:
        return self.ctype == CellType.EMPTY or (self.ctype == CellType.TEXT and not self.text.strip())


class Severity(str, Enum):
    BLOCK = "Block"
    CONFIRMABLE = "Block (confirmable)"
    WARNING = "Warning"
    INFO = "Information"


@dataclass
class Issue:
    rule: str            # e.g. "V-03"
    severity: Severity
    message: str
    field: str = ""
    suggestion: str = ""

    @property
    def key(self) -> str:
        """Identifier used to record an officer confirmation."""
        return f"{self.rule}:{self.field}"


class Status(str, Enum):
    READY = "Ready"
    WARNING = "Warning"
    BLOCKED = "Blocked"
    EXCLUDED = "Excluded"


@dataclass
class SourceRef:
    file: str
    sheet: str
    row: int  # 1-based worksheet row

    def __str__(self) -> str:
        return f"{self.file} / {self.sheet} / row {self.row}"


@dataclass
class Record:
    serial: int
    source_ref: SourceRef
    source: dict[str, RawCell]
    edits: dict[str, str] = field(default_factory=dict)
    include: bool = True
    confirmations: set[str] = field(default_factory=set)
    issues: list[Issue] = field(default_factory=list)

    # ----- values -----
    def raw(self, name: str) -> RawCell:
        return self.source.get(name, RawCell())

    def text(self, name: str) -> str:
        """Current text value: the officer's edit if any, else the source text."""
        if name in self.edits:
            return self.edits[name]
        return self.raw(name).text

    def typed_value(self, name: str) -> Any:
        """Date/time object from the source cell, unless the value was edited."""
        if name in self.edits:
            return None
        return self.raw(name).value

    def is_edited(self, name: str) -> bool:
        return name in self.edits and self.edits[name] != self.raw(name).text

    # ----- status -----
    def blocking_issues(self) -> list[Issue]:
        out = []
        for i in self.issues:
            if i.severity == Severity.BLOCK:
                out.append(i)
            elif i.severity == Severity.CONFIRMABLE and i.key not in self.confirmations:
                out.append(i)
        return out

    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == Severity.WARNING]

    @property
    def status(self) -> Status:
        if not self.include:
            return Status.EXCLUDED
        if self.blocking_issues():
            return Status.BLOCKED
        if self.warnings():
            return Status.WARNING
        return Status.READY


@dataclass
class NoticeSettings:
    """Batch-level notice configuration (FR-05, FR-06)."""

    officer_name: str = ""
    officer_designation: str = ""
    appearance_date: _dt.date | None = None
    appearance_time: _dt.time | None = None
    notice_date: _dt.date = field(default_factory=_dt.date.today)
    use_file_values: bool = False          # FR-06 precedence switch
    victim_bank: str = ""                  # case setting, fills empty file values
    victim_account: str = ""
    case_ref: str = ""
    remove_address_label: bool = False     # FR-05, off by default
    prefix_mode: str = "keep"              # keep | remove | standardise
    amount_suffix: str = "/-"
    transfer_date_format: str = "%d.%m.%Y"
    appearance_date_format: str = "%d/%m/%Y"
    holidays: list[_dt.date] = field(default_factory=list)
