"""Build placeholder values for one record (FR-05 formatting, FR-06 precedence)."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from docxtpl import Listing

from . import formatting as fmt
from .models import NoticeSettings, Record
from .normaliser import apply_prefix_mode, normalise_account, normalise_address, normalise_text


@dataclass
class Effective:
    officer_name: str = ""
    appearance_date: dt.date | None = None
    appearance_time: dt.time | None = None
    conflicts: list[tuple[str, str, str]] = field(default_factory=list)  # (field, file, gui)


def effective_gui_fields(rec: Record, s: NoticeSettings) -> Effective:
    e = Effective()
    file_officer = normalise_text(rec.text("OfficerName")).replace("\n", " ")
    file_date = fmt.to_date(rec.typed_value("AppearanceDate"), rec.text("AppearanceDate"))
    file_time = fmt.to_time(rec.typed_value("AppearanceTime"), rec.text("AppearanceTime"))
    gui_officer, gui_date, gui_time = s.officer_name.strip(), s.appearance_date, s.appearance_time

    def pick(file_v, gui_v):
        if s.use_file_values:
            return file_v if file_v else gui_v
        return gui_v if gui_v else file_v

    e.officer_name = pick(file_officer, gui_officer)
    e.appearance_date = pick(file_date, gui_date)
    e.appearance_time = pick(file_time, gui_time)

    if file_officer and gui_officer and file_officer.lower() != gui_officer.lower():
        e.conflicts.append(("OfficerName", file_officer, gui_officer))
    if file_date and gui_date and file_date != gui_date:
        e.conflicts.append(("AppearanceDate", file_date.strftime(s.appearance_date_format),
                            gui_date.strftime(s.appearance_date_format)))
    if file_time and gui_time and file_time.replace(second=0) != gui_time.replace(second=0):
        e.conflicts.append(("AppearanceTime", fmt.format_time(file_time), fmt.format_time(gui_time)))
    return e


def build_values(rec: Record, s: NoticeSettings) -> dict[str, str]:
    """Formatted plain-text values keyed by placeholder name. Empty string = missing."""
    v: dict[str, str] = {}
    v["SuspectName"] = apply_prefix_mode(rec.text("SuspectName"), s.prefix_mode).replace("\n", " ")
    v["Address"] = normalise_address(rec.text("Address"), s.remove_address_label)
    v["SuspectBankName"] = normalise_text(rec.text("SuspectBankName")).replace("\n", " ")
    v["SuspectAccountNumber"] = normalise_account(rec.text("SuspectAccountNumber"))
    v["BankAccountName"] = normalise_text(rec.text("BankAccountName")).replace("\n", " ") or s.victim_bank.strip()
    v["BankAccountNumber"] = normalise_account(rec.text("BankAccountNumber")) or normalise_account(s.victim_account)

    amt = fmt.parse_amount(rec.text("AmountDeposited"))
    v["AmountDeposited"] = fmt.format_amount(amt, s.amount_suffix) if amt is not None and amt > 0 else ""
    v["AmountInWords"] = fmt.amount_in_words(amt) if amt is not None and amt > 0 else ""

    dates, _bad = fmt.format_transfer_dates(rec.text("DateOfTransfer"), s.transfer_date_format,
                                            rec.typed_value("DateOfTransfer"))
    v["DateOfTransfer"] = dates

    e = effective_gui_fields(rec, s)
    v["OfficerName"] = e.officer_name or ""
    v["OfficerDesignation"] = s.officer_designation.strip()
    v["AppearanceDate"] = e.appearance_date.strftime(s.appearance_date_format) if e.appearance_date else ""
    v["AppearanceTime"] = fmt.format_time(e.appearance_time) if e.appearance_time else ""
    v["NoticeDate"] = s.notice_date.strftime(s.appearance_date_format) if s.notice_date else ""
    v["CaseRef"] = s.case_ref.strip()
    v["CaseNumber"] = s.case_ref.strip()

    # Extra columns pass through as normalised text.
    for name in rec.source:
        if name not in v:
            v[name] = normalise_text(rec.text(name))
    return v


def build_context(rec: Record, s: NoticeSettings) -> dict[str, object]:
    """docxtpl context. Multi-line values become ``Listing`` so ``\\n`` prints as a Word
    line break inside the same paragraph and run (keeping its formatting)."""
    ctx: dict[str, object] = {}
    for k, val in build_values(rec, s).items():
        if val == "":
            continue  # left undefined → reported as unresolved (V-14)
        ctx[k] = Listing(val) if "\n" in val else val
    return ctx
