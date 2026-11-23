"""Validation rules V-01 … V-14 (FR-04). V-15 is checked after PDF conversion."""
from __future__ import annotations

import datetime as dt
import difflib
import json
import re
from pathlib import Path

from . import formatting as fmt
from .grouping import address_variant_groups, exact_duplicates, shared_accounts
from .models import ACCOUNT_FIELDS, REQUIRED_FIELDS, CellType, Issue, NoticeSettings, Record, Severity
from .notice_builder import build_values, effective_gui_fields
from .normaliser import normalise_text

_BANKS_FILE = Path(__file__).resolve().parent.parent / "resources" / "bank_names.json"


def load_bank_names(path: Path | None = None) -> list[str]:
    p = path or _BANKS_FILE
    try:
        return list(json.loads(p.read_text(encoding="utf-8"))["banks"])
    except (OSError, ValueError, KeyError):
        return []


def _bank_key(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip().lower()


def check_bank(name: str, banks: list[str]) -> str | None:
    """Return None if known, '' if unknown with no suggestion, or the closest match."""
    if not name:
        return None
    keys = {_bank_key(b): b for b in banks}
    k = _bank_key(name)
    if k in keys:
        return None
    match = difflib.get_close_matches(k, list(keys), n=1, cutoff=0.75)
    return keys[match[0]] if match else ""


def validate_record(rec: Record, s: NoticeSettings, banks: list[str],
                    template_placeholders: set[str] | None = None,
                    today: dt.date | None = None, spec=None) -> list[Issue]:
    """``spec`` (TemplateSpec) switches the field lists to the template's own fields and types."""
    today = today or dt.date.today()
    issues: list[Issue] = []
    vals = build_values(rec, s, spec)
    if spec is not None:
        required = [f.name for f in spec.fields if f.required and f.source == "file"]
        account_fields = spec.names_of_type("account")
        amount_fields = spec.names_of_type("amount")
        date_fields = spec.names_of_type("date", "datelist")
    else:
        required = REQUIRED_FIELDS
        account_fields = list(ACCOUNT_FIELDS)
        amount_fields = ["AmountDeposited"]
        date_fields = ["DateOfTransfer"]

    # V-01 required fields
    for f in required:
        if f in amount_fields:
            if not normalise_text(rec.text(f)):
                issues.append(Issue("V-01", Severity.BLOCK, f"{f} (amount) is empty.", f))
            continue
        if not vals.get(f):
            issues.append(Issue("V-01", Severity.BLOCK, f"{f} is empty.", f))
    e = effective_gui_fields(rec, s)
    for f, v in (("OfficerName", e.officer_name), ("AppearanceDate", e.appearance_date),
                 ("AppearanceTime", e.appearance_time)):
        if not v and (template_placeholders is None or f in template_placeholders):
            issues.append(Issue("V-01", Severity.BLOCK, f"{f} is not set. Choose it in the notice settings.", f))

    # V-02 / V-03 / V-04 account numbers
    for f in account_fields:
        acc = vals.get(f, "")
        if not acc:
            continue
        if not acc.isdigit():
            issues.append(Issue("V-02", Severity.WARNING,
                                f"{f} contains characters other than digits: '{acc}'.", f))
        raw = rec.raw(f)
        if f not in rec.edits and (
                (raw.ctype in (CellType.INT, CellType.FLOAT) and raw.digits >= 16) or raw.scientific
                or "e+" in raw.text.lower()):
            issues.append(Issue("V-03", Severity.CONFIRMABLE,
                                f"{f} {acc}: possibly truncated by Excel. Verify against the bank statement.", f))
        digits = re.sub(r"\D", "", acc)
        if digits and not 9 <= len(digits) <= 18:
            issues.append(Issue("V-04", Severity.WARNING,
                                f"{f} has {len(digits)} digits (expected 9–18).", f))

    # V-05 amount
    for af in amount_fields:
        raw_amt = normalise_text(rec.text(af))
        if raw_amt:
            amt = fmt.parse_amount(raw_amt)
            if amt is None:
                issues.append(Issue("V-05", Severity.BLOCK, f"{af} '{raw_amt}' cannot be read as a number.", af))
            elif amt <= 0:
                issues.append(Issue("V-05", Severity.BLOCK, f"{af} must be greater than zero.", af))

    # V-06 dates from the file
    for df in date_fields:
        tv = rec.typed_value(df)
        parts = [tv] if tv else [fmt.parse_date(p) for p in fmt.split_dates(rec.text(df))]
        raw_parts = fmt.split_dates(rec.text(df)) if not tv else [str(tv)]
        for raw_p, d in zip(raw_parts, parts):
            if isinstance(d, dt.datetime):
                d = d.date()
            if d is None:
                issues.append(Issue("V-06", Severity.WARNING,
                                    f"{df}: date '{raw_p}' cannot be read; it will be printed as written.", df))
            elif s.notice_date and d > s.notice_date:
                issues.append(Issue("V-06", Severity.WARNING,
                                    f"{df}: date {d:%d.%m.%Y} is after the notice date.", df))

    # V-07 / V-08 appearance date
    ad = e.appearance_date
    if ad:
        if ad < today or (s.notice_date and ad < s.notice_date):
            issues.append(Issue("V-07", Severity.BLOCK,
                                f"Appearance date {ad:%d/%m/%Y} is in the past or before the notice date.",
                                "AppearanceDate"))
        if ad.weekday() == 6:
            issues.append(Issue("V-08", Severity.WARNING, f"Appearance date {ad:%d/%m/%Y} is a Sunday.",
                                "AppearanceDate"))
        elif ad in s.holidays:
            issues.append(Issue("V-08", Severity.WARNING, f"Appearance date {ad:%d/%m/%Y} is a holiday.",
                                "AppearanceDate"))

    # V-12 bank names
    for f in ("SuspectBankName", "BankAccountName"):
        sug = check_bank(vals.get(f, ""), banks)
        if sug is not None:
            msg = f"Bank name '{vals[f]}' is not in the bank list."
            if sug:
                msg += f" Did you mean '{sug}'?"
            issues.append(Issue("V-12", Severity.WARNING, msg, f, suggestion=sug))

    # V-13 file vs GUI
    for f, file_v, gui_v in e.conflicts:
        used = file_v if s.use_file_values else gui_v
        issues.append(Issue("V-13", Severity.WARNING,
                            f"{f}: file has '{file_v}', settings have '{gui_v}'. Using '{used}'.", f))

    # V-14 (pre-check) every template placeholder must get a value
    if template_placeholders:
        for ph in sorted(template_placeholders):
            if not vals.get(ph) and not any(i.field == ph and i.rule == "V-01" for i in issues):
                issues.append(Issue("V-14", Severity.BLOCK, f"Placeholder {{{{{ph}}}}} has no value.", ph))
    return issues


def validate_batch(records: list[Record], s: NoticeSettings, banks: list[str] | None = None,
                   template_placeholders: set[str] | None = None, today: dt.date | None = None,
                   spec=None) -> None:
    """Run all rules and store results on each record (in place)."""
    banks = load_bank_names() if banks is None else banks
    for r in records:
        r.issues = validate_record(r, s, banks, template_placeholders, today, spec)

    for grp in address_variant_groups(records):
        for r in grp:
            others = ", ".join(str(o.serial) for o in grp if o is not r)
            r.issues.append(Issue("V-09", Severity.INFO,
                                  f"Address variant: same name and account as record(s) {others}. "
                                  f"{len(grp)} separate notices will be generated.", "SuspectName"))
    for acc, rs in shared_accounts(records).items():
        names = sorted({normalise_text(r.text('SuspectName')) for r in rs})
        for r in rs:
            r.issues.append(Issue("V-10", Severity.WARNING,
                                  f"Account {acc} also appears under a different name: {', '.join(names)}.",
                                  "SuspectAccountNumber"))
    for grp in exact_duplicates(records):
        for r in grp:
            others = ", ".join(str(o.serial) for o in grp if o is not r)
            r.issues.append(Issue("V-11", Severity.WARNING,
                                  f"Exact duplicate of record(s) {others}.", "SuspectName"))
