"""Command line entry point. The engine runs without the GUI (SRS §7.1).

    python -m notice_generator gui
    python -m notice_generator validate --data notices_data.xlsx [--template T.docx]
    python -m notice_generator generate --data notices_data.xlsx --template T.docx --out OUT \
        --officer "Name" --date 02/12/2026 --time 11:00 --confirm-truncation
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys

from .core.models import NoticeSettings, Status
from .core.validation import validate_batch
from .importers.prepared_importer import import_prepared


def _date(s: str) -> dt.date:
    return dt.datetime.strptime(s, "%d/%m/%Y").date()


def _time(s: str) -> dt.time:
    for f in ("%H:%M", "%I:%M %p", "%I:%M%p"):
        try:
            return dt.datetime.strptime(s.upper(), f).time()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError(f"Invalid time '{s}'")


def _settings(a) -> NoticeSettings:
    return NoticeSettings(officer_name=a.officer or "", officer_designation=a.designation or "",
                          appearance_date=a.date, appearance_time=a.time,
                          notice_date=a.notice_date or dt.date.today(), use_file_values=a.use_file_values,
                          victim_bank=a.victim_bank or "", victim_account=a.victim_account or "",
                          remove_address_label=a.remove_address_label, prefix_mode=a.prefix)


def _common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--data", required=True, help="xlsx / xls / csv file")
    p.add_argument("--sheet")
    p.add_argument("--template")
    p.add_argument("--officer")
    p.add_argument("--designation")
    p.add_argument("--date", type=_date, help="Appearance date dd/mm/yyyy")
    p.add_argument("--time", type=_time, help="Appearance time, e.g. 11:00 or '11:00 AM'")
    p.add_argument("--notice-date", type=_date)
    p.add_argument("--use-file-values", action="store_true")
    p.add_argument("--victim-bank")
    p.add_argument("--victim-account")
    p.add_argument("--remove-address-label", action="store_true")
    p.add_argument("--prefix", choices=["keep", "remove", "standardise"], default="keep")
    p.add_argument("--confirm-truncation", action="store_true",
                   help="Officer confirms V-03 (possible Excel truncation) for every record")


def _report(records) -> None:
    for r in records:
        print(f"#{r.serial:03d} [{r.status.value:8}] {r.text('SuspectName').strip()}")
        for i in r.issues:
            print(f"      {i.rule} {i.severity.value}: {i.message}")
    counts = {s: sum(1 for r in records if r.status == s) for s in Status}
    print("Summary:", ", ".join(f"{k.value} {v}" for k, v in counts.items()))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="notice_generator")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("gui")
    pv = sub.add_parser("validate")
    _common(pv)
    pg = sub.add_parser("generate")
    _common(pg)
    pg.add_argument("--out", required=True)
    pg.add_argument("--no-docx", action="store_true")
    pg.add_argument("--no-pdf", action="store_true")
    pg.add_argument("--no-combined", action="store_true")
    pg.add_argument("--duplex", action="store_true")
    pg.add_argument("--rows", help="Comma-separated serial numbers to generate (default: all)")
    a = ap.parse_args(argv)

    if a.cmd == "gui":
        from .main import main as gui_main
        return gui_main()

    res = import_prepared(a.data, sheet=a.sheet)
    print(f"Imported {len(res.records)} records from sheet '{res.sheet}' "
          f"(header row {res.mapping.header_row}; {res.skipped_empty} empty rows skipped).")
    print("Mapping:", ", ".join(f"{f}←col{c + 1}({res.mapping.how[f]})" for f, c in res.mapping.field_to_col.items()))
    for w in res.warnings:
        print("WARNING:", w)
    s = _settings(a)
    if a.confirm_truncation:
        for r in res.records:
            r.confirmations.update({"V-03:SuspectAccountNumber", "V-03:BankAccountNumber"})
    if getattr(a, "rows", None):
        keep = {int(x) for x in a.rows.split(",")}
        for r in res.records:
            r.include = r.serial in keep

    placeholders = None
    if a.template:
        from .templates.placeholder_parser import find_placeholders
        placeholders = find_placeholders(a.template)
    validate_batch(res.records, s, template_placeholders=placeholders)
    _report(res.records)

    if a.cmd == "validate":
        return 0
    if not a.template:
        print("ERROR: --template is required for generate", file=sys.stderr)
        return 2
    from .exporters.batch_exporter import ExportOptions, run

    opts = ExportOptions(docx=not a.no_docx, pdf=not a.no_pdf, combined=not a.no_combined, duplex=a.duplex)
    result = run(res.records, s, a.template, a.out, opts,
                 progress=lambda d, t, m: print(f"  [{d}/{t}] {m}"), revalidate=False)
    print(f"\nOutput folder: {result.folder}")
    for it in result.items:
        line = f"#{it.record.serial:03d} {it.status:9}"
        if it.pdf:
            line += f" {it.pdf.name} ({it.pages} p)"
        elif it.docx:
            line += f" {it.docx.name}"
        if it.error:
            line += f" — {it.error}"
        print(line)
        for w in it.warnings:
            if w.rule == "V-15":
                print(f"      {w.rule}: {w.message}")
    if result.combined:
        print(f"Combined: {result.combined.name} ({result.combined_pages} pages)")
    print(f"Generated {result.count('generated')}, failed {result.count('failed')}, "
          f"blocked {result.count('blocked')}, excluded {result.count('excluded')}.")
    return 0 if result.count("failed") == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
