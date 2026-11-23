# Offline LEA Notice Generator — v0.1 proof of concept

This builds the first deliverable of SRS v1.1 (§12, milestone M0). It
imports *Prepared Notice Data* (one row = one notice), validates and
normalises it, fills an approved Word template, and generates DOCX, PDF
and a combined PDF. It runs fully offline.

## Run

```bash
pip install -r notice_generator/requirements.txt
# LibreOffice must be installed (or bundled in resources/bundled_renderer/)

python -m notice_generator gui                       # desktop window
python -m notice_generator validate --data notices_data.xlsx --template Notice_Template.docx
python -m notice_generator generate --data notices_data.xlsx --template Notice_Template.docx \
    --out output --officer "Name" --designation "Inspector" --date 02/12/2026 --time 11:00 \
    --confirm-truncation
```

Run the tests with `python -m pytest notice_generator/tests`. They use
synthetic fixtures only. No real case data is committed.

## What v0.1 covers (SRS references)

| Area | Implemented |
|------|-------------|
| Import (FR-02) | xlsx / xls / csv, read as **raw text plus cell type**; header-row detection; auto-mapping (exact → case/space-insensitive → synonyms); empty rows skipped |
| Account numbers (NFR-12) | Never stored as float. Numeric cells become integer text; 16+ digit numeric cells or CSV scientific notation → **V-03** (block until the officer confirms or edits) |
| Validation (FR-04) | V-01 to V-14 run before generation. V-15 (page count / blank last page) runs after PDF conversion. V-08 checks Sundays plus holidays configured in settings |
| Review (FR-03) | Searchable, sortable table; status filter; include/exclude; inline edits kept apart from the source and highlighted; source vs printed values; address-variant groups (V-09), shared accounts (V-10) |
| Formatting (FR-05) | `1,37,00,000/-` (the template supplies ₹), `d1, d2 and d3` transfer dates, `dd/mm/yyyy`, `hh:mm AM`, address line breaks inside the same paragraph, optional `Address:` label removal, name-prefix modes, `{{AmountInWords}}` |
| Precedence (FR-06) | GUI officer/date/time win by default; "use values from file" switch; conflicts → V-13 |
| Rendering (FR-07) | docxtpl with autoescape; split-run placeholders; headers/footers/tables; unresolved placeholders → V-14 (the record fails, its DOCX is not kept) |
| PDF (§6.1) | Headless LibreOffice with an isolated profile, one call per batch, timeout + kill + one retry |
| Output (FR-09) | `NNN_<Name>_<last4>`; Windows-safe names; never overwrites (`(2)`, `(3)`…); a notice counts as generated only after the PDF opens and its pages are counted; per-record failure isolation; combined PDF with optional duplex padding |
| Preview (FR-08) | QtPdf view of the actual generated PDF, with page count and V-15 result; "Preview selected record" |

Not in v0.1 (later milestones): case mode and Quick-mode temp handling,
generation log CSV, template manager and font preflight, Word COM
engine, backup, installer.

## Layout

```
importers/  cell_reader, column_mapper (header detection + mapping), prepared_importer
core/       models, normaliser, formatting, validation, grouping, notice_builder
templates/  placeholder_parser, renderer
exporters/  pdf_engine, combiner, batch_exporter
storage/    configuration (officer profiles), file_manager (naming, no-overwrite)
ui/         main_window (PySide6)
```

The engine has no GUI dependency. The CLI and the GUI both call
`exporters.batch_exporter.run()`.

Configuration lives in `%APPDATA%\NoticeGenerator` on Windows and in
`~/.notice_generator` elsewhere. Set `NOTICE_GEN_HOME` to override it.
Set `NOTICE_GEN_SOFFICE` to point at a specific `soffice` executable.

## Templates: creating, choosing and data sheets

**Templates…** (toolbar) opens the Template Library. Every saved template
is listed there each time the app starts. For the selected template you
can:

- use it;
- view or download the template (Word);
- view or download its **Excel data sheet**;
- edit its fields (this saves a new version, and earlier versions are kept);
- add an existing `.docx` that already contains `{{Field}}` placeholders.

The last template you used is selected automatically next time.

**Create template…** opens a three-step wizard:

1. Choose an existing notice (.docx), for example one already filled in
   for a case, and give the template a name. The original file is never
   changed.
2. Select text with the mouse, for example the recipient's name, and click
   **Make field**. Then:
   - choose a field name: a standard one (SuspectName, Address,
     AmountDeposited, SuspectAccountNumber, DateOfTransfer, …) or your own,
     e.g. `CaseNo`;
   - choose a type: Text, Multi-line, Amount, Date, Date list, Time or
     Account number.

   Repeated text can be replaced everywhere at once. A selection can span
   several lines, such as a whole address. The new placeholder keeps the
   formatting of the text it replaced. **Remove field** puts the original
   text back; **Undo** reverses the last step.
3. Review the fields and checks (for example, a fixed date still left in
   the header), then save.

When a template is saved, the app writes an **Excel data sheet** whose
column headings are the template's fields. Fields the officer sets on the
settings screen are not columns: OfficerName, OfficerDesignation,
AppearanceDate, AppearanceTime, NoticeDate and CaseRef. In the sheet:

- account-number columns are pre-formatted as Text, so Excel keeps all
  16+ digits;
- a "Field guide" tab explains every field.

Fill in one row per notice and import it with **Import data…**. The
columns map automatically. If a file's headings differ, use **Column
mapping…** to choose the worksheet, the header row and the column for
each field.

Values are formatted by field type:

| Type | Printed as |
|------|------------|
| Amount | `20,00,000/-` |
| Date list | `d1, d2 and d3` |
| Date | `dd.mm.yyyy` |
| Time | `hh:mm AM` |
| Multi-line | Line breaks kept |

Library files are stored in `Templates/<id>/v<N>/` (`template.docx`,
`data_sheet.xlsx`, `template.json` with the field list and SHA-256).

## Windows exe (DOCX-only)

The GitHub workflow `.github/workflows/build-windows.yml` builds
`NoticeGenerator.exe` on a Windows runner with PyInstaller
(`packaging/NoticeGenerator.spec`). It runs on every push to this
branch and can also be started by hand. Each run:

1. runs the tests;
2. builds the exe;
3. runs `NoticeGenerator.exe --selftest`, which renders a test DOCX and
   opens the main window offscreen;
4. uploads `NoticeGenerator-windows.zip` as a build artifact.

To install, download the zip from the workflow run, unzip it anywhere and
run `NoticeGenerator\NoticeGenerator.exe`. No Python and no installer
are needed.

PDF output needs LibreOffice. When it is not installed, the app runs in
**DOCX-only mode**: the PDF options are switched off, and "Preview
selected record" opens the generated Word document. If LibreOffice is
installed later, the PDF options turn on automatically.

The exe is not code-signed, so Windows SmartScreen or antivirus software
may warn about it until it is signed or whitelisted (SRS §6.1).
