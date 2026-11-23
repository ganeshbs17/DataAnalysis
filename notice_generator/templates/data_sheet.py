"""Excel import sheet that matches a template's fields."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .template_store import SOURCE_LABELS, STANDARD_HELP, TYPE_HELP, TYPE_LABELS, TemplateSpec

DATA_ROWS = 500  # rows pre-formatted for entry


def write_data_sheet(spec: TemplateSpec, out: str | Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Notice Data"
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F4E79")
    fields = spec.file_fields
    for col, f in enumerate(fields, start=1):
        c = ws.cell(row=1, column=col, value=f.name)
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(vertical="center")
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 45 if f.type == "multiline" else 24
        if f.type in ("account", "datelist", "text", "multiline"):
            fmt = "@"  # text: keeps every digit of account numbers, no auto-dates
        elif f.type == "date":
            fmt = "dd.mm.yyyy"
        elif f.type == "time":
            fmt = "hh:mm AM/PM"
        else:
            fmt = "General"
        for r in range(2, DATA_ROWS + 2):
            cell = ws.cell(row=r, column=col)
            cell.number_format = fmt
            if f.type == "multiline":
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"

    g = wb.create_sheet("Field guide")
    cols = ["Field (column heading)", "Type", "Required", "Filled from", "What to enter", "Example from source notice"]
    for i, h in enumerate(cols, start=1):
        c = g.cell(row=1, column=i, value=h)
        c.font, c.fill = head_font, head_fill
    for r, f in enumerate(spec.fields, start=2):
        help_text = STANDARD_HELP.get(f.name, "")
        tip = TYPE_HELP.get(f.type, "")
        if f.source != "file":
            tip = "Not in this sheet: chosen on the notice settings screen" if f.source == "settings" \
                else "Not in this sheet: calculated automatically"
        g.append([f.name, TYPE_LABELS.get(f.type, f.type), "Yes" if f.required else "No",
                  SOURCE_LABELS.get(f.source, f.source), "; ".join(x for x in (help_text, tip) if x),
                  f.example])
    for i, w in enumerate([26, 16, 10, 16, 70, 40], start=1):
        g.column_dimensions[get_column_letter(i)].width = w
    g.append([])
    g.append([f"Template: {spec.name} (version {spec.version})"])
    g.append(["One row in 'Notice Data' = one notice. Do not rename the column headings."])

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(out.name + ".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(out)
    return out
