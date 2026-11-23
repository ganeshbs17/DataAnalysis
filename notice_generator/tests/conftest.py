"""Synthetic fixtures with the same quirks as the real sample files (SRS §3.1).

No real personal data is used: names, addresses and accounts are invented.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl
import pytest
from docx import Document
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from notice_generator.core.models import FIELDS, NoticeSettings  # noqa: E402

HEADER = ["SuspectName", "Address", "BankAccountName", "BankAccountNumber", "AmountDeposited",
          "DateOfTransfer", "OfficerName", "SuspectBankName", "SuspectAccountNumber", "AppearanceDate",
          "AppearanceTime"]
APPEAR = dt.datetime(2030, 12, 2)   # a Monday, safely in the future
TIME = dt.time(11, 0)

ROWS = [
    # 1-2: address variants (same person + account, prefix differs)
    ["RAVI SAMPLE", "SAMPLE TRADERS\nAddress:SHOP 1, MAIN ROAD, TOWN, STATE, 100001, INDIA\n\n", "Karnataka Bank ",
     1234567890123400, 2000000, "04.08.2025", "Officer A", "IDFC Bank", 51910100001, APPEAR, TIME],
    ["Mr. RAVI SAMPLE", " S/o Test Person  Near School Village 100001", "Karnataka Bank ",
     1234567890123400, 2000000, "04.08.2025", "Officer A", "IDFC Bank", 51910100001, APPEAR, TIME],
    # 3: multi-date text, amount in crores, & in address, float account cell
    ["Mr.KUMAR DEMO", "\nDEMO & SONS\nPLOT 5 <B>, CITY\n", "Karnataka Bank ", 1234567890123400, 13700000,
     "18.02.2025, 20.02.2025,\n21.02.2025,\n25.02.2025 ", "Officer A", "City Unin Bank",
     20100036980152.0, APPEAR, TIME],
    # 4-5: shared account under two different names; text account cell with spaces
    ["Asha Example", "House 7, Lane 2", "Karnataka Bank ", 1234567890123400, 200000, "20.01.2025",
     "Officer A", "Axis Bank", "2010 3065 3952", APPEAR, TIME],
    ["Rohit Example", "House 9, Lane 4", "Karnataka Bank ", 1234567890123400, 200000, "20.01.2025",
     "Officer A", "Axis Bank", "201030653952", APPEAR, TIME],
    # 6: empty row (skipped)
    [None] * 11,
    # 7: missing name → V-01
    [None, "Somewhere", "Karnataka Bank ", 1234567890123400, 50000, "01.03.2025", "Officer A", "Axis Bank",
     "998877665544", APPEAR, TIME],
]


@pytest.fixture
def sample_xlsx(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEADER)
    for r in ROWS:
        ws.append(r)
    p = tmp_path / "notices.xlsx"
    wb.save(p)
    return p


@pytest.fixture
def settings() -> NoticeSettings:
    return NoticeSettings(officer_name="Officer A", officer_designation="Inspector",
                          appearance_date=APPEAR.date(), appearance_time=TIME,
                          notice_date=dt.date(2026, 11, 23))


def _make_template(path: Path, extra: str | None = None, split_run: bool = True) -> Path:
    doc = Document()
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "Header for {{OfficerName}}"
    doc.add_paragraph("NOTICE")
    p = doc.add_paragraph("Appear on {{AppearanceDate}} at {{AppearanceTime}}.")
    p = doc.add_paragraph("A sum of ₹{{AmountDeposited}} was credited to {{SuspectBankName}} "
                          "(Account No: {{SuspectAccountNumber}}) from {{BankAccountName}} A/c No. "
                          "{{BankAccountNumber}} on {{DateOfTransfer}}.")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Officer: {{OfficerName}}"
    doc.add_paragraph("To,")
    p = doc.add_paragraph()
    if split_run:
        r1 = p.add_run("{{Suspect")
        r2 = p.add_run("Name}}")
        r2.bold = True
        r1.font.size = Pt(12)
    else:
        p.add_run("{{SuspectName}}")
    doc.add_paragraph("{{Address}}")
    if extra:
        doc.add_paragraph(extra)
    doc.save(path)
    return path


@pytest.fixture
def template(tmp_path: Path) -> Path:
    return _make_template(tmp_path / "template.docx")


@pytest.fixture
def template_unknown(tmp_path: Path) -> Path:
    return _make_template(tmp_path / "template_unknown.docx", extra="Ref {{CaseFileCode}}")


__all__ = ["FIELDS"]
