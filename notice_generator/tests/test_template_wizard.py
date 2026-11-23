"""Create-Template wizard engine: field marking, library, data sheet, end-to-end generation."""
import datetime as dt
import os
import time
import zipfile

import openpyxl
import pytest
from docx import Document
from docx.shared import Pt

from notice_generator.core.models import NoticeSettings
from notice_generator.exporters.batch_exporter import ExportOptions, run
from notice_generator.importers.prepared_importer import import_prepared
from notice_generator.templates.docx_editor import LINE_SEP, DocxEditor, EditError
from notice_generator.templates.placeholder_parser import find_placeholders, iter_part_texts
from notice_generator.templates.template_store import FieldSpec, TemplateStore, fields_from_placeholders


@pytest.fixture
def filled_notice(tmp_path):
    """A notice already filled for one case: the investigator's starting point."""
    d = Document()
    d.sections[0].header.paragraphs[0].text = "Cyber Crime PS   Cr. No. 17/2026"
    d.add_paragraph("NOTICE UNDER SECTION 35(3) BNSS")
    p = d.add_paragraph("A sum of ₹")
    r = p.add_run("1,37,00,000/-")
    r.bold = True
    p.add_run(" was credited to your account ")
    p.add_run("6309").font.size = Pt(14)
    p.add_run("5881591")             # account split across two runs with different formatting
    p.add_run(" held by RAVI KUMAR.")
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "Case"
    t.cell(0, 1).text = "17/2026"
    d.add_paragraph("To,")
    d.add_paragraph("RAVI KUMAR")
    a = d.add_paragraph("SANJU HALLS")
    a.add_run().add_break()
    a.add_run("7-2-35 MAIN ROAD")
    d.add_paragraph("HYDERABAD 500018")
    d.add_paragraph("Signed")
    path = tmp_path / "filled.docx"
    d.save(path)
    return path


def _pid(ed, startswith):
    return next(p.pid for p in ed.paragraphs if p.text.startswith(startswith))


def test_field_inside_formatted_runs_keeps_first_run_format(filled_notice, tmp_path):
    ed = DocxEditor(filled_notice)
    pid = _pid(ed, "A sum of")
    text = ed.paragraphs[pid].text
    i = text.index("63095881591")
    ed.make_field(pid, i, pid, i + 11, "SuspectAccountNumber")
    assert ed.paragraphs[pid].text == ("A sum of ₹1,37,00,000/- was credited to your account "
                                       "{{SuspectAccountNumber}} held by RAVI KUMAR.")
    out = tmp_path / "t.docx"
    ed.save(out)
    para = next(p for p in Document(str(out)).paragraphs if "{{SuspectAccountNumber}}" in p.text)
    run = next(r for r in para.runs if "{{SuspectAccountNumber}}" in r.text)
    assert run.font.size == Pt(14)
    assert ed.originals["SuspectAccountNumber"] == "63095881591"


def test_replace_all_body_table_header(filled_notice, tmp_path):
    ed = DocxEditor(filled_notice)
    assert ed.count_occurrences("17/2026") == 2
    assert ed.replace_all("17/2026", "CrimeNo") == 2
    assert ed.replace_all("RAVI KUMAR", "SuspectName") == 2
    out = tmp_path / "t.docx"
    ed.save(out)
    texts = {part: t for part, t in iter_part_texts(out) if "CrimeNo" in t}
    assert any("header" in p for p in texts) and any("document" in p for p in texts)
    assert find_placeholders(out) == {"CrimeNo", "SuspectName"}
    Document(str(out))


def test_multi_paragraph_address_and_line_break(filled_notice, tmp_path):
    ed = DocxEditor(filled_notice)
    a, b = _pid(ed, "SANJU"), _pid(ed, "HYDERABAD")
    assert LINE_SEP in ed.paragraphs[a].text
    ed.make_field(a, 0, b, len(ed.paragraphs[b].text), "Address")
    assert ed.originals["Address"] == "SANJU HALLS\n7-2-35 MAIN ROAD\nHYDERABAD 500018"
    texts = [p.text for p in ed.paragraphs]
    assert "{{Address}}" in texts and not any("HYDERABAD" in t for t in texts)
    assert texts[texts.index("{{Address}}") + 1] == "Signed"
    out = tmp_path / "t.docx"
    ed.save(out)
    assert "<w:br/>" not in zipfile.ZipFile(out).read("word/document.xml").decode().split("{{Address}}")[1][:200]


def test_invalid_selections(filled_notice):
    ed = DocxEditor(filled_notice)
    pid = _pid(ed, "To,")
    with pytest.raises(EditError):
        ed.make_field(pid, 0, pid, 2, "bad name")
    with pytest.raises(EditError):
        ed.make_field(pid, 0, pid, 0, "X")
    header = next(p.pid for p in ed.paragraphs if p.region == "Header")
    with pytest.raises(EditError):
        ed.make_field(pid, 0, header, 3, "X")          # body → header
    ed.make_field(pid, 0, pid, 3, "Greeting")
    with pytest.raises(EditError):
        ed.make_field(pid, 0, pid, len(ed.paragraphs[pid].text), "Again")  # already a field


def test_remove_and_undo(filled_notice):
    ed = DocxEditor(filled_notice)
    ed.replace_all("RAVI KUMAR", "SuspectName")
    assert ed.remove_field("SuspectName") == 2
    assert ed.placeholders() == [] and ed.count_occurrences("RAVI KUMAR") == 2
    ed.undo()
    assert ed.placeholders() == ["SuspectName"]
    ed.undo()
    assert ed.placeholders() == []


def test_name_inside_token_does_not_loop(filled_notice):
    ed = DocxEditor(filled_notice)
    assert ed.replace_all("Signed", "Signed") == 1


def _make_template(filled_notice, store):
    ed = DocxEditor(filled_notice)
    ed.replace_all("1,37,00,000/-", "AmountDeposited")
    ed.replace_all("63095881591", "SuspectAccountNumber")
    ed.replace_all("RAVI KUMAR", "SuspectName")
    ed.replace_all("17/2026", "CrimeNo")
    ed.replace_all("Signed", "OfficerName")
    a, b = _pid(ed, "SANJU"), _pid(ed, "HYDERABAD")
    ed.make_field(a, 0, b, len(ed.paragraphs[b].text), "Address")
    fields = fields_from_placeholders(ed.placeholders(), ed.originals)
    return store.save("Bank credit notice", ed.to_bytes(), fields, "test", source_doc="filled.docx")


def test_store_and_data_sheet(filled_notice, tmp_path):
    store = TemplateStore(tmp_path / "Templates")
    spec = _make_template(filled_notice, store)
    assert {f.name: f.source for f in spec.fields}["OfficerName"] == "settings"
    again = TemplateStore(tmp_path / "Templates").list()          # after a restart
    assert [s.name for s in again] == ["Bank credit notice"] and again[0].sha256 == spec.sha256
    wb = openpyxl.load_workbook(spec.data_sheet)
    ws = wb["Notice Data"]
    heads = [c.value for c in ws[1]]
    assert heads == [f.name for f in spec.file_fields]           # order of first appearance
    assert sorted(heads) == ["Address", "AmountDeposited", "CrimeNo", "SuspectAccountNumber", "SuspectName"]
    assert ws.cell(2, heads.index("SuspectAccountNumber") + 1).number_format == "@"
    guide = [r[0].value for r in wb["Field guide"].iter_rows(min_row=2) if r[0].value]
    assert "OfficerName" in guide
    v2 = store.save(spec.name, spec.docx.read_bytes(), spec.fields, template_id=spec.id)
    assert v2.version == 2 and store.versions(spec.id) == [1, 2] and spec.docx.exists()


def test_end_to_end_from_data_sheet(filled_notice, tmp_path):
    store = TemplateStore(tmp_path / "Templates")
    spec = _make_template(filled_notice, store)
    wb = openpyxl.load_workbook(spec.data_sheet)
    ws = wb["Notice Data"]
    row = {"CrimeNo": "21/2026", "AmountDeposited": 2365000, "SuspectAccountNumber": "1234567890123456",
           "SuspectName": "A & B Traders", "Address": "Line 1\nLine 2"}
    for c, h in enumerate([c.value for c in ws[1]], start=1):
        ws.cell(2, c, row[h])
    filled = tmp_path / "data.xlsx"
    wb.save(filled)
    res = import_prepared(filled, fields=[f.name for f in spec.file_fields])
    assert len(res.records) == 1 and res.skipped_empty == 0 and not res.mapping.unmapped_fields
    s = NoticeSettings(officer_name="Officer X", appearance_date=dt.date(2030, 1, 7),
                       appearance_time=dt.time(11, 0), notice_date=dt.date(2026, 11, 23))
    out = run(res.records, s, spec.docx, tmp_path / "out", ExportOptions(pdf=False, combined=False))
    it = out.items[0]
    assert it.status == "generated", it.error
    text = "\n".join(t for _p, t in iter_part_texts(it.docx))
    for expect in ["23,65,000/-", "1234567890123456", "A & B Traders", "21/2026", "Officer X"]:
        assert expect in text
    assert "{{" not in text and "V-03" not in {i.rule for i in it.record.issues}


def test_import_overrides_map_renamed_column(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Name of accused", "Amt", "CrimeNo"])
    ws.append(["X", 100, "1/2026"])
    p = tmp_path / "d.xlsx"
    wb.save(p)
    fields = ["SuspectName", "AmountDeposited", "CrimeNo"]
    res = import_prepared(p, header_row=1, fields=fields, overrides={"SuspectName": 0, "AmountDeposited": 1})
    assert res.records[0].text("SuspectName") == "X" and res.records[0].text("AmountDeposited") == "100"
    assert not res.mapping.unmapped_fields


@pytest.mark.skipif(os.environ.get("QT_QPA_PLATFORM") != "offscreen" and not os.environ.get("DISPLAY"),
                    reason="no display")
def test_wizard_gui(filled_notice, tmp_path, monkeypatch):
    qtw = pytest.importorskip("PySide6.QtWidgets")
    from PySide6.QtGui import QTextCursor
    from PySide6.QtWidgets import QApplication, QMessageBox

    from notice_generator.ui.template_library import TemplateLibrary
    from notice_generator.ui.template_wizard import TemplateWizard

    app = QApplication.instance() or qtw.QApplication([])
    monkeypatch.setattr(QMessageBox, "exec", lambda self: 0)
    store = TemplateStore(tmp_path / "Templates")
    w = TemplateWizard(store=store)
    w.path.setText(str(filled_notice))
    w.tname.setText("GUI template")
    w._next()
    doc = w.viewer.document()

    def select(text, end_text=None):
        c = doc.find(text)
        if end_text:
            e = doc.find(end_text, c.selectionStart())
            c.setPosition(c.selectionStart())
            c.setPosition(e.selectionEnd(), QTextCursor.KeepAnchor)
        w.viewer.setTextCursor(c)

    select("RAVI KUMAR")
    assert w.make_field_from_selection(FieldSpec("SuspectName", "text"), True)
    select("SANJU", "500018")
    assert w.make_field_from_selection(FieldSpec("Address", "multiline"), True)
    select("Signed")
    assert w.make_field_from_selection(FieldSpec("OfficerName", "text"), True)
    assert w.fields["OfficerName"].source == "settings"
    w._next()
    w._next()   # save
    assert w.saved_spec and [f.name for f in w.saved_spec.file_fields] == ["SuspectName", "Address"]
    lib = TemplateLibrary(store=store)
    assert lib.table.rowCount() == 1 and lib.current().id == w.saved_spec.id
    app.processEvents()
    time.sleep(0)
