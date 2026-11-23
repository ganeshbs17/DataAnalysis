import re
import zipfile

from docx import Document

from notice_generator.core.notice_builder import build_context
from notice_generator.importers.prepared_importer import import_prepared
from notice_generator.templates.placeholder_parser import find_placeholders, find_unresolved
from notice_generator.templates.renderer import render_docx


def _text(docx):
    d = Document(str(docx))
    parts = [p.text for p in d.paragraphs]
    parts += [c.text for t in d.tables for row in t.rows for c in row.cells]
    parts += [p.text for p in d.sections[0].header.paragraphs]
    return "\n".join(parts)


def test_placeholders_found_including_split_runs(template):
    ph = find_placeholders(template)
    assert {"SuspectName", "Address", "OfficerName", "AmountDeposited"} <= ph
    assert len(ph) == 11


def test_render_values(sample_xlsx, settings, template, tmp_path):
    rec = import_prepared(sample_xlsx).records[2]
    out = tmp_path / "out.docx"
    assert render_docx(template, build_context(rec, settings), out) == []
    t = _text(out)
    assert "₹1,37,00,000/-" in t and t.count("₹") == 1
    assert "18.02.2025, 20.02.2025, 21.02.2025 and 25.02.2025" in t
    assert "02/12/2030 at 11:00 AM" in t
    assert "Account No: 20100036980152" in t
    assert "Mr.KUMAR DEMO" in t
    assert "Header for Officer A" in t and "Officer: Officer A" in t


def test_split_run_keeps_formatting(sample_xlsx, settings, template, tmp_path):
    rec = import_prepared(sample_xlsx).records[0]
    out = tmp_path / "out.docx"
    render_docx(template, build_context(rec, settings), out)
    d = Document(str(out))
    p = next(p for p in d.paragraphs if "RAVI SAMPLE" in p.text)
    assert p.text == "RAVI SAMPLE"


def test_address_line_breaks_and_escaping(sample_xlsx, settings, template, tmp_path):
    rec = import_prepared(sample_xlsx).records[2]
    out = tmp_path / "out.docx"
    render_docx(template, build_context(rec, settings), out)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    m = re.search(r"<w:p>(?:(?!</w:p>).)*DEMO &amp; SONS(?:(?!</w:p>).)*</w:p>", xml, re.S)
    assert m, "address paragraph with escaped & not found"
    para = m.group(0)
    assert para.count("<w:br/>") == 1                 # 2 lines, one break, no leading blank line
    assert "PLOT 5 &lt;B&gt;, CITY" in para
    Document(str(out))  # opens without error


def test_unresolved_placeholder_detected(sample_xlsx, settings, template_unknown, tmp_path):
    rec = import_prepared(sample_xlsx).records[0]
    out = tmp_path / "out.docx"
    assert render_docx(template_unknown, build_context(rec, settings), out) == ["CaseFileCode"]
    assert find_unresolved(out) == ["CaseFileCode"]
