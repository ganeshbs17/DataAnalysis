"""End-to-end batch tests. Need LibreOffice; skipped when it is not installed."""
import pytest

from notice_generator.exporters.batch_exporter import ExportOptions, run
from notice_generator.exporters.combiner import page_count
from notice_generator.exporters.pdf_engine import find_soffice
from notice_generator.importers.prepared_importer import import_prepared
from notice_generator.storage.file_manager import notice_basename, safe_name, unique_path

needs_lo = pytest.mark.skipif(find_soffice() is None, reason="LibreOffice not installed")


def _records(path):
    recs = import_prepared(path).records
    for r in recs:
        r.confirmations.update({"V-03:BankAccountNumber", "V-03:SuspectAccountNumber"})
    return recs


@needs_lo
def test_batch_generates_and_isolates_failures(sample_xlsx, settings, template, tmp_path):
    recs = _records(sample_xlsx)
    recs[4].include = False
    res = run(recs, settings, template, tmp_path, ExportOptions(duplex=False))
    by = {it.record.serial: it for it in res.items}
    assert by[6].status == "blocked"            # missing name (V-01); batch continued
    assert by[5].status == "excluded"
    gen = [it for it in res.items if it.status == "generated"]
    assert len(gen) == 4
    for it in gen:
        assert it.docx.exists() and it.pdf.exists() and it.pages >= 1
    assert it.docx.name.startswith(f"{it.record.serial:03d}_")
    assert res.combined_pages == sum(it.pages for it in gen) == page_count(res.combined)


@needs_lo
def test_duplex_padding(sample_xlsx, settings, template, tmp_path):
    recs = _records(sample_xlsx)[:2]
    res = run(recs, settings, template, tmp_path, ExportOptions(docx=False, duplex=True))
    pages = [it.pages for it in res.items]
    assert res.combined_pages == sum(p + (p % 2) for p in pages)
    assert not (res.folder / "DOCX").exists()


@needs_lo
def test_unresolved_placeholder_fails_record(sample_xlsx, settings, template_unknown, tmp_path):
    recs = _records(sample_xlsx)[:1]
    res = run(recs, settings, template_unknown, tmp_path, ExportOptions(pdf=False, combined=False),
              revalidate=False)
    assert res.items[0].status == "failed" and "V-14" in res.items[0].error
    assert not list((res.folder / "DOCX").glob("*.docx"))


@needs_lo
def test_no_overwrite(sample_xlsx, settings, template, tmp_path):
    recs = _records(sample_xlsx)[:1]
    opts = ExportOptions(pdf=False, combined=False, batch_subfolder=False)
    a = run(recs, settings, template, tmp_path, opts).items[0].docx
    b = run(recs, settings, template, tmp_path, opts).items[0].docx
    assert a != b and b.name.endswith(" (2).docx") and a.exists()


def test_file_naming(tmp_path):
    assert safe_name('Mr. A/B:C*"?') == "Mr._ABC"
    assert notice_basename(3, "Mr. X Y", "51910198214") == "003_Mr._X_Y_8214"
    assert safe_name("CON") == "_CON"
    (tmp_path / "a.pdf").write_text("x")
    assert unique_path(tmp_path, "a", ".pdf").name == "a (2).pdf"
