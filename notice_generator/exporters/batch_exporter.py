"""Batch generation: DOCX → PDF → combined PDF (FR-09).

A notice is reported as *generated* only after its DOCX has been written, its
PDF exists and opens, and its page count has been checked. One failed record
does not stop the batch.
"""
from __future__ import annotations

import datetime as dt
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from ..core.models import Issue, NoticeSettings, Record, Severity, Status
from ..core.notice_builder import build_context, build_values
from ..core.validation import validate_batch
from ..storage.file_manager import notice_basename, unique_path
from ..templates.placeholder_parser import expected_page_count, find_placeholders
from ..templates.renderer import render_docx
from ..templates.template_store import spec_for_docx
from .combiner import combine, ends_with_blank_page, page_count
from .pdf_engine import LibreOfficeEngine, PdfEngineError, get_engine

Progress = Callable[[int, int, str], None]


@dataclass
class ExportOptions:
    docx: bool = True
    pdf: bool = True
    combined: bool = True
    duplex: bool = False
    batch_subfolder: bool = True
    expected_pages: int | None = None   # None → template metadata


@dataclass
class ItemResult:
    record: Record
    status: str                      # generated | failed | blocked | excluded
    docx: Path | None = None
    pdf: Path | None = None
    pages: int | None = None
    warnings: list[Issue] = field(default_factory=list)
    error: str = ""


@dataclass
class BatchResult:
    batch_id: str
    folder: Path
    items: list[ItemResult]
    combined: Path | None = None
    combined_pages: int | None = None
    expected_pages: int | None = None
    error: str = ""

    def count(self, status: str) -> int:
        return sum(1 for i in self.items if i.status == status)


def run(records: list[Record], settings: NoticeSettings, template: str | Path, out_dir: str | Path,
        options: ExportOptions | None = None, engine: LibreOfficeEngine | None = None,
        progress: Progress | None = None, revalidate: bool = True) -> BatchResult:
    opts = options or ExportOptions()
    template = Path(template)
    out_dir = Path(out_dir)
    progress = progress or (lambda d, t, m: None)
    batch_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    placeholders = find_placeholders(template)
    spec = spec_for_docx(template)  # field types when the template comes from the library
    if revalidate:
        validate_batch(records, settings, template_placeholders=placeholders, spec=spec)
    expected = opts.expected_pages or (spec.expected_pages if spec else None) or expected_page_count(template)

    folder = out_dir / f"batch_{batch_id}" if opts.batch_subfolder else out_dir
    n = 2
    while opts.batch_subfolder and folder.exists():
        folder = out_dir / f"batch_{batch_id}_{n}"
        n += 1
    folder.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix=".work_", dir=folder))
    docx_dir = folder / "DOCX" if opts.docx else work / "docx"
    pdf_dir = folder / "PDF"
    docx_dir.mkdir(parents=True, exist_ok=True)

    items: list[ItemResult] = []
    to_make: list[ItemResult] = []
    for r in records:
        st = r.status
        if st == Status.EXCLUDED:
            items.append(ItemResult(r, "excluded"))
        elif st == Status.BLOCKED:
            items.append(ItemResult(r, "blocked",
                                    error="; ".join(f"{i.rule} {i.message}" for i in r.blocking_issues())))
        else:
            it = ItemResult(r, "pending", warnings=list(r.warnings()))
            items.append(it)
            to_make.append(it)

    total = len(to_make) * (2 if opts.pdf else 1) + (1 if opts.combined and opts.pdf else 0)
    done = 0

    # 1. DOCX
    for it in to_make:
        r = it.record
        vals = build_values(r, settings, spec)
        name_f = "SuspectName" if vals.get("SuspectName") else next(
            (f.name for f in (spec.file_fields if spec else []) if f.type == "text" and vals.get(f.name)), "")
        acc_f = "SuspectAccountNumber" if vals.get("SuspectAccountNumber") else next(
            (n for n in (spec.names_of_type("account") if spec else []) if vals.get(n)), "")
        stem = notice_basename(r.serial, vals.get(name_f, "") or "notice", vals.get(acc_f, ""))
        target = unique_path(docx_dir, stem, ".docx")
        try:
            unresolved = render_docx(template, build_context(r, settings, spec), target)
            if unresolved:
                target.unlink(missing_ok=True)
                it.status = "failed"
                it.error = "V-14 Unresolved placeholder(s): " + ", ".join(sorted(set(unresolved)))
            else:
                it.docx = target
                it.status = "docx"
        except Exception as e:  # noqa: BLE001 - plain-language error per record
            target.unlink(missing_ok=True)
            it.status = "failed"
            it.error = f"Could not create the Word document: {e}"
        done += 1
        progress(done, total, f"Created document {done} of {len(to_make)}")

    # 2. PDF (one LibreOffice call for the whole batch)
    made = [it for it in to_make if it.status == "docx"]
    if opts.pdf and made:
        progress(done, total, "Converting to PDF…")
        eng = engine or get_engine()
        try:
            conv = eng.convert([it.docx for it in made], work / "pdf")
        except PdfEngineError as e:
            conv = {}
            for it in made:
                it.status, it.error = "failed", str(e)
        pdf_dir.mkdir(parents=True, exist_ok=True)
        for it in made:
            if it.status == "failed":
                continue
            src = conv.get(it.docx)
            if src is None:
                it.status, it.error = "failed", "PDF conversion failed. Retry the record."
            else:
                try:
                    pages = page_count(src)
                    blank = ends_with_blank_page(src)
                except Exception:  # noqa: BLE001
                    it.status, it.error = "failed", "The PDF was created but cannot be opened."
                else:
                    dest = unique_path(pdf_dir, it.docx.stem, ".pdf")
                    shutil.move(str(src), dest)
                    it.pdf, it.pages = dest, pages
                    if expected and pages > expected:
                        it.warnings.append(Issue("V-15", Severity.WARNING,
                                                 f"PDF has {pages} pages; the template expects {expected}."))
                    if blank:
                        it.warnings.append(Issue("V-15", Severity.WARNING, "The PDF ends with a blank page."))
                    it.status = "generated"
            done += 1
            progress(done, total, f"Checked PDF {done - len(to_make)} of {len(made)}")
    elif not opts.pdf:
        for it in made:
            it.status = "generated"

    if not opts.docx:
        for it in to_make:
            it.docx = None

    # 3. Combined PDF
    result = BatchResult(batch_id, folder, items, expected_pages=expected)
    pdfs = [it.pdf for it in to_make if it.status == "generated" and it.pdf]
    if opts.combined and opts.pdf and pdfs:
        comb = unique_path(folder / "Combined", f"combined_{batch_id}", ".pdf")
        result.combined_pages = combine(pdfs, comb, duplex=opts.duplex)
        result.combined = comb
        done += 1
        progress(done, total, "Combined PDF created")

    shutil.rmtree(work, ignore_errors=True)
    return result
