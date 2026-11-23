"""PDF inspection and combined PDF (FR-09, V-15)."""
from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader, PdfWriter


def page_count(pdf: str | Path) -> int:
    return len(PdfReader(str(pdf)).pages)


def ends_with_blank_page(pdf: str | Path) -> bool:
    reader = PdfReader(str(pdf))
    if not reader.pages:
        return True
    last = reader.pages[-1]
    if (last.extract_text() or "").strip():
        return False
    try:
        if len(last.images) > 0:
            return False
    except Exception:  # pragma: no cover - malformed image streams
        return False
    return True


def combine(pdfs: list[Path], out: Path, duplex: bool = False) -> int:
    """Merge PDFs; with duplex, pad so every notice starts on an odd page. Returns page count."""
    w = PdfWriter()
    for p in pdfs:
        r = PdfReader(str(p))
        for page in r.pages:
            w.add_page(page)
        if duplex and len(r.pages) % 2 == 1:
            last = r.pages[-1]
            w.add_blank_page(width=last.mediabox.width, height=last.mediabox.height)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "xb") as fh:  # never overwrite
        w.write(fh)
    return len(w.pages)
