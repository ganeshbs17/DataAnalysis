"""Find ``{{Placeholder}}`` names in a DOCX (FR-07 step 3).

Text is joined per paragraph across runs, so placeholders split across
formatting runs are found. Body, tables, headers, footers, foot/endnotes and
text boxes (paragraphs inside ``w:txbxContent``) are all scanned.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")
_PART_RE = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")


def _paragraph_texts(xml: bytes):
    root = etree.fromstring(xml)
    for p in root.iter(f"{{{W}}}p"):
        # Only direct text of this paragraph (exclude nested text-box paragraphs; they are visited too).
        texts = []
        for t in p.iter(f"{{{W}}}t"):
            owner = next(t.iterancestors(f"{{{W}}}p"))
            if owner is p:
                texts.append(t.text or "")
        yield "".join(texts)


def iter_part_texts(docx: str | Path):
    with zipfile.ZipFile(docx) as z:
        for name in z.namelist():
            if _PART_RE.match(name):
                for text in _paragraph_texts(z.read(name)):
                    yield name, text


def find_placeholders(docx: str | Path) -> set[str]:
    found: set[str] = set()
    for _part, text in iter_part_texts(docx):
        found.update(PLACEHOLDER_RE.findall(text))
    return found


def find_unresolved(docx: str | Path) -> list[str]:
    """Placeholders (or stray ``{{``/``}}``) remaining in a rendered document."""
    out = []
    for _part, text in iter_part_texts(docx):
        out.extend(PLACEHOLDER_RE.findall(text))
        if ("{{" in text or "}}" in text) and not PLACEHOLDER_RE.search(text):
            out.append(text.strip()[:60])
    return out


def expected_page_count(docx: str | Path) -> int | None:
    """Word's saved page count from docProps/app.xml (used by V-15)."""
    try:
        with zipfile.ZipFile(docx) as z:
            xml = z.read("docProps/app.xml").decode("utf-8", "replace")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None
    m = re.search(r"<(?:\w+:)?Pages>(\d+)</", xml)
    return int(m.group(1)) if m else None
