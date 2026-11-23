"""Turn selected text in a notice into ``{{Field}}`` placeholders (Create-Template wizard).

Works directly on the WordprocessingML parts with lxml, so images, numbering,
headers/footers and section settings are untouched. A placeholder takes the
formatting of the run where the selection starts; the selected text in other
runs is removed. Line breaks (``w:br``) and tabs inside the selection are
removed too, so a multi-line address becomes one ``{{Address}}``.

Paragraph text as exposed here uses ``\\u2028`` for a ``w:br`` and ``\\t`` for a
``w:tab``, one character each, so offsets match a Qt text block exactly.
"""
from __future__ import annotations

import copy
import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
LINE_SEP = " "
_PART_RE = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
FIELD_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


class EditError(Exception):
    """Plain-language reason a selection cannot become a field."""


@dataclass
class Segment:
    start: int
    end: int
    kind: str              # t | br | tab
    elem: etree._Element
    run: etree._Element | None


@dataclass
class Paragraph:
    pid: int
    part: str
    region: str            # Body | Header | Footer | Notes
    elem: etree._Element
    segments: list[Segment] = field(default_factory=list)
    in_table: bool = False

    @property
    def text(self) -> str:
        out = []
        for s in self.segments:
            out.append(s.elem.text or "" if s.kind == "t" else (LINE_SEP if s.kind == "br" else "\t"))
        return "".join(out)

    @property
    def align(self) -> str:
        jc = self.elem.find(f"{_w('pPr')}/{_w('jc')}")
        return jc.get(_w("val"), "left") if jc is not None else "left"


def _flag(rpr, tag: str) -> bool:
    if rpr is None:
        return False
    e = rpr.find(_w(tag))
    if e is None:
        return False
    return e.get(_w("val"), "true").lower() not in ("0", "false", "none")


def run_style(run) -> dict:
    """Bold/italic/underline/size of a run, for display only."""
    rpr = run.find(_w("rPr")) if run is not None else None
    size = None
    if rpr is not None:
        sz = rpr.find(_w("sz"))
        if sz is not None and sz.get(_w("val"), "").isdigit():
            size = int(sz.get(_w("val"))) / 2
    return {"bold": _flag(rpr, "b"), "italic": _flag(rpr, "i"), "underline": _flag(rpr, "u"), "size": size}


def _region(part: str) -> str:
    if "header" in part:
        return "Header"
    if "footer" in part:
        return "Footer"
    if "notes" in part:
        return "Notes"
    return "Body"


class DocxEditor:
    def __init__(self, path: str | Path):
        self.source = Path(path)
        with zipfile.ZipFile(self.source) as z:
            self._infos = z.infolist()
            self._data = {i.filename: z.read(i.filename) for i in self._infos}
        self.trees: dict[str, etree._Element] = {}
        for name, data in self._data.items():
            if _PART_RE.match(name):
                self.trees[name] = etree.fromstring(data)
        self.originals: dict[str, str] = {}   # field name → original text it replaced (first occurrence)
        self._undo: list[tuple[dict[str, bytes], dict[str, str]]] = []
        self.paragraphs: list[Paragraph] = []
        self._index()

    # ------------------------------------------------------------ indexing
    def _part_order(self) -> list[str]:
        names = list(self.trees)
        return sorted(names, key=lambda n: (0 if "document" in n else 1 if "header" in n else
                                            2 if "footer" in n else 3, n))

    def _index(self) -> None:
        self.paragraphs = []
        for part in self._part_order():
            root = self.trees[part]
            for p in root.iter(_w("p")):
                para = Paragraph(len(self.paragraphs), part, _region(part), p)
                para.in_table = any(a.tag == _w("tc") for a in p.iterancestors())
                pos = 0
                for el in p.iter(_w("t"), _w("br"), _w("tab"), _w("cr")):
                    owner = next(el.iterancestors(_w("p")))
                    if owner is not p:
                        continue  # belongs to a nested (text-box) paragraph
                    if el.tag == _w("tab") and el.getparent().tag != _w("r"):
                        continue  # tab stop definition in pPr, not a tab character
                    run = el.getparent() if el.getparent().tag == _w("r") else None
                    if el.tag == _w("t"):
                        n = len(el.text or "")
                        kind = "t"
                    else:
                        n = 1
                        kind = "tab" if el.tag == _w("tab") else "br"
                    para.segments.append(Segment(pos, pos + n, kind, el, run))
                    pos += n
                self.paragraphs.append(para)

    # ------------------------------------------------------------ queries
    def placeholders(self) -> list[str]:
        seen: list[str] = []
        for p in self.paragraphs:
            for name in PLACEHOLDER_RE.findall(p.text):
                if name not in seen:
                    seen.append(name)
        return seen

    def count_occurrences(self, text: str) -> int:
        if not text:
            return 0
        return sum(p.text.count(text) for p in self.paragraphs)

    # ------------------------------------------------------------ edits
    def _snapshot(self) -> None:
        self._undo.append(({k: etree.tostring(v) for k, v in self.trees.items()}, dict(self.originals)))
        del self._undo[:-30]

    def can_undo(self) -> bool:
        return bool(self._undo)

    def undo(self) -> None:
        if not self._undo:
            return
        parts, originals = self._undo.pop()
        self.trees = {k: etree.fromstring(v) for k, v in parts.items()}
        self.originals = originals
        self._index()

    @staticmethod
    def _set_text(t, text: str) -> None:
        t.text = text
        t.set(f"{{{XML_NS}}}space", "preserve")

    def _cut(self, para: Paragraph, start: int, end: int, insert: str | None) -> None:
        """Remove [start, end) from a paragraph and, if given, put ``insert`` at ``start``
        inside the first text run of the selection (so it keeps that run's formatting)."""
        hit = [s for s in para.segments if s.start < end and s.end > start]
        host = next((s for s in hit if s.kind == "t"), None)
        if insert is not None and host is None:
            host = next((s for s in reversed(para.segments) if s.kind == "t" and s.start <= start), None) \
                or next((s for s in para.segments if s.kind == "t"), None)
            if host is None:
                raise EditError("The selection has no text that can hold a field.")
            txt = host.elem.text or ""
            a = min(max(start - host.start, 0), len(txt))
            self._set_text(host.elem, txt[:a] + insert + txt[a:])
            insert = None
        for s in hit:
            if s.kind == "t":
                txt = s.elem.text or ""
                a, b = max(start, s.start) - s.start, min(end, s.end) - s.start
                ins = insert if (s is host and insert is not None) else ""
                self._set_text(s.elem, txt[:a] + ins + txt[b:])
            else:
                s.elem.getparent().remove(s.elem)

    def make_field(self, pid_start: int, start: int, pid_end: int, end: int, name: str) -> str:
        """Replace a selection (may span consecutive paragraphs of one part) with ``{{name}}``.

        Returns the original selected text (paragraph breaks as ``\\n``)."""
        if not FIELD_NAME_RE.match(name or ""):
            raise EditError("Field names may contain only letters, digits and _ and must not start with a digit.")
        if pid_end < pid_start or (pid_end == pid_start and end <= start):
            raise EditError("Select some text first.")
        paras = self.paragraphs[pid_start:pid_end + 1]
        if len({p.part for p in paras}) > 1:
            raise EditError("A field cannot run from the document body into a header or footer.")
        if len({p.elem.getparent() for p in paras}) > 1:
            raise EditError("A field cannot cross a table cell. Select text inside one cell.")
        texts = [p.text for p in paras]
        if pid_start == pid_end:
            original = texts[0][start:end]
        else:
            original = "\n".join([texts[0][start:]] + texts[1:-1] + [texts[-1][:end]])
        if "{{" in original or "}}" in original:
            raise EditError("The selection already contains a field. Remove that field first.")
        if not original.replace(LINE_SEP, "").replace("\n", "").strip():
            raise EditError("Select some text (not only spaces or line breaks).")

        self._snapshot()
        token = "{{" + name + "}}"
        if pid_start == pid_end:
            self._cut(paras[0], start, end, token)
        else:
            self._cut(paras[0], start, len(texts[0]), token)
            for p in paras[1:-1]:
                p.elem.getparent().remove(p.elem)
            last = paras[-1]
            if end >= len(texts[-1]) or not texts[-1][end:].strip():
                last.elem.getparent().remove(last.elem)
            else:
                self._cut(last, 0, end, None)
        self.originals.setdefault(name, original.replace(LINE_SEP, "\n"))
        self._index()
        return original

    def replace_all(self, text: str, name: str) -> int:
        """Make every single-paragraph occurrence of ``text`` into ``{{name}}``. Returns the count."""
        if not text:
            return 0
        n = 0
        token = "{{" + name + "}}"
        for pid in range(len(self.paragraphs)):
            pos = 0
            while True:
                p = self.paragraphs[pid]
                i = p.text.find(text, pos)
                if i < 0:
                    break
                if n == 0:
                    self.make_field(pid, i, pid, i + len(text), name)  # snapshot + validation once
                else:
                    self._cut(p, i, i + len(text), token)
                    self._index()
                pos = i + len(token)
                n += 1
        return n

    def remove_field(self, name: str) -> int:
        """Put the original text back wherever ``{{name}}`` appears. Returns the count."""
        original = self.originals.get(name, name)
        token = "{{" + name + "}}"
        self._snapshot()
        n = 0
        for p in list(self.paragraphs):
            for s in p.segments:
                if s.kind == "t" and s.elem.text and token in s.elem.text:
                    before, _, after = s.elem.text.partition(token)
                    lines = original.split("\n")
                    self._set_text(s.elem, before + lines[0])
                    anchor = s.elem
                    for line in lines[1:]:
                        br = etree.Element(_w("br"))
                        anchor.addnext(br)
                        t = etree.Element(_w("t"))
                        self._set_text(t, line)
                        br.addnext(t)
                        anchor = t
                    self._set_text(anchor, (anchor.text or "") + after)
                    n += 1
        self.originals.pop(name, None)
        self._index()
        return n

    # ------------------------------------------------------------ output
    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self._infos:
                data = self._data[info.filename]
                if info.filename in self.trees:
                    data = etree.tostring(self.trees[info.filename], xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                zi = copy.copy(info)
                zi.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(zi, data)
        return buf.getvalue()

    def save(self, path: str | Path) -> None:
        Path(path).write_bytes(self.to_bytes())
