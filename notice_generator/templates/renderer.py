"""Render a DOCX template with docxtpl (FR-07 rendering rules).

* Values are XML-escaped (``autoescape=True``).
* Run formatting at each placeholder is kept; multi-line values use docxtpl
  ``Listing`` so ``\\n`` becomes ``<w:br/>`` inside the same run.
* Missing values are rendered back as ``{{Name}}`` and then detected (V-14),
  instead of silently printing nothing.
"""
from __future__ import annotations

from pathlib import Path

from docxtpl import DocxTemplate
from jinja2 import Environment, Undefined

from .placeholder_parser import find_unresolved


class _KeepUndefined(Undefined):
    def __str__(self) -> str:  # pragma: no cover - trivial
        return "{{" + str(self._undefined_name) + "}}"

    __html__ = __str__


def _env() -> Environment:
    return Environment(undefined=_KeepUndefined, autoescape=True)


def render_docx(template: str | Path, context: dict, out_path: str | Path) -> list[str]:
    """Render and save. Returns the unresolved placeholders found in the output."""
    tpl = DocxTemplate(str(template))
    tpl.render(context, jinja_env=_env(), autoescape=True)
    tpl.save(str(out_path))
    return find_unresolved(out_path)
