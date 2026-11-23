"""Saved template library (SRS FR-07, FR-10 ``Templates/`` folder).

    Templates/template_config.json          index of templates
    Templates/<id>/v<N>/template.docx
    Templates/<id>/v<N>/data_sheet.xlsx     matching import sheet
    Templates/<id>/v<N>/template.json       field list, hash, metadata

Earlier versions are kept when a template is edited.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path

from ..storage.configuration import app_dir
from .placeholder_parser import expected_page_count, find_placeholders

FIELD_TYPES = ["text", "multiline", "amount", "date", "datelist", "time", "account"]
TYPE_LABELS = {"text": "Text", "multiline": "Multi-line text", "amount": "Amount", "date": "Date",
               "datelist": "Date list", "time": "Time", "account": "Account number"}

# Fields filled from the notice settings screen, not from the Excel sheet.
SETTINGS_FIELDS = {"OfficerName", "OfficerDesignation", "AppearanceDate", "AppearanceTime", "NoticeDate",
                   "CaseRef", "CaseNumber"}
COMPUTED_FIELDS = {"AmountInWords"}

STANDARD_TYPES = {
    "SuspectName": "text", "Address": "multiline", "AmountDeposited": "amount", "SuspectBankName": "text",
    "SuspectAccountNumber": "account", "BankAccountName": "text", "BankAccountNumber": "account",
    "DateOfTransfer": "datelist", "OfficerName": "text", "OfficerDesignation": "text",
    "AppearanceDate": "date", "AppearanceTime": "time", "NoticeDate": "date", "CaseRef": "text",
    "CaseNumber": "text", "AmountInWords": "text",
}
STANDARD_HELP = {
    "SuspectName": "Name of the notice recipient, as it should be printed",
    "Address": "Recipient address; line breaks are kept",
    "AmountDeposited": "Total amount, digits only (e.g. 2000000)",
    "SuspectBankName": "Bank of the account credited",
    "SuspectAccountNumber": "Recipient account number",
    "BankAccountName": "Victim's (originating) bank",
    "BankAccountNumber": "Victim's (originating) account number",
    "DateOfTransfer": "One or more dates",
}
TYPE_HELP = {
    "text": "Text",
    "multiline": "Text; press Alt+Enter in Excel for a new line",
    "amount": "Number only, no ₹ or commas (e.g. 2000000). Printed as 20,00,000/-",
    "date": "Date as dd.mm.yyyy",
    "datelist": "One or more dates as dd.mm.yyyy, separated by commas",
    "time": "Time, e.g. 11:00 AM",
    "account": "Type as text. Column is pre-formatted as Text so Excel keeps every digit",
}


def guess_type(name: str) -> str:
    if name in STANDARD_TYPES:
        return STANDARD_TYPES[name]
    n = name.lower()
    if "account" in n or n.endswith(("acno", "acc")):
        return "account"
    if "amount" in n or "sum" in n:
        return "amount"
    if "dates" in n:
        return "datelist"
    if "date" in n:
        return "date"
    if "time" in n:
        return "time"
    if "address" in n:
        return "multiline"
    return "text"


def source_of(name: str) -> str:
    if name in SETTINGS_FIELDS:
        return "settings"
    if name in COMPUTED_FIELDS:
        return "computed"
    return "file"


SOURCE_LABELS = {"file": "Excel sheet", "settings": "Notice settings", "computed": "Calculated"}


@dataclass
class FieldSpec:
    name: str
    type: str = "text"
    required: bool = True
    source: str = "file"
    example: str = ""


@dataclass
class TemplateSpec:
    id: str
    name: str
    version: int
    fields: list[FieldSpec]
    description: str = ""
    sha256: str = ""
    source_doc: str = ""
    created: str = ""
    expected_pages: int | None = None
    folder: str = ""  # runtime only

    @property
    def docx(self) -> Path:
        return Path(self.folder) / "template.docx"

    @property
    def data_sheet(self) -> Path:
        return Path(self.folder) / "data_sheet.xlsx"

    @property
    def file_fields(self) -> list[FieldSpec]:
        return [f for f in self.fields if f.source == "file"]

    def get_field(self, name: str) -> FieldSpec | None:
        return next((f for f in self.fields if f.name == name), None)

    def names_of_type(self, *types: str) -> list[str]:
        return [f.name for f in self.fields if f.type in types and f.source == "file"]

    def to_json(self) -> dict:
        d = asdict(self)
        d.pop("folder", None)
        return d

    @classmethod
    def from_json(cls, d: dict, folder: Path) -> "TemplateSpec":
        fields = [FieldSpec(**f) for f in d.get("fields", [])]
        keep = {k: v for k, v in d.items() if k in cls.__dataclass_fields__ and k not in ("fields", "folder")}
        return cls(fields=fields, folder=str(folder), **keep)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def fields_from_placeholders(names: list[str], examples: dict[str, str] | None = None) -> list[FieldSpec]:
    examples = examples or {}
    return [FieldSpec(n, guess_type(n), True, source_of(n), examples.get(n, "")[:80]) for n in names]


def spec_for_docx(docx: str | Path) -> TemplateSpec | None:
    """Template metadata stored next to a library template, if any."""
    p = Path(docx).with_name("template.json")
    if Path(docx).name != "template.docx" or not p.exists():
        return None
    try:
        return TemplateSpec.from_json(json.loads(p.read_text(encoding="utf-8")), p.parent)
    except (OSError, ValueError, TypeError):
        return None


def _atomic_write(path: Path, data: bytes) -> None:
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


class TemplateStore:
    def __init__(self, root: str | Path | None = None):
        self.root = Path(root) if root else app_dir() / "Templates"
        self.root.mkdir(parents=True, exist_ok=True)
        self.index_file = self.root / "template_config.json"

    # ------------------------------------------------------------ index
    def _index(self) -> list[dict]:
        try:
            return json.loads(self.index_file.read_text(encoding="utf-8")).get("templates", [])
        except (OSError, ValueError):
            return []

    def _save_index(self, entries: list[dict]) -> None:
        _atomic_write(self.index_file, json.dumps({"templates": entries}, indent=2,
                                                  ensure_ascii=False).encode("utf-8"))

    def list(self) -> list[TemplateSpec]:
        out = []
        for e in self._index():
            s = self.get(e["id"])
            if s:
                out.append(s)
        return out

    def get(self, template_id: str, version: int | None = None) -> TemplateSpec | None:
        entry = next((e for e in self._index() if e["id"] == template_id), None)
        if entry is None:
            return None
        v = version or entry["current_version"]
        folder = self.root / template_id / f"v{v}"
        try:
            return TemplateSpec.from_json(json.loads((folder / "template.json").read_text(encoding="utf-8")),
                                          folder)
        except (OSError, ValueError, TypeError):
            return None

    def versions(self, template_id: str) -> list[int]:
        d = self.root / template_id
        return sorted(int(p.name[1:]) for p in d.glob("v*") if p.name[1:].isdigit()) if d.exists() else []

    # ------------------------------------------------------------ save
    def save(self, name: str, docx_bytes: bytes, fields: list[FieldSpec], description: str = "",
             source_doc: str = "", template_id: str | None = None) -> TemplateSpec:
        """Create a template, or a new version of ``template_id``."""
        from .data_sheet import write_data_sheet

        entries = self._index()
        if template_id:
            entry = next(e for e in entries if e["id"] == template_id)
            version = max(self.versions(template_id) or [0]) + 1
        else:
            slug = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")[:30] or "template"
            template_id = f"{slug}_{uuid.uuid4().hex[:6]}"
            entry = {"id": template_id}
            entries.append(entry)
            version = 1
        folder = self.root / template_id / f"v{version}"
        folder.mkdir(parents=True, exist_ok=False)
        docx = folder / "template.docx"
        _atomic_write(docx, docx_bytes)
        spec = TemplateSpec(id=template_id, name=name, version=version, fields=fields, description=description,
                            sha256=sha256_file(docx), source_doc=source_doc,
                            created=dt.datetime.now().isoformat(timespec="seconds"),
                            expected_pages=expected_page_count(docx), folder=str(folder))
        write_data_sheet(spec, spec.data_sheet)
        _atomic_write(folder / "template.json",
                      json.dumps(spec.to_json(), indent=2, ensure_ascii=False).encode("utf-8"))
        entry.update({"name": name, "description": description, "current_version": version,
                      "updated": spec.created})
        self._save_index(entries)
        return spec

    def import_existing(self, path: str | Path, name: str, description: str = "") -> TemplateSpec:
        """Add a .docx that already contains ``{{Field}}`` placeholders."""
        from .docx_editor import DocxEditor

        names = DocxEditor(path).placeholders()
        if not names:
            names = sorted(find_placeholders(path))
        return self.save(name, Path(path).read_bytes(), fields_from_placeholders(names), description,
                         source_doc=Path(path).name)

    def delete(self, template_id: str) -> None:
        """Remove from the library index (files are kept on disk for traceability)."""
        self._save_index([e for e in self._index() if e["id"] != template_id])
