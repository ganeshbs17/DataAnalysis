"""Template Library: view saved templates and download their DOCX and Excel data sheet."""
from __future__ import annotations

import html
import shutil
import tempfile
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QFileDialog, QGridLayout, QHBoxLayout, QHeaderView,
                               QInputDialog, QLabel, QMessageBox, QPushButton, QSplitter, QTableWidget,
                               QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget)

from ..storage.configuration import temp_dir
from ..templates.template_store import SOURCE_LABELS, TYPE_LABELS, TemplateSpec, TemplateStore
from .template_wizard import TemplateWizard, download_copy


class TemplateLibrary(QDialog):
    def __init__(self, parent=None, store: TemplateStore | None = None, current_id: str | None = None):
        super().__init__(parent)
        self.setWindowTitle("Templates")
        self.resize(1100, 620)
        self.store = store or TemplateStore()
        self.selected_spec: TemplateSpec | None = None
        self.specs: list[TemplateSpec] = []

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Template", "Version", "Excel columns", "Saved", "Description"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.itemSelectionChanged.connect(self._show_details)
        self.table.itemDoubleClicked.connect(lambda _i: self._use())
        self.details = QTextBrowser()

        b = {}
        for key, text, slot in [
            ("use", "Use this template", self._use),
            ("sheet", "Download data sheet (Excel)…", self._download_sheet),
            ("docx", "Download template (Word)…", self._download_docx),
            ("open_sheet", "Open data sheet", lambda: self._open_copy("sheet")),
            ("open_docx", "Open template in Word", lambda: self._open_copy("docx")),
            ("edit", "Edit fields…", self._edit),
            ("new", "Create new template…", self._create),
            ("add", "Add existing template (.docx with {{fields}})…", self._add_existing),
            ("remove", "Remove from list", self._remove),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            b[key] = btn
        b["use"].setStyleSheet("font-weight: bold; padding: 6px;")
        b["new"].setStyleSheet("font-weight: bold;")
        self.buttons = b
        grid = QGridLayout()
        for i, k in enumerate(["use", "sheet", "docx", "open_sheet", "open_docx", "edit"]):
            grid.addWidget(b[k], i // 3, i % 3)
        bottom = QHBoxLayout()
        bottom.addWidget(b["new"])
        bottom.addWidget(b["add"])
        bottom.addStretch(1)
        bottom.addWidget(b["remove"])

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.addWidget(self.details, 1)
        rv.addLayout(grid)
        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.table)
        split.addWidget(right)
        split.setSizes([560, 540])
        v = QVBoxLayout(self)
        v.addWidget(split, 1)
        v.addLayout(bottom)
        self.reload(select_id=current_id)

    # ------------------------------------------------------------------
    def reload(self, select_id: str | None = None):
        self.specs = self.store.list()
        self.table.setRowCount(len(self.specs))
        for r, s in enumerate(self.specs):
            vals = [s.name, f"v{s.version}", str(len(s.file_fields)), s.created.replace("T", " ")[:16],
                    s.description]
            for c, val in enumerate(vals):
                self.table.setItem(r, c, QTableWidgetItem(val))
        row = next((i for i, s in enumerate(self.specs) if s.id == select_id), 0)
        if self.specs:
            self.table.selectRow(row)
        self._show_details()

    def current(self) -> TemplateSpec | None:
        rows = self.table.selectionModel().selectedRows()
        return self.specs[rows[0].row()] if rows else None

    def _show_details(self):
        s = self.current()
        for k in ("use", "sheet", "docx", "open_sheet", "open_docx", "edit", "remove"):
            self.buttons[k].setEnabled(s is not None)
        if s is None:
            self.details.setHtml("<p><i>No templates yet. Click <b>Create new template…</b> to make one from an "
                                 "existing notice.</i></p>")
            return
        rows = "".join(
            f"<tr><td><b>{html.escape(f.name)}</b></td><td>{TYPE_LABELS.get(f.type, f.type)}</td>"
            f"<td>{SOURCE_LABELS.get(f.source, f.source)}</td><td>{'Yes' if f.required else 'No'}</td>"
            f"<td>{html.escape(f.example.replace(chr(10), ' ⏎ ')[:50])}</td></tr>"
            for f in s.fields)
        versions = self.store.versions(s.id)
        self.details.setHtml(
            f"<h3>{html.escape(s.name)} — version {s.version}</h3>"
            f"<p>{html.escape(s.description)}</p>"
            f"<p>Made from: {html.escape(s.source_doc)}<br>Saved: {s.created.replace('T', ' ')}<br>"
            f"Expected pages: {s.expected_pages or '?'}<br>Versions kept: {', '.join(f'v{v}' for v in versions)}<br>"
            f"<small>SHA-256 {s.sha256[:16]}…</small></p>"
            "<table border=1 cellspacing=0 cellpadding=3><tr><th>Field</th><th>Type</th><th>Filled from</th>"
            f"<th>Required</th><th>Example</th></tr>{rows}</table>")

    # ------------------------------------------------------------------ actions
    def _use(self):
        s = self.current()
        if s:
            self.selected_spec = s
            self.accept()

    def _download_sheet(self):
        s = self.current()
        if s:
            if download_copy(self, s.data_sheet, f"{s.name}_data_sheet.xlsx", "Excel workbook (*.xlsx)"):
                QMessageBox.information(self, "Download", "The data sheet was saved. Fill one row per notice, "
                                        "then import it with “Import data…”.")

    def _download_docx(self):
        s = self.current()
        if s:
            download_copy(self, s.docx, f"{s.name}_template_v{s.version}.docx", "Word document (*.docx)")

    def _open_copy(self, which: str):
        s = self.current()
        if not s:
            return
        src = s.data_sheet if which == "sheet" else s.docx
        d = Path(tempfile.mkdtemp(prefix="view_", dir=temp_dir()))
        dest = d / (f"{s.name}_data_sheet.xlsx" if which == "sheet" else f"{s.name}_v{s.version}.docx")
        shutil.copyfile(src, dest)  # a copy, so the stored template is never edited by accident
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(dest)))

    def _edit(self):
        s = self.current()
        if s:
            w = TemplateWizard(self, self.store, edit=s)
            if w.exec() == QDialog.Accepted and w.saved_spec:
                self.reload(select_id=w.saved_spec.id)

    def _create(self):
        w = TemplateWizard(self, self.store)
        if w.exec() == QDialog.Accepted and w.saved_spec:
            self.reload(select_id=w.saved_spec.id)

    def _add_existing(self):
        p, _ = QFileDialog.getOpenFileName(self, "Add existing template", "", "Word documents (*.docx)")
        if not p:
            return
        name, ok = QInputDialog.getText(self, "Template name", "Name for this template:", text=Path(p).stem)
        if not ok or not name.strip():
            return
        try:
            spec = self.store.import_existing(p, name.strip())
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, "Add template", f"The template could not be added: {e}")
            return
        if not spec.fields:
            QMessageBox.information(self, "Add template", "No {{fields}} were found in this document. "
                                    "Use “Edit fields…” to mark them.")
        self.reload(select_id=spec.id)

    def _remove(self):
        s = self.current()
        if s and QMessageBox.question(self, "Remove", f"Remove “{s.name}” from the list?\n\nIts files stay on "
                                      "disk, so notices already generated can still be traced.") == QMessageBox.Yes:
            self.store.delete(s.id)
            self.reload()
