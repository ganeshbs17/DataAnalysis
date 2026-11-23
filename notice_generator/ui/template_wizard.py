"""Create-Template wizard: open a notice, select text with the mouse, name it as a field.

Step 1  choose the notice (.docx), template name and description
Step 2  mark fields: select text → "Make field" → name, type, required
Step 3  review fields and checks → save template + matching Excel data sheet
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QTextBlockFormat, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog,
                               QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QMenu,
                               QListWidgetItem, QMessageBox, QPushButton, QSplitter, QStackedWidget,
                               QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget)

from ..templates.docx_editor import FIELD_NAME_RE, LINE_SEP, PLACEHOLDER_RE, DocxEditor, EditError, run_style
from ..templates.template_store import (SOURCE_LABELS, STANDARD_TYPES, TYPE_LABELS, FieldSpec, TemplateSpec,
                                        TemplateStore, guess_type, source_of)

FIXED_DATE_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b")
HIGHLIGHT = QColor("#ffe58a")


class FieldDialog(QDialog):
    """Name a selected piece of text as a field."""

    def __init__(self, parent, selected: str, occurrences: int, fields: dict[str, FieldSpec]):
        super().__init__(parent)
        self.setWindowTitle("Make field")
        self.fields = fields
        shown = selected.replace(LINE_SEP, " ⏎ ").replace("\n", " ⏎ ")
        lbl = QLabel(f"Selected text:<br><b>{_esc(shown[:300])}</b>")
        lbl.setWordWrap(True)
        self.name = QComboBox(editable=True)
        names = list(dict.fromkeys(list(fields) + list(STANDARD_TYPES)))
        self.name.addItems(names)
        self.name.setCurrentText("")
        self.name.lineEdit().setPlaceholderText("e.g. SuspectName or CaseNo")
        self.type = QComboBox()
        for k, v in TYPE_LABELS.items():
            self.type.addItem(v, k)
        self.required = QCheckBox("Required (a notice is blocked if this is empty)", checked=True)
        self.replace_all = QCheckBox(f"Also replace the other {occurrences - 1} place(s) where this exact "
                                     "text appears", checked=True)
        self.replace_all.setVisible(occurrences > 1)
        self.source = QLabel()
        self.name.currentTextChanged.connect(self._name_changed)
        form = QFormLayout()
        form.addRow("Field name:", self.name)
        form.addRow("Type:", self.type)
        form.addRow("Filled from:", self.source)
        form.addRow("", self.required)
        form.addRow("", self.replace_all)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        v = QVBoxLayout(self)
        v.addWidget(lbl)
        v.addLayout(form)
        hint = QLabel("<small>Use the standard names (SuspectName, Address, AmountDeposited …) where they fit; "
                      "any other name becomes a new column in the Excel data sheet.</small>")
        hint.setWordWrap(True)
        v.addWidget(hint)
        v.addWidget(bb)
        self.resize(520, 300)

    def _name_changed(self, name: str):
        name = name.strip()
        existing = self.fields.get(name)
        t = existing.type if existing else guess_type(name)
        self.type.setCurrentIndex(max(0, self.type.findData(t)))
        self.type.setEnabled(existing is None)
        src = source_of(name)
        self.source.setText(SOURCE_LABELS[src] + ("" if src == "file" else " (not a column in the Excel sheet)"))

    def _accept(self):
        name = self.name.currentText().strip()
        if not FIELD_NAME_RE.match(name):
            QMessageBox.warning(self, "Field name", "Use only letters, digits and _ (no spaces), "
                                "starting with a letter. Example: CaseNo")
            return
        self.accept()

    def spec(self) -> FieldSpec:
        name = self.name.currentText().strip()
        return FieldSpec(name, self.type.currentData(), self.required.isChecked(), source_of(name))


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class TemplateWizard(QDialog):
    def __init__(self, parent=None, store: TemplateStore | None = None, edit: TemplateSpec | None = None):
        super().__init__(parent)
        self.store = store or TemplateStore()
        self.edit = edit
        self.editor: DocxEditor | None = None
        self.fields: dict[str, FieldSpec] = {}
        self.block_pid: dict[int, int | None] = {}
        self.saved_spec: TemplateSpec | None = None
        self.setWindowTitle("Edit template fields" if edit else "Create new template")
        self.resize(1250, 820)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._page_source())
        self.stack.addWidget(self._page_fields())
        self.stack.addWidget(self._page_review())
        self.step_label = QLabel()
        self.btn_back = QPushButton("< Back")
        self.btn_next = QPushButton("Next >")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_back.clicked.connect(lambda: self._go(self.stack.currentIndex() - 1))
        self.btn_next.clicked.connect(self._next)
        self.btn_cancel.clicked.connect(self.reject)
        nav = QHBoxLayout()
        nav.addWidget(self.step_label, 1)
        nav.addWidget(self.btn_cancel)
        nav.addWidget(self.btn_back)
        nav.addWidget(self.btn_next)
        v = QVBoxLayout(self)
        v.addWidget(self.stack, 1)
        v.addLayout(nav)
        if edit:
            self.path.setText(str(edit.docx))
            self.path.setEnabled(False)
            self.btn_browse.setEnabled(False)
            self.tname.setText(edit.name)
            self.desc.setText(edit.description)
            self.fields = {f.name: f for f in edit.fields}
        self._go(0)

    # ------------------------------------------------------------------ pages
    def _page_source(self) -> QWidget:
        w = QWidget()
        self.path = QLineEdit(placeholderText="Choose a notice in Word format (.docx)")
        self.btn_browse = QPushButton("Browse…")
        self.btn_browse.clicked.connect(self._browse)
        row = QHBoxLayout()
        row.addWidget(self.path, 1)
        row.addWidget(self.btn_browse)
        self.tname = QLineEdit(placeholderText="e.g. BNSS 35(3) notice – bank credit")
        self.desc = QLineEdit(placeholderText="Optional")
        form = QFormLayout()
        form.addRow("Notice document:", row)
        form.addRow("Template name:", self.tname)
        form.addRow("Description:", self.desc)
        info = QLabel(
            "<p>Choose an existing notice — for example one already filled in for a real case. "
            "In the next step you select the parts that change from notice to notice (name, address, "
            "amount, account number, dates…) and give each one a field name.</p>"
            "<p>The chosen file is never changed. The template is saved as a copy, together with an Excel "
            "data sheet whose column headings are the field names.</p>")
        info.setWordWrap(True)
        v = QVBoxLayout(w)
        v.addWidget(info)
        v.addLayout(form)
        v.addStretch(1)
        return w

    def _page_fields(self) -> QWidget:
        w = QWidget()
        self.viewer = QTextEdit()
        self.viewer.setReadOnly(True)
        self.viewer.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        self.viewer.setLineWrapMode(QTextEdit.WidgetWidth)
        self.viewer.setStyleSheet("QTextEdit { background: white; color: black; }")
        self.btn_make = QPushButton("Make field from selection")
        self.btn_make.setStyleSheet("font-weight: bold; padding: 6px;")
        # A no-argument slot: clicked(bool) must not reach make_field_from_selection(spec=...).
        self.btn_make.clicked.connect(self.on_make_field)
        self.btn_make.setEnabled(False)
        self.viewer.selectionChanged.connect(
            lambda: self.btn_make.setEnabled(self.viewer.textCursor().hasSelection()))
        self.viewer.setContextMenuPolicy(Qt.CustomContextMenu)
        self.viewer.customContextMenuRequested.connect(self._viewer_menu)
        self.field_list = QListWidget()
        self.field_list.itemDoubleClicked.connect(self._edit_field)
        btn_edit = QPushButton("Change type…")
        btn_edit.clicked.connect(lambda: self._edit_field(self.field_list.currentItem()))
        btn_remove = QPushButton("Remove field")
        btn_remove.clicked.connect(self._remove_field)
        self.btn_undo = QPushButton("Undo")
        self.btn_undo.clicked.connect(self._undo)
        side = QWidget()
        sv = QVBoxLayout(side)
        sv.addWidget(self.btn_make)
        sv.addWidget(QLabel("<b>Fields</b>"))
        sv.addWidget(self.field_list, 1)
        bh = QHBoxLayout()
        bh.addWidget(btn_edit)
        bh.addWidget(btn_remove)
        sv.addLayout(bh)
        sv.addWidget(self.btn_undo)
        help_ = QLabel("<small>Select text with the mouse, then click <b>Make field</b>. A selection may cover "
                       "several lines (e.g. a whole address). Fields are shown in yellow.</small>")
        help_.setWordWrap(True)
        sv.addWidget(help_)
        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.viewer)
        split.addWidget(side)
        split.setSizes([900, 330])
        v = QVBoxLayout(w)
        v.addWidget(split)
        return w

    def _page_review(self) -> QWidget:
        w = QWidget()
        self.review = QTableWidget(0, 5)
        self.review.setHorizontalHeaderLabels(["Field", "Type", "Required", "Filled from", "Example (original text)"])
        self.review.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.review.horizontalHeader().setStretchLastSection(True)
        self.review.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.checks = QLabel()
        self.checks.setWordWrap(True)
        v = QVBoxLayout(w)
        v.addWidget(QLabel("<b>Review the fields.</b> Fields filled from the <i>Excel sheet</i> become its column "
                           "headings; the others are chosen on the notice settings screen."))
        v.addWidget(self.review, 1)
        v.addWidget(self.checks)
        return w

    # ------------------------------------------------------------------ navigation
    def _go(self, i: int):
        i = max(0, min(2, i))
        if i == 2:
            self._fill_review()
        self.stack.setCurrentIndex(i)
        self.btn_back.setEnabled(i > 0)
        self.btn_next.setText("Save template" if i == 2 else "Next >")
        self.step_label.setText(["Step 1 of 3: choose the notice",
                                 "Step 2 of 3: mark the fields",
                                 "Step 3 of 3: review and save"][i])

    def _browse(self):
        p, _ = QFileDialog.getOpenFileName(self, "Choose notice", "", "Word documents (*.docx)")
        if p:
            self.path.setText(p)
            if not self.tname.text():
                self.tname.setText(Path(p).stem.replace("_", " "))

    def _next(self):
        i = self.stack.currentIndex()
        if i == 0:
            if not self.load(self.path.text().strip(), self.tname.text().strip()):
                return
        elif i == 1:
            if not self.editor.placeholders():
                QMessageBox.information(self, "Fields", "Mark at least one field before continuing.")
                return
        elif i == 2:
            self.save()
            return
        self._go(i + 1)

    def load(self, path: str, name: str) -> bool:
        if not path or not Path(path).is_file():
            QMessageBox.warning(self, "Notice", "Choose a Word (.docx) notice first.")
            return False
        if not name:
            QMessageBox.warning(self, "Template name", "Give the template a name.")
            return False
        if self.editor is None or self.editor.source != Path(path):
            try:
                self.editor = DocxEditor(path)
            except Exception as e:  # noqa: BLE001
                QMessageBox.warning(self, "Notice", f"The document could not be opened: {e}\n\n"
                                    "Only .docx files are supported. Save .doc files as .docx in Word first.")
                return False
            for n in self.editor.placeholders():  # already-marked fields in the document
                self.fields.setdefault(n, FieldSpec(n, guess_type(n), True, source_of(n)))
            self.render()
        return True

    # ------------------------------------------------------------------ viewer
    def render(self):
        sb = self.viewer.verticalScrollBar()
        pos = sb.value()
        doc = self.viewer.document()
        doc.clear()
        cur = QTextCursor(doc)
        self.block_pid = {}
        region = None
        first = True
        heading = QTextCharFormat()
        heading.setForeground(QColor("#777"))
        heading.setFontItalic(True)
        for p in self.editor.paragraphs:
            if p.region != region:
                region = p.region
                if region != "Body":
                    if not first:
                        cur.insertBlock()
                    first = False
                    bf = QTextBlockFormat()
                    bf.setTopMargin(14)
                    cur.setBlockFormat(bf)
                    cur.insertText(f"──────── {region} (printed on the page {region.lower()}) ────────", heading)
                    self.block_pid[cur.block().blockNumber()] = None
            if not first:
                cur.insertBlock()
            first = False
            bf = QTextBlockFormat()
            bf.setAlignment({"center": Qt.AlignHCenter, "right": Qt.AlignRight, "end": Qt.AlignRight,
                             "both": Qt.AlignJustify}.get(p.align, Qt.AlignLeft))
            if p.in_table:
                bf.setLeftMargin(18)
                bf.setBackground(QColor("#f3f6fa"))
            cur.setBlockFormat(bf)
            self.block_pid[cur.block().blockNumber()] = p.pid
            for s in p.segments:
                cf = QTextCharFormat()
                st = run_style(s.run)
                if st["bold"]:
                    cf.setFontWeight(QFont.Bold)
                cf.setFontItalic(st["italic"])
                cf.setFontUnderline(st["underline"])
                if st["size"]:
                    cf.setFontPointSize(st["size"])
                cf.setForeground(QColor("black"))
                text = (s.elem.text or "") if s.kind == "t" else (LINE_SEP if s.kind == "br" else "\t")
                if text:
                    cur.insertText(text, cf)
            block_start = cur.block().position()
            for m in PLACEHOLDER_RE.finditer(p.text):
                hc = QTextCursor(doc)
                hc.setPosition(block_start + m.start())
                hc.setPosition(block_start + m.end(), QTextCursor.KeepAnchor)
                f = QTextCharFormat()
                f.setBackground(HIGHLIGHT)
                f.setFontWeight(QFont.Bold)
                hc.mergeCharFormat(f)
        self._refresh_field_list()
        sb.setValue(pos)
        self.btn_undo.setEnabled(self.editor.can_undo())

    def _refresh_field_list(self):
        present = self.editor.placeholders() if self.editor else []
        for n in present:
            self.fields.setdefault(n, FieldSpec(n, guess_type(n), True, source_of(n)))
        for n in list(self.fields):
            if n not in present:
                del self.fields[n]
        self.field_list.clear()
        for n in present:
            f = self.fields[n]
            it = QListWidgetItem(f"{n}  —  {TYPE_LABELS[f.type]}  ·  {SOURCE_LABELS[f.source]}"
                                 + ("" if f.required else "  · optional"))
            it.setData(Qt.UserRole, n)
            self.field_list.addItem(it)

    def selection_range(self):
        """(pid_start, start, pid_end, end, text) of the viewer selection, or raise EditError."""
        c = self.viewer.textCursor()
        if not c.hasSelection():
            raise EditError("Select some text in the notice first.")
        doc = self.viewer.document()
        a, b = c.selectionStart(), c.selectionEnd()
        ba, bb = doc.findBlock(a), doc.findBlock(b)
        if bb.blockNumber() > ba.blockNumber() and b == bb.position():
            bb = bb.previous()  # selection ends at the very start of the next paragraph
            b = bb.position() + bb.length() - 1
        pa, pb = self.block_pid.get(ba.blockNumber()), self.block_pid.get(bb.blockNumber())
        if pa is None or pb is None:
            raise EditError("The selection includes a section label. Select only text from the notice.")
        return pa, a - ba.position(), pb, b - bb.position(), c.selectedText()

    def _viewer_menu(self, pos):
        menu = QMenu(self)
        act = menu.addAction("Make field from selection")
        act.setEnabled(self.viewer.textCursor().hasSelection())
        act.triggered.connect(self.on_make_field)
        menu.addSeparator()
        menu.addAction("Copy").triggered.connect(self.viewer.copy)
        menu.exec(self.viewer.mapToGlobal(pos))

    def on_make_field(self, *_):
        """Button / context-menu slot: ask for the field details, then mark the selection."""
        try:
            self.make_field_from_selection()
        except Exception as e:  # noqa: BLE001 - never fail silently in the windowed exe
            QMessageBox.warning(self, "Make field", f"The field could not be created: {e}")

    def make_field_from_selection(self, spec: FieldSpec | None = None, replace_all: bool | None = None):
        try:
            pa, sa, pb, sb, text = self.selection_range()
        except EditError as e:
            QMessageBox.information(self, "Make field", str(e))
            return None
        if "{{" in text or "}}" in text:
            QMessageBox.information(self, "Make field", "The selection already contains a field.")
            return None
        single = self.editor.paragraphs[pa].text[sa:sb] if pa == pb else None
        occurrences = self.editor.count_occurrences(single) if single else 1
        if spec is None:
            dlg = FieldDialog(self, text, occurrences, self.fields)
            if dlg.exec() != QDialog.Accepted:
                return None
            spec, replace_all = dlg.spec(), dlg.replace_all.isChecked()
        try:
            self.editor.make_field(pa, sa, pb, sb, spec.name)
            if replace_all and single and occurrences > 1:
                self.editor.replace_all(single, spec.name)
        except EditError as e:
            QMessageBox.information(self, "Make field", str(e))
            return None
        spec.source = source_of(spec.name)  # settings/computed fields never become Excel columns
        spec.example = spec.example or self.editor.originals.get(spec.name, "")[:80]
        if spec.name in self.fields:
            spec.type = self.fields[spec.name].type
            spec.example = self.fields[spec.name].example or spec.example
        self.fields[spec.name] = spec
        self.render()
        return spec

    def _edit_field(self, item):
        if item is None:
            return
        name = item.data(Qt.UserRole)
        f = self.fields[name]
        dlg = FieldDialog(self, f"{{{{{name}}}}}", 1, {})
        dlg.name.setCurrentText(name)
        dlg.name.setEnabled(False)
        dlg.type.setEnabled(True)
        dlg.type.setCurrentIndex(max(0, dlg.type.findData(f.type)))
        dlg.required.setChecked(f.required)
        if dlg.exec() == QDialog.Accepted:
            f.type, f.required = dlg.type.currentData(), dlg.required.isChecked()
            self._refresh_field_list()

    def _remove_field(self):
        it = self.field_list.currentItem()
        if it is None:
            return
        self.editor.remove_field(it.data(Qt.UserRole))
        self.render()

    def _undo(self):
        self.editor.undo()
        self.render()

    # ------------------------------------------------------------------ review & save
    def ordered_fields(self) -> list[FieldSpec]:
        return [self.fields[n] for n in self.editor.placeholders() if n in self.fields]

    def preflight(self) -> list[str]:
        msgs = []
        dates = []
        for p in self.editor.paragraphs:
            plain = PLACEHOLDER_RE.sub("", p.text)
            if "{{" in plain or "}}" in plain:
                msgs.append(f"⚠ Incomplete field braces in: “{plain.strip()[:70]}”")
            dates += [(p.region, d) for d in FIXED_DATE_RE.findall(plain)]
        if dates:
            shown = ", ".join(f"{d} ({r.lower()})" for r, d in dates[:6])
            msgs.append(f"ℹ Fixed dates remain in the text and will print on every notice: {shown}. "
                        "Make them fields (e.g. NoticeDate) if they should change.")
        if not any(f.source == "file" for f in self.fields.values()):
            msgs.append("⚠ No field comes from the Excel sheet, so every notice would be identical.")
        return msgs

    def _fill_review(self):
        fields = self.ordered_fields()
        self.review.setRowCount(len(fields))
        for r, f in enumerate(fields):
            for c, val in enumerate([f.name, TYPE_LABELS[f.type], "Yes" if f.required else "No",
                                     SOURCE_LABELS[f.source], f.example.replace("\n", " ⏎ ")]):
                self.review.setItem(r, c, QTableWidgetItem(val))
        msgs = self.preflight()
        n_cols = sum(1 for f in fields if f.source == "file")
        self.checks.setText(f"<p>The Excel data sheet will have <b>{n_cols}</b> column(s).</p>"
                            + "".join(f"<p>{_esc(m)}</p>" for m in msgs))

    def save(self) -> TemplateSpec | None:
        fields = self.ordered_fields()
        try:
            spec = self.store.save(self.tname.text().strip(), self.editor.to_bytes(), fields,
                                   self.desc.text().strip(), source_doc=Path(self.path.text()).name,
                                   template_id=self.edit.id if self.edit else None)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Save", f"The template could not be saved: {e}")
            return None
        self.saved_spec = spec
        box = QMessageBox(self)
        box.setWindowTitle("Template saved")
        box.setText(f"“{spec.name}” (version {spec.version}) was saved with {len(fields)} field(s).\n\n"
                    "Its Excel data sheet is ready. You can download it now, or later from Templates.")
        dl = box.addButton("Download data sheet…", QMessageBox.AcceptRole)
        box.addButton("Close", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is dl:
            download_copy(self, spec.data_sheet, f"{spec.name}_data_sheet.xlsx", "Excel workbook (*.xlsx)")
        self.accept()
        return spec


def download_copy(parent, src: Path, default_name: str, filt: str) -> Path | None:
    safe = re.sub(r'[<>:"/\\|?*]+', "_", default_name)
    dest, _ = QFileDialog.getSaveFileName(parent, "Save a copy", str(Path.home() / safe), filt)
    if not dest:
        return None
    shutil.copyfile(src, dest)
    return Path(dest)
