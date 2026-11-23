"""v0.1 main window (SRS §12 steps 1–9).

Import → review table with validation status → select records → choose officer,
appearance date/time → generate DOCX/PDF/combined → preview the actual PDF.
"""
from __future__ import annotations

import html
import tempfile
from collections import Counter
from pathlib import Path

from PySide6.QtCore import (QAbstractTableModel, QDate, QModelIndex, QObject, QSortFilterProxyModel, Qt,
                            QThread, QTime, QUrl, Signal)
from PySide6.QtGui import QAction, QBrush, QColor, QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDateEdit, QFileDialog, QFormLayout,
                               QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMainWindow, QMessageBox, QProgressBar, QPushButton, QSplitter,
                               QTableView, QTabWidget, QTextBrowser, QTimeEdit, QToolBar, QVBoxLayout, QWidget)

from ..core.models import NoticeSettings, Record, Severity, Status
from ..core.notice_builder import build_values
from ..core.validation import load_bank_names, validate_batch
from ..exporters.batch_exporter import BatchResult, ExportOptions
from ..exporters.batch_exporter import run as run_batch
from ..exporters.pdf_engine import find_soffice
from ..importers.prepared_importer import ImportError_, import_prepared
from ..storage.configuration import OfficerProfile, load_profiles, save_profiles, temp_dir
from ..templates.placeholder_parser import expected_page_count, find_placeholders

# (header, field or None, editable)
COLUMNS = [
    ("", None, False),
    ("#", None, False),
    ("Status", None, False),
    ("Name", "SuspectName", True),
    ("Recipient account", "SuspectAccountNumber", True),
    ("Amount (₹)", "AmountDeposited", True),
    ("Recipient bank", "SuspectBankName", True),
    ("Transfer date(s)", "DateOfTransfer", True),
    ("Victim bank", "BankAccountName", True),
    ("Victim account", "BankAccountNumber", True),
    ("Issues", None, False),
]
STATUS_COLOURS = {Status.READY: "#d7f5dd", Status.WARNING: "#fff1c2", Status.BLOCKED: "#f9d0d0",
                  Status.EXCLUDED: "#e4e4e4"}
EDIT_COLOUR = QColor("#ffe58a")
STATUS_ROLE = Qt.UserRole + 1


class RecordModel(QAbstractTableModel):
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.records: list[Record] = []
        self.settings = NoticeSettings()
        self._values: list[dict[str, str]] = []

    def set_data(self, records: list[Record], settings: NoticeSettings):
        self.beginResetModel()
        self.records, self.settings = records, settings
        self._values = [build_values(r, settings) for r in records]
        self.endResetModel()

    def refresh(self):
        self._values = [build_values(r, self.settings) for r in self.records]
        if self.records:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self.records) - 1, len(COLUMNS) - 1))

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.records)

    def columnCount(self, parent=QModelIndex()):
        return len(COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return COLUMNS[section][0]
        return None

    def flags(self, index):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() == 0:
            f |= Qt.ItemIsUserCheckable
        if COLUMNS[index.column()][2]:
            f |= Qt.ItemIsEditable
        return f

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r = self.records[index.row()]
        col = index.column()
        _h, fld, _e = COLUMNS[col]
        if role == STATUS_ROLE:
            return r.status.value
        if col == 0 and role == Qt.CheckStateRole:
            return Qt.Checked if r.include else Qt.Unchecked
        if role == Qt.DisplayRole:
            if col == 1:
                return r.serial
            if col == 2:
                return r.status.value
            if col == 10:
                c = Counter(i.rule for i in r.issues if i.severity != Severity.INFO)
                return ", ".join(sorted(c)) if c else ""
            if fld:
                return self._values[index.row()].get(fld, "").replace("\n", " ⏎ ")
        if role == Qt.EditRole and fld:
            return r.text(fld)
        if role == Qt.BackgroundRole:
            if fld and r.is_edited(fld):
                return QBrush(EDIT_COLOUR)
            if col == 2:
                return QBrush(QColor(STATUS_COLOURS[r.status]))
        if role == Qt.ToolTipRole and fld:
            src = r.raw(fld).text
            tip = f"Source: {src!r}"
            if r.is_edited(fld):
                tip += f"\nEdited: {r.text(fld)!r}"
            return tip
        return None

    def setData(self, index, value, role=Qt.EditRole):
        r = self.records[index.row()]
        if index.column() == 0 and role == Qt.CheckStateRole:
            r.include = Qt.CheckState(value) == Qt.Checked
        elif role == Qt.EditRole and COLUMNS[index.column()][1]:
            fld = COLUMNS[index.column()][1]
            if value == r.raw(fld).text:
                r.edits.pop(fld, None)
            else:
                r.edits[fld] = str(value)
        else:
            return False
        self.changed.emit()
        return True


class FilterProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.status_filter = ""
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.setFilterKeyColumn(-1)

    def set_status(self, s: str):
        self.status_filter = s
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        if self.status_filter:
            idx = self.sourceModel().index(row, 0, parent)
            if self.sourceModel().data(idx, STATUS_ROLE) != self.status_filter:
                return False
        return super().filterAcceptsRow(row, parent)


class GenerateWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, records, settings, template, out_dir, options):
        super().__init__()
        self.args = (records, settings, template, out_dir, options)

    def run(self):
        try:
            records, settings, template, out_dir, options = self.args
            res = run_batch(records, settings, template, out_dir, options,
                            progress=lambda d, t, m: self.progress.emit(d, t, m), revalidate=False)
            self.finished.emit(res)
        except Exception as e:  # noqa: BLE001 - shown in plain language
            self.failed.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Notice Generator — v0.1")
        self.resize(1500, 900)
        self.template: Path | None = None
        self.placeholders: set[str] | None = None
        self.out_dir: Path | None = None
        self.banks = load_bank_names()
        self.profiles = load_profiles()
        self._thread: QThread | None = None
        self._worker: GenerateWorker | None = None
        self.last_result: BatchResult | None = None
        self.pdf_available = find_soffice() is not None  # PDF needs LibreOffice; DOCX works without it
        self._preview_mode = False

        self.model = RecordModel(self)
        self.model.changed.connect(self.revalidate)
        self.proxy = FilterProxy(self)
        self.proxy.setSourceModel(self.model)

        self._build_toolbar()
        self._build_ui()
        self._load_profiles_into_combo()
        self.statusBar().showMessage("Import a prepared notice workbook to begin.")

    # ------------------------------------------------------------------ UI
    def _build_toolbar(self):
        tb = QToolBar("Main")
        tb.setMovable(False)
        self.addToolBar(tb)
        for text, slot in (("Import data…", self.on_import), ("Choose template…", self.on_choose_template),
                           ("Output folder…", self.on_choose_output)):
            act = QAction(text, self)
            act.triggered.connect(slot)
            tb.addAction(act)

    def _build_ui(self):
        # Left: table + detail
        self.search = QLineEdit(placeholderText="Search records…")
        self.search.textChanged.connect(self.proxy.setFilterFixedString)
        self.status_combo = QComboBox()
        self.status_combo.addItems(["All", *(s.value for s in Status)])
        self.status_combo.currentTextChanged.connect(lambda t: self.proxy.set_status("" if t == "All" else t))
        btn_all = QPushButton("Select all")
        btn_all.clicked.connect(lambda: self._set_include(True))
        btn_none = QPushButton("Select none")
        btn_none.clicked.connect(lambda: self._set_include(False))
        top = QHBoxLayout()
        top.addWidget(self.search, 1)
        top.addWidget(QLabel("Status:"))
        top.addWidget(self.status_combo)
        top.addWidget(btn_all)
        top.addWidget(btn_none)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.selectionModel().selectionChanged.connect(self.show_detail)

        self.detail = QTextBrowser()
        btn_conf = QPushButton("Confirm V-03 for selected record")
        btn_conf.setToolTip("Confirm that the account number was verified against the bank statement.")
        btn_conf.clicked.connect(lambda: self.confirm_truncation(all_records=False))
        btn_conf_all = QPushButton("Confirm V-03 for all checked records")
        btn_conf_all.clicked.connect(lambda: self.confirm_truncation(all_records=True))
        conf_row = QHBoxLayout()
        conf_row.addWidget(btn_conf)
        conf_row.addWidget(btn_conf_all)
        conf_row.addStretch(1)

        left = QWidget()
        lv = QVBoxLayout(left)
        lv.addLayout(top)
        split_v = QSplitter(Qt.Vertical)
        split_v.addWidget(self.table)
        dw = QWidget()
        dl = QVBoxLayout(dw)
        dl.setContentsMargins(0, 0, 0, 0)
        dl.addWidget(self.detail)
        dl.addLayout(conf_row)
        split_v.addWidget(dw)
        split_v.setSizes([550, 250])
        lv.addWidget(split_v)

        # Right: settings + preview
        right_tabs = QTabWidget()
        right_tabs.addTab(self._build_settings_panel(), "Notice settings")
        right_tabs.addTab(self._build_preview_panel(), "Preview && results")
        self.right_tabs = right_tabs

        split = QSplitter(Qt.Horizontal)
        split.addWidget(left)
        split.addWidget(right_tabs)
        split.setSizes([900, 600])
        self.setCentralWidget(split)

    def _build_settings_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout()
        self.template_label = QLabel("<i>No template chosen</i>")
        self.template_label.setWordWrap(True)
        form.addRow("Template:", self.template_label)

        self.officer_combo = QComboBox(editable=True)
        self.officer_combo.currentIndexChanged.connect(self._on_profile_selected)
        self.officer_combo.editTextChanged.connect(lambda _t: self.revalidate())
        self.designation = QLineEdit()
        self.designation.textChanged.connect(lambda _t: self.revalidate())
        save_prof = QPushButton("Save as profile")
        save_prof.clicked.connect(self.save_profile)
        oh = QHBoxLayout()
        oh.addWidget(self.officer_combo, 1)
        oh.addWidget(save_prof)
        form.addRow("Investigating officer:", oh)
        form.addRow("Designation:", self.designation)

        today = QDate.currentDate()
        self.notice_date = QDateEdit(today, calendarPopup=True, displayFormat="dd/MM/yyyy")
        self.app_date = QDateEdit(today.addDays(7), calendarPopup=True, displayFormat="dd/MM/yyyy")
        self.app_time = QTimeEdit(QTime(11, 0), displayFormat="hh:mm AP")
        for wdg in (self.notice_date, self.app_date):
            wdg.dateChanged.connect(lambda _d: self.revalidate())
        self.app_time.timeChanged.connect(lambda _t: self.revalidate())
        form.addRow("Notice date:", self.notice_date)
        form.addRow("Appearance date:", self.app_date)
        form.addRow("Appearance time:", self.app_time)

        self.use_file_values = QCheckBox("Use officer / appearance values from the file")
        self.remove_label = QCheckBox("Remove leading 'Address:' labels")
        self.prefix_mode = QComboBox()
        self.prefix_mode.addItems(["keep", "remove", "standardise"])
        for c in (self.use_file_values, self.remove_label):
            c.toggled.connect(lambda _c: self.revalidate())
        self.prefix_mode.currentTextChanged.connect(lambda _t: self.revalidate())
        form.addRow("", self.use_file_values)
        form.addRow("", self.remove_label)
        form.addRow("Name prefixes:", self.prefix_mode)

        self.victim_bank = QLineEdit(placeholderText="Used only where the file value is empty")
        self.victim_account = QLineEdit(placeholderText="Used only where the file value is empty")
        for le in (self.victim_bank, self.victim_account):
            le.textChanged.connect(lambda _t: self.revalidate())
        form.addRow("Victim bank:", self.victim_bank)
        form.addRow("Victim account:", self.victim_account)

        out = QGroupBox("Output")
        ol = QVBoxLayout(out)
        self.out_label = QLabel("<i>No output folder chosen</i>")
        self.out_label.setWordWrap(True)
        self.chk_docx = QCheckBox("Individual DOCX", checked=True)
        self.chk_pdf = QCheckBox("Individual PDF", checked=True)
        self.chk_comb = QCheckBox("Combined PDF", checked=True)
        self.chk_duplex = QCheckBox("Duplex padding (each notice starts on an odd page)")
        for wdg in (self.out_label, self.chk_docx, self.chk_pdf, self.chk_comb, self.chk_duplex):
            ol.addWidget(wdg)
        if not self.pdf_available:
            for c in (self.chk_pdf, self.chk_comb, self.chk_duplex):
                c.setChecked(False)
                c.setEnabled(False)
                c.setToolTip("PDF output needs LibreOffice, which is not installed on this PC.")
            note = QLabel("<i>DOCX only: LibreOffice is not installed, so PDF output is off. "
                          "Preview opens the Word document.</i>")
            note.setWordWrap(True)
            ol.addWidget(note)

        self.btn_preview = QPushButton("Preview selected record")
        self.btn_preview.clicked.connect(self.on_preview)
        self.btn_generate = QPushButton("Generate notices")
        self.btn_generate.setStyleSheet("font-weight: bold; padding: 6px;")
        self.btn_generate.clicked.connect(self.on_generate)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress_label = QLabel("")

        v = QVBoxLayout(w)
        v.addLayout(form)
        v.addWidget(out)
        v.addWidget(self.btn_preview)
        v.addWidget(self.btn_generate)
        v.addWidget(self.progress)
        v.addWidget(self.progress_label)
        v.addStretch(1)
        return w

    def _build_preview_panel(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        self.results = QListWidget()
        self.results.currentItemChanged.connect(self._on_result_selected)
        self.results.itemDoubleClicked.connect(self._open_result_file)
        self.preview_info = QLabel("")
        self.preview_info.setWordWrap(True)
        self.pdf_doc = QPdfDocument(self)
        self.pdf_view = QPdfView()
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        open_btn = QPushButton("Open output folder")
        open_btn.clicked.connect(self._open_output_folder)
        split = QSplitter(Qt.Vertical)
        split.addWidget(self.results)
        split.addWidget(self.pdf_view)
        split.setSizes([150, 650])
        v.addWidget(split, 1)
        v.addWidget(self.preview_info)
        v.addWidget(open_btn)
        return w

    # ------------------------------------------------------------------ profiles
    def _load_profiles_into_combo(self):
        self.officer_combo.blockSignals(True)
        self.officer_combo.clear()
        for p in self.profiles:
            self.officer_combo.addItem(p.name)
        self.officer_combo.blockSignals(False)
        if self.profiles:
            self._on_profile_selected(0)

    def _on_profile_selected(self, i: int):
        if 0 <= i < len(self.profiles):
            self.designation.setText(self.profiles[i].designation)
        self.revalidate()

    def save_profile(self):
        name = self.officer_combo.currentText().strip()
        if not name:
            return
        self.profiles = [p for p in self.profiles if p.name != name]
        self.profiles.append(OfficerProfile(name=name, designation=self.designation.text().strip()))
        save_profiles(self.profiles)
        self._load_profiles_into_combo()
        self.officer_combo.setCurrentText(name)

    # ------------------------------------------------------------------ settings / validation
    def current_settings(self) -> NoticeSettings:
        return NoticeSettings(
            officer_name=self.officer_combo.currentText().strip(),
            officer_designation=self.designation.text().strip(),
            appearance_date=self.app_date.date().toPython(),
            appearance_time=self.app_time.time().toPython(),
            notice_date=self.notice_date.date().toPython(),
            use_file_values=self.use_file_values.isChecked(),
            victim_bank=self.victim_bank.text(),
            victim_account=self.victim_account.text(),
            remove_address_label=self.remove_label.isChecked(),
            prefix_mode=self.prefix_mode.currentText(),
        )

    def revalidate(self):
        if not hasattr(self, "model") or not self.model.records:
            return
        s = self.current_settings()
        self.model.settings = s
        validate_batch(self.model.records, s, self.banks, self.placeholders)
        self.model.refresh()
        self.show_detail()
        c = Counter(r.status for r in self.model.records)
        self.statusBar().showMessage(
            f"{len(self.model.records)} records — " + ", ".join(f"{st.value}: {c.get(st, 0)}" for st in Status))

    def _set_include(self, value: bool):
        for r in self.model.records:
            r.include = value
        self.revalidate()

    # ------------------------------------------------------------------ actions
    def on_import(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import notice data", "",
                                              "Workbooks (*.xlsx *.xls *.csv)")
        if path:
            self.import_file(path)

    def import_file(self, path: str | Path, sheet: str | None = None):
        try:
            res = import_prepared(path, sheet=sheet)
        except ImportError_ as e:
            QMessageBox.warning(self, "Import", str(e))
            return None
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, "Import", f"The file could not be read: {e}")
            return None
        self.model.set_data(res.records, self.current_settings())
        self.table.resizeColumnsToContents()
        self.revalidate()
        mapped = len(res.mapping.field_to_col)
        msg = (f"Imported {len(res.records)} records from sheet '{res.sheet}' "
               f"(header row {res.mapping.header_row}). {mapped} of 11 columns mapped automatically; "
               f"{res.skipped_empty} empty rows skipped.")
        if len(res.sheets) > 1:
            msg += "\nOther sheets: " + ", ".join(f"{n} ({c} rows)" for n, c in res.sheets if n != res.sheet)
        if res.warnings:
            msg += "\n\nWarnings:\n• " + "\n• ".join(res.warnings)
        self.statusBar().showMessage(msg.splitlines()[0])
        self.last_import = res
        return res

    def on_choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose notice template", "", "Word documents (*.docx)")
        if path:
            self.set_template(path)

    def set_template(self, path: str | Path):
        self.template = Path(path)
        self.placeholders = find_placeholders(self.template)
        pages = expected_page_count(self.template)
        self.template_label.setText(
            f"<b>{html.escape(self.template.name)}</b><br>{len(self.placeholders)} placeholders: "
            f"{html.escape(', '.join(sorted(self.placeholders)))}<br>Expected pages: {pages or 'unknown'}")
        self.revalidate()

    def on_choose_output(self):
        d = QFileDialog.getExistingDirectory(self, "Choose output folder")
        if d:
            self.set_output(d)

    def set_output(self, d: str | Path):
        self.out_dir = Path(d)
        self.out_label.setText(html.escape(str(self.out_dir)))

    def selected_record(self) -> Record | None:
        rows = self.table.selectionModel().selectedRows() if self.table.selectionModel() else []
        if not rows:
            return None
        return self.model.records[self.proxy.mapToSource(rows[0]).row()]

    def show_detail(self, *_):
        r = self.selected_record()
        if r is None:
            self.detail.setHtml("<i>Select a record to see its values and validation results.</i>")
            return
        vals = build_values(r, self.model.settings)
        rows = []
        for f in ("SuspectName", "Address", "AmountDeposited", "SuspectBankName", "SuspectAccountNumber",
                  "BankAccountName", "BankAccountNumber", "DateOfTransfer", "OfficerName", "AppearanceDate",
                  "AppearanceTime"):
            src = html.escape(r.raw(f).text).replace("\n", "<br>")
            out = html.escape(vals.get(f, "")).replace("\n", "<br>")
            style = " style='background:#ffe58a'" if r.is_edited(f) else ""
            rows.append(f"<tr><td><b>{f}</b></td><td>{src}</td><td{style}>{out}</td></tr>")
        colours = {Severity.BLOCK: "#b00020", Severity.CONFIRMABLE: "#b00020", Severity.WARNING: "#8a6d00",
                   Severity.INFO: "#225"}
        iss = []
        for i in r.issues:
            done = " ✔ confirmed" if i.key in r.confirmations else ""
            iss.append(f"<li style='color:{colours[i.severity]}'><b>{i.rule}</b> [{i.severity.value}] "
                       f"{html.escape(i.message)}{done}</li>")
        self.detail.setHtml(
            f"<p><b>Record {r.serial}</b> — {html.escape(str(r.source_ref))} — status <b>{r.status.value}</b></p>"
            f"<ul>{''.join(iss) or '<li>No issues</li>'}</ul>"
            "<table border=1 cellspacing=0 cellpadding=3><tr><th>Field</th><th>Source</th>"
            f"<th>Printed value</th></tr>{''.join(rows)}</table>")

    def confirm_truncation(self, all_records: bool):
        recs = [r for r in self.model.records if r.include] if all_records else \
            [r for r in [self.selected_record()] if r]
        targets = [(r, i) for r in recs for i in r.issues if i.rule == "V-03" and i.key not in r.confirmations]
        if not targets:
            QMessageBox.information(self, "Confirm", "There are no unconfirmed V-03 issues in the selection.")
            return
        accounts = sorted({i.message.split(":")[0] for _r, i in targets})
        ok = QMessageBox.question(
            self, "Confirm account numbers",
            f"{len(targets)} account number(s) may have lost their last digit in Excel:\n\n"
            + "\n".join(accounts[:10]) + ("\n…" if len(accounts) > 10 else "")
            + "\n\nConfirm that you have verified them against the bank statement?")
        if ok != QMessageBox.Yes:
            return
        for r, i in targets:
            r.confirmations.add(i.key)
        self.revalidate()

    def _acknowledge_warnings(self, recs: list[Record]) -> bool:
        c = Counter(i.rule for r in recs for i in r.warnings())
        if not c:
            return True
        text = "\n".join(f"{rule}: {n} record(s)" for rule, n in sorted(c.items()))
        return QMessageBox.question(self, "Warnings", f"The batch has warnings:\n\n{text}\n\n"
                                    "Acknowledge and continue?") == QMessageBox.Yes

    def _check_ready(self) -> bool:
        if not self.model.records:
            QMessageBox.information(self, "Generate", "Import notice data first.")
            return False
        if not self.template:
            QMessageBox.information(self, "Generate", "Choose a notice template first.")
            return False
        return True

    def on_preview(self):
        if not self._check_ready():
            return
        r = self.selected_record()
        if r is None:
            QMessageBox.information(self, "Preview", "Select a record in the table.")
            return
        self.revalidate()
        if r.status == Status.BLOCKED:
            QMessageBox.warning(self, "Preview", "This record is blocked:\n\n" +
                                "\n".join(f"{i.rule} {i.message}" for i in r.blocking_issues()))
            return
        out = Path(tempfile.mkdtemp(prefix="preview_", dir=temp_dir()))
        if self.pdf_available:
            opts = ExportOptions(docx=False, pdf=True, combined=False, batch_subfolder=False)
        else:
            opts = ExportOptions(docx=True, pdf=False, combined=False, batch_subfolder=False)
        self._preview_mode = True
        was = r.include
        r.include = True
        self._start([r], out, opts, restore=(r, was))

    def on_generate(self):
        if not self._check_ready():
            return
        if not self.out_dir:
            self.on_choose_output()
            if not self.out_dir:
                return
        self.revalidate()
        included = [r for r in self.model.records if r.include]
        if not included:
            QMessageBox.information(self, "Generate", "No records are selected.")
            return
        blocked = [r for r in included if r.status == Status.BLOCKED]
        if blocked and QMessageBox.question(
                self, "Blocked records",
                f"{len(blocked)} selected record(s) are blocked and will be skipped "
                f"(e.g. record {blocked[0].serial}: {blocked[0].blocking_issues()[0].message}).\n\nContinue?"
        ) != QMessageBox.Yes:
            return
        if not self._acknowledge_warnings([r for r in included if r.status == Status.WARNING]):
            return
        if not (self.chk_docx.isChecked() or self.chk_pdf.isChecked()):
            QMessageBox.information(self, "Generate", "Choose at least one output: DOCX or PDF.")
            return
        self._preview_mode = False
        opts = ExportOptions(docx=self.chk_docx.isChecked(), pdf=self.chk_pdf.isChecked(),
                             combined=self.chk_comb.isChecked(), duplex=self.chk_duplex.isChecked())
        self._start(list(self.model.records), self.out_dir, opts)

    def _start(self, records, out_dir, opts, restore=None):
        self.btn_generate.setEnabled(False)
        self.btn_preview.setEnabled(False)
        self.table.setEnabled(False)  # records must not change while the worker reads them
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.progress_label.setText("Starting…")
        self._restore = restore
        self._thread = QThread(self)
        self._worker = GenerateWorker(records, self.model.settings, self.template, out_dir, opts)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_progress(self, done, total, msg):
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(done)
        self.progress_label.setText(msg)

    def _end_run(self):
        self.btn_generate.setEnabled(True)
        self.btn_preview.setEnabled(True)
        self.table.setEnabled(True)
        self.progress.setVisible(False)
        if self._restore:
            r, was = self._restore
            r.include = was
            self._restore = None

    def _on_failed(self, msg):
        self._end_run()
        self.progress_label.setText("")
        QMessageBox.critical(self, "Generation failed", msg)

    def _on_finished(self, res: BatchResult):
        self._end_run()
        if self._preview_mode and not self.pdf_available:
            self._preview_mode = False
            it = res.items[0] if res.items else None
            if it and it.docx:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(it.docx)))
                self.progress_label.setText(f"Preview opened: {it.docx.name}")
            else:
                QMessageBox.warning(self, "Preview", it.error if it else "Preview failed.")
            return
        self._preview_mode = False
        self.last_result = res
        self.results.clear()
        for it in res.items:
            if it.status in ("excluded",):
                continue
            label = f"#{it.record.serial:03d}  {it.status.upper():9}  "
            label += (it.pdf.name if it.pdf else it.docx.name if it.docx else it.error)
            if it.pages:
                label += f"  ({it.pages} p)"
            if any(w.rule == "V-15" for w in it.warnings):
                label += "  ⚠ V-15"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, it)
            if it.status in ("failed", "blocked"):
                item.setForeground(QBrush(QColor("#b00020")))
            self.results.addItem(item)
        if res.combined:
            item = QListWidgetItem(f"COMBINED  {res.combined.name}  ({res.combined_pages} p)")
            item.setData(Qt.UserRole, res.combined)
            self.results.addItem(item)
        self.progress_label.setText(
            f"Generated {res.count('generated')}, failed {res.count('failed')}, blocked {res.count('blocked')}. "
            f"Folder: {res.folder}")
        self.right_tabs.setCurrentIndex(1)
        for i in range(self.results.count()):
            if isinstance(self.results.item(i).data(Qt.UserRole), Path) or \
                    getattr(self.results.item(i).data(Qt.UserRole), "pdf", None):
                self.results.setCurrentRow(i)
                break

    def _on_result_selected(self, cur: QListWidgetItem | None, _prev=None):
        if cur is None:
            return
        data = cur.data(Qt.UserRole)
        if isinstance(data, Path):
            self.show_pdf(data, f"Combined PDF: {self.last_result.combined_pages} pages.")
            return
        if data.pdf:
            warn = "; ".join(w.message for w in data.warnings if w.rule in ("V-14", "V-15"))
            self.show_pdf(data.pdf, f"Record {data.record.serial}: {data.pages} page(s)"
                          f" (template expects {self.last_result.expected_pages or '?'})."
                          + (f" ⚠ {warn}" if warn else " Page check OK."))
        else:
            self.pdf_doc.close()
            self.preview_info.setText(data.error or (f"{data.docx.name} — double-click to open in Word."
                                                     if data.docx else "No output for this record."))

    def _open_result_file(self, item: QListWidgetItem):
        data = item.data(Qt.UserRole)
        path = data if isinstance(data, Path) else (data.pdf or data.docx)
        if path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def show_pdf(self, path: Path, info: str):
        self.pdf_doc.close()
        self.pdf_doc.load(str(path))
        self.preview_info.setText(info)

    def _open_output_folder(self):
        if self.last_result:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_result.folder)))

    def closeEvent(self, ev):
        if self._thread and self._thread.isRunning():
            QMessageBox.information(self, "Busy", "Please wait for generation to finish.")
            ev.ignore()
            return
        super().closeEvent(ev)

