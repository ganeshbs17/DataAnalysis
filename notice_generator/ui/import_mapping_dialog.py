"""Choose sheet, header row and which column fills each field (FR-02 steps 2–4)."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHeaderView, QLabel, QSpinBox,
                               QTableWidget, QTableWidgetItem, QVBoxLayout)

from ..importers.cell_reader import read_workbook
from ..importers.column_mapper import detect_header_row, map_columns

NOT_IN_FILE = "(not in file)"


def _col_letter(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


class ImportMappingDialog(QDialog):
    def __init__(self, parent, path: str | Path, fields: list[str], required: set[str] | None = None,
                 sheet: str | None = None, header_row: int | None = None,
                 current: dict[str, int] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Column mapping")
        self.resize(760, 560)
        self.fields = fields
        self.required = required or set()
        self.sheets = {s.name: s for s in read_workbook(path)}
        self.sheet_combo = QComboBox()
        for name, s in self.sheets.items():
            self.sheet_combo.addItem(f"{name}  ({s.non_empty_rows} rows)", name)
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 1000)
        self.table = QTableWidget(len(fields), 3)
        self.table.setHorizontalHeaderLabels(["Field", "Column in the file", "First value"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.combos: list[QComboBox] = []
        for r, f in enumerate(fields):
            self.table.setItem(r, 0, QTableWidgetItem(f + (" *" if f in self.required else "")))
            cb = QComboBox()
            cb.currentIndexChanged.connect(lambda _i, row=r: self._sample(row))
            self.table.setCellWidget(r, 1, cb)
            self.combos.append(cb)
        form = QFormLayout()
        form.addRow("Worksheet:", self.sheet_combo)
        form.addRow("Header row:", self.header_spin)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Choose which column of the file fills each field. * = required."))
        v.addLayout(form)
        v.addWidget(self.table, 1)
        v.addWidget(bb)

        if sheet and sheet in self.sheets:
            self.sheet_combo.setCurrentIndex(list(self.sheets).index(sheet))
        self._current = current
        self._preset_header = header_row
        self.sheet_combo.currentIndexChanged.connect(self._sheet_changed)
        self.header_spin.valueChanged.connect(lambda _v: self._load_headers())
        self._sheet_changed()

    def _sheet(self):
        return self.sheets[self.sheet_combo.currentData()]

    def _sheet_changed(self):
        hr = self._preset_header or detect_header_row(self._sheet().rows, fields=self.fields) or 1
        self._preset_header = None
        self.header_spin.blockSignals(True)
        self.header_spin.setValue(hr)
        self.header_spin.blockSignals(False)
        self._load_headers()

    def _load_headers(self):
        rows = self._sheet().rows
        hr = self.header_spin.value()
        headers = [c.text.strip() for c in rows[hr - 1]] if hr <= len(rows) else []
        auto = map_columns(headers, hr, fields=self.fields).field_to_col
        if self._current:
            auto = {**auto, **self._current}
            self._current = None
        for r, f in enumerate(self.fields):
            cb = self.combos[r]
            cb.blockSignals(True)
            cb.clear()
            cb.addItem(NOT_IN_FILE, -1)
            for i, h in enumerate(headers):
                cb.addItem(f"{_col_letter(i)}: {h or '(no heading)'}", i)
            cb.setCurrentIndex(cb.findData(auto.get(f, -1)))
            cb.blockSignals(False)
            self._sample(r)

    def _sample(self, row: int):
        rows = self._sheet().rows
        hr = self.header_spin.value()
        col = self.combos[row].currentData()
        val = ""
        if col is not None and col >= 0:
            for data in rows[hr:]:
                if col < len(data) and not data[col].is_empty:
                    val = data[col].text.replace("\n", " ⏎ ")[:60]
                    break
        self.table.setItem(row, 2, QTableWidgetItem(val))

    def result_mapping(self) -> tuple[str, int, dict[str, int | None]]:
        overrides = {f: (cb.currentData() if cb.currentData() >= 0 else None) for f, cb in zip(self.fields, self.combos)}
        return self.sheet_combo.currentData(), self.header_spin.value(), overrides
