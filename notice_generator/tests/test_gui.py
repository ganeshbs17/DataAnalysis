"""Offscreen GUI smoke test: import → confirm → generate (worker thread) → preview."""
import os
import time

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
qtw = pytest.importorskip("PySide6.QtWidgets")

from notice_generator.exporters.pdf_engine import find_soffice  # noqa: E402


@pytest.mark.skipif(find_soffice() is None, reason="LibreOffice not installed")
def test_gui_end_to_end(sample_xlsx, template, tmp_path, monkeypatch):
    monkeypatch.setenv("NOTICE_GEN_HOME", str(tmp_path / "home"))
    from PySide6.QtCore import QDate
    from PySide6.QtWidgets import QApplication, QMessageBox

    from notice_generator.ui.main_window import MainWindow

    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    monkeypatch.setattr(QMessageBox, "information", staticmethod(lambda *a, **k: None))
    w = MainWindow()
    w.import_file(sample_xlsx)
    w.set_template(template)
    w.officer_combo.setCurrentText("Officer A")
    w.app_date.setDate(QDate(2030, 12, 2))
    assert w.model.rowCount() == 6
    w.confirm_truncation(all_records=True)
    w.set_output(tmp_path / "out")
    w.on_generate()
    t = time.time()
    while w._thread.isRunning() and time.time() - t < 120:
        app.processEvents()
        time.sleep(0.05)
    for _ in range(10):
        app.processEvents()
    res = w.last_result
    assert res is not None and res.count("generated") == 5 and res.count("blocked") == 1
    assert w.pdf_doc.pageCount() >= 1
    w.close()
