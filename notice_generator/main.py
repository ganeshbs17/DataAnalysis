"""GUI entry point.

``--selftest`` runs a headless check used by the Windows build: it imports
the engine and the Qt window, renders a small DOCX from a generated
template, writes ``selftest.log`` next to the executable and exits 0 on
success.
"""
from __future__ import annotations

import sys
import tempfile
import traceback
from pathlib import Path


def selftest() -> int:
    log_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
    log = log_dir / "selftest.log"
    try:
        import datetime as dt

        from docx import Document

        from .core.models import CellType, NoticeSettings, RawCell, Record, SourceRef
        from .core.notice_builder import build_context
        from .core.validation import load_bank_names
        from .templates.renderer import render_docx

        assert load_bank_names(), "bank list missing from bundle"
        with tempfile.TemporaryDirectory() as td:
            tpl = Path(td) / "t.docx"
            d = Document()
            d.add_paragraph("₹{{AmountDeposited}} {{SuspectName}} {{AppearanceDate}}")
            d.add_paragraph("{{Address}}")
            d.save(tpl)
            src = {"SuspectName": RawCell("A & B", CellType.TEXT),
                   "Address": RawCell("L1\nL2", CellType.TEXT),
                   "AmountDeposited": RawCell("2000000", CellType.INT, digits=7)}
            rec = Record(1, SourceRef("selftest", "s", 2), src)
            s = NoticeSettings(appearance_date=dt.date(2030, 1, 2))
            out = Path(td) / "o.docx"
            assert render_docx(tpl, build_context(rec, s), out) == []
            text = "\n".join(p.text for p in Document(str(out)).paragraphs)
            assert "₹20,00,000/- A & B 02/01/2030" in text, text

        import os

        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication

        from .ui.main_window import MainWindow

        app = QApplication.instance() or QApplication([])
        MainWindow().close()
        app.processEvents()
        log.write_text("OK\n", encoding="utf-8")
        return 0
    except Exception:  # noqa: BLE001
        log.write_text("FAILED\n" + traceback.format_exc(), encoding="utf-8")
        return 1


def main() -> int:
    if "--selftest" in sys.argv:
        return selftest()
    from PySide6.QtWidgets import QApplication

    from .ui.main_window import MainWindow

    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Notice Generator")
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
