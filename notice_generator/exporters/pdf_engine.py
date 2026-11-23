"""DOCX → PDF with headless LibreOffice (SRS §6.1).

* Dedicated profile (``-env:UserInstallation``) so an installed LibreOffice is untouched.
* One process call converts the whole batch (startup takes seconds).
* Timeout per batch; a hung process is killed and the conversion retried once.
"""
from __future__ import annotations

import os
import shutil
import signal
import subprocess
import sys
from pathlib import Path


class PdfEngineError(Exception):
    pass


def _candidate_paths() -> list[Path]:
    out: list[Path] = []
    env = os.environ.get("NOTICE_GEN_SOFFICE")
    if env:
        out.append(Path(env))
    exe = "soffice.exe" if os.name == "nt" else "soffice"
    bases = [Path(__file__).resolve().parent.parent / "resources" / "bundled_renderer"]
    if getattr(sys, "_MEIPASS", None):  # PyInstaller bundle
        bases.insert(0, Path(sys._MEIPASS) / "resources" / "bundled_renderer")
    for b in bases:
        out += [b / "program" / exe, b / exe]
    for name in ("soffice", "libreoffice"):
        w = shutil.which(name)
        if w:
            out.append(Path(w))
    if os.name == "nt":
        for pf in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                   os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            out.append(Path(pf) / "LibreOffice" / "program" / "soffice.exe")
    else:
        out.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    return out


def find_soffice() -> Path | None:
    for p in _candidate_paths():
        if p.is_file():
            return p
    return None


def _default_profile_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(Path.home(), ".notice_generator")
    return Path(base) / "NoticeGenerator" / "lo_profile"


class LibreOfficeEngine:
    name = "LibreOffice"

    def __init__(self, soffice: str | Path | None = None, profile_dir: str | Path | None = None,
                 timeout: int = 300):
        self.soffice = Path(soffice) if soffice else find_soffice()
        self.profile_dir = Path(profile_dir) if profile_dir else _default_profile_dir()
        self.timeout = timeout

    @property
    def available(self) -> bool:
        return self.soffice is not None and self.soffice.is_file()

    def _run(self, docx_files: list[Path], outdir: Path) -> None:
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        cmd = [str(self.soffice), f"-env:UserInstallation={self.profile_dir.resolve().as_uri()}",
               "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nofirststartwizard",
               "--convert-to", "pdf", "--outdir", str(outdir)] + [str(p) for p in docx_files]
        kwargs: dict = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE}
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP | 0x08000000  # CREATE_NO_WINDOW
        else:
            kwargs["start_new_session"] = True
        proc = subprocess.Popen(cmd, **kwargs)
        try:
            proc.communicate(timeout=self.timeout)
        except subprocess.TimeoutExpired:
            self._kill(proc)
            raise

    @staticmethod
    def _kill(proc: subprocess.Popen) -> None:
        try:
            if os.name == "nt":
                subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
            else:
                os.killpg(proc.pid, signal.SIGKILL)
        except OSError:
            proc.kill()
        proc.communicate()

    def convert(self, docx_files: list[Path], outdir: Path) -> dict[Path, Path | None]:
        """Convert in one batch; retry missing outputs once. Returns docx → pdf (or None)."""
        if not self.available:
            raise PdfEngineError("LibreOffice was not found. Install it or set its location in Settings.")
        outdir.mkdir(parents=True, exist_ok=True)
        result: dict[Path, Path | None] = {}
        pending = list(docx_files)
        for _attempt in range(2):
            if not pending:
                break
            try:
                self._run(pending, outdir)
            except subprocess.TimeoutExpired:
                pass
            still = []
            for d in pending:
                pdf = outdir / (d.stem + ".pdf")
                if pdf.is_file() and pdf.stat().st_size > 0:
                    result[d] = pdf
                else:
                    still.append(d)
            pending = still
        for d in pending:
            result[d] = None
        return result


def get_engine(name: str = "auto", **kw) -> LibreOfficeEngine:
    """Engine selection (Auto / LibreOffice / Word). The Word COM engine is planned for M4;
    the PoC always uses LibreOffice."""
    return LibreOfficeEngine(**kw)
