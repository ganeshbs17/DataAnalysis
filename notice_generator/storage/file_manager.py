"""Safe file naming; existing files are never overwritten (FR-09, NFR-08)."""
from __future__ import annotations

import re
from pathlib import Path

_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def safe_name(s: str, max_len: int = 60) -> str:
    t = _INVALID.sub("", s or "")
    t = re.sub(r"\s+", "_", t.strip()).strip("._")
    t = t[:max_len].rstrip("._")
    if not t or t.upper() in _RESERVED:
        t = f"_{t}" if t else "notice"
    return t


def notice_basename(serial: int, name: str, account: str) -> str:
    digits = re.sub(r"\D", "", account or "")
    last4 = digits[-4:] if digits else "XXXX"
    return f"{serial:03d}_{safe_name(name)}_{last4}"


def unique_path(folder: Path, stem: str, ext: str) -> Path:
    """``stem.ext``, or ``stem (2).ext``, ``stem (3).ext``… if the name is taken."""
    p = folder / f"{stem}{ext}"
    n = 2
    while p.exists():
        p = folder / f"{stem} ({n}){ext}"
        n += 1
    return p
