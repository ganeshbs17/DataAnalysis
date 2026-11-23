"""Normalisation on a working copy (FR-05). The source is never changed."""
from __future__ import annotations

import re

# Prefix followed by a dot and/or spaces ("Mr. X", "Mr X", "Mr.X"); "Mrinal" does not match.
PREFIX_RE = re.compile(r"^\s*(mr|mrs|ms|miss|shri|sri|smt|kum|kumari|dr)(?:\.\s*|\s+)", re.IGNORECASE)
_MALE = {"mr", "shri", "sri"}
_ADDRESS_LABEL_RE = re.compile(r"^\s*address\s*:\s*", re.IGNORECASE)


def normalise_text(s: str) -> str:
    """Trim, collapse spaces within lines, drop leading/trailing blank lines, keep inner breaks."""
    if not s:
        return ""
    lines = [re.sub(r"[ \t ]+", " ", ln).strip() for ln in s.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def normalise_address(s: str, remove_label: bool = False) -> str:
    t = normalise_text(s)
    if remove_label:
        t = "\n".join(_ADDRESS_LABEL_RE.sub("", ln) for ln in t.split("\n"))
        t = normalise_text(t)
    return t


def name_key(name: str) -> str:
    """Name used for grouping: prefix removed, case-insensitive, spaces collapsed."""
    t = normalise_text(name).replace("\n", " ")
    t = PREFIX_RE.sub("", t)
    return re.sub(r"\s+", " ", t).strip().lower()


def apply_prefix_mode(name: str, mode: str) -> str:
    t = normalise_text(name)
    if mode == "keep":
        return t
    m = PREFIX_RE.match(t)
    bare = PREFIX_RE.sub("", t) if m else t
    if mode == "remove":
        return bare
    if mode == "standardise":
        if not m:
            return t
        return ("Mr. " if m.group(1).lower() in _MALE else "Ms. ") + bare
    return t


def normalise_account(s: str) -> str:
    """Remove spaces only. Other non-digits are reported (V-02), not silently dropped."""
    return re.sub(r"\s+", "", s or "")
