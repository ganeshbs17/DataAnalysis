"""Local JSON configuration: application folder and officer profiles (FR-01)."""
from __future__ import annotations

import json
import os
import shutil
from dataclasses import asdict, dataclass
from pathlib import Path

RESOURCES = Path(__file__).resolve().parent.parent / "resources"


def app_dir() -> Path:
    base = os.environ.get("NOTICE_GEN_HOME")
    if base:
        return Path(base)
    if os.name == "nt" and os.environ.get("APPDATA"):
        return Path(os.environ["APPDATA"]) / "NoticeGenerator"
    return Path.home() / ".notice_generator"


def config_dir() -> Path:
    d = app_dir() / "Configuration"
    d.mkdir(parents=True, exist_ok=True)
    return d


def temp_dir() -> Path:
    d = app_dir() / "Temp"
    d.mkdir(parents=True, exist_ok=True)
    return d


@dataclass
class OfficerProfile:
    name: str
    designation: str = ""
    unit: str = ""
    station: str = ""
    contact: str = ""


def _profiles_file() -> Path:
    f = config_dir() / "officer_profiles.json"
    if not f.exists():
        shutil.copy(RESOURCES / "officer_profiles.json", f)
    return f


def load_profiles() -> list[OfficerProfile]:
    try:
        data = json.loads(_profiles_file().read_text(encoding="utf-8"))
        return [OfficerProfile(**p) for p in data.get("profiles", [])]
    except (OSError, ValueError, TypeError):
        return []


def save_profiles(profiles: list[OfficerProfile]) -> None:
    f = _profiles_file()
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps({"profiles": [asdict(p) for p in profiles]}, indent=2, ensure_ascii=False),
                   encoding="utf-8")
    os.replace(tmp, f)


def _settings_file() -> Path:
    return config_dir() / "application_settings.json"


def load_app_settings() -> dict:
    try:
        return json.loads(_settings_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_app_setting(key: str, value) -> None:
    data = load_app_settings()
    data[key] = value
    f = _settings_file()
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, f)
