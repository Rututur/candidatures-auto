"""User settings, stored outside the repository (never committed).

Windows: %APPDATA%\\CandidaturesAuto\\settings.json
Other:   ~/.config/candidatures-auto/settings.json
The Gmail app password goes to the system keyring (Windows Credential Manager).
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, fields
from pathlib import Path

from . import templates

KEYRING_SERVICE = "candidatures-auto"


def settings_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
        return base / "CandidaturesAuto"
    return Path.home() / ".config" / "candidatures-auto"


SETTINGS_FILE = settings_dir() / "settings.json"


@dataclass
class Settings:
    gmail_address: str = ""
    sender_name: str = ""
    phone: str = ""
    cv_path: str = ""
    workbook_path: str = ""
    last_letter_dir: str = ""
    period: str = "du 4 janvier au 12 février 2027"
    delay_seconds: int = 20
    reminder_days: int = 7
    test_mode: bool = True
    test_address: str = ""
    subject_template: str = templates.DEFAULT_SUBJECT
    body_template: str = templates.DEFAULT_BODY
    followup_subject_template: str = templates.DEFAULT_FOLLOWUP_SUBJECT
    followup_body_template: str = templates.DEFAULT_FOLLOWUP_BODY

    @classmethod
    def load(cls) -> "Settings":
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return cls()
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in data.items() if k in known})

    def save(self) -> None:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        tmp = SETTINGS_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, SETTINGS_FILE)

    @property
    def effective_test_address(self) -> str:
        return (self.test_address or self.gmail_address).strip()


def get_password(address: str) -> str:
    if not address:
        return ""
    try:
        import keyring

        return keyring.get_password(KEYRING_SERVICE, address.lower()) or ""
    except Exception:
        return ""


def set_password(address: str, password: str) -> None:
    import keyring

    keyring.set_password(KEYRING_SERVICE, address.lower(), password.replace(" ", ""))
