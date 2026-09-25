# state_manager.py
"""
Simple JSON-backed persistence for the TSSR app.
Survives browser refresh. Cleared explicitly by user action.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any


SESSION_DIR = Path(".session")
SESSION_FILE = SESSION_DIR / "state.json"
WORKDIR_FILE = SESSION_DIR / "workdir.txt"


# Keys we actually persist (avoid polluting the file with widget junk)
PERSISTED_KEYS = {
    "selected_plaid",
    "site_data",
    "image_map",
    "materials",
    "ericsson_docx",
    "ericsson_images",
    "ericsson_fields",
}


def ensure_session_dir() -> None:
    SESSION_DIR.mkdir(parents=True, exist_ok=True)


def save_state(state: dict) -> None:
    """Persist the app state to disk."""
    ensure_session_dir()
    data = {k: state.get(k) for k in PERSISTED_KEYS}
    try:
        SESSION_FILE.write_text(json.dumps(data, indent=2, default=str))
    except Exception as e:
        # Never crash the app because of a save failure
        print(f"⚠️ State save failed: {e}")


def load_state() -> dict:
    """Load the previous state, or return {} if none exists."""
    if not SESSION_FILE.exists():
        return {}
    try:
        return json.loads(SESSION_FILE.read_text())
    except Exception as e:
        print(f"⚠️ State load failed: {e}")
        return {}


def save_workdir(path: str) -> None:
    ensure_session_dir()
    WORKDIR_FILE.write_text(str(path))


def load_workdir() -> str | None:
    if not WORKDIR_FILE.exists():
        return None
    try:
        return WORKDIR_FILE.read_text().strip()
    except Exception:
        return None


def clear_state() -> None:
    """Remove all persisted session data."""
    if SESSION_FILE.exists():
        SESSION_FILE.unlink()
    if WORKDIR_FILE.exists():
        WORKDIR_FILE.unlink()
