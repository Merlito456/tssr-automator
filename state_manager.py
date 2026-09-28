# state_manager.py
"""
Local-only, per-user persistence for the TSSR Automator.

This module NEVER writes to shared or remote locations. State lives
in a `.session/` folder next to this file, and only if the working
directory looks like a local checkout.

If the app is running on Streamlit Cloud (or any remote host), every
save/load call is a no-op and returns empty results. This is
intentional: your site data must never persist on a shared host.
"""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any


# ─────────────────────────────────────────────────────────────
# Where state lives — always next to this file
# ─────────────────────────────────────────────────────────────

APP_DIR      = Path(__file__).resolve().parent
SESSION_DIR  = APP_DIR / ".session"
SESSION_FILE = SESSION_DIR / "state.json"
WORKDIR_FILE = SESSION_DIR / "workdir.txt"


# ─────────────────────────────────────────────────────────────
# Detect remote / shared environments and refuse to persist
# ─────────────────────────────────────────────────────────────

# Paths that indicate we're on a remote host or a shared container.
# If APP_DIR contains any of these, we disable persistence.
_REMOTE_MARKERS = (
    "/mount/src/",        # Streamlit Cloud
    "/home/adminuser/",   # Streamlit Cloud
    "/home/appuser/",     # Hugging Face Spaces / similar
    "/workspace/",        # some CI/cloud IDEs
    "/app/",              # some Docker images
)


def _is_remote_environment() -> bool:
    """True if the app is running somewhere that isn't this machine."""
    # Streamlit Cloud sets this env var
    if os.environ.get("STREAMLIT_SHARING_MODE") or \
       os.environ.get("IS_STREAMLIT_CLOUD"):
        return True
    # Match against well-known shared-host path prefixes
    app_str = str(APP_DIR)
    for marker in _REMOTE_MARKERS:
        if app_str.startswith(marker):
            return True
    return False


_REMOTE = _is_remote_environment()


# ─────────────────────────────────────────────────────────────
# Keys we persist
# ─────────────────────────────────────────────────────────────

PERSISTED_KEYS = {
    "selected_plaid",
    "site_data",
    "image_map",
    "materials",
    "ericsson_docx",
    "ericsson_images",
    "ericsson_fields",
}


# ─────────────────────────────────────────────────────────────
# Internals
# ─────────────────────────────────────────────────────────────

def _ensure_session_dir() -> bool:
    """Create `.session/` if allowed. Returns False if we refuse."""
    if _REMOTE:
        return False
    try:
        SESSION_DIR.mkdir(parents=True, exist_ok=True)
        _ensure_gitignore()
        return True
    except Exception as e:
        print(f"⚠️ Could not create {SESSION_DIR}: {e}")
        return False


def _ensure_gitignore() -> None:
    """Make sure `.session/` is never accidentally committed."""
    gi = SESSION_DIR / ".gitignore"
    if not gi.exists():
        try:
            gi.write_text("# Local-only session state\n*\n")
        except Exception:
            pass


def _atomic_write(path: Path, text: str) -> None:
    """Write via a temp file + rename so a crash can't corrupt state."""
    tmp_fd, tmp_path = tempfile.mkstemp(
        prefix=path.name + ".", dir=str(path.parent)
    )
    try:
        with os.fdopen(tmp_fd, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


# ─────────────────────────────────────────────────────────────
# Public API (same signatures as the Streamlit version)
# ─────────────────────────────────────────────────────────────

def save_state(state: dict) -> None:
    """Persist selected keys to `.session/state.json` (local only)."""
    if not _ensure_session_dir():
        return
    data = {k: state.get(k) for k in PERSISTED_KEYS}
    try:
        _atomic_write(
            SESSION_FILE,
            json.dumps(data, indent=2, default=str),
        )
    except Exception as e:
        print(f"⚠️ State save failed: {e}")


def load_state() -> dict:
    """Load the previous state, or return {} if none exists."""
    if _REMOTE or not SESSION_FILE.exists():
        return {}
    try:
        return json.loads(SESSION_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"⚠️ State load failed: {e}")
        return {}


def save_workdir(path: str) -> None:
    if not _ensure_session_dir():
        return
    try:
        _atomic_write(WORKDIR_FILE, str(path))
    except Exception as e:
        print(f"⚠️ Workdir save failed: {e}")


def load_workdir() -> str | None:
    if _REMOTE or not WORKDIR_FILE.exists():
        return None
    try:
        return WORKDIR_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        return None


def clear_state() -> None:
    """Remove all persisted session data (local only)."""
    if _REMOTE:
        return
    for f in (SESSION_FILE, WORKDIR_FILE):
        try:
            if f.exists():
                f.unlink()
        except OSError as e:
            print(f"⚠️ Could not remove {f}: {e}")
