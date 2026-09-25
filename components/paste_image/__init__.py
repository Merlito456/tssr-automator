# components/paste_image/__init__.py
"""
Custom Streamlit component for pasting images from clipboard.
"""
from __future__ import annotations

import base64
from pathlib import Path

import streamlit.components.v1 as components


# Path to the frontend HTML
_FRONTEND_DIR = Path(__file__).parent


_component_func = components.declare_component(
    "paste_image",
    path=str(_FRONTEND_DIR),
)


def paste_image(key: str = "paste_image") -> dict | None:
    """
    Render the paste-image component.

    Returns:
        None if no image pasted yet.
        dict with keys: base64, mime, size — if user clicked "Use this image".
    """
    return _component_func(key=key, default=None)


def save_pasted_image(result: dict, dest_path: str) -> str:
    """Decode base64 image and save to disk."""
    raw = base64.b64decode(result["base64"])
    dest = Path(dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(raw)
    return str(dest)
