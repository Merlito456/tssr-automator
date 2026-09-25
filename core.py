# core.py
"""
TSSR Automator — core pipeline.
Pure functions, no Streamlit imports. Testable independently.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from pdf2docx import Converter


# ─────────────────────────────────────────────────────────────
# 1. PDF → DOCX
# ─────────────────────────────────────────────────────────────

def pdf_to_docx(pdf_path: str, out_dir: str) -> str:
    """Try LibreOffice first (best tables), fall back to pdf2docx."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = Path(pdf_path)
    target = out_dir / f"{pdf_path.stem}.docx"

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless",
                 "--convert-to", "docx:MS Word 2007 XML",
                 "--outdir", str(out_dir), str(pdf_path)],
                check=True, timeout=180,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            )
            if target.exists():
                return str(target)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            pass

    cv = Converter(str(pdf_path))
    try:
        cv.convert(str(target))
    finally:
        cv.close()
    return str(target)


# ─────────────────────────────────────────────────────────────
# 2. Extract Fields
# ─────────────────────────────────────────────────────────────

# Ordered by specificity — first match wins.
FIELD_PATTERNS = {
    "site_id":      r"Site\s*ID[:\s]+([A-Z0-9\-]+)",
    "site_name":    r"Site\s*Name[:\s]+([A-Za-z0-9\s\-]+?)(?:\n|$)",
    "address":      r"Site\s*Address[:\s]+(.+?)(?:\n|$)",
    "latitude":     r"Latitude\s*\(N\)\s*([\d.]+)",
    "longitude":    r"Longitude\s*\(E\)\s*([\d.]+)",
    "contact_name": r"(?:Site\s*Contact\s*(?:Person)?|Contact\s*Person)[:\s]+([A-Za-z\s.]+)",
    "contact_no":   r"Contact\s*(?:Number|No)[:\s]+([\d/,\s]+)",
    "owner":        r"Site\s*Owner[:\s]+([A-Za-z\s]+)",
    "region":       r"Region[:\s]+(.+?)(?:\n|$)",
}


def extract_fields(docx_path: str) -> dict:
    """Return dict of field → value (None if not found)."""
    doc = Document(docx_path)
    text = "\n".join(p.text for p in doc.paragraphs)

    # Also scan table cells (some fields live in tables)
    for table in doc.tables:
        for row in table.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)

    fields = {}
    for key, pattern in FIELD_PATTERNS.items():
        m = re.search(pattern, text, re.IGNORECASE)
        fields[key] = m.group(1).strip() if m else ""
    return fields


# ─────────────────────────────────────────────────────────────
# 3. Extract Images
# ─────────────────────────────────────────────────────────────

def extract_images(docx_path: str, out_dir: str) -> list[str]:
    """DOCX is a zip — pull word/media/* out."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    images = []
    with zipfile.ZipFile(docx_path) as z:
        for name in sorted(z.namelist()):
            if name.startswith("word/media/"):
                dest = out_dir / Path(name).name
                dest.write_bytes(z.read(name))
                images.append(str(dest))
    return images


# ─────────────────────────────────────────────────────────────
# 4. Generate Nokia TSSR
# ─────────────────────────────────────────────────────────────

def generate_nokia_tssr(
    template_path: str,
    output_path: str,
    fields: dict,
    image_map: dict[str, str] | None = None,
) -> str:
    """
    Populate the Nokia template.
    image_map: {"vicinity_map": "/path/img.png", "site_photo": "...", ...}
    """
    doc = DocxTemplate(template_path)
    image_map = image_map or {}

    ctx = dict(fields)  # copy

    # Wrap image paths as InlineImage at fixed width
    for slot, path in image_map.items():
        if path and Path(path).exists():
            try:
                ctx[slot] = InlineImage(doc, path, width=Mm(150))
            except Exception:
                ctx[slot] = ""
        else:
            ctx[slot] = ""

    doc.render(ctx)
    doc.save(output_path)
    return output_path


# ─────────────────────────────────────────────────────────────
# 5. Filename helper
# ─────────────────────────────────────────────────────────────

def make_output_name(fields: dict, ext: str = "docx") -> str:
    site_id = (fields.get("site_id") or "SITE").replace(" ", "")
    name = (fields.get("site_name") or "TSSR").replace(" ", "")
    return f"{site_id}_{name}_TSSR.{ext}"
