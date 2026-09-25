# core.py
"""
TSSR Automator — core pipeline (dual source).

Excel masterlist → site identity
Ericsson TSSR    → images + supplementary fields
User uploads     → optional image overrides
"""

from __future__ import annotations
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from pdf2docx import Converter
from PIL import Image


# ─────────────────────────────────────────────────────────────
# 1. PDF → DOCX (unchanged from v0.1)
# ─────────────────────────────────────────────────────────────

def pdf_to_docx(pdf_path: str, out_dir: str) -> str:
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
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
# 2. Extract images from the Ericsson DOCX
# ─────────────────────────────────────────────────────────────

def extract_images(docx_path: str, out_dir: str) -> list[str]:
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    images = []
    with zipfile.ZipFile(docx_path) as z:
        for name in sorted(z.namelist()):
            if name.startswith("word/media/"):
                dest = out_dir / Path(name).name
                dest.write_bytes(z.read(name))
                images.append(str(dest))
    return images


# ─────────────────────────────────────────────────────────────
# 3. Grid composer for multi-image slots
# ─────────────────────────────────────────────────────────────

def make_grid(images: list[str], max_width: int = 1200,
              cols: int = 2, padding: int = 10) -> str:
    if not images:
        return ""
    if len(images) == 1:
        return images[0]

    loaded = [Image.open(p).convert("RGB") for p in images
              if Path(p).exists()]
    if not loaded:
        return ""

    cell_w = max_width // cols
    resized = []
    for img in loaded:
        ratio = cell_w / img.width
        resized.append(img.resize((cell_w, int(img.height * ratio))))

    rows = (len(resized) + cols - 1) // cols
    row_heights = [
        max(img.height for img in resized[r * cols:(r + 1) * cols])
        for r in range(rows)
    ]
    total_h = sum(row_heights) + padding * (rows + 1)
    total_w = cell_w * cols + padding * (cols + 1)

    canvas = Image.new("RGB", (total_w, total_h), "white")
    y = padding
    for r in range(rows):
        x = padding
        for img in resized[r * cols:(r + 1) * cols]:
            canvas.paste(img, (x, y))
            x += cell_w + padding
        y += row_heights[r] + padding

    out = Path(images[0]).parent / "_grid.jpg"
    canvas.save(out, quality=85)
    return str(out)


# ─────────────────────────────────────────────────────────────
# 4. Checkbox helper
# ─────────────────────────────────────────────────────────────

def cb(flag: bool) -> str:
    """Return a checked or unchecked box character."""
    return "☑" if flag else "☐"


# ─────────────────────────────────────────────────────────────
# 5. Build the docxtpl context from Excel + Ericsson
# ─────────────────────────────────────────────────────────────

def build_context(
    site: dict,
    ericsson_fields: dict,
    images: dict,
) -> dict:
    """
    Merge site data (Excel), Ericsson TSSR values, and image paths
    into a single dict that matches the template's tag contract.

    Priority: Excel wins for identity fields; Ericsson fills gaps.
    """
    ctx = {}

    # ─── Identity (Excel wins) ─────────────────────────────
    ctx["site_id"]      = site.get("site_id") or ericsson_fields.get("site_id", "")
    ctx["site_name"]    = site.get("site_name") or ericsson_fields.get("site_name", "")
    ctx["region"]       = site.get("region") or ericsson_fields.get("region", "MINDANAO")
    ctx["site_address"] = site.get("site_add") or compose_address_from(site)
    ctx["site_coords"]  = site.get("site_coords") or compose_coords_from(site)
    ctx["towercо"]      = site.get("towerco") or ericsson_fields.get("towerco", "")

    # ─── Contact (Excel wins; FO = column S, phone = column U) ───
    ctx["site_contact_person"] = site.get("site_contact_person") or ""
    ctx["site_contact_mobile"] = site.get("site_contact_mobile") or ""
    ctx["fo_name"]             = site.get("fo_name") or ""
    ctx["fo_mobile"]           = site.get("fo_mobile") or ""

    # ─── Lessor (usually N/A) ──────────────────────────────
    ctx["lessor_details"]      = ericsson_fields.get("lessor_details", "N/A")
    ctx["lessor_mobile"]       = ericsson_fields.get("lessor_mobile", "N/A")

    # ─── TCO / Site class ──────────────────────────────────
    ctx["tco_name"]            = site.get("towerco", "")
    ctx["site_class"]          = ericsson_fields.get("site_class", "")

    # ─── Site type checkboxes ──────────────────────────────
    # Default: assume Greenfield/Outdoor if Ericsson says "outdoor"
    room_access = ericsson_fields.get("room_access", "").lower()
    ctx["site_type_greenfield"] = cb("outdoor" in room_access or True)
    ctx["site_type_street"]     = cb("street" in room_access)
    ctx["site_type_indoor"]     = cb("indoor" in room_access)
    ctx["site_type_others"]     = cb(False)

    # ─── Site owner checkboxes ─────────────────────────────
    owner = ericsson_fields.get("owner", "").upper()
    ctx["owner_globe"]   = cb("GLOBE" in owner)
    ctx["owner_private"] = cb("PRIVATE" in owner)
    ctx["owner_govt"]    = cb("GOVERNMENT" in owner or "GOVT" in owner)
    ctx["owner_tco"]     = cb("TCO" in owner or "TOWER" in owner)

    # ─── Room access + cabin location ──────────────────────
    ctx["room_access"]    = ericsson_fields.get("room_access", "")
    ctx["cabin_location"] = ericsson_fields.get("cabin_location", "Ground Level")

    # ─── Flood history + hauling ───────────────────────────
    ctx["flood_history"]   = ericsson_fields.get("flood_history", "None")
    ctx["hauling_remarks"] = ericsson_fields.get("hauling_remarks", "N/A")

    # ─── Site profile ──────────────────────────────────────
    ctx["site_profile"] = ericsson_fields.get("site_profile", "GT Wireless")

    # ─── Site key location (Excel column M or Q) ───────────
    ctx["site_key_location"] = (
        site.get("site_key_location")
        or ericsson_fields.get("site_key_location", "")
    )

    # ─── Site security checkboxes ──────────────────────────
    sec = ericsson_fields.get("site_security", "").upper()
    ctx["security_elock"] = cb("E-LOCK" in sec)
    ctx["sec_caretaker"]  = cb("CARETAKER" in sec)
    ctx["sec_manned"]     = cb("MANNED" in sec)
    ctx["sec_roving"]     = cb("ROVING" in sec)
    ctx["sec_others"]     = cb("OTHERS" in sec)

    # ─── Work permit ───────────────────────────────────────
    permit = ericsson_fields.get("work_permit", "").upper()
    ctx["workpermit_raawa"]  = cb("RAAWA" in permit)
    ctx["workpermit_others"] = cb(False)

    # ─── Site access requirement ───────────────────────────
    ctx["access_requirement"] = ericsson_fields.get("access_requirement", "")

    # ─── Vehicle accessibility (usually N/A) ───────────────
    ctx["no_bridge"]    = ericsson_fields.get("no_bridge", "N/A")
    ctx["foot_trail"]   = ericsson_fields.get("foot_trail", "N/A")
    ctx["bridge_ton"]   = ericsson_fields.get("bridge_ton", "N/A")
    ctx["foot_bridge"]  = ericsson_fields.get("foot_bridge", "N/A")
    ctx["distance_m"]   = ericsson_fields.get("distance_m", "N/A")
    ctx["by_boat"]      = ericsson_fields.get("by_boat", "N/A")

    # ─── Site remarks (generated) ──────────────────────────
    ctx["site_remarks"] = build_site_remarks(site, ericsson_fields)

    # ─── Images (paths; wrapped later) ─────────────────────
    # These are dict entries — core generator wraps as InlineImage.
    for slot in [
        "img_vicinity_map",
        "img_site_photo_1", "img_site_photo_2",
        "img_site_photo_3", "img_site_photo_4",
        "img_olt_existing", "img_olt_proposed",
        "RS1_load_sched_img", "RS1_load_calc_img",
        "RS2_load_sched_img", "RS2_load_calc_img",
        "RS1_tapping_img",    "RS2_tapping_img",
        "equipment_room_layout_img", "equipment_cable_routing_img",
    ]:
        ctx[slot] = images.get(slot, "")

    # ─── Transport (multi-image grid) ──────────────────────
    ctx["transport_existing_1"] = images.get("transport_existing_1", "")
    ctx["transport_existing_2"] = images.get("transport_existing_2", "")
    ctx["transport_existing_3"] = images.get("transport_existing_3", "")

    # ─── Materials (user-filled via Streamlit) ─────────────
    MATERIAL_KEYS = [
        "grounding", "patchcord", "powercable", "lugs", "tube",
        "tiewrap", "termlog", "conduit", "dcbreaker",
    ]
    for key in MATERIAL_KEYS:
        ctx[f"mat_{key}_unit"] = images.get(f"mat_{key}_unit", "")
        ctx[f"mat_{key}_qty"]  = images.get(f"mat_{key}_qty", "")

    return ctx


# ─────────────────────────────────────────────────────────────
# 6. Helpers
# ─────────────────────────────────────────────────────────────

def compose_address_from(site: dict) -> str:
    parts = [site.get("barangay", ""), site.get("municipality", ""),
             site.get("province", "")]
    return ", ".join(p.strip() for p in parts if p and p.strip())


def compose_coords_from(site: dict) -> str:
    lat, lon = site.get("latitude", ""), site.get("longitude", "")
    if not lat or not lon:
        return ""
    try:
        return f"{float(lat):.5f}, {float(lon):.5f}"
    except ValueError:
        return f"{lat}, {lon}"


def build_site_remarks(site: dict, ericsson: dict) -> str:
    """Compose the multi-line Remarks block."""
    tco         = site.get("towerco", "")
    fo_name     = site.get("fo_name", "")
    fo_mobile   = site.get("fo_mobile", "")
    flood       = ericsson.get("flood_history", "None")
    cabin_loc   = ericsson.get("cabin_location", "Ground Level")
    accessible  = "accessible" in ericsson.get("site_accessible", "accessible").lower()

    return (
        f"Pre-requisite / Access: RAAWA, HSWP, and approved {tco} "
        f"ticket are required before entry. "
        f"(Site is {accessible} to vehicles).\n\n"
        f"Health & Safety / Site Condition: Site is an outdoor Greenfield "
        f"setup with {flood.lower()} flood history. {cabin_loc} cabin location.\n\n"
        f"Contact Person: {fo_name} (Mobile: {fo_mobile})."
    )


# ─────────────────────────────────────────────────────────────
# 7. Generate the final DOCX
# ─────────────────────────────────────────────────────────────

def generate_nokia_tssr(
    template_path: str,
    output_path: str,
    context: dict,
    image_widths: dict | None = None,
) -> str:
    """
    Render the Nokia template with the given context.
    Any context value that points to an existing file path
    will be treated as an InlineImage.
    """
    doc = DocxTemplate(template_path)

    # Default image widths per slot (mm)
    DEFAULT_WIDTHS = {
        "img_vicinity_map": Mm(160),
        "img_site_photo_1": Mm(80),
        "img_site_photo_2": Mm(80),
        "img_site_photo_3": Mm(80),
        "img_site_photo_4": Mm(80),
        "img_olt_existing": Mm(80),
        "img_olt_proposed": Mm(80),
        "RS1_load_sched_img": Mm(160),
        "RS1_load_calc_img": Mm(160),
        "RS2_load_sched_img": Mm(160),
        "RS2_load_calc_img": Mm(160),
        "RS1_tapping_img": Mm(80),
        "RS2_tapping_img": Mm(80),
        "equipment_room_layout_img": Mm(170),
        "equipment_cable_routing_img": Mm(170),
        "transport_existing_1": Mm(80),
        "transport_existing_2": Mm(80),
        "transport_existing_3": Mm(80),
    }
    widths = {**DEFAULT_WIDTHS, **(image_widths or {})}

    # Convert any file-path values to InlineImage
    rendered = {}
    for key, val in context.items():
        if isinstance(val, str) and val and Path(val).exists():
            # Heuristic: if it's an image file and we have a width, wrap it
            if key in widths:
                try:
                    rendered[key] = InlineImage(doc, val, width=widths[key])
                    continue
                except Exception:
                    pass
            rendered[key] = val
        else:
            rendered[key] = val

    doc.render(rendered)
    doc.save(output_path)
    return output_path


# ─────────────────────────────────────────────────────────────
# 8. Output filename helper
# ─────────────────────────────────────────────────────────────

def make_output_name(site_id: str, site_name: str,
                     ext: str = "docx") -> str:
    site_id = (site_id or "SITE").replace(" ", "")
    site_name = (site_name or "TSSR").replace(" ", "")
    return f"{site_id}_{site_name}_TSSR.{ext}"
