# core.py
"""
TSSR Automator - core pipeline (triple source).

Excel masterlist -> site identity (authoritative)
AI extraction    -> supplementary fields (fills gaps)
Ericsson TSSR    -> images + fallback fields
User uploads     -> optional image overrides

Work permit + Access requirement are HARDCODED from towercо rules
(see permit_rules.py) - not from AI.

Images are auto-fitted to their target cell aspect ratio to
fill the box edge-to-edge.
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

from permit_rules import get_permits_for_towerco


# =============================================================
# 1. PDF to DOCX
# =============================================================

def pdf_to_docx(pdf_path: str, out_dir: str) -> str:
    """Convert PDF to DOCX using LibreOffice, with pdf2docx as fallback."""
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


# =============================================================
# 2. Extract images from DOCX
# =============================================================

def extract_images(docx_path: str, out_dir: str) -> list[str]:
    """Extract all images from word/media/ inside a DOCX."""
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


# =============================================================
# 3. Grid composer for multi-image slots
# =============================================================

def make_grid(images: list[str], max_width: int = 1200,
              cols: int = 2, padding: int = 10) -> str:
    """Combine multiple images into a single grid image."""
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


# =============================================================
# 3b. Image aspect-ratio fitter (crop to match target shape)
# =============================================================

def fit_image_to_aspect(
    src_path: str,
    dest_path: str,
    target_aspect: float,
) -> str:
    """
    Crop an image to match the target aspect ratio (width/height).
    Centers the crop. Returns the path to the fitted image.

    If the image is already close to the target aspect, saves a copy as-is.
    """
    try:
        img = Image.open(src_path).convert("RGB")
    except Exception as e:
        print(f"[fit_image_to_aspect] Could not open {src_path}: {e}")
        return src_path

    w, h = img.size
    if h == 0 or w == 0:
        return src_path

    current_aspect = w / h

    # Already close enough — just copy
    if abs(current_aspect - target_aspect) < 0.02:
        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        img.save(dest, quality=88)
        return str(dest)

    if current_aspect > target_aspect:
        # Image is wider than target — crop the left/right sides
        new_w = int(h * target_aspect)
        x_offset = (w - new_w) // 2
        img = img.crop((x_offset, 0, x_offset + new_w, h))
    else:
        # Image is taller than target — crop top/bottom
        new_h = int(w / target_aspect)
        y_offset = (h - new_h) // 2
        img = img.crop((0, y_offset, w, y_offset + new_h))

    dest = Path(dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, quality=88)
    return str(dest)


# =============================================================
# 4. Checkbox helper
# =============================================================

def cb(flag: bool) -> str:
    """Return a checked or unchecked box character."""
    return "[x]" if flag else "[ ]"


# =============================================================
# 5. Merge AI-extracted fields into site data
# =============================================================

def merge_ai_fields(site: dict, ai_fields: dict) -> dict:
    """
    Merge AI-extracted fields into the site dict.
    Priority: Excel > AI > existing site dict > empty.
    """
    SKIP_FIELDS = {"work_permit", "access_requirement", "site_key_location"}

    merged = dict(site)

    for key, val in ai_fields.items():
        if key in SKIP_FIELDS:
            continue

        existing = merged.get(key)
        if existing not in ("", None, [], {}):
            continue

        if isinstance(val, str):
            clean_val = val.strip()
            if clean_val:
                merged[key] = clean_val
        elif isinstance(val, list):
            if val:
                merged[key] = val
        elif isinstance(val, bool):
            merged[key] = val

    return merged


# =============================================================
# 6. Build the docxtpl context
# =============================================================

def build_context(
    site: dict,
    ericsson_fields: dict,
    images: dict,
) -> dict:
    """Merge site data (Excel + AI), Ericsson TSSR values, image paths."""
    ctx = {}

    def pick(*sources, default=""):
        for s in sources:
            if s not in ("", None, [], {}):
                return s
        return default

    # --- Identity ---
    ctx["site_id"]      = pick(site.get("site_id"),      ericsson_fields.get("site_id"))
    ctx["site_name"]    = pick(site.get("site_name"),    ericsson_fields.get("site_name"))
    ctx["region"]       = pick(site.get("region"),       ericsson_fields.get("region"), default="MINDANAO")
    ctx["site_address"] = pick(site.get("site_add"),     site.get("site_address"), compose_address_from(site))
    ctx["site_coords"]  = pick(site.get("site_coords"),  compose_coords_from(site))
    ctx["towerco"]      = pick(site.get("towerco"),      ericsson_fields.get("towerco"))

    # --- Contact ---
    ctx["site_contact_person"] = pick(site.get("site_contact_person"), site.get("fo_name"))
    ctx["site_contact_mobile"] = pick(site.get("site_contact_mobile"), site.get("fo_mobile"))
    ctx["fo_name"]             = pick(site.get("fo_name"))
    ctx["fo_mobile"]           = pick(site.get("fo_mobile"))

    # --- Lessor ---
    ctx["lessor_details"] = pick(site.get("lessor_details"), ericsson_fields.get("lessor_details"), default="N/A")
    ctx["lessor_mobile"]  = pick(site.get("lessor_mobile"),  ericsson_fields.get("lessor_mobile"),  default="N/A")

    # --- TCO / Site class ---
    ctx["tco_name"]   = pick(site.get("towerco"), ericsson_fields.get("tco_name"))
    ctx["site_class"] = pick(site.get("site_class"), ericsson_fields.get("site_class"))

    # --- Site type checkboxes ---
    site_type = (
        site.get("site_type")
        or ericsson_fields.get("site_type")
        or ""
    ).strip().lower()

    room_access = pick(
        site.get("room_access"),
        ericsson_fields.get("room_access"),
    ).lower()

    if not site_type:
        if "outdoor" in room_access or not room_access:
            site_type = "greenfield"
        elif "indoor" in room_access:
            site_type = "indoor"
        elif "street" in room_access:
            site_type = "street cabinet"

    ctx["site_type_greenfield"] = cb("greenfield" in site_type or "outdoor" in site_type)
    ctx["site_type_street"]     = cb("street" in site_type)
    ctx["site_type_indoor"]     = cb("indoor" in site_type)
    ctx["site_type_others"]     = cb(
        site_type not in ("greenfield", "outdoor", "street cabinet", "indoor")
        and site_type != ""
    )

    # --- Site owner checkboxes ---
    owner = pick(
        site.get("site_owner"),
        ericsson_fields.get("owner"),
    ).upper()

    ctx["owner_globe"]   = cb("GLOBE" in owner)
    ctx["owner_private"] = cb("PRIVATE" in owner)
    ctx["owner_govt"]    = cb("GOVERNMENT" in owner or "GOVT" in owner)
    ctx["owner_tco"]     = cb("TCO" in owner or "TOWER" in owner)

    # --- Room access + cabin location ---
    ctx["room_access"]    = pick(site.get("room_access"),    ericsson_fields.get("room_access"))
    ctx["cabin_location"] = pick(site.get("cabin_location"), ericsson_fields.get("cabin_location"), default="Ground Level")

    # --- Flood history + hauling ---
    ctx["flood_history"]   = pick(site.get("flood_history"),   ericsson_fields.get("flood_history"),   default="None")
    ctx["hauling_remarks"] = pick(site.get("hauling_remarks"), ericsson_fields.get("hauling_remarks"), default="N/A")

    # --- Site profile ---
    ctx["site_profile"] = pick(site.get("site_profile"), ericsson_fields.get("site_profile"), default="GT Wireless")

    # --- Site key location ---
    ctx["site_key_location"] = pick(
        site.get("site_key_location"),
        ericsson_fields.get("site_key_location"),
    )

    # --- Site security checkboxes ---
    sec_raw = site.get("site_security") or ericsson_fields.get("site_security", "")
    if isinstance(sec_raw, list):
        sec = " ".join(str(s) for s in sec_raw).upper()
    else:
        sec = str(sec_raw).upper()

    ctx["security_elock"] = cb("E-LOCK" in sec or "ELOCK" in sec)
    ctx["sec_caretaker"]  = cb("CARETAKER" in sec)
    ctx["sec_manned"]     = cb("MANNED" in sec)
    ctx["sec_roving"]     = cb("ROVING" in sec)
    ctx["sec_others"]     = cb("OTHERS" in sec)

    # --- Work permit + Access requirement (HARDCODED FROM TOWERCO) ---
    towercо = pick(site.get("towerco"), ericsson_fields.get("towerco"))
    permits = get_permits_for_towerco(towercо)

    ctx["work_permit"]        = permits["work_permit"]
    ctx["access_requirement"] = permits["access_requirement"]

    # Legacy checkbox tags
    permit_upper = ctx["work_permit"].upper()
    ctx["workpermit_raawa"]  = cb("RAAWA" in permit_upper)
    ctx["workpermit_others"] = cb("OTHERS" in permit_upper)

    # --- Vehicle accessibility ---
    ctx["no_bridge"]   = pick(site.get("no_bridge"),   ericsson_fields.get("no_bridge"),   default="N/A")
    ctx["foot_trail"]  = pick(site.get("foot_trail"),  ericsson_fields.get("foot_trail"),  default="N/A")
    ctx["bridge_ton"]  = pick(site.get("bridge_ton"),  ericsson_fields.get("bridge_ton"),  default="N/A")
    ctx["foot_bridge"] = pick(site.get("foot_bridge"), ericsson_fields.get("foot_bridge"), default="N/A")
    ctx["distance_m"]  = pick(site.get("distance_m"),  ericsson_fields.get("distance_m"),  default="N/A")
    ctx["by_boat"]     = pick(site.get("by_boat"),     ericsson_fields.get("by_boat"),     default="N/A")

    # --- Site remarks ---
    ctx["site_remarks"] = build_site_remarks(site, ericsson_fields)

    # --- Images ---
    IMAGE_SLOTS = [
        "img_vicinity_map",
        "img_site_photo_1", "img_site_photo_2",
        "img_site_photo_3", "img_site_photo_4",
        "img_olt_existing", "img_olt_proposed",
        "RS1_load_sched_img", "RS1_load_calc_img",
        "RS2_load_sched_img", "RS2_load_calc_img",
        "RS1_tapping_img",    "RS2_tapping_img",
        "equipment_room_layout_img", "equipment_cable_routing_img",
        "transport_existing_1", "transport_existing_2", "transport_existing_3",
    ]
    for slot in IMAGE_SLOTS:
        ctx[slot] = images.get(slot, "")

    # --- Materials ---
    MATERIAL_KEYS = [
        "grounding", "patchcord", "powercable", "lugs", "tube",
        "tiewrap", "termlog", "conduit", "dcbreaker",
    ]
    for key in MATERIAL_KEYS:
        ctx[f"mat_{key}_unit"] = images.get(f"mat_{key}_unit", "")
        ctx[f"mat_{key}_qty"]  = images.get(f"mat_{key}_qty", "")

    return ctx


# =============================================================
# 7. Helpers
# =============================================================

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
    except (ValueError, TypeError):
        return f"{lat}, {lon}"


def build_site_remarks(site: dict, ericsson: dict) -> str:
    """Compose the multi-line Remarks block."""
    tco       = site.get("towerco", "") or ericsson.get("towerco", "")
    fo_name   = site.get("fo_name", "")
    fo_mobile = site.get("fo_mobile", "")
    flood     = site.get("flood_history") or ericsson.get("flood_history", "None")
    cabin_loc = site.get("cabin_location") or ericsson.get("cabin_location", "Ground Level")
    accessible = site.get("site_accessible", ericsson.get("site_accessible", True))

    if isinstance(accessible, bool):
        accessible_str = "accessible" if accessible else "not accessible"
    else:
        accessible_str = "accessible" if "accessible" in str(accessible).lower() else "not accessible"

    return (
        f"Pre-requisite / Access: RAAWA, HSWP, and approved {tco} "
        f"ticket are required before entry. "
        f"(Site is {accessible_str} to vehicles).\n\n"
        f"Health & Safety / Site Condition: Site is an outdoor Greenfield "
        f"setup with {flood.lower()} flood history. {cabin_loc} cabin location.\n\n"
        f"Contact Person: {fo_name} (Mobile: {fo_mobile})."
    )


# =============================================================
# 8. Generate the final DOCX
# =============================================================

# Target image widths (mm) — tuned to fill their template cells
IMAGE_WIDTHS_MM = {
    # Full width (single image rows)
    "img_vicinity_map":            Mm(170),
    "equipment_room_layout_img":   Mm(170),
    "equipment_cable_routing_img": Mm(170),

    # RS load schedules — REDUCED from 160 to 130 (they were too big)
    "RS1_load_sched_img":          Mm(130),
    "RS1_load_calc_img":           Mm(130),
    "RS2_load_sched_img":          Mm(130),
    "RS2_load_calc_img":           Mm(130),

    # Half width (2-across table cells)
    "img_site_photo_1":            Mm(85),
    "img_site_photo_2":            Mm(85),
    "img_site_photo_3":            Mm(85),
    "img_site_photo_4":            Mm(85),
    "img_olt_existing":            Mm(85),
    "img_olt_proposed":            Mm(85),
    "RS1_tapping_img":             Mm(85),
    "RS2_tapping_img":             Mm(85),

    # Third width (3-across table cells)
    "transport_existing_1":        Mm(56),
    "transport_existing_2":        Mm(56),
    "transport_existing_3":        Mm(56),
}


# Target aspect ratios (width / height) per slot — for crop-to-fit
TARGET_ASPECTS = {
    # Full width
    "img_vicinity_map":            1.70,
    "RS1_load_sched_img":          1.70,
    "RS1_load_calc_img":           1.70,
    "RS2_load_sched_img":          1.70,
    "RS2_load_calc_img":           1.70,
    "equipment_room_layout_img":   1.10,
    "equipment_cable_routing_img": 1.10,

    # Half width
    "img_site_photo_1":            1.40,
    "img_site_photo_2":            1.40,
    "img_site_photo_3":            1.40,
    "img_site_photo_4":            1.40,
    "img_olt_existing":            1.40,
    "img_olt_proposed":            1.40,
    "RS1_tapping_img":             1.40,
    "RS2_tapping_img":             1.40,

    # Third width
    "transport_existing_1":        1.40,
    "transport_existing_2":        1.40,
    "transport_existing_3":        1.40,
}


def generate_nokia_tssr(
    template_path: str,
    output_path: str,
    context: dict,
    image_widths: dict | None = None,
) -> str:
    """
    Render the Nokia template with the given context.

    Images are first cropped to match their target aspect ratio
    (so they fill their cell edge-to-edge), then inserted at the
    tuned width for that slot.
    """
    doc = DocxTemplate(template_path)

    # Merge caller overrides on top of defaults
    widths = {**IMAGE_WIDTHS_MM, **(image_widths or {})}

    # Directory for cropped-fitted images
    fitted_dir = Path(output_path).parent / "_fitted"

    rendered = {}
    for key, val in context.items():
        # Only process string paths that exist and have a target width
        if not (isinstance(val, str) and val and Path(val).exists() and key in widths):
            rendered[key] = val
            continue

        try:
            target_aspect = TARGET_ASPECTS.get(key)

            if target_aspect:
                # Crop the image to match the cell's aspect ratio
                fitted_dir.mkdir(parents=True, exist_ok=True)
                fitted_path = fitted_dir / f"{key}_fitted.jpg"
                fit_image_to_aspect(val, str(fitted_path), target_aspect)
                rendered[key] = InlineImage(doc, str(fitted_path), width=widths[key])
            else:
                # No aspect target — insert as-is
                rendered[key] = InlineImage(doc, val, width=widths[key])
        except Exception as e:
            print(f"[generate_nokia_tssr] Image render failed for {key}: {e}")
            rendered[key] = val

    doc.render(rendered)
    doc.save(output_path)
    return output_path


# =============================================================
# 9. Output filename helper
# =============================================================

def make_output_name(site_id: str, site_name: str, ext: str = "docx") -> str:
    site_id = (site_id or "SITE").replace(" ", "")
    site_name = (site_name or "TSSR").replace(" ", "")
    return f"{site_id}_{site_name}_TSSR.{ext}"
