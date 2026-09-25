# load_calc_render.py
"""
Fill load_calculation.xlsx (C7 block) and render it to a PNG.

Merged-safe: writes to the top-left anchor of any merged range.
Label matching is FUZZY (normalized) so minor template wording
changes don't break filling.

Renderer: Playwright + Chromium.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from load_calc_helper import compute_sufficiency, NOKIA_MF2_LOAD


# ─────────────────────────────────────────────────────────────
# Label normalization + matching
# ─────────────────────────────────────────────────────────────

def _norm(s) -> str:
    """Uppercase, strip punctuation, collapse whitespace."""
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _find_label_cells(ws) -> dict[str, "Cell"]:
    """
    Return { normalized_label : cell }
    For merged ranges, always stores the top-left anchor.
    """
    anchors = {}
    for mr in ws.merged_cells.ranges:
        anchors[(mr.min_row, mr.min_col)] = str(mr)

    found: dict[str, object] = {}
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            key = _norm(v)
            if not key or key in found:
                continue
            # If this cell is the top-left anchor of a merged range,
            # keep it (it's writable). Otherwise it's a MergedCell — skip.
            if isinstance(cell, MergedCell):
                continue
            found[key] = cell
    return found


def _lookup(label_cells: dict, wanted: str):
    """Exact normalized match, then substring fallback."""
    key = _norm(wanted)
    if key in label_cells:
        return label_cells[key]
    # Substring fallback (either direction)
    for k, cell in label_cells.items():
        if key in k or k in key:
            return cell
    return None


# ─────────────────────────────────────────────────────────────
# Merged-safe writer
# ─────────────────────────────────────────────────────────────

def _safe_set(ws, row: int, col: int, value):
    """
    Write `value` to (row, col). If the target is a MergedCell,
    walk to the top-left anchor of its merged range.
    """
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value
        return

    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            ws.cell(row=mr.min_row, column=mr.min_col).value = value
            return

    raise RuntimeError(
        f"Cannot write to {cell.coordinate}: merged range lookup failed."
    )


# ─────────────────────────────────────────────────────────────
# Playwright browser bootstrap (unchanged from your version)
# ─────────────────────────────────────────────────────────────

_PLAYWRIGHT_READY = False


def _ensure_playwright_browser():
    global _PLAYWRIGHT_READY
    if _PLAYWRIGHT_READY:
        return

    cache_root = Path.home() / ".cache" / "ms-playwright"
    already_installed = False

    if cache_root.exists():
        for d in cache_root.iterdir():
            if not d.is_dir() or "chromium" not in d.name.lower():
                continue
            candidates = (
                list(d.rglob("chrome-headless-shell"))
                + list(d.rglob("chrome"))
                + list(d.rglob("headless_shell"))
            )
            if candidates:
                already_installed = True
                break

    if not already_installed:
        cmd_candidates = [
            ["playwright", "install", "--with-deps", "chromium"],
            ["python", "-m", "playwright", "install", "--with-deps", "chromium"],
        ]
        installed_ok = False
        for cmd in cmd_candidates:
            try:
                subprocess.run(cmd, check=True, timeout=600,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                installed_ok = True
                break
            except FileNotFoundError:
                continue
            except subprocess.CalledProcessError as e:
                print(f"[playwright bootstrap] {cmd[0]} failed: {e}")
                continue

        if not installed_ok:
            try:
                subprocess.run(["python", "-m", "playwright", "install", "chromium"],
                               check=True, timeout=600,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            except Exception as e:
                print(f"[playwright bootstrap] final fallback failed: {e}")

    _PLAYWRIGHT_READY = True


# ─────────────────────────────────────────────────────────────
# Template filler
# ─────────────────────────────────────────────────────────────

# Labels as they appear in the template (any of these will match)
LABEL_ALIASES = {
    "rectifier_brand": [
        "1) EXISTING RECTIFIER SYSTEM BRAND - MODEL",
        "EXISTING RECTIFIER SYSTEM BRAND MODEL",
        "RECTIFIER SYSTEM BRAND MODEL",
    ],
    "max_modules": [
        "2) MAXIMUM RECTIFIER MODULES / INSTALLED",
        "MAXIMUM RECTIFIER MODULES INSTALLED",
        "MAX RECTIFIER MODULES INSTALLED",
    ],
    "module_rating_w": [
        "3) RECTIFIER MODULE RATING, WATTS / AMPS",
        "RECTIFIER MODULE RATING WATTS AMPS",
        "RECTIFIER MODULE RATING",
    ],
    "battery_brand_voltage": [
        "4) BATTERY BRAND / VOLTAGE RATING",
        "BATTERY BRAND VOLTAGE RATING",
        "BATTERY BRAND",
    ],
    "battery_banks_capacity": [
        "5) NUMBER OF BATTERIES in BANKS / CAPACITY per CELL (AH)",
        "NUMBER OF BATTERIES IN BANKS CAPACITY PER CELL AH",
        "NUMBER OF BATTERIES IN BANKS",
    ],
    "present_load_a": [
        "6) PRESENT LOAD READING, PANEL DISPLAY / CLAMP METER (A)",
        "PRESENT LOAD READING PANEL DISPLAY CLAMP METER A",
        "PRESENT LOAD READING",
    ],
    "modules_in_operation": [
        "7) RECTIFIER MODULES IN OPERATION",
        "RECTIFIER MODULES IN OPERATION",
        "MODULES IN OPERATION",
    ],
    "actual_float_voltage_v": [
        "8) ACTUAL FLOAT VOLTAGE (V)",
        "ACTUAL FLOAT VOLTAGE V",
        "ACTUAL FLOAT VOLTAGE",
    ],
    # Computed labels (these are the ones you'll confirm via the diagnostic)
    "total_full_load_a": [
        "TOTAL FULL LOAD CURRENT",
        "TOTAL FULL LOAD",
    ],
    "existing_rectifier_capacity_a": [
        "EXISTING RECTIFIER CAPACITY",
    ],
    "existing_battery_capacity_ah": [
        "EXISTING BATTERY CAPACITY",
    ],
    "available_rectifier_capacity_a": [
        "AVAILABLE RECTIFIER CAPACITY",
    ],
    "percent_utilization": [
        "PERCENT UTILIZATION",
        "PERCENT UTILISATION",
    ],
    "bbut_hours": [
        "BBUT",
        "BATTERY BACK UP TIME",
        "BATTERY BACKUP TIME",
    ],
}


def fill_template(template_path: str, out_path: str,
                  data: dict, sheet_name: str) -> str:
    """Copy template → out_path, write values into the C7 block."""
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )
    ws = wb[sheet_name]

    label_cells = _find_label_cells(ws)
    print(f"[fill_template] Found {len(label_cells)} labels in {sheet_name!r}")

    # ── Helper: write N columns to the right of a label's anchor
    def put(label_keys: list[str], value, *, col_offset: int = 1):
        cell = None
        for cand in label_keys:
            cell = _lookup(label_cells, cand)
            if cell:
                break
        if not cell:
            print(f"[fill_template] ⚠ label not found: {label_keys[0]!r}")
            return False
        _safe_set(ws, cell.row, cell.column + col_offset, value)
        return True

    # ── Write C7.1 (existing rectifier) fields
    # col_offset=1 assumes the input cell is immediately right of the label.
    # Adjust if your template differs (the diagnostic will tell you).
    put(LABEL_ALIASES["rectifier_brand"],
        data["rectifier_brand"], col_offset=8)
    put(LABEL_ALIASES["max_modules"],
        data["max_modules"], col_offset=8)
    put(LABEL_ALIASES["module_rating_w"],
        data["module_rating_w"], col_offset=8)
    put(LABEL_ALIASES["battery_brand_voltage"],
        f'{data["battery_brand"]} {data["battery_voltage"]}'.strip(),
        col_offset=8)
    put(LABEL_ALIASES["battery_banks_capacity"],
        f'{data["battery_banks"]} / {data["battery_capacity_ah"]}',
        col_offset=8)
    put(LABEL_ALIASES["present_load_a"],
        data["present_load_a"], col_offset=8)
    put(LABEL_ALIASES["modules_in_operation"],
        data["modules_in_operation"], col_offset=8)
    put(LABEL_ALIASES["actual_float_voltage_v"],
        data["actual_float_voltage_v"], col_offset=8)

    # ── Proposed load (fixed Nokia MF-2)
    start_row = _find_first_load_row(ws)
    _safe_set(ws, start_row, 1,  NOKIA_MF2_LOAD["load_no"])
    _safe_set(ws, start_row, 3,  NOKIA_MF2_LOAD["equipment"])
    _safe_set(ws, start_row, 12, NOKIA_MF2_LOAD["power_w"])
    _safe_set(ws, start_row, 15, NOKIA_MF2_LOAD["current_a"])
    _safe_set(ws, start_row, 21, NOKIA_MF2_LOAD["cable_awg"])
    _safe_set(ws, start_row, 27, NOKIA_MF2_LOAD["breaker_a"])
    for r in range(start_row + 1, start_row + 7):
        for col in (1, 3, 12, 15, 21, 27):
            _safe_set(ws, r, col, "")

    # ── Computed block
    comp = compute_sufficiency(data)
    for key, value in (
        ("total_full_load_a",              comp["total_full_load_a"]),
        ("existing_rectifier_capacity_a",  comp["existing_rectifier_capacity_a"]),
        ("existing_battery_capacity_ah",   comp["existing_battery_capacity_ah"]),
        ("available_rectifier_capacity_a", comp["available_rectifier_capacity_a"]),
        ("percent_utilization",            comp["percent_utilization"]),
        ("bbut_hours",                     comp["bbut_hours"]),
    ):
        put(LABEL_ALIASES[key], value, col_offset=1)

    wb.save(out_path)
    return out_path


def _find_first_load_row(ws) -> int:
    """Find the row after the 'Load No' header. Raise if not found."""
    target = _norm("Load No")
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and _norm(cell.value) == target:
                return cell.row + 1
    raise ValueError(
        "Could not find 'Load No' header. Run diagnose_template.py "
        "and adjust _find_first_load_row()."
    )


# ─────────────────────────────────────────────────────────────
# PNG rendering (unchanged)
# ─────────────────────────────────────────────────────────────

def render_sheet_png(xlsx_path: str, sheet_name: str, out_png: str) -> str:
    xlsx_path = str(xlsx_path)
    out_png = Path(out_png)
    out_png.parent.mkdir(parents=True, exist_ok=True)

    _ensure_playwright_browser()

    try:
        from xlsx2html import xlsx2html
    except ImportError as e:
        raise RuntimeError("Missing dependency: xlsx2html.") from e

    html_path = out_png.with_suffix(".html")
    with open(html_path, "w", encoding="utf-8") as f:
        xlsx2html(xlsx_path, sheet=sheet_name, output=f, locale="en_US")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError("Playwright not installed.") from e

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
        )
        page = browser.new_page(viewport={"width": 1400, "height": 1300})
        page.goto(f"file://{html_path.resolve()}")
        page.wait_for_timeout(800)
        page.screenshot(
            path=str(out_png), full_page=False,
            clip={"x": 0, "y": 0, "width": 1400, "height": 1300},
        )
        browser.close()

    if not out_png.exists() or out_png.stat().st_size < 1000:
        raise RuntimeError("Playwright did not produce a valid PNG.")
    return str(out_png)
