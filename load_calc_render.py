# load_calc_render.py
"""
Fill the load_calculation.xlsx template with data and render the
'C7. RectifierN Computation' block to a PNG image.

Merged-safe: writes to the top-left anchor of any merged cell range.
"""
from __future__ import annotations
import shutil
from pathlib import Path

from openpyxl import load_workbook

from load_calc_helper import compute_sufficiency


SHEET_BY_INDEX = {
    1: "RS1-computation",
    2: "RS2-computation",
    3: "RS3-computation (existing)",
    4: "RS4-computation (existing)",
}


# ─────────────────────────────────────────────────────────────
# Merged-safe cell writer
# ─────────────────────────────────────────────────────────────

def _safe_set(ws, row: int, col: int, value):
    """
    Write `value` at (row, col). If the target is a MergedCell,
    walk to the top-left anchor of its merged range and write there.
    """
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ != "MergedCell":
        cell.value = value
        return

    # Find the merged range that contains (row, col)
    for mr in ws.merged_cells.ranges:
        if (mr.min_row <= row <= mr.max_row) and \
           (mr.min_col <= col <= mr.max_col):
            ws.cell(row=mr.min_row, column=mr.min_col).value = value
            return

    # Fallback — shouldn't happen
    try:
        cell.value = value
    except AttributeError:
        pass


# ─────────────────────────────────────────────────────────────
# Template filler
# ─────────────────────────────────────────────────────────────

def fill_template(
    template_path: str,
    out_path: str,
    data: dict,
    sheet_name: str,
) -> str:
    """
    Copy template -> out_path, write values into the C7 block on sheet_name.
    Returns out_path.
    """
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not in workbook")
    ws = wb[sheet_name]

    # ---- Fill C7.1 ----
    # These cells are found by searching for the label text, so the
    # template can shift rows without breaking us.
    label_to_cell = _find_label_cells(ws)

    def put(label: str, value, *, row_offset: int = 0, col_offset: int = 8):
        """Write `value` at `col_offset` columns to the right of the label."""
        cell = label_to_cell.get(label)
        if not cell:
            return
        r, c = cell.row, cell.column
        _safe_set(ws, r + row_offset, c + col_offset, value)

    # C7.1 block (label -> value is 8 cols right in the sample layout)
    put("1) EXISTING RECTIFIER SYSTEM BRAND - MODEL",
        data["rectifier_brand"])
    put("2) MAXIMUM RECTIFIER MODULES / INSTALLED",
        data["max_modules"])
    put("3) RECTIFIER MODULE RATING, WATTS / AMPS",
        data["module_rating_w"])
    put("4) BATTERY BRAND / VOLTAGE RATING",
        f'{data["battery_brand"]} {data["battery_voltage"]}'.strip())
    put("5) NUMBER OF BATTERIES in BANKS / CAPACITY per CELL (AH)",
        f'{data["battery_banks"]} / {data["battery_capacity_ah"]}')
    put("6) PRESENT LOAD READING, PANEL DISPLAY / CLAMP METER (A)",
        data["present_load_a"])
    put("7) RECTIFIER MODULES IN OPERATION",
        data["modules_in_operation"])
    put("8) ACTUAL FLOAT VOLTAGE (V)",
        data["actual_float_voltage_v"])

    # ---- Fill C7.2 (proposed loads) — merged-safe ----
    start_row = _find_first_load_row(ws)
    for i, load in enumerate(data.get("proposed_loads", [])):
        r = start_row + i
        _safe_set(ws, r, 1,  load.get("load_no", i + 1))
        _safe_set(ws, r, 3,  load.get("equipment", ""))
        _safe_set(ws, r, 12, load.get("power_w", 0))
        _safe_set(ws, r, 15, load.get("current_a", 0))
        _safe_set(ws, r, 21, load.get("cable_awg", ""))
        _safe_set(ws, r, 27, load.get("breaker_a", ""))

    # ---- Compute sufficiency, write into the computed cells ----
    comp = compute_sufficiency(data)

    _set_by_label(ws, label_to_cell, "Total Full Load Current",
                  comp["total_full_load_a"])
    _set_by_label(ws, label_to_cell, "Existing Rectifier Capacity:",
                  comp["existing_rectifier_capacity_a"])
    _set_by_label(ws, label_to_cell, "Existing Battery Capacity:",
                  comp["existing_battery_capacity_ah"])
    _set_by_label(ws, label_to_cell, "Available Rectifier Capacity =",
                  comp["available_rectifier_capacity_a"])
    _set_by_label(ws, label_to_cell, "Percent Utilization",
                  comp["percent_utilization"])
    _set_by_label(ws, label_to_cell, "BBUT =",
                  comp["bbut_hours"])

    wb.save(out_path)
    return out_path


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def _find_label_cells(ws) -> dict:
    """Map exact label text -> cell object."""
    found = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                key = cell.value.strip()
                if key and key not in found:
                    found[key] = cell
    return found


def _find_first_load_row(ws) -> int:
    """Find the row after the 'Load No' header in the C7.2 block."""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "Load No":
                return cell.row + 1
    print("⚠️ Could not find 'Load No' header — proposed loads "
          "will be written from row 21 by default.")
    return 21


def _set_by_label(ws, label_cells: dict, prefix: str, value):
    """Find the first label cell whose value starts with `prefix`, write value 1 col right."""
    for text, cell in label_cells.items():
        if text.startswith(prefix):
            _safe_set(ws, cell.row, cell.column + 1, value)
            return


# ─────────────────────────────────────────────────────────────
# PNG rendering
# ─────────────────────────────────────────────────────────────

def render_sheet_png(xlsx_path: str, sheet_name: str, out_png: str) -> str:
    """
    Render the named sheet to a PNG.
    Strategy: convert the sheet region to HTML via xlsx2html, then
    use playwright (Chromium) to screenshot it.
    """
    from xlsx2html import xlsx2html
    from playwright.sync_api import sync_playwright

    html_path = Path(out_png).with_suffix(".html")

    with open(html_path, "w", encoding="utf-8") as f:
        xlsx2html(
            xlsx_path,
            sheet_name=sheet_name,
            output=f,
            locale="en_US",
        )

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1300})
        page.goto(f"file://{html_path.resolve()}")
        page.wait_for_timeout(400)
        page.screenshot(
            path=out_png,
            full_page=True,
        )
        browser.close()

    return out_png
