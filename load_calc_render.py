# load_calc_render.py
"""
Fill load_calculation.xlsx (C7 block) by writing to fixed cells.

The Excel template stores Jinja-style placeholders in specific cells
(see CELL_MAP below). This module writes the matching values directly —
no placeholder scanning, no browser, no PNG render.

The user downloads the filled XLSX and screenshots the C7 block
themselves in Excel / LibreOffice / Google Sheets.
"""
from __future__ import annotations

import shutil

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from load_calc_helper import compute_sufficiency, NOKIA_MF2_LOAD


# ─────────────────────────────────────────────────────────────
# Cell mapping for the template
# ─────────────────────────────────────────────────────────────
# Keys are the placeholder names; values are (row, column) 1-based.
# Column letters → 1-based indexes:
#   M  = 13
#   Y  = 25
#   AE = 31

CELL_MAP = {
    "SITE_NAME":              (4,  13),   # M4
    "SITE_ID":                (5,  13),   # M5
    "max_modules":            (9,  31),   # AE9
    "module_rating_w":        (10, 25),   # Y10
    "module_rating_a":        (10, 31),   # AE10
    "battery_brand":          (11, 25),   # Y11
    "battery_voltage":        (11, 31),   # AE11
    "battery_banks":          (12, 25),   # Y12
    "battery_capacity_ah":    (12, 31),   # AE12
    "present_load_a":         (13, 25),   # Y13
    "modules_in_operation":   (14, 25),   # Y14
    "actual_float_voltage_v": (15, 25),   # Y15
}


# ─────────────────────────────────────────────────────────────
# Merged-safe writer
# ─────────────────────────────────────────────────────────────

def _safe_set(ws, row: int, col: int, value):
    """Write to (row, col). If it's a MergedCell, target the anchor."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value
        return
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            ws.cell(row=mr.min_row, column=mr.min_col).value = value
            return
    raise RuntimeError(f"Cannot write to {cell.coordinate}")


# ─────────────────────────────────────────────────────────────
# Template filler
# ─────────────────────────────────────────────────────────────

def fill_template(template_path: str, out_path: str,
                  data: dict, sheet_name: str) -> str:
    """
    Copy the template to `out_path` and write values into the cells
    defined by CELL_MAP on the given sheet.

    `data` must contain all the keys referenced in CELL_MAP:
      SITE_NAME, SITE_ID, max_modules, module_rating_w,
      module_rating_a, battery_brand, battery_voltage,
      battery_banks, battery_capacity_ah, present_load_a,
      modules_in_operation, actual_float_voltage_v
    """
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )
    ws = wb[sheet_name]

    # ── Build the value dict
    values = dict(data)
    values.pop("proposed_loads", None)

    # Derived: module_rating_a = W / system voltage
    voltage = data.get("battery_voltage") or 48
    values["module_rating_a"] = round(
        (data.get("module_rating_w") or 0) / voltage, 2
    )

    # ── Write each mapped cell
    written, missing = [], []
    for key, (row, col) in CELL_MAP.items():
        if key in values and values[key] not in (None, ""):
            _safe_set(ws, row, col, values[key])
            written.append(key)
        else:
            missing.append(key)

    print(f"[fill_template] wrote {len(written)} cells to {sheet_name!r}: {written}")
    if missing:
        print(f"[fill_template] ⚠ missing keys (left untouched): {missing}")

    wb.save(out_path)
    return out_path
