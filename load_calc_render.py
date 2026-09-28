# load_calc_render.py
"""
Fill load_calculation.xlsx (C7 block) by writing to fixed cells.

Only the sheets the user actually parsed are kept in the output.
Any other computation sheets in the template are deleted.

The user downloads the filled XLSX and screenshots the C7 block
themselves in Excel / LibreOffice / Google Sheets.
"""
from __future__ import annotations

import shutil

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from load_calc_helper import compute_sufficiency, NOKIA_MF2_LOAD


# ─────────────────────────────────────────────────────────────
# Cell mapping (same layout for all RS sheets)
# ─────────────────────────────────────────────────────────────

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
                  data: dict, sheet_name: str,
                  keep_sheets: list[str] | None = None) -> str:
    """
    Copy the template to `out_path`, fill the target sheet, and then
    delete every other computation sheet that isn't in `keep_sheets`.

    Args:
        template_path: source XLSX template
        out_path: destination XLSX
        data: dict of values to write (see CELL_MAP for keys)
        sheet_name: the sheet to write into (e.g. "RS1-computation")
        keep_sheets: list of computation sheets to preserve in the
            output. If None, only `sheet_name` is kept.
    """
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )

    # ── Build the value dict
    values = dict(data)
    values.pop("proposed_loads", None)

    voltage = data.get("battery_voltage") or 48
    values["module_rating_a"] = round(
        (data.get("module_rating_w") or 0) / voltage, 2
    )

    # ── Write each mapped cell
    ws = wb[sheet_name]
    written, missing = [], []
    for key, (row, col) in CELL_MAP.items():
        if key in values and values[key] not in (None, ""):
            _safe_set(ws, row, col, values[key])
            written.append(key)
        else:
            missing.append(key)

    print(f"[fill_template] wrote {len(written)} cells to {sheet_name!r}: {written}")
    if missing:
        print(f"[fill_template] ⚠ missing keys: {missing}")

    # ── Decide which computation sheets to keep
    if keep_sheets is None:
        keep_sheets = [sheet_name]

    # Never delete the current sheet, even if it's not in keep_sheets
    keep_set = set(keep_sheets) | {sheet_name}

    # Only delete sheets that look like computation sheets. This
    # preserves any non-RS sheets in the template (config, notes, etc.).
    def _is_computation_sheet(name: str) -> bool:
        n = name.lower()
        return n.startswith("rs") and "-computation" in n

    deleted = []
    for name in list(wb.sheetnames):
        if not _is_computation_sheet(name):
            continue
        if name in keep_set:
            continue
        del wb[name]
        deleted.append(name)

    if deleted:
        print(f"[fill_template] deleted unused sheets: {deleted}")

    wb.save(out_path)
    return out_path
