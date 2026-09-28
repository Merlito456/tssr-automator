# load_calc_render.py
"""
Fill load_calculation.xlsx (C7 block) by writing to fixed cells.

Two entry points:
  - fill_template()      — one sheet, one file (legacy)
  - fill_all_templates() — many sheets, one combined file (preferred)

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
# Internal helpers
# ─────────────────────────────────────────────────────────────

def _build_values(data: dict, site_info: dict | None = None) -> dict:
    """Compose the value dict for a single sheet."""
    values = dict(data)
    values.pop("proposed_loads", None)

    if site_info:
        values["SITE_NAME"] = site_info.get("site_name", "")
        values["SITE_ID"]   = site_info.get("site_id", "")

    # Derived: module_rating_a = W / system voltage
    voltage = data.get("battery_voltage") or 48
    values["module_rating_a"] = round(
        (data.get("module_rating_w") or 0) / voltage, 2
    )
    return values


def _is_computation_sheet(name: str) -> bool:
    n = name.lower()
    return n.startswith("rs") and "-computation" in n


def _prune_unused(wb, keep_set: set[str]) -> list[str]:
    """Delete computation sheets not in keep_set. Return names deleted."""
    deleted = []
    for name in list(wb.sheetnames):
        if not _is_computation_sheet(name):
            continue
        if name in keep_set:
            continue
        del wb[name]
        deleted.append(name)
    return deleted


# ─────────────────────────────────────────────────────────────
# Entry point 1 — single sheet (legacy)
# ─────────────────────────────────────────────────────────────

def fill_template(template_path: str, out_path: str,
                  data: dict, sheet_name: str,
                  keep_sheets: list[str] | None = None,
                  site_info: dict | None = None) -> str:
    """
    Fill one sheet in the template and prune the rest.
    """
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )

    values = _build_values(data, site_info)
    ws = wb[sheet_name]

    written, missing = [], []
    for key, (row, col) in CELL_MAP.items():
        if key in values and values[key] not in (None, ""):
            _safe_set(ws, row, col, values[key])
            written.append(key)
        else:
            missing.append(key)

    print(f"[fill_template] wrote {len(written)} cells to {sheet_name!r}")
    if missing:
        print(f"[fill_template] ⚠ missing keys: {missing}")

    if keep_sheets is None:
        keep_sheets = [sheet_name]
    keep_set = set(keep_sheets) | {sheet_name}

    deleted = _prune_unused(wb, keep_set)
    if deleted:
        print(f"[fill_template] deleted unused sheets: {deleted}")

    wb.save(out_path)
    return out_path


# ─────────────────────────────────────────────────────────────
# Entry point 2 — many sheets, one combined file
# ─────────────────────────────────────────────────────────────

def fill_all_templates(template_path: str, out_path: str,
                       data_by_sheet: dict,
                       keep_sheets: list[str] | None = None,
                       site_info: dict | None = None) -> str:
    """
    Fill multiple RS sheets in a single output XLSX.

    Args:
        template_path: source XLSX template
        out_path: destination XLSX
        data_by_sheet: {sheet_name: values_dict}
        keep_sheets: sheets to preserve. Defaults to data_by_sheet keys.
        site_info: optional {"site_name": ..., "site_id": ...}

    Returns:
        out_path
    """
    if not data_by_sheet:
        raise ValueError("No sheets to fill.")

    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    missing_sheets = [s for s in data_by_sheet if s not in wb.sheetnames]
    if missing_sheets:
        raise ValueError(
            f"Sheets not found in template: {missing_sheets}. "
            f"Available: {wb.sheetnames}"
        )

    total_written = 0

    for sheet_name, data in data_by_sheet.items():
        ws = wb[sheet_name]
        values = _build_values(data, site_info)

        written, missing = [], []
        for key, (row, col) in CELL_MAP.items():
            if key in values and values[key] not in (None, ""):
                _safe_set(ws, row, col, values[key])
                written.append(key)
            else:
                missing.append(key)

        total_written += len(written)
        print(f"[fill_all] {sheet_name}: wrote {len(written)} cells")
        if missing:
            print(f"[fill_all] {sheet_name}: ⚠ missing {missing}")

    if keep_sheets is None:
        keep_sheets = list(data_by_sheet.keys())
    keep_set = set(keep_sheets)

    deleted = _prune_unused(wb, keep_set)
    if deleted:
        print(f"[fill_all] deleted unused sheets: {deleted}")

    print(f"[fill_all] total cells written across all sheets: {total_written}")
    wb.save(out_path)
    return out_path
