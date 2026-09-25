# load_calc_render.py
"""
Fill load_calculation.xlsx (C7 block) and render it to a PNG.

The Excel template stores Jinja-style placeholders directly in the
value cells, e.g. `{{max_modules}}`. This module scans EVERY cell of
the target sheet and substitutes any `{{key}}` with the matching
value from the AI-provided dict.

Merged-safe: if the target cell is inside a merged range, writes to
the top-left anchor.

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
# Placeholder scan & replace
# ─────────────────────────────────────────────────────────────

PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")


def _safe_set(ws, row: int, col: int, value):
    """Write to (row, col). If the cell is a MergedCell, write to the
    top-left anchor of its range instead."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value
        return
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            ws.cell(row=mr.min_row, column=mr.min_col).value = value
            return
    raise RuntimeError(f"Cannot write to {cell.coordinate}")


def replace_placeholders(ws, values: dict) -> tuple[int, list[str]]:
    """
    Walk every cell in the sheet. For each cell whose value is
    a string containing `{{key}}`:
      - If the whole cell is exactly one placeholder -> write raw value
      - Otherwise -> substitute inline, preserving surrounding text.
    Returns (count_replaced, list_of_unmatched_keys).
    """
    replaced = 0
    unmatched = []

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str) or "{{" not in v:
                continue

            # Case 1: the entire cell is a single placeholder
            whole = PLACEHOLDER_RE.fullmatch(v.strip())
            if whole:
                key = whole.group(1)
                if key in values:
                    _safe_set(ws, cell.row, cell.column, values[key])
                    replaced += 1
                else:
                    unmatched.append(key)
                continue

            # Case 2: placeholders embedded in surrounding text
            def sub(m):
                nonlocal replaced, unmatched
                key = m.group(1)
                if key in values:
                    replaced += 1
                    return str(values[key])
                unmatched.append(key)
                return m.group(0)

            new_text = PLACEHOLDER_RE.sub(sub, v)
            if new_text != v:
                _safe_set(ws, cell.row, cell.column, new_text)

    return replaced, unmatched


# ─────────────────────────────────────────────────────────────
# Playwright bootstrap (unchanged)
# ─────────────────────────────────────────────────────────────

_PLAYWRIGHT_READY = False


def _ensure_playwright_browser():
    global _PLAYWRIGHT_READY
    if _PLAYWRIGHT_READY:
        return
    cache_root = Path.home() / ".cache" / "ms-playwright"
    already = False
    if cache_root.exists():
        for d in cache_root.iterdir():
            if not d.is_dir() or "chromium" not in d.name.lower():
                continue
            if (list(d.rglob("chrome-headless-shell"))
                    or list(d.rglob("chrome"))
                    or list(d.rglob("headless_shell"))):
                already = True
                break
    if not already:
        for cmd in (["playwright", "install", "--with-deps", "chromium"],
                    ["python", "-m", "playwright", "install", "--with-deps", "chromium"]):
            try:
                subprocess.run(cmd, check=True, timeout=600,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                break
            except FileNotFoundError:
                continue
            except subprocess.CalledProcessError as e:
                print(f"[playwright bootstrap] {cmd[0]} failed: {e}")
    _PLAYWRIGHT_READY = True


# ─────────────────────────────────────────────────────────────
# Template filler
# ─────────────────────────────────────────────────────────────

def fill_template(template_path: str, out_path: str,
                  data: dict, sheet_name: str) -> str:
    """
    Fill `{{...}}` placeholders in the given sheet with values from `data`.
    Additionally computes sufficiency values and exposes them under the
    same placeholder convention.
    """
    shutil.copy(template_path, out_path)
    wb = load_workbook(out_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )
    ws = wb[sheet_name]

    # ── Build the substitution dict
    values = dict(data)                      # AI-provided fields
    values.pop("proposed_loads", None)       # not a scalar

    # Expose proposed-load fields as `proposed_*` placeholders too
    values["proposed_equipment"] = NOKIA_MF2_LOAD["equipment"]
    values["proposed_power_w"]   = NOKIA_MF2_LOAD["power_w"]
    values["proposed_current_a"] = NOKIA_MF2_LOAD["current_a"]
    values["proposed_breaker_a"] = NOKIA_MF2_LOAD["breaker_a"]

    # Compute sufficiency and expose it
    comp = compute_sufficiency(data)
    for k, v in comp.items():
        values[k] = v

    # Convenience: expose module_rating_a (W / 48 V) — the template
    # references this in the RECTIFIER MODULE RATING row.
    w = data.get("module_rating_w") or 0
    values["module_rating_a"] = round(w / 48.0, 2)

    # ── Replace placeholders
    replaced, unmatched = replace_placeholders(ws, values)
    print(f"[fill_template] replaced {replaced} placeholders in {sheet_name!r}")
    if unmatched:
        print(f"[fill_template] ⚠ unmatched keys: {sorted(set(unmatched))}")

    wb.save(out_path)
    return out_path


# ─────────────────────────────────────────────────────────────
# PNG rendering (unchanged)
# ─────────────────────────────────────────────────────────────

def render_sheet_png(xlsx_path: str, sheet_name: str, out_png: str) -> str:
    xlsx_path = str(xlsx_path)
    out_png = Path(out_png)
    out_png.parent.mkdir(parents=True, exist_ok=True)

    _ensure_playwright_browser()

    from xlsx2html import xlsx2html
    html_path = out_png.with_suffix(".html")
    with open(html_path, "w", encoding="utf-8") as f:
        xlsx2html(xlsx_path, sheet=sheet_name, output=f, locale="en_US")

    from playwright.sync_api import sync_playwright
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
