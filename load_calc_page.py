# load_calc_page.py
"""
Streamlit page: AI Load Calculator for RS1 / RS2 / RS3 / RS4.
Called from app.py as:
    load_calc_page.render(ensure_workdir, persist)

Workflow:
  1. Pick rectifier (RS1..RS4)
  2. Copy prompt → paste into AI → copy JSON back
  3. Parse JSON & preview  (kept per-RS in session_state)
  4. Fill & download ONE combined XLSX (only filled sheets kept)
  5. Screenshot each C7 block and upload into Section 5
"""
from __future__ import annotations

from pathlib import Path

import streamlit as st

import load_calc_helper as lch
import load_calc_render as lcr


TEMPLATE = Path(__file__).parent / "data" / "load_calculation.xlsx"

RS_OPTIONS = {
    "RS1 — Rectifier 1": "RS1-computation",
    "RS2 — Rectifier 2": "RS2-computation",
    "RS3 — Rectifier 3": "RS3-computation",
    "RS4 — Rectifier 4": "RS4-computation",
}

RS_SHORT = {
    "RS1 — Rectifier 1": "RS1",
    "RS2 — Rectifier 2": "RS2",
    "RS3 — Rectifier 3": "RS3",
    "RS4 — Rectifier 4": "RS4",
}


# ─────────────────────────────────────────────────────────────
# Per-RS storage helpers
# ─────────────────────────────────────────────────────────────

def _store() -> dict:
    st.session_state.setdefault("load_calc_data_by_rs", {})
    return st.session_state["load_calc_data_by_rs"]


def _get_for(rs_label: str) -> dict | None:
    return _store().get(rs_label)


def _set_for(rs_label: str, data: dict) -> None:
    _store()[rs_label] = data


def _clear_for(rs_label: str) -> None:
    _store().pop(rs_label, None)
    st.session_state.pop(f"load_calc_json_{rs_label}", None)


def render(ensure_workdir, persist) -> None:
    st.subheader("⚡ AI Load Calculator (C7)")

    if not TEMPLATE.exists():
        st.error(f"❌ Template not found: `{TEMPLATE}`")
        st.info("Place your `load_calculation.xlsx` in the `data/` folder.")
        return

    # 1) Pick which rectifier system
    rs_label = st.radio(
        "Which rectifier system are you documenting?",
        options=list(RS_OPTIONS.keys()),
        horizontal=True,
        key="load_calc_rs",
    )
    sheet_name = RS_OPTIONS[rs_label]
    rs_short = RS_SHORT.get(rs_label, "EXISTING")
    st.caption(f"→ Will write to sheet `{sheet_name}` (extracting **{rs_short}**).")

    # Indicator: which RSs have data already stored
    stored = _store()
    if stored:
        done = ", ".join(sorted(RS_SHORT.get(k, k) for k in stored))
        st.info(f"📦 Parsed so far: {done}")
    else:
        st.caption("No rectifier parsed yet.")

    # 2) Prompt
    with st.expander("📋 Step 1 — Copy this prompt", expanded=False):
        st.code(lch.build_prompt(rs_short), language="text")

    # 3) Paste JSON (per-RS textarea)
    with st.expander("📥 Step 2 — Paste AI JSON response", expanded=True):
        response = st.text_area(
            "Paste the AI's JSON here:",
            height=220,
            key=f"load_calc_json_{rs_label}",
            placeholder='{"rectifier_brand": "Eltek", ...}',
        )

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("✔ Parse JSON & preview",
                         type="primary",
                         use_container_width=True,
                         key=f"lc_parse_{rs_label}"):
                parsed = lch.extract_json(response)
                if not parsed:
                    st.error("❌ Could not parse JSON.")
                else:
                    clean, warnings = lch.validate(parsed)
                    _set_for(rs_label, clean)

                    # Push into site_data for the DOCX too
                    site = dict(st.session_state.get("site_data") or {})
                    for k in (
                        "rectifier_brand", "max_modules", "module_rating_w",
                        "battery_brand", "battery_voltage", "battery_banks",
                        "battery_capacity_ah", "present_load_a",
                        "modules_in_operation", "actual_float_voltage_v",
                    ):
                        site[k] = clean.get(k, "")

                    comp = lch.compute_sufficiency(clean)
                    for k in ("existing_rectifier_capacity_a",
                              "existing_battery_capacity_ah",
                              "total_full_load_a",
                              "available_rectifier_capacity_a",
                              "percent_utilization", "bbut_hours",
                              "system_voltage_v"):
                        site[k] = comp.get(k, 0)

                    voltage = clean.get("battery_voltage") or 48
                    site["module_rating_a"] = round(
                        (clean.get("module_rating_w") or 0) / voltage, 2
                    )
                    st.session_state.site_data = site

                    if warnings:
                        st.warning("⚠️ " + "\n- ".join(warnings))
                    st.success(f"✅ Parsed {rs_short} OK")
                    persist()

        with col_b:
            if st.button("🗑 Clear this RS",
                         use_container_width=True,
                         key=f"lc_clear_{rs_label}"):
                _clear_for(rs_label)
                persist()
                st.rerun()

    # 4) Preview for the currently selected RS
    data = _get_for(rs_label)
    if not data:
        st.info(f"No data parsed for {rs_short} yet.")
    else:
        st.divider()
        st.markdown(f"### Preview — {rs_short}")
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Existing rectifier**")
            st.json({k: v for k, v in data.items() if k != "proposed_loads"})
        with col2:
            st.write("**Proposed load (fixed)**")
            st.json(data.get("proposed_loads", []))

        comp = lch.compute_sufficiency(data)
        st.write("**Computed**")
        st.json(comp)

    # 5) SINGLE Fill & download button — for ALL parsed RSs
    if not stored:
        return

    st.divider()
    st.markdown("### Fill all parsed RSs into one XLSX")

    parsed_summary = ", ".join(
        sorted(RS_SHORT.get(k, k) for k in stored)
    )
    st.caption(f"Sheets that will be written: {parsed_summary}")

    if st.button("📄 Fill all → download XLSX",
                 type="primary",
                 use_container_width=True,
                 key="lc_fill_all"):
        workdir = Path(ensure_workdir())
        out_xlsx = workdir / "load_calculation_filled.xlsx"

        site = st.session_state.get("site_data") or {}
        site_info = {
            "site_name": site.get("site_name", ""),
            "site_id":   site.get("site_id", ""),
        }

        # Build {sheet_name: data}
        data_by_sheet = {
            RS_OPTIONS[lbl]: d
            for lbl, d in stored.items()
            if lbl in RS_OPTIONS
        }

        with st.spinner("Filling template…"):
            try:
                lcr.fill_all_templates(
                    str(TEMPLATE), str(out_xlsx),
                    data_by_sheet,
                    keep_sheets=list(data_by_sheet.keys()),
                    site_info=site_info,
                )
            except Exception as e:
                st.error(f"❌ Fill failed: {e}")
                st.exception(e)
                return

        st.success(f"✅ Filled {len(data_by_sheet)} sheet(s) — download below.")
        st.caption(
            "Open the file in Excel / LibreOffice / Google Sheets, screenshot "
            "each C7 block, and upload them into the matching image slots in "
            "**Section 5 · Images**."
        )

        st.download_button(
            "⬇️ Download filled XLSX",
            data=Path(out_xlsx).read_bytes(),
            file_name=out_xlsx.name,
            mime=("application/vnd.openxmlformats-officedocument"
                  ".spreadsheetml.sheet"),
            use_container_width=True,
        )
