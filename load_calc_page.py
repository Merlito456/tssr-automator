# load_calc_page.py
"""
Streamlit page: AI Load Calculator for RS1 / RS2 / RS3 / RS4.
Called from app.py as:
    load_calc_page.render(ensure_workdir, persist)

Workflow:
  1. Pick rectifier (RS1..RS4)
  2. Copy prompt → paste into AI → copy JSON back
  3. Parse JSON & preview
  4. Fill template → download the filled XLSX
  5. Screenshot the C7 block in Excel and upload into Section 5
"""
from __future__ import annotations

from pathlib import Path

import streamlit as st

import load_calc_helper as lch
import load_calc_render as lcr


TEMPLATE = Path(__file__).parent / "data" / "load_calculation.xlsx"

RS_OPTIONS = {
    "RS1 — Rectifier 1":       "RS1-computation",
    "RS2 — Rectifier 2":       "RS2-computation",
    "RS3 — Rectifier 3":       "RS3-computation (existing)",
    "RS4 — Rectifier 4":       "RS4-computation (existing)",
}

# Short labels used inside the AI prompt ("RS1", "RS2", ...)
RS_SHORT = {
    "RS1 — Rectifier 1": "RS1",
    "RS2 — Rectifier 2": "RS2",
    "RS3 — Rectifier 3": "RS3",
    "RS4 — Rectifier 4": "RS4",
}


def render(ensure_workdir, persist) -> None:
    """Two-arg entry point called from app.py."""
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

    # 2) Prompt — customized so the AI extracts the right rectifier
    with st.expander("📋 Step 1 — Copy this prompt", expanded=False):
        st.code(lch.build_prompt(rs_short), language="text")

    # 3) Paste JSON
    with st.expander("📥 Step 2 — Paste AI JSON response", expanded=True):
        response = st.text_area(
            "Paste the AI's JSON here:",
            height=220,
            key="load_calc_json",
            placeholder='{"rectifier_brand": "Eltek", ...}',
        )

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("✔ Parse JSON & preview",
                         type="primary",
                         use_container_width=True,
                         key="lc_parse"):
                parsed = lch.extract_json(response)
                if not parsed:
                    st.error("❌ Could not parse JSON.")
                else:
                    clean, warnings = lch.validate(parsed)
                    st.session_state["load_calc_data"] = clean

                    # Push values into site_data so the DOCX gets them too.
                    site = dict(st.session_state.get("site_data") or {})
                    for k in (
                        "rectifier_brand", "max_modules", "module_rating_w",
                        "battery_brand", "battery_voltage", "battery_banks",
                        "battery_capacity_ah", "present_load_a",
                        "modules_in_operation", "actual_float_voltage_v",
                    ):
                        site[k] = clean.get(k, "")

                    # Computed sufficiency fields (dynamic voltage)
                    comp = lch.compute_sufficiency(clean)
                    for k in ("existing_rectifier_capacity_a",
                              "existing_battery_capacity_ah",
                              "total_full_load_a",
                              "available_rectifier_capacity_a",
                              "percent_utilization", "bbut_hours",
                              "system_voltage_v"):
                        site[k] = comp.get(k, 0)

                    # module_rating_a uses the same dynamic voltage
                    voltage = clean.get("battery_voltage") or 48
                    site["module_rating_a"] = round(
                        (clean.get("module_rating_w") or 0) / voltage, 2
                    )
                    st.session_state.site_data = site

                    if warnings:
                        st.warning("⚠️ " + "\n- ".join(warnings))
                    st.success("✅ Parsed OK")
                    persist()

        with col_b:
            if st.button("🗑 Clear", use_container_width=True, key="lc_clear"):
                st.session_state["load_calc_data"] = None
                st.session_state.pop("load_calc_json", None)
                persist()
                st.rerun()

    # 4) Preview + fill
    data = st.session_state.get("load_calc_data")
    if not data:
        return

    st.divider()
    st.markdown("### Preview")
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

    # 5) Fill template → download XLSX
    if st.button("📄 Fill template → download XLSX",
                 type="primary",
                 use_container_width=True,
                 key="lc_fill"):
        workdir = Path(ensure_workdir())
        out_xlsx = workdir / f"{sheet_name.replace(' ', '_')}_filled.xlsx"

        # Inject site identity from the masterlist (used by M4 / M5).
        site = st.session_state.get("site_data") or {}
        payload = {
            **data,
            "SITE_NAME": site.get("site_name", ""),
            "SITE_ID":   site.get("site_id", ""),
        }

        with st.spinner("Filling template…"):
            try:
                lcr.fill_template(str(TEMPLATE), str(out_xlsx),
                                  payload, sheet_name)
            except Exception as e:
                st.error(f"❌ Fill failed: {e}")
                st.exception(e)
                return

        st.success(f"✅ Filled `{rs_label}` — download below.")
        st.caption(
            "Open the file in Excel / LibreOffice / Google Sheets, then "
            "screenshot the C7 block and upload it into the matching "
            "image slot in **Section 5 · Images**."
        )

        st.download_button(
            "⬇️ Download filled XLSX",
            data=Path(out_xlsx).read_bytes(),
            file_name=out_xlsx.name,
            mime=("application/vnd.openxmlformats-officedocument"
                  ".spreadsheetml.sheet"),
            use_container_width=True,
        )
