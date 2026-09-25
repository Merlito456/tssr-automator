# load_calc_page.py
"""
Streamlit page: AI Load Calculator for RS1 / RS2 / RS3 / RS4.
Called from app.py as:
    load_calc_page.render(ensure_workdir, persist)
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

SLOT_BY_RS = {
    "RS1 — Rectifier 1": "RS1_load_calc_img",
    "RS2 — Rectifier 2": "RS2_load_calc_img",
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
    st.caption(f"→ Will write to sheet `{sheet_name}`.")

    # 2) Prompt
    with st.expander("📋 Step 1 — Copy this prompt", expanded=False):
        st.code(lch.build_prompt(), language="text")

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
                    comp = lch.compute_sufficiency(clean)
                    for k in ("existing_rectifier_capacity_a",
                              "existing_battery_capacity_ah",
                              "total_full_load_a",
                              "available_rectifier_capacity_a",
                              "percent_utilization", "bbut_hours"):
                        site[k] = comp.get(k, 0)
                    site["module_rating_a"] = round(
                        (clean.get("module_rating_w") or 0) / 48.0, 2)
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

    # 4) Preview + render
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

    if st.button("🖼 Render filled sheet → PNG",
                 type="primary",
                 use_container_width=True,
                 key="lc_render"):
        workdir = Path(ensure_workdir())
        out_xlsx = workdir / f"{sheet_name.replace(' ', '_')}.xlsx"
        out_png = workdir / "uploads" / f"{sheet_name.replace(' ', '_')}.png"
        out_png.parent.mkdir(parents=True, exist_ok=True)

        with st.spinner("Filling template…"):
            try:
                lcr.fill_template(str(TEMPLATE), str(out_xlsx), data, sheet_name)
            except Exception as e:
                st.error(f"❌ Fill failed: {e}")
                st.exception(e)
                return

        with st.spinner("Rendering PNG…"):
            try:
                lcr.render_sheet_png(str(out_xlsx), sheet_name, str(out_png))
            except Exception as e:
                st.error(f"❌ Render failed: {e}")
                st.info("Fallback: filled XLSX is still available.")
                st.download_button(
                    "⬇️ Download filled XLSX",
                    data=Path(out_xlsx).read_bytes(),
                    file_name=out_xlsx.name,
                    mime=("application/vnd.openxmlformats-officedocument"
                          ".spreadsheetml.sheet"),
                )
                return

        slot = SLOT_BY_RS.get(rs_label)
        if slot:
            st.session_state.image_map[slot] = str(out_png)
            persist()
            st.success(f"✅ Saved to slot `{slot}`")
        else:
            st.info(f"Rendered but {rs_label} has no Nokia slot — kept at `{out_png}`")

        st.image(str(out_png), caption=f"{rs_label} — C7 block")
        st.download_button(
            "⬇️ Download PNG",
            data=Path(out_png).read_bytes(),
            file_name=out_png.name,
            mime="image/png",
        )
