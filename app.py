# app.py
"""
TSSR Automator v0.2 — Dual Source
Excel masterlist + Ericsson TSSR + user-uploaded images.
"""
import shutil
import tempfile
import time
from pathlib import Path

import pandas as pd
import streamlit as st

import core
from excel_loader import SiteMasterlist, compose_address, compose_coords


# ─────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TSSR Automator",
    page_icon="📡",
    layout="wide",
)


# ─────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────
DEFAULTS = {
    "workdir": None,
    "masterlist": None,
    "ericsson_docx": None,
    "ericsson_images": [],
    "ericsson_fields": {},
    "site_data": {},
    "image_map": {},
    "materials": {},
}
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)


def reset():
    if st.session_state.workdir and Path(st.session_state.workdir).exists():
        shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


def save_upload(uploaded_file) -> str:
    workdir = Path(st.session_state.workdir) / "uploads"
    workdir.mkdir(parents=True, exist_ok=True)
    dest = workdir / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return str(dest)


# ─────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────
col_a, col_b = st.columns([4, 1])
with col_a:
    st.title("📡 TSSR Automator")
    st.caption("Excel masterlist + Ericsson TSSR → Nokia TSSR DOCX")
with col_b:
    if st.button("↺ Reset", use_container_width=True):
        reset()
        st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 1 — Load Masterlist
# ═════════════════════════════════════════════════════════════
st.subheader("1 · Load Site Masterlist")

if st.session_state.masterlist is None:
    default_xlsx = Path("data/MINDANAO_Site_Activity_Monitoring_OLT_PROJECT.xlsx")
    if default_xlsx.exists():
        st.info(f"Using bundled masterlist: `{default_xlsx.name}`")
        st.session_state.masterlist = SiteMasterlist(str(default_xlsx))
    else:
        xlsx_file = st.file_uploader(
            "Upload the MINDANAO site masterlist (.xlsx)",
            type=["xlsx"],
        )
        if xlsx_file:
            workdir = tempfile.mkdtemp(prefix="tssr_")
            st.session_state.workdir = workdir
            xlsx_path = Path(workdir) / xlsx_file.name
            xlsx_path.write_bytes(xlsx_file.getbuffer())
            st.session_state.masterlist = SiteMasterlist(str(xlsx_path))
            st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 2 — Look Up Site by PLAID
# ═════════════════════════════════════════════════════════════
if st.session_state.masterlist:
    st.divider()
    st.subheader("2 · Select Site")

    plaids = st.session_state.masterlist.list_plaids()
    site_id = st.selectbox(
        "PLAID",
        options=[""] + plaids,
        help="Select from the masterlist",
    )

    if site_id:
        site = st.session_state.masterlist.get_site(site_id)
        st.session_state.site_data = site

        # Preview
        st.markdown("**Excel data loaded:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.text_input("Site ID",   site.get("site_id", ""), disabled=True)
            st.text_input("Site Name", site.get("site_name", ""), disabled=True)
        with col2:
            st.text_input("Region",    site.get("region", ""), disabled=True)
            st.text_input("Towerco",   site.get("towerco", ""), disabled=True)
        with col3:
            st.text_input("FO Name",   site.get("fo_name", ""), disabled=True)
            st.text_input("FO Mobile", site.get("fo_mobile", ""), disabled=True)


# ═════════════════════════════════════════════════════════════
# STEP 3 — Upload Ericsson TSSR (optional, for images)
# ═════════════════════════════════════════════════════════════
if st.session_state.site_data:
    st.divider()
    st.subheader("3 · Upload Ericsson TSSR (optional)")
    st.caption("Upload to auto-extract images. Skip if you'll upload photos manually.")

    uploaded_pdf = st.file_uploader("Ericsson TSSR PDF", type=["pdf"])

    if uploaded_pdf and st.session_state.ericsson_docx is None:
        if not st.session_state.workdir:
            st.session_state.workdir = tempfile.mkdtemp(prefix="tssr_")

        workdir = Path(st.session_state.workdir)
        pdf_path = workdir / uploaded_pdf.name
        pdf_path.write_bytes(uploaded_pdf.getbuffer())

        with st.status("Converting PDF → DOCX…", expanded=True) as status:
            t0 = time.time()
            docx_path = core.pdf_to_docx(str(pdf_path), str(workdir))
            st.write(f"✅ Converted in {time.time() - t0:.1f}s")

            st.write("Extracting images…")
            imgs = core.extract_images(docx_path, str(workdir / "images"))
            st.session_state.ericsson_images = imgs
            st.session_state.ericsson_docx = docx_path
            st.write(f"✅ Found {len(imgs)} images")

            status.update(label="Done", state="complete")


# ═════════════════════════════════════════════════════════════
# STEP 4 — Image Upload Grid
# ═════════════════════════════════════════════════════════════
if st.session_state.site_data:
    st.divider()
    st.subheader("4 · Images")
    st.caption("Auto-extracted images are pre-selected. Override by uploading your own.")

    extracted = st.session_state.ericsson_images

    IMAGE_SLOTS = [
        ("img_vicinity_map",         "Vicinity Map"),
        ("img_site_photo_1",         "Site Photo 1"),
        ("img_site_photo_2",         "Site Photo 2"),
        ("img_site_photo_3",         "Site Photo 3"),
        ("img_site_photo_4",         "Site Photo 4"),
        ("img_olt_existing",         "OLT — Existing"),
        ("img_olt_proposed",         "OLT — Proposed"),
        ("RS1_load_sched_img",       "RS1 — Load Schedule"),
        ("RS1_load_calc_img",        "RS1 — Load Calc"),
        ("RS2_load_sched_img",       "RS2 — Load Schedule"),
        ("RS2_load_calc_img",        "RS2 — Load Calc"),
        ("RS1_tapping_img",          "RS1 — Tapping"),
        ("RS2_tapping_img",          "RS2 — Tapping"),
        ("equipment_room_layout_img","Equipment — Room Layout"),
        ("equipment_cable_routing_img","Equipment — Cable Routing"),
        ("transport_existing_1",     "Transport — Existing 1"),
        ("transport_existing_2",     "Transport — Existing 2"),
        ("transport_existing_3",     "Transport — Existing 3"),
    ]

    # Two-column grid for compact layout
    cols = st.columns(2)
    for i, (slot, label) in enumerate(IMAGE_SLOTS):
        with cols[i % 2]:
            st.markdown(f"**{label}**")

            # Option A — pick from extracted
            if extracted:
                options = ["(none)"] + extracted
                choice = st.selectbox(
                    "From TSSR",
                    options,
                    key=f"ext_{slot}",
                    label_visibility="collapsed",
                )
                if choice != "(none)":
                    st.session_state.image_map[slot] = choice
                elif st.session_state.image_map.get(slot) in extracted:
                    del st.session_state.image_map[slot]

            # Option B — upload override
            upload = st.file_uploader(
                "Or upload",
                type=["png", "jpg", "jpeg"],
                key=f"up_{slot}",
                label_visibility="collapsed",
            )
            if upload:
                path = save_upload(upload)
                st.session_state.image_map[slot] = path
                st.image(upload, use_container_width=True)


# ═════════════════════════════════════════════════════════════
# STEP 5 — Materials
# ═════════════════════════════════════════════════════════════
if st.session_state.site_data:
    st.divider()
    st.subheader("5 · Materials / Inventory")

    MATERIALS = [
        ("mat_grounding",  "Grounding (OLT MF-2 to Existing ground) #8", "m",   "10"),
        ("mat_patchcord",  "Patchcord (OLT to Transport)",              "pc",  "3"),
        ("mat_powercable", "Power Cable #14/2c (Blue/Black)",           "m",   "undetermined"),
        ("mat_lugs",       "Terminal lugs (#12 AWG)",                   "pcs", "4"),
        ("mat_tube",       "Shrinkable tube",                           "pcs", "6"),
        ("mat_tiewrap",    "Tie wrap x 8 inches",                       "pc",  "1"),
        ("mat_termlog",    "Terminal Log #14-8",                        "pcs", "6"),
        ("mat_conduit",    "Liquid-tight Flexible Steel Conduit",       "m",   "20"),
        ("mat_dcbreaker",  "DC breaker",                                "pcs", "2"),
    ]

    for key, label, default_unit, default_qty in MATERIALS:
        st.markdown(f"**{label}**")
        c1, c2 = st.columns([1, 2])
        with c1:
            st.session_state.materials[f"{key}_unit"] = st.text_input(
                "unit", value=default_unit, key=f"{key}_unit",
                label_visibility="collapsed", placeholder="unit",
            )
        with c2:
            st.session_state.materials[f"{key}_qty"] = st.text_input(
                "qty", value=default_qty, key=f"{key}_qty",
                label_visibility="collapsed", placeholder="qty",
            )


# ═════════════════════════════════════════════════════════════
# STEP 6 — Generate
# ═════════════════════════════════════════════════════════════
if st.session_state.site_data:
    st.divider()
    st.subheader("6 · Generate Nokia TSSR")

    template = Path("templates/nokia_template.docx")
    if not template.exists():
        st.error(f"Template not found: {template}")
    elif st.button("⚙️ Generate DOCX", type="primary", use_container_width=True):
        with st.spinner("Building context…"):
            # Merge image_map + materials into one lookup
            all_assets = {
                **st.session_state.image_map,
                **st.session_state.materials,
            }

            ctx = core.build_context(
                site=st.session_state.site_data,
                ericsson_fields=st.session_state.ericsson_fields,
                images=all_assets,
            )

        with st.spinner("Rendering template…"):
            out_dir = Path(st.session_state.workdir or tempfile.mkdtemp())
            out_name = core.make_output_name(
                ctx["site_id"], ctx["site_name"]
            )
            out_path = out_dir / out_name

            core.generate_nokia_tssr(
                template_path=str(template),
                output_path=str(out_path),
                context=ctx,
            )

        st.success("✅ Nokia TSSR generated")
        st.download_button(
            "⬇️ Download DOCX",
            data=Path(out_path).read_bytes(),
            file_name=out_name,
            mime=("application/vnd.openxmlformats-officedocument"
                  ".wordprocessingml.document"),
            use_container_width=True,
        )
