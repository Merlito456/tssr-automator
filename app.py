# app.py
"""
TSSR Automator v0.2 — Dual Source
Excel masterlist + Ericsson TSSR + user-uploaded images.
"""
import shutil
import tempfile
import time
from pathlib import Path

import streamlit as st

import core
from excel_loader import SiteMasterlist


# ═════════════════════════════════════════════════════════════
# Configuration
# ═════════════════════════════════════════════════════════════

MASTERLIST_PATH = Path("data/MINDANAO_Site_Activity_Monitoring_OLT_PROJECT.xlsx")
TEMPLATE_PATH   = Path("templates/nokia_template.docx")


# ═════════════════════════════════════════════════════════════
# Page config
# ═════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="TSSR Automator",
    page_icon="📡",
    layout="wide",
)


# ═════════════════════════════════════════════════════════════
# Session state
# ═════════════════════════════════════════════════════════════

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
    """Clear session state and temp files."""
    if st.session_state.workdir and Path(st.session_state.workdir).exists():
        shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


def save_upload(uploaded_file) -> str:
    """Persist an uploaded file to a temp workdir; return its path."""
    workdir = Path(st.session_state.workdir or tempfile.mkdtemp(prefix="tssr_"))
    st.session_state.workdir = str(workdir)
    uploads = workdir / "uploads"
    uploads.mkdir(parents=True, exist_ok=True)
    dest = uploads / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return str(dest)


# ═════════════════════════════════════════════════════════════
# Header
# ═════════════════════════════════════════════════════════════

col_a, col_b = st.columns([4, 1])
with col_a:
    st.title("📡 TSSR Automator")
    st.caption("Excel masterlist + Ericsson TSSR → Nokia TSSR DOCX")
with col_b:
    if st.button("↺ Reset", use_container_width=True):
        reset()
        st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 1 — Load Masterlist (hardcoded path)
# ═════════════════════════════════════════════════════════════

st.subheader("1 · Site Masterlist")

if st.session_state.masterlist is None:
    if not MASTERLIST_PATH.exists():
        st.error(f"❌ Masterlist not found at `{MASTERLIST_PATH}`.")
        with st.expander("🔍 Debug info"):
            st.write("**Working directory:**", Path.cwd())
            st.write("**`data/` exists:**", Path("data").exists())
            if Path("data").exists():
                st.write(
                    "**`data/` contents:**",
                    [p.name for p in Path("data").iterdir()],
                )
            st.write("**Repo root contents:**",
                     [p.name for p in Path.cwd().iterdir()])
        st.stop()

    try:
        st.session_state.masterlist = SiteMasterlist(str(MASTERLIST_PATH))
        st.success(f"✅ Loaded: `{MASTERLIST_PATH.name}`")
    except Exception as e:
        st.error(f"❌ Failed to load masterlist: {e}")
        st.exception(e)
        st.stop()
else:
    st.success(f"✅ Loaded: `{MASTERLIST_PATH.name}`")


# ═════════════════════════════════════════════════════════════
# STEP 2 — Select Site
# ═════════════════════════════════════════════════════════════

st.divider()
st.subheader("2 · Select Site")

if st.session_state.masterlist:
    plaids = st.session_state.masterlist.list_plaids()

    if not plaids:
        st.warning("No PLAIDs found in the masterlist.")
        st.stop()

    site_id = st.selectbox(
        "PLAID",
        options=[""] + plaids,
        index=0,
        help="Start typing to filter",
    )

    if site_id:
        site = st.session_state.masterlist.get_site(site_id)
        if site:
            st.session_state.site_data = site

            st.markdown("**Site details:**")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Site ID",   site.get("site_id", ""),   disabled=True, key="disp_site_id")
                st.text_input("Site Name", site.get("site_name", ""), disabled=True, key="disp_site_name")
            with c2:
                st.text_input("Region",    site.get("region", ""),    disabled=True, key="disp_region")
                st.text_input("Towerco",   site.get("towerco", ""),   disabled=True, key="disp_towerco")
            with c3:
                st.text_input("FO Name",   site.get("fo_name", ""),   disabled=True, key="disp_fo_name")
                st.text_input("FO Mobile", site.get("fo_mobile", ""), disabled=True, key="disp_fo_mobile")

            with st.expander("📍 Full address & coordinates"):
                st.write("**Address:**",    site.get("site_address", ""))
                st.write("**Coordinates:**", site.get("site_coords", ""))
                st.write("**Key location:**", site.get("site_key_location", ""))
        else:
            st.warning(f"No data found for PLAID `{site_id}`.")
            st.session_state.site_data = {}


# ═════════════════════════════════════════════════════════════
# STEP 3 — Upload Ericsson TSSR (optional)
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("3 · Ericsson TSSR (optional)")
    st.caption("Upload to auto-extract images. Skip if uploading photos manually.")

    if st.session_state.ericsson_docx is None:
        uploaded_pdf = st.file_uploader(
            "Ericsson TSSR PDF",
            type=["pdf"],
            key="ericsson_pdf_upload",
        )

        if uploaded_pdf:
            workdir = Path(
                st.session_state.workdir
                or tempfile.mkdtemp(prefix="tssr_")
            )
            st.session_state.workdir = str(workdir)
            pdf_path = workdir / uploaded_pdf.name
            pdf_path.write_bytes(uploaded_pdf.getbuffer())

            with st.status("Converting PDF → DOCX…", expanded=True) as status:
                t0 = time.time()
                docx_path = core.pdf_to_docx(str(pdf_path), str(workdir))
                st.write(f"✅ Converted in {time.time() - t0:.1f}s")

                st.write("Extracting images…")
                imgs = core.extract_images(
                    docx_path, str(workdir / "images")
                )
                st.session_state.ericsson_images = imgs
                st.session_state.ericsson_docx = docx_path
                st.write(f"✅ Found {len(imgs)} images")

                status.update(label="Done", state="complete", expanded=False)
            st.rerun()
    else:
        st.success(f"✅ Ericsson TSSR loaded ({len(st.session_state.ericsson_images)} images)")

        if st.button("🗑 Clear Ericsson TSSR", key="clear_ericsson"):
            st.session_state.ericsson_docx = None
            st.session_state.ericsson_images = []
            st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 4 — Images
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("4 · Images")
    st.caption(
        "Pick from auto-extracted images, or upload your own. "
        "Leave blank to skip a slot."
    )

    extracted = st.session_state.ericsson_images

    IMAGE_SLOTS = [
        ("img_vicinity_map",           "Vicinity Map"),
        ("img_site_photo_1",           "Site Photo 1"),
        ("img_site_photo_2",           "Site Photo 2"),
        ("img_site_photo_3",           "Site Photo 3"),
        ("img_site_photo_4",           "Site Photo 4"),
        ("img_olt_existing",           "OLT — Existing"),
        ("img_olt_proposed",           "OLT — Proposed"),
        ("RS1_load_sched_img",         "RS1 — Load Schedule"),
        ("RS1_load_calc_img",          "RS1 — Load Calc"),
        ("RS2_load_sched_img",         "RS2 — Load Schedule"),
        ("RS2_load_calc_img",          "RS2 — Load Calc"),
        ("RS1_tapping_img",            "RS1 — Tapping"),
        ("RS2_tapping_img",            "RS2 — Tapping"),
        ("equipment_room_layout_img",  "Equipment — Room Layout"),
        ("equipment_cable_routing_img","Equipment — Cable Routing"),
        ("transport_existing_1",       "Transport — Existing 1"),
        ("transport_existing_2",       "Transport — Existing 2"),
        ("transport_existing_3",       "Transport — Existing 3"),
    ]

    cols = st.columns(2)

    for i, (slot, label) in enumerate(IMAGE_SLOTS):
        with cols[i % 2]:
            st.markdown(f"**{label}**")

            # Option 1 — pick from extracted
            if extracted:
                options = ["(none)"] + extracted
                current = st.session_state.image_map.get(slot, "(none)")
                idx = options.index(current) if current in options else 0

                choice = st.selectbox(
                    "From TSSR",
                    options,
                    index=idx,
                    key=f"sel_{slot}",
                    label_visibility="collapsed",
                )

                if choice != "(none)":
                    st.session_state.image_map[slot] = choice
                elif st.session_state.image_map.get(slot) in extracted:
                    del st.session_state.image_map[slot]

            # Option 2 — upload override
            upload = st.file_uploader(
                "Or upload",
                type=["png", "jpg", "jpeg"],
                key=f"up_{slot}",
                label_visibility="collapsed",
            )
            if upload:
                path = save_upload(upload)
                st.session_state.image_map[slot] = path

            # Preview if set
            current_path = st.session_state.image_map.get(slot)
            if current_path and Path(current_path).exists():
                st.image(current_path, use_container_width=True)


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
            unit_val = st.text_input(
                "unit",
                value=st.session_state.materials.get(f"{key}_unit", default_unit),
                key=f"input_{key}_unit",
                label_visibility="collapsed",
                placeholder="unit",
            )
        with c2:
            qty_val = st.text_input(
                "qty",
                value=st.session_state.materials.get(f"{key}_qty", default_qty),
                key=f"input_{key}_qty",
                label_visibility="collapsed",
                placeholder="length / pcs",
            )
        st.session_state.materials[f"{key}_unit"] = unit_val
        st.session_state.materials[f"{key}_qty"]  = qty_val


# ═════════════════════════════════════════════════════════════
# STEP 6 — Generate
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("6 · Generate Nokia TSSR")

    if not TEMPLATE_PATH.exists():
        st.error(f"❌ Template not found: `{TEMPLATE_PATH}`")
        st.stop()

    if st.button("⚙️ Generate DOCX", type="primary", use_container_width=True):
        # Ensure we have a workdir
        if not st.session_state.workdir:
            st.session_state.workdir = tempfile.mkdtemp(prefix="tssr_")
        workdir = Path(st.session_state.workdir)

        # Merge image_map + materials for build_context
        all_assets = {
            **st.session_state.image_map,
            **st.session_state.materials,
        }

        with st.spinner("Building context…"):
            ctx = core.build_context(
                site=st.session_state.site_data,
                ericsson_fields=st.session_state.ericsson_fields,
                images=all_assets,
            )

        with st.spinner("Rendering template…"):
            out_dir = workdir / "output"
            out_dir.mkdir(parents=True, exist_ok=True)

            out_name = core.make_output_name(
                ctx.get("site_id", "SITE"),
                ctx.get("site_name", "TSSR"),
            )
            out_path = out_dir / out_name

            try:
                core.generate_nokia_tssr(
                    template_path=str(TEMPLATE_PATH),
                    output_path=str(out_path),
                    context=ctx,
                )
            except Exception as e:
                st.error(f"❌ Generation failed: {e}")
                st.exception(e)
                st.stop()

        st.success("✅ Nokia TSSR generated")
        st.download_button(
            "⬇️ Download DOCX",
            data=Path(out_path).read_bytes(),
            file_name=out_name,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
            use_container_width=True,
        )


# ═════════════════════════════════════════════════════════════
# Sidebar
# ═════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### About")
    st.markdown(
        "Converts **Excel masterlist** + **Ericsson TSSR** into a "
        "populated **Nokia TSSR** DOCX."
    )
    st.markdown("---")
    st.markdown("**Workflow**")
    st.markdown(
        "1. Masterlist auto-loads\n"
        "2. Select a PLAID\n"
        "3. (Optional) Upload Ericsson PDF\n"
        "4. Pick or upload images\n"
        "5. Fill materials\n"
        "6. Generate & download"
    )
    st.markdown("---")
    st.caption("v0.2 · Python + Streamlit")

    if st.session_state.masterlist:
        with st.expander("🔍 Masterlist info"):
            info = st.session_state.masterlist.debug_info()
            st.write(f"**Sheet:** {info['sheet']}")
            st.write(f"**Rows:** {info['rows']}")
            st.write(f"**Columns:** {len(info['columns_raw'])}")
            if info["missing_fields"]:
                st.warning(
                    f"Missing fields: {info['missing_fields']}"
                )
