# app.py
"""
TSSR Automator v0.5
- Excel masterlist (auto-loaded)
- Ericsson TSSR PDF (image extraction)
- Persistent state across refresh
- Image grid with blue/black status indicators
- Real Ctrl+V paste via custom component
"""
import shutil
import tempfile
import time
from pathlib import Path

import streamlit as st

import core
import state_manager as sm
from excel_loader import SiteMasterlist
from components.paste_image import paste_image, save_pasted_image


# ═════════════════════════════════════════════════════════════
# Path resolution
# ═════════════════════════════════════════════════════════════

APP_DIR = Path(__file__).resolve().parent
MASTERLIST_PATH = APP_DIR / "data" / "MINDANAO_Site_Activity_Monitoring_OLT_PROJECT.xlsx"
TEMPLATE_PATH   = APP_DIR / "templates" / "nokia_template.docx"


# ═════════════════════════════════════════════════════════════
# Page config
# ═════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="TSSR Automator",
    page_icon="📡",
    layout="wide",
)


# ═════════════════════════════════════════════════════════════
# Session state — restore on startup
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
    "selected_plaid": "",
}

for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)

if "_restored" not in st.session_state:
    saved = sm.load_state()
    for k, v in saved.items():
        if k in DEFAULTS:
            st.session_state[k] = v

    saved_workdir = sm.load_workdir()
    if saved_workdir and Path(saved_workdir).exists():
        st.session_state.workdir = saved_workdir

    st.session_state["_restored"] = True


# ═════════════════════════════════════════════════════════════
# Helpers
# ═════════════════════════════════════════════════════════════

def persist():
    """Save current session state to disk."""
    sm.save_state(dict(st.session_state))
    if st.session_state.workdir:
        sm.save_workdir(st.session_state.workdir)


def reset_to_new_site():
    """Clear everything except masterlist."""
    if st.session_state.workdir and Path(st.session_state.workdir).exists():
        shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    sm.clear_state()

    for k, v in DEFAULTS.items():
        if k == "masterlist":
            continue
        if isinstance(v, (list, dict)):
            st.session_state[k] = type(v)()
        else:
            st.session_state[k] = v

    st.session_state["_restored"] = True


def ensure_workdir() -> Path:
    workdir = Path(st.session_state.workdir or tempfile.mkdtemp(prefix="tssr_"))
    st.session_state.workdir = str(workdir)
    sm.save_workdir(str(workdir))
    return workdir


def save_upload(uploaded_file) -> str:
    workdir = ensure_workdir()
    uploads = workdir / "uploads"
    uploads.mkdir(parents=True, exist_ok=True)
    dest = uploads / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return str(dest)


# ═════════════════════════════════════════════════════════════
# Image slot definitions
# ═════════════════════════════════════════════════════════════

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


# ═════════════════════════════════════════════════════════════
# Image editor dialog
# ═════════════════════════════════════════════════════════════

@st.dialog("Image slot", width="large")
def image_dialog(slot: str, label: str, extracted: list[str]):
    """Modal dialog for editing one image slot."""
    st.markdown(f"### {label}")
    st.caption(f"Slot key: `{slot}`")

    # Current image preview
    current = st.session_state.image_map.get(slot, "")
    if current and Path(current).exists():
        st.image(current, use_container_width=True)
        if st.button(
            "🗑 Remove current image",
            key=f"d_del_{slot}",
            use_container_width=True,
        ):
            st.session_state.image_map.pop(slot, None)
            persist()
            st.rerun()
        st.markdown("---")

    method = st.radio(
        "How to provide this image:",
        ["📁 Upload", "📋 Paste (Ctrl+V)", "🖼 Pick from TSSR"],
        key=f"d_method_{slot}",
        horizontal=True,
        label_visibility="collapsed",
    )

    # ─── Upload ───
    if method == "📁 Upload":
        upload = st.file_uploader(
            "Upload image",
            type=["png", "jpg", "jpeg"],
            key=f"d_up_{slot}",
            label_visibility="collapsed",
        )
        if upload:
            path = save_upload(upload)
            st.session_state.image_map[slot] = path
            persist()
            st.rerun()

    # ─── Paste (Ctrl+V) ───
    elif method == "📋 Paste (Ctrl+V)":
        st.caption(
            "**Click the dashed box below, then press Ctrl+V.** "
            "Or drag an image onto it."
        )
        result = paste_image(key=f"d_paste_{slot}")
        if result:
            cache_key = f"_pasted_{slot}_{result.get('size', 0)}"
            if cache_key not in st.session_state:
                workdir = ensure_workdir()
                ext = result["mime"].split("/")[-1].replace("jpeg", "jpg")
                dest = workdir / "uploads" / f"{slot}_{int(time.time())}.{ext}"
                save_pasted_image(result, str(dest))
                st.session_state.image_map[slot] = str(dest)
                st.session_state[cache_key] = True
                persist()
                st.rerun()

    # ─── Pick from TSSR extracts ───
    elif method == "🖼 Pick from TSSR":
        if not extracted:
            st.warning("No extracted images. Upload an Ericsson TSSR first.")
        else:
            st.caption(f"{len(extracted)} images — click to select")
            COLS = 4
            rows = (len(extracted) + COLS - 1) // COLS
            for r in range(rows):
                cols = st.columns(COLS)
                for c in range(COLS):
                    idx = r * COLS + c
                    if idx >= len(extracted):
                        break
                    img_path = extracted[idx]
                    is_selected = current == img_path
                    with cols[c]:
                        st.image(img_path, use_container_width=True)
                        if st.button(
                            "✓" if is_selected else "Select",
                            key=f"d_pick_{slot}_{idx}",
                            use_container_width=True,
                            type="primary" if is_selected else "secondary",
                        ):
                            st.session_state.image_map[slot] = img_path
                            persist()
                            st.rerun()


# ═════════════════════════════════════════════════════════════
# Header
# ═════════════════════════════════════════════════════════════

col_a, col_b, col_c = st.columns([3, 1, 1])
with col_a:
    st.title("📡 TSSR Automator")
    st.caption("Excel masterlist + Ericsson TSSR → Nokia TSSR DOCX")
with col_b:
    if st.button(
        "🆕 New Site",
        use_container_width=True,
        help="Clear current work and start over",
    ):
        reset_to_new_site()
        st.rerun()
with col_c:
    if st.button(
        "↺ Clear All",
        use_container_width=True,
        help="Reset everything including masterlist",
    ):
        reset_to_new_site()
        st.session_state.masterlist = None
        st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 1 — Masterlist
# ═════════════════════════════════════════════════════════════

st.subheader("1 · Site Masterlist")

if st.session_state.masterlist is None:
    if not MASTERLIST_PATH.exists():
        st.error(f"❌ Masterlist not found at `{MASTERLIST_PATH}`.")
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

plaids = st.session_state.masterlist.list_plaids()
if not plaids:
    st.warning("No PLAIDs found.")
    st.stop()

current_idx = 0
if st.session_state.selected_plaid in plaids:
    current_idx = plaids.index(st.session_state.selected_plaid) + 1

site_id = st.selectbox(
    "PLAID",
    options=[""] + plaids,
    index=current_idx,
    help="Start typing to filter",
)

if site_id and site_id != st.session_state.selected_plaid:
    st.session_state.selected_plaid = site_id
    st.session_state.site_data = (
        st.session_state.masterlist.get_site(site_id) or {}
    )
    persist()
    st.rerun()

if st.session_state.site_data:
    site = st.session_state.site_data
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


# ═════════════════════════════════════════════════════════════
# STEP 3 — Ericsson TSSR
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("3 · Ericsson TSSR (optional)")

    if st.session_state.ericsson_docx is None:
        uploaded_pdf = st.file_uploader(
            "Upload Ericsson TSSR PDF to extract images",
            type=["pdf"],
            key="ericsson_pdf_upload",
        )

        if uploaded_pdf:
            workdir = ensure_workdir()
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
                status.update(label="Done", state="complete", expanded=False)

            persist()
            st.rerun()
    else:
        st.success(
            f"✅ Ericsson TSSR loaded "
            f"({len(st.session_state.ericsson_images)} images)"
        )
        if st.button("🗑 Clear Ericsson TSSR", key="clear_ericsson"):
            st.session_state.ericsson_docx = None
            st.session_state.ericsson_images = []
            persist()
            st.rerun()


# ═════════════════════════════════════════════════════════════
# STEP 4 — Images (grid + dialog)
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("4 · Images")

    filled = sum(
        1 for slot, _ in IMAGE_SLOTS
        if st.session_state.image_map.get(slot)
        and Path(st.session_state.image_map[slot]).exists()
    )
    st.caption(
        f"**{filled} / {len(IMAGE_SLOTS)}** slots filled. "
        f"Click any row to upload, paste, or pick an image."
    )

    extracted = st.session_state.ericsson_images

    COLS = 3
    rows = (len(IMAGE_SLOTS) + COLS - 1) // COLS

    for r in range(rows):
        cols = st.columns(COLS)
        for c in range(COLS):
            idx = r * COLS + c
            if idx >= len(IMAGE_SLOTS):
                break
            slot, label = IMAGE_SLOTS[idx]
            current = st.session_state.image_map.get(slot, "")
            has_image = bool(current and Path(current).exists())

            with cols[c]:
                icon = "🔵" if has_image else "⚫"
                if st.button(
                    f"{icon}  {label}",
                    key=f"open_{slot}",
                    use_container_width=True,
                    type="primary" if has_image else "secondary",
                ):
                    image_dialog(slot, label, extracted)


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
            )
        with c2:
            qty_val = st.text_input(
                "qty",
                value=st.session_state.materials.get(f"{key}_qty", default_qty),
                key=f"input_{key}_qty",
                label_visibility="collapsed",
            )
        st.session_state.materials[f"{key}_unit"] = unit_val
        st.session_state.materials[f"{key}_qty"]  = qty_val

    if st.button("💾 Save materials", key="save_materials"):
        persist()
        st.success("Saved")


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
        workdir = ensure_workdir()

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
    st.markdown("Excel masterlist + Ericsson TSSR → Nokia TSSR DOCX")

    st.markdown("---")
    st.markdown("**Workflow**")
    st.markdown(
        "1. Masterlist auto-loads\n"
        "2. Select a PLAID\n"
        "3. Upload Ericsson TSSR\n"
        "4. Per image: upload, paste, or pick\n"
        "5. Fill materials\n"
        "6. Generate & download"
    )

    st.markdown("---")
    st.caption("v0.5 · Image grid + Ctrl+V paste")

    if st.session_state.get("selected_plaid"):
        st.success(f"Working on: **{st.session_state.selected_plaid}**")

    if st.button("💾 Save state now", use_container_width=True):
        persist()
        st.success("Saved")

    with st.expander("🔍 Session info"):
        st.write("**Workdir:**", st.session_state.workdir)
        st.write("**Selected PLAID:**", st.session_state.selected_plaid)
        st.write("**Images set:**", len(st.session_state.image_map))
        st.write("**Materials set:**", len(st.session_state.materials))
        st.write("**Extracted imgs:**", len(st.session_state.ericsson_images))
