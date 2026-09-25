# app.py
"""
TSSR Automator v0.9
- Excel masterlist (auto-loaded)
- AI-assisted field extraction via Gemini/ChatGPT
- AI-assisted load calculation (C7 block, RS1–RS4)
- Ericsson TSSR PDF (image extraction)
- Hardcoded work_permit + access_requirement from towerco rules
- Map + Street View capture (no API key)
- Persistent state across refresh
- Image grid with blue/black status indicators
- Real Ctrl+V paste via custom component
"""
import shutil
import tempfile
import time
import traceback
import urllib.request
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import core
import ai_helper
import state_manager as sm
from excel_loader import SiteMasterlist
from components.paste_image import paste_image, save_pasted_image
from permit_rules import get_permits_for_towerco


APP_DIR = Path(__file__).resolve().parent
MASTERLIST_PATH = APP_DIR / "data" / "MINDANAO_Site_Activity_Monitoring_OLT_PROJECT.xlsx"
TEMPLATE_PATH   = APP_DIR / "templates" / "nokia_template.docx"


st.set_page_config(
    page_title="TSSR Automator",
    page_icon="📡",
    layout="wide",
)


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
    "ai_applied": False,
    # --- AI Load Calculator ---
    "load_calc_data": None,
    "load_calc_rs": "RS1 — Rectifier 1",
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
    sm.save_state(dict(st.session_state))
    if st.session_state.workdir:
        sm.save_workdir(st.session_state.workdir)


def reset_to_new_site():
    if st.session_state.workdir and Path(st.session_state.workdir).exists():
        shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    sm.clear_state()
    for k, v in DEFAULTS.items():
        if k == "masterlist":
            continue
        if isinstance(v, (list, dict)):
            st.session_state[k] = type(v)()
        elif isinstance(v, bool):
            st.session_state[k] = False
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
# Map helpers — no API key needed
# ═════════════════════════════════════════════════════════════

def _osm_embed_html(lat: str, lon: str, height: int = 450,
                    zoom: int = 17) -> str:
    """
    OpenStreetMap embed via iframe. Free, no API key.
    Uses a bbox around the point for the visible extent.
    """
    lat_f = float(lat)
    lon_f = float(lon)
    delta = 0.003
    bbox = f"{lon_f - delta},{lat_f - delta},{lon_f + delta},{lat_f + delta}"
    return f"""
    <iframe
        width="100%"
        height="{height}"
        frameborder="0"
        scrolling="no"
        marginheight="0"
        marginwidth="0"
        src="https://www.openstreetmap.org/export/embed.html?bbox={bbox}&layer=mapnik&marker={lat_f},{lon_f}"
        style="border: 1px solid #ccc; border-radius: 8px;">
    </iframe>
    """


def _download_static_map(lat: str, lon: str, dest_path: str,
                         zoom: int = 17, size: str = "800x600") -> str | None:
    """Download a static map PNG from OSM-based staticmap service."""
    try:
        url = (
            f"https://staticmap.openstreetmap.de/staticmap.php"
            f"?center={lat},{lon}"
            f"&zoom={zoom}"
            f"&size={size}"
            f"&maptype=mapnik"
            f"&markers={lat},{lon},red-pushpin"
        )
        req = urllib.request.Request(
            url, headers={"User-Agent": "TSSR-Automator/1.0"}
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read()

        if len(data) < 2000:
            print(f"Static map response too small: {len(data)} bytes")
            return None

        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data)
        return str(dest)
    except Exception as e:
        print(f"Static map download failed: {e}")
        return None


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
    st.markdown(f"### {label}")
    st.caption(f"Slot key: `{slot}`")
    current = st.session_state.image_map.get(slot, "")
    if current and Path(current).exists():
        st.image(current, use_container_width=True)
        if st.button("🗑 Remove current image",
                     key=f"d_del_{slot}", use_container_width=True):
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

    if method == "📁 Upload":
        upload = st.file_uploader(
            "Upload image", type=["png", "jpg", "jpeg"],
            key=f"d_up_{slot}", label_visibility="collapsed",
        )
        if upload:
            path = save_upload(upload)
            st.session_state.image_map[slot] = path
            persist()
            st.rerun()

    elif method == "📋 Paste (Ctrl+V)":
        st.caption("**Click the dashed box below, then press Ctrl+V.**")
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
    st.caption("Excel masterlist + AI + Ericsson TSSR → Nokia TSSR DOCX")
with col_b:
    if st.button("🆕 New Site", use_container_width=True):
        reset_to_new_site()
        st.rerun()
with col_c:
    if st.button("↺ Clear All", use_container_width=True):
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
    st.session_state.ai_applied = False
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
# STEP 3 — AI-Assisted Fields
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("3 · AI-Assisted Fields (optional)")
    st.caption(
        "Copy the prompt below, feed it to **Gemini** or **ChatGPT** "
        "along with the Ericsson TSSR PDF, then paste the JSON response back."
    )

    with st.expander("📋 Step 1 — Copy this prompt", expanded=False):
        prompt_text = ai_helper.build_prompt()
        st.code(prompt_text, language="text")

    with st.expander("📥 Step 2 — Paste AI response",
                     expanded=not st.session_state.get("ai_applied", False)):
        ai_response = st.text_area(
            "Paste the AI's JSON response here:",
            height=200,
            key="ai_json_input",
            placeholder='{"site_class": "C3", ...}',
        )

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("✔ Apply AI fields", key="apply_ai",
                         type="primary", use_container_width=True):
                if not ai_response.strip():
                    st.warning("Paste a JSON response first.")
                else:
                    parsed = ai_helper.extract_json_from_response(ai_response)
                    if not parsed:
                        st.error("❌ Could not parse JSON from the response.")
                    else:
                        clean, warnings = ai_helper.validate_and_normalize(parsed)
                        merged = core.merge_ai_fields(
                            st.session_state.site_data, clean
                        )
                        st.session_state.site_data = merged
                        st.session_state["ai_applied"] = True
                        persist()

                        if warnings:
                            st.warning(
                                f"⚠️ Applied with {len(warnings)} warnings:\n\n"
                                + "\n".join(f"- {w}" for w in warnings)
                            )
                        st.success("✅ AI fields applied")
                        st.rerun()

        with col_b:
            if st.button("🗑 Clear response", key="clear_ai",
                         use_container_width=True):
                st.session_state.pop("ai_json_input", None)
                st.rerun()

    if st.session_state.get("ai_applied"):
        st.success("🤖 AI fields applied")

    with st.expander("🔍 Preview merged site data", expanded=False):
        site = st.session_state.site_data
        towerco_value = site.get("towerco", "")
        permits = get_permits_for_towerco(towerco_value)

        display_fields = [
            ("Site Class",        site.get("site_class", "")),
            ("Room Access",       site.get("room_access", "")),
            ("Cabin Location",    site.get("cabin_location", "")),
            ("Flood History",     site.get("flood_history", "")),
            ("Hauling Remarks",   site.get("hauling_remarks", "")),
            ("Site Profile",      site.get("site_profile", "")),
            ("Site Key Location", site.get("site_key_location", "")),
            ("Site Owner",        site.get("site_owner", "")),
            ("Site Security",     site.get("site_security", "")),
            ("Site Type",         site.get("site_type", "")),
            ("Site Accessible",   site.get("site_accessible", "")),
            ("No. of Bridge",     site.get("no_bridge", "")),
            ("Foot Trail",        site.get("foot_trail", "")),
            ("— DERIVED FROM TOWERCO —", ""),
            ("Work Permit",        permits["work_permit"]),
            ("Access Requirement", permits["access_requirement"]),
        ]

        for label, val in display_fields:
            if val not in (None, "", [], {}):
                st.write(f"**{label}:** `{val}`")


# ═════════════════════════════════════════════════════════════
# STEP 3.2 — AI Load Calculator (C7 block)
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()

    # ── DIAGNOSTIC: prove which load_calc_page.py is actually imported ──
    try:
        import importlib
        import load_calc_page
        importlib.reload(load_calc_page)

        st.caption(
            f"🔧 `load_calc_page` loaded from: "
            f"`{getattr(load_calc_page, '__file__', '?')}`"
        )
        public = [x for x in dir(load_calc_page) if not x.startswith("_")]
        st.caption(f"🔧 public names: `{public}`")

        if not hasattr(load_calc_page, "render"):
            st.error(
                "❌ `load_calc_page` has no `render` function. "
                "The file on the server is stale or wrong."
            )
            st.stop()

        load_calc_page.render(ensure_workdir, persist)

    except ImportError as e:
        st.warning(
            f"⚠️ Load calculator module not available: {e}\n\n"
            "Create `load_calc_page.py`, `load_calc_helper.py`, "
            "and `load_calc_render.py`, then place "
            "`data/load_calculation.xlsx`."
        )
    except Exception as e:
        st.error(f"❌ Load calculator crashed: {type(e).__name__}: {e}")
        with st.expander("🔎 Full traceback", expanded=True):
            st.code(traceback.format_exc())


# ═════════════════════════════════════════════════════════════
# STEP 3.5 — Map & Street View
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("3.5 · Site Map & Street View")
    st.caption(
        "Explore the site location using its coordinates. "
        "Capture the map as your vicinity map, or open Street View "
        "in a new tab and paste a screenshot."
    )

    site = st.session_state.site_data
    lat = str(site.get("latitude", "")).strip()
    lon = str(site.get("longitude", "")).strip()

    if not lat or not lon:
        st.warning("⚠️ No coordinates found in the masterlist for this site.")
    else:
        st.info(f"📍 **Coordinates:** `{lat}, {lon}`")

        tab_map, tab_street, tab_help = st.tabs([
            "🗺️ Open Street Map",
            "📸 Google Street View",
            "💡 How to use",
        ])

        # ─── Tab 1: OSM Map ───
        with tab_map:
            st.markdown("**Vicinity map preview**")
            st.caption(
                "Zoom in/out and pan. When you're happy with the view, "
                "click 'Use as Vicinity Map' below to save a static snapshot."
            )

            components.html(
                _osm_embed_html(lat, lon, height=450),
                height=470,
            )

            col_a, col_b = st.columns([1, 1])
            with col_a:
                if st.button(
                    "🖼 Use as Vicinity Map",
                    key="use_osm_as_vicinity",
                    use_container_width=True,
                    type="primary",
                ):
                    workdir = ensure_workdir()
                    dest = workdir / "uploads" / "vicinity_map_from_osm.png"

                    with st.spinner("Downloading static map…"):
                        result = _download_static_map(lat, lon, str(dest))

                    if result:
                        st.session_state.image_map["img_vicinity_map"] = result
                        persist()
                        st.success("✅ Vicinity map captured")
                        st.rerun()
                    else:
                        st.error(
                            "❌ Static map service unavailable. "
                            "Use the 'How to use' tab for manual capture."
                        )

            with col_b:
                st.markdown(
                    f"[🔗 Open larger map ↗]"
                    f"(https://www.openstreetmap.org/?mlat={lat}&mlon={lon}"
                    f"#map=17/{lat}/{lon})"
                )

        # ─── Tab 2: Google Street View ───
        with tab_street:
            st.markdown("**Street View preview**")
            st.caption(
                "Street View can't be embedded without a Google API key. "
                "Click the button below to open it in a new tab, then "
                "screenshot and paste it back into the app."
            )

            gmaps_embed_url = (
                f"https://maps.google.com/maps?q={lat},{lon}"
                f"&t=k&z=17&output=embed"
            )
            components.html(
                f"""
                <iframe
                    width="100%"
                    height="450"
                    frameborder="0"
                    style="border: 1px solid #ccc; border-radius: 8px;"
                    src="{gmaps_embed_url}"
                    allowfullscreen>
                </iframe>
                """,
                height=470,
            )

            st.markdown("### Open in Google Maps")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(
                    f"[🗺️ Satellite view ↗]"
                    f"(https://www.google.com/maps/@{lat},{lon},17z/data=!3m1!1e3)"
                )
            with c2:
                st.markdown(
                    f"[📸 Street View ↗]"
                    f"(https://www.google.com/maps/@?api=1&map_action=pano"
                    f"&viewpoint={lat},{lon})"
                )

        # ─── Tab 3: Instructions ───
        with tab_help:
            st.markdown("""
            ### How to capture maps and photos

            **Option 1 — Auto-capture vicinity map (fastest)**
            1. Go to the **🗺️ Open Street Map** tab
            2. Click **"🖼 Use as Vicinity Map"**
            3. The app downloads a static map PNG and saves it to the
               vicinity map slot. Done!

            **Option 2 — Manual screenshot (for any map)**
            1. Open the **🗺️ OSM** or **📸 Street View** tab
            2. Click the "Open larger map" or "Open Street View" link
            3. The map opens in a new browser tab
            4. Take a screenshot:
               - **Windows:** `Win + Shift + S`
               - **macOS:** `Cmd + Shift + 4`
               - **Chrome:** `Ctrl + Shift + P` (DevTools screenshot)
            5. Come back to this app
            6. Scroll to **Section 5 · Images**
            7. Click any slot (e.g. Vicinity Map or Site Photo 1)
            8. Choose **📋 Paste (Ctrl+V)**
            9. Click the dashed box → press `Ctrl+V` → done

            **Option 3 — Satellite view**
            For a satellite view instead of the street map, use the
            **🗺️ Satellite view** link in the Street View tab.

            ---

            **Tips:**
            - Street View coverage varies by region. Some remote sites
              may not have Street View — use the satellite view instead.
            - The static OSM map works best for rural areas.
            - Google Maps often has higher-resolution imagery for urban sites.
            """)

        # ─── Quick actions ───
        st.markdown("---")
        st.markdown("**Quick actions:**")
        qc1, qc2, qc3 = st.columns(3)

        with qc1:
            if st.button(
                "🖼 Capture as Vicinity Map",
                key="quick_vicinity",
                use_container_width=True,
            ):
                workdir = ensure_workdir()
                dest = workdir / "uploads" / "vicinity_map.png"
                with st.spinner("Downloading…"):
                    result = _download_static_map(lat, lon, str(dest))
                if result:
                    st.session_state.image_map["img_vicinity_map"] = result
                    persist()
                    st.success("✅ Saved as Vicinity Map")
                    st.rerun()
                else:
                    st.error("❌ Map service unavailable — try manual capture")

        with qc2:
            if st.button(
                "📸 Capture as Site Photo 1",
                key="quick_site_photo",
                use_container_width=True,
            ):
                workdir = ensure_workdir()
                dest = workdir / "uploads" / "site_photo_from_map.png"
                with st.spinner("Downloading…"):
                    result = _download_static_map(lat, lon, str(dest), zoom=18)
                if result:
                    st.session_state.image_map["img_site_photo_1"] = result
                    persist()
                    st.success("✅ Saved as Site Photo 1")
                    st.rerun()
                else:
                    st.error("❌ Map service unavailable")

        with qc3:
            if st.button(
                "🌐 Open in new tab",
                key="open_new_tab",
                use_container_width=True,
            ):
                st.markdown(
                    f"**→ [Click to open Google Maps](https://www.google.com/maps/@{lat},{lon},17z)**",
                    unsafe_allow_html=True,
                )
                st.caption(
                    "Right-click → **Open in new tab**, then screenshot."
                )


# ═════════════════════════════════════════════════════════════
# STEP 4 — Ericsson TSSR Images
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("4 · Ericsson TSSR Images")
    st.caption("Upload the Ericsson TSSR PDF to auto-extract images.")

    if st.session_state.ericsson_docx is None:
        uploaded_pdf = st.file_uploader(
            "Upload Ericsson TSSR PDF",
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
# STEP 5 — Images
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("5 · Images")

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
# STEP 6 — Materials
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("6 · Materials / Inventory")

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
# STEP 7 — Generate
# ═════════════════════════════════════════════════════════════

if st.session_state.site_data:
    st.divider()
    st.subheader("7 · Generate Nokia TSSR")

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
            mime=("application/vnd.openxmlformats-officedocument"
                  ".wordprocessingml.document"),
            use_container_width=True,
        )


# ═════════════════════════════════════════════════════════════# Sidebar
# ═════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### About")
    st.markdown("Excel masterlist + AI + Ericsson TSSR → Nokia TSSR DOCX")
    st.markdown("---")
    st.markdown("**Workflow**")
    st.markdown(
        "1. Masterlist auto-loads\n"
        "2. Select a PLAID\n"
        "3. (Optional) AI-assisted fields\n"
        "4. (Optional) AI load calculator (C7)\n"
        "5. Capture map + street view\n"
        "6. Upload Ericsson TSSR for images\n"
        "7. Fill image slots\n"
        "8. Fill materials\n"
        "9. Generate & download"
    )
    st.markdown("---")
    st.caption("v0.9 · AI load calculator")

    if st.session_state.get("selected_plaid"):
        st.success(f"Working on: **{st.session_state.selected_plaid}**")

    if st.button("💾 Save state now", use_container_width=True):
        persist()
        st.success("Saved")

    with st.expander("🔍 Session info"):
        st.write("**Workdir:**", st.session_state.workdir)
        st.write("**Selected PLAID:**", st.session_state.selected_plaid)
        st.write("**AI applied:**", st.session_state.get("ai_applied", False))
        st.write("**Load calc RS:**", st.session_state.get("load_calc_rs", "—"))
        st.write("**Load calc data:**", "✓" if st.session_state.get("load_calc_data") else "—")
        st.write("**Images set:**", len(st.session_state.image_map))
        st.write("**Materials set:**", len(st.session_state.materials))
        st.write("**Extracted imgs:**", len(st.session_state.ericsson_images))
