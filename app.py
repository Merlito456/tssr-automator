# app.py
"""
Compact TSSR Automator.
Upload Ericsson PDF → download Nokia TSSR DOCX.
"""
import shutil
import tempfile
import time
from pathlib import Path

import streamlit as st

import core

# ─────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TSSR Automator",
    page_icon="📡",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────
# Session state defaults
# ─────────────────────────────────────────────────────────────
DEFAULTS = {
    "workdir": None,
    "ericsson_docx": None,
    "images": [],
    "fields": {},
    "image_map": {},
}
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)


def reset():
    if st.session_state.workdir and Path(st.session_state.workdir).exists():
        shutil.rmtree(st.session_state.workdir, ignore_errors=True)
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


# ─────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────
col_a, col_b = st.columns([4, 1])
with col_a:
    st.title("📡 TSSR Automator")
    st.caption("Ericsson TSSR PDF → Nokia TSSR Word")
with col_b:
    if st.button("↺ Reset", use_container_width=True):
        reset()
        st.rerun()

# ─────────────────────────────────────────────────────────────
# Step 1 — Upload
# ─────────────────────────────────────────────────────────────
st.subheader("1 · Upload Ericsson TSSR")

uploaded = st.file_uploader(
    "Drop the Ericsson TSSR PDF",
    type=["pdf"],
    label_visibility="collapsed",
)

if uploaded and st.session_state.ericsson_docx is None:
    workdir = tempfile.mkdtemp(prefix="tssr_")
    st.session_state.workdir = workdir
    pdf_path = Path(workdir) / uploaded.name
    pdf_path.write_bytes(uploaded.getbuffer())

    with st.status("Converting PDF → DOCX…", expanded=True) as status:
        t0 = time.time()
        docx_path = core.pdf_to_docx(str(pdf_path), workdir)
        st.write(f"✅ Converted in {time.time() - t0:.1f}s")

        st.write("Extracting fields…")
        fields = core.extract_fields(docx_path)
        st.session_state.fields = fields

        st.write("Extracting images…")
        images = core.extract_images(docx_path, str(Path(workdir) / "images"))
        st.session_state.images = images
        st.write(f"✅ Found {len(images)} images")

        st.session_state.ericsson_docx = docx_path
        status.update(label="Done", state="complete", expanded=False)
    st.rerun()

# ─────────────────────────────────────────────────────────────
# Step 2 — Review
# ─────────────────────────────────────────────────────────────
if st.session_state.fields:
    st.divider()
    st.subheader("2 · Review Extracted Fields")
    st.caption("Edit anything the parser got wrong before generating.")

    fields = st.session_state.fields
    c1, c2, c3 = st.columns(3)
    with c1:
        fields["site_id"]   = st.text_input("Site ID",   fields.get("site_id", ""))
        fields["site_name"] = st.text_input("Site Name", fields.get("site_name", ""))
        fields["region"]    = st.text_input("Region",    fields.get("region", ""))
    with c2:
        fields["latitude"]  = st.text_input("Latitude",  fields.get("latitude", ""))
        fields["longitude"] = st.text_input("Longitude", fields.get("longitude", ""))
        fields["owner"]     = st.text_input("Owner",     fields.get("owner", ""))
    with c3:
        fields["contact_name"] = st.text_input("Contact",    fields.get("contact_name", ""))
        fields["contact_no"]   = st.text_input("Contact No", fields.get("contact_no", ""))
        fields["address"]      = st.text_input("Address",    fields.get("address", ""))

    st.session_state.fields = fields

    # ─── Image mapping ───────────────────────────────────────
    st.divider()
    st.subheader("3 · Map Images")
    st.caption("Pick which extracted image goes into each Nokia template slot.")

    SLOTS = ["vicinity_map", "site_photo", "proposed_space"]
    images = st.session_state.images

    if not images:
        st.info("No images found in the source document.")
    else:
        options = ["(none)"] + images
        img_cols = st.columns(len(SLOTS))
        for col, slot in zip(img_cols, SLOTS):
            with col:
                current = st.session_state.image_map.get(slot, "(none)")
                idx = options.index(current) if current in options else 0
                choice = st.selectbox(
                    slot.replace("_", " ").title(),
                    options,
                    index=idx,
                    key=f"sel_{slot}",
                )
                st.session_state.image_map[slot] = (
                    "" if choice == "(none)" else choice
                )
                if choice != "(none)":
                    st.image(choice, use_container_width=True)

    # ─── Generate ────────────────────────────────────────────
    st.divider()
    st.subheader("4 · Generate Nokia TSSR")

    template = Path("templates/nokia_template.docx")
    if not template.exists():
        st.error(f"Template not found: {template}")
    else:
        if st.button("⚙️ Generate DOCX", type="primary", use_container_width=True):
            out_dir = Path(st.session_state.workdir) / "output"
            out_dir.mkdir(exist_ok=True)
            out_name = core.make_output_name(st.session_state.fields, "docx")
            out_path = out_dir / out_name

            with st.spinner("Rendering template…"):
                core.generate_nokia_tssr(
                    template_path=str(template),
                    output_path=str(out_path),
                    fields=st.session_state.fields,
                    image_map=st.session_state.image_map,
                )

            st.success("✅ Nokia TSSR generated")
            st.download_button(
                "⬇️ Download Nokia TSSR",
                data=out_path.read_bytes(),
                file_name=out_name,
                mime=("application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document"),
                use_container_width=True,
            )

# ─────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### About")
    st.markdown(
        "Converts **Ericsson TSSR** PDFs into populated "
        "**Nokia TSSR** Word documents for the GLOBE HERMES project."
    )
    st.markdown("---")
    st.markdown("**Workflow**")
    st.markdown("1. Upload PDF\n2. Review fields\n3. Map images\n4. Generate")
    st.markdown("---")
    st.caption("v0.1 · Python + Streamlit")
