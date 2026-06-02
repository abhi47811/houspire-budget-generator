import os, tempfile, re
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

from config import CITIES, CITIES_WITH_MULTIPLIERS, TIERS, ROOM_TYPES
from modules.room_analyzer import analyze_all_renders, RoomAnalysis
from modules.boq_generator import generate_boq
from modules.vendor_finder import generate_vendors
from modules.budget_excel import write_budget_excel
from modules.vendor_excel import write_vendor_excel
from modules.docx_generator import generate_branded_boq_docx, generate_branded_vendor_docx
from modules.pdf_converter import docx_to_pdf
from modules.db import save_project, list_projects, load_project, delete_project, _available as db_available

st.set_page_config(page_title="Houspire Budget Generator", page_icon="🏠", layout="wide")
st.title("🏠 Houspire Budget Generator")
st.caption("Internal tool — generates BOQ + Vendor Excel from client renders")

# ── Sidebar: Project History ──────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Past Projects")
    if db_available():
        projects = list_projects()
        if not projects:
            st.caption("No saved projects yet.")
        else:
            for p in projects:
                created = p["created_at"][:10]
                label = f"{p['client_name']} · {p['city']} · {created}"
                col_a, col_b = st.columns([4, 1])
                with col_a:
                    if st.button(label, key=f"load_{p['id']}"):
                        data = load_project(p["id"])
                        if data:
                            proj = data["project"]
                            db_rows = data["boq_rows"]
                            db_sources = data["rate_sources"]
                            db_vendors = data["vendors"]
                            db_notes = data["notes"]
                            # Rebuild Excel bytes from DB rows
                            if db_rows:
                                from modules.boq_generator import BOQRow, RateSource
                                from modules.vendor_finder import VendorRow
                                rows = [BOQRow(r["category"], r["description"], r["unit"], r["qty"], r["rate"]) for r in db_rows]
                                sources = [RateSource(s["item"], s["basis"], s["source"]) for s in db_sources]
                                safe = re.sub(r"[^A-Za-z0-9_]", "", proj["client_name"].replace(" ", "_"))
                                with tempfile.TemporaryDirectory() as tmpdir:
                                    xlsx_name = f"{safe}_BOQ_{proj['city']}.xlsx"
                                    xlsx_path = os.path.join(tmpdir, xlsx_name)
                                    write_budget_excel(proj["client_name"], proj["city"], proj["pincode"], proj["tier"], rows, sources, xlsx_path)
                                    with open(xlsx_path, "rb") as fh:
                                        st.session_state.budget_excel_bytes = fh.read()
                                    st.session_state.budget_excel_name = xlsx_name
                            if db_vendors:
                                from modules.vendor_finder import VendorRow
                                vendors = [VendorRow(v["category"], v["vendor"], v["specialty"] or "", v["area"] or "", v["lat"] or 0.0, v["lng"] or 0.0, v["rating"] or "", v["phone"] or "") for v in db_vendors]
                                safe = re.sub(r"[^A-Za-z0-9_]", "", proj["client_name"].replace(" ", "_"))
                                centroid = (0.0, 0.0)
                                with tempfile.TemporaryDirectory() as tmpdir:
                                    xlsx_name = f"{proj['city']}_{safe}_Vendors_{proj['pincode']}.xlsx"
                                    xlsx_path = os.path.join(tmpdir, xlsx_name)
                                    write_vendor_excel(proj["client_name"], proj["city"], proj["pincode"], centroid, vendors, db_notes or "", xlsx_path)
                                    with open(xlsx_path, "rb") as fh:
                                        st.session_state.vendor_excel_bytes = fh.read()
                                    st.session_state.vendor_excel_name = xlsx_name
                            st.success(f"Loaded {proj['client_name']} — downloads ready")
                with col_b:
                    if st.button("🗑", key=f"del_{p['id']}"):
                        delete_project(p["id"])
                        st.rerun()
    else:
        st.caption("DB not configured — set SUPABASE_URL and SUPABASE_ANON_KEY.")

# ── Step 1: Client Details ────────────────────────────────────────────────────
st.header("Step 1 — Client Details")
c1, c2, c3, c4 = st.columns(4)
with c1:
    client_name = st.text_input("Client Name", placeholder="e.g. Sharma Residence")
with c2:
    city = st.selectbox("City", CITIES)
with c3:
    pincode = st.text_input("Pincode", placeholder="e.g. 500032")
with c4:
    tier = st.selectbox("Tier", TIERS)

multiplier = CITIES_WITH_MULTIPLIERS.get(city)
if multiplier:
    st.caption(f"City multiplier: ×{multiplier:.2f} (applied silently to all rates)")
elif city == "Other":
    st.caption("City multiplier: unknown — enter manual estimate if required")

# ── Step 2: Room Renders ──────────────────────────────────────────────────────
st.header("Step 2 — Upload Room Renders")
render_files = st.file_uploader(
    "Upload renders (JPG/PNG)", type=["jpg", "jpeg", "png"],
    accept_multiple_files=True, key="renders"
)

# ── Step 3: Floor Plan (Optional) ─────────────────────────────────────────────
st.header("Step 3 — Floor Plan (Optional)")
floor_plan = st.file_uploader(
    "Upload floor plan (JPG/PNG)", type=["jpg", "jpeg", "png"],
    key="floorplan"
)

# ── Step 4: Analyse ───────────────────────────────────────────────────────────
analyse_disabled = not (render_files and client_name and pincode)
if st.button("🔍 Analyse Rooms", disabled=analyse_disabled):
    images = []
    for f in render_files:
        data = f.read()
        mt = "image/jpeg" if f.name.lower().endswith((".jpg", ".jpeg")) else "image/png"
        images.append((data, mt, f.name))
    with st.spinner("Analysing renders with Claude Vision…"):
        st.session_state.analyses = analyze_all_renders(images)
    st.success(f"Detected {len(st.session_state.analyses)} room(s)")

# ── Step 5: Review / Edit Rooms ───────────────────────────────────────────────
if "analyses" in st.session_state and st.session_state.analyses:
    st.header("Step 5 — Review Detected Rooms")
    updated = []
    for i, a in enumerate(st.session_state.analyses):
        conf_color = {"high": "🟢", "medium": "🟡", "low": "🔴"}.get(a.confidence, "⚪")
        with st.expander(f"{conf_color} {a.image_filename or f'Image {i+1}'} — {a.room_type}"):
            col1, col2 = st.columns(2)
            with col1:
                rt = st.selectbox("Room Type", ROOM_TYPES,
                                  index=ROOM_TYPES.index(a.room_type) if a.room_type in ROOM_TYPES else 0,
                                  key=f"rt_{i}")
                sqft = st.number_input("Estimated sqft", min_value=20, max_value=2000,
                                       value=a.estimated_sqft, key=f"sqft_{i}")
            with col2:
                elements = st.text_area("Design Elements", value=a.design_elements,
                                        height=120, key=f"elem_{i}")
            updated.append(RoomAnalysis(
                room_type=rt, estimated_sqft=sqft, confidence=a.confidence,
                design_elements=elements, image_filename=a.image_filename
            ))
    st.session_state.analyses = updated

# ── Step 6: Output Format ─────────────────────────────────────────────────────
st.header("Step 6 — Output Format")
output_format = st.radio(
    "Choose format",
    ["📊 Plain Excel (internal use)", "🎨 Branded DOCX + PDF (client delivery)"],
    horizontal=True,
)
branded = output_format.startswith("🎨")

# ── Step 7: Generate ──────────────────────────────────────────────────────────
st.header("Step 7 — Generate")

analyses = st.session_state.get("analyses", [])
can_generate = bool(analyses and client_name and pincode and city)

if st.button("⚡ Generate BOQ + Vendors (parallel)", disabled=not can_generate, type="primary"):
    with st.spinner("Running BOQ and vendor search simultaneously…"):
        try:
            with ThreadPoolExecutor(max_workers=2) as executor:
                fut_boq    = executor.submit(generate_boq, analyses, city, pincode, tier)
                fut_vendor = executor.submit(generate_vendors, analyses, city, pincode, tier)
                rows, sources = fut_boq.result()
                vendors, notes = fut_vendor.result()
        except Exception as exc:
            st.error(f"API error: {exc}")
            st.stop()

    safe = re.sub(r"[^A-Za-z0-9_]", "", client_name.replace(" ", "_"))
    errors = []

    # ── BOQ outputs ────────────────────────────────────────────────────────
    if not rows:
        errors.append("No BOQ rows returned.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_name = f"{safe}_BOQ_{city}.xlsx"
            xlsx_path = os.path.join(tmpdir, xlsx_name)
            write_budget_excel(client_name, city, pincode, tier, rows, sources, xlsx_path)
            with open(xlsx_path, "rb") as fh:
                st.session_state.budget_excel_bytes = fh.read()
            st.session_state.budget_excel_name = xlsx_name

            if branded:
                docx_name = f"Houspire_BOQ_{safe}_{city}.docx"
                docx_path = os.path.join(tmpdir, docx_name)
                generate_branded_boq_docx(client_name, city, tier, xlsx_path, docx_path)
                with open(docx_path, "rb") as fh:
                    st.session_state.budget_docx_bytes = fh.read()
                st.session_state.budget_docx_name = docx_name
                try:
                    pdf_path = docx_to_pdf(docx_path, tmpdir)
                    with open(pdf_path, "rb") as fh:
                        st.session_state.budget_pdf_bytes = fh.read()
                    st.session_state.budget_pdf_name = docx_name.replace(".docx", ".pdf")
                except Exception as e:
                    st.warning(f"PDF conversion failed: {e}")
        st.session_state.boq_rows = rows
        st.session_state.boq_sources = sources

    # ── Vendor outputs ─────────────────────────────────────────────────────
    if not vendors:
        errors.append("No vendors returned.")
    else:
        centroid = (0.0, 0.0)
        m = re.search(r"centroid[:\s]+\(?([\d.]+),\s*([\d.]+)\)?", notes, re.IGNORECASE)
        if m:
            centroid = (float(m.group(1)), float(m.group(2)))
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_name = f"{city}_{safe}_Vendors_{pincode}.xlsx"
            xlsx_path = os.path.join(tmpdir, xlsx_name)
            write_vendor_excel(client_name, city, pincode, centroid, vendors, notes, xlsx_path)
            with open(xlsx_path, "rb") as fh:
                st.session_state.vendor_excel_bytes = fh.read()
            st.session_state.vendor_excel_name = xlsx_name

            if branded:
                docx_name = f"Houspire_Vendors_{safe}_{city}.docx"
                docx_path = os.path.join(tmpdir, docx_name)
                generate_branded_vendor_docx(client_name, city, vendors, docx_path)
                with open(docx_path, "rb") as fh:
                    st.session_state.vendor_docx_bytes = fh.read()
                st.session_state.vendor_docx_name = docx_name
                try:
                    pdf_path = docx_to_pdf(docx_path, tmpdir)
                    with open(pdf_path, "rb") as fh:
                        st.session_state.vendor_pdf_bytes = fh.read()
                    st.session_state.vendor_pdf_name = docx_name.replace(".docx", ".pdf")
                except Exception as e:
                    st.warning(f"PDF conversion failed: {e}")
        st.session_state.vendors = vendors
        st.session_state.vendor_notes = notes

    # ── Save + status ──────────────────────────────────────────────────────
    if not errors:
        pid = save_project(client_name, city, pincode, tier, rows, sources, vendors, notes)
        if pid:
            st.caption(f"✅ Saved to DB (project {pid[:8]}…)")
        st.success(f"✅ BOQ: {len(rows)} line items | Vendors: {len(vendors)} entries")
    else:
        for e in errors:
            st.error(e)

st.caption("Or generate separately:")
col_boq, col_vendor = st.columns(2)
with col_boq:
    if st.button("📊 BOQ only", disabled=not can_generate):
        with st.spinner("Generating BOQ…"):
            rows, sources = generate_boq(analyses, city, pincode, tier)
        if rows:
            safe = re.sub(r"[^A-Za-z0-9_]", "", client_name.replace(" ", "_"))
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_name = f"{safe}_BOQ_{city}.xlsx"
                xlsx_path = os.path.join(tmpdir, xlsx_name)
                write_budget_excel(client_name, city, pincode, tier, rows, sources, xlsx_path)
                with open(xlsx_path, "rb") as fh:
                    st.session_state.budget_excel_bytes = fh.read()
                st.session_state.budget_excel_name = xlsx_name
            st.session_state.boq_rows = rows
            st.session_state.boq_sources = sources
            st.success(f"BOQ: {len(rows)} items")
        else:
            st.error("No BOQ rows returned.")

with col_vendor:
    if st.button("📍 Vendors only", disabled=not can_generate):
        with st.spinner("Searching vendors…"):
            vendors, notes = generate_vendors(analyses, city, pincode, tier)
        if vendors:
            safe = re.sub(r"[^A-Za-z0-9_]", "", client_name.replace(" ", "_"))
            centroid = (0.0, 0.0)
            m = re.search(r"centroid[:\s]+\(?([\d.]+),\s*([\d.]+)\)?", notes, re.IGNORECASE)
            if m:
                centroid = (float(m.group(1)), float(m.group(2)))
            with tempfile.TemporaryDirectory() as tmpdir:
                xlsx_name = f"{city}_{safe}_Vendors_{pincode}.xlsx"
                xlsx_path = os.path.join(tmpdir, xlsx_name)
                write_vendor_excel(client_name, city, pincode, centroid, vendors, notes, xlsx_path)
                with open(xlsx_path, "rb") as fh:
                    st.session_state.vendor_excel_bytes = fh.read()
                st.session_state.vendor_excel_name = xlsx_name
            st.session_state.vendors = vendors
            st.session_state.vendor_notes = notes
            st.success(f"Vendors: {len(vendors)} entries")
        else:
            st.error("No vendors returned.")

# ── Downloads ─────────────────────────────────────────────────────────────────
downloads = [
    ("budget_excel_bytes", "budget_excel_name", "📊 Download Budget Excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("vendor_excel_bytes", "vendor_excel_name", "📍 Download Vendor Excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("budget_docx_bytes",  "budget_docx_name",  "📄 Download Budget DOCX",  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ("vendor_docx_bytes",  "vendor_docx_name",  "📄 Download Vendor DOCX",  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ("budget_pdf_bytes",   "budget_pdf_name",   "📑 Download Budget PDF",   "application/pdf"),
    ("vendor_pdf_bytes",   "vendor_pdf_name",   "📑 Download Vendor PDF",   "application/pdf"),
]

has_any = any(k in st.session_state for k, _, _, _ in downloads)
if has_any:
    st.header("Downloads")
    cols = st.columns(3)
    col_idx = 0
    for key, name_key, label, mime in downloads:
        if key in st.session_state:
            with cols[col_idx % 3]:
                st.download_button(
                    label=label,
                    data=st.session_state[key],
                    file_name=st.session_state[name_key],
                    mime=mime,
                )
            col_idx += 1
