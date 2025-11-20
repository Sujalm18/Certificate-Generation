# app.py
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from pathlib import Path
import io
from PIL import Image, ImageDraw, ImageFont
from zipfile import ZipFile
import random
import time
import base64
import re

# --------------------------
# PAGE CONFIG
# --------------------------
st.set_page_config(page_title="Certificate Generator", layout="wide")

# --------------------------
# SITE LOGO (UI ONLY, NOT PDF)
# --------------------------
LOGO_FILE = "logo.png"
if Path(LOGO_FILE).exists():
    b64logo = base64.b64encode(Path(LOGO_FILE).read_bytes()).decode()
    # centered & modest size
    st.markdown(
        f"""
        <div style='display:flex; justify-content:center; margin-top:-20px; margin-bottom:6px;'>
            <img src="data:image/png;base64,{b64logo}" style="width:120px; max-width:25%; height:auto;" />
        </div>
        """,
        unsafe_allow_html=True
    )

# --------------------------
# TITLE
# --------------------------
st.markdown(
    "<h1 style='text-align:center; margin-top:0.1rem;'>PHN Certificate Generator</h1>",
    unsafe_allow_html=True
)

# --------------------------
# CONSTANTS / DEFAULTS
# --------------------------
DEFAULT_FONT_FILE = "Times New Roman Italic.ttf"
FONT_PATH = Path(DEFAULT_FONT_FILE)

# default templates in repo root (exact filenames you provided)
DEFAULT_QUALIFIED = "phnscholar qualified certificate.pdf"
DEFAULT_PARTICIPATED = "phnscholar participation certificate.pdf"
DEFAULT_SMARTEDGE = "smart edge workshop certificate.pdf"

DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19.0
DEFAULT_MAX_WIDTH_CM = 16.0
DPI = 300

def cm_to_px(cm: float, dpi: int = DPI) -> int:
    return int((cm / 2.54) * dpi)

def safe_filename(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

# --------------------------
# MESSAGES
# --------------------------
FUNNY_ERRORS = [
    "You selected NOTHING. I can't make certificates out of vibes 😅",
    "Did you mean invisible certificates? Pick at least one checkbox! 🫥",
    "No selection detected. My crystal ball is on lunch. Pick something! 🔮🍔",
    "I need a target — pick a group or I'll generate imaginary friends. 👻",
    "Zero choices found. The app prefers options, not silence. 😶‍🌫️"
]

MISSING_TEMPLATE_ERRORS = [
    "Template missing! Even superheroes need costumes — upload the PDF. 🦸‍♂️",
    "No template found. Please upload the PDF unless you want blank sheets. 📝❌",
]

MISSING_SHEET_ERRORS = [
    "Excel missing the needed sheet. Did it go on vacation? 🏖️",
    "No matching sheet — try renaming it to Names / Name / Smart Edge / Certificates.",
]

# --------------------------
# DRAW / PDF HELPERS
# --------------------------
def draw_name_on_template(template_bytes: bytes, name: str, x_cm: float, y_cm: float,
                          font_size_pt: float, max_width_cm: float) -> Image.Image:
    """Render first page of PDF template to image, draw centered name at x_cm,y_cm from left/bottom."""
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=DPI)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)

    # load TTF if available
    if FONT_PATH.exists():
        try:
            font_px = max(8, int(round(font_size_pt * DPI / 72.0)))
            font = ImageFont.truetype(str(FONT_PATH), font_px)
        except Exception:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    x_px = cm_to_px(x_cm)
    y_px_from_bottom = cm_to_px(y_cm)
    y_px = img.height - y_px_from_bottom

    # compute text size (try textbbox, fallback to textsize)
    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except Exception:
        text_w, text_h = draw.textsize(name, font=font)

    max_w_px = cm_to_px(max_width_cm)
    if text_w > max_w_px:
        # scale font down if using truetype
        try:
            if hasattr(font, "path"):
                current_px = getattr(font, "size", font_px)
                scale = max_w_px / text_w
                new_px = max(8, int(round(current_px * scale)))
                font = ImageFont.truetype(str(FONT_PATH), new_px)
                bbox = draw.textbbox((0, 0), name, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
        except Exception:
            pass

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # draw white outline for contrast then black fill
    outline_color = "white"
    fill_color = "black"
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill=outline_color)
    draw.text((draw_x, draw_y), name, font=font, fill=fill_color)

    return img

def image_to_pdf_bytes(img: Image.Image) -> bytes:
    out = io.BytesIO()
    img.convert("RGB").save(out, format="PDF")
    return out.getvalue()

# --------------------------
# UI: Uploads & Settings
# --------------------------
st.header("1) Upload files (Excel must contain sheets QUALIFIED, PARTICIPATED and Smart Edge sheet)")

excel_file = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx", "xls"])
uploaded_qual = st.file_uploader("Qualified template PDF (optional)", type=["pdf"])
uploaded_part = st.file_uploader("Participated template PDF (optional)", type=["pdf"])
uploaded_smart = st.file_uploader("SMART EDGE template PDF (optional)", type=["pdf"])
ttf_upload = st.file_uploader("Times New Roman Italic TTF (optional)", type=["ttf","otf"])

# read uploaded bytes once
qual_bytes = uploaded_qual.read() if uploaded_qual else None
part_bytes = uploaded_part.read() if uploaded_part else None
smart_bytes = uploaded_smart.read() if uploaded_smart else None

# fallback to defaults in repo root if upload not provided
if qual_bytes is None and Path(DEFAULT_QUALIFIED).exists():
    qual_bytes = Path(DEFAULT_QUALIFIED).read_bytes()
if part_bytes is None and Path(DEFAULT_PARTICIPATED).exists():
    part_bytes = Path(DEFAULT_PARTICIPATED).read_bytes()
if smart_bytes is None and Path(DEFAULT_SMARTEDGE).exists():
    smart_bytes = Path(DEFAULT_SMARTEDGE).read_bytes()

# optionally override TTF (save to disk so PIL can load)
if ttf_upload:
    with open("uploaded_times.ttf", "wb") as f:
        f.write(ttf_upload.getbuffer())
    FONT_PATH = Path("uploaded_times.ttf")

# --------------------------
# SIDEBAR: position & font
# --------------------------
st.sidebar.header("Rasterize output (recommended)")
rasterize = st.sidebar.checkbox("Rasterize certificates", value=True)

st.sidebar.header("Position & font settings")
# enforce float defaults to avoid Streamlit mixed numeric type errors
X_CM = float(st.sidebar.number_input("X (cm from left)", value=float(DEFAULT_X_CM), format="%.2f", step=0.01))
Y_CM = float(st.sidebar.number_input("Y (cm from bottom)", value=float(DEFAULT_Y_CM), format="%.2f", step=0.01))
BASE_FONT_PT = float(st.sidebar.number_input("Base font size (pt)", value=float(DEFAULT_FONT_PT), step=1.0))
MAX_WIDTH_CM = float(st.sidebar.number_input("Max name width (cm)", value=float(DEFAULT_MAX_WIDTH_CM), step=0.5))

# --------------------------
# LIVE PREVIEW
# --------------------------
st.markdown("---")
st.subheader("Live Preview")

preview_name = st.text_input("Preview name", "Aarav Sharma")

preview_option = st.selectbox("Template for preview", ["Qualified (upload or default)", "Participated (upload or default)", "Smart Edge (upload or default)"])

preview_template_bytes = None
if preview_option.startswith("Qualified"):
    preview_template_bytes = qual_bytes
elif preview_option.startswith("Participated"):
    preview_template_bytes = part_bytes
else:
    preview_template_bytes = smart_bytes

preview_col = st.container()
if preview_template_bytes is not None:
    try:
        preview_img = draw_name_on_template(preview_template_bytes, preview_name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
        # use_container_width to avoid deprecation warning
        preview_col.image(preview_img, caption="Live certificate preview (rasterized)", use_container_width=True)
    except Exception as e:
        preview_col.error(f"Preview error: {e}")
else:
    preview_col.info("Upload or provide at least one template (or ensure default templates exist in repo root) to enable preview.")

# --------------------------
# CHECKBOXES & centered caption
# --------------------------
st.markdown("---")
st.markdown("### 2) Select which certificates to generate")
col1, col2, col3 = st.columns([1,1,1])
with col1:
    gen_qualified = st.checkbox("Generate QUALIFIED", value=False)
with col2:
    gen_participated = st.checkbox("Generate PARTICIPATED", value=False)
with col3:
    gen_smartedge = st.checkbox("Generate SMART EDGE WORKSHOP", value=False)

st.markdown(
    "<div style='text-align:center; opacity:0.85; margin-top:10px;'>"
    "Select which certificates to include in the ZIP. Uncheck to exclude a group."
    "</div>",
    unsafe_allow_html=True
)

# add vertical space before button
st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)

# --------------------------
# GENERATE ZIP
# --------------------------
if st.button("Generate certificates ZIP"):

    # funny error if nothing selected
    if not (gen_qualified or gen_participated or gen_smartedge):
        st.error(random.choice(FUNNY_ERRORS))
        st.stop()

    if excel_file is None:
        st.error(random.choice(MISSING_SHEET_ERRORS))
        st.stop()

    # read excel
    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Cannot read Excel: {e}")
        st.stop()

    # Smart Edge allowed sheet names
    smart_allowed = {"NAMES", "NAME", "SMART EDGE", "CERTIFICATES"}
    smart_sheet = None
    for s in xls.sheet_names:
        if s.strip().upper() in smart_allowed:
            smart_sheet = s
            break

    # template checks
    if gen_qualified and not qual_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Qualified)")
        st.stop()
    if gen_participated and not part_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Participated)")
        st.stop()
    if gen_smartedge and not smart_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Smart Edge)")
        st.stop()
    if gen_smartedge and not smart_sheet:
        st.error("Smart Edge sheet missing! Use Names / Name / Smart Edge / Certificates.")
        st.stop()

    # read sheets
    df_q = pd.read_excel(excel_file, sheet_name="QUALIFIED", dtype=object) if ("QUALIFIED" in [s.upper() for s in xls.sheet_names]) else pd.DataFrame()
    df_p = pd.read_excel(excel_file, sheet_name="PARTICIPATED", dtype=object) if ("PARTICIPATED" in [s.upper() for s in xls.sheet_names]) else pd.DataFrame()
    df_s = pd.read_excel(excel_file, sheet_name=smart_sheet, dtype=object) if smart_sheet else pd.DataFrame()

    # build tasks
    tasks = []
    group_counts = {"QUALIFIED": 0, "PARTICIPATED": 0, "SMART_EDGE_WORKSHOP": 0}

    if gen_qualified and not df_q.empty:
        q_names = df_q.iloc[:,0].dropna().astype(str).tolist()
        group_counts["QUALIFIED"] = len(q_names)
        tasks += [("QUALIFIED", n.strip()) for n in q_names]

    if gen_participated and not df_p.empty:
        p_names = df_p.iloc[:,0].dropna().astype(str).tolist()
        group_counts["PARTICIPATED"] = len(p_names)
        tasks += [("PARTICIPATED", n.strip()) for n in p_names]

    if gen_smartedge and not df_s.empty:
        s_names = df_s.iloc[:,0].dropna().astype(str).tolist()
        group_counts["SMART_EDGE_WORKSHOP"] = len(s_names)
        tasks += [("SMART_EDGE_WORKSHOP", n.strip()) for n in s_names]

    if len(tasks) == 0:
        st.warning("No names found in the selected sheets. Nothing to generate.")
        st.stop()

    total = len(tasks)
    overall_progress = st.progress(0.0)
    overall_status = st.empty()

    # per-group placeholders
    group_done = {g: 0 for g in group_counts}
    group_status_placeholders = {}
    group_progress_bars = {}
    for g, cnt in group_counts.items():
        if cnt > 0:
            group_status_placeholders[g] = st.empty()
            group_progress_bars[g] = st.progress(0.0)

    zip_buf = io.BytesIO()
    # Keep track of used filenames to prevent duplicates in ZIP
    used_names = {}

    with ZipFile(zip_buf, "w") as zf:
        for idx, (group, name) in enumerate(tasks, start=1):
            group_done[group] += 1
            overall_status.info(f"Overall: {idx}/{total} — Generating {group} / {name}")

            for g in group_status_placeholders:
                done = group_done.get(g, 0)
                total_g = group_counts.get(g, 0)
                group_status_placeholders[g].text(f"{g.replace('_',' ')}: {done}/{total_g} done")
                if total_g > 0:
                    group_progress_bars[g].progress(done / total_g)

            time.sleep(0.005)

            try:
                if group == "QUALIFIED":
                    tpl_bytes = qual_bytes
                elif group == "PARTICIPATED":
                    tpl_bytes = part_bytes
                else:
                    tpl_bytes = smart_bytes

                img = draw_name_on_template(tpl_bytes, name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
                pdf_bytes = image_to_pdf_bytes(img)

                base_name = safe_filename(name)
                # ensure uniqueness within group folder
                key = f"{group}/{base_name}"
                count = used_names.get(key, 0)
                if count == 0:
                    filename = f"{group}/{base_name}.pdf"
                else:
                    filename = f"{group}/{base_name}_{count}.pdf"
                used_names[key] = count + 1

                zf.writestr(filename, pdf_bytes)
            except Exception as e:
                err_msg = f"Failed to generate for {name}: {e}"
                base_name = safe_filename(name) or "error"
                key = f"{group}/{base_name}_ERROR"
                count = used_names.get(key, 0)
                if count == 0:
                    filename = f"{group}/{base_name}_ERROR.txt"
                else:
                    filename = f"{group}/{base_name}_ERROR_{count}.txt"
                used_names[key] = count + 1
                zf.writestr(filename, err_msg.encode("utf-8"))

            overall_progress.progress(idx / total)

        overall_status.success("All items processed. Finalizing ZIP...")

    st.balloons()
    st.success(f"Done — {total} certificates generated (errors, if any, are in ZIP).")
    zip_buf.seek(0)
    st.download_button("Download certificates ZIP", data=zip_buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
