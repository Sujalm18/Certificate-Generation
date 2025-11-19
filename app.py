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
import re

# --------------------------
# CONFIG
# --------------------------
st.set_page_config(page_title="Certificate Generator", layout="wide")
ROOT = Path(".")
LOGO_FILENAME = "logo.png"

# default template filenames (must match exactly in repo root)
DEFAULT_QUALIFIED = "phnscholar qualified certificate.pdf"
DEFAULT_PARTICIPATED = "phnscholar participation certificate.pdf"
DEFAULT_SMARTEDGE = "smart edge workshop certificate.pdf"

# constants (floats to avoid mixed numeric-type errors)
DEFAULT_FONT_FILE = "Times New Roman Italic.ttf"
FONT_PATH = Path(DEFAULT_FONT_FILE)
DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19.0
DEFAULT_MAX_WIDTH_CM = 16.0
DPI = 300

def cm_to_px(cm, dpi=DPI):
    return int((cm / 2.54) * dpi)

def safe_filename(s: str):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

# --------------------------
# SITE LOGO (root) - only for UI
# --------------------------
if Path(LOGO_FILENAME).exists():
    st.image(LOGO_FILENAME, width=150)

st.markdown(
    "<h1 style='text-align:center;'>Certificate Generator — QUALIFIED, PARTICIPATED & SMART EDGE WORKSHOP</h1>",
    unsafe_allow_html=True,
)

# --------------------------
# MESSAGES
# --------------------------
FUNNY_ERRORS = [
    "You selected NOTHING. I can't make certificates out of vibes 😅",
    "Did you mean invisible certificates? Pick at least one checkbox! 🫥",
    "No selection detected. My crystal ball is on lunch. Pick something! 🔮🍔",
    "I need a target — pick a group or I'll generate imaginary friends. 👻",
    "Zero choices found. The app prefers options, not silence. 😶‍🌫️",
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
    font = None
    if FONT_PATH.exists():
        try:
            font_px = max(8, int(round(font_size_pt * DPI / 72.0)))
            font = ImageFont.truetype(str(FONT_PATH), font_px)
        except Exception:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    # center coordinates in pixels
    x_px = cm_to_px(x_cm)
    y_px_from_bottom = cm_to_px(y_cm)
    y_px = img.height - y_px_from_bottom

    # compute text size and autoscale if needed
    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except Exception:
        text_w, text_h = draw.textsize(name, font=font)

    max_w_px = cm_to_px(max_width_cm)
    if text_w > max_w_px:
        try:
            if hasattr(font, "path"):
                scale = max_w_px / text_w
                new_font_px = max(8, int((font.size if hasattr(font, "size") else font_px) * scale))
                font = ImageFont.truetype(str(FONT_PATH), new_font_px)
                bbox = draw.textbbox((0, 0), name, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
        except Exception:
            pass

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # draw white outline + black fill for visibility
    outline_color = "white"
    fill_color = "black"
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x + dx, draw_y + dy), name, font=font, fill=outline_color)
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

# optionally override TTF
if ttf_upload:
    with open("uploaded_times.ttf", "wb") as f:
        f.write(ttf_upload.getbuffer())
    FONT_PATH = Path("uploaded_times.ttf")

# sidebar controls (float values)
st.sidebar.header("Rasterize output (recommended)")
rasterize = st.sidebar.checkbox("Rasterize certificates", value=True)

st.sidebar.header("Position & font settings")
X_CM = st.sidebar.number_input("X (cm from left)", value=float(DEFAULT_X_CM), format="%.2f", step=0.01)
Y_CM = st.sidebar.number_input("Y (cm from bottom)", value=float(DEFAULT_Y_CM), format="%.2f", step=0.01)
BASE_FONT_PT = st.sidebar.number_input("Base font size (pt)", value=float(DEFAULT_FONT_PT), step=1.0)
MAX_WIDTH_CM = st.sidebar.number_input("Max name width (cm)", value=float(DEFAULT_MAX_WIDTH_CM), step=0.5)

# --------------------------
# Live preview (choose which template to preview)
# --------------------------
st.markdown("---")
st.subheader("Live Preview (select template and enter name)")

preview_name = st.text_input("Preview name", "Aarav Sharma")

# preview source selection
preview_choice = st.selectbox(
    "Select template for preview",
    options=[
        ("Qualified (upload or default)", "QUAL"),
        ("Participated (upload or default)", "PART"),
        ("Smart Edge (upload or default)", "SMART")
    ],
    format_func=lambda x: x[0]
)

preview_template = None
if preview_choice[1] == "QUAL":
    preview_template = qual_bytes
elif preview_choice[1] == "PART":
    preview_template = part_bytes
else:
    preview_template = smart_bytes

if preview_template:
    try:
        img_prev = draw_name_on_template(preview_template, preview_name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
        st.image(img_prev, caption="Live certificate preview", use_column_width=True)
    except Exception as e:
        st.error(f"Preview error: {e}")
else:
    st.info("Upload or provide at least one template (or ensure default templates exist in repo root).")

# --------------------------
# checkboxes (none selected by default)
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
    "<div style='text-align:center; opacity:0.75; margin-top:10px;'>"
    "Select which certificates to include in the ZIP. Uncheck to exclude a group."
    "</div>",
    unsafe_allow_html=True
)

# --------------------------
# Generate ZIP with progress
# --------------------------
if st.button("Generate certificates ZIP"):

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

    # detect smart edge sheet (names allowed)
    smartedge_allowed = {"NAMES", "NAME", "SMART EDGE", "CERTIFICATES"}
    smartedge_sheet = None
    for s in xls.sheet_names:
        if s.strip().upper() in smartedge_allowed:
            smartedge_sheet = s
            break

    # template presence checks (use default bytes or uploaded bytes)
    if gen_qualified and not qual_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Qualified)")
        st.stop()
    if gen_participated and not part_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Participated)")
        st.stop()
    if gen_smartedge and not smart_bytes:
        st.error(MISSING_TEMPLATE_ERRORS[0] + " (Smart Edge)")
        st.stop()
    if gen_smartedge and not smartedge_sheet:
        st.error(MISSING_SHEET_ERRORS[2])
        st.stop()

    # read sheets into dataframes (if present)
    df_q = pd.read_excel(excel_file, sheet_name="QUALIFIED", dtype=object) if ("QUALIFIED" in [s.upper() for s in xls.sheet_names]) else pd.DataFrame()
    df_p = pd.read_excel(excel_file, sheet_name="PARTICIPATED", dtype=object) if ("PARTICIPATED" in [s.upper() for s in xls.sheet_names]) else pd.DataFrame()
    df_s = pd.read_excel(excel_file, sheet_name=smartedge_sheet, dtype=object) if smartedge_sheet else pd.DataFrame()

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
    group_text = {}
    group_bars = {}
    for g, cnt in group_counts.items():
        if cnt > 0:
            group_text[g] = st.empty()
            group_bars[g] = st.progress(0.0)

    # generate and write to in-memory zip
    zip_buf = io.BytesIO()
    with ZipFile(zip_buf, "w") as zf:
        for idx, (group, name) in enumerate(tasks, start=1):
            group_done[group] += 1
            overall_status.info(f"Overall: {idx}/{total} — Generating {group} / {name}")
            # update per-group UI
            for g in group_text:
                done = group_done.get(g, 0)
                total_g = group_counts.get(g, 0)
                group_text[g].text(f"{g.replace('_', ' ')}: {done}/{total_g} done")
                if total_g > 0:
                    group_bars[g].progress(done / total_g)

            time.sleep(0.01)  # small sleep so UI updates

            try:
                tpl = qual_bytes if group == "QUALIFIED" else (part_bytes if group == "PARTICIPATED" else smart_bytes)
                img = draw_name_on_template(tpl, name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
                pdf_bytes = image_to_pdf_bytes(img)
                zf.writestr(f"{group}/{safe_filename(name)}.pdf", pdf_bytes)
            except Exception as e:
                err = f"Failed to generate for {name}: {e}"
                zf.writestr(f"{group}/{safe_filename(name)}_ERROR.txt", err.encode("utf-8"))

            overall_progress.progress(idx / total)

        overall_status.success("All items processed. Finalizing ZIP...")

    st.balloons()
    st.success(f"Done — {total} certificates generated (errors, if any, are in ZIP).")
    zip_buf.seek(0)
    st.download_button("Download certificates ZIP", data=zip_buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
