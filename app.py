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

# --------------------------
# PAGE CONFIG
# --------------------------
st.set_page_config(page_title="Certificate Generator", layout="wide")

# --------------------------
# CENTERED LOGO (UI ONLY, NOT PDF)
# --------------------------
logo_file = "logo.png"
if Path(logo_file).exists():
    b64logo = base64.b64encode(Path(logo_file).read_bytes()).decode()
    st.markdown(
        f"""
        <div style='display:flex; justify-content:center; margin-top:-20px; margin-bottom:-10px;'>
            <img src="data:image/png;base64,{b64logo}" style="width:140px;" />
        </div>
        """,
        unsafe_allow_html=True
    )

# --------------------------
# TITLE
# --------------------------
st.markdown(
    "<h1 style='text-align:center;'>Certificate Generator — QUALIFIED, PARTICIPATED & SMART EDGE WORKSHOP</h1>",
    unsafe_allow_html=True
)

# --------------------------
# CONSTANTS
# --------------------------
DEFAULT_FONT_FILE = "Times New Roman Italic.ttf"
FONT_PATH = Path(DEFAULT_FONT_FILE)

DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19
DEFAULT_MAX_WIDTH_CM = 16
DPI = 300

def cm_to_px(cm, dpi=DPI):
    return int((cm / 2.54) * dpi)

# --------------------------
# FUNNY ERRORS
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
    "Template not uploaded — certificates won’t dress themselves. 👔"
]

MISSING_SHEET_ERRORS = [
    "Excel missing the needed sheet. Did it go on vacation? 🏖️",
    "Required sheet not found. Please use the correct sheet name. 📄",
    "No matching sheet — try renaming it to Names / Name / Smart Edge / Certificates."
]

# --------------------------
# DRAW NAME ON PDF TEMPLATE
# --------------------------
def draw_name_on_template(template_bytes, name, x_cm, y_cm, font_size_pt, max_width_cm):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]

    pix = page.get_pixmap(dpi=DPI)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)

    if FONT_PATH.exists():
        try:
            font_px = max(8, int(font_size_pt * DPI / 72))
            font = ImageFont.truetype(str(FONT_PATH), font_px)
        except:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    x_px = cm_to_px(x_cm)
    y_px = img.height - cm_to_px(y_cm)

    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except:
        text_w, text_h = draw.textsize(name, font=font)

    max_w_px = cm_to_px(max_width_cm)
    if text_w > max_w_px:
        try:
            scale = max_w_px / text_w
            new_font_px = int(font.size * scale)
            font = ImageFont.truetype(str(FONT_PATH), new_font_px)
            bbox = draw.textbbox((0, 0), name, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
        except:
            pass

    draw_x = int(x_px - text_w/2)
    draw_y = int(y_px - text_h/2)

    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill="white")

    draw.text((draw_x, draw_y), name, font=font, fill="black")

    return img

def image_to_pdf_bytes(img):
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PDF")
    return buf.getvalue()

# --------------------------
# UPLOAD UI
# --------------------------
st.header("1) Upload files (Excel must contain sheets QUALIFIED, PARTICIPATED and Smart Edge sheet)")

excel_file = st.file_uploader("Upload Excel", type=["xlsx","xls"])
qualified_pdf_file = st.file_uploader("Qualified template PDF", type=["pdf"])
participated_pdf_file = st.file_uploader("Participated template PDF", type=["pdf"])
smartedge_pdf_file = st.file_uploader("Smart Edge Workshop template PDF", type=["pdf"])

# --------------------------
# SIDEBAR SETTINGS
# --------------------------
st.sidebar.header("Position & font settings")
X_CM = st.sidebar.number_input("X (cm from left)", value=float(DEFAULT_X_CM), step=0.01)
Y_CM = st.sidebar.number_input("Y (cm from bottom)", value=float(DEFAULT_Y_CM), step=0.01)
BASE_FONT_PT = st.sidebar.number_input("Base font size", value=int(DEFAULT_FONT_PT), step=1)
MAX_WIDTH_CM = st.sidebar.number_input("Max name width (cm)", value=float(DEFAULT_MAX_WIDTH_CM), step=0.5)

# --------------------------
# CHECKBOXES
# --------------------------
st.markdown("### 3) Generate and download final ZIP")

col1, col2, col3 = st.columns([1,1,1])
with col1: gen_q = st.checkbox("Generate QUALIFIED")
with col2: gen_p = st.checkbox("Generate PARTICIPATED")
with col3: gen_s = st.checkbox("Generate SMART EDGE WORKSHOP")

st.markdown(
    "<div style='text-align:center; opacity:0.75;'>Select which certificates to include in the ZIP.</div>",
    unsafe_allow_html=True
)

# --------------------------
# GENERATION
# --------------------------
if st.button("Generate certificates ZIP"):

    if not (gen_q or gen_p or gen_s):
        st.error(random.choice(FUNNY_ERRORS))
        st.stop()

    if excel_file is None:
        st.error("Upload the Excel file first.")
        st.stop()

    xls = pd.ExcelFile(excel_file)

    # Smart Edge sheet detection
    allowed_sheets = ["NAMES", "NAME", "SMART EDGE", "CERTIFICATES"]
    smart_sheet = None
    for s in xls.sheet_names:
        if s.strip().upper() in allowed_sheets:
            smart_sheet = s
            break

    # Read templates
    qual_bytes = qualified_pdf_file.read() if (gen_q and qualified_pdf_file) else None
    part_bytes = participated_pdf_file.read() if (gen_p and participated_pdf_file) else None
    smart_bytes = smartedge_pdf_file.read() if (gen_s and smartedge_pdf_file) else None

    if gen_s and smart_sheet is None:
        st.error("Smart Edge sheet missing! Use Names / Name / Smart Edge / Certificates.")
        st.stop()

    # Read names
    df_q = pd.read_excel(excel_file, "QUALIFIED") if gen_q else pd.DataFrame()
    df_p = pd.read_excel(excel_file, "PARTICIPATED") if gen_p else pd.DataFrame()
    df_s = pd.read_excel(excel_file, smart_sheet) if (gen_s and smart_sheet) else pd.DataFrame()

    tasks = []
    if gen_q: tasks += [("QUALIFIED", n) for n in df_q.iloc[:,0].dropna()]
    if gen_p: tasks += [("PARTICIPATED", n) for n in df_p.iloc[:,0].dropna()]
    if gen_s: tasks += [("SMART", n) for n in df_s.iloc[:,0].dropna()]

    total = len(tasks)
    overall = st.progress(0)

    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, "w") as z:
        for idx, (group, name) in enumerate(tasks, start=1):

            tpl = qual_bytes if group=="QUALIFIED" else part_bytes if group=="PARTICIPATED" else smart_bytes
            img = draw_name_on_template(tpl, str(name), X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
            pdf_bytes = image_to_pdf_bytes(img)

            safe = str(name).replace("/", "_")
            z.writestr(f"{group}/{safe}.pdf", pdf_bytes)

            overall.progress(idx/total)
            time.sleep(0.01)

    st.success("Done!")
    st.download_button("Download ZIP", zip_buffer.getvalue(), "certificates.zip", "application/zip")
