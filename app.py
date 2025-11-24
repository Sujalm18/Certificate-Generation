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

# --------------------------
# PAGE CONFIG
# --------------------------
st.set_page_config(page_title="Certificate Generator", layout="wide")

# --------------------------
# SITE LOGO (Streamlit UI only — NOT inserted into PDFs)
# --------------------------
# Path to your logo inside the repo
# --------------------------
# SITE LOGO (UI only — not added to PDFs)
# --------------------------
# --------------------------
# SMALL CENTERED LOGO (Streamlit only)
# --------------------------
from PIL import Image

logo_path = Path("logo.png")

if logo_path.exists():
    try:
        img = Image.open(logo_path)
        st.image(img, width=45)   # <-- SUPER SMALL (45px)
    except:
        st.warning("Logo found but could not be loaded.")
else:
    st.info("logo.png not found. Add it to the root directory.")



# --------------------------
# TITLE
# --------------------------
st.markdown(
    "<h1 style='text-align:center;'>PHN Certificate Generator</h1>",
    unsafe_allow_html=True
)


# -------------------------------------------
# CONSTANTS
# -------------------------------------------
DEFAULT_FONT_FILE = "Times New Roman Italic.ttf"
FONT_PATH = Path(DEFAULT_FONT_FILE)

DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19
DEFAULT_MAX_WIDTH_CM = 16
DPI = 300


def cm_to_px(cm, dpi=DPI):
    return int((cm / 2.54) * dpi)


# -------------------------------------------
# FUNNY MESSAGES
# -------------------------------------------
FUNNY_ERRORS = [
    "You selected NOTHING. I can't make certificates out of vibes 😅",
    "Did you mean invisible certificates? Pick at least one checkbox! 🫥",
    "No selection detected. My crystal ball is on lunch. Pick something! 🔮🍔",
    "I need a target — pick a group or I'll generate imaginary friends. 👻",
    "Zero choices found. The app prefers options, not silence. 😶‍🌫️"
]


# -------------------------------------------
# DRAW NAME ON TEMPLATE
# -------------------------------------------
def draw_name_on_template(template_bytes, name, x_cm, y_cm, font_size_pt, max_width_cm):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]

    pix = page.get_pixmap(dpi=DPI)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)

    # Load font
    if FONT_PATH.exists():
        try:
            font_px = int(font_size_pt * DPI / 72)
            font = ImageFont.truetype(str(FONT_PATH), font_px)
        except:
            font = ImageFont.load_default()
    else:
        font = ImageFont.load_default()

    # Compute pixel coords
    x_px = cm_to_px(x_cm)
    y_px = img.height - cm_to_px(y_cm)

    # Get text width
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
            new_font_px = max(8, int(font.size * scale))
            font = ImageFont.truetype(str(FONT_PATH), new_font_px)
            bbox = draw.textbbox((0, 0), name, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
        except:
            pass

    draw_x = int(x_px - text_w / 2)
    draw_y = int(y_px - text_h / 2)

    # White outline
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill="white")

    draw.text((draw_x, draw_y), name, font=font, fill="black")

    return img


def image_to_pdf_bytes(img):
    out = io.BytesIO()
    img.convert("RGB").save(out, format="PDF")
    return out.getvalue()


# -------------------------------------------
# FILE UPLOAD SECTION
# -------------------------------------------
st.header("1) Upload files")

excel_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
qualified_pdf = st.file_uploader("Qualified Certificate Template", type=["pdf"])
participated_pdf = st.file_uploader("Participated Certificate Template", type=["pdf"])
smartedge_pdf = st.file_uploader("Smart Edge Workshop Template", type=["pdf"])
ttf_upload = st.file_uploader("Upload Times New Roman Italic (optional)", type=["ttf", "otf"])

if ttf_upload:
    with open("uploaded_font.ttf", "wb") as f:
        f.write(ttf_upload.getbuffer())
    FONT_PATH = Path("uploaded_font.ttf")

# -------------------------------------------
# LIVE PREVIEW — ALWAYS VISIBLE
# -------------------------------------------
st.subheader("Live Preview (Sample Name)")

if participated_pdf:
    sample_bytes = participated_pdf.read()
    sample_img = draw_name_on_template(sample_bytes, "Sample Student", DEFAULT_X_CM, DEFAULT_Y_CM, DEFAULT_FONT_PT, DEFAULT_MAX_WIDTH_CM)
    st.image(sample_img, caption="Live Preview", use_column_width=True)


# -------------------------------------------
# SIDEBAR SETTINGS
# -------------------------------------------
st.sidebar.header("Position & Font")
X_CM = st.sidebar.number_input("X (cm from left)", value=DEFAULT_X_CM, step=0.01, format="%.2f")
Y_CM = st.sidebar.number_input("Y (cm from bottom)", value=DEFAULT_Y_CM, step=0.01, format="%.2f")
FONT_PT = st.sidebar.number_input("Font Size (pt)", value=DEFAULT_FONT_PT)
MAX_WIDTH_CM = st.sidebar.number_input("Max Width (cm)", value=DEFAULT_MAX_WIDTH_CM)

# -------------------------------------------
# CHECKBOX OPTIONS
# -------------------------------------------
st.subheader("3) Choose certificate types")
col1, col2, col3 = st.columns(3)
gen_qualified = col1.checkbox("Generate QUALIFIED")
gen_participated = col2.checkbox("Generate PARTICIPATED")
gen_smartedge = col3.checkbox("Generate SMART EDGE WORKSHOP")

st.markdown(
    "<div style='text-align:center; margin-top:10px; margin-bottom:25px; opacity:0.75;'>"
    "Select which certificates to include in the ZIP. Uncheck to exclude a group."
    "</div>",
    unsafe_allow_html=True
)

# -------------------------------------------
# GENERATE BUTTON
# -------------------------------------------
if st.button("Generate Certificates ZIP"):

    if not (gen_qualified or gen_participated or gen_smartedge):
        st.error(random.choice(FUNNY_ERRORS))
        st.stop()

    if excel_file is None:
        st.error("Upload Excel file first.")
        st.stop()

    xls = pd.ExcelFile(excel_file)

    # smart edge accepted sheets
    smartedge_sheets = ["NAMES", "NAME", "SMART EDGE", "CERTIFICATES"]
    smartedge_sheet = None

    for s in xls.sheet_names:
        if s.strip().upper() in smartedge_sheets:
            smartedge_sheet = s
            break

    # read data
    df_q = pd.read_excel(excel_file, sheet_name="QUALIFIED") if "QUALIFIED" in [s.upper() for s in xls.sheet_names] else pd.DataFrame()
    df_p = pd.read_excel(excel_file, sheet_name="PARTICIPATED") if "PARTICIPATED" in [s.upper() for s in xls.sheet_names] else pd.DataFrame()
    df_s = pd.read_excel(excel_file, sheet_name=smartedge_sheet) if smartedge_sheet else pd.DataFrame()

    zip_buffer = io.BytesIO()

    with ZipFile(zip_buffer, "w") as zf:

        # QUALIFIED
        if gen_qualified and not df_q.empty:
            for name in df_q.iloc[:, 0].dropna():
                img = draw_name_on_template(qualified_pdf.read(), str(name), X_CM, Y_CM, FONT_PT, MAX_WIDTH_CM)
                zf.writestr(f"QUALIFIED/{name}.pdf", image_to_pdf_bytes(img))

        # PARTICIPATED
        if gen_participated and not df_p.empty:
            for name in df_p.iloc[:, 0].dropna():
                img = draw_name_on_template(participated_pdf.read(), str(name), X_CM, Y_CM, FONT_PT, MAX_WIDTH_CM)
                zf.writestr(f"PARTICIPATED/{name}.pdf", image_to_pdf_bytes(img))

        # SMART EDGE WORKSHOP
        if gen_smartedge and not df_s.empty:
            for name in df_s.iloc[:, 0].dropna():
                img = draw_name_on_template(smartedge_pdf.read(), str(name), X_CM, Y_CM, FONT_PT, MAX_WIDTH_CM)
                zf.writestr(f"SMART_EDGE/{name}.pdf", image_to_pdf_bytes(img))

    zip_buffer.seek(0)
    st.success("ZIP generated successfully!")
    st.download_button("Download ZIP", zip_buffer, "Certificates.zip", mime="application/zip")
