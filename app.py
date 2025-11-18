import streamlit as st
import pandas as pd
import os, re, zipfile
from io import BytesIO
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter

# ----------------- Config / defaults -----------------
DATA_DIR = Path(".")
DEFAULT_QUAL_PDF = "phnscholar qualified certificate.pdf"
DEFAULT_PART_PDF = "phnscholar participation certificate.pdf"
DEFAULT_TTF = "Times New Roman Italic.ttf"  # place your TTF here if available
OUTPUT_DIR = "cert_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ----------------- Helpers -----------------
def safe_filename(s):
    s = str(s)
    s = re.sub(r'[\\/*?:"<>|]', "_", s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

def get_first_name_column(df):
    for col in df.columns:
        if df[col].notna().any():
            return col
    return None

def make_overlay_pdf(name, page_width_pt, page_height_pt, x_cm, y_cm, font_name, font_size_pt):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_width_pt, page_height_pt))
    c.setFont(font_name, font_size_pt)
    x_pt = x_cm * cm
    y_pt = y_cm * cm
    c.drawCentredString(x_pt, y_pt, str(name).strip())
    c.save()
    buf.seek(0)
    return buf

def overlay_template_with_names(template_pdf_path, names_list, out_pdf_path, x_cm, y_cm, font_size_pt, font_name):
    reader = PdfReader(str(template_pdf_path))
    if len(reader.pages) == 0:
        raise ValueError("Template PDF has no pages")
    writer = PdfWriter()
    # get page size from first page
    mediabox = reader.pages[0].mediabox
    page_width_pt = float(mediabox.width)
    page_height_pt = float(mediabox.height)
    for name in names_list:
        overlay_buf = make_overlay_pdf(name, page_width_pt, page_height_pt, x_cm, y_cm, font_name, font_size_pt)
        overlay_reader = PdfReader(overlay_buf)
        overlay_page = overlay_reader.pages[0]
        # re-read template fresh to avoid in-place merge reuse issues
        reader_fresh = PdfReader(str(template_pdf_path))
        base = reader_fresh.pages[0]
        base.merge_page(overlay_page)
        writer.add_page(base)
    with open(out_pdf_path, "wb") as f:
        writer.write(f)
    return out_pdf_path

# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="Certificate overlay - QUALIFIED/PARTICIPATED", layout="centered")
st.title("Certificate Generator — QUALIFIED & PARTICIPATED")

st.markdown("""
Upload an Excel file that contains two sheets named **QUALIFIED** and **PARTICIPATED** (case-insensitive).
Each sheet should contain one column of student names (header can be any text).  
You can either place your template PDFs and TTF in the project folder or upload them below.
""")

# Excel upload
excel_file = st.file_uploader("Upload Excel file (.xlsx/.xls) with sheets QUALIFIED and PARTICIPATED", type=["xlsx","xls"])

# Optional template PDF uploads (override defaults)
st.write("If you don't upload templates, the app will use files in the script folder named:")
st.write(f"- Qualified template: `{DEFAULT_QUAL_PDF}`") 
st.write(f"- Participated template: `{DEFAULT_PART_PDF}`")
qual_pdf_upload = st.file_uploader("Qualified template PDF (optional)", type=["pdf"])
part_pdf_upload = st.file_uploader("Participated template PDF (optional)", type=["pdf"])

# Optional TTF upload (Times New Roman Italic) or fallback
st.write("Optional: upload Times New Roman Italic TTF (exact matching font). If not provided, built-in Times-Italic is used.")
ttf_upload = st.file_uploader("Upload TTF (optional)", type=["ttf","otf"])

# Sidebar controls for positioning & font
st.sidebar.header("Name position & font (cm / pt)")
x_cm = st.sidebar.number_input("Horizontal center (x cm)", value=14.85, step=0.1,
                               help="Horizontal center position in cm (from left). Default centers on A4 landscape.")
y_cm_qual = st.sidebar.number_input("Qualified name vertical (y cm)", value=11.0, step=0.1)
y_cm_part = st.sidebar.number_input("Participated name vertical (y cm)", value=11.0, step=0.1)
font_size = st.sidebar.number_input("Font size (pt)", value=16, step=1)
# font selection will include registered TTF or built-in fallback
font_choice = st.sidebar.selectbox("Font fallback", options=["Times-Italic","Helvetica-Bold","Times-Roman"])

# Save uploaded files to disk (if uploaded) or use defaults
if qual_pdf_upload:
    qual_pdf_path = Path("qualified_uploaded.pdf")
    with open(qual_pdf_path, "wb") as f:
        f.write(qual_pdf_upload.getbuffer())
else:
    qual_pdf_path = Path(DEFAULT_QUAL_PDF)

if part_pdf_upload:
    part_pdf_path = Path("participated_uploaded.pdf")
    with open(part_pdf_path, "wb") as f:
        f.write(part_pdf_upload.getbuffer())
else:
    part_pdf_path = Path(DEFAULT_PART_PDF)

if ttf_upload:
    ttf_path = Path("uploaded_times_italic.ttf")
    with open(ttf_path, "wb") as f:
        f.write(ttf_upload.getbuffer())
elif Path(DEFAULT_TTF).exists():
    ttf_path = Path(DEFAULT_TTF)
else:
    ttf_path = None

# Register font if TTF present, else fallback to Times-Italic
registered_font_name = None
if ttf_path and ttf_path.exists():
    try:
        pdfmetrics.registerFont(TTFont("TimesNewRomanItalicUser", str(ttf_path)))
        registered_font_name = "TimesNewRomanItalicUser"
        st.sidebar.success(f"Registered font from {ttf_path.name}")
    except Exception as e:
        st.sidebar.warning(f"Failed to register TTF: {e}. Using fallback.")
        registered_font_name = font_choice
else:
    registered_font_name = font_choice

# Generate button
if st.button("Generate certificates (Overlay)"):
    if excel_file is None:
        st.error("Please upload the Excel file with QUALIFIED and PARTICIPATED sheets.")
        st.stop()

    # Read Excel
    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    # case-insensitive sheet name mapping
    sheets_map = {name.strip().upper(): name for name in xls.sheet_names}
    required = ["QUALIFIED", "PARTICIPATED"]
    missing = [s for s in required if s not in sheets_map]
    if missing:
        st.error(f"Missing sheet(s): {missing}. Please ensure the Excel contains sheets named QUALIFIED and PARTICIPATED.")
        st.stop()

    progress = st.progress(0)
    generated_files = []
    total_steps = 2
    step = 0

    # Process QUALIFIED
    st.info("Processing QUALIFIED sheet...")
    q_actual = sheets_map["QUALIFIED"]
    df_q = pd.read_excel(xls, sheet_name=q_actual, dtype=object)
    name_col_q = get_first_name_column(df_q)
    if not name_col_q or df_q[name_col_q].dropna().empty:
        st.warning("QUALIFIED sheet is empty or has no names - skipping.")
    else:
        names_q = df_q[name_col_q].dropna().astype(str).tolist()
        out_q = Path(OUTPUT_DIR) / "QUALIFIED.pdf"
        try:
            overlay_template_with_names(qual_pdf_path, names_q, out_q, x_cm, y_cm_qual, font_size, registered_font_name)
            generated_files.append(out_q)
            st.success(f"QUALIFIED PDF created ({len(names_q)} names).")
        except Exception as e:
            st.error(f"Failed QUALIFIED generation: {e}")

    step += 1
    progress.progress(int(step/total_steps*100))

    # Process PARTICIPATED
    st.info("Processing PARTICIPATED sheet...")
    p_actual = sheets_map["PARTICIPATED"]
    df_p = pd.read_excel(xls, sheet_name=p_actual, dtype=object)
    name_col_p = get_first_name_column(df_p)
    if not name_col_p or df_p[name_col_p].dropna().empty:
        st.warning("PARTICIPATED sheet is empty or has no names - skipping.")
    else:
        names_p = df_p[name_col_p].dropna().astype(str).tolist()
        out_p = Path(OUTPUT_DIR) / "PARTICIPATED.pdf"
        try:
            overlay_template_with_names(part_pdf_path, names_p, out_p, x_cm, y_cm_part, font_size, registered_font_name)
            generated_files.append(out_p)
            st.success(f"PARTICIPATED PDF created ({len(names_p)} names).")
        except Exception as e:
            st.error(f"Failed PARTICIPATED generation: {e}")

    step += 1
    progress.progress(int(step/total_steps*100))

    if not generated_files:
        st.error("No PDF files were generated.")
        st.stop()

    # create ZIP in-memory
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in generated_files:
            zf.write(p, p.name)
    zip_buffer.seek(0)

    # optional: cleanup generated pdfs on disk
    for p in generated_files:
        try:
            os.remove(p)
        except:
            pass

    st.success("Certificates generated. Download the ZIP below.")
    st.download_button("Download certificates ZIP", data=zip_buffer.getvalue(),
                       file_name="certificates_two_sheets.zip", mime="application/zip")
