# app.py
import streamlit as st
from pathlib import Path
from io import BytesIO
import zipfile
import pandas as pd
import math
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

# ---------- Basic setup ----------
st.set_page_config(page_title="Certificate Generator", layout="wide")
# define BASE_DIR early so Streamlit Cloud pathing works
BASE_DIR = Path(__file__).parent

# ---------- Defaults ----------
DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_SIZE = 16
DEFAULT_MAX_TEXT_WIDTH_CM = 16.0

DEFAULT_QUAL = BASE_DIR / "phnscholar qualified certificate.pdf"
DEFAULT_PART = BASE_DIR / "phnscholar participation certificate.pdf"
DEFAULT_TTF = BASE_DIR / "Times New Roman Italic.ttf"
DEFAULT_LOGO = BASE_DIR / "logo.png"

OUT_DIR = BASE_DIR / "output"
OUT_DIR.mkdir(exist_ok=True)

# ---------- Helper functions ----------
def register_ttf_if_present(ttf_path: Path):
    if ttf_path and ttf_path.exists():
        try:
            pdfmetrics.registerFont(TTFont("UserTime", str(ttf_path)))
            return "UserTime"
        except Exception:
            return "Times-Italic"
    return "Times-Italic"

def get_first_name_column(df):
    for col in df.columns:
        if df[col].notna().any():
            return col
    return None

def scaled_font_size(name, font_name, desired_size_pt, max_width_cm, min_font_size=8):
    try:
        width_pt = pdfmetrics.stringWidth(str(name), font_name, desired_size_pt)
    except Exception:
        # fallback width estimate if font not registered for PDF text
        width_pt = len(str(name)) * desired_size_pt * 0.5
    max_width_pt = max_width_cm * cm
    if width_pt <= max_width_pt:
        return desired_size_pt
    scale = max_width_pt / width_pt
    new_size = max(min_font_size, int(desired_size_pt * scale))
    return new_size

def make_overlay_pdf_bytes(name, page_w_pt, page_h_pt, x_cm, y_cm, font_name, font_size_pt):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_w_pt, page_h_pt))
    c.setFont(font_name, font_size_pt)
    x_pt = x_cm * cm
    y_pt = y_cm * cm
    c.drawCentredString(x_pt, y_pt, str(name).strip())
    c.save()
    packet.seek(0)
    return packet

def merge_overlay(template_path: Path, overlay_buf: BytesIO):
    reader_fresh = PdfReader(str(template_path))
    overlay_reader = PdfReader(overlay_buf)
    base = reader_fresh.pages[0]
    base.merge_page(overlay_reader.pages[0])
    writer = PdfWriter()
    writer.add_page(base)
    out = BytesIO()
    writer.write(out)
    out.seek(0)
    return out

def render_pdf_to_png_bytes(pdf_bytes: bytes, dpi=150):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")

# Raster utilities
def render_page_to_image(pdf_path: Path, dpi=300):
    doc = fitz.open(str(pdf_path))
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    mediabox = page.mediabox
    page_w_pt = float(mediabox.width)
    page_h_pt = float(mediabox.height)
    return img, page_w_pt, page_h_pt

def draw_name_on_image(img: Image.Image, name: str, x_cm_val, y_cm_val, page_w_pt, page_h_pt,
                       font_path=None, font_size_pt=16, dpi=300, logo_path: Path = None):
    # px per point mapping
    px_per_pt = img.width / page_w_pt
    # convert cm -> points (1 cm = 28.3464567 pt)
    x_pt = x_cm_val * 28.3464567
    y_pt = y_cm_val * 28.3464567
    x_px = x_pt * px_per_pt
    y_px_from_bottom = y_pt * px_per_pt
    y_px = img.height - y_px_from_bottom

    draw = ImageDraw.Draw(img)

    # load font
    try:
        if font_path and Path(font_path).exists():
            font_px = int(round(font_size_pt * dpi / 72.0))
            font = ImageFont.truetype(str(font_path), size=font_px)
        else:
            font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()

    # measure with bbox for Pillow >= 10
    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except Exception:
        text_w, text_h = font.getsize(name)

    # autoscale if too wide
    max_w = img.width * 0.75
    if text_w > max_w:
        scale = max_w / text_w
        new_font_px = max(8, int((font.size if hasattr(font, "size") else 16) * scale))
        try:
            if font_path and Path(font_path).exists():
                font = ImageFont.truetype(str(font_path), size=new_font_px)
            else:
                font = ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()
        try:
            bbox = draw.textbbox((0, 0), name, font=font)
            text_w = bbox[2] - bbox[0]; text_h = bbox[3] - bbox[1]
        except Exception:
            text_w, text_h = font.getsize(name)

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # draw outline + fill for visibility
    outline = "white"; fill = "black"
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill=outline)
    draw.text((draw_x, draw_y), name, font=font, fill=fill)

    # draw logo if provided (centered above the name)
    if logo_path and Path(logo_path).exists():
        try:
            logo_img = Image.open(str(logo_path)).convert("RGBA")
            # scale logo to 20% of page width by default
            max_logo_width = int(img.width * 0.20)
            ratio = logo_img.width / logo_img.height if logo_img.height != 0 else 1
            logo_img = logo_img.resize((max_logo_width, int(max_logo_width / ratio)))
            logo_x = int((img.width - logo_img.width) / 2)
            # place above name with some gap
            logo_y = int(draw_y - logo_img.height - (0.5 * text_h))
            if logo_y < 10:
                logo_y = 10
            img.paste(logo_img, (logo_x, logo_y), logo_img)
        except Exception:
            pass

    return img

def image_to_pdf_bytes(img: Image.Image):
    out = BytesIO()
    img_rgb = img.convert("RGB")
    img_rgb.save(out, format="PDF")
    out.seek(0)
    return out.read()

# ---------- UI: Header logo (centered) ----------
# Show logo from repo root if present
logo_path = None
if DEFAULT_LOGO.exists():
    logo_path = DEFAULT_LOGO

# centered header logo
if logo_path and logo_path.exists():
    col1, col2, col3 = st.columns([1, 0.5, 1])
    with col2:
        st.image(str(logo_path), width=150)
# title after logo
st.title("Certificate Generator — QUALIFIED & PARTICIPATED")

# ---------- UI: Uploads ----------
st.markdown("### 1) Upload files (Excel must contain sheets QUALIFIED & PARTICIPATED)")

excel_file = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"])
qual_upload = st.file_uploader("Qualified template PDF (optional)", type=["pdf"])
part_upload = st.file_uploader("Participated template PDF (optional)", type=["pdf"])
ttf_upload = st.file_uploader("Times New Roman Italic TTF (optional)", type=["ttf","otf"])
logo_upload = st.file_uploader("Logo (optional) — used on site header & in PDF", type=["png","jpg","jpeg"])

# save uploaded files to disk if provided
if qual_upload:
    qual_path = BASE_DIR / "qual_uploaded.pdf"
    qual_path.write_bytes(qual_upload.getbuffer())
else:
    qual_path = DEFAULT_QUAL if DEFAULT_QUAL.exists() else None

if part_upload:
    part_path = BASE_DIR / "part_uploaded.pdf"
    part_path.write_bytes(part_upload.getbuffer())
else:
    part_path = DEFAULT_PART if DEFAULT_PART.exists() else None

if ttf_upload:
    ttf_path = BASE_DIR / "uploaded_times.ttf"
    ttf_path.write_bytes(ttf_upload.getbuffer())
elif DEFAULT_TTF.exists():
    ttf_path = DEFAULT_TTF
else:
    ttf_path = None

if logo_upload:
    logo_path = BASE_DIR / "uploaded_logo.png"
    logo_path.write_bytes(logo_upload.getbuffer())

# ---------- UI: Position & font settings ----------
st.sidebar.header("Position & font settings")
x_cm = st.sidebar.number_input("X (cm from left)", value=DEFAULT_X_CM, format="%.2f", step=0.01)
y_cm = st.sidebar.number_input("Y (cm from bottom)", value=DEFAULT_Y_CM, format="%.2f", step=0.01)
font_size_pt = st.sidebar.number_input("Base font size (pt)", value=DEFAULT_FONT_SIZE, step=1)
max_text_width_cm = st.sidebar.number_input("Max name width (cm) for autoscale", value=DEFAULT_MAX_TEXT_WIDTH_CM, step=0.5)
rasterize = st.sidebar.checkbox("Rasterize output (recommended)", value=True)

# register font for vector overlay if possible
font_name = register_ttf_if_present(Path(ttf_path) if ttf_path else None)
st.sidebar.write("Font used:", font_name)

# ---------- UI: Preview ----------
st.markdown("### 2) Preview (live)")
preview_choice = st.selectbox("Preview template", options=["QUALIFIED", "PARTICIPATED"])
template_for_preview = qual_path if preview_choice == "QUALIFIED" else part_path

if template_for_preview and Path(template_for_preview).exists():
    sample_name = st.text_input("Sample name for preview", "Aarav Sharma")
    # determine preview mode
    if not rasterize:
        reader = PdfReader(str(template_for_preview))
        mediabox = reader.pages[0].mediabox
        page_w_pt = float(mediabox.width); page_h_pt = float(mediabox.height)
        fs_preview = scaled_font_size(sample_name, font_name, font_size_pt, max_text_width_cm)
        overlay_buf = make_overlay_pdf_bytes(sample_name, page_w_pt, page_h_pt, x_cm, y_cm, font_name, fs_preview)
        merged = merge_overlay(template_for_preview, overlay_buf)
        png = render_pdf_to_png_bytes(merged.read(), dpi=150)
        st.image(png, use_column_width=True, caption="Preview (vector overlay)")
    else:
        img, page_w_pt, page_h_pt = render_page_to_image(template_for_preview, dpi=300)
        img_preview = draw_name_on_image(img.copy(), sample_name, x_cm, y_cm, page_w_pt, page_h_pt,
                                        font_path=ttf_path, font_size_pt=font_size_pt, dpi=300, logo_path=logo_path)
        buf = BytesIO()
        img_preview.save(buf, format="PNG")
        st.image(buf.getvalue(), use_column_width=True, caption="Preview (raster)")
else:
    st.info("Template not found. Upload template or place default file in repo.")

# ---------- UI: Export options ----------
st.markdown("### 3) Generate and download final ZIP")
st.write("Export options:")
col_a, col_b, col_c = st.columns([1,1,6])
with col_a:
    gen_qualified = st.checkbox("Generate QUALIFIED", value=True)
with col_b:
    gen_participated = st.checkbox("Generate PARTICIPATED", value=True)
with col_c:
    st.caption("Select which certificates to include in the ZIP. Uncheck to exclude a group.")

# ---------- Generation logic ----------
if st.button("Generate certificates ZIP"):
    if not excel_file:
        st.error("Please upload Excel with QUALIFIED and PARTICIPATED sheets.")
        st.stop()
    if not (qual_path and Path(qual_path).exists()):
        st.error("Qualified template missing. Upload or place it in repo.")
        st.stop()
    if not (part_path and Path(part_path).exists()):
        st.error("Participated template missing. Upload or place it in repo.")
        st.stop()
    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Cannot read Excel: {e}")
        st.stop()

    sheets_map = {name.strip().upper(): name for name in xls.sheet_names}
    missing = [s for s in ("QUALIFIED","PARTICIPATED") if s not in sheets_map]
    if missing:
        st.error(f"Missing sheets in Excel: {missing}")
        st.stop()

    df_q = pd.read_excel(xls, sheet_name=sheets_map["QUALIFIED"], dtype=object)
    df_p = pd.read_excel(xls, sheet_name=sheets_map["PARTICIPATED"], dtype=object)

    # prepare zip buffer
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # QUALIFIED
        if gen_qualified:
            col_q = get_first_name_column(df_q)
            if col_q:
                for nm in df_q[col_q].dropna().astype(str):
                    nm_clean = str(nm).strip()
                    if not rasterize:
                        fs = scaled_font_size(nm_clean, font_name, font_size_pt, max_text_width_cm)
                        reader = PdfReader(str(qual_path))
                        mediabox = reader.pages[0].mediabox
                        page_w = float(mediabox.width); page_h = float(mediabox.height)
                        overlay_buf = make_overlay_pdf_bytes(nm_clean, page_w, page_h, x_cm, y_cm, font_name, fs)
                        merged = merge_overlay(qual_path, overlay_buf)
                        zf.writestr(f"QUALIFIED/{nm_clean}.pdf", merged.read())
                    else:
                        img, page_w_pt, page_h_pt = render_page_to_image(qual_path, dpi=300)
                        img_w_name = draw_name_on_image(img.copy(), nm_clean, x_cm, y_cm, page_w_pt, page_h_pt,
                                                        font_path=ttf_path, font_size_pt=font_size_pt, dpi=300, logo_path=logo_path)
                        pdf_bytes = image_to_pdf_bytes(img_w_name)
                        zf.writestr(f"QUALIFIED/{nm_clean}.pdf", pdf_bytes)

        # PARTICIPATED
        if gen_participated:
            col_p = get_first_name_column(df_p)
            if col_p:
                for nm in df_p[col_p].dropna().astype(str):
                    nm_clean = str(nm).strip()
                    if not rasterize:
                        fs = scaled_font_size(nm_clean, font_name, font_size_pt, max_text_width_cm)
                        reader = PdfReader(str(part_path))
                        mediabox = reader.pages[0].mediabox
                        page_w = float(mediabox.width); page_h = float(mediabox.height)
                        overlay_buf = make_overlay_pdf_bytes(nm_clean, page_w, page_h, x_cm, y_cm, font_name, fs)
                        merged = merge_overlay(part_path, overlay_buf)
                        zf.writestr(f"PARTICIPATED/{nm_clean}.pdf", merged.read())
                    else:
                        img, page_w_pt, page_h_pt = render_page_to_image(part_path, dpi=300)
                        img_w_name = draw_name_on_image(img.copy(), nm_clean, x_cm, y_cm, page_w_pt, page_h_pt,
                                                        font_path=ttf_path, font_size_pt=font_size_pt, dpi=300, logo_path=logo_path)
                        pdf_bytes = image_to_pdf_bytes(img_w_name)
                        zf.writestr(f"PARTICIPATED/{nm_clean}.pdf", pdf_bytes)

    zip_buf.seek(0)
    # build filename reflecting selection
    sel = []
    if gen_qualified: sel.append("qualified")
    if gen_participated: sel.append("participated")
    if not sel:
        st.error("No group selected. Please select at least one group.")
        st.stop()
    zip_name = "certificates_" + "_".join(sel) + ".zip"

    st.success("Certificates ZIP ready.")
    st.download_button("Download certificates ZIP", data=zip_buf.getvalue(), file_name=zip_name, mime="application/zip")
