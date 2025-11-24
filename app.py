# app.py
# Safer certificate generator: heavy imports are lazy-loaded inside functions to avoid import-time crashes.
import streamlit as st
import pandas as pd
import io
import time
from zipfile import ZipFile
from pathlib import Path

st.set_page_config(page_title="Certificate Generator", layout="wide")

# ---------- config ----------
LOGO_PATH = Path("logo.png")
LOGO_WIDTH_PX = 64
DPI = 300
DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19
DEFAULT_MAX_WIDTH_CM = 16.0

# ---------- small helpers ----------
def cm_to_px(cm, dpi=DPI):
    return int((cm / 2.54) * dpi)

def get_bytes_from_uploader_or_default(uploader, default_path: Path):
    if uploader is not None:
        try:
            return uploader.read()
        except Exception:
            return None
    if default_path.exists():
        try:
            return default_path.read_bytes()
        except Exception:
            return None
    return None

def safe_list_sheets(excel_file):
    try:
        xls = pd.ExcelFile(excel_file)
        return xls.sheet_names
    except Exception as e:
        return []

# Lazy imports for heavy libs (pymupdf, PIL) so app starts even if something goes wrong
def render_pdf_first_page_to_image(pdf_bytes: bytes, dpi=DPI):
    # lazy imports
    from PIL import Image
    import fitz  # pymupdf
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def image_to_pdf_bytes(img):
    from io import BytesIO
    out = BytesIO()
    img.convert("RGB").save(out, format="PDF")
    return out.getvalue()

def draw_name_on_template(template_bytes, name, x_cm, y_cm, font_size_pt, max_width_cm, font_path=None):
    # lazy imports
    from PIL import Image, ImageDraw, ImageFont
    import fitz  # pymupdf

    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=DPI)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)

    # font handling (best-effort)
    try:
        if font_path and Path(font_path).exists():
            font_px = max(8, int(round(font_size_pt * DPI / 72.0)))
            font = ImageFont.truetype(str(font_path), font_px)
        else:
            # fallback to a reasonably sized default font if available
            font_px = max(8, int(round(font_size_pt * DPI / 72.0)))
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", font_px)
            except Exception:
                font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()

    x_px = cm_to_px(x_cm)
    y_px_from_bottom = cm_to_px(y_cm)
    y_px = img.height - y_px_from_bottom

    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except Exception:
        text_w, text_h = draw.textsize(name, font=font)

    max_w_px = cm_to_px(max_width_cm)
    if text_w > max_w_px:
        # try scaling font down (best-effort)
        try:
            scale = max_w_px / text_w
            new_font_px = max(8, int(getattr(font, "size", font_px) * scale))
            if isinstance(font, ImageFont.FreeTypeFont):
                font = ImageFont.truetype(font.path, new_font_px)
            else:
                # fallback - keep same font
                pass
            bbox = draw.textbbox((0, 0), name, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
        except Exception:
            pass

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # draw small outline for contrast
    outline_color = "white"
    fill_color = "black"
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill=outline_color)
    draw.text((draw_x, draw_y), name, font=font, fill=fill_color)

    return img

# ---------- UI header & logo ----------
try:
    if LOGO_PATH.exists():
        from PIL import Image
        img_logo = Image.open(LOGO_PATH)
        c1, c2, c3 = st.columns([1, 0.3, 1])
        with c2:
            st.image(img_logo, width=LOGO_WIDTH_PX, use_container_width=False)
except Exception:
    # don't block UI if logo fails
    pass

st.markdown("<h1 style='text-align:center;'>Certificate Generator</h1>", unsafe_allow_html=True)
st.write("Upload Excel + PDF templates and press **Generate certificates ZIP**. If a worksheet name is missing the app will warn instead of crashing.")

# ---------- upload controls ----------
st.header("Upload files")
excel_file = st.file_uploader("Excel (.xlsx/.xls)", type=["xlsx","xls"])
qualified_pdf_file = st.file_uploader("Qualified template PDF (optional)", type=["pdf"])
participated_pdf_file = st.file_uploader("Participated template PDF (optional)", type=["pdf"])
smartedge_pdf_file = st.file_uploader("Smart Edge template PDF (optional)", type=["pdf"])
ttf_upload = st.file_uploader("Custom TTF font (optional)", type=["ttf","otf"])

uploaded_font_path = None
if ttf_upload:
    uploaded_font_path = "uploaded_font.ttf"
    try:
        with open(uploaded_font_path, "wb") as f:
            f.write(ttf_upload.getbuffer())
    except Exception:
        st.warning("Could not save uploaded font; fallback to defaults will be used.")

st.header("Live preview (first page)")
colq, colp, cols = st.columns(3)
with colq:
    st.subheader("QUALIFIED")
    try:
        qual_bytes_preview = get_bytes_from_uploader_or_default(qualified_pdf_file, Path("phnscholar qualified certificate.pdf"))
        if qual_bytes_preview:
            st.image(render_pdf_first_page_to_image(qual_bytes_preview), caption="QUALIFIED — first page", use_container_width=True)
        else:
            st.info("No QUALIFIED template uploaded / found.")
    except Exception as e:
        st.error("Preview error for QUALIFIED.")

with colp:
    st.subheader("PARTICIPATED")
    try:
        part_bytes_preview = get_bytes_from_uploader_or_default(participated_pdf_file, Path("phnscholar participation certificate.pdf"))
        if part_bytes_preview:
            st.image(render_pdf_first_page_to_image(part_bytes_preview), caption="PARTICIPATED — first page", use_container_width=True)
        else:
            st.info("No PARTICIPATED template uploaded / found.")
    except Exception:
        st.error("Preview error for PARTICIPATED.")

with cols:
    st.subheader("SMART EDGE")
    try:
        smart_bytes_preview = get_bytes_from_uploader_or_default(smartedge_pdf_file, Path("smart edge workshop certificate.pdf"))
        if smart_bytes_preview:
            st.image(render_pdf_first_page_to_image(smart_bytes_preview), caption="SMART EDGE — first page", use_container_width=True)
        else:
            st.info("No SMART EDGE template uploaded / found.")
    except Exception:
        st.error("Preview error for SMART EDGE.")

# ---------- options ----------
st.markdown("### Export options")
col1, col2, col3 = st.columns(3)
with col1:
    gen_qualified = st.checkbox("Generate QUALIFIED", value=False)
with col2:
    gen_participated = st.checkbox("Generate PARTICIPATED", value=False)
with col3:
    gen_smartedge = st.checkbox("Generate SMART EDGE", value=False)

st.sidebar.header("Position & font settings")
X_CM = float(st.sidebar.number_input("X (cm from left)", value=DEFAULT_X_CM))
Y_CM = float(st.sidebar.number_input("Y (cm from bottom)", value=DEFAULT_Y_CM))
BASE_FONT_PT = int(st.sidebar.number_input("Base font size (pt)", value=DEFAULT_FONT_PT))
MAX_WIDTH_CM = float(st.sidebar.number_input("Max name width (cm)", value=DEFAULT_MAX_WIDTH_CM))

if excel_file is not None:
    sheets = safe_list_sheets(excel_file)
    if sheets:
        st.info(f"Workbook sheets: {', '.join(sheets)}")
    else:
        st.warning("Could not read workbook sheets list.")

# ---------- run generation (button only) ----------
if st.button("Generate certificates ZIP"):
    # defensive checks
    if not (gen_qualified or gen_participated or gen_smartedge):
        st.error("Select at least one certificate type.")
        st.stop()

    if excel_file is None:
        st.error("Please upload Excel file.")
        st.stop()

    # try to parse workbook safely
    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Cannot open Excel file: {e}")
        st.stop()

    def read_sheet_if_exists(excel_io, xls_obj, wanted_upper):
        match = None
        for s in xls_obj.sheet_names:
            if s.strip().upper() == wanted_upper:
                match = s
                break
        if match is None:
            return pd.DataFrame(), None
        try:
            df = pd.read_excel(excel_io, sheet_name=match, dtype=object)
            return df, match
        except Exception as e:
            return pd.DataFrame(), match

    df_q, used_q = read_sheet_if_exists(excel_file, xls, "QUALIFIED")
    df_p, used_p = read_sheet_if_exists(excel_file, xls, "PARTICIPATED")

    # smart-edge tolerant names
    smart_candidates = {"SMART EDGE", "SMARTEDGE", "SMART_EDGE", "NAMES", "NAME", "CERTIFICATES", "PARTICIPANTS"}
    df_s = pd.DataFrame(); used_s = None
    for s in xls.sheet_names:
        if s.strip().upper() in smart_candidates:
            try:
                df_s = pd.read_excel(excel_file, sheet_name=s, dtype=object)
                used_s = s
                break
            except Exception:
                df_s = pd.DataFrame()
                used_s = s
                break

    # warn if requested but missing
    if gen_qualified and df_q.empty:
        st.warning(f"QUALIFIED requested but sheet not found or empty. Sheets: {', '.join(xls.sheet_names)}")
    if gen_participated and df_p.empty:
        st.warning(f"PARTICIPATED requested but sheet not found or empty. Sheets: {', '.join(xls.sheet_names)}")
    if gen_smartedge and df_s.empty:
        st.warning(f"SMART EDGE requested but matching sheet not found or empty. Sheets: {', '.join(xls.sheet_names)}")

    # collect tasks
    tasks = []
    if gen_qualified and not df_q.empty:
        names_q = df_q.iloc[:,0].dropna().astype(str).tolist()
        tasks += [("QUALIFIED", n.strip()) for n in names_q]
    if gen_participated and not df_p.empty:
        names_p = df_p.iloc[:,0].dropna().astype(str).tolist()
        tasks += [("PARTICIPATED", n.strip()) for n in names_p]
    if gen_smartedge and not df_s.empty:
        names_s = df_s.iloc[:,0].dropna().astype(str).tolist()
        tasks += [("SMART_EDGE", n.strip()) for n in names_s]

    if len(tasks) == 0:
        st.error("No names to process (check sheets).")
        st.stop()

    # templates (must exist if generation requested)
    qual_tpl = get_bytes_from_uploader_or_default(qualified_pdf_file, Path("phnscholar qualified certificate.pdf")) if gen_qualified else None
    part_tpl = get_bytes_from_uploader_or_default(participated_pdf_file, Path("phnscholar participation certificate.pdf")) if gen_participated else None
    smart_tpl = get_bytes_from_uploader_or_default(smartedge_pdf_file, Path("smart edge workshop certificate.pdf")) if gen_smartedge else None

    missing_templates = []
    if gen_qualified and not qual_tpl:
        missing_templates.append("QUALIFIED")
    if gen_participated and not part_tpl:
        missing_templates.append("PARTICIPATED")
    if gen_smartedge and not smart_tpl:
        missing_templates.append("SMART EDGE")
    if missing_templates:
        st.error(f"Missing templates for: {', '.join(missing_templates)}. Upload PDFs or add to repo.")
        st.stop()

    # generate zip
    zip_buf = io.BytesIO()
    total = len(tasks)
    prog = st.progress(0.0)
    status = st.empty()

    try:
        with ZipFile(zip_buf, "w") as zf:
            for idx, (group, name) in enumerate(tasks, start=1):
                status.info(f"Generating {group} — {name} ({idx}/{total})")
                time.sleep(0.02)  # allow UI update but keep small
                try:
                    tpl = qual_tpl if group == "QUALIFIED" else (part_tpl if group == "PARTICIPATED" else smart_tpl)
                    img = draw_name_on_template(tpl, name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM, font_path=uploaded_font_path)
                    pdf_bytes = image_to_pdf_bytes(img)
                    safe_name = name.replace("/", "_").replace("\\", "_")
                    zf.writestr(f"{group}/{safe_name}.pdf", pdf_bytes)
                except Exception as e:
                    zf.writestr(f"{group}/{name}_ERROR.txt", str(e).encode("utf-8"))
                # update progress less frequently for performance safety
                if idx % 5 == 0 or idx == total:
                    prog.progress(min(1.0, idx/total))
    except Exception as gen_e:
        st.error(f"Generation failed: {gen_e}")
        st.stop()

    zip_buf.seek(0)
    st.success("Certificates generated.")
    st.download_button("Download ZIP", data=zip_buf.getvalue(), file_name="certificates.zip", mime="application/zip")
