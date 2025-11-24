# app.py
import streamlit as st
import pandas as pd
import fitz  # pymupdf
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
# CONSTANTS & DEFAULTS
# --------------------------
DEFAULT_FONT_FILE = "Times New Roman Italic.ttf"
FONT_PATH = Path(DEFAULT_FONT_FILE)

DEFAULT_X_CM = 10.46
DEFAULT_Y_CM = 16.50
DEFAULT_FONT_PT = 19
DEFAULT_MAX_WIDTH_CM = 16.0
DPI = 300

# Default template file names in repo (change if different)
DEFAULT_QUALIFIED_PDF = Path("phnscholar qualified certificate.pdf")
DEFAULT_PARTICIPATED_PDF = Path("phnscholar participation certificate.pdf")
DEFAULT_SMARTEDGE_PDF = Path("smart edge workshop certificate.pdf")

# fallback preview in deployment (optional)
FALLBACK_PREVIEW = Path("/mnt/data/PARTICIPATED_preview.pdf")

def cm_to_px(cm, dpi=DPI):
    return int((cm / 2.54) * dpi)

# --------------------------
# FUNNY / HELPFUL MESSAGES
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
# PDF -> Image preview helper
# --------------------------
def render_pdf_first_page_to_image(pdf_bytes: bytes, dpi=DPI) -> Image.Image:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

# --------------------------
# Draw name onto template (rasterize)
# --------------------------
def draw_name_on_template(template_bytes, name, x_cm, y_cm, font_size_pt, max_width_cm):
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
                new_font_px = max(8, int(font.size * scale))
                font = ImageFont.truetype(str(FONT_PATH), new_font_px)
                bbox = draw.textbbox((0, 0), name, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
        except Exception:
            pass

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # outline and fill
    outline_color = "white"
    fill_color = "black"
    for dx, dy in [(-1,0),(1,0),(0,-1),(0,1)]:
        draw.text((draw_x+dx, draw_y+dy), name, font=font, fill=outline_color)
    draw.text((draw_x, draw_y), name, font=font, fill=fill_color)

    return img

def image_to_pdf_bytes(img: Image.Image):
    out = io.BytesIO()
    img.convert("RGB").save(out, format="PDF")
    return out.getvalue()

# --------------------------
# UI: Header + centered small logo (no raw HTML)
# --------------------------
logo_path = Path("logo.png")
if logo_path.exists():
    try:
        img_logo = Image.open(logo_path)
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            # small width keeps logo small and centered column centers it
            st.image(img_logo, width=64, use_column_width=False)
    except Exception as e:
        st.warning(f"Logo found but could not be displayed: {e}")

st.markdown("<h1 style='text-align:center;'>Certificate Generator — QUALIFIED, PARTICIPATED & SMART EDGE WORKSHOP</h1>", unsafe_allow_html=True)

# --------------------------
# UPLOADS & SETTINGS (UI)
# --------------------------
st.header("1) Upload files (Excel must contain sheets QUALIFIED, PARTICIPATED and Smart Edge sheet if used)")

excel_file = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx", "xls"])
qualified_pdf_file = st.file_uploader("Qualified template PDF (optional)", type=["pdf"])
participated_pdf_file = st.file_uploader("Participated template PDF (optional)", type=["pdf"])
smartedge_pdf_file = st.file_uploader("SMART EDGE template PDF (optional)", type=["pdf"])
ttf_upload = st.file_uploader("Times New Roman Italic TTF (optional)", type=["ttf","otf"])
site_logo_upload = st.file_uploader("Logo for website header only (optional)", type=["png","jpg","jpeg"])

# allow overriding font
if ttf_upload:
    with open("uploaded_times.ttf", "wb") as f:
        f.write(ttf_upload.getbuffer())
    FONT_PATH = Path("uploaded_times.ttf")

# --------------------------
# Sidebar controls (force numeric types explicitly to avoid mixed-type bug)
# --------------------------
st.sidebar.header("Rasterize output (recommended)")
rasterize = st.sidebar.checkbox("Rasterize certificates", value=True)

st.sidebar.header("Position & font settings")
X_CM = float(st.sidebar.number_input("X (cm from left)", value=float(DEFAULT_X_CM), format="%.2f", step=0.01))
Y_CM = float(st.sidebar.number_input("Y (cm from bottom)", value=float(DEFAULT_Y_CM), format="%.2f", step=0.01))
BASE_FONT_PT = int(st.sidebar.number_input("Base font size (pt)", value=int(DEFAULT_FONT_PT), step=1))
MAX_WIDTH_CM = float(st.sidebar.number_input("Max name width (cm)", value=float(DEFAULT_MAX_WIDTH_CM), step=0.5))

# --------------------------
# LIVE PREVIEW area
# --------------------------
st.header("2) Live preview (first page)")

def get_template_bytes(uploader_obj, default_path: Path):
    if uploader_obj is not None:
        try:
            return uploader_obj.read()
        except Exception:
            return None
    if default_path.exists():
        try:
            return default_path.read_bytes()
        except Exception:
            return None
    if FALLBACK_PREVIEW.exists():
        try:
            return FALLBACK_PREVIEW.read_bytes()
        except Exception:
            return None
    return None

colq, colp, cols = st.columns([1,1,1])
with colq:
    st.subheader("QUALIFIED")
    qual_bytes_preview = get_template_bytes(qualified_pdf_file, DEFAULT_QUALIFIED_PDF)
    if qual_bytes_preview:
        try:
            st.image(render_pdf_first_page_to_image(qual_bytes_preview), caption="QUALIFIED — first page", use_column_width=True)
        except Exception as e:
            st.error(f"Preview error: {e}")
    else:
        st.info("No QUALIFIED template uploaded / found.")

with colp:
    st.subheader("PARTICIPATED")
    part_bytes_preview = get_template_bytes(participated_pdf_file, DEFAULT_PARTICIPATED_PDF)
    if part_bytes_preview:
        try:
            st.image(render_pdf_first_page_to_image(part_bytes_preview), caption="PARTICIPATED — first page", use_column_width=True)
        except Exception as e:
            st.error(f"Preview error: {e}")
    else:
        st.info("No PARTICIPATED template uploaded / found.")

with cols:
    st.subheader("SMART EDGE")
    smart_bytes_preview = get_template_bytes(smartedge_pdf_file, DEFAULT_SMARTEDGE_PDF)
    if smart_bytes_preview:
        try:
            st.image(render_pdf_first_page_to_image(smart_bytes_preview), caption="SMART EDGE — first page", use_column_width=True)
        except Exception as e:
            st.error(f"Preview error: {e}")
    else:
        st.info("No SMART EDGE template uploaded / found.")

# --------------------------
# CHECKBOXES & centered caption
# --------------------------
st.markdown("### 3) Generate and download final ZIP")
st.write("Export options:")

col1, col2, col3 = st.columns([1,1,1])
with col1:
    gen_qualified = st.checkbox("Generate QUALIFIED", value=False)
with col2:
    gen_participated = st.checkbox("Generate PARTICIPATED", value=False)
with col3:
    gen_smartedge = st.checkbox("Generate SMART EDGE WORKSHOP", value=False)

st.markdown("<div style='text-align:center; opacity:0.75; margin-top:20px; margin-bottom:16px;'>"
            "Select which certificates to include in the ZIP. Uncheck to exclude a group."
            "</div>", unsafe_allow_html=True)

# spacing before button
st.write("")
st.write("")

# --------------------------
# Helpers for sheet detection & safe reading
# --------------------------
def find_sheet_name(xls: pd.ExcelFile, target_upper: str):
    """
    Return original sheet name whose upper() matches target_upper, else None.
    """
    for s in xls.sheet_names:
        if s.strip().upper() == target_upper:
            return s
    return None

def read_excel_safe(excel_io, xls_obj: pd.ExcelFile, target_upper: str):
    """
    If a matching sheet exists (case-insensitive) return dataframe, else return empty DataFrame.
    """
    name = find_sheet_name(xls_obj, target_upper)
    if name is None:
        return pd.DataFrame(), None
    try:
        df = pd.read_excel(excel_io, sheet_name=name, dtype=object)
        return df, name
    except Exception as e:
        # bubble up a descriptive exception
        raise ValueError(f"Failed to read sheet '{name}': {e}")

# --------------------------
# GENERATION
# --------------------------
if st.button("Generate certificates ZIP"):

    # check selections
    if not (gen_qualified or gen_participated or gen_smartedge):
        st.error(random.choice(FUNNY_ERRORS))
        st.stop()

    if excel_file is None:
        st.error("Please upload the Excel file first.")
        st.stop()

    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Cannot parse Excel file: {e}")
        st.stop()

    # Find smartedge sheet if any allowed names present
    smartedge_allowed = {"NAMES", "NAME", "SMART EDGE", "CERTIFICATES"}
    smartedge_sheet_found = None
    for s in xls.sheet_names:
        if s.strip().upper() in smartedge_allowed:
            smartedge_sheet_found = s
            break

    # read templates (uploaded -> repo default -> error)
    def get_template_or_none(uploader_obj, default_path: Path):
        b = get_template_bytes(uploader_obj, default_path)
        return b

    qual_bytes = None
    part_bytes = None
    smart_bytes = None

    if gen_qualified:
        qual_bytes = get_template_or_none(qualified_pdf_file, DEFAULT_QUALIFIED_PDF)
        if qual_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Qualified)")
            st.stop()

    if gen_participated:
        part_bytes = get_template_or_none(participated_pdf_file, DEFAULT_PARTICIPATED_PDF)
        if part_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Participated)")
            st.stop()

    if gen_smartedge:
        smart_bytes = get_template_or_none(smartedge_pdf_file, DEFAULT_SMARTEDGE_PDF)
        if smart_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Smart Edge)")
            st.stop()
        if smartedge_sheet_found is None:
            st.error(random.choice(MISSING_SHEET_ERRORS) + " (Smart Edge sheet must be one of: Names / Name / Smart Edge / Certificates)")
            st.stop()

    # Now read Excel sheets safely (case-insensitive)
    # QUALIFIED
    try:
        df_q, qual_sheet_name = read_excel_safe(excel_file, xls, "QUALIFIED")
    except ValueError as e:
        st.error(str(e))
        st.stop()

    # PARTICIPATED
    try:
        df_p, part_sheet_name = read_excel_safe(excel_file, xls, "PARTICIPATED")
    except ValueError as e:
        st.error(str(e))
        st.stop()

    # SMART EDGE: use detected original sheet name if present
    if smartedge_sheet_found:
        try:
            df_s = pd.read_excel(excel_file, sheet_name=smartedge_sheet_found, dtype=object)
        except Exception as e:
            st.error(f"Failed to read Smart Edge sheet '{smartedge_sheet_found}': {e}")
            st.stop()
    else:
        df_s = pd.DataFrame()

    # Build tasks
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

    # UI progress placeholders
    total = len(tasks)
    overall_progress = st.progress(0)
    overall_status = st.empty()

    group_done = {g: 0 for g in group_counts}
    group_status_placeholders = {}
    group_progress_bars = {}
    for g, cnt in group_counts.items():
        if cnt > 0:
            group_status_placeholders[g] = st.empty()
            group_progress_bars[g] = st.progress(0.0)

    zip_buf = io.BytesIO()
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

            # slight pause so UI updates smoothly
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
                safe_name = str(name).replace("/", "_").replace("\\", "_")
                zf.writestr(f"{group}/{safe_name}.pdf", pdf_bytes)
            except Exception as e:
                err_msg = f"Failed to generate for {name}: {e}"
                safe_name = str(name).replace("/", "_").replace("\\", "_")
                zf.writestr(f"{group}/{safe_name}_ERROR.txt", err_msg.encode("utf-8"))

            overall_progress.progress(idx / total)

        overall_status.success("All items processed. Finalizing ZIP...")

    st.balloons()
    st.success(f"Done — {total} certificates generated (errors, if any, are in ZIP).")
    zip_buf.seek(0)
    st.download_button("Download certificates ZIP", data=zip_buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
