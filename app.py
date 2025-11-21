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
from collections import Counter

# --------------------------
# PAGE CONFIG & SITE LOGO (site-only)
# --------------------------
st.set_page_config(page_title="Certificate Generator", layout="wide")
BASE_DIR = Path.cwd()

# Root/default templates that exist in your repo or uploaded test files used earlier in the convo.
# Dev note: I included some local /mnt/data paths that appeared in the conversation history.
DEFAULT_ROOT_TEMPLATES = {
    "QUALIFIED": BASE_DIR / "phnscholar qualified certificate.pdf",
    "PARTICIPATED": BASE_DIR / "phnscholar participation certificate.pdf",
    "SMART_EDGE": BASE_DIR / "smart edge workshop certificate.pdf",
    # fallback uploaded preview paths I saw in the conversation (used for local testing)
    "FALLBACK_1": Path("/mnt/data/d2d2db0d-3933-45cd-8123-39e8a6753858.pdf"),
    "FALLBACK_2": Path("/mnt/data/PARTICIPATED_preview.pdf"),
}

# site-only logo (center)
site_logo_path = BASE_DIR / "logo.png"
site_logo_uploaded = None  # will hold uploaded bytes if provided via uploader

# top logo (site-only, center)
if site_logo_path.exists():
    st.markdown(
        """
        <div style='text-align:center; margin-top:-10px; margin-bottom:6px;'>
            <img src="logo.png" style="width:140px; height:auto; display:block; margin-left:auto; margin-right:auto;">
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    "<h1 style='text-align:center; margin-bottom:4px;'>PHN ertificate Generator</h1>",
    unsafe_allow_html=True,
)

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

# helper conversions
def cm_to_px(cm, dpi=DPI):
    return int(round((cm / 2.54) * dpi))

# --------------------------
# FUNNY MESSAGES
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
    "Template not uploaded — certificates won’t dress themselves. 👔",
]

MISSING_SHEET_ERRORS = [
    "Excel missing the needed sheet. Did it go on vacation? 🏖️",
    "Required sheet not found. Please use the correct sheet name. 📄",
    "No matching sheet — try renaming it to Names / Name / Smart Edge / Certificates.",
]

# --------------------------
# PDF / IMAGE helpers
# --------------------------
def load_font(font_path: Path, font_pt: int):
    """Return a PIL ImageFont instance sized approximately for font_pt (pt)."""
    # convert pt -> px
    try:
        font_px = max(8, int(round(font_pt * DPI / 72.0)))
        if font_path.exists():
            return ImageFont.truetype(str(font_path), font_px)
    except Exception:
        pass
    return ImageFont.load_default()

def draw_name_on_template(template_bytes, name, x_cm, y_cm, font_size_pt, max_width_cm, font_path=FONT_PATH):
    """
    Render the first page of a PDF template to an image, write centered name at x_cm (from left)
    and y_cm (from bottom) and return PIL Image.
    """
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=DPI)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)

    # load font
    font = load_font(font_path, font_size_pt)

    # compute coordinates
    x_px = cm_to_px(x_cm)
    y_px_from_bottom = cm_to_px(y_cm)
    y_px = img.height - y_px_from_bottom

    # measure text
    try:
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    except Exception:
        text_w, text_h = draw.textsize(name, font=font)

    max_w_px = cm_to_px(max_width_cm)
    if text_w > max_w_px:
        # try to scale font down proportionally (only if truetype available)
        try:
            if hasattr(font, "size"):
                scale = max_w_px / text_w
                new_font_px = max(8, int(font.size * scale))
                if font_path.exists():
                    font = ImageFont.truetype(str(font_path), new_font_px)
                    bbox = draw.textbbox((0, 0), name, font=font)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
        except Exception:
            pass

    draw_x = int(round(x_px - text_w / 2.0))
    draw_y = int(round(y_px - text_h / 2.0))

    # outline (thin) and fill
    outline_color = "white"
    fill_color = "black"
    for dx, dy in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
        draw.text((draw_x + dx, draw_y + dy), name, font=font, fill=outline_color)
    draw.text((draw_x, draw_y), name, font=font, fill=fill_color)
    return img

def image_to_pdf_bytes(img: Image.Image):
    out = io.BytesIO()
    img.convert("RGB").save(out, format="PDF")
    return out.getvalue()

def read_template_bytes_from_path(path: Path):
    return path.read_bytes() if path.exists() else None

# --------------------------
# UI: Uploads & Settings
# --------------------------
st.header("1) Upload files (Excel must contain sheets QUALIFIED, PARTICIPATED and Smart Edge sheet if used)")

excel_file = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx", "xls"])
qualified_pdf_file = st.file_uploader("Qualified template PDF (optional — fallback used if not uploaded)", type=["pdf"])
participated_pdf_file = st.file_uploader("Participated template PDF (optional — fallback used if not uploaded)", type=["pdf"])
smartedge_pdf_file = st.file_uploader("SMART EDGE template PDF (optional — fallback used if not uploaded)", type=["pdf"])

ttf_upload = st.file_uploader("Times New Roman Italic TTF (optional)", type=["ttf", "otf"])
site_logo_upload = st.file_uploader("Logo for website header only (optional) — will NOT be printed on PDFs", type=["png", "jpg", "jpeg"])

# use uploaded TTF if provided
if ttf_upload is not None:
    with open("uploaded_times.ttf", "wb") as f:
        f.write(ttf_upload.getbuffer())
    FONT_PATH = Path("uploaded_times.ttf")

# use uploaded site logo if provided
if site_logo_upload is not None:
    site_logo_uploaded = site_logo_upload.getvalue()
    # write it so the markdown <img src="logo.png"> in this app can pick it up
    with open("logo.png", "wb") as f:
        f.write(site_logo_uploaded)
    # display centered logo
    st.markdown(
        """
        <div style='text-align:center; margin-top:-10px; margin-bottom:6px;'>
            <img src="logo.png" style="width:140px; height:auto; display:block; margin-left:auto; margin-right:auto;">
        </div>
        """,
        unsafe_allow_html=True,
    )

# --------------------------
# Sidebar settings (ensure same numeric types)
# --------------------------
st.sidebar.header("Rasterize output (recommended)")
rasterize = st.sidebar.checkbox("Rasterize certificates", value=True)

st.sidebar.header("Position & font settings")
# ensure we pass floats for cm fields and steps as floats so Streamlit doesn't mix numeric types
X_CM = float(st.sidebar.number_input("X (cm from left)", value=float(DEFAULT_X_CM), format="%.2f", step=0.01))
Y_CM = float(st.sidebar.number_input("Y (cm from bottom)", value=float(DEFAULT_Y_CM), format="%.2f", step=0.01))
BASE_FONT_PT = int(st.sidebar.number_input("Base font size (pt)", value=int(DEFAULT_FONT_PT), step=1))
MAX_WIDTH_CM = float(st.sidebar.number_input("Max name width (cm)", value=float(DEFAULT_MAX_WIDTH_CM), step=0.5))

# --------------------------
# CHECKBOXES & centered caption
# --------------------------
st.markdown("### 3) Generate and download final ZIP")
st.write("Export options:")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    gen_qualified = st.checkbox("Generate QUALIFIED", value=False)
with col2:
    gen_participated = st.checkbox("Generate PARTICIPATED", value=False)
with col3:
    gen_smartedge = st.checkbox("Generate SMART EDGE WORKSHOP", value=False)

st.markdown(
    "<div style='text-align:center; opacity:0.85; margin-top:12px;'>"
    "Select which certificates to include in the ZIP. Uncheck to exclude a group."
    "</div>",
    unsafe_allow_html=True,
)

# add spacing before button
st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

# --------------------------
# LIVE PREVIEW
# --------------------------
st.markdown("---")
st.header("Live preview (select a template and sample name)")
preview_col1, preview_col2 = st.columns([2, 1])

with preview_col1:
    preview_kind = st.selectbox("Choose template to preview", ["(none)", "QUALIFIED", "PARTICIPATED", "SMART EDGE"])
    sample_name = st.text_input("Sample name for preview", value="SAMPLE NAME")
    show_preview_btn = st.button("Show preview")

with preview_col2:
    st.write("Template preview:")
    tpl_preview_place = st.empty()

def get_template_bytes_for(kind):
    # prefer uploaded file; if not present, try root defaults and fallback paths
    if kind == "QUALIFIED":
        if qualified_pdf_file:
            return qualified_pdf_file.read()
        for p in (DEFAULT_ROOT_TEMPLATES["QUALIFIED"], DEFAULT_ROOT_TEMPLATES["FALLBACK_1"]):
            if p.exists():
                return p.read_bytes()
    elif kind == "PARTICIPATED":
        if participated_pdf_file:
            return participated_pdf_file.read()
        for p in (DEFAULT_ROOT_TEMPLATES["PARTICIPATED"], DEFAULT_ROOT_TEMPLATES["FALLBACK_2"]):
            if p.exists():
                return p.read_bytes()
    elif kind == "SMART EDGE":
        if smartedge_pdf_file:
            return smartedge_pdf_file.read()
        if DEFAULT_ROOT_TEMPLATES["SMART_EDGE"].exists():
            return DEFAULT_ROOT_TEMPLATES["SMART_EDGE"].read_bytes()
    return None

if show_preview_btn:
    if preview_kind == "(none)":
        tpl_preview_place.warning("Pick a template to preview.")
    else:
        tb = get_template_bytes_for(preview_kind)
        if tb is None:
            tpl_preview_place.error("Template not available (upload or put default in repo root).")
        else:
            try:
                img = draw_name_on_template(tb, sample_name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
                # display image (streamlit wants bytes or PIL image)
                tpl_preview_place.image(img, use_column_width=True)
            except Exception as e:
                tpl_preview_place.error(f"Preview failed: {e}")

st.markdown("---")

# --------------------------
# GENERATION
# --------------------------
if st.button("Generate certificates ZIP"):

    # funny error if nothing selected
    if not (gen_qualified or gen_participated or gen_smartedge):
        st.error(random.choice(FUNNY_ERRORS))
        st.stop()

    if excel_file is None:
        st.error(random.choice(MISSING_SHEET_ERRORS))
        st.stop()

    try:
        xls = pd.ExcelFile(excel_file)
    except Exception as e:
        st.error(f"Cannot read Excel: {e}")
        st.stop()

    # Smart Edge allowed sheet names (case-insensitive)
    smartedge_allowed = ["NAMES", "NAME", "SMART EDGE", "CERTIFICATES"]
    smartedge_sheet = None
    for s in xls.sheet_names:
        if s.strip().upper() in smartedge_allowed:
            smartedge_sheet = s
            break

    # find actual qualified/participated sheet names (prefix match)
    qualified_sheet = None
    participated_sheet = None
    for s in xls.sheet_names:
        s_up = s.strip().upper()
        if qualified_sheet is None and s_up.startswith("QUALIFIED"):
            qualified_sheet = s
        if participated_sheet is None and s_up.startswith("PARTICIPATED"):
            participated_sheet = s

    # template checks and read template bytes once
    qual_bytes = None
    part_bytes = None
    smart_bytes = None

    # use uploaded templates first, else try default root templates
    def fetch_template_bytes(uploaded_file_obj, default_path: Path):
        if uploaded_file_obj:
            return uploaded_file_obj.read()
        if default_path.exists():
            return default_path.read_bytes()
        return None

    # pick qualified template if requested
    if gen_qualified:
        qual_bytes = fetch_template_bytes(qualified_pdf_file, DEFAULT_ROOT_TEMPLATES["QUALIFIED"]) or fetch_template_bytes(None, DEFAULT_ROOT_TEMPLATES["FALLBACK_1"])
        if qual_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Qualified)")
            st.stop()

    if gen_participated:
        part_bytes = fetch_template_bytes(participated_pdf_file, DEFAULT_ROOT_TEMPLATES["PARTICIPATED"]) or fetch_template_bytes(None, DEFAULT_ROOT_TEMPLATES["FALLBACK_2"])
        if part_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Participated)")
            st.stop()

    if gen_smartedge:
        smart_bytes = fetch_template_bytes(smartedge_pdf_file, DEFAULT_ROOT_TEMPLATES["SMART_EDGE"])
        if smart_bytes is None:
            st.error(random.choice(MISSING_TEMPLATE_ERRORS) + " (Smart Edge)")
            st.stop()
        if smartedge_sheet is None:
            st.error(random.choice(MISSING_SHEET_ERRORS) + " (Smart Edge sheet must be one of: Names / Name / Smart Edge / Certificates)")
            st.stop()

    # read dataframes for actual sheet names (use the found sheet names)
    df_q = pd.DataFrame()
    df_p = pd.DataFrame()
    df_s = pd.DataFrame()
    try:
        if qualified_sheet and gen_qualified:
            df_q = pd.read_excel(excel_file, sheet_name=qualified_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read Qualified sheet '{qualified_sheet}': {e}")
        st.stop()
    try:
        if participated_sheet and gen_participated:
            df_p = pd.read_excel(excel_file, sheet_name=participated_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read Participated sheet '{participated_sheet}': {e}")
        st.stop()
    try:
        if smartedge_sheet and gen_smartedge:
            df_s = pd.read_excel(excel_file, sheet_name=smartedge_sheet, dtype=object)
    except Exception as e:
        st.error(f"Failed to read Smart Edge sheet '{smartedge_sheet}': {e}")
        st.stop()

    # build list of names (assume name column is first column)
    tasks = []
    group_counts = {"QUALIFIED": 0, "PARTICIPATED": 0, "SMART_EDGE_WORKSHOP": 0}

    if gen_qualified and not df_q.empty:
        q_names = df_q.iloc[:, 0].dropna().astype(str).tolist()
        group_counts["QUALIFIED"] = len(q_names)
        tasks += [("QUALIFIED", n.strip()) for n in q_names]

    if gen_participated and not df_p.empty:
        p_names = df_p.iloc[:, 0].dropna().astype(str).tolist()
        group_counts["PARTICIPATED"] = len(p_names)
        tasks += [("PARTICIPATED", n.strip()) for n in p_names]

    if gen_smartedge and not df_s.empty:
        s_names = df_s.iloc[:, 0].dropna().astype(str).tolist()
        group_counts["SMART_EDGE_WORKSHOP"] = len(s_names)
        tasks += [("SMART_EDGE_WORKSHOP", n.strip()) for n in s_names]

    if len(tasks) == 0:
        st.warning("No names found in the selected sheets. Nothing to generate.")
        st.stop()

    total = len(tasks)
    overall_progress = st.progress(0)
    overall_status = st.empty()

    # placeholders for per-group small progress UI
    group_done = {g: 0 for g in group_counts}
    group_status_placeholders = {}
    group_progress_bars = {}
    for g, cnt in group_counts.items():
        if cnt > 0:
            group_status_placeholders[g] = st.empty()
            group_progress_bars[g] = st.progress(0.0)

    # avoid duplicate filenames inside zip by tracking used paths
    used_zip_names = Counter()

    zip_buf = io.BytesIO()
    with ZipFile(zip_buf, "w") as zf:
        for idx, (group, name) in enumerate(tasks, start=1):
            group_done[group] += 1
            overall_status.info(f"Overall: {idx}/{total} — Generating {group} / {name}")

            # update per-group status and progress
            for g in group_status_placeholders:
                done = group_done.get(g, 0)
                total_g = group_counts.get(g, 0)
                group_status_placeholders[g].text(f"{g.replace('_',' ')}: {done}/{total_g} done")
                if total_g > 0:
                    group_progress_bars[g].progress(done / total_g)

            time.sleep(0.005)  # give UI a tiny moment to update

            # pick template bytes
            try:
                if group == "QUALIFIED":
                    tpl_bytes = qual_bytes
                elif group == "PARTICIPATED":
                    tpl_bytes = part_bytes
                else:
                    tpl_bytes = smart_bytes

                img = draw_name_on_template(tpl_bytes, name, X_CM, Y_CM, BASE_FONT_PT, MAX_WIDTH_CM)
                pdf_bytes = image_to_pdf_bytes(img)

                # safe filename and unique check
                safe_name = str(name).replace("/", "_").replace("\\", "_")
                folder = group
                base_filename = f"{safe_name}.pdf"
                zip_path = f"{folder}/{base_filename}"
                if used_zip_names[zip_path] > 0:
                    # append suffix
                    suffix = used_zip_names[zip_path] + 1
                    zip_path = f"{folder}/{safe_name}_{suffix}.pdf"
                used_zip_names[zip_path] += 1

                zf.writestr(zip_path, pdf_bytes)
            except Exception as e:
                # try to include error file inside the zip so user can inspect
                err_msg = f"Failed to generate for {name}: {e}"
                safe_name = str(name).replace("/", "_").replace("\\", "_")
                folder = group
                err_path = f"{folder}/{safe_name}_ERROR.txt"
                zf.writestr(err_path, err_msg.encode("utf-8"))

            overall_progress.progress(idx / total)

        overall_status.success("All items processed. Finalizing ZIP...")

    st.balloons()
    st.success(f"Done — {total} certificates generated (errors, if any, are in the ZIP).")
    zip_buf.seek(0)
    st.download_button("Download certificates ZIP", data=zip_buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
