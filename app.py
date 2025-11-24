# Updated app.py
# - Safe, case-insensitive sheet detection when reading Excel
# - Graceful fallbacks to empty DataFrame when sheet not present
# - Centered, resizable logo using Streamlit columns (no raw HTML)
# - Avoids deprecated "use_column_width" API by using width or explicit sizing

import io
import os
import streamlit as st
import pandas as pd
from PIL import Image

st.set_page_config(page_title="Certificate Generator (Updated)", layout="wide")

# --------------------------
# HELPERS
# --------------------------

def find_sheet_exact(xls: pd.ExcelFile, target_name: str):
    """Find a sheet by name case-insensitively and ignoring surrounding whitespace.
    Returns the actual sheet name if found, otherwise None.
    """
    target_up = target_name.strip().upper()
    for s in xls.sheet_names:
        if s.strip().upper() == target_up:
            return s
    return None


def read_sheet_safely(excel_path_or_buffer, sheet_target: str, dtype=None, **kwargs):
    """Return (df, sheet_name_used).
    If sheet not found returns (pd.DataFrame(), None).
    """
    try:
        xls = pd.ExcelFile(excel_path_or_buffer)
    except Exception as e:
        st.error(f"Failed to open Excel file: {e}")
        return pd.DataFrame(), None

    sheet_name = find_sheet_exact(xls, sheet_target)
    if sheet_name is None:
        return pd.DataFrame(), None

    try:
        df = pd.read_excel(excel_path_or_buffer, sheet_name=sheet_name, dtype=dtype, **kwargs)
        return df, sheet_name
    except Exception as e:
        st.error(f"Error reading sheet '{sheet_name}': {e}")
        return pd.DataFrame(), sheet_name


# --------------------------
# LOGO (centered + resizable)
# --------------------------

with st.sidebar:
    st.header("UI Settings")
    logo_width = st.slider("Logo width (px)", min_value=50, max_value=500, value=150)
    show_logo = st.checkbox("Show logo", value=True)

if show_logo:
    # Attempt to locate logo in repo (relative path). Adjust path as needed.
    # Common places: ./logo.png, ./assets/logo.png, ./static/logo.png
    possible_paths = [
        "logo.png",
        "assets/logo.png",
        "static/logo.png",
        "images/logo.png",
        "./logo.png",
    ]
    logo_path = None
    for p in possible_paths:
        if os.path.exists(p):
            logo_path = p
            break

    if logo_path:
        try:
            img = Image.open(logo_path)
            # center using columns
            col_left, col_middle, col_right = st.columns([1, 2, 1])
            with col_middle:
                st.image(img, width=logo_width)
        except Exception as e:
            st.warning(f"Could not load logo image: {e}")
    else:
        st.info("No logo file found in repo. Put logo.png at project root or update the path.")

# --------------------------
# UPLOAD / EXCEL handling
# --------------------------

st.title("Certificate Generator (Safe Excel read)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"], accept_multiple_files=False)

# show useful debug/info in an expander
with st.expander("Excel file diagnostics"):
    if uploaded_file is None:
        st.write("No file uploaded yet.")
    else:
        try:
            xls = pd.ExcelFile(uploaded_file)
            st.write("Detected sheets:", xls.sheet_names)
        except Exception as e:
            st.write("Failed to read excel file:", e)


if uploaded_file is not None:
    # Read sheets safely using the helper
    df_qualified, qualified_sheet = read_sheet_safely(uploaded_file, "QUALIFIED", dtype=object)
    # NOTE: pd.ExcelFile consumes the buffer's pointer; to reuse uploaded_file we need to reset buffer
    # UploadedFile is a BytesIO-like object; rewind it between reads
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    df_participated, participated_sheet = read_sheet_safely(uploaded_file, "PARTICIPATED", dtype=object)
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    # Example: read 'SMARTEDGE' sheet where name might be 'SmartEdge' or user defined
    df_smartedge, smartedge_sheet = read_sheet_safely(uploaded_file, "SMARTEDGE", dtype=object)

    st.subheader("Read results")
    st.write(f"QUALIFIED sheet used: {qualified_sheet}")
    st.write(f"PARTICIPATED sheet used: {participated_sheet}")
    st.write(f"SMARTEDGE sheet used: {smartedge_sheet}")

    if not df_qualified.empty:
        st.markdown("**QUALIFIED preview**")
        st.dataframe(df_qualified.head())
    else:
        st.info("QUALIFIED sheet not found or empty.")

    if not df_participated.empty:
        st.markdown("**PARTICIPATED preview**")
        st.dataframe(df_participated.head())
    else:
        st.info("PARTICIPATED sheet not found or empty.")

    # Continue with certificate generation logic below — placeholder
    st.success("Sheets loaded (if present). You can now continue to certificate generation steps.")

# --------------------------
# NOTES for maintainers (printed to logs)
# --------------------------
st.write("\n---\nMaintainer notes:\n- This app now looks for sheet names case-insensitively.\n- If you still see Worksheet not found errors, ensure the uploaded Excel actually contains the exact sheet name (QUALIFIED / PARTICIPATED) ignoring case and whitespace.")

# End of file
