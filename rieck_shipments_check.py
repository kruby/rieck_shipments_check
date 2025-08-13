import os
import re
import zipfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


# --------------------------- Config ---------------------------
st.set_page_config(page_title="Shipment Presence Checker", page_icon="📦", layout="wide")

TITLE = "📦 Shipment Presence Checker"
INTRO = (
    "Compare Excel **Shipment No.** values against files from your processed folder. "
    "We assume each processed filename begins with an **8-digit shipment number** "
    "(e.g., `12345678_invoice.pdf`)."
)

# --------------------------- Helpers ---------------------------
def extract_shipment_from_filename(filename: str) -> str | None:
    """
    Return the first 8 digits at the very start of a filename.
    Example: '12345678_invoice.pdf' -> '12345678'
             'AB12345678.pdf'      -> None (not at start)
    """
    base = os.path.basename(filename)
    m = re.match(r"^(\d{8})", base)
    return m.group(1) if m else None


def clean_excel_shipment(val: object) -> str | None:
    """
    Keep digits only from Excel 'Shipment No.' values and accept exactly 8 digits.
    """
    if pd.isna(val):
        return None
    digits = re.sub(r"\D", "", str(val))
    return digits if len(digits) == 8 else None


def read_excel_shipment_column(uploaded_file) -> pd.Series:
    """
    Read the uploaded Excel and return a Series of cleaned 8-digit shipment numbers.
    Expects a column named exactly 'Shipment No.'.
    Supports .xlsx/.xlsm (openpyxl) and .xls (xlrd).
    """
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, dtype=str, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    if "Shipment No." not in df.columns:
        st.error("Excel must contain a column named **'Shipment No.'** (exact match).")
        st.stop()

    series = df["Shipment No."].map(clean_excel_shipment)
    valid = series.dropna().drop_duplicates().astype(str)
    invalid = df[series.isna()].copy()
    invalid = invalid[["Shipment No."]] if "Shipment No." in invalid.columns else invalid

    return valid, invalid


def extract_from_uploaded_files(files) -> set[str]:
    """
    From a list of UploadedFile objects, derive a set of 8-digit shipment numbers
    using the filename rule (first 8 digits at the start).
    """
    shipments = set()
    for f in files:
        ship = extract_shipment_from_filename(f.name)
        if ship:
            shipments.add(ship)
    return shipments


def extract_from_zip(uploaded_zip) -> set[str]:
    """
    From an uploaded ZIP, derive a set of 8-digit shipment numbers by scanning entry names.
    """
    shipments = set()
    try:
        with zipfile.ZipFile(uploaded_zip) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                base = os.path.basename(info.filename)
                ship = extract_shipment_from_filename(base)
                if ship:
                    shipments.add(ship)
    except Exception as e:
        st.error(f"Failed to read ZIP: {e}")
        st.stop()
    return shipments


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name="Missing", file_name="missing_shipments.xlsx") -> tuple[bytes, str]:
    """
    Convert a DataFrame to a downloadable XLSX (bytes, filename).
    """
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf.read(), file_name


# --------------------------- UI ---------------------------
st.title(TITLE)
st.caption(INTRO)

with st.expander("How it works", expanded=False):
    st.markdown(
        """
        **Steps**
        1. Upload the Excel that contains a column named **`Shipment No.`** (your column B). Create this file in BC from the list Posted Warehouse Shipments. (Filter for Rieck, and customer NOT Nordisk)
        2. Provide the processed files either by **uploading multiple files** or **uploading a ZIP of the processed folder**. (Download all the Posted Sales Shipment files from Rieck > Processed folder on SharePoint)
        3. Click **Run check** to compare.
        4. Download the **Missing Shipments** as an Excel file.

        **Rules & Notes**
        - A processed file counts for a shipment if its **filename starts with 8 digits** (e.g., `12345678_document.pdf`).
        - In Excel, values in **`Shipment No.`** are cleaned to digits only and must be exactly **8 digits**.
        - Works the same on Windows, macOS, Linux, and Streamlit Cloud.
        """
    )

st.subheader("1) Upload Excel (must contain 'Shipment No.')")
excel_file = st.file_uploader("Excel (.xlsx, .xlsm, .xls)", type=["xlsx", "xlsm", "xls"])

st.subheader("2) Provide processed files")
mode = st.radio(
    "How will you provide the processed files?",
    options=["Upload multiple files", "Upload a ZIP of the processed folder"],
    horizontal=True,
)

uploaded_files = None
uploaded_zip = None
if mode == "Upload multiple files":
    uploaded_files = st.file_uploader(
        "Drop or browse processed files (you can select many)",
        type=None,  # allow any extension
        accept_multiple_files=True,
        help="We only read the filenames; contents are ignored."
    )
else:
    uploaded_zip = st.file_uploader(
        "Upload a ZIP of the processed folder",
        type=["zip"],
        help="On Windows: right-click the folder → Send to → Compressed (zipped) folder."
    )

with st.sidebar:
    st.header("Options")
    show_preview = st.checkbox("Show preview tables", value=True)

run = st.button(
    "▶️ Run check",
    type="primary",
    disabled=excel_file is None or ((uploaded_files is None or len(uploaded_files) == 0) and uploaded_zip is None),
)

# --------------------------- Logic ---------------------------
if run:
    # Read Excel shipments
    valid_shipments, invalid_rows = read_excel_shipment_column(excel_file)
    if valid_shipments.empty:
        st.warning("No valid 8-digit shipment numbers found in Excel after cleaning.")
        st.stop()

    # Read processed files (filenames only)
    if uploaded_zip is not None:
        folder_shipments = extract_from_zip(uploaded_zip)
    else:
        folder_shipments = extract_from_uploaded_files(uploaded_files or [])

    # Compare sets
    folder_set = set(folder_shipments)
    present_mask = valid_shipments.isin(folder_set)
    result_df = pd.DataFrame({"Shipment No.": valid_shipments.values, "Present": present_mask.values})
    missing_df = result_df[~result_df["Present"]].drop(columns=["Present"]).reset_index(drop=True)

    # Metrics
    st.success("Done!")
    c1, c2, c3 = st.columns(3)
    c1.metric("Unique shipments in folder", f"{len(folder_set):,}")
    c2.metric("Unique shipments in Excel", f"{len(valid_shipments):,}")
    c3.metric("Missing shipments", f"{len(missing_df):,}")

    # Previews
    if show_preview:
        st.subheader("Missing shipments (preview)")
        st.dataframe(missing_df, use_container_width=True, height=320)

        if not invalid_rows.empty:
            st.subheader("Rows with invalid 'Shipment No.' (not exactly 8 digits after cleaning)")
            st.dataframe(invalid_rows, use_container_width=True, height=240)

    # Downloads
    xlsx_bytes, fname = df_to_xlsx_bytes(missing_df, sheet_name="Missing", file_name="missing_shipments.xlsx")
    st.download_button(
        "⬇️ Download Missing Shipments (Excel)",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Advanced: Download full presence summary"):
        full_buf = BytesIO()
        with pd.ExcelWriter(full_buf, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Summary")
            pd.DataFrame({"FolderShipments": sorted(folder_set)}).to_excel(writer, index=False, sheet_name="FolderSet")
        full_buf.seek(0)
        st.download_button(
            "⬇️ Download Full Summary (Excel)",
            data=full_buf.read(),
            file_name="shipment_presence_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
