import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Master Day-by-Day Revenue Dashboard",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. EXACT TARGET DAY-BY-DAY HEADERS & ORDER
# ------------------------------------------------------------------------------
DESIRED_DD_HEADERS = [
    "DOW",
    "Date",
    "Days Left",
    "Events",
    # Rooms Section
    "Rooms Current",
    "Rooms Change",
    "Rooms OTB STLY",
    "Rooms Variance",
    # ADR Section
    "Estimated ADR",
    "ADR Current",
    "ADR Change",
    "ADR OTB STLY",
    "ADR Variance",
    # Revenue Section
    "Rev Current",
    "Rev Change",
    "Rev OTB STLY",
    "Rev Variance (TY - STLY)",
    # Demand, Pickup & Inventory
    "Blocked",
    "P/U",
    "Remaining Demand - Total Hotel",
    "Variance (TY - STLY)"
]


# ------------------------------------------------------------------------------
# 2. DATA CLEANING & MERGING LOGIC
# ------------------------------------------------------------------------------

def normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Strips timestamps from date columns to prevent blank joins/reports."""
    for col in df.columns:
        if "date" in str(col).lower():
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except Exception:
                pass
    return df


def parse_ideas_data(file) -> pd.DataFrame:
    """Parses IDeaS Data Extraction export."""
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    df = normalize_dates(df)
    
    # Map IDeaS specific fields
    ideas_mapping = {
        "System Total Demand - Total This Year": "Remaining Demand - Total Hotel",
        "PickUp": "P/U",
        "Blocked Rooms": "Blocked"
    }
    return df.rename(columns=ideas_mapping)


def parse_synxis_data(file) -> pd.DataFrame:
    """Parses SynXis reservation report."""
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    return normalize_dates(df)


def parse_lighthouse_data(file) -> pd.DataFrame:
    """Parses Lighthouse rate shop data."""
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    return normalize_dates(df)


def build_master_dd_grid(ideas_df, synxis_df=None, lighthouse_df=None) -> pd.DataFrame:
    """
    Merges multi-source inputs and strictly enforces exact DD column structure and order.
    """
    # Start with base dataset
    master_df = ideas_df.copy()

    # Merge SynXis if provided
    if synxis_df is not None and "Date" in synxis_df.columns and "Date" in master_df.columns:
        master_df = pd.merge(master_df, synxis_df, on="Date", how="left", suffixes=("", "_synxis"))

    # Merge Lighthouse if provided
    if lighthouse_df is not None and "Date" in lighthouse_df.columns and "Date" in master_df.columns:
        master_df = pd.merge(master_df, lighthouse_df, on="Date", how="left", suffixes=("", "_lh"))

    # Handle MultiIndex columns if present
    if isinstance(master_df.columns, pd.MultiIndex):
        master_df.columns = [" ".join(str(c) for c in col if str(c) and not str(c).startswith("Unnamed")).strip() for col in master_df.columns.values]

    # Map generic source headers to exact target DD display headers
    field_mapping = {
        "Rooms_Current": "Rooms Current",
        "Rooms_Change": "Rooms Change",
        "Rooms_STLY": "Rooms OTB STLY",
        "Rooms_Var": "Rooms Variance",
        "ADR_Current": "ADR Current",
        "ADR_Change": "ADR Change",
        "ADR_STLY": "ADR OTB STLY",
        "ADR_Var": "ADR Variance",
        "Rev_Current": "Rev Current",
        "Rev_Change": "Rev Change",
        "Rev_STLY": "Rev OTB STLY",
        "Rev_Var": "Rev Variance (TY - STLY)",
    }
    master_df = master_df.rename(columns=field_mapping)

    # Force missing columns to exist with null values
    for header in DESIRED_DD_HEADERS:
        if header not in master_df.columns:
            master_df[header] = None

    # Reindex columns to strictly match DESIRED_DD_HEADERS layout
    final_grid = master_df[DESIRED_DD_HEADERS]

    return final_grid


# ------------------------------------------------------------------------------
# 3. STREAMLIT APP INTERFACE
# ------------------------------------------------------------------------------

st.title("Day-by-Day Revenue Management Dashboard")
st.write("Upload your data sources to generate the Master Day-by-Day layout.")

col1, col2, col3 = st.columns(3)

with col1:
    ideas_file = st.file_uploader("1. Upload IDeaS Export", type=["csv", "xlsx"])

with col2:
    synxis_file = st.file_uploader("2. Upload SynXis Export (Optional)", type=["csv", "xlsx"])

with col3:
    lighthouse_file = st.file_uploader("3. Upload Lighthouse Export (Optional)", type=["csv", "xlsx"])

if ideas_file is not None:
    try:
        # Load inputs
        ideas_df = parse_ideas_data(ideas_file)
        synxis_df = parse_synxis_data(synxis_file) if synxis_file else None
        lighthouse_df = parse_lighthouse_data(lighthouse_file) if lighthouse_file else None

        # Build master grid
        master_dd_grid = build_master_dd_grid(ideas_df, synxis_df, lighthouse_df)

        st.subheader("Master Day-by-Day Report")
        st.dataframe(master_dd_grid, use_container_width=True)

        # Export options
        st.subheader("Exports")
        e_col1, e_col2 = st.columns(2)

        # CSV Download (Flattened single-row header format)
        csv_bytes = master_dd_grid.to_csv(index=False).encode("utf-8")
        e_col1.download_button(
            label="Download CSV Report",
            data=csv_bytes,
            file_name="day_by_day_report.csv",
            mime="text/csv",
            use_container_width=True
        )

        # Excel Download
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            master_dd_grid.to_excel(writer, index=False, sheet_name="Day_by_Day")

        e_col2.download_button(
            label="Download Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="day_by_day_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Error building Master Day-by-Day Report: {str(e)}")
else:
    st.info("Please upload at least your primary IDeaS file to populate the report.")
