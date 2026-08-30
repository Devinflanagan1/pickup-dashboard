import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Master Day-by-Day Revenue Management Dashboard",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. EXACT MASTER DAY-BY-DAY HEADERS & SEQUENCE
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
    # Inventory, Demand & Pickup Section
    "Blocked",
    "P/U",
    "Remaining Demand - Total Hotel",
    "Variance (TY - STLY)"
]


# ------------------------------------------------------------------------------
# 2. HELPER & CLEANING FUNCTIONS
# ------------------------------------------------------------------------------

def normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Strips timestamps (time component) to ensure exact date joins."""
    for col in df.columns:
        if "date" in str(col).lower():
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except Exception:
                pass
    return df


def load_file(file):
    if file is None:
        return None
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    return normalize_dates(df)


# ------------------------------------------------------------------------------
# 3. MASTER GRID BUILDER LOGIC
# ------------------------------------------------------------------------------

def build_master_day_by_day_grid(pcdc_df, extraction_df, mkt_df, lighthouse_df, synxis_df) -> pd.DataFrame:
    """
    Merges all 5 source files and strictly enforces exact DD column order and header naming.
    """
    # Start base grid from primary PCDC export
    master_df = pcdc_df.copy() if pcdc_df is not None else pd.DataFrame()

    # Merge Data Extraction (Sourcing "Remaining Demand - Total Hotel")
    if extraction_df is not None and "Date" in extraction_df.columns:
        if "System Total Demand - Total This Year" in extraction_df.columns:
            extraction_df = extraction_df.rename(
                columns={"System Total Demand - Total This Year": "Remaining Demand - Total Hotel"}
            )
        master_df = pd.merge(master_df, extraction_df, on="Date", how="left", suffixes=("", "_ext"))

    # Merge Market Segmentation
    if mkt_df is not None and "Date" in mkt_df.columns:
        master_df = pd.merge(master_df, mkt_df, on="Date", how="left", suffixes=("", "_mkt"))

    # Merge Lighthouse Rate Shop
    if lighthouse_df is not None and "Date" in lighthouse_df.columns:
        master_df = pd.merge(master_df, lighthouse_df, on="Date", how="left", suffixes=("", "_lh"))

    # Merge SynXis Rate Plan
    if synxis_df is not None and "Date" in synxis_df.columns:
        master_df = pd.merge(master_df, synxis_df, on="Date", how="left", suffixes=("", "_syn"))

    # Flatten MultiIndex columns if present
    if isinstance(master_df.columns, pd.MultiIndex):
        master_df.columns = [" ".join(str(c) for c in col if str(c) and not str(c).startswith("Unnamed")).strip() for col in master_df.columns.values]

    # Map alternative field names to exact target DD display headers
    field_mapping = {
        "PickUp": "P/U",
        "Blocked Rooms": "Blocked",
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

    # Force exact column ordering matching DESIRED_DD_HEADERS
    final_grid = master_df[DESIRED_DD_HEADERS]

    return final_grid


# ------------------------------------------------------------------------------
# 4. STREAMLIT APP UI & LAYOUT
# ------------------------------------------------------------------------------

st.title("Day-by-Day Revenue Management Dashboard")
st.write("Upload all required source files below to construct the Master Day-by-Day Report.")

# Section A: All 5 Required Source Uploaders
st.subheader("1. Source File Uploads")

col1, col2, col3 = st.columns(3)
with col1:
    pcdc_file = st.file_uploader("IDeaS PCDC Export", type=["csv", "xlsx"], key="pcdc")
    extraction_file = st.file_uploader("IDeaS Data Extraction", type=["csv", "xlsx"], key="ext")

with col2:
    mkt_file = st.file_uploader("IDeaS Market Segmentation", type=["csv", "xlsx"], key="mkt")
    lh_file = st.file_uploader("Lighthouse Rate Shop", type=["csv", "xlsx"], key="lh")

with col3:
    synxis_file = st.file_uploader("SynXis Rate Plan Export", type=["csv", "xlsx"], key="syn")

st.divider()

# Section B: Process & Tabbed Dashboard View
if pcdc_file is not None:
    try:
        # Load and normalize all 5 files
        pcdc_df = load_file(pcdc_file)
        extraction_df = load_file(extraction_file)
        mkt_df = load_file(mkt_file)
        lh_df = load_file(lh_file)
        synxis_df = load_file(synxis_file)

        # Build master grid with explicit headers and order
        master_grid = build_master_day_by_day_grid(
            pcdc_df, extraction_df, mkt_df, lh_df, synxis_df
        )

        # Tabbed Dashboard Layout
        tab_dd, tab_rate_plans, tab_export = st.tabs([
            "📊 Master Day-by-Day Grid", 
            "📈 Rate Plan Analysis", 
            "📥 Exports"
        ])

        with tab_dd:
            st.subheader("Day-by-Day Performance View")
            st.dataframe(master_grid, use_container_width=True)

        with tab_rate_plans:
            st.subheader("SynXis & Lighthouse Rate Plan Breakdown")
            if synxis_df is not None:
                st.write("### SynXis Rate Plans")
                st.dataframe(synxis_df, use_container_width=True)
            else:
                st.info("Upload SynXis Rate Plan export to view this tab.")

            if lh_df is not None:
                st.write("### Lighthouse Rate Shops")
                st.dataframe(lh_df, use_container_width=True)

        with tab_export:
            st.subheader("Download Reports")
            e1, e2 = st.columns(2)

            # CSV Export
            csv_data = master_grid.to_csv(index=False).encode("utf-8")
            e1.download_button(
                label="Download Flattened CSV Report",
                data=csv_data,
                file_name="master_day_by_day_report.csv",
                mime="text/csv",
                use_container_width=True
            )

            # Excel Export
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                master_grid.to_excel(writer, index=False, sheet_name="Day_by_Day")
                if synxis_df is not None:
                    synxis_df.to_excel(writer, index=False, sheet_name="SynXis_Rate_Plans")

            e2.download_button(
                label="Download Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="master_day_by_day_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as err:
        st.error(f"Error generating dashboard: {str(err)}")
else:
    st.info("Please upload at least your primary IDeaS PCDC file to load the report.")
