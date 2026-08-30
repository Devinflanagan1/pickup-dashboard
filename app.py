import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Master Day-by-Day Revenue Dashboard",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. DEFINE EXACT MULTI-LEVEL HEADERS STRUCTURE
# ------------------------------------------------------------------------------
def build_multiindex_headers():
    """
    Constructs the exact 3-level header tuple hierarchy matching your required layout.
    """
    header_tuples = [
        ("DOW", "", ""),
        ("Date", "", ""),
        ("Days Left", "", ""),
        ("Events", "", ""),
        ("Rooms Left to Sell", "", ""),
        ("BAR", "", ""),
        ("Comp Set Avg", "", ""),
        ("Last Room Value", "", ""),
        # Competitor Shops
        ("Competitor Shops", "21c Museum Hotel Bentonville - MGallery", ""),
        ("Competitor Shops", "Motto By Hilton Bentonville Downtown", ""),
        ("Competitor Shops", "AC Hotel by Marriott Bentonville", ""),
        ("Competitor Shops", "DoubleTree Suites by Hilton Bentonville", ""),
        ("Competitor Shops", "Current", ""),
        # Inventory & Sales
        ("OOO", "", ""),
        ("Ovrbk", "", ""),
        ("Rooms Sold", "Total Hotel", "Current"),
        ("Rooms Sold", "Total Hotel", "Change"),
        ("Rooms Sold", "Total Transient", "Current"),
        ("Rooms Sold", "Total Transient", "Change"),
        ("Rooms Sold", "Total Group", "Current"),
        ("Rooms Sold", "Total Group", "Change"),
        ("Rooms Sold", "Blocked", "Blocked"),
        ("Rooms Sold", "P/U", "P/U"),
        # Rooms OTB STLY
        ("Rooms OTB STLY", "Total OTB STLY", ""),
        ("Rooms OTB STLY", "Variance (TY - STLY)", ""),
        ("Rooms OTB STLY", "Transient OTB STLY", ""),
        ("Rooms OTB STLY", "Variance (TY - STLY)", ""),
        ("Rooms OTB STLY", "Group OTB STLY", ""),
        ("Rooms OTB STLY", "Variance (TY - STLY)", ""),
        # Demand & Occupancy Forecasts
        ("Remaining Demand", "Remaining", ""),
        ("Occupancy Forecast", "Total Hotel", ""),
        ("Occupancy Forecast", "Total Transient", ""),
        ("Occupancy Forecast", "Total Group", ""),
        ("Occupancy Forecast %", "Total Hotel", ""),
        ("Occupancy Forecast %", "Total Transient", ""),
        ("Occupancy Forecast %", "Total Group", ""),
        # ADR
        ("Booked ADR(USD)", "Total Hotel", ""),
        ("Booked ADR(USD)", "Total Transient", ""),
        ("Booked ADR(USD)", "Total Group", ""),
        ("Estimated ADR", "Total Hotel", ""),
        ("Estimated ADR", "Total Transient", ""),
        ("Estimated ADR", "Total Group", "")
    ]
    
    return pd.MultiIndex.from_tuples(header_tuples, names=["Category", "SubCategory", "Metric"])


# ------------------------------------------------------------------------------
# 2. DATA PROCESSING & CLEANING HELPERS
# ------------------------------------------------------------------------------

def normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Strips timestamps from date fields to fix join issues."""
    for col in df.columns:
        if "date" in str(col).lower():
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except Exception:
                pass
    return df


def load_file(uploaded_file):
    if uploaded_file is None:
        return None
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
    return normalize_dates(df)


def create_master_dd_grid(pcdc_df, extraction_df, mkt_df, lh_df, synxis_df) -> pd.DataFrame:
    """
    Combines input sources and shapes data into the exact multi-level header structure.
    """
    multi_index = build_multiindex_headers()
    
    # Establish base rows (preferring PCDC date range)
    if pcdc_df is not None and "Date" in pcdc_df.columns:
        base_dates = pcdc_df["Date"].unique()
    else:
        base_dates = pd.date_range(start=pd.Timestamp.now(), periods=30).date

    # Create empty dataframe initialized with the exact MultiIndex columns
    final_df = pd.DataFrame(index=range(len(base_dates)), columns=multi_index)
    final_df[("Date", "", "")] = base_dates

    # Calculate/populate DOW
    final_df[("DOW", "", "")] = pd.to_datetime(final_df[("Date", "", "")]).dt.strftime("%a")

    # Map IDeaS Data Extraction field for Remaining Demand
    if extraction_df is not None and "Date" in extraction_df.columns:
        demand_col = "System Total Demand - Total This Year"
        if demand_col in extraction_df.columns:
            merged = pd.merge(
                final_df[[("Date", "", "")]], 
                extraction_df[["Date", demand_col]], 
                left_on=("Date", "", ""), 
                right_on="Date", 
                how="left"
            )
            final_df[("Remaining Demand", "Remaining", "")] = merged[demand_col].values

    return final_df


# ------------------------------------------------------------------------------
# 3. STREAMLIT APP LAYOUT & WIDGETS
# ------------------------------------------------------------------------------

st.title("Master Day-by-Day Revenue Management Dashboard")
st.write("Upload all 5 required source exports to construct your complete Day-by-Day report.")

# Section 1: File Uploaders
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

# Section 2: Report Processing & Displays
if pcdc_file is not None or extraction_file is not None:
    try:
        # Read & normalize input files
        pcdc_df = load_file(pcdc_file)
        extraction_df = load_file(extraction_file)
        mkt_df = load_file(mkt_file)
        lh_df = load_file(lh_file)
        synxis_df = load_file(synxis_file)

        # Build Master Grid
        master_grid = create_master_dd_grid(pcdc_df, extraction_df, mkt_df, lh_df, synxis_df)

        # Tabbed Views
        tab_dd, tab_rate_plans, tab_export = st.tabs([
            "📊 Master Day-by-Day Grid", 
            "📈 Rate Plan Breakdown", 
            "📥 Exports"
        ])

        with tab_dd:
            st.subheader("Master Day-by-Day Report")
            st.dataframe(master_grid, use_container_width=True)

        with tab_rate_plans:
            st.subheader("SynXis & Lighthouse Data")
            if synxis_df is not None:
                st.write("### SynXis Rate Plans")
                st.dataframe(synxis_df, use_container_width=True)
            if lh_df is not None:
                st.write("### Lighthouse Rate Shops")
                st.dataframe(lh_df, use_container_width=True)

        with tab_export:
            st.subheader("Download Formatted Reports")
            e1, e2 = st.columns(2)

            # Export Option 1: Flattened CSV (single-row column headers)
            csv_df = master_grid.copy()
            csv_df.columns = [" ".join(str(c) for c in col if str(c) and not str(c).startswith("Unnamed")).strip() for col in csv_df.columns.values]
            csv_bytes = csv_df.to_csv(index=False).encode("utf-8")
            
            e1.download_button(
                label="Download Flattened CSV",
                data=csv_bytes,
                file_name="master_day_by_day_report.csv",
                mime="text/csv",
                use_container_width=True
            )

            # Export Option 2: Excel File preserving exact multi-row headers
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                master_grid.to_excel(writer, sheet_name="Day_by_Day")
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
        st.error(f"Error building Master Day-by-Day Report: {str(err)}")
else:
    st.info("Please upload your files above to generate the report.")
