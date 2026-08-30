import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Master Day-by-Day Revenue Management Dashboard",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. EXACT MULTI-LEVEL HEADERS STRUCTURE
# ------------------------------------------------------------------------------
def build_multiindex_headers():
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


def create_master_dd_grid(pcdc_df, extraction_df, mkt_df, lh_df, synxis_df, show_comp_shops) -> pd.DataFrame:
    multi_index = build_multiindex_headers()
    
    # Establish base rows
    if pcdc_df is not None and "Date" in pcdc_df.columns:
        base_dates = pcdc_df["Date"].unique()
    else:
        base_dates = pd.date_range(start=pd.Timestamp.now(), periods=30).date

    final_df = pd.DataFrame(index=range(len(base_dates)), columns=multi_index)
    final_df[("Date", "", "")] = base_dates
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

    # Toggle control for Competitor Shops display
    if not show_comp_shops:
        comp_cols = [c for c in final_df.columns if c[0] == "Competitor Shops" or c[0] == "Comp Set Avg"]
        final_df = final_df.drop(columns=comp_cols)

    return final_df


# ------------------------------------------------------------------------------
# 3. SIDEBAR LAYOUT: FILE UPLOADERS & TOGGLES
# ------------------------------------------------------------------------------
with st.sidebar:
    st.header("1. Upload Source Files")
    
    pcdc_file = st.file_uploader("IDeaS PCDC Export", type=["csv", "xlsx"], key="sb_pcdc")
    extraction_file = st.file_uploader("IDeaS Data Extraction", type=["csv", "xlsx"], key="sb_ext")
    mkt_file = st.file_uploader("IDeaS Market Segmentation", type=["csv", "xlsx"], key="sb_mkt")
    lh_file = st.file_uploader("Lighthouse Rate Shop", type=["csv", "xlsx"], key="sb_lh")
    synxis_file = st.file_uploader("SynXis Rate Plan Export", type=["csv", "xlsx"], key="sb_syn")

    st.divider()
    st.header("2. Display Options")
    
    # Interactive View Toggles
    show_comp_shops = st.toggle("Show Competitor Shops", value=True)
    rate_shop_view = st.radio("Rate Shop Display Mode", ["Standard", "Comp Set Average", "Variance Only"])


# ------------------------------------------------------------------------------
# 4. MAIN DASHBOARD CONTENT
# ------------------------------------------------------------------------------
st.title("Master Day-by-Day Revenue Management Dashboard")

if pcdc_file is not None or extraction_file is not None:
    try:
        pcdc_df = load_file(pcdc_file)
        extraction_df = load_file(extraction_file)
        mkt_df = load_file(mkt_file)
        lh_df = load_file(lh_file)
        synxis_df = load_file(synxis_file)

        master_grid = create_master_dd_grid(
            pcdc_df, extraction_df, mkt_df, lh_df, synxis_df, show_comp_shops
        )

        tab_dd, tab_rate_plans, tab_export = st.tabs([
            "📊 Master Day-by-Day Grid", 
            "📈 Rate Plan Analysis", 
            "📥 Exports"
        ])

        with tab_dd:
            st.subheader("Master Day-by-Day Report")
            st.dataframe(master_grid, use_container_width=True)

        with tab_rate_plans:
            st.subheader("SynXis & Lighthouse Rate Plan Data")
            if synxis_df is not None:
                st.write("### SynXis Rate Plans")
                st.dataframe(synxis_df, use_container_width=True)
            if lh_df is not None:
                st.write(f"### Lighthouse Rate Shops ({rate_shop_view} View)")
                st.dataframe(lh_df, use_container_width=True)

        with tab_export:
            st.subheader("Download Formatted Reports")
            e1, e2 = st.columns(2)

            # Flattened CSV Export
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

            # Excel Export
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
    st.info("Use the sidebar on the left to upload your source files and configure display toggles.")
