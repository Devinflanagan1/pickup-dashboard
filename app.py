import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Master Day-by-Day Revenue Management Dashboard",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. EXACT TARGET HEADERS
# ------------------------------------------------------------------------------
DESIRED_DD_HEADERS = [
    "DOW",
    "Date",
    "Days Left",
    "Events",
    "Rooms Left to Sell",
    "BAR",
    "Comp Set Avg",
    "Last Room Value",
    "21c Museum Hotel Bentonville - MGallery",
    "Motto By Hilton Bentonville Downtown",
    "AC Hotel by Marriott Bentonville",
    "DoubleTree Suites by Hilton Bentonville",
    "Current",
    "OOO",
    "Ovrbk",
    "Rooms Sold Total Hotel Current",
    "Rooms Sold Total Hotel Change",
    "Rooms Sold Total Transient Current",
    "Rooms Sold Total Transient Change",
    "Rooms Sold Total Group Current",
    "Rooms Sold Total Group Change",
    "Blocked",
    "P/U",
    "Rooms OTB STLY Total OTB STLY",
    "Rooms OTB STLY Variance (TY - STLY)",
    "Rooms OTB STLY Transient OTB STLY",
    "Rooms OTB STLY Variance (TY - STLY)",
    "Rooms OTB STLY Group OTB STLY",
    "Rooms OTB STLY Variance (TY - STLY)",
    "Remaining Demand - Total Hotel",
    "Occupancy Forecast Total Hotel",
    "Occupancy Forecast Total Transient",
    "Occupancy Forecast Total Group",
    "Occupancy Forecast % Total Hotel",
    "Occupancy Forecast % Total Transient",
    "Occupancy Forecast % Total Group",
    "Booked ADR(USD) Total Hotel",
    "Booked ADR(USD) Total Transient",
    "Booked ADR(USD) Total Group",
    "Estimated ADR Total Hotel",
    "Estimated ADR Total Transient",
    "Estimated ADR Total Group"
]

def normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if "date" in str(col).lower():
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except Exception:
                pass
    return df

def process_day_by_day_grid(raw_df: pd.DataFrame, synxis_df: pd.DataFrame = None, reservation_statuses = None, price_change_mode = "Standard") -> pd.DataFrame:
    df = raw_df.copy()

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"]).dt.date

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(c) for c in col).strip() for col in df.columns.values]

    # Filter SynXis reservation status if applied
    if synxis_df is not None and reservation_statuses:
        status_col = next((c for c in synxis_df.columns if "status" in c.lower()), None)
        if status_col:
            synxis_df = synxis_df[synxis_df[status_col].isin(reservation_statuses)]

    # Handle competitor price change adjustments if requested
    if price_change_mode == "Variance Only":
        for comp_col in ["21c Museum Hotel Bentonville - MGallery", "Motto By Hilton Bentonville Downtown", "AC Hotel by Marriott Bentonville", "DoubleTree Suites by Hilton Bentonville"]:
            if comp_col in df.columns:
                df[comp_col] = df[comp_col] - df.get("BAR", 0)

    field_mappings = {
        "System Total Demand - Total This Year": "Remaining Demand - Total Hotel",
        "PickUp": "P/U",
        "Blocked Rooms": "Blocked",
        "Rooms_Current": "Rooms Sold Total Hotel Current",
        "Rooms_Change": "Rooms Sold Total Hotel Change",
    }
    df = df.rename(columns=field_mappings)

    for col in DESIRED_DD_HEADERS:
        if col not in df.columns:
            df[col] = None

    df_final = df[DESIRED_DD_HEADERS]
    return df_final


# ------------------------------------------------------------------------------
# 2. SIDEBAR LAYOUT (FILE UPLOADERS & FILTERS)
# ------------------------------------------------------------------------------
with st.sidebar:
    st.header("Source File Uploads")
    uploaded_file = st.file_uploader("Upload Master Data File", type=["csv", "xlsx"], key="master_upload")
    synxis_file = st.file_uploader("SynXis Rate Plan Export (Optional)", type=["csv", "xlsx"], key="synxis_upload")
    lighthouse_file = st.file_uploader("Lighthouse Rate Shop Export (Optional)", type=["csv", "xlsx"], key="lh_upload")

    st.divider()
    st.header("Report Filters & Toggles")
    
    # SynXis Reservation Status Filter
    reservation_status_filter = st.multiselect(
        "SynXis Reservation Status Filter",
        options=["Confirmed", "Cancelled", "Guaranteed", "Tentative"],
        default=["Confirmed", "Guaranteed"]
    )

    # Competitor Price Change Toggle/Mode
    price_change_mode = st.selectbox(
        "Competitor Price View",
        options=["Standard", "Variance Only", "Comp Set Average Focus"]
    )


# ------------------------------------------------------------------------------
# 3. MAIN APP INTERFACE & EXPORTS
# ------------------------------------------------------------------------------
st.title("Day-by-Day Revenue Report Generator")

if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        raw_data = pd.read_csv(uploaded_file)
    else:
        raw_data = pd.read_excel(uploaded_file)

    synxis_data = None
    if synxis_file is not None:
        synxis_data = pd.read_csv(synxis_file) if synxis_file.name.endswith(".csv") else pd.read_excel(synxis_file)

    processed_df = process_day_by_day_grid(raw_data, synxis_data, reservation_status_filter, price_change_mode)

    st.subheader("Day-by-Day Report Preview")
    st.dataframe(processed_df, use_container_width=True)

    col1, col2 = st.columns(2)

    csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
    col1.download_button(
        label="Download Flattened CSV",
        data=csv_bytes,
        file_name="day_by_day_report.csv",
        mime="text/csv",
        use_container_width=True
    )

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        processed_df.to_excel(writer, index=False, sheet_name="Day_by_Day")
    
    col2.download_button(
        label="Download Excel (.xlsx)",
        data=excel_buffer.getvalue(),
        file_name="day_by_day_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("Please upload your files via the sidebar to get started.")
