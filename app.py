import io
import pandas as pd
import streamlit as st

# Set page configuration
st.set_page_config(
    page_title="Day-by-Day Revenue Report Generator",
    layout="wide"
)

# ------------------------------------------------------------------------------
# 1. EXACT MASTER DAY-BY-DAY DISPLAY HEADERS
# ------------------------------------------------------------------------------
DESIRED_DD_HEADERS = [
    "DOW",
    "Date",
    "Days Left",
    "Events",
    # Rooms Metrics
    "Rooms Current",
    "Rooms Change",
    "Rooms OTB STLY",
    "Rooms Variance",
    # ADR Metrics
    "Estimated ADR",
    "ADR Current",
    "ADR Change",
    "ADR OTB STLY",
    "ADR Variance",
    # Revenue Metrics
    "Rev Current",
    "Rev Change",
    "Rev OTB STLY",
    "Rev Variance (TY - STLY)",
    # Demand & Pickup
    "Blocked",
    "P/U",
    "Remaining Demand - Total Hotel",
    "Variance (TY - STLY)"
]


# ------------------------------------------------------------------------------
# 2. HELPER & TRANSFORMATION FUNCTIONS
# ------------------------------------------------------------------------------

def clean_dates_and_timestamps(df: pd.DataFrame) -> pd.DataFrame:
    """Strip time components from date columns to prevent join/display failures."""
    for col in df.columns:
        if "date" in str(col).lower():
            try:
                df[col] = pd.to_datetime(df[col]).dt.date
            except Exception:
                pass
    return df


def map_and_order_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Maps raw output/export fields to target header names and strictly orders
    columns according to DESIRED_DD_HEADERS.
    """
    df = df.copy()

    # Flatten MultiIndex columns if present
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(c) for c in col if str(c) and not str(c).startswith("Unnamed")).strip() for col in df.columns.values]

    # Map IDeaS / PMS / Export source field names to exact target headers
    field_mapping = {
        "System Total Demand - Total This Year": "Remaining Demand - Total Hotel",
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
    
    df = df.rename(columns=field_mapping)

    # Ensure all required target headers exist in the dataframe
    for header in DESIRED_DD_HEADERS:
        if header not in df.columns:
            df[header] = None

    # Force exact column ordering according to specified DD headers
    df_ordered = df[DESIRED_DD_HEADERS]
    
    return df_ordered


def process_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Reads CSV or Excel input and applies cleaning and header normalization."""
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Step A: Clean date/timestamp issues
    df = clean_dates_and_timestamps(df)

    # Step B: Apply custom column structure & mapping
    df_final = map_and_order_headers(df)

    return df_final


# ------------------------------------------------------------------------------
# 3. MAIN STREAMLIT APPLICATION INTERFACE
# ------------------------------------------------------------------------------

st.title("Day-by-Day Revenue Management Report")
st.markdown("Upload your raw PMS/IDeaS data file to transform date formats and build the master Day-by-Day layout.")

uploaded_file = st.file_uploader(
    "Choose a file (CSV or XLSX)",
    type=["csv", "xlsx"],
    key="master_file_uploader"
)

if uploaded_file is not None:
    try:
        with st.spinner("Processing report and formatting columns..."):
            processed_df = process_uploaded_file(uploaded_file)

        st.success("File successfully processed!")

        # Display Data Summary Metrics
        st.subheader("Summary View")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Days Reported", len(processed_df))
        m2.metric("Total Rooms OTB", f"{processed_df['Rooms Current'].sum():,.0f}" if processed_df['Rooms Current'].notnull().any() else "N/A")
        m3.metric("Total Revenue OTB", f"${processed_df['Rev Current'].sum():,.2f}" if processed_df['Rev Current'].notnull().any() else "N/A")

        # Main Table Preview
        st.subheader("Day-by-Day Master Grid")
        st.dataframe(processed_df, use_container_width=True)

        # Download Section
        st.subheader("Export Formatted Reports")
        col_csv, col_excel = st.columns(2)

        # CSV Download (Flattened single-row headers)
        csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
        col_csv.download_button(
            label="📥 Download CSV Report",
            data=csv_bytes,
            file_name="day_by_day_report.csv",
            mime="text/csv",
            use_container_width=True
        )

        # Excel Download
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            processed_df.to_excel(writer, index=False, sheet_name="Day_by_Day")
        
        col_excel.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="day_by_day_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")
else:
    st.info("Please upload a file to view and export your Day-by-Day report.")
