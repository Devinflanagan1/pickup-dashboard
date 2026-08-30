import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Day-by-Day Revenue Report Generator",
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

def process_day_by_day_grid(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans date formats, maps raw export fields to target headers,
    and forces strict column ordering.
    """
    df = raw_df.copy()

    # --- Step 1: Strip Time Component from Date Column ---
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"]).dt.date

    # --- Step 2: Handle MultiIndex Columns if present ---
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(c) for c in col).strip() for col in df.columns.values]

    # --- Step 3: Map Raw Export Source Fields ---
    field_mappings = {
        "System Total Demand - Total This Year": "Remaining Demand - Total Hotel",
        "PickUp": "P/U",
        "Blocked Rooms": "Blocked",
        "Rooms_Current": "Rooms Sold Total Hotel Current",
        "Rooms_Change": "Rooms Sold Total Hotel Change",
    }
    df = df.rename(columns=field_mappings)

    # --- Step 4: Enforce Exact Header List & Ordering ---
    for col in DESIRED_DD_HEADERS:
        if col not in df.columns:
            df[col] = None

    # Reorder columns to strictly match DESIRED_DD_HEADERS
    df_final = df[DESIRED_DD_HEADERS]

    return df_final


# ------------------------------------------------------------------------------
# 2. STREAMLIT UI & DOWNLOADS
# ------------------------------------------------------------------------------

st.title("Day-by-Day Revenue Report Generator")

uploaded_file = st.file_uploader("Upload Master Data File", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Read input file
    if uploaded_file.name.endswith(".csv"):
        raw_data = pd.read_csv(uploaded_file)
    else:
        raw_data = pd.read_excel(uploaded_file)

    # Process DataFrame into the exact structure
    processed_df = process_day_by_day_grid(raw_data)

    st.subheader("Day-by-Day Report Preview")
    st.dataframe(processed_df, use_container_width=True)

    col1, col2 = st.columns(2)

    # Download Option 1: Clean Single-Row Flattened CSV
    csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
    col1.download_button(
        label="Download Flattened CSV",
        data=csv_bytes,
        file_name="day_by_day_report.csv",
        mime="text/csv",
        use_container_width=True
    )

    # Download Option 2: Excel (.xlsx) Format
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
    st.info("Please upload your master data file to get started.")
