import io
import pandas as pd
import streamlit as st

# 1. Define your exact requested column order
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

def process_day_by_day_grid(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans date formats, maps raw export fields to your exact target headers,
    and forces strict column ordering.
    """
    df = raw_df.copy()

    # --- Step 1: Strip Time Component from Date Column ---
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"]).dt.date

    # --- Step 2: Handle MultiIndex Columns if present ---
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(c) for c in col).strip() for col in df.columns.values]

    # --- Step 3: Map Raw Export Source Fields to Your Target Names ---
    field_mappings = {
        # IDeaS Data Extraction field mapping
        "System Total Demand - Total This Year": "Remaining Demand - Total Hotel",
        "PickUp": "P/U",
        "Blocked Rooms": "Blocked",
        # Generic multi-source fallback mappings (adjust if your merged keys differ)
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
    df = df.rename(columns=field_mappings)

    # --- Step 4: Enforce Exact Header List & Ordering ---
    # Ensure all target columns exist (fill missing ones with None/0 to avoid KeyError)
    for col in DESIRED_DD_HEADERS:
        if col not in df.columns:
            df[col] = None

    # Reorder columns to strictly match DESIRED_DD_HEADERS and drop unrequested columns
    df_final = df[DESIRED_DD_HEADERS]

    return df_final


# --- Streamlit Application UI & Downloads ---

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
    st.dataframe(processed_df)

    col1, col2 = st.columns(2)

    # Download Option 1: Clean Single-Row Flattened CSV
    csv_bytes = processed_df.to_csv(index=False).encode("utf-8")
    col1.download_button(
        label="Download Flattened CSV",
        data=csv_bytes,
        file_name="day_by_day_report.csv",
        mime="text/csv"
    )

    # Download Option 2: Excel (.xlsx) Format
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        processed_df.to_excel(writer, index=False, sheet_name="Day_by_Day")
    
    col2.download_button(
        label="Download Excel (.xlsx)",
        data=excel_buffer.getvalue(),
        file_name="day_by_day_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
