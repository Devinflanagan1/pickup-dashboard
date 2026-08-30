import io
import openpyxl
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Day-by-Day Pickup Report", layout="wide")
st.title("📅 Day-by-Day (DD) Pickup Grid")

# ------------------------------------------------------------------------------
# 1. SIDEBAR UPLOADERS (SEPARATE FILES)
# ------------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload Raw Data")
    extract_file = st.file_uploader(
        "1. IDeaS Data Extraction File", type=["csv", "xlsx"]
    )
    pcdc_file = st.file_uploader(
        "2. IDeaS PCDC Business Type File", type=["csv", "xlsx"]
    )
    lh_file = st.file_uploader("3. Lighthouse Comp Set File", type=["csv", "xlsx"])
    synxis_file = st.file_uploader("4. SynXis Rate Plan File", type=["csv", "xlsx"])
    prior_file = st.file_uploader(
        "5. Yesterday's Pickup Report", type=["xlsx", "xlsm", "xls"]
    )


def read_uploaded_file(file_obj):
    if file_obj is None:
        return None
    if file_obj.name.endswith(".csv"):
        return pd.read_csv(file_obj)
    return pd.read_excel(file_obj)


# ------------------------------------------------------------------------------
# 2. MULTIINDEX HEADER STRUCTURE
# ------------------------------------------------------------------------------
columns = pd.MultiIndex.from_tuples(
    [
        ("Date Info", "DOW"),
        ("Date Info", "Date"),
        ("Date Info", "Days Left"),
        ("Date Info", "Events"),
        ("Pricing & Capacity", "Rooms Left to Sell"),
        ("Pricing & Capacity", "BAR Current"),
        ("Pricing & Capacity", "Comp Set Avg"),
        ("Pricing & Capacity", "Last Room Value"),
        ("Competitor Shops", "21c Museum Hotel"),
        ("Competitor Shops", "Motto By Hilton"),
        ("Competitor Shops", "AC Hotel by Marriott"),
        ("Competitor Shops", "DoubleTree Suites"),
        ("House Status", "OOO"),
        ("House Status", "Ovrbk"),
        ("Rooms Sold - Total Hotel", "Current"),
        ("Rooms Sold - Total Hotel", "Change"),
        ("Rooms Sold - Total Transient", "Current"),
        ("Rooms Sold - Total Transient", "Change"),
        ("Rooms Sold - Total Group", "Current"),
        ("Rooms Sold - Total Group", "Change"),
        ("Rooms Sold - Total Group", "Blocked"),
        ("Rooms Sold - Total Group", "P/U"),
        ("Rooms Sold - Total Group", "Remaining"),
        ("Rooms OTB STLY", "Total OTB STLY"),
        ("Rooms OTB STLY", "Variance (TY - STLY)"),
        ("Rooms OTB STLY", "Transient OTB STLY"),
        ("Rooms OTB STLY", "Trans Variance (TY - STLY)"),
        ("Rooms OTB STLY", "Group OTB STLY"),
        ("Rooms OTB STLY", "Group Variance (TY - STLY)"),
        ("Remaining Demand", "Total Hotel"),
        ("Remaining Demand", "Total Transient"),
        ("Remaining Demand", "Total Group"),
        ("Occupancy Forecast", "Total Hotel"),
        ("Occupancy Forecast", "Total Transient"),
        ("Occupancy Forecast", "Total Group"),
        ("Occupancy Forecast %", "Total Hotel"),
        ("Occupancy Forecast %", "Total Transient"),
        ("Occupancy Forecast %", "Total Group"),
        ("Booked ADR (USD)", "Total Hotel"),
        ("Booked ADR (USD)", "Total Transient"),
        ("Booked ADR (USD)", "Total Group"),
        ("Yield Metrics", "Estimated ADR"),
    ]
)

# ------------------------------------------------------------------------------
# 3. DATA PROCESSING & MERGING ENGINE
# ------------------------------------------------------------------------------
df_extract = read_uploaded_file(extract_file)
df_pcdc_biz = read_uploaded_file(pcdc_file)

if df_extract is None or df_pcdc_biz is None:
    st.info(
        "👈 Upload both **1. IDeaS Data Extraction** and **2. IDeaS PCDC Business Type** in the sidebar to build the grid."
    )
    st.dataframe(pd.DataFrame(columns=columns), use_container_width=True)
else:
    # Identify date column in Data Extraction
    date_col_extract = next(
        (c for c in df_extract.columns if "Date" in str(c)), df_extract.columns[0]
    )
    df_extract["Date_Clean"] = pd.to_datetime(df_extract[date_col_extract])

    # Identify date column in PCDC
    date_col_pcdc = next(
        (c for c in df_pcdc_biz.columns if "Date" in str(c)),
        df_pcdc_biz.columns[0],
    )
    df_pcdc_biz["Date_Clean"] = pd.to_datetime(df_pcdc_biz[date_col_pcdc])

    # Generate continuous date range: Start date from import -> Dec 31 current year
    start_date = df_extract["Date_Clean"].min()
    current_year = pd.Timestamp.now().year
    end_date = pd.Timestamp(year=current_year, month=12, day=31)
    date_range = pd.date_range(start=start_date, end=end_date, freq="D")

    # Merge individual files to master date sequence
    base_df = pd.DataFrame({"Date_Clean": date_range})
    merged_extract = pd.merge(
        base_df, df_extract, on="Date_Clean", how="left"
    )
    merged_pcdc = pd.merge(
        base_df, df_pcdc_biz, on="Date_Clean", how="left"
    )

    # Initialize master grid output
    df_dd = pd.DataFrame(index=range(len(date_range)), columns=columns)

    # Date Info Calculations
    today = pd.to_datetime("today").normalize()
    df_dd[("Date Info", "Date")] = date_range.strftime("%Y-%m-%d")
    df_dd[("Date Info", "DOW")] = date_range.strftime("%a")
    df_dd[("Date Info", "Days Left")] = (date_range - today).days

    # Helper function for positional index retrieval
    def get_col(df, idx):
        if idx < len(df.columns):
            return pd.to_numeric(df.iloc[:, idx], errors="coerce").fillna(0)
        return pd.Series(0, index=df.index)

    # --- PCDC BUSINESS TYPE LOOKUPS ---
    trans_current = get_col(merged_pcdc, 3)  # Col 4
    trans_change = get_col(merged_pcdc, 4)  # Col 5
    group_current = get_col(merged_pcdc, 5)  # Col 6
    group_change = get_col(merged_pcdc, 6)  # Col 7
    group_blocked = get_col(merged_pcdc, 7)  # Col 8
    group_pu = get_col(merged_pcdc, 9)  # Col 10

    # ADR Numerator/Denominator columns from PCDC
    trans_rn = get_col(merged_pcdc, 10)  # Col 11
    group_rn = get_col(merged_pcdc, 12)  # Col 13
    trans_rev = get_col(merged_pcdc, 18)  # Col 19
    group_rev = get_col(merged_pcdc, 20)  # Col 21

    # --- DATA EXTRACTION LOOKUPS ---
    tot_otb_stly = get_col(merged_extract, 5)  # Col 6
    grp_otb_stly = get_col(merged_extract, 7)  # Col 8
    trans_otb_stly = get_col(merged_extract, 9)  # Col 10

    # --- POPULATE GRID ---
    # Transient
    df_dd[("Rooms Sold - Total Transient", "Current")] = trans_current
    df_dd[("Rooms Sold - Total Transient", "Change")] = trans_change

    # Group
    df_dd[("Rooms Sold - Total Group", "Current")] = group_current
    df_dd[("Rooms Sold - Total Group", "Change")] = group_change
    df_dd[("Rooms Sold - Total Group", "Blocked")] = group_blocked
    df_dd[("Rooms Sold - Total Group", "P/U")] = group_pu
    df_dd[("Rooms Sold - Total Group", "Remaining")] = (
        group_blocked - group_pu
    ).clip(lower=0)

    # Total Hotel (= Transient + Group)
    df_dd[("Rooms Sold - Total Hotel", "Current")] = (
        trans_current + group_current
    )
    df_dd[("Rooms Sold - Total Hotel", "Change")] = trans_change + group_change

    # Rooms OTB STLY & Variances
    df_dd[("Rooms OTB STLY", "Total OTB STLY")] = tot_otb_stly
    df_dd[("Rooms OTB STLY", "Variance (TY - STLY)")] = (
        df_dd[("Rooms Sold - Total Hotel", "Current")] - tot_otb_stly
    )

    df_dd[("Rooms OTB STLY", "Transient OTB STLY")] = trans_otb_stly
    df_dd[("Rooms OTB STLY", "Trans Variance (TY - STLY)")] = (
        trans_current - trans_otb_stly
    )

    df_dd[("Rooms OTB STLY", "Group OTB STLY")] = grp_otb_stly
    df_dd[("Rooms OTB STLY", "Group Variance (TY - STLY)")] = (
        group_current - grp_otb_stly
    )

    # Estimated ADR
    total_rev = trans_rev + group_rev
    total_rn = trans_rn + group_rn
    df_dd[("Yield Metrics", "Estimated ADR")] = np.where(
        total_rn > 0, total_rev / total_rn, 0.0
    )

    # ------------------------------------------------------------------------------
    # 4. DISPLAY & EXPORT GRID
    # ------------------------------------------------------------------------------
    st.subheader("Day-by-Day Master Grid")
    st.dataframe(df_dd, use_container_width=True, height=600)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_dd.to_excel(writer, sheet_name="DD")

    st.download_button(
        label="📥 Download DD Excel Report",
        data=buffer.getvalue(),
        file_name="Day_By_Day_Pickup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
