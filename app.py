import io
import openpyxl
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Day-by-Day Pickup Report", layout="wide")
st.title("📅 Day-by-Day (DD) Pickup Grid")

# ------------------------------------------------------------------------------
# 1. SIDEBAR UPLOADERS (EXACT USER ORDER)
# ------------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload Raw Data")
    pcdc_file = st.file_uploader("1. IDeaS PCDC", type=["csv", "xlsx"])
    extract_file = st.file_uploader("2. IDeaS Data Extraction", type=["csv", "xlsx"])
    market_file = st.file_uploader("3. IDeaS Market Segment", type=["csv", "xlsx"])
    lh_file = st.file_uploader("4. Lighthouse Rate Shop", type=["csv", "xlsx"])
    synxis_file = st.file_uploader("5. SynXis Rate Plan", type=["csv", "xlsx"])
    prior_file = st.file_uploader(
        "6. Yesterday's Pickup Report", type=["xlsx", "xlsm", "xls"]
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
df_pcdc_biz = read_uploaded_file(pcdc_file)
df_extract = read_uploaded_file(extract_file)
df_market = read_uploaded_file(market_file)
df_lh = read_uploaded_file(lh_file)
df_synxis = read_uploaded_file(synxis_file)

if df_pcdc_biz is None or df_extract is None:
    st.info(
        "👈 Please upload **1. IDeaS PCDC** and **2. IDeaS Data Extraction** to populate the Day-by-Day grid."
    )
    st.dataframe(pd.DataFrame(columns=columns), use_container_width=True)
else:
    # Identify date columns
    date_col_extract = next(
        (c for c in df_extract.columns if "Date" in str(c)), df_extract.columns[0]
    )
    df_extract["Date_Clean"] = pd.to_datetime(df_extract[date_col_extract])

    date_col_pcdc = next(
        (c for c in df_pcdc_biz.columns if "Date" in str(c)),
        df_pcdc_biz.columns[0],
    )
    df_pcdc_biz["Date_Clean"] = pd.to_datetime(df_pcdc_biz[date_col_pcdc])

    # Date range generation: Start from Data Extraction min date -> Dec 31 current year
    start_date = df_extract["Date_Clean"].min()
    current_year = pd.Timestamp.now().year
    end_date = pd.Timestamp(year=current_year, month=12, day=31)
    date_range = pd.date_range(start=start_date, end=end_date, freq="D")

    # Base date grid merge
    base_df = pd.DataFrame({"Date_Clean": date_range})
    merged_extract = pd.merge(base_df, df_extract, on="Date_Clean", how="left")
    merged_pcdc = pd.merge(base_df, df_pcdc_biz, on="Date_Clean", how="left")

    df_dd = pd.DataFrame(index=range(len(date_range)), columns=columns)

    today = pd.to_datetime("today").normalize()
    df_dd[("Date Info", "Date")] = date_range.strftime("%Y-%m-%d")
    df_dd[("Date Info", "DOW")] = date_range.strftime("%a")
    df_dd[("Date Info", "Days Left")] = (date_range - today).days

    def get_col(df, idx):
        if idx < len(df.columns):
            return pd.to_numeric(df.iloc[:, idx], errors="coerce").fillna(0)
        return pd.Series(0, index=df.index)

    # --- 1. PCDC LOOKUPS ---
    trans_current = get_col(merged_pcdc, 3)
    trans_change = get_col(merged_pcdc, 4)
    group_current = get_col(merged_pcdc, 5)
    group_change = get_col(merged_pcdc, 6)
    group_blocked = get_col(merged_pcdc, 7)
    group_pu = get_col(merged_pcdc, 9)

    trans_rn = get_col(merged_pcdc, 10)
    group_rn = get_col(merged_pcdc, 12)
    trans_rev = get_col(merged_pcdc, 18)
    group_rev = get_col(merged_pcdc, 20)

    # --- 2. DATA EXTRACTION LOOKUPS ---
    tot_otb_stly = get_col(merged_extract, 5)
    grp_otb_stly = get_col(merged_extract, 7)
    trans_otb_stly = get_col(merged_extract, 9)

    # --- POPULATE GRID ---
    df_dd[("Rooms Sold - Total Transient", "Current")] = trans_current
    df_dd[("Rooms Sold - Total Transient", "Change")] = trans_change

    df_dd[("Rooms Sold - Total Group", "Current")] = group_current
    df_dd[("Rooms Sold - Total Group", "Change")] = group_change
    df_dd[("Rooms Sold - Total Group", "Blocked")] = group_blocked
    df_dd[("Rooms Sold - Total Group", "P/U")] = group_pu
    df_dd[("Rooms Sold - Total Group", "Remaining")] = (
        group_blocked - group_pu
    ).clip(lower=0)

    df_dd[("Rooms Sold - Total Hotel", "Current")] = trans_current + group_current
    df_dd[("Rooms Sold - Total Hotel", "Change")] = trans_change + group_change

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
