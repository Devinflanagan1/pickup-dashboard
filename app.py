import io
import pandas as pd
import streamlit as st

# =============================================================================
# PAGE CONFIGURATION
# =============================================================================
st.set_page_config(
    page_title="2026 Master Revenue Management Dashboard",
    page_icon="🏨",
    layout="wide",
)


# =============================================================================
# 1. LIGHTHOUSE RATE SHOP PARSER
# =============================================================================
def parse_lighthouse_rate_shop(lh_file):
    if lh_file is None:
        return {}

    try:
        lh_file.seek(0)
    except Exception:
        pass

    excel_file = pd.ExcelFile(lh_file)
    sheet_names = excel_file.sheet_names
    tabs_data = {}

    def get_sheet(target_key):
        for s in sheet_names:
            if target_key.lower() in s.lower().replace(".", ""):
                return s
        return None

    target_tabs = {
        "Rates": get_sheet("rates"),
        "vs. Yesterday": get_sheet("vs yesterday"),
        "vs. 3 Days Ago": get_sheet("vs 3 days ago"),
        "vs. 7 Days Ago": get_sheet("vs 7 days ago"),
        "Overview": get_sheet("overview"),
    }

    comp_map = {
        "21c Museum Hotel Bentonville - MGallery": [
            "21c Museum Hotel",
            "21c Museum Hotel Bentonville",
        ],
        "Motto By Hilton Bentonville Downtown": [
            "Motto By Hilton",
            "Motto Bentonville",
        ],
        "AC Hotel by Marriott Bentonville": ["AC Hotel", "AC Hotel Bentonville"],
        "DoubleTree Suites by Hilton Bentonville": [
            "DoubleTree",
            "DoubleTree Suites Bentonville",
        ],
    }

    for view_label, sheet_name in target_tabs.items():
        if sheet_name is None:
            continue

        try:
            lh_file.seek(0)
        except Exception:
            pass

        df_temp = pd.read_excel(lh_file, sheet_name=sheet_name)
        header_row_idx = 0
        for i, row in df_temp.head(10).iterrows():
            row_str = " ".join([str(val) for val in row.values])
            if "Day Date" in row_str or "Date" in row_str:
                header_row_idx = i
                break

        try:
            lh_file.seek(0)
        except Exception:
            pass

        df_clean = pd.read_excel(
            lh_file, sheet_name=sheet_name, header=header_row_idx
        )
        df_clean.columns = df_clean.columns.astype(str).str.strip()

        date_col = next((c for c in df_clean.columns if "Date" in c), None)
        if not date_col:
            continue

        df_clean["Occupancy Date"] = pd.to_datetime(
            df_clean[date_col], errors="coerce"
        )
        df_clean = df_clean.dropna(subset=["Occupancy Date"])

        parsed_df = pd.DataFrame({"Occupancy Date": df_clean["Occupancy Date"]})

        for alias, search_terms in comp_map.items():
            found_col = None
            for term in search_terms:
                found_col = next(
                    (
                        c
                        for c in df_clean.columns
                        if term.lower() in c.lower()
                    ),
                    None,
                )
                if found_col:
                    break

            if found_col:
                parsed_df[alias] = pd.to_numeric(
                    df_clean[found_col], errors="coerce"
                ).fillna(0)
            else:
                parsed_df[alias] = 0.0

        comp_cols = list(comp_map.keys())
        parsed_df["Comp Set Avg"] = parsed_df[comp_cols].mean(axis=1).round(2)
        tabs_data[view_label] = parsed_df

    return tabs_data


# =============================================================================
# 2. IDEAS PCDC REPORT PARSER
# =============================================================================
def parse_ideas_pcdc(pcdc_file):
    if pcdc_file is None:
        return pd.DataFrame()

    try:
        pcdc_file.seek(0)
    except Exception:
        pass

    if pcdc_file.name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(pcdc_file)
    else:
        try:
            df = pd.read_csv(pcdc_file)
        except Exception:
            pcdc_file.seek(0)
            df = pd.read_csv(pcdc_file, encoding="latin1")

    df.columns = df.columns.astype(str).str.strip()
    date_col = next((c for c in df.columns if "Date" in c), None)

    if date_col:
        df["Occupancy Date"] = pd.to_datetime(df[date_col], errors="coerce")
        return df.dropna(subset=["Occupancy Date"])

    return pd.DataFrame()


# =============================================================================
# 3. IDEAS DATA EXTRACTION PARSER
# =============================================================================
def parse_ideas_data_extraction(extract_file):
    if extract_file is None:
        return pd.DataFrame()

    try:
        extract_file.seek(0)
    except Exception:
        pass

    if extract_file.name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(extract_file)
    else:
        try:
            df = pd.read_csv(extract_file)
        except Exception:
            extract_file.seek(0)
            df = pd.read_csv(extract_file, encoding="latin1")

    df.columns = df.columns.astype(str).str.strip()
    date_col = next((c for c in df.columns if "Date" in c), None)

    if date_col:
        df["Occupancy Date"] = pd.to_datetime(df[date_col], errors="coerce")
        return df.dropna(subset=["Occupancy Date"])

    return pd.DataFrame()


# =============================================================================
# 4. MARKET SEGMENTATION PARSER
# =============================================================================
def parse_market_segmentation(market_seg_file):
    if market_seg_file is None:
        return pd.DataFrame()

    try:
        market_seg_file.seek(0)
    except Exception:
        pass

    if market_seg_file.name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(market_seg_file)
    else:
        try:
            df = pd.read_csv(market_seg_file)
        except Exception:
            market_seg_file.seek(0)
            df = pd.read_csv(market_seg_file, encoding="latin1")

    df.columns = df.columns.astype(str).str.strip()
    date_col = next((c for c in df.columns if "Date" in c), None)

    if date_col:
        df["Occupancy Date"] = pd.to_datetime(df[date_col], errors="coerce")
        return df.dropna(subset=["Occupancy Date"])

    return pd.DataFrame()


# =============================================================================
# 5. SYNXIS RATE PLAN PARSER
# =============================================================================
def parse_synxis_rate_plan(synxis_file, selected_statuses=None):
    if synxis_file is None:
        return None, [], pd.DataFrame()

    try:
        synxis_file.seek(0)
    except Exception:
        pass

    if synxis_file.name.endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(synxis_file)
    else:
        try:
            df_raw = pd.read_csv(synxis_file)
        except Exception:
            synxis_file.seek(0)
            df_raw = pd.read_csv(synxis_file, encoding="latin1")

    df_raw.columns = df_raw.columns.astype(str).str.strip()

    status_col = next(
        (c for c in df_raw.columns if "Rez_Status" in c or "Status" in c),
        "Rez_Status",
    )

    available_statuses = []
    if status_col in df_raw.columns:
        available_statuses = (
            df_raw[status_col].dropna().astype(str).unique().tolist()
        )

    if selected_statuses and status_col in df_raw.columns:
        df_filtered = df_raw[df_raw[status_col].isin(selected_statuses)].copy()
    else:
        df_filtered = df_raw.copy()

    arr_col = next(
        (c for c in df_filtered.columns if "Arrival_Dt" in c or "Arrival" in c),
        "Arrival_Dt",
    )
    dep_col = next(
        (c for c in df_filtered.columns if "Depart_Dt" in c or "Depart" in c),
        "Depart_Dt",
    )

    df_filtered["Arrival_Dt"] = pd.to_datetime(
        df_filtered[arr_col], errors="coerce"
    )
    df_filtered["Depart_Dt"] = pd.to_datetime(
        df_filtered[dep_col], errors="coerce"
    )
    df_filtered = df_filtered.dropna(subset=["Arrival_Dt", "Depart_Dt"])

    exploded_rows = []
    for _, row in df_filtered.iterrows():
        stay_dates = pd.date_range(
            start=row["Arrival_Dt"],
            end=row["Depart_Dt"] - pd.Timedelta(days=1),
            freq="D",
        )

        room_qty = pd.to_numeric(row.get("Room_Qty", 1), errors="coerce")
        if pd.isna(room_qty) or room_qty <= 0:
            room_qty = 1

        avg_rate = pd.to_numeric(row.get("Rez_Avg_R", 0), errors="coerce")
        if pd.isna(avg_rate):
            avg_rate = 0.0

        rate_type = row.get("Rate_Type", "Unknown")
        rate_code = row.get("Rate_Cate", "Unknown")
        rez_status = row.get(status_col, "Unknown")

        for single_date in stay_dates:
            exploded_rows.append(
                {
                    "Occupancy Date": single_date,
                    "Room_Qty": room_qty,
                    "Daily_Revenue": avg_rate * room_qty,
                    "Rate_Type": rate_type,
                    "Rate_Code": rate_code,
                    "Rez_Status": rez_status,
                }
            )

    if not exploded_rows:
        return (
            pd.DataFrame(
                columns=[
                    "Occupancy Date",
                    "Rate_Plan_Rooms",
                    "Rate_Plan_Rev",
                    "Rate_Plan_ADR",
                ]
            ),
            available_statuses,
            pd.DataFrame(),
        )

    df_exploded = pd.DataFrame(exploded_rows)

    df_daily = (
        df_exploded.groupby("Occupancy Date")
        .agg(
            Rate_Plan_Rooms=("Room_Qty", "sum"),
            Rate_Plan_Rev=("Daily_Revenue", "sum"),
        )
        .reset_index()
    )

    df_daily["Rate_Plan_ADR"] = (
        df_daily["Rate_Plan_Rev"]
        / df_daily["Rate_Plan_Rooms"].replace(0, 1)
    ).round(2)

    return df_daily, available_statuses, df_exploded


# =============================================================================
# 6. BASE 2026 GRID GENERATOR
# =============================================================================
def build_2026_base_grid():
    dates = pd.date_range(start="2026-01-01", end="2026-12-31", freq="D")
    base_df = pd.DataFrame({"Occupancy Date": dates})
    base_df["DOW"] = base_df["Occupancy Date"].dt.strftime("%a")
    base_df["Date"] = base_df["Occupancy Date"].dt.strftime("%Y-%m-%d")

    today = pd.Timestamp.now().normalize()
    base_df["Days Left"] = (base_df["Occupancy Date"] - today).dt.days
    base_df["Events"] = ""

    return base_df


# =============================================================================
# 7. STREAMLIT APP ENGINE
# =============================================================================
def main():
    st.title("🏨 2026 Master Revenue Management System")

    # --- SIDEBAR UPLOADERS (ORDERED: LIGHTHOUSE, PCDC, EXTRACTION, SEGMENTATION, SYNXIS) ---
    st.sidebar.header("📁 Data Source Uploads")
    lh_file = st.sidebar.file_uploader(
        "1. Lighthouse Rate Shop (.xlsx)", type=["xlsx"]
    )
    pcdc_file = st.sidebar.file_uploader(
        "2. IDeaS PCDC Report (.csv / .xlsx)", type=["csv", "xlsx"]
    )
    extract_file = st.sidebar.file_uploader(
        "3. IDeaS Data Extraction (.csv / .xlsx)", type=["csv", "xlsx"]
    )
    market_seg_file = st.sidebar.file_uploader(
        "4. Market Segmentation (.csv / .xlsx)", type=["csv", "xlsx"]
    )
    synxis_file = st.sidebar.file_uploader(
        "5. SynXis Rate Plan Export (.csv / .xlsx)", type=["csv", "xlsx"]
    )

    st.sidebar.markdown("---")
    st.sidebar.header("🕹️ Rate Shop Controls")
    rate_view_option = st.sidebar.radio(
        "Competitor View Mode:",
        options=["Rates", "vs. Yesterday", "vs. 3 Days Ago", "vs. 7 Days Ago"],
        index=0,
    )

    # --- SETUP TABS ---
    tab_dashboard, tab_rate_plan = st.tabs(
        ["📊 Main Day-by-Day Dashboard", "📑 SynXis Rate Plan Analysis"]
    )

    # Parse SynXis Data
    selected_status_list = []
    df_synxis_daily = None
    df_synxis_exploded = pd.DataFrame()
    available_statuses = []

    if synxis_file is not None:
        _, available_statuses, _ = parse_synxis_rate_plan(
            synxis_file, selected_statuses=None
        )

        st.sidebar.markdown("---")
        st.sidebar.header("🛎️ SynXis Status Filter")
        default_selected = [
            s for s in available_statuses if "confirm" in s.lower()
        ]
        if not default_selected and available_statuses:
            default_selected = available_statuses

        selected_status_list = st.sidebar.multiselect(
            "Active Reservation Statuses:",
            options=available_statuses,
            default=default_selected,
        )

        (
            df_synxis_daily,
            _,
            df_synxis_exploded,
        ) = parse_synxis_rate_plan(
            synxis_file, selected_statuses=selected_status_list
        )

    # Parse Lighthouse Data
    lh_tabs = parse_lighthouse_rate_shop(lh_file) if lh_file else {}
    df_lh_selected = lh_tabs.get(rate_view_option, pd.DataFrame())

    # Build Base Grid
    base_df = build_2026_base_grid()

    if not df_lh_selected.empty:
        base_df = pd.merge(
            base_df, df_lh_selected, on="Occupancy Date", how="left"
        ).fillna(0)
    else:
        for comp in [
            "21c Museum Hotel Bentonville - MGallery",
            "Motto By Hilton Bentonville Downtown",
            "AC Hotel by Marriott Bentonville",
            "DoubleTree Suites by Hilton Bentonville",
            "Comp Set Avg",
        ]:
            base_df[comp] = 0.0

    if df_synxis_daily is not None and not df_synxis_daily.empty:
        base_df = pd.merge(
            base_df, df_synxis_daily, on="Occupancy Date", how="left"
        ).fillna(0)
    else:
        base_df["Rate_Plan_Rooms"] = 0
        base_df["Rate_Plan_Rev"] = 0.0
        base_df["Rate_Plan_ADR"] = 0.0

    # Execute Parsers for PCDC, Data Extraction, Market Segmentation
    df_pcdc = parse_ideas_pcdc(pcdc_file)
    df_extract = parse_ideas_data_extraction(extract_file)
    df_market_seg = parse_market_segmentation(market_seg_file)

    # Initialize placeholder metrics
    metric_cols = [
        "Rem_Demand_Total",
        "Rem_Demand_Trans",
        "Rem_Demand_Group",
        "Occ_Fcst_Total",
        "Occ_Fcst_Trans",
        "Occ_Fcst_Group",
        "Occ_Fcst_Pct_Total",
        "Occ_Fcst_Pct_Trans",
        "Occ_Fcst_Pct_Group",
        "Booked_ADR_Total",
        "Booked_ADR_Trans",
        "Booked_ADR_Group",
        "Rooms_Left_to_Sell",
        "BAR_Current",
        "Estimated_ADR",
        "OOO",
        "Ovrbk",
        "Sold_Total_Current",
        "Sold_Total_Change",
        "Sold_Trans_Current",
        "Sold_Trans_Change",
        "Sold_Group_Current",
        "Sold_Group_Change",
        "Sold_Group_Blocked",
        "Sold_Group_PU",
        "Sold_Group_Remaining",
        "OTB_STLY_Total",
        "Variance_Total_STLY",
        "OTB_STLY_Trans",
        "Variance_Trans_STLY",
        "OTB_STLY_Group",
        "Variance_Group_STLY",
    ]
    for col in metric_cols:
        base_df[col] = 0.0

    # Map SynXis OTB into Sold Total Current
    base_df["Sold_Total_Current"] = base_df["Rate_Plan_Rooms"]
    base_df["Booked_ADR_Total"] = base_df["Rate_Plan_ADR"]

    # =========================================================================
    # TAB 1: MAIN DASHBOARD GRID
    # =========================================================================
    with tab_dashboard:
        columns_tuple = [
            ("", "", "DOW"),
            ("", "", "Date"),
            ("", "", "Days Left"),
            ("", "", "Events"),
            ("Remaining Demand", "Total Hotel", ""),
            ("Remaining Demand", "Total Transient", ""),
            ("Remaining Demand", "Total Group", ""),
            ("Occupancy Forecast", "Total Hotel", ""),
            ("Occupancy Forecast", "Total Transient", ""),
            ("Occupancy Forecast", "Total Group", ""),
            ("Occupancy Forecast %", "Total Hotel", ""),
            ("Occupancy Forecast %", "Total Transient", ""),
            ("Occupancy Forecast %", "Total Group", ""),
            ("Booked ADR(USD)", "Total Hotel", ""),
            ("Booked ADR(USD)", "Total Transient", ""),
            ("Booked ADR(USD)", "Total Group", ""),
            ("Pricing & Capacity", "Rooms Left to Sell", ""),
            ("Pricing & Capacity", "BAR", "Current"),
            ("Pricing & Capacity", "Comp Set Avg", ""),
            ("Pricing & Capacity", "Last Room Value", "Estimated ADR"),
            (
                "Competitor Shops",
                "21c Museum Hotel Bentonville - MGallery",
                "",
            ),
            ("Competitor Shops", "Motto By Hilton Bentonville Downtown", ""),
            ("Competitor Shops", "AC Hotel by Marriott Bentonville", ""),
            (
                "Competitor Shops",
                "DoubleTree Suites by Hilton Bentonville",
                "",
            ),
            ("Inventory", "OOO", ""),
            ("Inventory", "Ovrbk", ""),
            ("Rooms Sold", "Total Hotel", "Current"),
            ("Rooms Sold", "Total Hotel", "Change"),
            ("Rooms Sold", "Total Transient", "Current"),
            ("Rooms Sold", "Total Transient", "Change"),
            ("Rooms Sold", "Total Group", "Current"),
            ("Rooms Sold", "Total Group", "Change"),
            ("Rooms Sold", "Total Group", "Blocked"),
            ("Rooms Sold", "Total Group", "P/U"),
            ("Rooms Sold", "Total Group", "Remaining"),
            ("Rooms OTB STLY", "Total Hotel", "OTB STLY"),
            ("Rooms OTB STLY", "Total Hotel", "Variance (TY - STLY)"),
            ("Rooms OTB STLY", "Transient", "OTB STLY"),
            ("Rooms OTB STLY", "Transient", "Variance (TY - STLY)"),
            ("Rooms OTB STLY", "Group", "OTB STLY"),
            ("Rooms OTB STLY", "Group", "Variance (TY - STLY)"),
        ]

        multi_index_cols = pd.MultiIndex.from_tuples(columns_tuple)
        dash_df = pd.DataFrame(index=base_df.index, columns=multi_index_cols)

        # Populate Values
        dash_df[("", "", "DOW")] = base_df["DOW"]
        dash_df[("", "", "Date")] = base_df["Date"]
        dash_df[("", "", "Days Left")] = base_df["Days Left"]
        dash_df[("", "", "Events")] = base_df["Events"]

        dash_df[("Remaining Demand", "Total Hotel", "")] = base_df[
            "Rem_Demand_Total"
        ]
        dash_df[("Remaining Demand", "Total Transient", "")] = base_df[
            "Rem_Demand_Trans"
        ]
        dash_df[("Remaining Demand", "Total Group", "")] = base_df[
            "Rem_Demand_Group"
        ]

        dash_df[("Occupancy Forecast", "Total Hotel", "")] = base_df[
            "Occ_Fcst_Total"
        ]
        dash_df[("Occupancy Forecast", "Total Transient", "")] = base_df[
            "Occ_Fcst_Trans"
        ]
        dash_df[("Occupancy Forecast", "Total Group", "")] = base_df[
            "Occ_Fcst_Group"
        ]

        dash_df[("Occupancy Forecast %", "Total Hotel", "")] = base_df[
            "Occ_Fcst_Pct_Total"
        ]
        dash_df[("Occupancy Forecast %", "Total Transient", "")] = base_df[
            "Occ_Fcst_Pct_Trans"
        ]
        dash_df[("Occupancy Forecast %", "Total Group", "")] = base_df[
            "Occ_Fcst_Pct_Group"
        ]

        dash_df[("Booked ADR(USD)", "Total Hotel", "")] = base_df[
            "Booked_ADR_Total"
        ]
        dash_df[("Booked ADR(USD)", "Total Transient", "")] = base_df[
            "Booked_ADR_Trans"
        ]
        dash_df[("Booked ADR(USD)", "Total Group", "")] = base_df[
            "Booked_ADR_Group"
        ]

        dash_df[("Pricing & Capacity", "Rooms Left to Sell", "")] = base_df[
            "Rooms_Left_to_Sell"
        ]
        dash_df[("Pricing & Capacity", "BAR", "Current")] = base_df[
            "BAR_Current"
        ]
        dash_df[("Pricing & Capacity", "Comp Set Avg", "")] = base_df[
            "Comp Set Avg"
        ]
        dash_df[
            ("Pricing & Capacity", "Last Room Value", "Estimated ADR")
        ] = base_df["Estimated_ADR"]

        dash_df[
            (
                "Competitor Shops",
                "21c Museum Hotel Bentonville - MGallery",
                "",
            )
        ] = base_df["21c Museum Hotel Bentonville - MGallery"]
        dash_df[
            ("Competitor Shops", "Motto By Hilton Bentonville Downtown", "")
        ] = base_df["Motto By Hilton Bentonville Downtown"]
        dash_df[
            ("Competitor Shops", "AC Hotel by Marriott Bentonville", "")
        ] = base_df["AC Hotel by Marriott Bentonville"]
        dash_df[
            (
                "Competitor Shops",
                "DoubleTree Suites by Hilton Bentonville",
                "",
            )
        ] = base_df["DoubleTree Suites by Hilton Bentonville"]

        dash_df[("Inventory", "OOO", "")] = base_df["OOO"]
        dash_df[("Inventory", "Ovrbk", "")] = base_df["Ovrbk"]

        dash_df[("Rooms Sold", "Total Hotel", "Current")] = base_df[
            "Sold_Total_Current"
        ]
        dash_df[("Rooms Sold", "Total Hotel", "Change")] = base_df[
            "Sold_Total_Change"
        ]
        dash_df[("Rooms Sold", "Total Transient", "Current")] = base_df[
            "Sold_Trans_Current"
        ]
        dash_df[("Rooms Sold", "Total Transient", "Change")] = base_df[
            "Sold_Trans_Change"
        ]
        dash_df[("Rooms Sold", "Total Group", "Current")] = base_df[
            "Sold_Group_Current"
        ]
        dash_df[("Rooms Sold", "Total Group", "Change")] = base_df[
            "Sold_Group_Change"
        ]
        dash_df[("Rooms Sold", "Total Group", "Blocked")] = base_df[
            "Sold_Group_Blocked"
        ]
        dash_df[("Rooms Sold", "Total Group", "P/U")] = base_df["Sold_Group_PU"]
        dash_df[("Rooms Sold", "Total Group", "Remaining")] = base_df[
            "Sold_Group_Remaining"
        ]

        dash_df[("Rooms OTB STLY", "Total Hotel", "OTB STLY")] = base_df[
            "OTB_STLY_Total"
        ]
        dash_df[
            ("Rooms OTB STLY", "Total Hotel", "Variance (TY - STLY)")
        ] = base_df["Variance_Total_STLY"]
        dash_df[("Rooms OTB STLY", "Transient", "OTB STLY")] = base_df[
            "OTB_STLY_Trans"
        ]
        dash_df[
            ("Rooms OTB STLY", "Transient", "Variance (TY - STLY)")
        ] = base_df["Variance_Trans_STLY"]
        dash_df[("Rooms OTB STLY", "Group", "OTB STLY")] = base_df[
            "OTB_STLY_Group"
        ]
        dash_df[
            ("Rooms OTB STLY", "Group", "Variance (TY - STLY)")
        ] = base_df["Variance_Group_STLY"]

        st.subheader("2026 Day-by-Day Master Grid")
        st.dataframe(dash_df, use_container_width=True, height=650)

    # =========================================================================
    # TAB 2: SYNXIS RATE PLAN ANALYSIS
    # =========================================================================
    with tab_rate_plan:
        st.subheader("📑 SynXis Rate Plan Production & Segmentation")

        if synxis_file is None:
            st.info(
                "Please upload a SynXis Rate Plan file in the sidebar to view detailed rate plan breakdowns."
            )
        elif df_synxis_exploded.empty:
            st.warning(
                "No records found matching the selected status filter."
            )
        else:
            rp_col1, rp_col2, rp_col3 = st.columns(3)
            total_rp_rooms = df_synxis_exploded["Room_Qty"].sum()
            total_rp_rev = df_synxis_exploded["Daily_Revenue"].sum()
            avg_rp_adr = (
                total_rp_rev / total_rp_rooms if total_rp_rooms > 0 else 0.0
            )

            rp_col1.metric("Filtered Total Rooms", f"{total_rp_rooms:,}")
            rp_col2.metric("Filtered Total Revenue", f"${total_rp_rev:,.2f}")
            rp_col3.metric("Filtered Overall ADR", f"${avg_rp_adr:,.2f}")

            st.markdown("---")

            st.write("### 📊 Production by Rate Code")
            rate_code_summary = (
                df_synxis_exploded.groupby(["Rate_Code", "Rate_Type"])
                .agg(
                    Rooms=("Room_Qty", "sum"),
                    Revenue=("Daily_Revenue", "sum"),
                )
                .reset_index()
            )
            rate_code_summary["ADR"] = (
                rate_code_summary["Revenue"]
                / rate_code_summary["Rooms"].replace(0, 1)
            ).round(2)
            rate_code_summary = rate_code_summary.sort_values(
                by="Revenue", ascending=False
            )

            st.dataframe(
                rate_code_summary.style.format(
                    {"Revenue": "${:,.2f}", "ADR": "${:,.2f}"}
                ),
                use_container_width=True,
            )

            st.write("### 📅 Daily Production Breakdown")
            st.dataframe(
                df_synxis_daily.style.format(
                    {
                        "Rate_Plan_Rev": "${:,.2f}",
                        "Rate_Plan_ADR": "${:,.2f}",
                    }
                ),
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
