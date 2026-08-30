import io
import pandas as pd
import streamlit as st

# =============================================================================
# STREAMLIT PAGE CONFIGURATION
# =============================================================================
st.set_page_config(
    page_title="2026 Master Day-by-Day Revenue Grid",
    page_icon="📊",
    layout="wide",
)


# =============================================================================
# 1. LIGHTHOUSE RATE SHOP PARSER
# =============================================================================
def parse_lighthouse_rate_shop(lh_file):
    if lh_file is None:
        return {}

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
        "21c Museum Hotel": [
            "21c Museum Hotel",
            "21c Museum Hotel Bentonville - MGallery",
        ],
        "Motto By Hilton": ["Motto By Hilton", "Motto By Hilton Bentonville"],
        "AC Hotel by Marriott": [
            "AC Hotel",
            "AC Hotel by Marriott Bentonville",
        ],
        "DoubleTree Suites": [
            "DoubleTree",
            "DoubleTree Suites by Hilton Bentonville",
        ],
    }

    for view_label, sheet_name in target_tabs.items():
        if sheet_name is None:
            continue

        df_temp = pd.read_excel(lh_file, sheet_name=sheet_name)
        header_row_idx = 0
        for i, row in df_temp.head(10).iterrows():
            row_str = " ".join([str(val) for val in row.values])
            if "Day Date" in row_str or "Date" in row_str:
                header_row_idx = i
                break

        df_clean = pd.read_excel(lh_file, sheet_name=sheet_name, header=header_row_idx)
        df_clean.columns = df_clean.columns.astype(str).str.strip()

        date_col = next((c for c in df_clean.columns if "Date" in c), None)
        if not date_col:
            continue

        df_clean["Occupancy Date"] = pd.to_datetime(df_clean[date_col], errors="coerce")
        df_clean = df_clean.dropna(subset=["Occupancy Date"])

        parsed_df = pd.DataFrame({"Occupancy Date": df_clean["Occupancy Date"]})

        for alias, search_terms in comp_map.items():
            found_col = None
            for term in search_terms:
                found_col = next(
                    (c for c in df_clean.columns if term.lower() in c.lower()), None
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
        parsed_df["Comp Set Avg"] = parsed_df[comp_cols].mean(axis=1)
        tabs_data[view_label] = parsed_df

    return tabs_data


# =============================================================================
# 2. SYNXIS RATE PLAN PARSER
# =============================================================================
def parse_synxis_rate_plan(synxis_file, selected_statuses=None):
    if synxis_file is None:
        return None, []

    if synxis_file.name.endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(synxis_file)
    else:
        df_raw = pd.read_csv(synxis_file)

    df_raw.columns = df_raw.columns.str.strip()

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

        for single_date in stay_dates:
            exploded_rows.append(
                {
                    "Occupancy Date": single_date,
                    "Room_Qty": room_qty,
                    "Daily_Revenue": avg_rate * room_qty,
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

    return df_daily, available_statuses


# =============================================================================
# 3. IDEAS PCDC PARSER
# =============================================================================
def parse_ideas_pcdc(pcdc_file):
    if pcdc_file is None:
        return None

    if pcdc_file.name.endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(pcdc_file)
    else:
        df_raw = pd.read_csv(pcdc_file)

    df_raw.columns = df_raw.columns.str.strip()

    date_col = next((c for c in df_raw.columns if "Date" in c or "Occupancy" in c), None)
    if not date_col:
        return None

    df_raw["Occupancy Date"] = pd.to_datetime(df_raw[date_col], errors="coerce")
    df_raw = df_raw.dropna(subset=["Occupancy Date"])

    rooms_col = next((c for c in df_raw.columns if "Room" in c or "Occ" in c or "Stays" in c), None)
    rev_col = next((c for c in df_raw.columns if "Rev" in c or "Revenue" in c), None)
    adr_col = next((c for c in df_raw.columns if "ADR" in c or "Rate" in c), None)

    df_pcdc = pd.DataFrame({"Occupancy Date": df_raw["Occupancy Date"]})
    df_pcdc["PCDC_Rooms"] = pd.to_numeric(df_raw[rooms_col], errors="coerce").fillna(0) if rooms_col else 0
    df_pcdc["PCDC_Rev"] = pd.to_numeric(df_raw[rev_col], errors="coerce").fillna(0.0) if rev_col else 0.0
    
    if adr_col:
        df_pcdc["PCDC_ADR"] = pd.to_numeric(df_raw[adr_col], errors="coerce").fillna(0.0)
    else:
        df_pcdc["PCDC_ADR"] = (df_pcdc["PCDC_Rev"] / df_pcdc["PCDC_Rooms"].replace(0, 1)).round(2)

    return df_pcdc


# =============================================================================
# 4. IDEAS DATA EXTRACTION PARSER
# =============================================================================
def parse_ideas_data_extraction(extract_file):
    if extract_file is None:
        return None

    if extract_file.name.endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(extract_file)
    else:
        df_raw = pd.read_csv(extract_file)

    df_raw.columns = df_raw.columns.str.strip()

    date_col = next((c for c in df_raw.columns if "Date" in c or "Occupancy" in c), None)
    if not date_col:
        return None

    df_raw["Occupancy Date"] = pd.to_datetime(df_raw[date_col], errors="coerce")
    df_raw = df_raw.dropna(subset=["Occupancy Date"])

    rooms_col = next((c for c in df_raw.columns if "Room" in c or "Occ" in c or "Sold" in c), None)
    rev_col = next((c for c in df_raw.columns if "Rev" in c or "Revenue" in c), None)

    df_ext = pd.DataFrame({"Occupancy Date": df_raw["Occupancy Date"]})
    df_ext["Extract_Rooms"] = pd.to_numeric(df_raw[rooms_col], errors="coerce").fillna(0) if rooms_col else 0
    df_ext["Extract_Rev"] = pd.to_numeric(df_raw[rev_col], errors="coerce").fillna(0.0) if rev_col else 0.0
    df_ext["Extract_ADR"] = (df_ext["Extract_Rev"] / df_ext["Extract_Rooms"].replace(0, 1)).round(2)

    return df_ext


# =============================================================================
# 5. BASE GRID BUILDER (CALENDAR YEAR 2026)
# =============================================================================
def build_2026_base_grid():
    dates = pd.date_range(start="2026-01-01", end="2026-12-31", freq="D")
    base_df = pd.DataFrame({"Occupancy Date": dates})
    base_df["Day of Week"] = base_df["Occupancy Date"].dt.strftime("%a")
    base_df["Month"] = base_df["Occupancy Date"].dt.strftime("%b %Y")
    return base_df


# =============================================================================
# 6. STREAMLIT MAIN APPLICATION
# =============================================================================
def main():
    st.title("🏨 2026 Master Day-by-Day Revenue Grid")
    st.caption(
        "Consolidates Lighthouse Rate Shops, SynXis Rate Plans, IDeaS PCDC, and IDeaS Data Extractions."
    )

    # --- SIDEBAR: FILE UPLOAD CENTER ---
    st.sidebar.header("📁 File Upload Center")
    lh_file = st.sidebar.file_uploader("1. Lighthouse Rate Shop (.xlsx)", type=["xlsx"])
    synxis_file = st.sidebar.file_uploader("2. SynXis Rate Plan Export (.csv / .xlsx)", type=["csv", "xlsx"])
    pcdc_file = st.sidebar.file_uploader("3. IDeaS PCDC Report (.csv / .xlsx)", type=["csv", "xlsx"])
    extract_file = st.sidebar.file_uploader("4. IDeaS Data Extraction (.csv / .xlsx)", type=["csv", "xlsx"])

    # --- SIDEBAR: CONTROLS ---
    st.sidebar.markdown("---")
    st.sidebar.header("🕹️ Rate Shop Controls")
    rate_view_option = st.sidebar.radio(
        "Competitor Rate View Mode:",
        options=["Rates", "vs. Yesterday", "vs. 3 Days Ago", "vs. 7 Days Ago"],
        index=0,
    )

    st.sidebar.markdown("---")
    st.sidebar.header("🛎️ SynXis Status Filter")
    selected_status_list = []
    if synxis_file is not None:
        _, available_statuses = parse_synxis_rate_plan(synxis_file, selected_statuses=None)
        default_selected = [s for s in available_statuses if "confirm" in s.lower()]
        if not default_selected and available_statuses:
            default_selected = available_statuses

        selected_status_list = st.sidebar.multiselect(
            "Filter SynXis `Rez_Status`:",
            options=available_statuses,
            default=default_selected,
        )

    # --- BUILD & MERGE GRID ---
    base_df = build_2026_base_grid()

    # 1. Lighthouse Integration
    lh_tabs = parse_lighthouse_rate_shop(lh_file) if lh_file else {}
    if rate_view_option in lh_tabs:
        df_lh_selected = lh_tabs[rate_view_option]
        base_df = pd.merge(base_df, df_lh_selected, on="Occupancy Date", how="left").fillna(0)
    else:
        for comp in [
            "21c Museum Hotel",
            "Motto By Hilton",
            "AC Hotel by Marriott",
            "DoubleTree Suites",
            "Comp Set Avg",
        ]:
            base_df[comp] = 0.0

    # 2. SynXis Integration
    if synxis_file is not None:
        df_synxis_daily, _ = parse_synxis_rate_plan(synxis_file, selected_statuses=selected_status_list)
        base_df = pd.merge(base_df, df_synxis_daily, on="Occupancy Date", how="left").fillna(0)
    else:
        base_df["Rate_Plan_Rooms"] = 0
        base_df["Rate_Plan_Rev"] = 0.0
        base_df["Rate_Plan_ADR"] = 0.0

    # 3. IDeaS PCDC Integration
    if pcdc_file is not None:
        df_pcdc = parse_ideas_pcdc(pcdc_file)
        if df_pcdc is not None:
            base_df = pd.merge(base_df, df_pcdc, on="Occupancy Date", how="left").fillna(0)
        else:
            base_df["PCDC_Rooms"], base_df["PCDC_Rev"], base_df["PCDC_ADR"] = 0, 0.0, 0.0
    else:
        base_df["PCDC_Rooms"], base_df["PCDC_Rev"], base_df["PCDC_ADR"] = 0, 0.0, 0.0

    # 4. IDeaS Data Extraction Integration
    if extract_file is not None:
        df_ext = parse_ideas_data_extraction(extract_file)
        if df_ext is not None:
            base_df = pd.merge(base_df, df_ext, on="Occupancy Date", how="left").fillna(0)
        else:
            base_df["Extract_Rooms"], base_df["Extract_Rev"], base_df["Extract_ADR"] = 0, 0.0, 0.0
    else:
        base_df["Extract_Rooms"], base_df["Extract_Rev"], base_df["Extract_ADR"] = 0, 0.0, 0.0

    # --- CONSTRUCT MULTIINDEX DATAFRAME ---
    multi_columns = pd.MultiIndex.from_tuples(
        [
            ("Date Info", "Occupancy Date"),
            ("Date Info", "Day of Week"),
            ("Date Info", "Month"),
            ("SynXis Rate Plan OTB", "Rooms"),
            ("SynXis Rate Plan OTB", "Revenue"),
            ("SynXis Rate Plan OTB", "ADR"),
            ("IDeaS PCDC", "Rooms"),
            ("IDeaS PCDC", "Revenue"),
            ("IDeaS PCDC", "ADR"),
            ("IDeaS Data Extract", "Rooms"),
            ("IDeaS Data Extract", "Revenue"),
            ("IDeaS Data Extract", "ADR"),
            ("Pricing & Capacity", "Comp Set Avg"),
            ("Competitor Shops", "21c Museum Hotel"),
            ("Competitor Shops", "Motto By Hilton"),
            ("Competitor Shops", "AC Hotel by Marriott"),
            ("Competitor Shops", "DoubleTree Suites"),
        ]
    )

    grid_df = pd.DataFrame(index=base_df.index, columns=multi_columns)

    grid_df[("Date Info", "Occupancy Date")] = base_df["Occupancy Date"].dt.strftime("%Y-%m-%d")
    grid_df[("Date Info", "Day of Week")] = base_df["Day of Week"]
    grid_df[("Date Info", "Month")] = base_df["Month"]

    grid_df[("SynXis Rate Plan OTB", "Rooms")] = base_df["Rate_Plan_Rooms"].astype(int)
    grid_df[("SynXis Rate Plan OTB", "Revenue")] = base_df["Rate_Plan_Rev"].round(2).map("{:,.2f}".format)
    grid_df[("SynXis Rate Plan OTB", "ADR")] = base_df["Rate_Plan_ADR"].round(2).map("{:,.2f}".format)

    grid_df[("IDeaS PCDC", "Rooms")] = base_df["PCDC_Rooms"].astype(int)
    grid_df[("IDeaS PCDC", "Revenue")] = base_df["PCDC_Rev"].round(2).map("{:,.2f}".format)
    grid_df[("IDeaS PCDC", "ADR")] = base_df["PCDC_ADR"].round(2).map("{:,.2f}".format)

    grid_df[("IDeaS Data Extract", "Rooms")] = base_df["Extract_Rooms"].astype(int)
    grid_df[("IDeaS Data Extract", "Revenue")] = base_df["Extract_Rev"].round(2).map("{:,.2f}".format)
    grid_df[("IDeaS Data Extract", "ADR")] = base_df["Extract_ADR"].round(2).map("{:,.2f}".format)

    grid_df[("Pricing & Capacity", "Comp Set Avg")] = base_df["Comp Set Avg"].round(2).map("{:,.2f}".format)
    grid_df[("Competitor Shops", "21c Museum Hotel")] = base_df["21c Museum Hotel"].round(2).map("{:,.2f}".format)
    grid_df[("Competitor Shops", "Motto By Hilton")] = base_df["Motto By Hilton"].round(2).map("{:,.2f}".format)
    grid_df[("Competitor Shops", "AC Hotel by Marriott")] = base_df["AC Hotel by Marriott"].round(2).map("{:,.2f}".format)
    grid_df[("Competitor Shops", "DoubleTree Suites")] = base_df["DoubleTree Suites"].round(2).map("{:,.2f}".format)

    # --- SUMMARY METRICS BAR ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("SynXis Rooms OTB", f"{base_df['Rate_Plan_Rooms'].sum():,}")
    col2.metric("SynXis Revenue OTB", f"${base_df['Rate_Plan_Rev'].sum():,.2f}")
    col3.metric("IDeaS PCDC Revenue", f"${base_df['PCDC_Rev'].sum():,.2f}")
    col4.metric("IDeaS Extract Revenue", f"${base_df['Extract_Rev'].sum():,.2f}")

    st.markdown("---")

    # --- DISPLAY MAIN GRID ---
    st.subheader(f"📅 2026 Master Grid Breakdown (Active View: {rate_view_option})")
    st.dataframe(grid_df, use_container_width=True, height=650)

    # --- EXPORT EXCEL CAPABILITY ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        grid_df.to_excel(writer, sheet_name="2026_Master_DD_Grid")
    buffer.seek(0)

    st.download_button(
        label="📥 Download Master DD Grid (.xlsx)",
        data=buffer,
        file_name="2026_Master_DD_Grid_Export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
