import pandas as pd
import streamlit as st


def parse_ideas_pcdc(pcdc_file):
    if pcdc_file is None:
        return None, {}

    # 1. READ REPORT CRITERIA TAB (METADATA)
    metadata = {}
    try:
        # Read Report Criteria sheet
        df_crit = pd.read_excel(pcdc_file, sheet_name="Report Criteria", header=1)
        if not df_crit.empty:
            metadata["property_name"] = df_crit.get("Property Name", [None])[0]
            metadata["analysis_start"] = pd.to_datetime(
                df_crit.get("Analysis Start Date", [None])[0], errors="coerce"
            )
            metadata["analysis_end"] = pd.to_datetime(
                df_crit.get("Analysis End Date", [None])[0], errors="coerce"
            )
            metadata["activity_start"] = pd.to_datetime(
                df_crit.get("Activity Start Date", [None])[0], errors="coerce"
            )

            # Extract generated timestamp for current run date
            gen_on_str = str(df_crit.get("Generated On", [""])[0]).split(" ")[0]
            metadata["generated_on"] = pd.to_datetime(
                gen_on_str, errors="coerce"
            )
    except Exception as e:
        st.warning(f"Note: Could not parse 'Report Criteria' tab: {e}")

    # 2. READ MAIN PCDC DATA SHEET
    # Load first sheet with multi-level headers (rows 0, 1, 2 in Excel)
    df_raw = pd.read_excel(pcdc_file, sheet_name=0, header=[0, 1, 2])

    # Flatten multi-level column tuples into clean strings
    flat_cols = []
    for col in df_raw.columns:
        clean_parts = [
            str(c).strip() for c in col if "Unnamed" not in str(c) and str(c)
        ]
        flat_cols.append(" - ".join(clean_parts))
    df_raw.columns = flat_cols

    # Find and format date column
    date_col = next((c for c in df_raw.columns if "Occupancy Date" in c), None)
    if date_col:
        df_raw["Occupancy Date"] = pd.to_datetime(
            df_raw[date_col], errors="coerce"
        )
        df_raw = df_raw.dropna(subset=["Occupancy Date"])

    return df_raw, metadata
