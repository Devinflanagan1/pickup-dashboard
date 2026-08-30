"""
The Compton — Daily Pick-Up Report Builder
===========================================
Upload your four exports and this app rebuilds your pick-up report automatically,
so you never have to copy/paste again.

Files it expects (any order — it auto-detects each one):
  1. PCDC.xlsx .................. Change & Differential Control Report (Business Type)
  2. Data Extract.xlsx ......... Data Extraction Report (Property)
  3. Market Seg.xlsx ........... Data Extraction Report (Market Segment)
  4. Rate Shop .xlsx ........... Brand.com BAR / competitor shop (21c, Motto, AC, DoubleTree)

Run it with:   streamlit run app.py
"""

import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Compton Daily Pick-Up Report", layout="wide")

# ----------------------------------------------------------------------------- #
#  Helpers
# ----------------------------------------------------------------------------- #
def _num(x):
    """Convert to float; anything non-numeric ('Sold out', 'LOS3', blanks) -> NaN."""
    try:
        return float(x)
    except (TypeError, ValueError):
        return np.nan


def _read_all_sheets(uploaded_file):
    """Return {sheet_name: raw DataFrame (header=None)} for an uploaded workbook."""
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    return {s: xls.parse(s, header=None) for s in xls.sheet_names}


def detect_and_load(uploaded_files):
    """
    Look at every uploaded workbook and figure out which one is which,
    based on its sheet names / header text. Returns a dict of parsed frames.
    """
    found = {"pcdc": None, "data": None, "mseg": None, "shop": None,
             "shop_chg_1": None, "shop_chg_3": None, "shop_chg_7": None}

    for f in uploaded_files:
        try:
            sheets = _read_all_sheets(f)
        except Exception as e:
            st.warning(f"Could not open **{f.name}** ({e}).")
            continue
        names = [s.lower() for s in sheets.keys()]

        # ---- Rate shop: has a 'Rates' sheet ------------------------------- #
        if any("rate" in n for n in names) and any("overview" in n for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "rates"][0]]
            found["shop"] = _parse_rateshop(raw)
            # also grab the vs-comparison tabs, if present
            for tab, key in [("vs. yesterday", "shop_chg_1"),
                             ("vs. 3 days ago", "shop_chg_3"),
                             ("vs. 7 days ago", "shop_chg_7")]:
                match = [s for s in sheets if s.lower() == tab]
                if match:
                    found[key] = _parse_shop_change(sheets[match[0]])

        # ---- PCDC: sheet name contains 'ChangeReport' --------------------- #
        elif any("changereport" in n for n in names):
            raw = sheets[[s for s in sheets if "changereport" in s.lower()][0]]
            found["pcdc"] = _parse_pcdc(raw)

        # ---- Market Segment: has a 'Market Segment' sheet ----------------- #
        elif any("market segment" in n for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "market segment"][0]]
            found["mseg"] = _parse_mseg(raw)

        # ---- Data Extract: has a 'Property' sheet ------------------------- #
        elif any(n == "property" for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "property"][0]]
            found["data"] = _parse_data(raw)

        else:
            st.warning(f"Didn't recognise **{f.name}** — skipping. "
                       f"(sheets: {', '.join(sheets.keys())})")
    return found


# ----------------------------------------------------------------------------- #
#  Individual file parsers (all key on Occupancy Date)
# ----------------------------------------------------------------------------- #
def _parse_data(raw):
    """Data Extract › Property sheet. Header in row 0."""
    df = raw.copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    df["Occupancy Date"] = pd.to_datetime(df["Occupancy Date"], errors="coerce")
    df = df.dropna(subset=["Occupancy Date"])
    return df


def _parse_pcdc(raw):
    """
    PCDC › ChangeReport. Multi-row header (rows 0-2), data from row 3.
    Columns are taken by position (they're stable in this SynXis export).
    """
    df = raw.iloc[3:].copy()
    df[0] = pd.to_datetime(df[0], errors="coerce")
    df = df.dropna(subset=[0]).reset_index(drop=True)
    out = pd.DataFrame({"Occupancy Date": df[0]})
    out["pc_event"]      = df[2]
    out["trans_cur"]     = pd.to_numeric(df[3], errors="coerce")
    out["trans_chg"]     = pd.to_numeric(df[4], errors="coerce")
    out["grp_cur"]       = pd.to_numeric(df[5], errors="coerce")
    out["grp_chg"]       = pd.to_numeric(df[6], errors="coerce")
    out["grp_blocked"]   = pd.to_numeric(df[7], errors="coerce")
    out["grp_avail"]     = pd.to_numeric(df[8], errors="coerce")   # Remaining in block
    out["grp_pickup"]    = pd.to_numeric(df[9], errors="coerce")   # Picked up from block
    out["fcst_trans"]    = pd.to_numeric(df[10], errors="coerce")
    out["fcst_grp"]      = pd.to_numeric(df[12], errors="coerce")
    out["rev_trans"]     = pd.to_numeric(df[14], errors="coerce")
    out["rev_grp"]       = pd.to_numeric(df[16], errors="coerce")
    out["adr_trans"]     = pd.to_numeric(df[22], errors="coerce")
    out["adr_grp"]       = pd.to_numeric(df[24], errors="coerce")
    return out


def _parse_mseg(raw):
    """Market Segment sheet. Header row 0. Aggregate Transient vs Group by date."""
    df = raw.copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    df["Occupancy Date"] = pd.to_datetime(df["Occupancy Date"], errors="coerce")
    df = df.dropna(subset=["Occupancy Date"])
    df["Occupancy On Books This Year"] = pd.to_numeric(
        df["Occupancy On Books This Year"], errors="coerce").fillna(0)
    df["is_trans"] = df["Market Segment"].astype(str).str.startswith("Transient")
    agg = df.groupby("Occupancy Date").apply(
        lambda g: pd.Series({
            "ms_trans": g.loc[g["is_trans"], "Occupancy On Books This Year"].sum(),
            "ms_group": g.loc[~g["is_trans"], "Occupancy On Books This Year"].sum(),
        })
    ).reset_index()
    return agg


def _clean_shop(x):
    """Keep the shop value as-is for display ('Sold out', 'LOS3', 359),
    but tidy numbers so 359.0 shows as 359."""
    if x is None:
        return None
    v = _num(x)
    if not np.isnan(v):                       # it's a number
        return int(v) if float(v).is_integer() else v
    s = str(x).strip()
    return s if s else None                    # keep text like 'Sold out' / 'LOS3'


def _parse_rateshop(raw):
    """Rate shop › Rates sheet. Header row 4, data from row 5. Columns by position.
    Each competitor keeps a display value (text or number) AND a numeric value
    (used only for the Comp Set Avg)."""
    df = raw.iloc[5:].copy()
    df[2] = pd.to_datetime(df[2], errors="coerce")
    df = df.dropna(subset=[2]).reset_index(drop=True)
    out = pd.DataFrame({"Occupancy Date": df[2]})
    out["own_bar"] = df[4].map(_num)
    for key, col in [("c_21c", 5), ("c_motto", 6), ("c_ac", 7), ("c_dt", 8)]:
        out[key] = df[col].map(_clean_shop)          # display value (may be text)
        out[key + "_n"] = df[col].map(_num)          # numeric only (NaN for text)
    return out


def _parse_shop_change(raw):
    """
    One of the 'vs. Yesterday / 3 days ago / 7 days ago' tabs.
    Header row 4, data from row 5. Each competitor's CHANGE sits in the
    column immediately to the right of its rate:
        21c rate=7 chg=8 | Motto rate=9 chg=10 | AC rate=11 chg=12 | DT r
