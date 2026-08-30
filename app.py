"""
The Compton — Daily Pick-Up Report Builder
Run it with:   streamlit run app.py
"""

import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Compton Daily Pick-Up Report", layout="wide")


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return np.nan


def _read_all_sheets(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    return {s: xls.parse(s, header=None) for s in xls.sheet_names}


def detect_and_load(uploaded_files):
    found = {"pcdc": None, "data": None, "mseg": None, "shop": None,
             "shop_chg_1": None, "shop_chg_3": None, "shop_chg_7": None,
             "pcdc_asof": None}

    for f in uploaded_files:
        try:
            sheets = _read_all_sheets(f)
        except Exception as e:
            st.warning(f"Could not open **{f.name}** ({e}).")
            continue
        names = [s.lower() for s in sheets.keys()]

        if any("rate" in n for n in names) and any("overview" in n for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "rates"][0]]
            found["shop"] = _parse_rateshop(raw)
            for tab, key in [("vs. yesterday", "shop_chg_1"),
                             ("vs. 3 days ago", "shop_chg_3"),
                             ("vs. 7 days ago", "shop_chg_7")]:
                match = [s for s in sheets if s.lower() == tab]
                if match:
                    found[key] = _parse_shop_change(sheets[match[0]])

        elif any("changereport" in n for n in names):
            raw = sheets[[s for s in sheets if "changereport" in s.lower()][0]]
            found["pcdc"] = _parse_pcdc(raw)
            crit = [s for s in sheets if s.lower() == "report criteria"]
            if crit:
                found["pcdc_asof"] = _parse_pcdc_asof(sheets[crit[0]])

        elif any("market segment" in n for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "market segment"][0]]
            found["mseg"] = _parse_mseg(raw)

        elif any(n == "property" for n in names):
            raw = sheets[[s for s in sheets if s.lower() == "property"][0]]
            found["data"] = _parse_data(raw)

        else:
            st.warning(f"Didn't recognise **{f.name}** — skipping. "
                       f"(sheets: {', '.join(sheets.keys())})")
    return found


def _parse_data(raw):
    df = raw.copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    df["Occupancy Date"] = pd.to_datetime(df["Occupancy Date"], errors="coerce")
    df = df.dropna(subset=["Occupancy Date"])
    return df


def _parse_pcdc(raw):
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
    out["grp_avail"]     = pd.to_numeric(df[8], errors="coerce")
    out["grp_pickup"]    = pd.to_numeric(df[9], errors="coerce")
    out["fcst_trans"]    = pd.to_numeric(df[10], errors="coerce")
    out["fcst_grp"]      = pd.to_numeric(df[12], errors="coerce")
    out["rev_trans"]     = pd.to_numeric(df[14], errors="coerce")
    out["rev_grp"]       = pd.to_numeric(df[16], errors="coerce")
    out["adr_trans"]     = pd.to_numeric(df[22], errors="coerce")
    out["adr_grp"]       = pd.to_numeric(df[24], errors="coerce")
    return out


def _parse_pcdc_asof(raw):
    """
    Read the PCDC 'Report Criteria' sheet and return the reference date for
    'Days Left'.  Prefers 'Analysis Start Date'; falls back to 'Generated On'
    (the export/created date).  Returns a datetime.date or None.
    """
    try:
        # find the header row that contains the labels, then the value row below it
        for i in range(len(raw) - 1):
            rowvals = [str(x).strip().lower() for x in raw.iloc[i].tolist()]
            if "analysis start date" in rowvals or "generated on" in rowvals:
                header = raw.iloc[i].tolist()
                values = raw.iloc[i + 1].tolist()
                lut = {str(h).strip().lower(): v for h, v in zip(header, values)}
                for key in ("analysis start date", "activity start date", "generated on"):
                    if key in lut and lut[key] not in (None, ""):
                        d = pd.to_datetime(lut[key], errors="coerce")
                        if pd.notna(d):
                            return d.date()
    except Exception:
        pass
    return None


def _parse_mseg(raw):
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
    if x is None:
        return None
    v = _num(x)
    if not np.isnan(v):
        return int(v) if float(v).is_integer() else v
    s = str(x).strip()
    return s if s else None


def _parse_rateshop(raw):
    df = raw.iloc[5:].copy()
    df[2] = pd.to_datetime(df[2], errors="coerce")
    df = df.dropna(subset=[2]).reset_index(drop=True)
    out = pd.DataFrame({"Occupancy Date": df[2]})
    out["own_bar"] = df[4].map(_num)
    for key, col in [("c_21c", 5), ("c_motto", 6), ("c_ac", 7), ("c_dt", 8)]:
        out[key] = df[col].map(_clean_shop)
        out[key + "_n"] = df[col].map(_num)
    return out


def _parse_shop_change(raw):
    df = raw.iloc[5:].copy()
    df[2] = pd.to_datetime(df[2], errors="coerce")
    df = df.dropna(subset=[2]).reset_index(drop=True)
    out = pd.DataFrame({"Occupancy Date": df[2]})
    # our own BAR change sits in column 6 (right of our rate in column 5)
    out["own_chg"] = df[6].map(_num) if 6 in df.columns else np.nan
    for key, chg_col in [("c_21c", 8), ("c_motto", 10), ("c_ac", 12), ("c_dt", 14)]:
        out[key + "_chg"] = df[chg_col].map(_num) if chg_col in df.columns else np.nan
    return out


def build_report(parts, as_of, comp_method="Average", show_change="None", thr=None,
                 show_own_change=True):
    data, pcdc, mseg, shop = parts["data"], parts["pcdc"], parts["mseg"], parts["shop"]
    if data is None:
        st.error("The **Data Extract** file is required (couldn't find a 'Property' sheet).")
        st.stop()

    df = data.copy()
    for extra in (pcdc, mseg, shop):
        if extra is not None:
            df = df.merge(extra, on="Occupancy Date", how="left")

    chg_key = {"Yesterday": "shop_chg_1", "3 days ago": "shop_chg_3",
               "7 days ago": "shop_chg_7"}.get(show_change)
    chg_df = parts.get(chg_key) if chg_key else None
    if chg_df is not None:
        # avoid duplicate own_chg if we also pull the 7-day own change below
        df = df.merge(chg_df.drop(columns=[c for c in ["own_chg"] if c in chg_df], errors="ignore"),
                      on="Occupancy Date", how="left")

    # Always pull OUR OWN 7-day BAR change (independent of the competitor toggle)
    own7 = parts.get("shop_chg_7")
    if own7 is not None and "own_chg" in own7:
        df = df.merge(own7[["Occupancy Date", "own_chg"]].rename(columns={"own_chg": "bar_chg_7d"}),
                      on="Occupancy Date", how="left")

    def dnum(col):
        return pd.to_numeric(df.get(col), errors="coerce")

    cap        = dnum("Physical Capacity This Year")
    occ_ty     = dnum("Occupancy On Books This Year")
    occ_stly   = dnum("Occupancy On Books STLY")
    tr_ty      = dnum("Rooms Sold - Transient This Year")
    tr_stly    = dnum("Rooms Sold - Transient STLY")
    gr_ty      = dnum("Rooms Sold - Group This Year")
    gr_stly    = dnum("Rooms Sold - Group STLY")
    dem_tot    = dnum("User Total Demand - Total This Year")
    dem_grp    = dnum("User Constrained Total Demand - Group This Year")
    dem_tr     = dnum("User Unconstrained Total Demand - Transient This Year")

    r = pd.DataFrame()
    r["DOW"]  = df["Day of Week"]
    r["Date"] = df["Occupancy Date"].dt.date
    r["Days Left"] = (df["Occupancy Date"] - pd.Timestamp(as_of)).dt.days

    ev = df.get("Special Event This Year")
    ev = ev.fillna("") if ev is not None else ""
    if "pc_event" in df:
        ev = ev.where(ev.astype(str).str.strip() != "", df["pc_event"].fillna(""))
    r["Events"] = pd.Series(ev, index=df.index).replace("nan", "")

    r["Rooms Left to Sell"] = dnum("Remaining Capacity This Year")
    r["BAR"]                = dnum("BAR")
    if show_own_change:
        r["BAR Chg 7d"]     = dnum("bar_chg_7d")     # our own rate move over last 7 days

    if "c_21c_n" in df:
        comps_n = df[["c_21c_n", "c_motto_n", "c_ac_n", "c_dt_n"]]
        r["Comp Set Avg"] = (comps_n.mean(axis=1) if comp_method == "Average"
                             else comps_n.median(axis=1)).round(0)
    else:
        r["Comp Set Avg"] = np.nan

    r["Last Room Value"] = dnum("Last Room Value This Year")

    add_chg = chg_df is not None
    for name, key in [("21c Museum Hotel Bentonville - MGallery", "c_21c"),
                      ("Motto By Hilton Bentonville Downtown",    "c_motto"),
                      ("AC Hotel by Marriott Bentonville",        "c_ac"),
                      ("DoubleTree Suites by Hilton Bentonville", "c_dt")]:
        base = "Competitor Shops | " + name
        if add_chg:
            r[base + " | Current"] = df.get(key)
            r[base + " | Change"]  = df.get(key + "_chg")
        else:
            r[base] = df.get(key)

    r["OOO"]   = dnum("Rooms N/A - Out of Order This Year")
    r["Ovrbk"] = dnum("Overbooking This Year")

    tc, tch = df.get("trans_cur"), df.get("trans_chg")
    gc, gch = df.get("grp_cur"),   df.get("grp_chg")
    r["Rooms Sold | Total Hotel | Current"]   = (tc.fillna(0) + gc.fillna(0)) if tc is not None else occ_ty
    r["Rooms Sold | Total Hotel | Change"]    = (tch.fillna(0) + gch.fillna(0)) if tch is not None else np.nan
    r["Rooms Sold | Total Transient | Current"] = tc
    r["Rooms Sold | Total Transient | Change"]  = tch
    r["Rooms Sold | Total Group | Current"]     = gc
    r["Rooms Sold | Total Group | Change"]      = gch
    r["Rooms Sold | Total Group | Blocked"]     = df.get("grp_blocked")
    r["Rooms Sold | Total Group | P/U"]         = df.get("grp_pickup")
    r["Rooms Sold | Total Group | Remaining"]   = df.get("grp_avail")

    r["Rooms OTB STLY | Total Hotel | Current"]     = occ_stly
    r["Rooms OTB STLY | Total Hotel | Change"]      = occ_ty - occ_stly
    r["Rooms OTB STLY | Total Transient | Current"] = tr_stly
    r["Rooms OTB STLY | Total Transient | Change"]  = tr_ty - tr_stly
    r["Rooms OTB STLY | Total Group | Current"]     = gr_stly
    r["Rooms OTB STLY | Total Group | Change"]      = gr_ty - gr_stly

    r["Remaining Demand | Total Hotel"]     = (dem_tot - occ_ty).clip(lower=0)
    r["Remaining Demand | Total Transient"] = (dem_tr - tr_ty).clip(lower=0)
    r["Remaining Demand | Total Group"]     = (dem_grp - gr_ty).clip(lower=0)

    ft, fg = df.get("fcst_trans"), df.get("fcst_grp")
    fcst_tot = (ft.fillna(0) + fg.fillna(0)) if ft is not None else dnum("Forecasted Room Revenue This Year")*np.nan
    r["Occ Forecast | Total Hotel"]     = fcst_tot.round(1) if ft is not None else np.nan
    r["Occ Forecast | Total Transient"] = ft.round(1) if ft is not None else np.nan
    r["Occ Forecast | Total Group"]     = fg.round(1) if fg is not None else np.nan

    r["Occ Forecast % | Total Hotel"]     = (fcst_tot / cap) if ft is not None else np.nan
    r["Occ Forecast % | Total Transient"] = (ft / cap) if ft is not None else np.nan
    r["Occ Forecast % | Total Group"]     = (fg / cap) if fg is not None else np.nan

    rev_t, rev_g = df.get("rev_trans"), df.get("rev_grp")
    if tc is not None:
        total_rooms = (tc.fillna(0) + gc.fillna(0)).replace(0, np.nan)
        r["Booked ADR | Total Hotel"] = ((rev_t.fillna(0) + rev_g.fillna(0)) / total_rooms).round(2)
    else:
        r["Booked ADR | Total Hotel"] = dnum("ADR On Books This Year")
    r["Booked ADR | Total Transient"] = df.get("adr_trans")
    r["Booked ADR | Total Group"]     = df.get("adr_grp")

    r["Estimated ADR"] = dnum("ADR Forecast This Year")

    # ================================================================= #
    #  PRICING DECISION SIGNALS  (VP-of-Revenue toolkit)
    # ================================================================= #
    bar   = r["BAR"]
    comp  = r["Comp Set Avg"]
    rl    = r["Rooms Left to Sell"]
    rd    = r["Remaining Demand | Total Hotel"]
    ofp   = r["Occ Forecast % | Total Hotel"]
    pace  = r["Rooms OTB STLY | Total Hotel | Change"]      # TY − STLY (+ = ahead)
    dleft = r["Days Left"]
    floor   = dnum("Floor")
    ceiling = dnum("Ceiling")

    # --- Rate Rank badge:  where our BAR sits in the comp set (1 = highest) --- #
    rank_cols = [c for c in ("c_21c_n", "c_motto_n", "c_ac_n", "c_dt_n") if c in df]
    if rank_cols:
        price_df = pd.DataFrame({"own": pd.to_numeric(r["BAR"], errors="coerce").values})
        for cc in rank_cols:
            price_df[cc] = pd.to_numeric(df[cc], errors="coerce").values

        def _rank_row(row):
            own = row["own"]
            allv = [v for v in row.tolist() if pd.notna(v)]
            if pd.isna(own) or not allv:
                return ""
            higher = sum(1 for v in allv if v > own)
            return f"{higher + 1} of {len(allv)}"

        r["Rate Rank"] = [_rank_row(price_df.iloc[i]) for i in range(len(price_df))]
    else:
        r["Rate Rank"] = ""

    # --- Signal thresholds (tunable from the sidebar) --- #
    thr = thr or {}
    lead_buffer = thr.get("lead_buffer", 0)    # $ we want to sit ABOVE the top competitor
    comp_occ    = thr.get("comp_occ", 0.85)    # OTB occupancy % that means "compression"
    move_thresh = thr.get("move_thresh", 20)   # $ comp-set move that counts as "major"

    # Strategy: The Compton is the price LEADER — we want to be #1 (highest) in the
    # comp set, and we're pushing market rates up.  Three flags:
    #   1) Rate leadership  — are we still on top? (raise to reclaim #1)
    #   2) OTB compression  — our own occupancy is filling up (raise)
    #   3) Comp-set movement — a competitor moved rates materially (review)

    cap_safe = cap.replace(0, np.nan)
    otb_occ_pct = occ_ty / cap_safe                       # OTB occupancy (on the books)

    # top competitor rate per date (max of the 4 shops)
    if rank_cols:
        comp_matrix = pd.DataFrame({cc: pd.to_numeric(df[cc], errors="coerce").values
                                    for cc in rank_cols})
        max_comp = comp_matrix.max(axis=1)
        max_comp.index = r.index
    else:
        max_comp = pd.Series(np.nan, index=r.index)

    # biggest absolute comp-set move (uses the change window if one is selected)
    chg_cols = [c for c in ("c_21c_chg", "c_motto_chg", "c_ac_chg", "c_dt_chg") if c in df]
    if chg_cols:
        move_matrix = pd.DataFrame({cc: pd.to_numeric(df[cc], errors="coerce").abs().values
                                    for cc in chg_cols})
        max_move = move_matrix.max(axis=1)
        max_move.index = r.index
    else:
        max_move = pd.Series(np.nan, index=r.index)

    # store the gap-to-#1 so the user can see how much room they have to lead
    lead_gap = (bar - max_comp)                            # + = we're above top comp
    r["Lead Gap vs Comp"] = lead_gap.round(0)

    not_leader  = (bar - max_comp) < lead_buffer          # we're not clearly #1
    compression = otb_occ_pct >= comp_occ
    comp_moved  = max_move >= move_thresh
    at_ceiling  = bar >= ceiling

    def _label(i):
        if pd.isna(dleft.iloc[i]) or dleft.iloc[i] < 0:
            return ""                                     # past date
        # 1) Compression takes top priority — real demand on our own books
        if bool(compression.iloc[i]):
            return "🔴 Raise · compression" if not bool(at_ceiling.iloc[i]) else "🟠 At ceiling"
        # 2) Rate leadership — reclaim #1 if a competitor has caught us
        if bool(not_leader.iloc[i]) and pd.notna(max_comp.iloc[i]):
            return "🔴 Raise · reclaim #1" if not bool(at_ceiling.iloc[i]) else "🟠 At ceiling"
        # 3) Comp-set movement — a competitor moved materially; review positioning
        if bool(comp_moved.iloc[i]):
            return "🔵 Comp moved · review"
        return "⚪ Hold"

    r["⚡ Action"] = [_label(i) for i in range(len(r))]

    # ---- Enforce the exact original report column order ---- #
    lead = ["⚡ Action", "DOW", "Date", "Days Left", "Events", "Rooms Left to Sell",
            "BAR", "BAR Chg 7d", "Comp Set Avg", "Rate Rank", "Lead Gap vs Comp",
            "Last Room Value"]
    comp = [c for c in r.columns if c.startswith("Competitor Shops")]
    mid  = ["OOO", "Ovrbk"]
    groups = ["Rooms Sold", "Rooms OTB STLY", "Remaining Demand",
              "Occ Forecast |", "Occ Forecast %", "Booked ADR"]
    grouped = []
    for g in groups:
        grouped += [c for c in r.columns if c.startswith(g)]
    tail = ["Estimated ADR"]
    order = lead + comp + mid + grouped + tail
    order = [c for c in order if c in r.columns]          # keep only existing
    order += [c for c in r.columns if c not in order]     # safety: append any strays
    r = r[order]

    return r


def _levels(col):
    parts = [p.strip() for p in str(col).split("|")]
    while len(parts) < 3:
        parts.append("")
    return tuple(parts[:3])


def to_excel(r, kpi_ctx=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule

    lv = [_levels(c) for c in r.columns]
    n = len(lv)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pick-Up Report"

    NAVY, TEAL, ORANGE, PURPLE, GREEN = "16365C", "215967", "E26B0A", "60497A", "76933C"
    white = Font(color="FFFFFF", bold=True, size=9)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

    def PF(hexcol):
        return PatternFill("solid", fgColor=hexcol)

    def seg_color(txt):
        if "Total Hotel" in txt:
            return ORANGE
        if "Total Transient" in txt:
            return PURPLE
        if "Total Group" in txt:
            return GREEN
        return None

    def band_fill(l1, l2):
        return PF(NAVY) if l2 == "" else PF(TEAL)

    def lvl2_fill(l1, l2):
        sc = seg_color(l2)
        if sc:
            return PF(sc)
        if l1 == "Competitor Shops":
            return PF(NAVY)
        return PF(TEAL)

    def lvl3_fill(l1, l2):
        sc = seg_color(l2)
        return PF(sc) if sc else PF(NAVY)

    j = 0
    while j < n:
        l1 = lv[j][0]
        k = j
        while k + 1 < n and lv[k + 1][0] == l1 and lv[j][1]:
            k += 1
        fill = band_fill(l1, lv[j][1])
        if lv[j][1] == "" and l1 in ("Estimated ADR",):
            fill = PF(ORANGE)
        cell = ws.cell(row=1, column=j + 1, value=l1)
        cell.font = white
        cell.alignment = center
        cell.fill = fill
        if lv[j][1] == "":
            ws.merge_cells(start_row=1, start_column=j + 1, end_row=3, end_column=j + 1)
        elif k > j:
            ws.merge_cells(start_row=1, start_column=j + 1, end_row=1, end_column=k + 1)
        j = k + 1

    j = 0
    while j < n:
        l1, l2, l3 = lv[j]
        if l2:
            k = j
            while k + 1 < n and lv[k + 1][0] == l1 and lv[k + 1][1] == l2:
                k += 1
            cell = ws.cell(row=2, column=j + 1, value=l2)
            cell.font = white
            cell.alignment = center
            cell.fill = lvl2_fill(l1, l2)
            if l3 == "":
                ws.merge_cells(start_row=2, start_column=j + 1, end_row=3, end_column=j + 1)
            elif k > j:
                ws.merge_cells(start_row=2, start_column=j + 1, end_row=2, end_column=k + 1)
            j = k + 1
        else:
            j += 1

    for j, (l1, l2, l3) in enumerate(lv, start=1):
        if l3:
            cell = ws.cell(row=3, column=j, value=l3)
            cell.font = white
            cell.alignment = center
            cell.fill = lvl3_fill(l1, l2)

    _hotelnames = ("Bentonville", "MGallery")
    DATA0 = 4
    pct = {i for i, c0 in enumerate(r.columns) if c0.startswith("Occ Forecast %")}
    signed_cols = ("BAR Chg 7d", "Lead Gap vs Comp")
    chg = {i for i, c0 in enumerate(r.columns)
           if (any(h in c0 for h in _hotelnames) and c0.endswith("Change"))
           or c0 in signed_cols}
    money = {i for i, c0 in enumerate(r.columns)
             if (c0 in ("BAR", "Comp Set Avg", "Last Room Value", "Estimated ADR")
                 or "ADR" in c0 or any(h in c0 for h in _hotelnames))
             and i not in chg}
    otb_sold_col = None
    for i, c0 in enumerate(r.columns):
        if c0 == "Rooms Left to Sell":
            otb_sold_col = i

    def _action_fill(txt):
        t = str(txt)
        if "Raise" in t:      return PatternFill("solid", fgColor="F8D7DA")
        if "Comp moved" in t: return PatternFill("solid", fgColor="CFE2FF")
        if "At " in t:        return PatternFill("solid", fgColor="FFF3CD")
        if "Hold" in t:       return PatternFill("solid", fgColor="E9ECEF")
        return None

    for ri, (_, row) in enumerate(r.iterrows(), start=DATA0):
        for ci2, col in enumerate(r.columns, start=1):
            v = row[col]
            cell = ws.cell(row=ri, column=ci2, value=(None if pd.isna(v) else v))
            cell.border = thin
            cell.font = Font(size=9)
            # center everything; left-align the long-text "Events" column
            halign = "left" if col == "Events" else "center"
            cell.alignment = Alignment(horizontal=halign, vertical="center")
            if col == "⚡ Action":
                af = _action_fill(v)
                if af is not None:
                    cell.fill = af
            idx = ci2 - 1
            if idx in pct:
                cell.number_format = "0.0%"
            elif idx in chg:
                cell.number_format = "+#,##0;-#,##0;\"\""
            elif idx in money:
                cell.number_format = "$#,##0"
            elif isinstance(v, float) and float(v).is_integer():
                cell.number_format = "#,##0"

    if otb_sold_col is not None and len(r) > 0:
        L = get_column_letter(otb_sold_col + 1)
        rng = f"{L}{DATA0}:{L}{DATA0 + len(r) - 1}"
        # Rooms Left to Sell: RED = few left (near sell-out), GREEN = lots left
        ws.conditional_formatting.add(
            rng,
            ColorScaleRule(start_type="min", start_color="F8696B",         # red (few left)
                           mid_type="percentile", mid_value=50, mid_color="FFEB84",  # yellow
                           end_type="max", end_color="63BE7B"))            # green (lots left)

    ws.freeze_panes = "C4"
    for j in range(1, n + 1):
        ws.column_dimensions[get_column_letter(j)].width = 12
    ws.column_dimensions["A"].width = 16   # ⚡ Action
    ws.column_dimensions["E"].width = 22   # Events

    # Excel AutoFilter so the team can filter/sort right in Excel.
    # Anchored on the bottom header row (row 3) through the last data row.
    last_row = DATA0 + len(r) - 1
    ws.auto_filter.ref = f"A3:{get_column_letter(n)}{last_row}"

    # ================================================================= #
    #  KPI Summary tab (self-contained snapshot for sharing)
    # ================================================================= #
    _build_kpi_sheet(wb, r, NAVY, TEAL, ORANGE, PURPLE, GREEN, kpi_ctx)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _month_key(d):
    return (d.year, d.month)


def _month_label(y, m):
    import calendar
    return f"{calendar.month_abbr[m]} {y}"


def _build_kpi_sheet(wb, r, NAVY, TEAL, ORANGE, PURPLE, GREEN, kpi_ctx=None):
    """Monthly recap: OTB vs Forecast vs Budget for current / next / following month,
    pulled straight from the Data Extract (budget + forecast columns)."""
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    import calendar
    kpi_ctx = kpi_ctx or {}
    as_of = kpi_ctx.get("as_of", pd.Timestamp.today().date())
    raw = kpi_ctx.get("data_raw")          # the daily Data Extract (all columns)

    ks = wb.create_sheet("KPI Summary", 0)
    white = Font(color="FFFFFF", bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    thin = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

    ks.merge_cells("A1:K1")
    t = ks["A1"]; t.value = "The Compton - Monthly Revenue Recap  (On the Books vs Forecast vs Budget)"
    t.font = Font(color="FFFFFF", bold=True, size=13); t.alignment = center
    t.fill = PatternFill("solid", fgColor=NAVY); ks.row_dimensions[1].height = 28

    if raw is None or "Occupancy Date" not in getattr(raw, "columns", []):
        ks["A3"] = "Upload the Data Extract that includes Budget/Forecast to populate this recap."
        ks["A3"].font = Font(size=11, italic=True, color="884400")
        return ks

    d = raw.copy()
    d["Occupancy Date"] = pd.to_datetime(d["Occupancy Date"], errors="coerce")
    d = d.dropna(subset=["Occupancy Date"])

    def col(name):
        return pd.to_numeric(d[name], errors="coerce") if name in d.columns else pd.Series(0.0, index=d.index)

    # choose "My Forecast" when it has data, else fall back to the system forecast
    myf_rev = col("My Forecast Revenue This Year")
    fcst_rev_src = "My Forecast Revenue This Year" if myf_rev.fillna(0).abs().sum() > 0 else "Forecasted Room Revenue This Year"
    myf_occ = col("My Forecast Occupancy - Total This Year")
    fcst_occ_src = "My Forecast Occupancy - Total This Year" if myf_occ.fillna(0).abs().sum() > 0 else "Occupancy Forecast - Total This Year"
    fcst_kind = "Your forecast" if fcst_rev_src.startswith("My") else "System forecast"

    d["_otb_rev"]  = col("Booked Room Revenue This Year")
    d["_otb_rms"]  = col("Occupancy On Books This Year")
    d["_fc_rev"]   = col(fcst_rev_src)
    d["_fc_rms"]   = col(fcst_occ_src)
    d["_bud_rev"]  = col("Budget Room Revenue This Year")
    d["_bud_rms"]  = col("Budget Occupancy - Total This Year")
    # segment (rooms only — the file has no segment-level revenue)
    myf_g = col("My Forecast Occupancy - Group This Year")
    fc_g_src = "My Forecast Occupancy - Group This Year" if myf_g.fillna(0).abs().sum() > 0 else "Occupancy Forecast - Group This Year"
    myf_t = col("My Forecast Occupancy - Transient This Year")
    fc_t_src = "My Forecast Occupancy - Transient This Year" if myf_t.fillna(0).abs().sum() > 0 else "Occupancy Forecast - Transient This Year"
    d["_otb_g"] = col("Rooms Sold - Group This Year")
    d["_otb_t"] = col("Rooms Sold - Transient This Year")
    d["_fc_g"]  = col(fc_g_src)
    d["_fc_t"]  = col(fc_t_src)
    d["_bud_g"] = col("Budget Occupancy - Group This Year")
    d["_bud_t"] = col("Budget Occupancy - Transient This Year")
    d["_ym"] = list(zip(d["Occupancy Date"].dt.year, d["Occupancy Date"].dt.month))

    base = pd.Timestamp(as_of)
    months = [(base + pd.DateOffset(months=k)).year for k in range(3)], [(base + pd.DateOffset(months=k)).month for k in range(3)]
    months = list(zip(months[0], months[1]))

    def mstat(ym):
        s = d[d["_ym"] == ym]
        return dict(otb_rev=s["_otb_rev"].sum(), otb_rms=s["_otb_rms"].sum(),
                    fc_rev=s["_fc_rev"].sum(), fc_rms=s["_fc_rms"].sum(),
                    bud_rev=s["_bud_rev"].sum(), bud_rms=s["_bud_rms"].sum(),
                    otb_g=s["_otb_g"].sum(), otb_t=s["_otb_t"].sum(),
                    fc_g=s["_fc_g"].sum(), fc_t=s["_fc_t"].sum(),
                    bud_g=s["_bud_g"].sum(), bud_t=s["_bud_t"].sum())

    ks.merge_cells("A2:K2")
    s2 = ks["A2"]
    s2.value = f"As of {base:%b %d, %Y}   -   Forecast source: {fcst_kind}   -   $ = room revenue, ADR = revenue / room-nights"
    s2.font = Font(size=10, color="555555"); s2.alignment = center

    heads = ["Month", "OTB Rooms", "OTB Revenue", "OTB ADR",
             "Fcst Rooms", "Fcst Revenue", "Fcst ADR",
             "Budget Revenue", "Fcst vs Budget", "Fcst vs Bud %", "OTB vs Budget"]
    hr = 4
    for j, h in enumerate(heads, start=1):
        c = ks.cell(row=hr, column=j, value=h)
        c.font = white; c.alignment = center; c.fill = PatternFill("solid", fgColor=TEAL)
    ks.row_dimensions[hr].height = 26

    monthcolors = [ORANGE, PURPLE, "3A6EA5"]
    for i, ym in enumerate(months):
        m = mstat(ym)
        otb_adr = (m["otb_rev"] / m["otb_rms"]) if m["otb_rms"] else np.nan
        fc_adr  = (m["fc_rev"] / m["fc_rms"]) if m["fc_rms"] else np.nan
        fvb = (m["fc_rev"] - m["bud_rev"]) if m["bud_rev"] else np.nan
        fvb_pct = (fvb / m["bud_rev"]) if m["bud_rev"] else np.nan
        ovb = (m["otb_rev"] - m["bud_rev"]) if m["bud_rev"] else np.nan
        vals = [f"{calendar.month_abbr[ym[1]]} {ym[0]}", m["otb_rms"], m["otb_rev"], otb_adr,
                m["fc_rms"], m["fc_rev"], fc_adr, m["bud_rev"], fvb, fvb_pct, ovb]
        rr = hr + 1 + i
        for j, v in enumerate(vals, start=1):
            c = ks.cell(row=rr, column=j, value=(None if (isinstance(v, float) and pd.isna(v)) else v))
            c.border = thin; c.alignment = center; c.font = Font(size=10)
            if j == 1:
                c.font = Font(size=10, bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=monthcolors[i % 3])
            elif j in (3, 4, 6, 7, 8, 9, 11):
                c.number_format = "$#,##0"
            elif j == 10:
                c.number_format = "0.0%"
            else:
                c.number_format = "#,##0"
            if j in (9, 10, 11) and isinstance(v, (int, float)) and pd.notna(v):
                c.font = Font(size=10, bold=True, color=("1a7f37" if v >= 0 else "C00000"))

    # ---- Group vs Transient split (rooms) ----
    sr = hr + 5
    ks.cell(row=sr, column=1,
            value="Rooms by segment - Group vs Transient (room-nights; file has no segment revenue)"
            ).font = Font(bold=True, size=12)
    seg_heads = ["Month", "Seg", "OTB Rooms", "Fcst Rooms", "Budget Rooms",
                 "Fcst vs Bud", "OTB vs Bud"]
    sh = sr + 1
    for j, h in enumerate(seg_heads, start=1):
        c = ks.cell(row=sh, column=j, value=h)
        c.font = white; c.alignment = center; c.fill = PatternFill("solid", fgColor=NAVY)
    rowp = sh + 1
    for i, ym in enumerate(months):
        m = mstat(ym)
        for seg, (otb, fc, bud, segfill) in [
            ("Group",     (m["otb_g"], m["fc_g"], m["bud_g"], GREEN)),
            ("Transient", (m["otb_t"], m["fc_t"], m["bud_t"], PURPLE))]:
            fvb = fc - bud
            ovb = otb - bud
            vals = [calendar.month_abbr[ym[1]] + " " + str(ym[0]), seg, otb, fc, bud, fvb, ovb]
            for j, v in enumerate(vals, start=1):
                c = ks.cell(row=rowp, column=j, value=v)
                c.border = thin; c.alignment = center; c.font = Font(size=10)
                if j == 1:
                    c.font = Font(size=10, bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor=monthcolors[i % 3])
                elif j == 2:
                    c.font = Font(size=10, bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor=segfill)
                else:
                    c.number_format = "#,##0"
                if j in (6, 7) and isinstance(v, (int, float)) and pd.notna(v):
                    c.font = Font(size=10, bold=True, color=("1a7f37" if v >= 0 else "C00000"))
            rowp += 1

    # ---- Action breakdown (future dates) ----
    dl = pd.to_numeric(r.get("Days Left"), errors="coerce")
    fut = r[dl >= 0] if dl is not None else r
    act = fut.get("\u26a1 Action", pd.Series(dtype=str)).fillna("")
    br = rowp + 1
    ks.cell(row=br, column=1, value="Action breakdown (future dates)").font = Font(bold=True, size=12)
    bd = [("Raise", int(act.str.contains("Raise").sum()), "F8D7DA"),
          ("Comp moved", int(act.str.contains("Comp moved").sum()), "CFE2FF"),
          ("At limit", int(act.str.contains("At ").sum()), "FFF3CD"),
          ("Hold", int(act.str.contains("Hold").sum()), "E9ECEF")]
    for k, (name, cnt, fill) in enumerate(bd):
        c1 = ks.cell(row=br+1, column=1+k*2, value=name)
        c1.fill = PatternFill("solid", fgColor=fill); c1.alignment = center; c1.font = Font(size=10, bold=True)
        c2 = ks.cell(row=br+1, column=2+k*2, value=cnt)
        c2.fill = PatternFill("solid", fgColor=fill); c2.alignment = center; c2.font = Font(size=12, bold=True)

    for cc, w in zip("ABCDEFGHIJK", [11,10,13,10,10,13,10,13,13,12,13]):
        ks.column_dimensions[cc].width = w
    ks.sheet_view.showGridLines = False
    return ks


st.title("The Compton — Daily Pick-Up Report Builder")
st.caption("Upload your four SynXis / rate-shop exports and get the full pick-up "
           "report instantly. No copy/paste required.")

with st.sidebar:
    st.header("1 - Upload your files")
    ups = st.file_uploader(
        "Drop in PCDC, Data Extract, Market Seg and the Rate-Shop file "
        "(all four at once is fine)",
        type=["xlsx"], accept_multiple_files=True,
    )

if not ups:
    st.info("Upload your files in the sidebar to build the report.")
    st.stop()

parts = detect_and_load(ups)

# 'Days Left' reference date, pulled from the PCDC report (Analysis Start Date /
# Generated On).  Falls back to today's date if the PCDC criteria isn't found.
pcdc_asof = parts.get("pcdc_asof") or pd.Timestamp.today().date()

with st.sidebar:
    st.header("2 - Settings")
    as_of = st.date_input(
        "Report as-of date (drives 'Days Left')",
        value=pcdc_asof,
        help="Auto-filled from the PCDC report's Analysis Start Date. "
             "Override here if you want a different reference date.",
    )
    comp_method = st.radio("Comp Set Avg method", ["Average", "Median"], horizontal=True)
    show_change = st.radio(
        "Show competitor rate change vs.",
        ["None", "Yesterday", "3 days ago", "7 days ago"],
        help="Adds a 'Change' column beside each competitor, pulled from the "
             "matching 'vs.' tab in the rate-shop file.",
    )
    show_own_change = st.checkbox(
        "Show our BAR change (7 days)", value=True,
        help="Adds a 'BAR Chg 7d' column showing how our own rate moved over the "
             "last 7 days (from the rate-shop file).",
    )

    st.markdown("---")
    st.header("3 - Signal sensitivity")
    st.caption("The Compton prices as the market leader (#1 in the comp set). "
               "Tune the three triggers below.")
    with st.expander("🔴 Rate leadership (#1)", expanded=False):
        lead_buffer = st.slider("Minimum $ we want to sit ABOVE the top competitor",
                                -20, 100, 0, 5,
                                help="If our BAR isn't at least this far above the highest "
                                     "competitor, flag 'Raise · reclaim #1'. 0 = just need "
                                     "to be strictly highest.")
    with st.expander("🔴 OTB compression", expanded=False):
        comp_occ = st.slider("OTB occupancy % that means compression", 60, 100, 85, 1,
                             help="At or above this rooms-on-the-books occupancy, flag Raise.") / 100
    with st.expander("🔵 Comp-set movement", expanded=False):
        move_thresh = st.slider("Flag when a competitor moves rate by at least ($)",
                                5, 100, 20, 5,
                                help="Uses the change window selected above (Yesterday / 3 / 7 days). "
                                     "Set 'Show competitor rate change vs.' to enable this flag.")

    thr = {"lead_buffer": lead_buffer, "comp_occ": comp_occ, "move_thresh": move_thresh}

    st.markdown("---")
    st.caption("Budget & forecast are pulled automatically from the Data Extract "
               "(Budget / Forecast columns) into the KPI Summary tab.")
    st.caption("Tip: leave the file names as SynXis exports them — the app "
               "auto-detects each report by its sheets.")

c1, c2, c3, c4 = st.columns(4)
c1.metric("PCDC",         "OK" if parts["pcdc"] is not None else "-")
c2.metric("Data Extract", "OK" if parts["data"] is not None else "-")
c3.metric("Market Seg",   "OK" if parts["mseg"] is not None else "-")
c4.metric("Rate Shop",    "OK" if parts["shop"] is not None else "-")

if parts.get("pcdc_asof"):
    st.caption(f"📅 'Days Left' counts from the PCDC report date: "
               f"**{pd.Timestamp(pcdc_asof):%b %d, %Y}** (Analysis Start Date). "
               f"You can change this in the sidebar.")

report = build_report(parts, as_of, comp_method, show_change, thr, show_own_change)

# =============================================================== #
#  KPI strip — portfolio health at a glance
# =============================================================== #
_fut = report[report["Days Left"] >= 0].copy()
def _s(col):
    return pd.to_numeric(_fut.get(col), errors="coerce") if col in _fut else pd.Series(dtype=float)

otb_ty   = _s("Rooms Sold | Total Hotel | Current").sum()
otb_stly = _s("Rooms OTB STLY | Total Hotel | Current").sum()
pace_abs = otb_ty - otb_stly
pace_pct = (pace_abs / otb_stly * 100) if otb_stly else np.nan

n7  = _fut[_fut["Days Left"] <= 7]
n30 = _fut[_fut["Days Left"] <= 30]
occ7  = pd.to_numeric(n7.get("Occ Forecast % | Total Hotel"), errors="coerce").mean()
occ30 = pd.to_numeric(n30.get("Occ Forecast % | Total Hotel"), errors="coerce").mean()

act = _fut.get("⚡ Action", pd.Series(dtype=str)).fillna("")
n_raise = int(act.str.contains("Raise").sum())
n_moved = int(act.str.contains("Comp moved").sum())

# How many future dates are we #1 (price leader) on?
lg = pd.to_numeric(_fut.get("Lead Gap vs Comp"), errors="coerce")
n_leader = int((lg > 0).sum())
n_leader_pct = (n_leader / lg.notna().sum() * 100) if lg.notna().sum() else np.nan

bar_ty  = pd.to_numeric(_fut.get("BAR"), errors="coerce")
comp_ty = pd.to_numeric(_fut.get("Comp Set Avg"), errors="coerce")
rate_gap = (bar_ty - comp_ty).mean()

# =============================================================== #
#  Monthly recap — current / next / following month (OTB vs Fcst vs Budget)
# =============================================================== #
st.subheader("Monthly recap — OTB vs Forecast vs Budget")
_raw = parts.get("data")
if _raw is not None and "Occupancy Date" in getattr(_raw, "columns", []):
    import calendar as _cal
    _d = _raw.copy()
    _d["Occupancy Date"] = pd.to_datetime(_d["Occupancy Date"], errors="coerce")
    _d = _d.dropna(subset=["Occupancy Date"])

    def _c(name):
        return pd.to_numeric(_d[name], errors="coerce") if name in _d.columns else pd.Series(0.0, index=_d.index)

    _myf = _c("My Forecast Revenue This Year")
    _fc_rev_src = "My Forecast Revenue This Year" if _myf.fillna(0).abs().sum() > 0 else "Forecasted Room Revenue This Year"
    _myfo = _c("My Forecast Occupancy - Total This Year")
    _fc_occ_src = "My Forecast Occupancy - Total This Year" if _myfo.fillna(0).abs().sum() > 0 else "Occupancy Forecast - Total This Year"
    _fc_kind = "your forecast" if _fc_rev_src.startswith("My") else "system forecast"

    _d["_otb_rev"] = _c("Booked Room Revenue This Year")
    _d["_otb_rms"] = _c("Occupancy On Books This Year")
    _d["_fc_rev"]  = _c(_fc_rev_src)
    _d["_fc_rms"]  = _c(_fc_occ_src)
    _d["_bud_rev"] = _c("Budget Room Revenue This Year")
    _d["_bud_rms"] = _c("Budget Occupancy - Total This Year")
    _d["_ym"] = list(zip(_d["Occupancy Date"].dt.year, _d["Occupancy Date"].dt.month))

    _base = pd.Timestamp(as_of)
    _months = [((_base + pd.DateOffset(months=k)).year, (_base + pd.DateOffset(months=k)).month) for k in range(3)]

    _rows = []
    for ym in _months:
        s = _d[_d["_ym"] == ym]
        otb_rev, otb_rms = s["_otb_rev"].sum(), s["_otb_rms"].sum()
        fc_rev, fc_rms   = s["_fc_rev"].sum(), s["_fc_rms"].sum()
        bud_rev          = s["_bud_rev"].sum()
        _rows.append({
            "Month": f"{_cal.month_abbr[ym[1]]} {ym[0]}",
            "OTB Rooms": otb_rms, "OTB Revenue": otb_rev,
            "OTB ADR": (otb_rev / otb_rms) if otb_rms else np.nan,
            "Fcst Rooms": fc_rms, "Fcst Revenue": fc_rev,
            "Budget Revenue": bud_rev,
            "Fcst vs Budget": (fc_rev - bud_rev) if bud_rev else np.nan,
            "Fcst vs Bud %": ((fc_rev - bud_rev) / bud_rev) if bud_rev else np.nan,
            "OTB vs Budget": (otb_rev - bud_rev) if bud_rev else np.nan,
        })
    _mdf = pd.DataFrame(_rows).set_index("Month")

    def _hl(v):
        if isinstance(v, (int, float)) and pd.notna(v):
            return "color:#1a7f37;font-weight:600" if v >= 0 else "color:#cf222e;font-weight:600"
        return ""
    _sty = (_mdf.style
            .format({"OTB Rooms": "{:,.0f}", "OTB Revenue": "${:,.0f}", "OTB ADR": "${:,.0f}",
                     "Fcst Rooms": "{:,.0f}", "Fcst Revenue": "${:,.0f}",
                     "Budget Revenue": "${:,.0f}", "Fcst vs Budget": "${:,.0f}",
                     "Fcst vs Bud %": "{:.1%}", "OTB vs Budget": "${:,.0f}"}, na_rep="–")
            .map(_hl, subset=["Fcst vs Budget", "Fcst vs Bud %", "OTB vs Budget"]))
    st.dataframe(_sty, use_container_width=True)
    st.caption(f"Forecast source: {_fc_kind}. Budget & forecast pulled from the Data Extract. "
               "Full recap + Group/Transient split are in the Excel export's KPI Summary tab.")
else:
    st.info("Upload the Data Extract with Budget/Forecast columns to see the monthly recap.")

st.subheader("Portfolio snapshot")
k = st.columns(6)
k[0].metric("OTB vs STLY", f"{otb_ty:,.0f}",
            f"{pace_abs:+,.0f} ({pace_pct:+.0f}%)" if pd.notna(pace_pct) else None)
k[1].metric("Next 7 Occ %",  f"{occ7*100:,.0f}%" if pd.notna(occ7) else "–")
k[2].metric("#1 in comp set", f"{n_leader_pct:,.0f}%" if pd.notna(n_leader_pct) else "–",
            f"{n_leader} dates")
k[3].metric("🔴 Raise dates", f"{n_raise}")
k[4].metric("🔵 Comp moved", f"{n_moved}")
k[5].metric("Rate vs comp", f"{rate_gap:+,.0f}" if pd.notna(rate_gap) else "–")

# =============================================================== #
#  Filters
# =============================================================== #
st.subheader("Filters")
min_d, max_d = report["Date"].min(), report["Date"].max()

fa, fb, fc = st.columns([1.2, 1, 1])
window = fa.selectbox("Booking window",
                      ["Custom", "Next 7 days", "Next 14 days", "Next 30 days",
                       "Next 60 days", "Next 90 days", "All future dates"],
                      index=6)
dow_pick = fb.multiselect("Day of week", ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
                          default=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])
action_pick = fc.multiselect("Action", ["🔴 Raise", "🔵 Comp moved", "🟠 At limit", "⚪ Hold"],
                             default=["🔴 Raise", "🔵 Comp moved", "🟠 At limit", "⚪ Hold"])

fd, fe, ff = st.columns([1, 1, 1])
events_only = fd.checkbox("Events only", value=False)
occ_band = fe.select_slider("Occ Forecast % band",
                            options=[0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
                            value=(0, 100))
var_thresh = ff.number_input("Min |variance vs STLY| (rooms)", min_value=0, value=0, step=1)

# custom date range (used when window == Custom)
start, end = min_d, max_d          # defaults so the filename always has a range
if window == "Custom":
    d1, d2 = st.columns(2)
    default_start = max(min_d, as_of)
    start = d1.date_input("From", value=default_start, min_value=min_d, max_value=max_d)
    end   = d2.date_input("To",   value=max_d,         min_value=min_d, max_value=max_d)

# ---- apply filters ---- #
v = report.copy()
win_map = {"Next 7 days": 7, "Next 14 days": 14, "Next 30 days": 30,
           "Next 60 days": 60, "Next 90 days": 90}
if window in win_map:
    v = v[(v["Days Left"] >= 0) & (v["Days Left"] <= win_map[window])]
elif window == "All future dates":
    v = v[v["Days Left"] >= 0]
else:  # Custom
    v = v[(v["Date"] >= start) & (v["Date"] <= end)]

dow_full = {"Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday", "Thu": "Thursday",
            "Fri": "Friday", "Sat": "Saturday", "Sun": "Sunday"}
keep_dow = [dow_full[d] for d in dow_pick]
v = v[v["DOW"].isin(keep_dow)]

if events_only:
    v = v[v["Events"].astype(str).str.strip() != ""]

if "⚡ Action" in v:
    a = v["⚡ Action"].fillna("")
    sel = pd.Series(False, index=v.index)
    if "🔴 Raise" in action_pick:      sel |= a.str.contains("Raise")
    if "🔵 Comp moved" in action_pick: sel |= a.str.contains("Comp moved")
    if "🟠 At limit" in action_pick:   sel |= a.str.contains("At ")
    if "⚪ Hold" in action_pick:       sel |= a.str.contains("Hold")
    sel |= (a.str.strip() == "")      # keep past/blank rows
    v = v[sel]

ofp_col = pd.to_numeric(v.get("Occ Forecast % | Total Hotel"), errors="coerce") * 100
lo, hi = occ_band
v = v[(ofp_col.isna()) | ((ofp_col >= lo) & (ofp_col <= hi))]

if var_thresh > 0 and "Rooms OTB STLY | Total Hotel | Change" in v:
    var_col = pd.to_numeric(v["Rooms OTB STLY | Total Hotel | Change"], errors="coerce").abs()
    v = v[var_col >= var_thresh]

view = v.reset_index(drop=True)

st.subheader(f"Pick-Up Report — {len(view)} dates")

pct_cols   = [c for c in view.columns if c.startswith("Occ Forecast %")]
shop_all   = [c for c in view.columns if c.startswith("Competitor Shops")]
shop_chg   = [c for c in shop_all if c.endswith("Change")]
shop_cur   = [c for c in shop_all if c not in shop_chg]
money_cols = (["BAR", "Comp Set Avg", "Last Room Value", "Estimated ADR"]
              + [c for c in view.columns if c.startswith("Booked ADR")])
rooms_chg  = [c for c in view.columns
              if (c.startswith("Rooms Sold") or c.startswith("Rooms OTB STLY"))
              and c.endswith("Change")]
var_cols   = ([c for c in view.columns if "Variance" in c] + shop_chg + rooms_chg
              + [c for c in ("Lead Gap vs Comp", "BAR Chg 7d") if c in view.columns])
otb_sold   = "Rooms Left to Sell"      # heat-map column

def _money(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return f"${v:,.0f}"
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)

def _signed(v):
    if isinstance(v, (int, float)) and not pd.isna(v) and v != 0:
        return f"{v:+,.0f}"
    return ""

fmt = {}
for c in pct_cols:
    fmt[c] = "{:.1%}"
for c in money_cols:
    fmt[c] = "${:,.0f}"
for c in shop_cur:
    fmt[c] = _money
for c in shop_chg:
    fmt[c] = _signed
for c in view.columns:
    if c in ("Days Left", "OOO", "Ovrbk") or c.startswith(("Rooms Sold", "Rooms OTB STLY",
             "Remaining Demand", "Occ Forecast |")) and c not in pct_cols:
        fmt.setdefault(c, "{:,.0f}")

styler = (view.style
          .format(fmt, na_rep="")
          .map(lambda v: "color:#1a7f37;font-weight:600" if isinstance(v, (int, float)) and v > 0
               else ("color:#cf222e;font-weight:600" if isinstance(v, (int, float)) and v < 0 else ""),
               subset=var_cols))


# Heat map on the OTB rooms sold column — pure-Python red->yellow->green
# gradient (no matplotlib needed, so it works on Streamlit Cloud out of the box).
def _heat_bg(series):
    # For "Rooms Left to Sell": GREEN = lots of rooms left (wide open),
    # RED = few rooms left (near sell-out).  Low value -> red, high value -> green.
    vals = pd.to_numeric(series, errors="coerce")
    lo, hi = vals.min(), vals.max()
    span = (hi - lo) or 1
    out = []
    for v in vals:
        if pd.isna(v):
            out.append("")
            continue
        t = (v - lo) / span                    # low rooms left -> red, high -> green
        if t < 0.5:
            f = t / 0.5
            red, grn, blu = 248, int(105 + f * (235 - 105)), int(107 + f * (132 - 107))
        else:
            f = (t - 0.5) / 0.5
            red, grn, blu = int(255 - f * (255 - 99)), int(235 - f * (235 - 190)), int(132 - f * (132 - 123))
        out.append(f"background-color: #{red:02X}{grn:02X}{blu:02X}")
    return out

# Heat map is applied directly inside the pick-up report on the OTB column
if otb_sold in view.columns and view[otb_sold].notna().any():
    styler = styler.apply(_heat_bg, subset=[otb_sold])

# Colour the ⚡ Action column so recommendations pop
def _action_bg(series):
    out = []
    for v in series.astype(str):
        if "Raise" in v:
            out.append("background-color:#F8D7DA;color:#842029;font-weight:600")
        elif "Comp moved" in v:
            out.append("background-color:#CFE2FF;color:#084298;font-weight:600")
        elif "At " in v:
            out.append("background-color:#FFF3CD;color:#664D03;font-weight:600")
        elif "Hold" in v:
            out.append("background-color:#E9ECEF;color:#495057")
        else:
            out.append("")
    return out

if "⚡ Action" in view.columns:
    styler = styler.apply(_action_bg, subset=["⚡ Action"])

st.dataframe(styler, use_container_width=True, height=560)

# filename reflects the actual filtered date range (falls back to the full range)
if len(view) > 0:
    _fn_start, _fn_end = view["Date"].min(), view["Date"].max()
else:
    _fn_start, _fn_end = start, end

st.download_button(
    "Download report as Excel",
    data=to_excel(view, {"data_raw": parts.get("data"), "as_of": as_of}),
    file_name=f"Compton_PickUp_{_fn_start:%Y%m%d}_{_fn_end:%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
