import streamlit as st
import pandas as pd
import numpy as np

# --------------- PAGE CONFIG ---------------
st.set_page_config(
    page_title="RO Performance & Map",
    page_icon="⛽",
    layout="wide",
)

# --------------- HELPERS ---------------

@st.cache_data
def load_data_from_file(file):
    # Engine explicit to avoid ImportError
    df = pd.read_excel(file, sheet_name="Sheet1", engine="openpyxl")
    df = df.rename(columns=lambda c: str(c).strip())
    return df

def get_month_cols(df):
    """Return ordered lists of MS and HSD monthly columns (wide format)."""
    ms_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("M0")] + \
              [c for c in df.columns if isinstance(c, str) and c.startswith("M1")]
    h_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("H0")] + \
             [c for c in df.columns if isinstance(c, str) and c.startswith("H1")]
    ms_cols = sorted(set(ms_cols), key=lambda x: x)
    h_cols = sorted(set(h_cols), key=lambda x: x)
    return ms_cols, h_cols

def get_yoy_cols(df):
    """Return MS and HSD YoY aggregate columns like MS2425, HSD2425."""
    ms_yoy = [c for c in df.columns if isinstance(c, str) and c.startswith("MS") and len(c) == 6]
    hsd_yoy = [c for c in df.columns if isinstance(c, str) and c.startswith("HSD") and len(c) == 7]
    ms_yoy = sorted(set(ms_yoy), key=lambda x: x)
    hsd_yoy = sorted(set(hsd_yoy), key=lambda x: x)
    return ms_yoy, hsd_yoy

def format_number(val, decimals=1):
    if pd.isna(val):
        return "NA"
    try:
        return f"{val:,.{decimals}f}"
    except Exception:
        return str(val)

def get_last_n_months_series(row, ms_cols, h_cols, n=6):
    """
    For a single RO row, extract MS & HSD monthly series and
    keep only last `n` months where at least one product has non-null data.
    """
    if not ms_cols or not h_cols:
        return None

    # Align by shared month columns (intersection)
    shared = [c for c in ms_cols if c in h_cols]
    if not shared:
        return None

    ms_series = row[shared]
    hsd_series = row[shared]

    # Identify last index with any non-null volume
    mask_valid = (~ms_series.isna()) | (~hsd_series.isna())
    if not mask_valid.any():
        return None

    # Last non-null position
    last_idx = np.where(mask_valid.values)[0][-1]
    start_idx = max(0, last_idx - (n - 1))

    sel_cols = shared[start_idx:last_idx + 1]
    ms_last = ms_series[sel_cols].fillna(0)
    hsd_last = hsd_series[sel_cols].fillna(0)

    out = pd.DataFrame({
        "Month": sel_cols,
        "MS": ms_last.values,
        "HSD": hsd_last.values,
    })
    return out

def get_last_month_values(ts_df):
    """
    From the last-6-month DataFrame, return last-month MS & HSD and month label.
    """
    if ts_df is None or ts_df.empty:
        return None, None, None
    last_row = ts_df.iloc[-1]
    return last_row["Month"], last_row["MS"], last_row["HSD"]

# --------------- HEADER ---------------

st.title("RO Performance & Location Dashboard")

st.caption(
    "Upload the latest **SALES-TILL-NOV.xlsx** and analyse any RO with KPIs, "
    "last 6 months trend and map view. Designed to work well on mobile screens."
)

# --------------- FILE UPLOAD ---------------

uploaded_file = st.file_uploader("Upload SALES-TILL-NOV.xlsx", type=["xlsx"])

if uploaded_file is None:
    st.info("Upload the latest Excel file to start.")
    st.stop()

df = load_data_from_file(uploaded_file)

required_cols = ["RO Name", "SAP Code"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns in file: {', '.join(missing)}")
    st.stop()

df = df.dropna(subset=["RO Name", "SAP Code"])

# --------------- SEARCH & FILTERS BAR ---------------

st.markdown("### Search & Filters")

top1, top2, top3 = st.columns([2.5, 1.2, 1.2])

with top1:
    search_text = st.text_input(
        "Search RO (Name / SAP / District / TA)",
        value="",
        placeholder="e.g. BALASORE, 143302, KENDRAPARA",
    )

districts = (["All"] +
             sorted(df["District"].dropna().astype(str).unique().tolist())
             ) if "District" in df.columns else ["All"]
ta_names = (["All"] +
            sorted(df["Trading Area Name"].dropna().astype(str).unique().tolist())
            ) if "Trading Area Name" in df.columns else ["All"]

with top2:
    sel_district = st.selectbox("District", options=districts)

with top3:
    sel_ta = st.selectbox("Trading Area", options=ta_names)

filtered = df.copy()

if sel_district != "All" and "District" in df.columns:
    filtered = filtered[filtered["District"].astype(str) == sel_district]

if sel_ta != "All" and "Trading Area Name" in df.columns:
    filtered = filtered[filtered["Trading Area Name"].astype(str) == sel_ta]

if search_text:
    q = search_text.strip().lower()
    mask = (
        filtered["RO Name"].astype(str).str.lower().str.contains(q)
        | filtered["SAP Code"].astype(str).str.lower().str.contains(q)
    )
    if "Trading Area Name" in filtered.columns:
        mask |= filtered["Trading Area Name"].astype(str).str.lower().str.contains(q)
    if "District" in filtered.columns:
        mask |= filtered["District"].astype(str).str.lower().str.contains(q)
    filtered = filtered[mask]

if filtered.empty:
    st.warning("No ROs found for current filters / search.")
    st.stop()

# --------------- RO SELECTION ---------------

st.markdown("### Select RO")

ro_display = (
    filtered["RO Name"].astype(str)
    + " | SAP:" + filtered["SAP Code"].astype(str)
    + ((" | TA:" + filtered["TA Code"].astype(str))
       if "TA Code" in filtered.columns else "")
)

sel_ro = st.selectbox("Outlet", options=ro_display)

sel_sap = sel_ro.split("SAP:")[-1].split("|")[0].strip()
ro_row = filtered[filtered["SAP Code"].astype(str) == sel_sap]

if ro_row.empty:
    st.error("Selected RO not found.")
    st.stop()

ro_row = ro_row.iloc[0]

# --------------- SNAPSHOT ---------------

st.markdown("### RO Snapshot")

snap1, snap2 = st.columns(2)

with snap1:
    st.markdown(
        f"**{ro_row.get('RO Name', '')}**  \n"
        f"SAP: {ro_row.get('SAP Code', '')}  \n"
        f"OMC: {ro_row.get('OMC', '')} | PSU/PVT: {ro_row.get('PSU / PVT.', '')}  \n"
        f"Active: {ro_row.get('Active(Y/N)', '')}"
    )

with snap2:
    st.markdown(
        f"District: {ro_row.get('District', '')}  \n"
        f"Trading Area: {ro_row.get('Trading Area Name', '')}  \n"
        f"Location: {ro_row.get('Location', '')}  \n"
        f"NH / CC-DC: {ro_row.get('New NH', '')} / {ro_row.get('CC/DC', '')}"
    )

# --------------- KPIs ---------------

st.markdown("### KPIs & Indicators")

# Base KPIs from file
ms_latest = ro_row.get("MS2526", np.nan)
hsd_latest = ro_row.get("HSD2526", np.nan)
ty_total = ro_row.get("TY MS+ HSD", np.nan)
ms_avg = ro_row.get("MS(AVG2526)", np.nan)
hsd_avg = ro_row.get("HSD(AVG2526)", np.nan)
ms_peak = ro_row.get("Peak Consumption MS", np.nan)
hsd_peak = ro_row.get("Peak Consumption HSD", np.nan)
ms_hist = ro_row.get("MS Historical", np.nan)
hsd_hist = ro_row.get("HSD Historical", np.nan)

# Last 6-month trend and last month volumes
ms_cols, h_cols = get_month_cols(df)
ts_last6 = get_last_n_months_series(ro_row, ms_cols, h_cols, n=6)
last_month_label, last_ms, last_hsd = get_last_month_values(ts_last6)

# 3-month rolling average based on trend
ms_3m_avg = hsd_3m_avg = np.nan
if ts_last6 is not None and len(ts_last6) >= 3:
    ms_3m_avg = ts_last6["MS"].tail(3).mean()
    hsd_3m_avg = ts_last6["HSD"].tail(3).mean()

# YOY comparison: latest 2 MS/HSD years if available
ms_yoy_cols, hsd_yoy_cols = get_yoy_cols(df)
ms_yoy_delta = hsd_yoy_delta = None
if len(ms_yoy_cols) >= 2:
    ms_yoy_delta = ro_row.get(ms_yoy_cols[-1], np.nan) - ro_row.get(ms_yoy_cols[-2], np.nan)
if len(hsd_yoy_cols) >= 2:
    hsd_yoy_delta = ro_row.get(hsd_yoy_cols[-1], np.nan) - ro_row.get(hsd_yoy_cols[-2], np.nan)

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric("MS 25-26 (KL)", format_number(ms_latest))
    st.metric("HSD 25-26 (KL)", format_number(hsd_latest))

with k2:
    st.metric("TY MS+HSD (KL)", format_number(ty_total))
    st.metric("Last Month MS (KL)",
              format_number(last_ms) if last_ms is not None else "NA")

with k3:
    st.metric("Last Month HSD (KL)",
              format_number(last_hsd) if last_hsd is not None else "NA")
    st.metric("3M Avg MS / HSD",
              f"{format_number(ms_3m_avg)}/{format_number(hsd_3m_avg)}"
              if not pd.isna(ms_3m_avg) and not pd.isna(hsd_3m_avg) else "NA")

with k4:
    st.metric("Peak MS / HSD",
              f"{format_number(ms_peak)}/{format_number(hsd_peak)}")
    yoy_text = "NA"
    if ms_yoy_delta is not None and hsd_yoy_delta is not None:
        yoy_text = f"ΔMS:{format_number(ms_yoy_delta)}/ΔHSD:{format_number(hsd_yoy_delta)}"
    st.metric("YoY Vol Change", yoy_text)

if last_month_label is not None:
    st.caption(f"Last month in trend: **{last_month_label}**")

# --------------- TRENDS ---------------

st.markdown("### Last 6 Months Trend")

tcol1, tcol2 = st.columns(2)

with tcol1:
    st.caption("MS / HSD Volumes (last 6 non-empty months)")

    if ts_last6 is not None and not ts_last6.empty:
        ts_long = ts_last6.melt("Month", var_name="Product", value_name="Volume")
        st.line_chart(ts_long, x="Month", y="Volume", color="Product")
    else:
        st.info("No recent monthly data available for this RO.")

with tcol2:
    st.caption("Year-on-Year MS / HSD (if available)")

    yoy_rows = []
    if ms_yoy_cols or hsd_yoy_cols:
        for col in ms_yoy_cols[-3:]:
            yoy_rows.append({"Year": col.replace("MS", ""), "Product": "MS",
                             "Volume": ro_row.get(col, np.nan)})
        for col in hsd_yoy_cols[-3:]:
            yoy_rows.append({"Year": col.replace("HSD", ""), "Product": "HSD",
                             "Volume": ro_row.get(col, np.nan)})

        yoy_df = pd.DataFrame(yoy_rows).dropna(subset=["Volume"])
        if not yoy_df.empty:
            st.bar_chart(yoy_df, x="Year", y="Volume", color="Product")
        else:
            st.info("No YoY data found for this RO.")
    else:
        st.info("YoY columns not present in file.")

# --------------- MAP ---------------

st.markdown("### Map View")

m1, m2 = st.columns(2)

with m1:
    st.caption("Selected RO")

    lat = ro_row.get("Latitude", np.nan)
    lon = ro_row.get("Longitude", np.nan)

    if pd.notna(lat) and pd.notna(lon):
        map_df = pd.DataFrame({
            "lat": [lat],
            "lon": [lon],
            "RO Name": [ro_row.get("RO Name", "")]
        })
        st.map(map_df, size=18)
    else:
        st.info("No coordinates for this RO.")

with m2:
    st.caption("All Filtered ROs")

    if "Latitude" in filtered.columns and "Longitude" in filtered.columns:
        map_filtered = filtered.dropna(subset=["Latitude", "Longitude"]).copy()
        if not map_filtered.empty:
            map_filtered = map_filtered.rename(columns={"Latitude": "lat", "Longitude": "lon"})
            st.map(map_filtered[["lat", "lon"]], size=10)
        else:
            st.info("No coordinates for filtered ROs.")
    else:
        st.info("Latitude/Longitude columns not found in file.")
