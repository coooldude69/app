import streamlit as st
import pandas as pd
import numpy as np

# ---------------- BASIC CONFIG ----------------
st.set_page_config(
    page_title="RO Performance & Map",
    page_icon="⛽",
    layout="wide",
)

# ---------------- HELPERS ----------------
@st.cache_data
def load_data_from_file(file):
    df = pd.read_excel(file, sheet_name="Sheet1", engine="openpyxl")
    df = df.rename(columns=lambda c: str(c).strip())
    return df

def get_month_cols(df):
    ms_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("M0")] + \
              [c for c in df.columns if isinstance(c, str) and c.startswith("M1")]
    h_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("H0")] + \
             [c for c in df.columns if isinstance(c, str) and c.startswith("H1")]
    ms_cols = sorted(set(ms_cols), key=lambda x: x)
    h_cols = sorted(set(h_cols), key=lambda x: x)
    return ms_cols, h_cols

def get_yoy_cols(df):
    ms_yoy = [c for c in df.columns if isinstance(c, str) and c.startswith("MS") and len(c) == 6]
    hsd_yoy = [c for c in df.columns if isinstance(c, str) and c.startswith("HSD") and len(c) == 7]
    ms_yoy = sorted(set(ms_yoy), key=lambda x: x)
    hsd_yoy = sorted(set(hsd_yoy), key=lambda x: x)
    return ms_yoy, hsd_yoy

def format_number(val):
    if pd.isna(val):
        return "NA"
    try:
        return f"{val:,.1f}"
    except Exception:
        return str(val)

def build_last_n_months_ts(row, ms_cols, h_cols, n=6):
    """
    For one RO, build a dataframe with last n months of MS/HSD.
    Uses the last column index where either MS or HSD is non-null.
    """
    if not ms_cols or not h_cols:
        return None

    # Use only columns that exist for both MS and HSD so Month axis aligns
    shared = [c for c in ms_cols if c in h_cols]
    if not shared:
        return None

    ms_series = row[shared]
    hsd_series = row[shared]

    # month is valid if either MS or HSD has data
    valid_mask = (~ms_series.isna()) | (~hsd_series.isna())
    if not valid_mask.any():
        return None

    last_idx = np.where(valid_mask.values)[-1][-1]
    start_idx = max(0, last_idx - (n - 1))

    sel_cols = shared[start_idx:last_idx + 1]
    ms_last = ms_series[sel_cols].fillna(0)
    hsd_last = hsd_series[sel_cols].fillna(0)

    ts = pd.DataFrame({
        "Month": sel_cols,
        "MS": ms_last.values,
        "HSD": hsd_last.values
    })
    return ts

# ---------------- HEADER ----------------
st.title("RO Performance & Location Dashboard")

st.caption(
    "Upload the latest **SALES-TILL-NOV.xlsx** and analyse any RO with KPIs, trends and map view. "
    "Optimised for mobile: use the search bar to jump to an outlet."
)

# ---------------- FILE UPLOAD ----------------
file = st.file_uploader("Upload SALES-TILL-NOV.xlsx", type=["xlsx"])

if file is None:
    st.info("Upload the latest Excel file to begin.")
    st.stop()

df = load_data_from_file(file)

# Basic validation
required_cols = ["RO Name", "SAP Code"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns in file: {', '.join(missing)}")
    st.stop()

df = df.dropna(subset=["RO Name", "SAP Code"])

# ---------------- TOP BAR: SEARCH + FILTERS ----------------
st.markdown("### Search & Filters")

t1, t2, t3 = st.columns([2.5, 1.2, 1.2])

with t1:
    search_text = st.text_input(
        "Search RO (name / SAP / TA / District)",
        value="",
        placeholder="e.g. BALASORE, 143302, NH-5A, KENDRAPARA",
    )

districts = (["All"] +
             sorted(df["District"].dropna().astype(str).unique().tolist())
             ) if "District" in df.columns else ["All"]
ta_names = (["All"] +
            sorted(df["Trading Area Name"].dropna().astype(str).unique().tolist())
            ) if "Trading Area Name" in df.columns else ["All"]

with t2:
    sel_district = st.selectbox("District", options=districts)
with t3:
    sel_ta = st.selectbox("Trading Area", options=ta_names)

filtered = df.copy()

if sel_district != "All" and "District" in df.columns:
    filtered = filtered[filtered["District"].astype(str) == sel_district]
if sel_ta != "All" and "Trading Area Name" in df.columns:
    filtered = filtered[filtered["Trading Area Name"].astype(str) == sel_ta]

if search_text:
    s = search_text.strip()
    mask = (
        df["RO Name"].astype(str).str.contains(s, case=False, na=False)
        | df["SAP Code"].astype(str).str.contains(s, case=False, na=False)
    )
    if "Trading Area Name" in df.columns:
        mask |= df["Trading Area Name"].astype(str).str.contains(s, case=False, na=False)
    if "District" in df.columns:
        mask |= df["District"].astype(str).str.contains(s, case=False, na=False)
    filtered = filtered[mask]

if filtered.empty:
    st.warning("No ROs found for current search / filters.")
    st.stop()

# ---------------- RO SELECTION ----------------
st.markdown("### Select RO")

ro_display = (
    filtered["RO Name"].astype(str)
    + " | SAP:" + filtered["SAP Code"].astype(str)
    + ((" | TA:" + filtered["TA Code"].astype(str)) if "TA Code" in filtered.columns else "")
)

sel_ro = st.selectbox("Outlet", options=ro_display)

sel_sap = sel_ro.split("SAP:")[-1].split("|")[0].strip()
ro_row = filtered[filtered["SAP Code"].astype(str) == sel_sap]

if ro_row.empty:
    st.error("Selected RO not found in filtered data.")
    st.stop()

ro_row = ro_row.iloc[0]

# ---------------- RO SNAPSHOT ----------------
st.markdown("### RO Snapshot")

c1, c2 = st.columns(2)

with c1:
    st.markdown(
        f"**{ro_row.get('RO Name', '')}**  \n"
        f"SAP: {ro_row.get('SAP Code', '')}  \n"
        f"OMC: {ro_row.get('OMC', '')} | PSU/PVT: {ro_row.get('PSU / PVT.', '')}  \n"
        f"Active: {ro_row.get('Active(Y/N)', '')}"
    )

with c2:
    st.markdown(
        f"District: {ro_row.get('District', '')}  \n"
        f"Trading Area: {ro_row.get('Trading Area Name', '')}  \n"
        f"Location: {ro_row.get('Location', '')}  \n"
        f"NH / CC-DC: {ro_row.get('New NH', '')} / {ro_row.get('CC/DC', '')}"
    )

# ---------------- KPIs SECTION ----------------
st.markdown("### MS / HSD KPIs")

ms_latest = ro_row.get("MS2526", np.nan)
hsd_latest = ro_row.get("HSD2526", np.nan)
ty = ro_row.get("TY MS+ HSD", np.nan)
ms_avg = ro_row.get("MS(AVG2526)", np.nan)
hsd_avg = ro_row.get("HSD(AVG2526)", np.nan)
ms_peak = ro_row.get("Peak Consumption MS", np.nan)
hsd_peak = ro_row.get("Peak Consumption HSD", np.nan)
ms_hist = ro_row.get("MS Historical", np.nan)
hsd_hist = ro_row.get("HSD Historical", np.nan)

# Build last 6 months time series for extra KPIs
ms_cols, h_cols = get_month_cols(df)
ts_last6 = build_last_n_months_ts(ro_row, ms_cols, h_cols, n=6)

last_month_label = last_ms = last_hsd = None
ms_3m_avg = hsd_3m_avg = np.nan

if ts_last6 is not None and not ts_last6.empty:
    last_row = ts_last6.iloc[-1]
    last_month_label = last_row["Month"]
    last_ms = last_row["MS"]
    last_hsd = last_row["HSD"]

    if len(ts_last6) >= 3:
        ms_3m_avg = ts_last6["MS"].tail(3).mean()
        hsd_3m_avg = ts_last6["HSD"].tail(3).mean()

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric("MS 25-26 (KL)", format_number(ms_latest))
    st.metric("HSD 25-26 (KL)", format_number(hsd_latest))

with k2:
    st.metric("TY MS+HSD (KL)", format_number(ty))
    st.metric("Last Month MS (KL)",
              format_number(last_ms) if last_ms is not None else "NA")

with k3:
    st.metric("Last Month HSD (KL)",
              format_number(last_hsd) if last_hsd is not None else "NA")
    st.metric("3M Avg MS (KL)",
              format_number(ms_3m_avg) if not pd.isna(ms_3m_avg) else "NA")

with k4:
    st.metric("3M Avg HSD (KL)",
              format_number(hsd_3m_avg) if not pd.isna(hsd_3m_avg) else "NA")
    st.metric("MS/HSD Hist. (KL)", f"{format_number(ms_hist)} / {format_number(hsd_hist)}")

if last_month_label is not None:
    st.caption(f"Last month plotted: **{last_month_label}**")

# ---------------- TRENDS SECTION ----------------
st.markdown("### Trends (Last 6 Months)")

t_col1, t_col2 = st.columns(2)

with t_col1:
    st.caption("Monthly MS / HSD (last 6 non-empty months)")
    if ts_last6 is not None and not ts_last6.empty:
        ts_long = ts_last6.melt("Month", var_name="Product", value_name="Volume")
        st.line_chart(ts_long, x="Month", y="Volume", color="Product")
    else:
        st.info("No recent monthly MS/HSD data available for this RO.")

with t_col2:
    st.caption("Year-on-Year MS / HSD")
    ms_yoy, hsd_yoy = get_yoy_cols(df)
    yoy_rows = []
    if ms_yoy or hsd_yoy:
        for col in ms_yoy:
            yoy_rows.append({"Year": col.replace("MS", ""), "Product": "MS",
                             "Volume": ro_row.get(col, np.nan)})
        for col in hsd_yoy:
            yoy_rows.append({"Year": col.replace("HSD", ""), "Product": "HSD",
                             "Volume": ro_row.get(col, np.nan)})
        yoy_df = pd.DataFrame(yoy_rows).dropna(subset=["Volume"])
        if not yoy_df.empty:
            st.bar_chart(yoy_df, x="Year", y="Volume", color="Product")
        else:
            st.info("No YoY data for this RO.")
    else:
        st.info("YoY MS/HSD columns not found in file.")

# ---------------- MAP SECTION ----------------
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
            "name": [ro_row.get("RO Name", "")]
        })
        st.map(map_df, size=18)
    else:
        st.info("No coordinates for this RO.")

with m2:
    st.caption("All filtered ROs")
    if "Latitude" in filtered.columns and "Longitude" in filtered.columns:
        map_filtered = filtered.dropna(subset=["Latitude", "Longitude"]).copy()
        if not map_filtered.empty:
            map_filtered = map_filtered.rename(columns={"Latitude": "lat", "Longitude": "lon"})
            st.map(map_filtered[["lat", "lon"]], size=10)
        else:
            st.info("No coordinates for filtered ROs.")
    else:
        st.info("Latitude/Longitude not found in file.")
