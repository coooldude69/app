import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pydeck as pdk

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

def month_col_labels_last_n(n, today=None):
    """Last n full calendar months before today."""
    if today is None:
        today = datetime.today()
    labels_m = []
    labels_h = []
    for i in range(n, 0, -1):
        d = today - relativedelta(months=i)
        mm = f"{d.month:02d}"
        yy = str(d.year)[-2:]
        labels_m.append(f"M{mm}{yy}")
        labels_h.append(f"H{mm}{yy}")
    return labels_m, labels_h

# ---------------- SESSION STATE: DATA PERSISTENCE ----------------
if "df" not in st.session_state:
    st.session_state.df = None
if "file_name" not in st.session_state:
    st.session_state.file_name = None

st.title("RO Performance & Location Dashboard")

st.caption(
    "Upload the latest **SALES-TILL-NOV.xlsx** once per session and analyse any RO "
    "with last 3/6 month KPIs and detailed map. Optimised for mobile."
)

# File uploader with persistence
uploaded_file = st.file_uploader(
    "Upload SALES-TILL-NOV.xlsx (will stay loaded until you upload a new file)",
    type=["xlsx"]
)

if uploaded_file is not None:
    st.session_state.df = load_data_from_file(uploaded_file)
    st.session_state.file_name = uploaded_file.name

if st.session_state.df is None:
    st.info("Upload the Excel file to start the analysis.")
    st.stop()

df = st.session_state.df

st.caption(f"Current file: **{st.session_state.file_name}**")

# Basic validation
required_cols = ["RO Name", "SAP Code"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns in file: {', '.join(missing)}")
    st.stop()

df = df.dropna(subset=["RO Name", "SAP Code"])

# ---------------- SEARCH & FILTERS ----------------
st.markdown("### Search & Filters")

search_col, filt_col = st.columns([2.5, 1.5])

with search_col:
    search_text = st.text_input(
        "Search RO (Name / SAP / TA / District)",
        value="",
        placeholder="e.g. BALASORE, 143302, NH-5A, KENDRAPARA",
    )

with filt_col:
    districts = (["All"] +
                 sorted(df["District"].dropna().astype(str).unique().tolist())
                 ) if "District" in df.columns else ["All"]
    ta_names = (["All"] +
                sorted(df["Trading Area Name"].dropna().astype(str).unique().tolist())
                ) if "Trading Area Name" in df.columns else ["All"]

    sel_district = st.selectbox("District", options=districts)
    sel_ta = st.selectbox("Trading Area", options=ta_names)

filtered = df.copy()

if sel_district != "All" and "District" in df.columns:
    filtered = filtered[filtered["District"].astype(str) == sel_district]
if sel_ta != "All" and "Trading Area Name" in df.columns:
    filtered = filtered[filtered["Trading Area Name"].astype(str) == sel_ta]

if search_text:
    s = search_text.strip()
    mask = (
        filtered["RO Name"].astype(str).str.contains(s, case=False, na=False)
        | filtered["SAP Code"].astype(str).str.contains(s, case=False, na=False)
    )
    if "Trading Area Name" in filtered.columns:
        mask |= filtered["Trading Area Name"].astype(str).str.contains(s, case=False, na=False)
    if "District" in filtered.columns:
        mask |= filtered["District"].astype(str).str.contains(s, case=False, na=False)
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

# ---------------- KPIs: LAST 3 MONTHS ----------------
st.markdown("### MS / HSD KPIs (Last 3 Calendar Months)")

ms_latest = ro_row.get("MS2526", np.nan)
hsd_latest = ro_row.get("HSD2526", np.nan)
ty = ro_row.get("TY MS+ HSD", np.nan)
ms_peak = ro_row.get("Peak Consumption MS", np.nan)
hsd_peak = ro_row.get("Peak Consumption HSD", np.nan)
ms_hist = ro_row.get("MS Historical", np.nan)
hsd_hist = ro_row.get("HSD Historical", np.nan)

m3_labels, h3_labels = month_col_labels_last_n(3)

ms_last3_vals = []
hsd_last3_vals = []

for m_col, h_col in zip(m3_labels, h3_labels):
    ms_val = ro_row.get(m_col, 0.0)
    hsd_val = ro_row.get(h_col, 0.0)
    try:
        ms_val = float(ms_val)
    except Exception:
        ms_val = 0.0
    try:
        hsd_val = float(hsd_val)
    except Exception:
        hsd_val = 0.0
    ms_last3_vals.append(ms_val)
    hsd_last3_vals.append(hsd_val)

if m3_labels:
    last_month_label = m3_labels[-1]
    last_ms = ms_last3_vals[-1]
    last_hsd = hsd_last3_vals[-1]
else:
    last_month_label = None
    last_ms = 0.0
    last_hsd = 0.0

ms_3m_avg = sum(ms_last3_vals) / 3 if ms_last3_vals else 0.0
hsd_3m_avg = sum(hsd_last3_vals) / 3 if hsd_last3_vals else 0.0

k1, k2 = st.columns(2)

with k1:
    st.metric("MS 25-26 (KL)", format_number(ms_latest))
    st.metric("TY MS+HSD (KL)", format_number(ty))
    st.metric("Last Month MS (KL)", format_number(last_ms))

with k2:
    st.metric("HSD 25-26 (KL)", format_number(hsd_latest))
    st.metric("Last Month HSD (KL)", format_number(last_hsd))
    st.metric("3M Avg MS / HSD (KL)",
              f"{format_number(ms_3m_avg)}/{format_number(hsd_3m_avg)}")

st.caption(
    f"Reference 3‑month window: "
    f"**{m3_labels[0]} – {m3_labels[-1]}** (last completed calendar months before today)."
)

st.markdown(
    f"Historical & peaks: MS Hist {format_number(ms_hist)}, "
    f"HSD Hist {format_number(hsd_hist)}, "
    f"Peak MS {format_number(ms_peak)}, Peak HSD {format_number(hsd_peak)}"
)

# ---------------- TRENDS: LAST 6 MONTHS ----------------
st.markdown("### Trends (Last 6 Calendar Months)")

m6_labels, h6_labels = month_col_labels_last_n(6)

ms_vals_6 = []
hsd_vals_6 = []
for m_col, h_col in zip(m6_labels, h6_labels):
    ms_val = ro_row.get(m_col, 0.0)
    hsd_val = ro_row.get(h_col, 0.0)
    try:
        ms_val = float(ms_val)
    except Exception:
        ms_val = 0.0
    try:
        hsd_val = float(hsd_val)
    except Exception:
        hsd_val = 0.0
    ms_vals_6.append(ms_val)
    hsd_vals_6.append(hsd_val)

ts6 = pd.DataFrame({
    "Month": m6_labels,
    "MS": ms_vals_6,
    "HSD": hsd_vals_6,
})

trend_col1, trend_col2 = st.columns(2)

with trend_col1:
    st.caption("Monthly MS / HSD (last 6 calendar months)")
    ts_long = ts6.melt("Month", var_name="Product", value_name="Volume")
    st.line_chart(ts_long, x="Month", y="Volume", color="Product")

with trend_col2:
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

# ---------------- MAPS (DETAILED) ----------------
st.markdown("### Map View (Detailed)")

# Selected RO map with tooltip
lat = ro_row.get("Latitude", np.nan)
lon = ro_row.get("Longitude", np.nan)

if pd.notna(lat) and pd.notna(lon):
    selected_df = pd.DataFrame({
        "lat": [lat],
        "lon": [lon],
        "RO Name": [ro_row.get("RO Name", "")],
        "District": [ro_row.get("District", "")],
        "TY_MS_HSD": [ty if not pd.isna(ty) else 0.0],
    })

    view_state = pdk.ViewState(
        latitude=lat,
        longitude=lon,
        zoom=12,
        pitch=30,
    )

    layer_selected = pdk.Layer(
        "ScatterplotLayer",
        data=selected_df,
        get_position=["lon", "lat"],
        get_radius=800,
        get_fill_color=[255, 0, 0, 200],
        pickable=True,
    )

    tooltip = {
        "text": "{RO Name}\nDistrict: {District}\nTY MS+HSD: {TY_MS_HSD}"
    }

    st.pydeck_chart(
        pdk.Deck(
            layers=[layer_selected],
            initial_view_state=view_state,
            tooltip=tooltip,
            map_style="mapbox://styles/mapbox/streets-v11",
        )
    )
else:
    st.info("No coordinates for this RO.")

st.markdown("#### All Filtered ROs")

if "Latitude" in filtered.columns and "Longitude" in filtered.columns:
    map_filtered = filtered.dropna(subset=["Latitude", "Longitude"]).copy()
    if not map_filtered.empty:
        map_filtered["TY_MS_HSD"] = (
            map_filtered.get("TY MS+ HSD")
            .fillna(0)
            .astype(float)
        )
        center_lat = map_filtered["Latitude"].astype(float).mean()
        center_lon = map_filtered["Longitude"].astype(float).mean()

        view_state_all = pdk.ViewState(
            latitude=center_lat,
            longitude=center_lon,
            zoom=8,
            pitch=0,
        )

        layer_all = pdk.Layer(
            "ScatterplotLayer",
            data=map_filtered,
            get_position=["Longitude", "Latitude"],
            get_radius=1000,
            get_fill_color=[0, 120, 255, 160],
            pickable=True,
        )

        tooltip_all = {
            "text": "{RO Name}\nDistrict: {District}\nTY MS+HSD: {TY_MS_HSD}"
        }

        st.pydeck_chart(
            pdk.Deck(
                layers=[layer_all],
                initial_view_state=view_state_all,
                tooltip=tooltip_all,
                map_style="mapbox://styles/mapbox/light-v9",
            )
        )
    else:
        st.info("No coordinates for filtered ROs.")
else:
    st.info("Latitude/Longitude columns not found in file.")
