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
    df = pd.read_excel(file, sheet_name="Sheet1")
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

def get_last_n_months(series, n=6):
    """Get last N months of data"""
    valid = series.dropna()
    return valid.tail(n) if len(valid) > n else valid

def calculate_avg(series, n=3):
    """Calculate average of last N months"""
    valid = series.dropna()
    last_n = valid.tail(n) if len(valid) >= n else valid
    return last_n.mean() if len(last_n) > 0 else np.nan

# ---------------- HEADER ----------------
st.title("⛽ RO Performance Dashboard")

st.caption(
    "📱 **Mobile-Optimized** | Upload **SALES-TILL-NOV.xlsx** to analyze RO performance with KPIs, trends & maps"
)

# ---------------- FILE UPLOAD ----------------
file = st.file_uploader("📂 Upload SALES-TILL-NOV.xlsx", type=["xlsx"])

if file is None:
    st.info("👆 Upload the Excel file to begin analysis.")
    st.stop()

df = load_data_from_file(file)

# Basic validation
required_cols = ["RO Name", "SAP Code"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"❌ Missing required columns: {', '.join(missing)}")
    st.stop()

df = df.dropna(subset=["RO Name", "SAP Code"])

# ---------------- TOP BAR: SEARCH + FILTERS ----------------
st.markdown("### 🔍 Search & Filters")

# Mobile-friendly stacked layout
search_text = st.text_input(
    "Search RO",
    value="",
    placeholder="🔎 Enter RO name, SAP code, TA, or District...",
    help="Search across RO Name, SAP Code, Trading Area, and District"
)

# Filters in expandable section for mobile
with st.expander("🎯 Advanced Filters", expanded=False):
    col1, col2 = st.columns(2)
    
    districts = (["All"] +
                 sorted(df["District"].dropna().astype(str).unique().tolist())
                 ) if "District" in df.columns else ["All"]
    ta_names = (["All"] +
                sorted(df["Trading Area Name"].dropna().astype(str).unique().tolist())
                ) if "Trading Area Name" in df.columns else ["All"]
    
    with col1:
        sel_district = st.selectbox("📍 District", options=districts)
    with col2:
        sel_ta = st.selectbox("🏪 Trading Area", options=ta_names)

# Apply filters
filtered = df.copy()

if sel_district != "All" and "District" in df.columns:
    filtered = filtered[filtered["District"].astype(str) == sel_district]
if sel_ta != "All" and "Trading Area Name" in df.columns:
    filtered = filtered[filtered["Trading Area Name"].astype(str) == sel_ta]

# Apply search
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
    st.warning("⚠️ No ROs found. Try adjusting your search or filters.")
    st.stop()

st.success(f"✅ Found **{len(filtered)}** RO(s)")

# ---------------- RO SELECTION ----------------
st.markdown("### 🏢 Select Retail Outlet")

# Compact display
ro_display = (
    filtered["RO Name"].astype(str)
    + " | " + filtered["SAP Code"].astype(str)
    + ((" | " + filtered["TA Code"].astype(str)) if "TA Code" in filtered.columns else "")
)

sel_ro = st.selectbox("Select RO", options=ro_display, label_visibility="collapsed")

sel_sap = sel_ro.split("|"),[object Object],strip() if "|" in sel_ro else sel_ro.split("SAP:")[-1].split("|"),[object Object],strip()
ro_row = filtered[filtered["SAP Code"].astype(str) == sel_sap]

if ro_row.empty:
    st.error("❌ Selected RO not found.")
    st.stop()

ro_row = ro_row.iloc,[object Object],

# ---------------- RO SNAPSHOT ----------------
st.markdown("---")
st.markdown("### 📋 RO Details")

# Mobile-friendly stacked layout
st.markdown(f"""
**🏪 {ro_row.get('RO Name', 'N/A')}**  
**SAP Code:** {ro_row.get('SAP Code', 'N/A')}  
**OMC:** {ro_row.get('OMC', 'N/A')} | **Type:** {ro_row.get('PSU / PVT.', 'N/A')}  
**Status:** {'🟢 Active' if ro_row.get('Active(Y/N)', '') == 'Y' else '🔴 Inactive'}

**📍 Location Details:**  
- **District:** {ro_row.get('District', 'N/A')}  
- **Trading Area:** {ro_row.get('Trading Area Name', 'N/A')}  
- **Location:** {ro_row.get('Location', 'N/A')}  
- **Highway:** {ro_row.get('New NH', 'N/A')} | **CC/DC:** {ro_row.get('CC/DC', 'N/A')}
""")

# ---------------- KPIs SECTION ----------------
st.markdown("---")
st.markdown("### 📊 Key Performance Indicators")

# Get month columns
ms_cols, h_cols = get_month_cols(df)

# Calculate last month sales and 3-month average
if ms_cols:
    ms_series = ro_row[ms_cols]
    ms_last_month = ms_series.dropna().iloc[-1] if len(ms_series.dropna()) > 0 else np.nan
    ms_3m_avg = calculate_avg(ms_series, 3)
else:
    ms_last_month = np.nan
    ms_3m_avg = np.nan

if h_cols:
    hsd_series = ro_row[h_cols]
    hsd_last_month = hsd_series.dropna().iloc[-1] if len(hsd_series.dropna()) > 0 else np.nan
    hsd_3m_avg = calculate_avg(hsd_series, 3)
else:
    hsd_last_month = np.nan
    hsd_3m_avg = np.nan

# Mobile-friendly KPI layout (2 columns)
st.markdown("#### 🔵 MS (Motor Spirit) Performance")
k1, k2 = st.columns(2)

with k1:
    st.metric("📅 Last Month Sales", format_number(ms_last_month) + " KL", 
              help="Most recent month's MS sales")
with k2:
    st.metric("📈 3-Month Average", format_number(ms_3m_avg) + " KL",
              help="Average MS sales over last 3 months")

st.markdown("#### 🟠 HSD (High Speed Diesel) Performance")
k3, k4 = st.columns(2)

with k3:
    st.metric("📅 Last Month Sales", format_number(hsd_last_month) + " KL",
              help="Most recent month's HSD sales")
with k4:
    st.metric("📈 3-Month Average", format_number(hsd_3m_avg) + " KL",
              help="Average HSD sales over last 3 months")

# Additional KPIs in expandable section
with st.expander("📦 Additional Metrics", expanded=False):
    ty = ro_row.get("TY MS+ HSD", np.nan)
    ms_peak = ro_row.get("Peak Consumption MS", np.nan)
    hsd_peak = ro_row.get("Peak Consumption HSD", np.nan)
    ms_hist = ro_row.get("MS Historical", np.nan)
    hsd_hist = ro_row.get("HSD Historical", np.nan)
    
    a1, a2 = st.columns(2)
    with a1:
        st.metric("Total TY (MS+HSD)", format_number(ty) + " KL")
        st.metric("Peak MS Consumption", format_number(ms_peak) + " KL")
    with a2:
        st.metric("Peak HSD Consumption", format_number(hsd_peak) + " KL")
        st.metric("Historical MS/HSD", f"{format_number(ms_hist)} / {format_number(hsd_hist)} KL")

# ---------------- TRENDS SECTION (LAST 6 MONTHS) ----------------
st.markdown("---")
st.markdown("### 📈 Performance Trends (Last 6 Months)")

tab1, tab2 = st.tabs(["📊 Monthly Trends", "📅 Year-on-Year"])

with tab1:
    st.caption("Last 6 months MS & HSD sales comparison")
    
    if ms_cols and h_cols:
        # Get last 6 months
        ms_6m = get_last_n_months(ro_row[ms_cols], 6)
        hsd_6m = get_last_n_months(ro_row[h_cols], 6)
        
        min_len = min(len(ms_6m), len(hsd_6m))
        if min_len > 0:
            ms_6m = ms_6m.iloc[:min_len]
            hsd_6m = hsd_6m.iloc[:min_len]
            months = ms_6m.index.to_list()
            
            ts = pd.DataFrame({
                "Month": months,
                "MS": ms_6m.values,
                "HSD": hsd_6m.values
            })
            ts_long = ts.melt("Month", var_name="Product", value_name="Volume (KL)")
            ts_long = ts_long.replace({np.nan: 0})
            
            st.line_chart(ts_long, x="Month", y="Volume (KL)", color="Product", height=400)
            
            # Summary stats
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("MS Avg (6M)", format_number(ms_6m.mean()))
            with col2:
                st.metric("HSD Avg (6M)", format_number(hsd_6m.mean()))
            with col3:
                total_6m = ms_6m.sum() + hsd_6m.sum()
                st.metric("Total (6M)", format_number(total_6m))
        else:
            st.info("📭 Insufficient data for 6-month trend")
    else:
        st.info("📭 Monthly data columns not found")

with tab2:
    st.caption("Year-over-year comparison")
    
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
            st.bar_chart(yoy_df, x="Year", y="Volume", color="Product", height=400)
        else:
            st.info("📭 No year-on-year data available")
    else:
        st.info("📭 YoY columns not found")

# ---------------- MAP SECTION ----------------
st.markdown("---")
st.markdown("### 🗺️ Location Map")

map_tab1, map_tab2 = st.tabs(["📍 Selected RO", "🗺️ All Filtered ROs"])

with map_tab1:
    lat = ro_row.get("Latitude", np.nan)
    lon = ro_row.get("Longitude", np.nan)
    
    if pd.notna(lat) and pd.notna(lon):
        map_df = pd.DataFrame({
            "lat": [lat],
            "lon": [lon]
        })
        st.map(map_df, size=200, zoom=13)
        st.caption(f"📍 Coordinates: {lat:.6f}, {lon:.6f}")
    else:
        st.info("📍 No GPS coordinates available for this RO")

with map_tab2:
    if "Latitude" in filtered.columns and "Longitude" in filtered.columns:
        map_filtered = filtered.dropna(subset=["Latitude", "Longitude"]).copy()
        if not map_filtered.empty:
            map_filtered = map_filtered.rename(columns={"Latitude": "lat", "Longitude": "lon"})
            st.map(map_filtered[["lat", "lon"]], size=80, zoom=8)
            st.caption(f"📍 Showing {len(map_filtered)} ROs on map")
        else:
            st.info("📍 No coordinates available for filtered ROs")
    else:
        st.info("📍 Location data not found in file")

# ---------------- FOOTER ----------------
st.markdown("---")
st.caption("💡 **Tip:** Use landscape mode on mobile for better chart visibility | Data refreshed on upload")
