import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
from sklearn.linear_model import LinearRegression
from scipy import stats

# ═══════════════════════════════════════════════════════════════════
# PAGE CONFIG & GLOBAL CSS
# ═══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="PetroAnalytics — Outlet Intelligence Portal",
    layout="wide",
    initial_sidebar_state="expanded",
)

PORTAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; color: #0f172a; }
#MainMenu, footer { visibility: hidden; }
header
.block-container { padding: 0 2rem 2rem 2rem !important; max-width: 1440px !important; }

/* ── Portal Header ── */
.portal-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 55%, #0e4d92 100%);
    padding: 1.8rem 2.5rem 1.4rem;
    margin: -1rem -2rem 1.5rem -2rem;
    border-bottom: 3px solid #f59e0b;
}
.portal-title { font-size: 1.75rem; font-weight: 700; color: #ffffff; letter-spacing: -0.02em; margin: 0; }
.portal-subtitle { font-size: 0.82rem; color: #94a3b8; margin-top: 0.25rem; font-weight: 400; letter-spacing: 0.04em; text-transform: uppercase; }
.portal-badge { display: inline-block; background: #f59e0b; color: #0f172a; font-size: 0.68rem; font-weight: 700; padding: 2px 8px; border-radius: 3px; letter-spacing: 0.08em; text-transform: uppercase; vertical-align: middle; margin-left: 10px; }
.portal-meta { margin-top: 0.8rem; display: flex; gap: 1.5rem; flex-wrap: wrap; }
.portal-meta-item { font-size: 0.75rem; color: #64748b; }
.portal-meta-item b { color: #cbd5e1; }

/* ── KPI Cards ── */
.kpi-card {
    background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px;
    padding: 1rem 1.2rem; border-left: 4px solid #3b82f6;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05); margin-bottom: 0.75rem;
    transition: box-shadow 0.15s ease;
}
.kpi-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.09); }
.kpi-card.amber  { border-left-color: #f59e0b; }
.kpi-card.green  { border-left-color: #22c55e; }
.kpi-card.red    { border-left-color: #ef4444; }
.kpi-card.purple { border-left-color: #8b5cf6; }
.kpi-card.teal   { border-left-color: #14b8a6; }
.kpi-card.indigo { border-left-color: #6366f1; }
.kpi-label { font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.09em; color: #64748b; font-weight: 600; margin-bottom: 0.3rem; }
.kpi-value { font-size: 1.5rem; font-weight: 700; color: #0f172a; line-height: 1; font-family: 'DM Mono', monospace; }
.kpi-sub { font-size: 0.7rem; color: #94a3b8; margin-top: 0.25rem; }
.kpi-delta { font-size: 0.75rem; font-weight: 600; margin-top: 0.3rem; }
.kpi-delta.up   { color: #16a34a; }
.kpi-delta.down { color: #dc2626; }
.kpi-delta.neutral { color: #d97706; }

/* ── Mini KPI for sub-metrics ── */
.mini-kpi {
    background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
    padding: 0.7rem 1rem; margin-bottom: 0.5rem;
}
.mini-kpi-label { font-size: 0.65rem; text-transform: uppercase; letter-spacing: 0.07em; color: #94a3b8; font-weight: 600; }
.mini-kpi-value { font-size: 1.1rem; font-weight: 700; color: #0f172a; font-family: 'DM Mono', monospace; }
.mini-kpi-sub { font-size: 0.65rem; color: #94a3b8; }

/* ── Section Cards ── */
.section-card {
    background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
    padding: 1.4rem; margin-bottom: 1.1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.section-title {
    font-size: 0.92rem; font-weight: 700; color: #0f172a;
    margin-bottom: 0.9rem; padding-bottom: 0.55rem;
    border-bottom: 1px solid #f1f5f9; letter-spacing: -0.01em;
}
.section-subtitle { font-size: 0.75rem; color: #64748b; font-weight: 400; margin-left: 6px; }

/* ── Analytics Row (MS/HSD split) ── */
.analytics-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem; }
.analytics-block {
    background: #f8fafc; border: 1px solid #e9eef5; border-radius: 10px;
    padding: 1rem 1.2rem;
}
.analytics-block.ms-block { border-top: 3px solid #f97316; }
.analytics-block.hsd-block { border-top: 3px solid #3b82f6; }
.analytics-type { font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 0.7rem; }
.analytics-type.ms  { color: #ea580c; }
.analytics-type.hsd { color: #2563eb; }
.analytics-row { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 0.45rem; }
.analytics-metric { font-size: 0.68rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.06em; }
.analytics-value { font-size: 1rem; font-weight: 700; font-family: 'DM Mono', monospace; color: #0f172a; }
.analytics-growth { font-size: 0.78rem; font-weight: 600; }
.analytics-growth.up { color: #16a34a; }
.analytics-growth.down { color: #dc2626; }
.analytics-growth.neutral { color: #d97706; }

/* ── Banners ── */
.warn-banner    { background: #fffbeb; border: 1px solid #fde68a; border-left: 4px solid #f59e0b; border-radius: 8px; padding: 0.65rem 1rem; margin-bottom: 0.8rem; font-size: 0.81rem; color: #92400e; }
.info-banner    { background: #eff6ff; border: 1px solid #bfdbfe; border-left: 4px solid #3b82f6; border-radius: 8px; padding: 0.65rem 1rem; margin-bottom: 0.8rem; font-size: 0.81rem; color: #1e3a8a; }
.error-banner   { background: #fef2f2; border: 1px solid #fecaca; border-left: 4px solid #ef4444; border-radius: 8px; padding: 0.65rem 1rem; margin-bottom: 0.8rem; font-size: 0.81rem; color: #7f1d1d; }
.success-banner { background: #f0fdf4; border: 1px solid #bbf7d0; border-left: 4px solid #22c55e; border-radius: 8px; padding: 0.65rem 1rem; margin-bottom: 0.8rem; font-size: 0.81rem; color: #14532d; }

/* ── Sidebar ── */
[data-testid="stSidebar"] { background: #0f172a !important; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] hr { border-color: #1e293b !important; }
[data-testid="stSidebar"] label { font-size: 0.75rem !important; text-transform: uppercase; letter-spacing: 0.06em; }
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] { background: #1e3a5f !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { gap: 0; border-bottom: 2px solid #e2e8f0; background: transparent; }
.stTabs [data-baseweb="tab"] { font-size: 0.78rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; padding: 0.5rem 1.1rem; border-bottom: 2px solid transparent; margin-bottom: -2px; color: #64748b; }
.stTabs [aria-selected="true"] { color: #0f172a; border-bottom-color: #f59e0b; background: transparent; }

/* ── Tables ── */
[data-testid="stDataFrame"] { border-radius: 8px !important; overflow: hidden; }

/* ── File uploader ── */
[data-testid="stFileUploader"] { border: 2px dashed #334155 !important; border-radius: 10px; background: #1e293b !important; }

/* ── Status badge ── */
.status-badge {
    display: inline-block; font-size: 0.63rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.07em; padding: 2px 7px; border-radius: 4px;
}
.status-active   { background: #dcfce7; color: #15803d; }
.status-inactive { background: #fee2e2; color: #b91c1c; }
.status-restart  { background: #fef3c7; color: #b45309; }
.status-spike    { background: #ede9fe; color: #7c3aed; }
.status-outlier  { background: #f3e8ff; color: #9333ea; }

/* ── Divider ── */
.section-divider { border: none; border-top: 1px solid #f1f5f9; margin: 1rem 0; }

/* ── Responsive stacking ── */
@media (max-width: 768px) {
    .analytics-grid { grid-template-columns: 1fr; }
    .block-container { padding: 0 0.75rem 1rem 0.75rem !important; }
    .kpi-value { font-size: 1.2rem; }
}
</style>
"""
st.markdown(PORTAL_CSS, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════

OMC_MAP = {"IOC": "IOCL", "BPC": "BPCL", "HPC": "HPCL"}
PRIMARY_OMCS = ["IOCL", "BPCL", "HPCL"]
OMC_COLORS = {"IOCL": "#f97316", "BPCL": "#3b82f6", "HPCL": "#22c55e"}
INACTIVE_CONSEC_MONTHS = 3
LOW_SALES_THRESHOLD = 5
GROWTH_BASE_THRESHOLD = 10
ZSCORE_CUTOFF = 2.5
IQR_MULTIPLIER = 2.0
CHART_FONT = dict(family="DM Sans", size=12)

# COM-based segment grouping — COM is a column, not a category
COM_SEGMENT_MAP = {
    "D1_D2": ["D1", "D2"],
    "C":     ["C"],
    # E_Others = everything else (E, rural, NaN, etc.)
}
COM_SEGMENT_COLORS = {
    "D1_D2":    "#6366f1",
    "C":        "#f59e0b",
    "E_Others": "#14b8a6",
}


def assign_com_segment(com_val):
    """Map a single COM value to its segment label."""
    if pd.isna(com_val):
        return "E_Others"
    v = str(com_val).strip().upper()
    if v in ["D1", "D2"]:
        return "D1_D2"
    if v == "C":
        return "C"
    return "E_Others"

# ═══════════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════════

def col_to_date(col, prefix="M"):
    s = col[len(prefix):]
    return (2000 + int(s[2:]), int(s[:2]))

def col_to_label(col):
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    s = col[1:]
    return f"{months[int(s[:2])-1]}-{s[2:]}"

def chart_layout(height=300, margin=None, **kw):
    m = margin or dict(t=10, b=10, l=10, r=18)
    base_axis = dict(gridcolor="#f1f5f9", linecolor="#e2e8f0")
    xaxis = {**base_axis, **kw.pop("xaxis", {})}
    yaxis = {**base_axis, **kw.pop("yaxis", {})}
    return dict(
        height=height, margin=m, font=CHART_FONT,
        plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        xaxis=xaxis, yaxis=yaxis,
        **kw
    )

def safe_pct(curr, prev):
    """Growth % with division-by-zero protection."""
    if prev is None or prev == 0 or pd.isna(prev):
        return None
    return (curr - prev) / prev * 100

def growth_class(val):
    if val is None: return "neutral"
    if val > 0:  return "up"
    if val < 0:  return "down"
    return "neutral"

def growth_fmt(val, label=None):
    if label and val is None:
        return f'<span class="analytics-growth neutral">{label}</span>'
    if val is None:
        return '<span class="analytics-growth neutral">N/A</span>'
    arrow = "▲" if val > 0 else ("▼" if val < 0 else "—")
    cls = growth_class(val)
    return f'<span class="analytics-growth {cls}">{arrow} {abs(val):.1f}%</span>'

def delta_html(pct, label="vs prior L3M"):
    if pct is None: return ""
    arrow = "▲" if pct > 0 else "▼"
    cls   = "up" if pct > 0 else "down"
    return f'<div class="kpi-delta {cls}">{arrow} {abs(pct):.1f}% {label}</div>'

# ═══════════════════════════════════════════════════════════════════
# OUTLIER & STATUS DETECTION
# ═══════════════════════════════════════════════════════════════════

def detect_outlet_status(row, ms_cols, hsd_cols, n_recent=6):
    recent_ms  = [float(row.get(c, 0) or 0) for c in ms_cols[-n_recent:]]
    recent_hsd = [float(row.get(c, 0) or 0) for c in hsd_cols[-n_recent:]]
    recent_total = [m + h for m, h in zip(recent_ms, recent_hsd)]

    last3 = recent_total[-INACTIVE_CONSEC_MONTHS:]
    is_inactive = all(v <= LOW_SALES_THRESHOLD for v in last3)

    is_restart = False
    if len(recent_total) >= 4:
        first_half = recent_total[:-2]
        last2 = recent_total[-2:]
        if (all(v <= LOW_SALES_THRESHOLD for v in first_half) and
                any(v > LOW_SALES_THRESHOLD * 5 for v in last2)):
            is_restart = True

    has_spike = False
    if len(recent_total) >= 4 and not is_inactive and not is_restart:
        baseline = np.mean(recent_total[:-1])
        if baseline > LOW_SALES_THRESHOLD and recent_total[-1] > baseline * 3:
            has_spike = True

    if is_inactive:
        return {"is_inactive": True, "is_restart": False, "has_spike": False,
                "status_label": "Inactive", "status_type": "inactive"}
    if is_restart:
        return {"is_inactive": False, "is_restart": True, "has_spike": False,
                "status_label": "Restarted", "status_type": "restart"}
    if has_spike:
        return {"is_inactive": False, "is_restart": False, "has_spike": True,
                "status_label": "Unusual Spike", "status_type": "spike"}
    return {"is_inactive": False, "is_restart": False, "has_spike": False,
            "status_label": "Active", "status_type": "normal"}


def compute_growth(curr, prev, flags):
    if flags["is_inactive"]:  return None, "Inactive", False
    if flags["is_restart"]:   return None, "Restarted", False
    if flags["has_spike"]:    return None, "Spike — verify", False
    if prev is None or prev < GROWTH_BASE_THRESHOLD:
        return None, "Insufficient base", False
    g = safe_pct(curr, prev)
    if g is None: return None, "N/A", False
    if abs(g) > 300:
        return None, f"Extreme ({g:.0f}%)", False
    return g, f"{g:+.1f}%", True


def compute_segment_growth(curr_val, prev_val, flags):
    if flags["is_inactive"] or flags["is_restart"] or flags["has_spike"]:
        return None, False
    g = safe_pct(curr_val, prev_val)
    if g is None: return None, False
    if abs(g) > 300: return None, False
    return g, True


def flag_outliers_series(series):
    if len(series) < 6:
        return pd.Series(False, index=series.index)
    filled = series.fillna(0)
    z_mask = np.abs(stats.zscore(filled)) > ZSCORE_CUTOFF
    q1, q3 = filled.quantile(0.25), filled.quantile(0.75)
    iqr = q3 - q1
    if iqr < 1:
        return pd.Series(False, index=series.index)
    iqr_mask = (filled < q1 - IQR_MULTIPLIER * iqr) | (filled > q3 + IQR_MULTIPLIER * iqr)
    return z_mask & iqr_mask


def flag_outliers(df_active):
    result = pd.Series(False, index=df_active.index, dtype=bool)
    if "Commissioning Date" in df_active.columns:
        comm = pd.to_datetime(df_active["Commissioning Date"], errors="coerce")
        latest = comm.max()
        new_mask = ((latest - comm) < pd.Timedelta(days=548)).reindex(df_active.index, fill_value=False)
    else:
        new_mask = pd.Series(False, index=df_active.index, dtype=bool)

    eligible = df_active[~new_mask]
    group_col = ("Trading_Area_Clean" if "Trading_Area_Clean" in eligible.columns else
                 "Trading Area" if "Trading Area" in eligible.columns else "District")

    handled = set()
    for gval, grp in eligible.groupby(group_col):
        if len(grp) < 6:
            continue
        flags = flag_outliers_series(grp["Total_L3"])
        # Use loc assignment to preserve the original index alignment
        result.loc[flags.index] = flags.values
        handled.update(grp.index.tolist())

    leftover_idx = [i for i in eligible.index if i not in handled]
    if leftover_idx:
        leftover = eligible.loc[leftover_idx]
        for dval, grp in leftover.groupby("District"):
            if len(grp) < 6:
                continue
            flags = flag_outliers_series(grp["Total_L3"])
            result.loc[flags.index] = flags.values
    return result

# ═══════════════════════════════════════════════════════════════════
# DATA LOADING & PROCESSING
# ═══════════════════════════════════════════════════════════════════

@st.cache_data
def load_and_process(file_bytes):
    import io
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    # ── Sort monthly columns chronologically ──────────────────────
    ms_cols  = sorted([c for c in df.columns if re.match(r'^M\d{4}$', c)],
                      key=lambda c: col_to_date(c, "M"))
    hsd_cols = sorted([c for c in df.columns if re.match(r'^H\d{4}$', c)],
                      key=lambda c: col_to_date(c, "H"))

    for c in ms_cols + hsd_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["OMC_std"] = df["OMC"].map(lambda x: OMC_MAP.get(str(x).strip(), str(x).strip()))
    df["District"] = df["District"].str.strip().str.upper()
    if "RO Name" not in df.columns and "RO_Name" in df.columns:
        df.rename(columns={"RO_Name": "RO Name"}, inplace=True)

    # ── COM-based segmentation (values inside COM column) ─────────
    df["COM_Segment"] = df["COM"].apply(assign_com_segment)

    # ── L3M / Prior L3M / Same 3M Last Year windows ──────────────
    ms_last3   = ms_cols[-3:]
    hsd_last3  = hsd_cols[-3:]
    ms_prev3   = ms_cols[-6:-3]   if len(ms_cols) >= 6  else ms_cols[:-3]
    hsd_prev3  = hsd_cols[-6:-3]  if len(hsd_cols) >= 6 else hsd_cols[:-3]
    # Same 3 months one year ago (12 months back from end)
    ms_ly3     = ms_cols[-15:-12] if len(ms_cols) >= 15 else []
    hsd_ly3    = hsd_cols[-15:-12] if len(hsd_cols) >= 15 else []

    # ── L3M Averages (preferred over sums) ───────────────────────
    df["MS_L3"]     = df[ms_last3].sum(axis=1)
    df["HSD_L3"]    = df[hsd_last3].sum(axis=1)
    df["Total_L3"]  = df["MS_L3"] + df["HSD_L3"]
    df["MS_L3_Avg"]   = df["MS_L3"] / 3
    df["HSD_L3_Avg"]  = df["HSD_L3"] / 3
    df["Total_L3_Avg"] = df["Total_L3"] / 3

    # ── Prior L3M Averages ───────────────────────────────────────
    if ms_prev3 and hsd_prev3:
        df["MS_P3"]    = df[ms_prev3].sum(axis=1)
        df["HSD_P3"]   = df[hsd_prev3].sum(axis=1)
        df["Total_P3"] = df["MS_P3"] + df["HSD_P3"]
        df["MS_P3_Avg"]    = df["MS_P3"] / max(len(ms_prev3), 1)
        df["HSD_P3_Avg"]   = df["HSD_P3"] / max(len(hsd_prev3), 1)
        df["Total_P3_Avg"] = df["Total_P3"] / max(len(ms_prev3), 1)
    else:
        df[["MS_P3","HSD_P3","Total_P3","MS_P3_Avg","HSD_P3_Avg","Total_P3_Avg"]] = np.nan

    # ── Same Period Last Year Averages ───────────────────────────
    if ms_ly3 and hsd_ly3:
        df["MS_LY3"]    = df[ms_ly3].sum(axis=1)
        df["HSD_LY3"]   = df[hsd_ly3].sum(axis=1)
        df["Total_LY3"] = df["MS_LY3"] + df["HSD_LY3"]
        df["MS_LY3_Avg"]    = df["MS_LY3"] / max(len(ms_ly3), 1)
        df["HSD_LY3_Avg"]   = df["HSD_LY3"] / max(len(hsd_ly3), 1)
        df["Total_LY3_Avg"] = df["Total_LY3"] / max(len(ms_ly3), 1)
    else:
        df[["MS_LY3","HSD_LY3","Total_LY3","MS_LY3_Avg","HSD_LY3_Avg","Total_LY3_Avg"]] = np.nan

    # ── Status flags ─────────────────────────────────────────────
    statuses = df.apply(lambda r: detect_outlet_status(r, ms_cols, hsd_cols), axis=1)
    for key, col in [("is_inactive","Is_Inactive"), ("is_restart","Is_Restart"),
                     ("has_spike","Has_Spike"), ("status_label","Status_Label"),
                     ("status_type","Status_Type")]:
        df[col] = [s[key] for s in statuses]

    # ── MoM growth (latest single month vs previous single month) ─
    if len(ms_cols) >= 2:
        prev_t = df[ms_cols[-2]] + df[hsd_cols[-2]]
        curr_t = df[ms_cols[-1]] + df[hsd_cols[-1]]
        gv, gl, gm = [], [], []
        for i in df.index:
            flags = {"is_inactive": df.at[i,"Is_Inactive"],
                     "is_restart":  df.at[i,"Is_Restart"],
                     "has_spike":   df.at[i,"Has_Spike"]}
            v, lbl, m = compute_growth(float(curr_t[i]), float(prev_t[i]), flags)
            gv.append(v); gl.append(lbl); gm.append(m)
        df["MoM_Growth"] = gv
        df["MoM_Growth_Label"] = gl
        df["MoM_Growth_Meaningful"] = gm
    else:
        df["MoM_Growth"] = None
        df["MoM_Growth_Label"] = "N/A"
        df["MoM_Growth_Meaningful"] = False

    # ── Segment L3M growth: L3M Avg vs Prior L3M Avg ─────────────
    ms_seg_g, hsd_seg_g, ms_seg_m, hsd_seg_m = [], [], [], []
    ms_yoy_g, hsd_yoy_g = [], []
    for i in df.index:
        flags = {"is_inactive": df.at[i,"Is_Inactive"],
                 "is_restart":  df.at[i,"Is_Restart"],
                 "has_spike":   df.at[i,"Has_Spike"]}
        # L3M Avg vs Prior L3M Avg
        mg, mm = compute_segment_growth(
            df.at[i,"MS_L3_Avg"],
            df.at[i,"MS_P3_Avg"] if not pd.isna(df.at[i,"MS_P3_Avg"]) else None,
            flags)
        hg, hm = compute_segment_growth(
            df.at[i,"HSD_L3_Avg"],
            df.at[i,"HSD_P3_Avg"] if not pd.isna(df.at[i,"HSD_P3_Avg"]) else None,
            flags)
        ms_seg_g.append(mg); ms_seg_m.append(mm)
        hsd_seg_g.append(hg); hsd_seg_m.append(hm)
        # YoY: L3M Avg vs Same Period Last Year Avg
        myoy, _ = compute_segment_growth(
            df.at[i,"MS_L3_Avg"],
            df.at[i,"MS_LY3_Avg"] if not pd.isna(df.at[i,"MS_LY3_Avg"]) else None,
            flags)
        hyoy, _ = compute_segment_growth(
            df.at[i,"HSD_L3_Avg"],
            df.at[i,"HSD_LY3_Avg"] if not pd.isna(df.at[i,"HSD_LY3_Avg"]) else None,
            flags)
        ms_yoy_g.append(myoy); hsd_yoy_g.append(hyoy)

    df["MS_L3_Growth"]        = ms_seg_g
    df["MS_L3_Growth_Valid"]  = ms_seg_m
    df["HSD_L3_Growth"]       = hsd_seg_g
    df["HSD_L3_Growth_Valid"] = hsd_seg_m
    df["MS_YoY_Growth"]       = ms_yoy_g
    df["HSD_YoY_Growth"]      = hsd_yoy_g

    # ── Trading Area normalisation ────────────────────────────────
    if "Trading Area" in df.columns:
        df["Trading_Area_Clean"] = (pd.to_numeric(df["Trading Area"], errors="coerce")
                                    .fillna(0).astype(int).astype(str))
        if "Trading Area Name" in df.columns:
            df["TA_Name"] = df["Trading Area Name"].astype(str).str.strip()
            df["TA_Name"] = df["TA_Name"].replace({"nan": "", "#ERROR!": "", "0": ""})
    else:
        df["Trading_Area_Clean"] = "UNKNOWN"
        df["TA_Name"] = ""

    # ── Outlier detection ─────────────────────────────────────────
    df["Is_Outlier"] = False
    active_mask = ~df["Is_Inactive"]
    if active_mask.sum() > 10:
        outlier_flags = flag_outliers(df.loc[active_mask].copy())
        # Reindex to the active subset's index and cast to bool to avoid index bleed
        aligned = outlier_flags.reindex(df.index, fill_value=False).astype(bool)
        df["Is_Outlier"] = aligned

    # ── Ranks ─────────────────────────────────────────────────────
    df["Overall_Rank"]  = df["Total_L3_Avg"].rank(ascending=False, method="min").astype(int)
    df["District_Rank"] = df.groupby("District")["Total_L3_Avg"].rank(ascending=False, method="min").astype(int)
    df["OMC_Rank"]      = df.groupby("OMC_std")["Total_L3_Avg"].rank(ascending=False, method="min").astype(int)
    df["TA_Rank"]       = df.groupby("Trading_Area_Clean")["Total_L3_Avg"].rank(ascending=False, method="min").astype(int)

    # ── Monthly TA average (for effectivity) ─────────────────────
    # Pre-compute per Trading Area per month average volume (industry proxy)
    # This is attached to df for later use — returns raw columns list too
    return df, ms_cols, hsd_cols, ms_last3, hsd_last3, ms_prev3, hsd_prev3, ms_ly3, hsd_ly3

# ═══════════════════════════════════════════════════════════════════
# EFFECTIVITY HELPERS
# ═══════════════════════════════════════════════════════════════════

def compute_omc_effectivity(df_full, ms_cols, hsd_cols, n_months=12):
    """
    OMC Effectivity = OMC monthly volume / Avg industry monthly volume
    Returns a long-form DataFrame: Month, OMC, MS_Vol, HSD_Vol,
    Ind_MS_Avg, Ind_HSD_Avg, MS_Eff, HSD_Eff
    """
    records = []
    mc_window = ms_cols[-n_months:]
    hc_window = hsd_cols[-n_months:]
    # Industry monthly averages across ALL outlets
    for mc, hc in zip(mc_window, hc_window):
        month_lbl = col_to_label(mc)
        ind_ms_avg  = df_full[mc].mean()
        ind_hsd_avg = df_full[hc].mean()
        omc_grp = df_full.groupby("OMC_std").agg(
            MS_Vol=(mc, "sum"), HSD_Vol=(hc, "sum"), Outlets=("RO Name","count")
        ).reset_index()
        for _, row in omc_grp.iterrows():
            ms_eff  = safe_pct(row["MS_Vol"]  / max(row["Outlets"],1), ind_ms_avg)
            hsd_eff = safe_pct(row["HSD_Vol"] / max(row["Outlets"],1), ind_hsd_avg)
            # Effectivity = omc_avg_per_outlet / industry_avg_per_outlet
            omc_ms_avg  = row["MS_Vol"]  / max(row["Outlets"],1)
            omc_hsd_avg = row["HSD_Vol"] / max(row["Outlets"],1)
            ms_eff_ratio  = (omc_ms_avg  / ind_ms_avg * 100)  if ind_ms_avg  > 0 else None
            hsd_eff_ratio = (omc_hsd_avg / ind_hsd_avg * 100) if ind_hsd_avg > 0 else None
            records.append({
                "Month": month_lbl, "OMC": row["OMC_std"],
                "MS_Vol": row["MS_Vol"], "HSD_Vol": row["HSD_Vol"],
                "Outlets": row["Outlets"],
                "Ind_MS_Avg": ind_ms_avg, "Ind_HSD_Avg": ind_hsd_avg,
                "MS_Eff": ms_eff_ratio, "HSD_Eff": hsd_eff_ratio,
            })
    return pd.DataFrame(records)


def compute_ro_effectivity(outlet_row, df_full, ms_cols, hsd_cols, n_months=12):
    """
    RO Effectivity = RO monthly volume / Avg Trading Area monthly volume
    Separately for MS and HSD.
    Returns monthly DataFrame for the outlet.
    """
    ta_key = outlet_row.get("Trading_Area_Clean", "")
    ta_peers = df_full[df_full["Trading_Area_Clean"] == ta_key] if ta_key else pd.DataFrame()
    records = []
    for mc, hc in zip(ms_cols[-n_months:], hsd_cols[-n_months:]):
        month_lbl = col_to_label(mc)
        ro_ms  = float(outlet_row.get(mc, 0) or 0)
        ro_hsd = float(outlet_row.get(hc, 0) or 0)
        if not ta_peers.empty:
            ta_ms_avg  = ta_peers[mc].mean()
            ta_hsd_avg = ta_peers[hc].mean()
            ta_ms_tot  = ta_peers[mc].sum()
            ta_hsd_tot = ta_peers[hc].sum()
        else:
            ta_ms_avg = ta_hsd_avg = ta_ms_tot = ta_hsd_tot = np.nan

        ms_eff  = (ro_ms  / ta_ms_avg  * 100) if (ta_ms_avg  and ta_ms_avg  > 0) else None
        hsd_eff = (ro_hsd / ta_hsd_avg * 100) if (ta_hsd_avg and ta_hsd_avg > 0) else None
        ms_share  = (ro_ms  / ta_ms_tot  * 100) if (ta_ms_tot  and ta_ms_tot  > 0) else None
        hsd_share = (ro_hsd / ta_hsd_tot * 100) if (ta_hsd_tot and ta_hsd_tot > 0) else None

        records.append({
            "Month": month_lbl,
            "RO_MS": ro_ms, "RO_HSD": ro_hsd,
            "TA_MS_Avg": ta_ms_avg, "TA_HSD_Avg": ta_hsd_avg,
            "MS_Eff": ms_eff, "HSD_Eff": hsd_eff,
            "MS_Share": ms_share, "HSD_Share": hsd_share,
        })
    return pd.DataFrame(records)

# ═══════════════════════════════════════════════════════════════════
# PREDICTION
# ═══════════════════════════════════════════════════════════════════

def predict_outlet(row, ms_cols, hsd_cols, n=6):
    def _predict(cols):
        vals = [float(row.get(c, 0) or 0) for c in cols[-n:]]
        start = next((i for i, v in enumerate(vals) if v > LOW_SALES_THRESHOLD), len(vals) - 2)
        vals = vals[start:]
        if len(vals) < 2:
            return (vals[-1] if vals else 0), 0
        alpha, sm = 0.4, [vals[0]]
        for v in vals[1:]:
            sm.append(alpha * v + (1 - alpha) * sm[-1])
        X   = np.arange(len(sm)).reshape(-1, 1)
        mdl = LinearRegression().fit(X, np.array(sm))
        pred = max(float(mdl.predict([[len(sm)]])[0]), 0)
        pred = min(pred, max(sm) * 1.5)
        g = ((pred - sm[-1]) / sm[-1] * 100) if sm[-1] > GROWTH_BASE_THRESHOLD else 0
        return pred, g

    ms_pred,  ms_g  = _predict(ms_cols)
    hsd_pred, hsd_g = _predict(hsd_cols)
    total_pred = ms_pred + hsd_pred
    base = (row["Total_L3"] / 3) if row["Total_L3"] > GROWTH_BASE_THRESHOLD * 3 else None
    total_g = ((total_pred - base) / base * 100) if base else 0
    trend = "Increase" if total_g > 5 else ("Decline" if total_g < -5 else "Stable")
    return {"ms_pred": ms_pred, "hsd_pred": hsd_pred, "total_pred": total_pred,
            "ms_growth": ms_g, "hsd_growth": hsd_g, "total_growth": total_g, "trend": trend}

# ═══════════════════════════════════════════════════════════════════
# MS / HSD ANALYTICS BLOCK RENDERER (enhanced with prior & YoY)
# ═══════════════════════════════════════════════════════════════════

def render_ms_hsd_analytics(ms_avg, hsd_avg,
                             ms_prior_avg=None, hsd_prior_avg=None,
                             ms_ly_avg=None, hsd_ly_avg=None,
                             ms_growth=None, hsd_growth=None,
                             ms_growth_valid=False, hsd_growth_valid=False,
                             ms_yoy=None, hsd_yoy=None):
    """Renders a two-column MS/HSD block showing L3M avg, vs prior, vs same period LY."""

    def _growth_html(g, valid):
        if not valid or g is None:
            return '<span class="analytics-growth neutral">N/A</span>'
        arrow = "▲" if g > 0 else ("▼" if g < 0 else "—")
        cls   = "up" if g > 0 else ("down" if g < 0 else "neutral")
        return f'<span class="analytics-growth {cls}">{arrow} {abs(g):.1f}%</span>'

    def _val(v):
        return f"{v:,.1f} KL" if v is not None and not pd.isna(v) else "N/A"

    ms_g_html   = _growth_html(ms_growth,  ms_growth_valid)
    hsd_g_html  = _growth_html(hsd_growth, hsd_growth_valid)
    ms_yoy_html = _growth_html(ms_yoy,  ms_yoy is not None)
    hsd_yoy_html = _growth_html(hsd_yoy, hsd_yoy is not None)

    st.markdown(f"""
    <div class="analytics-grid">
      <div class="analytics-block ms-block">
        <div class="analytics-type ms">Motor Spirit (MS)</div>
        <div class="analytics-row">
          <span class="analytics-metric">L3M Avg/Month</span>
          <span class="analytics-value">{_val(ms_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">Prior L3M Avg</span>
          <span class="analytics-value">{_val(ms_prior_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">vs Prior L3M</span>
          {ms_g_html}
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">Same Period LY Avg</span>
          <span class="analytics-value">{_val(ms_ly_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">YoY Growth</span>
          {ms_yoy_html}
        </div>
      </div>
      <div class="analytics-block hsd-block">
        <div class="analytics-type hsd">High Speed Diesel (HSD)</div>
        <div class="analytics-row">
          <span class="analytics-metric">L3M Avg/Month</span>
          <span class="analytics-value">{_val(hsd_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">Prior L3M Avg</span>
          <span class="analytics-value">{_val(hsd_prior_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">vs Prior L3M</span>
          {hsd_g_html}
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">Same Period LY Avg</span>
          <span class="analytics-value">{_val(hsd_ly_avg)}</span>
        </div>
        <div class="analytics-row">
          <span class="analytics-metric">YoY Growth</span>
          {hsd_yoy_html}
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div style="padding:1.2rem 0 0.4rem 0;">
      <div style="font-size:1.05rem;font-weight:700;color:#f1f5f9;letter-spacing:-0.01em;">PetroAnalytics</div>
      <div style="font-size:0.65rem;color:#475569;text-transform:uppercase;letter-spacing:0.1em;margin-top:3px;">Outlet Intelligence Portal</div>
    </div>""", unsafe_allow_html=True)
    st.markdown("---")
    uploaded = st.file_uploader("Upload Dataset (.xlsx)", type=["xlsx"])
    st.markdown("---")

if uploaded is None:
    st.markdown("""
    <div class="portal-header">
      <div class="portal-title">PetroAnalytics <span class="portal-badge">Portal</span></div>
      <div class="portal-subtitle">Petroleum Retail Outlet Intelligence — IOCL · BPCL · HPCL</div>
    </div>
    <div class="info-banner">
      <b>Getting Started:</b> Upload an Excel (.xlsx) dataset using the sidebar.
      MS columns should follow the pattern <code>Mmmyy</code> (e.g. M0124) and HSD columns <code>Hmmyy</code>.
      All columns are auto-detected — no configuration needed.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

raw_bytes = uploaded.read()
df_raw, ms_cols, hsd_cols, ms_last3, hsd_last3, ms_prev3, hsd_prev3, ms_ly3, hsd_ly3 = load_and_process(raw_bytes)

# ── Sidebar filters ──
with st.sidebar:
    st.markdown('<div style="font-size:0.65rem;color:#475569;text-transform:uppercase;letter-spacing:0.1em;font-weight:600;margin-bottom:6px;">Filters</div>', unsafe_allow_html=True)
    selected_omcs      = st.multiselect("OMC", sorted(df_raw["OMC_std"].unique()), default=sorted(df_raw["OMC_std"].unique()))
    selected_districts = st.multiselect("District", sorted(df_raw["District"].unique()), default=sorted(df_raw["District"].unique()))
    selected_segments  = st.multiselect("COM Segment", ["D1_D2","C","E_Others"], default=["D1_D2","C","E_Others"])
    st.markdown("---")
    st.markdown('<div style="font-size:0.65rem;color:#475569;text-transform:uppercase;letter-spacing:0.1em;font-weight:600;margin-bottom:6px;">Data Controls</div>', unsafe_allow_html=True)
    exclude_inactive = st.toggle("Exclude Inactive Outlets", value=True)
    exclude_outliers = st.toggle("Exclude Outliers from Charts", value=False)
    show_norm_only   = st.toggle("Meaningful Growth Only", value=True)
    st.markdown("---")

    n_inactive = int(df_raw["Is_Inactive"].sum())
    n_restart  = int(df_raw["Is_Restart"].sum())
    n_outlier  = int(df_raw["Is_Outlier"].sum())
    n_spike    = int(df_raw["Has_Spike"].sum())

    st.markdown(f"""
    <div style="font-size:0.7rem;color:#94a3b8;line-height:2.0;">
      <div>Total outlets: <b style="color:#f1f5f9;">{len(df_raw)}</b></div>
      <div>Inactive: <b style="color:#f87171;">{n_inactive}</b></div>
      <div>Restarted: <b style="color:#fb923c;">{n_restart}</b></div>
      <div>Outliers: <b style="color:#c084fc;">{n_outlier}</b></div>
      <div>Spikes: <b style="color:#fbbf24;">{n_spike}</b></div>
      <div style="margin-top:6px;">Latest month: <b style="color:#f1f5f9;">{col_to_label(ms_cols[-1])}</b></div>
      <div>L3M window: <b style="color:#f1f5f9;">{col_to_label(ms_last3[0])}–{col_to_label(ms_last3[-1])}</b></div>
    </div>""", unsafe_allow_html=True)

# ── Apply filters ──
df = df_raw[
    df_raw["OMC_std"].isin(selected_omcs) &
    df_raw["District"].isin(selected_districts) &
    df_raw["COM_Segment"].isin(selected_segments)
].copy()
if exclude_inactive: df = df[~df["Is_Inactive"]]
if exclude_outliers: df = df[~df["Is_Outlier"]]

# ═══════════════════════════════════════════════════════════════════
# PORTAL HEADER
# ═══════════════════════════════════════════════════════════════════

latest_label = col_to_label(ms_cols[-1])
prev_label   = col_to_label(ms_prev3[0]) if ms_prev3 else "—"
ly_label     = col_to_label(ms_ly3[0])   if ms_ly3   else "—"

st.markdown(f"""
<div class="portal-header">
  <div class="portal-title">PetroAnalytics <span class="portal-badge">Live</span></div>
  <div class="portal-subtitle">Petroleum Retail Intelligence Portal — IOCL · BPCL · HPCL &nbsp;·&nbsp; Data through {latest_label}</div>
  <div class="portal-meta">
    <div class="portal-meta-item">Outlets loaded: <b>{len(df_raw)}</b></div>
    <div class="portal-meta-item">Active in view: <b>{len(df)}</b></div>
    <div class="portal-meta-item">L3M window: <b>{col_to_label(ms_last3[0])} — {latest_label}</b></div>
    <div class="portal-meta-item">Prior L3M: <b>{col_to_label(ms_prev3[0]) if ms_prev3 else '—'} — {col_to_label(ms_prev3[-1]) if ms_prev3 else '—'}</b></div>
    <div class="portal-meta-item">Same 3M LY: <b>{col_to_label(ms_ly3[0]) if ms_ly3 else '—'} — {col_to_label(ms_ly3[-1]) if ms_ly3 else '—'}</b></div>
  </div>
</div>""", unsafe_allow_html=True)

alerts = []
if n_inactive > 0: alerts.append(f"{n_inactive} outlet(s) flagged inactive (zero sales 3+ months). Excluded from rankings by default.")
if n_restart  > 0: alerts.append(f"{n_restart} outlet(s) show restart patterns — growth % suppressed.")
if n_outlier  > 0: alerts.append(f"{n_outlier} statistical outlier(s) detected (Z-score + IQR). Use sidebar toggle to exclude.")
if n_spike    > 0: alerts.append(f"{n_spike} outlet(s) show unusual sales spikes. Verify source data before acting.")
if alerts:
    with st.expander(f"Data Quality Alerts ({len(alerts)})", expanded=False):
        for msg in alerts:
            st.markdown(f'<div class="warn-banner">{msg}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# AGGREGATE KPIs (averages, not sums)
# ═══════════════════════════════════════════════════════════════════

total_ms_avg   = df["MS_L3_Avg"].sum()     # sum of per-outlet monthly avgs = aggregate monthly avg
total_hsd_avg  = df["HSD_L3_Avg"].sum()
total_vol_avg  = df["Total_L3_Avg"].sum()
total_ms_p_avg = df["MS_P3_Avg"].sum()
total_hsd_p_avg= df["HSD_P3_Avg"].sum()
total_p_avg    = total_ms_p_avg + total_hsd_p_avg
total_ms_ly_avg= df["MS_LY3_Avg"].sum()
total_hsd_ly_avg=df["HSD_LY3_Avg"].sum()
total_ly_avg   = total_ms_ly_avg + total_hsd_ly_avg

ms_delta_pct   = safe_pct(total_ms_avg,  total_ms_p_avg)
hsd_delta_pct  = safe_pct(total_hsd_avg, total_hsd_p_avg)
tot_delta_pct  = safe_pct(total_vol_avg, total_p_avg)
ms_yoy_pct     = safe_pct(total_ms_avg,  total_ms_ly_avg)
hsd_yoy_pct    = safe_pct(total_hsd_avg, total_hsd_ly_avg)
tot_yoy_pct    = safe_pct(total_vol_avg, total_ly_avg)

meaningful_growth = df[df["MoM_Growth_Meaningful"] == True]["MoM_Growth"]
avg_g_str = f"{meaningful_growth.mean():+.1f}%" if not meaningful_growth.empty else "N/A"

prim_df   = df[df["OMC_std"].isin(PRIMARY_OMCS)]
omc_leader = prim_df.groupby("OMC_std")["Total_L3_Avg"].sum().idxmax() if not prim_df.empty else "—"

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    st.markdown(f"""<div class="kpi-card blue">
    <div class="kpi-label">Total Outlets</div>
    <div class="kpi-value">{len(df)}</div>
    <div class="kpi-sub">filtered view</div></div>""", unsafe_allow_html=True)
with c2:
    inactive_ct = len(df[df["Is_Inactive"]]) if not exclude_inactive else 0
    st.markdown(f"""<div class="kpi-card green">
    <div class="kpi-label">Active Outlets</div>
    <div class="kpi-value">{len(df)}</div>
    <div class="kpi-sub">{inactive_ct} inactive in view</div></div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="kpi-card amber">
    <div class="kpi-label">MS Avg / Month</div>
    <div class="kpi-value">{total_ms_avg/1000:.1f}K</div>
    {delta_html(ms_delta_pct)}
    <div class="kpi-sub">KL · L3M avg · YoY {f'{ms_yoy_pct:+.1f}%' if ms_yoy_pct else 'N/A'}</div></div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="kpi-card purple">
    <div class="kpi-label">HSD Avg / Month</div>
    <div class="kpi-value">{total_hsd_avg/1000:.1f}K</div>
    {delta_html(hsd_delta_pct)}
    <div class="kpi-sub">KL · L3M avg · YoY {f'{hsd_yoy_pct:+.1f}%' if hsd_yoy_pct else 'N/A'}</div></div>""", unsafe_allow_html=True)
with c5:
    st.markdown(f"""<div class="kpi-card teal">
    <div class="kpi-label">Total Avg / Month</div>
    <div class="kpi-value">{total_vol_avg/1000:.1f}K</div>
    {delta_html(tot_delta_pct)}
    <div class="kpi-sub">KL · MS + HSD · YoY {f'{tot_yoy_pct:+.1f}%' if tot_yoy_pct else 'N/A'}</div></div>""", unsafe_allow_html=True)
with c6:
    st.markdown(f"""<div class="kpi-card indigo">
    <div class="kpi-label">Avg MoM Growth</div>
    <div class="kpi-value" style="font-size:1.3rem;">{avg_g_str}</div>
    <div class="kpi-sub">meaningful outlets only</div></div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════

tabs = st.tabs(["Overview", "MS / HSD Analytics", "District Analysis",
                "Search & Outlet", "AI Predictions", "Data Quality", "Insights"])

# ═══════════════════════════════════════════════════════════════════
# TAB 1: OVERVIEW
# ═══════════════════════════════════════════════════════════════════
with tabs[0]:

    # ── A. KPI summary using L3M Averages (not sums) ─────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Portfolio Averages — L3M <span class="section-subtitle">avg monthly volume · vs prior L3M · vs same period last year</span></div>', unsafe_allow_html=True)

    ov1, ov2, ov3 = st.columns(3)
    with ov1:
        st.markdown(f"""
        <div class="mini-kpi">
          <div class="mini-kpi-label">MS L3M Avg/Month</div>
          <div class="mini-kpi-value">{total_ms_avg:,.1f} KL</div>
          <div class="mini-kpi-sub">Prior: {total_ms_p_avg:,.1f} KL &nbsp;|&nbsp; {f'▲ {ms_delta_pct:.1f}%' if ms_delta_pct and ms_delta_pct>0 else f'▼ {abs(ms_delta_pct):.1f}%' if ms_delta_pct else 'N/A'}</div>
          <div class="mini-kpi-sub">Same Period LY: {total_ms_ly_avg:,.1f} KL &nbsp;|&nbsp; YoY: {f'{ms_yoy_pct:+.1f}%' if ms_yoy_pct else 'N/A'}</div>
        </div>""", unsafe_allow_html=True)
    with ov2:
        st.markdown(f"""
        <div class="mini-kpi">
          <div class="mini-kpi-label">HSD L3M Avg/Month</div>
          <div class="mini-kpi-value">{total_hsd_avg:,.1f} KL</div>
          <div class="mini-kpi-sub">Prior: {total_hsd_p_avg:,.1f} KL &nbsp;|&nbsp; {f'▲ {hsd_delta_pct:.1f}%' if hsd_delta_pct and hsd_delta_pct>0 else f'▼ {abs(hsd_delta_pct):.1f}%' if hsd_delta_pct else 'N/A'}</div>
          <div class="mini-kpi-sub">Same Period LY: {total_hsd_ly_avg:,.1f} KL &nbsp;|&nbsp; YoY: {f'{hsd_yoy_pct:+.1f}%' if hsd_yoy_pct else 'N/A'}</div>
        </div>""", unsafe_allow_html=True)
    with ov3:
        st.markdown(f"""
        <div class="mini-kpi">
          <div class="mini-kpi-label">Combined L3M Avg/Month</div>
          <div class="mini-kpi-value">{total_vol_avg:,.1f} KL</div>
          <div class="mini-kpi-sub">Prior: {total_p_avg:,.1f} KL &nbsp;|&nbsp; {f'▲ {tot_delta_pct:.1f}%' if tot_delta_pct and tot_delta_pct>0 else f'▼ {abs(tot_delta_pct):.1f}%' if tot_delta_pct else 'N/A'}</div>
          <div class="mini-kpi-sub">Same Period LY: {total_ly_avg:,.1f} KL &nbsp;|&nbsp; YoY: {f'{tot_yoy_pct:+.1f}%' if tot_yoy_pct else 'N/A'}</div>
        </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── B. COM Segment breakdown (using COM column values) ────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">COM Segment Performance <span class="section-subtitle">D1_D2 · C · E_Others — based on COM column values</span></div>', unsafe_allow_html=True)

    seg_agg = df.groupby("COM_Segment").agg(
        Outlets=("RO Name","count"),
        MS_L3_Avg=("MS_L3_Avg","sum"),
        HSD_L3_Avg=("HSD_L3_Avg","sum"),
        MS_P3_Avg=("MS_P3_Avg","sum"),
        HSD_P3_Avg=("HSD_P3_Avg","sum"),
        MS_LY3_Avg=("MS_LY3_Avg","sum"),
        HSD_LY3_Avg=("HSD_LY3_Avg","sum"),
    ).reset_index()
    seg_agg["Total_Avg"]   = seg_agg["MS_L3_Avg"] + seg_agg["HSD_L3_Avg"]
    seg_agg["Total_P_Avg"] = seg_agg["MS_P3_Avg"]  + seg_agg["HSD_P3_Avg"]
    seg_agg["Total_LY_Avg"]= seg_agg["MS_LY3_Avg"] + seg_agg["HSD_LY3_Avg"]
    seg_agg["vs_Prior"]    = seg_agg.apply(lambda r: safe_pct(r["Total_Avg"], r["Total_P_Avg"]),  axis=1)
    seg_agg["YoY"]         = seg_agg.apply(lambda r: safe_pct(r["Total_Avg"], r["Total_LY_Avg"]), axis=1)

    sc1, sc2, sc3 = st.columns(3)
    seg_order = ["D1_D2","C","E_Others"]
    seg_cols  = [sc1, sc2, sc3]
    seg_css   = ["indigo","amber","teal"]
    for col_w, seg, css in zip(seg_cols, seg_order, seg_css):
        row = seg_agg[seg_agg["COM_Segment"] == seg]
        if row.empty:
            with col_w:
                st.markdown(f'<div class="kpi-card {css}"><div class="kpi-label">{seg}</div><div class="kpi-value">—</div><div class="kpi-sub">No data</div></div>', unsafe_allow_html=True)
            continue
        r = row.iloc[0]
        vp  = f"{r['vs_Prior']:+.1f}%" if r["vs_Prior"] is not None else "N/A"
        vyoy= f"{r['YoY']:+.1f}%"      if r["YoY"]      is not None else "N/A"
        with col_w:
            st.markdown(f"""<div class="kpi-card {css}">
            <div class="kpi-label">{seg} ({int(r['Outlets'])} outlets)</div>
            <div class="kpi-value">{r['Total_Avg']:,.1f} KL</div>
            <div class="kpi-sub">avg/month · MS {r['MS_L3_Avg']:,.1f} + HSD {r['HSD_L3_Avg']:,.1f}</div>
            <div class="kpi-delta {'up' if r['vs_Prior'] and r['vs_Prior']>0 else 'down' if r['vs_Prior'] and r['vs_Prior']<0 else 'neutral'}">vs Prior: {vp} &nbsp;|&nbsp; YoY: {vyoy}</div>
            </div>""", unsafe_allow_html=True)

    # Segment bar chart
    seg_plot = seg_agg[seg_agg["COM_Segment"].isin(seg_order)].set_index("COM_Segment").reindex(seg_order).reset_index()
    fig_seg = go.Figure()
    fig_seg.add_trace(go.Bar(x=seg_plot["COM_Segment"], y=seg_plot["MS_L3_Avg"],
                              name="MS", marker_color="#f97316",
                              text=seg_plot["MS_L3_Avg"].map("{:,.0f}".format), textposition="outside"))
    fig_seg.add_trace(go.Bar(x=seg_plot["COM_Segment"], y=seg_plot["HSD_L3_Avg"],
                              name="HSD", marker_color="#3b82f6",
                              text=seg_plot["HSD_L3_Avg"].map("{:,.0f}".format), textposition="outside"))
    fig_seg.update_layout(**chart_layout(280, barmode="group",
                                          legend=dict(orientation="h", y=1.08),
                                          yaxis_title="Avg Monthly Volume (KL)"))
    st.plotly_chart(fig_seg, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    r1c1, r1c2 = st.columns([3, 2], gap="medium")

    with r1c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Top 5 Outlets <span class="section-subtitle">L3M avg monthly volume · outliers in purple</span></div>', unsafe_allow_html=True)
        is_filtered = (set(selected_omcs) != set(df_raw["OMC_std"].unique()) or
                       set(selected_districts) != set(df_raw["District"].unique()) or
                       exclude_inactive or exclude_outliers)
        t5_src = df if is_filtered else df_raw[~df_raw["Is_Inactive"]]
        top5   = t5_src.nlargest(5, "Total_L3_Avg").copy()

        fig = go.Figure()
        for _, row in top5.iterrows():
            clr = "#c084fc" if row["Is_Outlier"] else OMC_COLORS.get(row["OMC_std"], "#94a3b8")
            fig.add_trace(go.Bar(
                x=[row["Total_L3_Avg"]], y=[row["RO Name"]], orientation="h",
                name=row["OMC_std"], marker_color=clr, showlegend=False,
                text=f"{row['Total_L3_Avg']:,.1f}", textposition="outside",
            ))
        fig.update_layout(**chart_layout(280, yaxis=dict(autorange="reversed"), xaxis_title="Avg Monthly Volume (KL)"))
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with r1c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">OMC Market Share <span class="section-subtitle">by avg monthly volume</span></div>', unsafe_allow_html=True)
        if not prim_df.empty:
            share = prim_df.groupby("OMC_std")["Total_L3_Avg"].sum().reset_index()
            fig_p = px.pie(share, values="Total_L3_Avg", names="OMC_std", hole=0.45,
                           color="OMC_std", color_discrete_map=OMC_COLORS)
            fig_p.update_traces(textposition="inside", textinfo="percent+label",
                                textfont=dict(size=12, family="DM Sans"))
            fig_p.update_layout(showlegend=False, height=280, margin=dict(t=5,b=5,l=5,r=5),
                                paper_bgcolor="#ffffff",
                                annotations=[dict(text=f"<b>{omc_leader}</b><br>leads", x=0.5, y=0.5,
                                                  font_size=13, showarrow=False,
                                                  font_color="#0f172a", font_family="DM Sans")])
            st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.info("No primary OMC data in filter.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Monthly trend (L12M) — using averages
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Monthly Volume Trend <span class="section-subtitle">last 12 months · avg monthly volume MS and HSD</span></div>', unsafe_allow_html=True)
    trend_data = [{"Month": col_to_label(mc), "MS Avg": df[mc].mean(), "HSD Avg": df[hc].mean()}
                  for mc, hc in zip(ms_cols[-12:], hsd_cols[-12:])]
    tdf = pd.DataFrame(trend_data)
    fig_t = go.Figure()
    fig_t.add_trace(go.Scatter(x=tdf["Month"], y=tdf["MS Avg"], mode="lines+markers", name="MS Avg",
                               line=dict(color="#f97316", width=2.5), marker=dict(size=6),
                               fill="tozeroy", fillcolor="rgba(249,115,22,0.07)"))
    fig_t.add_trace(go.Scatter(x=tdf["Month"], y=tdf["HSD Avg"], mode="lines+markers", name="HSD Avg",
                               line=dict(color="#3b82f6", width=2.5), marker=dict(size=6),
                               fill="tozeroy", fillcolor="rgba(59,130,246,0.07)"))
    fig_t.update_layout(**chart_layout(290, yaxis=dict(gridcolor="#f1f5f9", title="Avg Volume per Outlet (KL)"),
                                       legend=dict(orientation="h", y=1.08, x=0)))
    st.plotly_chart(fig_t, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── C. Growth Metrics Table ───────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Growth Metrics by COM Segment <span class="section-subtitle">L3M Avg vs Prior · vs Same Period LY · YoY</span></div>', unsafe_allow_html=True)
    growth_rows = []
    for seg in seg_order:
        r = seg_agg[seg_agg["COM_Segment"] == seg]
        if r.empty: continue
        r = r.iloc[0]
        growth_rows.append({
            "Segment": seg,
            "Outlets": int(r["Outlets"]),
            "MS Avg/Mo (KL)": r["MS_L3_Avg"],
            "HSD Avg/Mo (KL)": r["HSD_L3_Avg"],
            "Total Avg/Mo (KL)": r["Total_Avg"],
            "vs Prior L3M": safe_pct(r["Total_Avg"], r["Total_P_Avg"]),
            "vs Same Period LY": safe_pct(r["Total_Avg"], r["Total_LY_Avg"]),
        })
    g_tbl = pd.DataFrame(growth_rows)

    def color_pct(val):
        if pd.isna(val): return "color:#94a3b8"
        if val > 2:  return "color:#16a34a;font-weight:600"
        if val < -2: return "color:#dc2626;font-weight:600"
        return "color:#d97706"

    if not g_tbl.empty:
        st.dataframe(
            g_tbl.style
            .format({"MS Avg/Mo (KL)":"{:,.1f}", "HSD Avg/Mo (KL)":"{:,.1f}",
                     "Total Avg/Mo (KL)":"{:,.1f}",
                     "vs Prior L3M":"{:+.1f}%", "vs Same Period LY":"{:+.1f}%"}, na_rep="N/A")
            .applymap(color_pct, subset=["vs Prior L3M","vs Same Period LY"]),
            use_container_width=True, hide_index=True,
        )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── D. OMC Effectivity ────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">OMC Effectivity <span class="section-subtitle">OMC avg per outlet / Industry avg per outlet · monthly trend</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="info-banner">Effectivity = OMC avg volume per outlet ÷ Industry avg volume per outlet × 100. Values above 100 indicate above-industry performance.</div>', unsafe_allow_html=True)

    eff_df = compute_omc_effectivity(df_raw, ms_cols, hsd_cols, n_months=12)

    ec1, ec2 = st.columns(2, gap="medium")
    with ec1:
        st.markdown("**MS Effectivity Trend**")
        fig_ems = go.Figure()
        for omc in sorted(eff_df["OMC"].unique()):
            odf = eff_df[eff_df["OMC"] == omc]
            color = OMC_COLORS.get(omc, "#94a3b8")
            fig_ems.add_trace(go.Scatter(x=odf["Month"], y=odf["MS_Eff"],
                                          mode="lines+markers", name=omc,
                                          line=dict(color=color, width=2), marker=dict(size=5)))
        fig_ems.add_hline(y=100, line_dash="dot", line_color="#94a3b8",
                           annotation_text="Industry baseline (100)", annotation_position="bottom right")
        fig_ems.update_layout(**chart_layout(280, yaxis_title="MS Effectivity (%)",
                                              legend=dict(orientation="h", y=1.08),
                                              xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_ems, use_container_width=True)

    with ec2:
        st.markdown("**HSD Effectivity Trend**")
        fig_ehsd = go.Figure()
        for omc in sorted(eff_df["OMC"].unique()):
            odf = eff_df[eff_df["OMC"] == omc]
            color = OMC_COLORS.get(omc, "#94a3b8")
            fig_ehsd.add_trace(go.Scatter(x=odf["Month"], y=odf["HSD_Eff"],
                                           mode="lines+markers", name=omc,
                                           line=dict(color=color, width=2), marker=dict(size=5)))
        fig_ehsd.add_hline(y=100, line_dash="dot", line_color="#94a3b8",
                            annotation_text="Industry baseline (100)", annotation_position="bottom right")
        fig_ehsd.update_layout(**chart_layout(280, yaxis_title="HSD Effectivity (%)",
                                               legend=dict(orientation="h", y=1.08),
                                               xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_ehsd, use_container_width=True)

    # Latest month effectivity table
    latest_eff = eff_df[eff_df["Month"] == col_to_label(ms_cols[-1])].copy()
    if not latest_eff.empty:
        st.markdown(f"**Latest Month Effectivity — {col_to_label(ms_cols[-1])}**")
        latest_eff["MS Eff (%)"]  = latest_eff["MS_Eff"].map(lambda x: f"{x:.1f}%" if x else "N/A")
        latest_eff["HSD Eff (%)"] = latest_eff["HSD_Eff"].map(lambda x: f"{x:.1f}%" if x else "N/A")
        st.dataframe(latest_eff[["OMC","Outlets","MS_Vol","HSD_Vol","MS Eff (%)","HSD Eff (%)"]]
                     .rename(columns={"MS_Vol":"MS Vol (KL)","HSD_Vol":"HSD Vol (KL)"})
                     .style.format({"MS Vol (KL)":"{:,.0f}","HSD Vol (KL)":"{:,.0f}"}),
                     use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("All Outlets Ranked Table", expanded=False):
        rank_tbl = df[["RO Name","OMC_std","District","COM_Segment","MS_L3_Avg","HSD_L3_Avg",
                        "Total_L3_Avg","MoM_Growth_Label","Status_Label","Is_Outlier","Overall_Rank"]].copy()
        rank_tbl = rank_tbl.sort_values("Total_L3_Avg", ascending=False).reset_index(drop=True)
        rank_tbl.index += 1
        rank_tbl["Outlier"] = rank_tbl["Is_Outlier"].map(lambda x: "Yes" if x else "")
        rank_tbl.drop(columns=["Is_Outlier"], inplace=True)
        rank_tbl.rename(columns={
            "OMC_std":"OMC","COM_Segment":"Segment",
            "MS_L3_Avg":"MS Avg/Mo","HSD_L3_Avg":"HSD Avg/Mo",
            "Total_L3_Avg":"Total Avg/Mo","MoM_Growth_Label":"MoM Growth",
            "Status_Label":"Status","Overall_Rank":"Rank"}, inplace=True)
        st.dataframe(rank_tbl.style.format({"MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}","Total Avg/Mo":"{:,.1f}"}),
                     use_container_width=True, height=380)

# ═══════════════════════════════════════════════════════════════════
# TAB 2: MS / HSD ANALYTICS
# ═══════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Aggregate MS & HSD Performance <span class="section-subtitle">L3M avg · vs prior 3M · vs same period last year</span></div>', unsafe_allow_html=True)

    render_ms_hsd_analytics(
        ms_avg=total_ms_avg, hsd_avg=total_hsd_avg,
        ms_prior_avg=total_ms_p_avg, hsd_prior_avg=total_hsd_p_avg,
        ms_ly_avg=total_ms_ly_avg,   hsd_ly_avg=total_hsd_ly_avg,
        ms_growth=ms_delta_pct, hsd_growth=hsd_delta_pct,
        ms_growth_valid=(ms_delta_pct is not None), hsd_growth_valid=(hsd_delta_pct is not None),
        ms_yoy=ms_yoy_pct, hsd_yoy=hsd_yoy_pct,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # OMC-wise MS/HSD breakdown with Prior and YoY
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">OMC-Wise MS & HSD Analytics <span class="section-subtitle">L3M avg · vs prior · vs same period LY</span></div>', unsafe_allow_html=True)

    omc_breakdown = df.groupby("OMC_std").agg(
        MS_L3_Avg=("MS_L3_Avg","sum"), HSD_L3_Avg=("HSD_L3_Avg","sum"),
        MS_P3_Avg=("MS_P3_Avg","sum"), HSD_P3_Avg=("HSD_P3_Avg","sum"),
        MS_LY3_Avg=("MS_LY3_Avg","sum"), HSD_LY3_Avg=("HSD_LY3_Avg","sum"),
        Outlets=("RO Name","count")
    ).reset_index()

    hdr = st.columns([2,1,2,2,2,2,2,2])
    for c, lbl in zip(hdr, ["OMC","Outlets","MS Avg/Mo","vs Prior","vs LY","HSD Avg/Mo","vs Prior","vs LY"]):
        c.markdown(f'<div style="font-size:0.67rem;font-weight:700;text-transform:uppercase;color:#64748b;letter-spacing:0.07em;">{lbl}</div>', unsafe_allow_html=True)

    for _, row in omc_breakdown.sort_values("MS_L3_Avg", ascending=False).iterrows():
        ms_g  = safe_pct(row["MS_L3_Avg"],  row["MS_P3_Avg"])
        ms_yoy= safe_pct(row["MS_L3_Avg"],  row["MS_LY3_Avg"])
        hsd_g = safe_pct(row["HSD_L3_Avg"], row["HSD_P3_Avg"])
        hsd_yoy=safe_pct(row["HSD_L3_Avg"], row["HSD_LY3_Avg"])
        cols = st.columns([2,1,2,2,2,2,2,2])
        cols[0].markdown(f'<b style="color:{OMC_COLORS.get(row["OMC_std"],"#64748b")}">{row["OMC_std"]}</b>', unsafe_allow_html=True)
        cols[1].markdown(str(int(row["Outlets"])))
        cols[2].markdown(f'{row["MS_L3_Avg"]:,.1f} KL')
        cols[3].markdown(growth_fmt(ms_g),   unsafe_allow_html=True)
        cols[4].markdown(growth_fmt(ms_yoy),  unsafe_allow_html=True)
        cols[5].markdown(f'{row["HSD_L3_Avg"]:,.1f} KL')
        cols[6].markdown(growth_fmt(hsd_g),  unsafe_allow_html=True)
        cols[7].markdown(growth_fmt(hsd_yoy), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # District-wise MS/HSD
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">District-Wise MS & HSD Analytics <span class="section-subtitle">L3M avg · monthly avg · vs prior · YoY</span></div>', unsafe_allow_html=True)

    dist_agg = df.groupby("District").agg(
        MS_L3_Avg=("MS_L3_Avg","sum"), HSD_L3_Avg=("HSD_L3_Avg","sum"),
        MS_P3_Avg=("MS_P3_Avg","sum"), HSD_P3_Avg=("HSD_P3_Avg","sum"),
        MS_LY3_Avg=("MS_LY3_Avg","sum"), HSD_LY3_Avg=("HSD_LY3_Avg","sum"),
        Outlets=("RO Name","count")
    ).reset_index()
    dist_agg["Total_Avg"]   = dist_agg["MS_L3_Avg"] + dist_agg["HSD_L3_Avg"]
    dist_agg["MS_Growth"]   = dist_agg.apply(lambda r: safe_pct(r["MS_L3_Avg"],  r["MS_P3_Avg"]),  axis=1)
    dist_agg["HSD_Growth"]  = dist_agg.apply(lambda r: safe_pct(r["HSD_L3_Avg"], r["HSD_P3_Avg"]), axis=1)
    dist_agg["MS_YoY"]      = dist_agg.apply(lambda r: safe_pct(r["MS_L3_Avg"],  r["MS_LY3_Avg"]), axis=1)
    dist_agg["HSD_YoY"]     = dist_agg.apply(lambda r: safe_pct(r["HSD_L3_Avg"], r["HSD_LY3_Avg"]),axis=1)
    dist_agg = dist_agg.sort_values("Total_Avg", ascending=False)

    da1, da2 = st.columns(2, gap="medium")
    with da1:
        fig_ms = px.bar(dist_agg, x="District", y="MS_L3_Avg", color="District",
                        title="MS Avg Monthly Volume by District",
                        color_discrete_sequence=px.colors.qualitative.Set2, text="MS_L3_Avg")
        fig_ms.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
        fig_ms.update_layout(**chart_layout(280, showlegend=False, xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_ms, use_container_width=True)
    with da2:
        fig_hsd = px.bar(dist_agg, x="District", y="HSD_L3_Avg", color="District",
                         title="HSD Avg Monthly Volume by District",
                         color_discrete_sequence=px.colors.qualitative.Pastel, text="HSD_L3_Avg")
        fig_hsd.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
        fig_hsd.update_layout(**chart_layout(280, showlegend=False, xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_hsd, use_container_width=True)

    tbl_data = dist_agg[["District","Outlets","MS_L3_Avg","MS_Growth","MS_YoY",
                           "HSD_L3_Avg","HSD_Growth","HSD_YoY","Total_Avg"]].copy()
    tbl_data.columns = ["District","Outlets","MS Avg/Mo","MS vs Prior %","MS YoY %",
                         "HSD Avg/Mo","HSD vs Prior %","HSD YoY %","Total Avg/Mo"]

    def color_growth(val):
        if pd.isna(val): return "color: #94a3b8"
        if val > 5:  return "color: #16a34a; font-weight:600"
        if val < -5: return "color: #dc2626; font-weight:600"
        return "color: #d97706"

    st.dataframe(
        tbl_data.style
        .format({"MS Avg/Mo":"{:,.1f}","MS vs Prior %":"{:+.1f}%","MS YoY %":"{:+.1f}%",
                 "HSD Avg/Mo":"{:,.1f}","HSD vs Prior %":"{:+.1f}%","HSD YoY %":"{:+.1f}%",
                 "Total Avg/Mo":"{:,.1f}"}, na_rep="N/A")
        .applymap(color_growth, subset=["MS vs Prior %","MS YoY %","HSD vs Prior %","HSD YoY %"]),
        use_container_width=True, hide_index=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # MS vs HSD trend
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">MS vs HSD Volume Trend <span class="section-subtitle">monthly avg per outlet · product mix evolution</span></div>', unsafe_allow_html=True)
    mix_data = [{"Month": col_to_label(mc),
                 "MS Avg": df[mc].mean(), "HSD Avg": df[hc].mean(),
                 "MS_Share": df[mc].mean() / (df[mc].mean() + df[hc].mean()) * 100
                              if (df[mc].mean() + df[hc].mean()) > 0 else 0}
                for mc, hc in zip(ms_cols[-12:], hsd_cols[-12:])]
    mdf = pd.DataFrame(mix_data)

    mc1, mc2 = st.columns(2, gap="medium")
    with mc1:
        fig_mx = go.Figure()
        fig_mx.add_trace(go.Bar(x=mdf["Month"], y=mdf["MS Avg"],  name="MS",  marker_color="#f97316"))
        fig_mx.add_trace(go.Bar(x=mdf["Month"], y=mdf["HSD Avg"], name="HSD", marker_color="#3b82f6"))
        fig_mx.update_layout(**chart_layout(270, barmode="stack",
                                            legend=dict(orientation="h", y=1.08),
                                            yaxis_title="Avg Volume per Outlet (KL)", xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_mx, use_container_width=True)
    with mc2:
        fig_sh = go.Figure()
        fig_sh.add_trace(go.Scatter(x=mdf["Month"], y=mdf["MS_Share"], mode="lines+markers",
                                    name="MS Share %", line=dict(color="#f97316", width=2.5),
                                    fill="tozeroy", fillcolor="rgba(249,115,22,0.1)"))
        fig_sh.add_hline(y=50, line_dash="dot", line_color="#94a3b8")
        fig_sh.update_layout(**chart_layout(270, yaxis=dict(range=[0,100], title="MS Share (%)"),
                                            xaxis=dict(tickangle=-20)))
        st.plotly_chart(fig_sh, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TAB 3: DISTRICT ANALYSIS
# ═══════════════════════════════════════════════════════════════════
with tabs[2]:
    if df["District"].nunique() == 0:
        st.warning("No districts in current filter.")
        st.stop()

    sel_d = st.selectbox("Select District", sorted(df["District"].unique()), key="dist_sel")
    ddf = df[df["District"] == sel_d].copy()

    # District KPIs (averages)
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    d1, d2, d3, d4, d5 = st.columns(5)
    for col, cls, lbl, val, sub in [
        (d1,"blue","Outlets",str(len(ddf)),"in district"),
        (d2,"amber","MS Avg/Mo",f"{ddf['MS_L3_Avg'].sum():,.1f}","KL"),
        (d3,"purple","HSD Avg/Mo",f"{ddf['HSD_L3_Avg'].sum():,.1f}","KL"),
        (d4,"teal","Total Avg/Mo",f"{ddf['Total_L3_Avg'].sum():,.1f}","KL"),
        (d5,"green","Active",str(len(ddf[~ddf["Is_Inactive"]])),"outlets"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi-card {cls}">
            <div class="kpi-label">{lbl}</div>
            <div class="kpi-value">{val}</div>
            <div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # COM Segment breakdown for the district
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">COM Segment Breakdown — {sel_d} <span class="section-subtitle">D1_D2 · C · E_Others</span></div>', unsafe_allow_html=True)
    d_seg = ddf.groupby("COM_Segment").agg(
        Outlets=("RO Name","count"),
        MS_Avg=("MS_L3_Avg","sum"), HSD_Avg=("HSD_L3_Avg","sum"),
        MS_Prior=("MS_P3_Avg","sum"), HSD_Prior=("HSD_P3_Avg","sum"),
        MS_LY=("MS_LY3_Avg","sum"), HSD_LY=("HSD_LY3_Avg","sum"),
    ).reset_index()
    d_seg["Total_Avg"] = d_seg["MS_Avg"] + d_seg["HSD_Avg"]
    d_seg["vs_Prior"]  = d_seg.apply(lambda r: safe_pct(r["Total_Avg"], r["MS_Prior"]+r["HSD_Prior"]), axis=1)
    d_seg["YoY"]       = d_seg.apply(lambda r: safe_pct(r["Total_Avg"], r["MS_LY"]+r["HSD_LY"]), axis=1)

    dseg_cols = st.columns(3)
    for col_w, seg, css in zip(dseg_cols, seg_order, seg_css):
        row = d_seg[d_seg["COM_Segment"] == seg]
        if row.empty:
            with col_w:
                st.markdown(f'<div class="mini-kpi"><div class="mini-kpi-label">{seg}</div><div class="mini-kpi-value">—</div></div>', unsafe_allow_html=True)
            continue
        r = row.iloc[0]
        with col_w:
            st.markdown(f"""<div class="mini-kpi">
            <div class="mini-kpi-label">{seg} ({int(r['Outlets'])} outlets)</div>
            <div class="mini-kpi-value">{r['Total_Avg']:,.1f} KL/mo</div>
            <div class="mini-kpi-sub">MS {r['MS_Avg']:,.1f} + HSD {r['HSD_Avg']:,.1f}</div>
            <div class="mini-kpi-sub">vs Prior: {f"{r['vs_Prior']:+.1f}%" if r['vs_Prior'] else 'N/A'} | YoY: {f"{r['YoY']:+.1f}%" if r['YoY'] else 'N/A'}</div>
            </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # MS/HSD analytics for district
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">MS & HSD Analytics — {sel_d}</div>', unsafe_allow_html=True)
    d_ms_g   = safe_pct(ddf["MS_L3_Avg"].sum(),  ddf["MS_P3_Avg"].sum())
    d_hsd_g  = safe_pct(ddf["HSD_L3_Avg"].sum(), ddf["HSD_P3_Avg"].sum())
    d_ms_yoy = safe_pct(ddf["MS_L3_Avg"].sum(),  ddf["MS_LY3_Avg"].sum())
    d_hsd_yoy= safe_pct(ddf["HSD_L3_Avg"].sum(), ddf["HSD_LY3_Avg"].sum())

    render_ms_hsd_analytics(
        ms_avg=ddf["MS_L3_Avg"].sum(), hsd_avg=ddf["HSD_L3_Avg"].sum(),
        ms_prior_avg=ddf["MS_P3_Avg"].sum(), hsd_prior_avg=ddf["HSD_P3_Avg"].sum(),
        ms_ly_avg=ddf["MS_LY3_Avg"].sum(), hsd_ly_avg=ddf["HSD_LY3_Avg"].sum(),
        ms_growth=d_ms_g, hsd_growth=d_hsd_g,
        ms_growth_valid=(d_ms_g is not None), hsd_growth_valid=(d_hsd_g is not None),
        ms_yoy=d_ms_yoy, hsd_yoy=d_hsd_yoy,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([3, 2], gap="medium")
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Outlet Rankings <span class="section-subtitle">avg monthly volume · purple = outlier</span></div>', unsafe_allow_html=True)
        r10 = ddf.nlargest(10, "Total_L3_Avg").copy()
        r10["bar_color"] = r10.apply(
            lambda r: "#c084fc" if r["Is_Outlier"] else OMC_COLORS.get(r["OMC_std"], "#94a3b8"), axis=1)
        fig_r = go.Figure(go.Bar(
            x=r10["Total_L3_Avg"], y=r10["RO Name"], orientation="h",
            marker_color=r10["bar_color"],
            text=r10["Total_L3_Avg"].map(lambda x: f"{x:,.1f}"), textposition="outside",
        ))
        fig_r.update_layout(**chart_layout(330, yaxis=dict(autorange="reversed"),
                                            xaxis_title="Avg Monthly Volume (KL)"))
        st.plotly_chart(fig_r, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">OMC Share in District</div>', unsafe_allow_html=True)
        prim_d = ddf[ddf["OMC_std"].isin(PRIMARY_OMCS)]
        if not prim_d.empty:
            ds = prim_d.groupby("OMC_std")["Total_L3_Avg"].sum().reset_index()
            fig_dp = px.pie(ds, values="Total_L3_Avg", names="OMC_std", hole=0.4,
                            color="OMC_std", color_discrete_map=OMC_COLORS)
            fig_dp.update_layout(height=330, margin=dict(t=5,b=5,l=5,r=5),
                                 paper_bgcolor="#ffffff", font=CHART_FONT)
            st.plotly_chart(fig_dp, use_container_width=True)
        else:
            st.info("No primary OMC data.")
        st.markdown('</div>', unsafe_allow_html=True)

    # COM Segment stacked bar (consistent with global grouping)
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Volume by COM Segment <span class="section-subtitle">D1_D2 · C · E_Others</span></div>', unsafe_allow_html=True)
    seg_bar = ddf.groupby("COM_Segment")[["MS_L3_Avg","HSD_L3_Avg"]].sum().reindex(seg_order).reset_index()
    fig_segbar = go.Figure()
    fig_segbar.add_trace(go.Bar(x=seg_bar["COM_Segment"], y=seg_bar["MS_L3_Avg"],
                                 name="MS", marker_color="#f97316",
                                 text=seg_bar["MS_L3_Avg"].map("{:,.1f}".format), textposition="outside"))
    fig_segbar.add_trace(go.Bar(x=seg_bar["COM_Segment"], y=seg_bar["HSD_L3_Avg"],
                                 name="HSD", marker_color="#3b82f6",
                                 text=seg_bar["HSD_L3_Avg"].map("{:,.1f}".format), textposition="outside"))
    fig_segbar.update_layout(**chart_layout(260, barmode="group",
                                             legend=dict(orientation="h", y=1.08),
                                             yaxis_title="Avg Monthly Volume (KL)"))
    st.plotly_chart(fig_segbar, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("Full District Table — with MS/HSD avg & growth", expanded=False):
        tbl = ddf[["RO Name","OMC_std","COM_Segment","MS_L3_Avg","MS_L3_Growth",
                    "HSD_L3_Avg","HSD_L3_Growth","Total_L3_Avg",
                    "MoM_Growth_Label","Status_Label","District_Rank"]].copy()
        tbl = tbl.sort_values("Total_L3_Avg", ascending=False).reset_index(drop=True)
        tbl.index += 1
        tbl.rename(columns={
            "OMC_std":"OMC","COM_Segment":"Segment",
            "MS_L3_Avg":"MS Avg/Mo","MS_L3_Growth":"MS Growth %",
            "HSD_L3_Avg":"HSD Avg/Mo","HSD_L3_Growth":"HSD Growth %",
            "Total_L3_Avg":"Total Avg/Mo","MoM_Growth_Label":"MoM Growth",
            "Status_Label":"Status","District_Rank":"Rank",
        }, inplace=True)
        st.dataframe(
            tbl.style.format({
                "MS Avg/Mo":"{:,.1f}","MS Growth %":"{:+.1f}%",
                "HSD Avg/Mo":"{:,.1f}","HSD Growth %":"{:+.1f}%",
                "Total Avg/Mo":"{:,.1f}",
            }, na_rep="N/A"),
            use_container_width=True,
        )

# ═══════════════════════════════════════════════════════════════════
# TAB 4: SEARCH & OUTLET DETAIL
# ═══════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Outlet Search</div>', unsafe_allow_html=True)
    search = st.text_input("Search by RO Name or SAP Code",
                            placeholder="e.g. ABC Petrol or 143302",
                            label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    if search:
        mask = (df_raw["RO Name"].str.contains(search, case=False, na=False) |
                df_raw["SAP Code"].astype(str).str.contains(search, case=False, na=False))
        results = df_raw[mask]
        if results.empty:
            st.markdown('<div class="error-banner">No outlets matched your search.</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="info-banner">Found {len(results)} outlet(s).</div>', unsafe_allow_html=True)
            sel_name = st.selectbox("Select outlet", results["RO Name"].tolist(), key="o_sel")
            outlet = df_raw[df_raw["RO Name"] == sel_name].iloc[0]

            # Status banners
            if outlet["Is_Inactive"]:
                st.markdown('<div class="error-banner"><b>Inactive Outlet.</b> Near-zero sales for 3+ consecutive months.</div>', unsafe_allow_html=True)
            elif outlet["Is_Restart"]:
                st.markdown('<div class="warn-banner"><b>Restart Pattern.</b> Outlet was inactive and has recently resumed.</div>', unsafe_allow_html=True)
            elif outlet["Has_Spike"]:
                st.markdown('<div class="warn-banner"><b>Unusual Sales Spike.</b> Latest month is 3× the recent average.</div>', unsafe_allow_html=True)
            elif outlet["Is_Outlier"]:
                st.markdown('<div class="warn-banner"><b>Statistical Outlier.</b> Deviates significantly from district peers.</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="section-title">{sel_name} <span class="section-subtitle">{outlet["OMC_std"]} · {outlet["District"]} · SAP {outlet["SAP Code"]} · COM: {outlet.get("COM","—")} · Segment: {outlet["COM_Segment"]}</span></div>', unsafe_allow_html=True)

            # Core KPI cards (averages)
            r1, r2, r3, r4 = st.columns(4)
            for col, cls, lbl, val, sub in [
                (r1,"amber","MS Avg/Mo",f"{outlet['MS_L3_Avg']:,.1f}","KL · L3M"),
                (r2,"blue","HSD Avg/Mo",f"{outlet['HSD_L3_Avg']:,.1f}","KL · L3M"),
                (r3,"purple","Total Avg/Mo",f"{outlet['Total_L3_Avg']:,.1f}","KL · L3M"),
                (r4,"green","MoM Growth",outlet["MoM_Growth_Label"],"latest month"),
            ]:
                with col:
                    st.markdown(f"""<div class="kpi-card {cls}">
                    <div class="kpi-label">{lbl}</div>
                    <div class="kpi-value" style="font-size:1.3rem;">{val}</div>
                    <div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)

            # Comparison metrics: Current L3M Avg, Prior L3M Avg, Same Period LY Avg
            st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
            st.markdown('<div style="font-size:0.78rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:0.6rem;">Comparison Metrics — L3M Avg vs Prior vs Same Period LY</div>', unsafe_allow_html=True)
            outlet_flags = {"is_inactive": outlet["Is_Inactive"], "is_restart": outlet["Is_Restart"], "has_spike": outlet["Has_Spike"]}
            o_ms_g,  _ = compute_segment_growth(outlet["MS_L3_Avg"],
                                                  outlet["MS_P3_Avg"]  if not pd.isna(outlet["MS_P3_Avg"])  else None,
                                                  outlet_flags)
            o_hsd_g, _ = compute_segment_growth(outlet["HSD_L3_Avg"],
                                                  outlet["HSD_P3_Avg"] if not pd.isna(outlet["HSD_P3_Avg"]) else None,
                                                  outlet_flags)
            o_ms_yoy, _ = compute_segment_growth(outlet["MS_L3_Avg"],
                                                   outlet["MS_LY3_Avg"] if not pd.isna(outlet["MS_LY3_Avg"]) else None,
                                                   outlet_flags)
            o_hsd_yoy,_ = compute_segment_growth(outlet["HSD_L3_Avg"],
                                                   outlet["HSD_LY3_Avg"] if not pd.isna(outlet["HSD_LY3_Avg"]) else None,
                                                   outlet_flags)
            render_ms_hsd_analytics(
                ms_avg=outlet["MS_L3_Avg"],  hsd_avg=outlet["HSD_L3_Avg"],
                ms_prior_avg=outlet["MS_P3_Avg"]  if not pd.isna(outlet["MS_P3_Avg"])  else None,
                hsd_prior_avg=outlet["HSD_P3_Avg"] if not pd.isna(outlet["HSD_P3_Avg"]) else None,
                ms_ly_avg=outlet["MS_LY3_Avg"]  if not pd.isna(outlet["MS_LY3_Avg"])  else None,
                hsd_ly_avg=outlet["HSD_LY3_Avg"] if not pd.isna(outlet["HSD_LY3_Avg"]) else None,
                ms_growth=o_ms_g, hsd_growth=o_hsd_g,
                ms_growth_valid=(o_ms_g is not None and not outlet["Is_Inactive"]),
                hsd_growth_valid=(o_hsd_g is not None and not outlet["Is_Inactive"]),
                ms_yoy=o_ms_yoy, hsd_yoy=o_hsd_yoy,
            )

            # Rank cards
            r2c1, r2c2, r2c3, r2c4 = st.columns(4)
            ta_name_disp = outlet.get("TA_Name","") or outlet.get("Trading_Area_Clean","")
            ta_peers_count = len(df_raw[df_raw["Trading_Area_Clean"] == outlet.get("Trading_Area_Clean","")]) if "Trading_Area_Clean" in df_raw.columns else 0
            for col, cls, lbl, val, sub in [
                (r2c1,"blue","Global Rank",f"#{int(outlet['Overall_Rank'])}",f"of {len(df_raw)}"),
                (r2c2,"amber","District Rank",f"#{int(outlet['District_Rank'])}",outlet["District"]),
                (r2c3,"green","OMC Rank",f"#{int(outlet['OMC_Rank'])}",outlet["OMC_std"]),
                (r2c4,"teal","Trading Area Rank",f"#{int(outlet['TA_Rank'])}",
                 f"{ta_peers_count} outlets in TA" if ta_peers_count else "TA"),
            ]:
                with col:
                    st.markdown(f"""<div class="kpi-card {cls}">
                    <div class="kpi-label">{lbl}</div>
                    <div class="kpi-value">{val}</div>
                    <div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Monthly history chart
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Monthly History (12M)</div>', unsafe_allow_html=True)
            hdata = [{"Month": col_to_label(mc), "MS": outlet[mc], "HSD": outlet[hc]}
                     for mc, hc in zip(ms_cols[-12:], hsd_cols[-12:])]
            hdf2 = pd.DataFrame(hdata)
            fig_hh = go.Figure()
            fig_hh.add_trace(go.Scatter(x=hdf2["Month"], y=hdf2["MS"], mode="lines+markers",
                                        name="MS", line=dict(color="#f97316", width=2.2)))
            fig_hh.add_trace(go.Scatter(x=hdf2["Month"], y=hdf2["HSD"], mode="lines+markers",
                                        name="HSD", line=dict(color="#3b82f6", width=2.2)))
            if len(hdata) >= 3:
                l3_x = [hdata[-3]["Month"], hdata[-1]["Month"]]
                fig_hh.add_vrect(x0=l3_x[0], x1=l3_x[1],
                                 fillcolor="rgba(99,102,241,0.06)", line_width=0,
                                 annotation_text="L3M", annotation_position="top left")
            fig_hh.update_layout(**chart_layout(280, legend=dict(orientation="h", y=1.08)))
            st.plotly_chart(fig_hh, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # ── RO Effectivity & Market Share ────────────────────
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">RO Effectivity & Market Share <span class="section-subtitle">RO volume ÷ avg trading area volume · monthly trend</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="info-banner">Effectivity = RO monthly volume ÷ Trading Area avg monthly volume × 100. Market Share = RO volume ÷ Total TA volume × 100.</div>', unsafe_allow_html=True)

            eff_ro = compute_ro_effectivity(outlet, df_raw, ms_cols, hsd_cols, n_months=12)
            if not eff_ro.empty and not eff_ro["MS_Eff"].isna().all():
                ec1, ec2 = st.columns(2, gap="medium")
                with ec1:
                    fig_roeff = go.Figure()
                    fig_roeff.add_trace(go.Scatter(x=eff_ro["Month"], y=eff_ro["MS_Eff"],
                                                    mode="lines+markers", name="MS Effectivity",
                                                    line=dict(color="#f97316", width=2.2)))
                    fig_roeff.add_trace(go.Scatter(x=eff_ro["Month"], y=eff_ro["HSD_Eff"],
                                                    mode="lines+markers", name="HSD Effectivity",
                                                    line=dict(color="#3b82f6", width=2.2)))
                    fig_roeff.add_hline(y=100, line_dash="dot", line_color="#94a3b8",
                                        annotation_text="TA Avg (100%)", annotation_position="bottom right")
                    fig_roeff.update_layout(**chart_layout(260, yaxis_title="Effectivity (%)",
                                                           legend=dict(orientation="h", y=1.08),
                                                           xaxis=dict(tickangle=-20)))
                    st.plotly_chart(fig_roeff, use_container_width=True)
                with ec2:
                    fig_share = go.Figure()
                    fig_share.add_trace(go.Scatter(x=eff_ro["Month"], y=eff_ro["MS_Share"],
                                                    mode="lines+markers", name="MS Market Share %",
                                                    line=dict(color="#f97316", width=2.2)))
                    fig_share.add_trace(go.Scatter(x=eff_ro["Month"], y=eff_ro["HSD_Share"],
                                                    mode="lines+markers", name="HSD Market Share %",
                                                    line=dict(color="#3b82f6", width=2.2)))
                    fig_share.update_layout(**chart_layout(260, yaxis_title="Market Share (%)",
                                                            legend=dict(orientation="h", y=1.08),
                                                            xaxis=dict(tickangle=-20)))
                    st.plotly_chart(fig_share, use_container_width=True)

                # Latest month summary
                latest_eff_row = eff_ro.iloc[-1]
                eff1, eff2, eff3, eff4 = st.columns(4)
                for col_w, lbl, val, css in [
                    (eff1,"MS Effectivity", f"{latest_eff_row['MS_Eff']:.1f}%" if latest_eff_row['MS_Eff'] else "N/A","amber"),
                    (eff2,"HSD Effectivity",f"{latest_eff_row['HSD_Eff']:.1f}%" if latest_eff_row['HSD_Eff'] else "N/A","blue"),
                    (eff3,"MS Market Share", f"{latest_eff_row['MS_Share']:.1f}%" if latest_eff_row['MS_Share'] else "N/A","indigo"),
                    (eff4,"HSD Market Share",f"{latest_eff_row['HSD_Share']:.1f}%" if latest_eff_row['HSD_Share'] else "N/A","teal"),
                ]:
                    with col_w:
                        st.markdown(f"""<div class="mini-kpi">
                        <div class="mini-kpi-label">{lbl} (latest month)</div>
                        <div class="mini-kpi-value">{val}</div></div>""", unsafe_allow_html=True)
            else:
                st.markdown('<div class="info-banner">No Trading Area peers found for effectivity computation.</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Peer comparisons — District AND Trading Area side by side
            oc1, oc2 = st.columns(2, gap="medium")

            with oc1:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-title">vs District Peers</div>', unsafe_allow_html=True)
                dist_peers = df_raw[df_raw["District"] == outlet["District"]]
                cdf_d = pd.DataFrame({
                    "Category": ["This Outlet", "District Avg", "District Top"],
                    "MS Avg/Mo":  [outlet["MS_L3_Avg"],  dist_peers["MS_L3_Avg"].mean(),  dist_peers["MS_L3_Avg"].max()],
                    "HSD Avg/Mo": [outlet["HSD_L3_Avg"], dist_peers["HSD_L3_Avg"].mean(), dist_peers["HSD_L3_Avg"].max()],
                })
                fig_cd = go.Figure()
                fig_cd.add_trace(go.Bar(name="MS",  x=cdf_d["Category"], y=cdf_d["MS Avg/Mo"],
                                        marker_color="#f97316",
                                        text=cdf_d["MS Avg/Mo"].map("{:,.1f}".format), textposition="outside"))
                fig_cd.add_trace(go.Bar(name="HSD", x=cdf_d["Category"], y=cdf_d["HSD Avg/Mo"],
                                        marker_color="#3b82f6",
                                        text=cdf_d["HSD Avg/Mo"].map("{:,.1f}".format), textposition="outside"))
                fig_cd.update_layout(**chart_layout(290, barmode="group", showlegend=True,
                                                    legend=dict(orientation="h", y=1.08),
                                                    yaxis_title="Avg Monthly Volume (KL)"))
                st.plotly_chart(fig_cd, use_container_width=True)
                st.markdown(f'<div class="info-banner" style="font-size:0.75rem;">'
                            f'District: <b>{outlet["District"]}</b> &nbsp;·&nbsp; '
                            f'{len(dist_peers)} outlets &nbsp;·&nbsp; '
                            f'Avg {dist_peers["Total_L3_Avg"].mean():,.1f} KL/mo &nbsp;·&nbsp; '
                            f'Top {dist_peers["Total_L3_Avg"].max():,.1f} KL/mo</div>',
                            unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with oc2:
                st.markdown('<div class="section-card">', unsafe_allow_html=True)
                ta_key  = outlet.get("Trading_Area_Clean","")
                ta_name = outlet.get("TA_Name","") or ta_key
                ta_peers = df_raw[df_raw["Trading_Area_Clean"] == ta_key] if ta_key else pd.DataFrame()

                if not ta_peers.empty and len(ta_peers) > 1:
                    st.markdown(f'<div class="section-title">vs Trading Area Peers'
                                f'<span class="section-subtitle">{ta_name}</span></div>',
                                unsafe_allow_html=True)

                    cdf_ta = pd.DataFrame({
                        "Category": ["This Outlet", "TA Avg", "TA Top"],
                        "MS Avg/Mo":  [outlet["MS_L3_Avg"],  ta_peers["MS_L3_Avg"].mean(),  ta_peers["MS_L3_Avg"].max()],
                        "HSD Avg/Mo": [outlet["HSD_L3_Avg"], ta_peers["HSD_L3_Avg"].mean(), ta_peers["HSD_L3_Avg"].max()],
                    })
                    fig_cta = go.Figure()
                    fig_cta.add_trace(go.Bar(name="MS",  x=cdf_ta["Category"], y=cdf_ta["MS Avg/Mo"],
                                             marker_color="#f97316",
                                             text=cdf_ta["MS Avg/Mo"].map("{:,.1f}".format), textposition="outside"))
                    fig_cta.add_trace(go.Bar(name="HSD", x=cdf_ta["Category"], y=cdf_ta["HSD Avg/Mo"],
                                             marker_color="#3b82f6",
                                             text=cdf_ta["HSD Avg/Mo"].map("{:,.1f}".format), textposition="outside"))
                    fig_cta.update_layout(**chart_layout(230, barmode="group", showlegend=True,
                                                         legend=dict(orientation="h", y=1.08),
                                                         yaxis_title="Avg Monthly Volume (KL)"))
                    st.plotly_chart(fig_cta, use_container_width=True)

                    # Peer rank table
                    st.markdown('<div style="font-size:0.72rem;font-weight:700;color:#64748b;'
                                'text-transform:uppercase;letter-spacing:0.07em;margin:0.5rem 0 0.4rem;">All Outlets in This Trading Area</div>',
                                unsafe_allow_html=True)
                    ta_tbl = ta_peers[["RO Name","OMC_std","COM_Segment",
                                        "Total_L3_Avg","MS_L3_Avg","HSD_L3_Avg","TA_Rank"]].copy()
                    ta_tbl = ta_tbl.sort_values("TA_Rank")
                    ta_tbl["This"] = ta_tbl["RO Name"].apply(lambda x: "◀" if x == sel_name else "")
                    # Market share within TA
                    ta_total_ms  = ta_peers["MS_L3_Avg"].sum()
                    ta_total_hsd = ta_peers["HSD_L3_Avg"].sum()
                    ta_tbl["MS Share%"]  = ta_peers["MS_L3_Avg"].values  / ta_total_ms  * 100 if ta_total_ms  > 0 else 0
                    ta_tbl["HSD Share%"] = ta_peers["HSD_L3_Avg"].values / ta_total_hsd * 100 if ta_total_hsd > 0 else 0
                    ta_tbl.rename(columns={
                        "OMC_std":"OMC","COM_Segment":"Segment",
                        "Total_L3_Avg":"Total Avg/Mo","MS_L3_Avg":"MS Avg/Mo",
                        "HSD_L3_Avg":"HSD Avg/Mo","TA_Rank":"TA Rank"}, inplace=True)
                    st.dataframe(
                        ta_tbl[["TA Rank","RO Name","OMC","Segment","MS Avg/Mo","MS Share%",
                                 "HSD Avg/Mo","HSD Share%","Total Avg/Mo","This"]]
                        .style.format({"MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}",
                                       "Total Avg/Mo":"{:,.1f}",
                                       "MS Share%":"{:.1f}%","HSD Share%":"{:.1f}%"})
                        .applymap(lambda v: "font-weight:700;color:#6366f1" if v == "◀" else "",
                                  subset=["This"]),
                        use_container_width=True, hide_index=True, height=220,
                    )
                    st.markdown(f'<div class="info-banner" style="font-size:0.75rem;">'
                                f'Trading Area: <b>{ta_name}</b> &nbsp;·&nbsp; '
                                f'{len(ta_peers)} outlets &nbsp;·&nbsp; '
                                f'Avg {ta_peers["Total_L3_Avg"].mean():,.1f} KL/mo &nbsp;·&nbsp; '
                                f'Top {ta_peers["Total_L3_Avg"].max():,.1f} KL/mo</div>',
                                unsafe_allow_html=True)
                else:
                    st.markdown('<div class="section-title">vs Trading Area Peers</div>', unsafe_allow_html=True)
                    st.markdown('<div class="info-banner">No other outlets found in this Trading Area.</div>',
                                unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-banner">Enter an outlet name or SAP code above to view detailed analysis, effectivity, market share, and peer comparisons.</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TAB 5: AI PREDICTIONS
# ═══════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">AI Prediction Module <span class="section-subtitle">smoothed linear regression · next month forecast · inactive/spike-adjusted</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="info-banner">Predictions use exponential smoothing + linear regression on last 6 months of data. Inactive/restart/spike outlets are noted without overriding predictions with unreliable baselines.</div>', unsafe_allow_html=True)
    pc1, pc2 = st.columns([3, 1])
    with pc1:
        pred_district = st.selectbox("Scope", ["All Districts"] + sorted(df["District"].unique()), key="pred_d")
    with pc2:
        top_n = st.number_input("Top N", min_value=5, max_value=50, value=10)
    st.markdown('</div>', unsafe_allow_html=True)

    pred_src = df if pred_district == "All Districts" else df[df["District"] == pred_district]
    pred_src = pred_src[~pred_src["Is_Inactive"]]

    with st.spinner("Computing predictions..."):
        preds = []
        for _, row in pred_src.iterrows():
            p    = predict_outlet(row, ms_cols, hsd_cols)
            note = ("Recent restart — caution" if row["Is_Restart"] else
                    "Spike detected" if row["Has_Spike"] else
                    "Outlier" if row["Is_Outlier"] else "")
            preds.append({
                "RO Name": row["RO Name"], "OMC": row["OMC_std"],
                "District": row["District"], "Segment": row["COM_Segment"],
                "Current Avg/Mo": row["Total_L3_Avg"],
                "Pred MS": p["ms_pred"], "Pred HSD": p["hsd_pred"], "Pred Total": p["total_pred"],
                "MS Growth %": p["ms_growth"], "HSD Growth %": p["hsd_growth"],
                "Pred Growth %": p["total_growth"], "Trend": p["trend"], "Note": note,
            })
        pred_df = pd.DataFrame(preds).sort_values("Pred Total", ascending=False).head(top_n)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">Top {top_n} Predicted Performers — Next Month</div>', unsafe_allow_html=True)
    st.dataframe(
        pred_df.style
        .format({"Current Avg/Mo":"{:,.1f}","Pred MS":"{:,.1f}","Pred HSD":"{:,.1f}",
                 "Pred Total":"{:,.1f}","MS Growth %":"{:.1f}%","HSD Growth %":"{:.1f}%",
                 "Pred Growth %":"{:.1f}%"})
        .applymap(lambda v: ("color:#16a34a;font-weight:600" if v=="Increase" else
                             "color:#ef4444;font-weight:600" if v=="Decline" else "color:#d97706"),
                  subset=["Trend"]),
        use_container_width=True, hide_index=True, height=360,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    pc1, pc2 = st.columns(2, gap="medium")
    with pc1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Trend Distribution</div>', unsafe_allow_html=True)
        tc = pd.Series([p["Trend"] for p in preds]).value_counts().reset_index()
        tc.columns = ["Trend","Count"]
        fig_tc = px.bar(tc, x="Trend", y="Count", color="Trend",
                        color_discrete_map={"Increase":"#22c55e","Stable":"#f59e0b","Decline":"#ef4444"},
                        text="Count")
        fig_tc.update_traces(textposition="outside")
        fig_tc.update_layout(**chart_layout(260, showlegend=False))
        st.plotly_chart(fig_tc, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with pc2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Top Predicted — History + Forecast</div>', unsafe_allow_html=True)
        if not pred_df.empty:
            top_name = pred_df.iloc[0]["RO Name"]
            top_row  = df_raw[df_raw["RO Name"] == top_name].iloc[0]
            tp       = predict_outlet(top_row, ms_cols, hsd_cols)
            hms  = [top_row[c] for c in ms_cols[-6:]]
            hhsd = [top_row[c] for c in hsd_cols[-6:]]
            xlbl = [col_to_label(c) for c in ms_cols[-6:]] + ["Next Mo."]
            fig_fp = go.Figure()
            fig_fp.add_trace(go.Scatter(x=xlbl[:-1], y=hms, mode="lines+markers", name="MS",
                                        line=dict(color="#f97316", width=2)))
            fig_fp.add_trace(go.Scatter(x=xlbl[:-1], y=hhsd, mode="lines+markers", name="HSD",
                                        line=dict(color="#3b82f6", width=2)))
            fig_fp.add_trace(go.Scatter(x=[xlbl[-2], xlbl[-1]], y=[hms[-1], tp["ms_pred"]],
                                        mode="lines+markers", name="MS Pred",
                                        line=dict(color="#f97316", dash="dash"),
                                        marker=dict(symbol="star", size=12)))
            fig_fp.add_trace(go.Scatter(x=[xlbl[-2], xlbl[-1]], y=[hhsd[-1], tp["hsd_pred"]],
                                        mode="lines+markers", name="HSD Pred",
                                        line=dict(color="#3b82f6", dash="dash"),
                                        marker=dict(symbol="star", size=12)))
            fig_fp.update_layout(**chart_layout(260, legend=dict(orientation="h", y=1.08)))
            st.plotly_chart(fig_fp, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TAB 6: DATA QUALITY
# ═══════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Data Quality Dashboard <span class="section-subtitle">outlier detection · inactive flagging · spike analysis · growth normalization</span></div>', unsafe_allow_html=True)
    q1, q2, q3, q4 = st.columns(4)
    for col, cls, lbl, val, sub in [
        (q1,"red","Inactive Outlets",str(n_inactive),"zero sales 3+ months"),
        (q2,"amber","Restart Events",str(n_restart),"growth suppressed"),
        (q3,"purple","Outliers",str(n_outlier),"Z-score + IQR method"),
        (q4,"amber","Spike Alerts",str(n_spike),"verify source data"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi-card {cls}">
            <div class="kpi-label">{lbl}</div>
            <div class="kpi-value">{val}</div>
            <div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander(f"Inactive Outlets ({n_inactive})", expanded=n_inactive > 0):
        idf = df_raw[df_raw["Is_Inactive"]][["RO Name","OMC_std","District","COM_Segment",
                                               "MS_L3_Avg","HSD_L3_Avg","Total_L3_Avg"]].copy()
        if idf.empty:
            st.markdown('<div class="success-banner">No inactive outlets detected.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="error-banner">Near-zero sales for 3+ consecutive months.</div>', unsafe_allow_html=True)
            st.dataframe(idf.rename(columns={"OMC_std":"OMC","COM_Segment":"Segment",
                                              "MS_L3_Avg":"MS Avg/Mo","HSD_L3_Avg":"HSD Avg/Mo",
                                              "Total_L3_Avg":"Total Avg/Mo"})
                         .style.format({"MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}","Total Avg/Mo":"{:,.1f}"}),
                         use_container_width=True, hide_index=True)

    with st.expander(f"Restart Pattern Outlets ({n_restart})", expanded=False):
        rdf = df_raw[df_raw["Is_Restart"]][["RO Name","OMC_std","District","COM_Segment",
                                              "MS_L3_Avg","HSD_L3_Avg","Total_L3_Avg"]].copy()
        if rdf.empty:
            st.markdown('<div class="success-banner">No restart patterns detected.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-banner">Previously inactive, recently resumed. MoM growth % suppressed.</div>', unsafe_allow_html=True)
            st.dataframe(rdf.rename(columns={"OMC_std":"OMC","COM_Segment":"Segment",
                                              "MS_L3_Avg":"MS Avg/Mo","HSD_L3_Avg":"HSD Avg/Mo",
                                              "Total_L3_Avg":"Total Avg/Mo"})
                         .style.format({"MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}","Total Avg/Mo":"{:,.1f}"}),
                         use_container_width=True, hide_index=True)

    with st.expander(f"Statistical Outliers ({n_outlier})", expanded=False):
        odf = df_raw[df_raw["Is_Outlier"]][["RO Name","OMC_std","District","COM_Segment","Total_L3_Avg"]].copy()
        if odf.empty:
            st.markdown('<div class="success-banner">No statistical outliers detected.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-banner">Detected via Z-score (2.5σ) AND IQR (2.0× fence) — both must agree.</div>', unsafe_allow_html=True)
            st.dataframe(odf.rename(columns={"OMC_std":"OMC","COM_Segment":"Segment","Total_L3_Avg":"Total Avg/Mo"})
                         .style.format({"Total Avg/Mo":"{:,.1f}"}),
                         use_container_width=True, hide_index=True)

    with st.expander(f"Unusual Spike Outlets ({n_spike})", expanded=False):
        sdf = df_raw[df_raw["Has_Spike"]][["RO Name","OMC_std","District","COM_Segment",
                                            "MS_L3_Avg","HSD_L3_Avg","Total_L3_Avg"]].copy()
        if sdf.empty:
            st.markdown('<div class="success-banner">No unusual spikes detected.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-banner">Latest-month sales exceed 3× the rolling average.</div>', unsafe_allow_html=True)
            st.dataframe(sdf.rename(columns={"OMC_std":"OMC","COM_Segment":"Segment",
                                              "MS_L3_Avg":"MS Avg/Mo","HSD_L3_Avg":"HSD Avg/Mo",
                                              "Total_L3_Avg":"Total Avg/Mo"})
                         .style.format({"MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}","Total Avg/Mo":"{:,.1f}"}),
                         use_container_width=True, hide_index=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Growth Computation Quality <span class="section-subtitle">meaningful vs suppressed</span></div>', unsafe_allow_html=True)
    total_o   = len(df_raw)
    meaningful = int(df_raw["MoM_Growth_Meaningful"].sum())
    suppressed = total_o - meaningful
    gq = pd.DataFrame({"Category": ["Meaningful growth computed","Growth suppressed"], "Count": [meaningful, suppressed]})
    fig_gq = px.bar(gq, x="Category", y="Count", color="Category",
                    color_discrete_sequence=["#22c55e","#ef4444"], text="Count")
    fig_gq.update_traces(textposition="outside")
    fig_gq.update_layout(**chart_layout(240, showlegend=False))
    st.plotly_chart(fig_gq, use_container_width=True)
    reasons = df_raw[df_raw["MoM_Growth_Meaningful"]==False]["MoM_Growth_Label"].value_counts().reset_index()
    reasons.columns = ["Reason for Suppression","Count"]
    st.dataframe(reasons, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════
# TAB 7: INSIGHTS
# ═══════════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Intelligent Insights Engine <span class="section-subtitle">auto-generated from data patterns</span></div>', unsafe_allow_html=True)

    active_df2 = df[~df["Is_Inactive"]]
    total_s    = active_df2["Total_L3_Avg"].sum()
    insights   = []

    if not active_df2.empty and total_s > 0:
        t5s = active_df2.nlargest(5,"Total_L3_Avg")["Total_L3_Avg"].sum()
        t5p = t5s / total_s * 100
        t5n = active_df2.nlargest(5,"Total_L3_Avg")["RO Name"].tolist()
        insights.append(("info", f"Top 5 outlets contribute <b>{t5p:.1f}%</b> of active avg monthly volume ({t5s:,.1f} KL of {total_s:,.1f} KL). Leaders: {', '.join(t5n[:3])}..."))

    prim_t = df[df["OMC_std"].isin(PRIMARY_OMCS)]["Total_L3_Avg"].sum()
    for omc in PRIMARY_OMCS:
        ov = df[df["OMC_std"]==omc]["Total_L3_Avg"].sum()
        if prim_t > 0:
            omc_ms  = df[df["OMC_std"]==omc]["MS_L3_Avg"].sum()
            omc_hsd = df[df["OMC_std"]==omc]["HSD_L3_Avg"].sum()
            insights.append(("info", f"<b>{omc}</b> holds <b>{ov/prim_t*100:.1f}%</b> avg monthly share — MS: {omc_ms:,.1f} KL/mo, HSD: {omc_hsd:,.1f} KL/mo."))

    # COM segment insights
    for seg in seg_order:
        seg_row = seg_agg[seg_agg["COM_Segment"] == seg]
        if seg_row.empty: continue
        r = seg_row.iloc[0]
        yoy = safe_pct(r["Total_Avg"], r["Total_LY_Avg"])
        if yoy is not None:
            icon = "success" if yoy > 0 else "warning"
            insights.append((icon, f"COM Segment <b>{seg}</b> ({int(r['Outlets'])} outlets): avg {r['Total_Avg']:,.1f} KL/mo — YoY: <b>{yoy:+.1f}%</b>."))

    fg_df = df[df["MoM_Growth_Meaningful"]==True]
    if not fg_df.empty:
        fastest = fg_df.loc[fg_df["MoM_Growth"].idxmax()]
        insights.append(("success", f"Fastest growing outlet: <b>{fastest['RO Name']}</b> ({fastest['OMC_std']}, {fastest['District']}, {fastest['COM_Segment']}) — <b>{fastest['MoM_Growth']:+.1f}%</b> MoM growth."))
        slowest = fg_df.loc[fg_df["MoM_Growth"].idxmin()]
        if slowest["MoM_Growth"] < -5:
            insights.append(("warning", f"Steepest decline: <b>{slowest['RO Name']}</b> ({slowest['OMC_std']}, {slowest['District']}) — <b>{slowest['MoM_Growth']:+.1f}%</b> MoM change."))

    ms_leaders = active_df2[active_df2["MS_L3_Growth_Valid"]].nlargest(1,"MS_L3_Growth")
    if not ms_leaders.empty:
        ml = ms_leaders.iloc[0]
        insights.append(("success", f"Fastest MS growth: <b>{ml['RO Name']}</b> ({ml['OMC_std']}) — <b>{ml['MS_L3_Growth']:+.1f}%</b> vs prior L3M (avg {ml['MS_L3_Avg']:,.1f} KL/mo)."))

    hsd_leaders = active_df2[active_df2["HSD_L3_Growth_Valid"]].nlargest(1,"HSD_L3_Growth")
    if not hsd_leaders.empty:
        hl = hsd_leaders.iloc[0]
        insights.append(("success", f"Fastest HSD growth: <b>{hl['RO Name']}</b> ({hl['OMC_std']}) — <b>{hl['HSD_L3_Growth']:+.1f}%</b> vs prior L3M (avg {hl['HSD_L3_Avg']:,.1f} KL/mo)."))

    in_list = df_raw[df_raw["Is_Inactive"]]["RO Name"].tolist()
    if in_list:
        insights.append(("error", f"<b>{len(in_list)} outlet(s)</b> consistently inactive: {', '.join(in_list[:4])}{'...' if len(in_list)>4 else ''}. Field verification recommended."))

    for name in df_raw[df_raw["Is_Restart"]]["RO Name"].tolist()[:3]:
        insights.append(("warning", f"Outlet <b>{name}</b> was inactive and recently resumed operations."))

    for _, row in df_raw[df_raw["Is_Outlier"]].nlargest(3,"Total_L3_Avg").iterrows():
        insights.append(("warning", f"Outlet <b>{row['RO Name']}</b> shows abnormal avg volume ({row['Total_L3_Avg']:,.1f} KL/mo) — statistical outlier."))

    for name in df_raw[df_raw["Has_Spike"]]["RO Name"].tolist()[:3]:
        insights.append(("warning", f"Outlet <b>{name}</b> shows an unusual sales spike in the latest month."))

    q25 = active_df2["Total_L3_Avg"].quantile(0.25) if not active_df2.empty else 0
    under = active_df2[active_df2["Total_L3_Avg"] < q25]
    if not under.empty:
        insights.append(("info", f"<b>{len(under)} outlet(s)</b> below 25th percentile (&lt;{q25:,.1f} KL/mo avg). Consider targeted field intervention."))

    with st.spinner("Computing predictions..."):
        pa = []
        for _, row in active_df2.iterrows():
            p = predict_outlet(row, ms_cols, hsd_cols)
            pa.append({"RO Name": row["RO Name"], "OMC": row["OMC_std"], "District": row["District"],
                       "Pred Total": p["total_pred"], "Trend": p["trend"]})
        if pa:
            pa_df = pd.DataFrame(pa)
            tp = pa_df.nlargest(1,"Pred Total").iloc[0]
            insights.append(("success", f"Predicted top performer next month: <b>{tp['RO Name']}</b> ({tp['OMC']}, {tp['District']}) — projected <b>{tp['Pred Total']:,.1f} KL</b>."))
            inc = (pa_df["Trend"]=="Increase").sum()
            dec = (pa_df["Trend"]=="Decline").sum()
            insights.append(("info", f"Next month outlook: <b>{inc} outlets</b> predicted to grow, <b>{dec}</b> predicted to decline."))

    type_map = {"info":"info-banner","success":"success-banner","warning":"warn-banner","error":"error-banner"}
    for itype, itext in insights:
        st.markdown(f'<div class="{type_map[itype]}">{itext}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">District Performance Summary — MS & HSD (Avg)</div>', unsafe_allow_html=True)
    ds = df.groupby("District").agg(
        Outlets=("RO Name","count"),
        MS_Avg=("MS_L3_Avg","sum"), HSD_Avg=("HSD_L3_Avg","sum"), Total_Avg=("Total_L3_Avg","sum"),
    ).sort_values("Total_Avg", ascending=False).reset_index()

    fig_ds = go.Figure()
    fig_ds.add_trace(go.Bar(x=ds["District"], y=ds["MS_Avg"],  name="MS",  marker_color="#f97316",
                             text=ds["MS_Avg"].map(lambda x: f"{x:,.1f}"), textposition="outside"))
    fig_ds.add_trace(go.Bar(x=ds["District"], y=ds["HSD_Avg"], name="HSD", marker_color="#3b82f6",
                             text=ds["HSD_Avg"].map(lambda x: f"{x:,.1f}"), textposition="outside"))
    fig_ds.update_layout(**chart_layout(300, barmode="group",
                                        legend=dict(orientation="h", y=1.08),
                                        yaxis_title="Avg Monthly Volume (KL)", xaxis=dict(tickangle=-15)))
    st.plotly_chart(fig_ds, use_container_width=True)

    ds_tbl = ds.copy()
    ds_tbl.columns = ["District","Outlets","MS Avg/Mo","HSD Avg/Mo","Total Avg/Mo"]
    st.dataframe(ds_tbl.style.format({
        "MS Avg/Mo":"{:,.1f}","HSD Avg/Mo":"{:,.1f}","Total Avg/Mo":"{:,.1f}",
    }), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
