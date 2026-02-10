import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# PAGE
# =====================================================

st.set_page_config(layout="wide")


# =====================================================
# DESIGN SYSTEM (CSS)
# =====================================================

st.markdown("""
<style>
html, body, [class*="css"]  {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

/* Title */
h1 {
    font-size: 32px !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em;
    color: #0f172a;
    margin-bottom: 6px !important;
}

.caption-date {
    font-size: 13px;
    color: #64748b;
    margin-bottom: 32px;
}

/* Metric cards */
[data-testid="metric-container"] {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px;
    transition: all 150ms ease;
}

[data-testid="metric-container"]:hover {
    background: #ffffff;
    box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
}

[data-testid="stMetricValue"] {
    font-size: 26px !important;
    font-weight: 700 !important;
    color: #0f172a !important;
}

[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    color: #64748b !important;
}

/* Dividers */
hr {
    border: none;
    border-top: 1px solid rgba(15,23,42,0.12);
    margin: 64px 0 32px 0;
}

/* Charts */
[data-testid="stVegaLiteChart"] {
    border-radius: 12px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# FILES
# =====================================================

BASE = Path(__file__).parent
ITEMS_FILE = "ItemsPT.csv"
PRES_FILE = "PresencePT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE CSV LOADER
# =====================================================

def read_csv_safe(path):
    try:
        return pd.read_csv(path, encoding="cp1252")
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t")


@st.cache_data
def load():
    items = read_csv_safe(BASE / ITEMS_FILE)
    pres  = read_csv_safe(BASE / PRES_FILE)

    for df in (items, pres):
        df.columns = df.columns.str.strip()
        for c in ["Start DT", "End DT"]:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    items = items.dropna(subset=["Start DT", "End DT"])
    pres  = pres.dropna(subset=["Start DT", "End DT"])

    return items, pres


# =====================================================
# HELPERS (bulletproof)
# =====================================================

def fmt(sec):
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


def clip(s, e, ws, we):
    if pd.isna(s) or pd.isna(e):
        return None
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None


def sum_seconds(intervals):
    total = 0
    for iv in intervals:
        if not iv:
            continue
        s, e = iv
        total += (e - s).total_seconds()
    return total


# =====================================================
# LOAD
# =====================================================

items, pres = load()

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

# -----------------------------------------------------
# TIME METRICS
# -----------------------------------------------------

items["HandleSec"] = (items["End DT"] - items["Start DT"]).dt.total_seconds()

# NEW — response time (customer wait)
items["ResponseSec"] = items["HandleSec"]

items = items.dropna(subset=["HandleSec"])

items["Date"] = items["Start DT"].dt.date


# =====================================================
# DATE RANGE (visible, not sidebar)
# =====================================================

min_d = items["Date"].min()
max_d = items["Date"].max()

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

ws = pd.Timestamp(start)
we = pd.Timestamp(end) + pd.Timedelta(days=1)

items = items[(items["Date"] >= start) & (items["Date"] <= end)]


# =====================================================
# HEADER
# =====================================================

st.title("Email Department Performance")
st.markdown(
    f"<div class='caption-date'>Viewing: {start:%b %d, %Y} – {end:%b %d, %Y}</div>",
    unsafe_allow_html=True
)


# =====================================================
# CAPACITY (presence)
# =====================================================

pres = pres[
    pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)
]

intervals = [
    x for x in (
        clip(s, e, ws, we)
        for s, e in zip(pres["Start DT"], pres["End DT"])
    ) if x
]

available_sec = sum_seconds(intervals)


# =====================================================
# METRICS
# =====================================================

handle_sec = items["HandleSec"].sum()
util = handle_sec / available_sec if available_sec else 0
emails_hr = len(items) / (available_sec / 3600) if available_sec else 0

avg_handle = items["HandleSec"].mean()
avg_response = items["ResponseSec"].mean()

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("Total Emails", f"{len(items):,}")
c2.metric("Avg Handle Time", fmt(avg_handle))
c3.metric("Avg Response Time", fmt(avg_response))
c4.metric("Utilisation", f"{util:.1%}")
c5.metric("Emails / Available Hr", f"{emails_hr:.1f}")


# =====================================================
# DAILY AGGREGATION
# =====================================================

st.markdown("<hr>", unsafe_allow_html=True)

daily = (
    items.groupby("Date")
    .agg(
        Volume=("HandleSec", "size"),
        Response=("ResponseSec", "mean")
    )
    .reset_index()
)

if daily.empty:
    st.info("No emails in this period.")
    st.stop()


# =====================================================
# CHARTS (styled)
# =====================================================

alt.themes.enable("none")

bars = alt.Chart(daily).mark_bar(
    opacity=0.15,
    color="#cbd5e1"
).encode(
    x="Date:T",
    y="Volume:Q"
)

vol_line = alt.Chart(daily).mark_line(
    strokeWidth=2.5,
    color="#2563eb",
    point=True
).encode(
    x="Date:T",
    y="Volume:Q"
)

resp_line = alt.Chart(daily).mark_line(
    strokeWidth=2.5,
    color="#ef4444",
    point=True
).encode(
    x="Date:T",
    y="Response:Q",
    tooltip=[
        alt.Tooltip("Date:T"),
        alt.Tooltip("Response:Q", title="Avg Response (sec)", format=".0f")
    ]
)

chart = alt.layer(bars, vol_line, resp_line).resolve_scale(y="independent")\
    .configure_axis(
        labelColor="#64748b",
        titleColor="#64748b",
        gridOpacity=0.08,
        domainOpacity=0.15
    ).configure_view(strokeWidth=0)

st.altair_chart(chart, use_container_width=True)
