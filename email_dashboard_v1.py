import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(layout="wide")


# =====================================================
# DESIGN SYSTEM (Phase 1 + 2 implemented)
# =====================================================

st.markdown("""
<style>

/* ---------- Typography ---------- */

html, body, [class*="css"]  {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

h1 {
    font-size: 32px !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em;
    color: #0f172a;
    margin-bottom: 8px !important;
}

.caption-date {
    font-size: 13px;
    color: #64748b;
    margin-bottom: 32px;
}

/* ---------- Metric Cards ---------- */

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
    font-size: 28px !important;
    font-weight: 700 !important;
    color: #0f172a !important;
}

[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #64748b !important;
}

/* ---------- Dividers ---------- */

hr {
    border: none;
    border-top: 1px solid rgba(15,23,42,0.12);
    margin: 64px 0 32px 0;
}

/* ---------- Chart container ---------- */

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
# SAFE LOADER
# =====================================================

def read_csv_safe(path):
    try:
        return pd.read_csv(path, encoding="cp1252")
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t")


@st.cache_data
def load():
    items = read_csv_safe(BASE / ITEMS_FILE)
    pres = read_csv_safe(BASE / PRES_FILE)

    for df in (items, pres):
        df.columns = df.columns.str.strip()
        for c in ["Start DT", "End DT"]:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    items = items.dropna(subset=["Start DT","End DT"])
    pres = pres.dropna(subset=["Start DT","End DT"])

    return items, pres


# =====================================================
# HELPERS (bulletproof intervals)
# =====================================================

def clip(s,e,ws,we):
    if pd.isna(s) or pd.isna(e):
        return None
    s2,e2 = max(s,ws), min(e,we)
    return (s2,e2) if e2>s2 else None


def sum_seconds(intervals):
    total = 0
    for iv in intervals:
        if not iv: continue
        s,e = iv
        total += (e-s).total_seconds()
    return total


def fmt(sec):
    if pd.isna(sec): return "—"
    m,s = divmod(int(sec),60)
    return f"{m:02}:{s:02}"


# =====================================================
# LOAD
# =====================================================

items, pres = load()

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL]
items["Duration"] = (items["End DT"] - items["Start DT"]).dt.total_seconds()
items["Date"] = items["Start DT"].dt.date


# =====================================================
# DATE RANGE (context always visible)
# =====================================================

min_d = items["Date"].min()
max_d = items["Date"].max()

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

ws = pd.Timestamp(start)
we = pd.Timestamp(end) + pd.Timedelta(days=1)

items = items[(items["Date"]>=start)&(items["Date"]<=end)]


# =====================================================
# HEADER
# =====================================================

st.title("Email Department Performance")
st.markdown(
    f"<div class='caption-date'>Viewing: {start:%b %d, %Y} – {end:%b %d, %Y}</div>",
    unsafe_allow_html=True
)


# =====================================================
# CAPACITY
# =====================================================

pres = pres[
    pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)
]

intervals = [
    x for x in (
        clip(s,e,ws,we)
        for s,e in zip(pres["Start DT"], pres["End DT"])
    ) if x
]

available_sec = sum_seconds(intervals)


# =====================================================
# METRICS
# =====================================================

handle_sec = items["Duration"].sum()
util = handle_sec / available_sec if available_sec else 0
emails_hr = len(items)/(available_sec/3600) if available_sec else 0

c1,c2,c3,c4 = st.columns(4)

c1.metric("Total Emails", f"{len(items):,}")
c2.metric("Avg Handle Time", fmt(items["Duration"].mean()))
c3.metric("Utilisation", f"{util:.1%}")
c4.metric("Emails / Available Hr", f"{emails_hr:.1f}")

st.markdown("<hr>", unsafe_allow_html=True)


# =====================================================
# DAILY CHART
# =====================================================

daily = (
    items.groupby("Date")
    .agg(Volume=("Duration","size"))
    .reset_index()
)

if daily.empty:
    st.info("No emails in this period.")
    st.stop()

bars = alt.Chart(daily).mark_bar(
    opacity=0.15,
    color="#cbd5e1"
).encode(
    x="Date:T",
    y="Volume:Q"
)

line = alt.Chart(daily).mark_line(
    strokeWidth=2.5,
    color="#2563eb",
    point=True
).encode(
    x="Date:T",
    y="Volume:Q"
)

chart = alt.layer(bars,line).configure_axis(
    labelColor="#64748b",
    titleColor="#64748b",
    gridOpacity=0.08,
    domainOpacity=0.15
).configure_view(strokeWidth=0)

st.altair_chart(chart, use_container_width=True)


# =====================================================
# HOURLY CHART
# =====================================================

st.markdown("<hr>", unsafe_allow_html=True)
st.subheader("Hourly Coverage")

sel = st.date_input("Select day", value=end)

d_start = pd.Timestamp(sel)
d_end = d_start + pd.Timedelta(days=1)

hours = pd.date_range(d_start,d_end,freq="h",inclusive="left")

rows=[]
for h in hours:
    vol=len(items[(items["Start DT"]>=h)&(items["Start DT"]<h+pd.Timedelta(hours=1))])
    rows.append([h,vol])

hourly=pd.DataFrame(rows,columns=["Hour","Volume"])

bars=alt.Chart(hourly).mark_bar(opacity=0.15,color="#cbd5e1").encode(x="Hour:T",y="Volume:Q")
line=alt.Chart(hourly).mark_line(strokeWidth=2.5,color="#2563eb",point=True).encode(x="Hour:T",y="Volume:Q")

st.altair_chart(alt.layer(bars,line).configure_view(strokeWidth=0), use_container_width=True)
