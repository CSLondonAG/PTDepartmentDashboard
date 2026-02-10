import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# PAGE
# =====================================================

st.set_page_config(layout="wide")

# =====================================================
# FILES
# =====================================================

BASE = Path(__file__).parent

ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"
ART_FILE   = "ART PT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE CSV
# =====================================================

def read_csv_safe(path):
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)


# =====================================================
# HELPERS
# =====================================================

def fmt(sec):
    if sec is None or pd.isna(sec):
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
        if iv:
            s, e = iv
            total += (e - s).total_seconds()
    return total


def safe_min(series):
    s = series.dropna()
    return s.min() if len(s) else None


def safe_max(series):
    s = series.dropna()
    return s.max() if len(s) else None


def find_col(cols, keywords):
    lower = {c.lower(): c for c in cols}
    for k in keywords:
        for lc, orig in lower.items():
            if k in lc:
                return orig
    return None


# =====================================================
# LOAD
# =====================================================

@st.cache_data
def load():
    items = read_csv_safe(BASE / ITEMS_FILE)
    pres  = read_csv_safe(BASE / PRES_FILE)
    art   = read_csv_safe(BASE / ART_FILE)

    for df in (items, pres, art):
        df.columns = df.columns.str.strip()

    return items, pres, art


items, pres, art = load()


# =====================================================
# AHT (ItemsPT)
# =====================================================

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["HandleSec"] = pd.to_numeric(items.get("Handle Time"), errors="coerce")

assign_dt = pd.to_datetime(
    items.get("Assign Date", "").astype(str) + " " + items.get("Assign Time", "").astype(str),
    errors="coerce",
    dayfirst=True
)

items["Date"] = assign_dt.dt.date


# =====================================================
# RESPONSE TIME (ART PT)
# =====================================================

open_col  = find_col(art.columns, ["opened", "received", "created"])
close_col = find_col(art.columns, ["closed", "resolved", "completed"])

if open_col and close_col:
    art[open_col]  = pd.to_datetime(art[open_col], errors="coerce", dayfirst=True)
    art[close_col] = pd.to_datetime(art[close_col], errors="coerce", dayfirst=True)
    art["ResponseSec"] = (art[close_col] - art[open_col]).dt.total_seconds()
    art["Date"] = art[open_col].dt.date
else:
    art["ResponseSec"] = pd.NA
    art["Date"] = pd.NaT


# =====================================================
# DATE RANGE (NaT-safe)
# =====================================================

min_candidates = [
    safe_min(items["Date"]),
    safe_min(art["Date"])
]

max_candidates = [
    safe_max(items["Date"]),
    safe_max(art["Date"])
]

min_candidates = [x for x in min_candidates if x is not None]
max_candidates = [x for x in max_candidates if x is not None]

if not min_candidates:
    st.error("No valid dates in dataset.")
    st.stop()

min_d = min(min_candidates)
max_d = max(max_candidates)

start, end = st.date_input("Date Range", value=(max_d - pd.Timedelta(days=6), max_d))

items = items[(items["Date"] >= start) & (items["Date"] <= end)]
art   = art[(art["Date"] >= start) & (art["Date"] <= end)]


# =====================================================
# PRESENCE → CAPACITY
# =====================================================

pres["Start DT"] = pd.to_datetime(pres.get("Start DT"), errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres.get("End DT"), errors="coerce", dayfirst=True)

ws = pd.Timestamp(start)
we = pd.Timestamp(end) + pd.Timedelta(days=1)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]

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

avg_aht = items["HandleSec"].mean()
avg_response = art["ResponseSec"].mean()

util = items["HandleSec"].sum() / available_sec if available_sec else 0
emails_hr = len(items) / (available_sec / 3600) if available_sec else 0

st.title("Email Department Performance")

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("Total Emails", f"{len(items):,}")
c2.metric("Avg Handle Time (AHT)", fmt(avg_aht))
c3.metric("Avg Response Time", fmt(avg_response))
c4.metric("Utilisation", f"{util:.1%}")
c5.metric("Emails / Available Hr", f"{emails_hr:.1f}")


# =====================================================
# DAILY TREND
# =====================================================

st.markdown("---")

daily_items = items.groupby("Date").size().reset_index(name="Volume")
daily_resp  = art.groupby("Date")["ResponseSec"].mean().reset_index(name="Response")

daily = daily_items.merge(daily_resp, on="Date", how="left")

if daily.empty:
    st.info("No data for selected period.")
    st.stop()

bars = alt.Chart(daily).mark_bar(opacity=0.15).encode(x="Date:T", y="Volume:Q")

line = alt.Chart(daily).mark_line(strokeWidth=2.5, point=True).encode(
    x="Date:T",
    y="Response:Q"
)

st.altair_chart(
    alt.layer(bars, line).resolve_scale(y="independent"),
    use_container_width=True
)
