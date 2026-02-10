import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(layout="wide")

BASE = Path(__file__).parent

ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE CSV LOADER
# =====================================================

def read_csv_safe(path: Path):
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)


@st.cache_data
def load_data():
    items = read_csv_safe(BASE / ITEMS_FILE)
    pres  = read_csv_safe(BASE / PRES_FILE)

    items.columns = items.columns.str.strip()
    pres.columns  = pres.columns.str.strip()

    # ===== FORCE DATETIME (critical) =====
    for df in (items, pres):
        for c in ["Start DT", "End DT"]:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    items = items.dropna(subset=["Start DT", "End DT"])
    pres  = pres.dropna(subset=["Start DT", "End DT"])

    return items, pres


# =====================================================
# HARDENED HELPERS
# =====================================================

def fmt_mmss(sec):
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
    """Never assume tuple validity"""
    total = 0.0
    for iv in intervals:
        if not iv or len(iv) != 2:
            continue
        s, e = iv
        if pd.isna(s) or pd.isna(e):
            continue
        if e > s:
            total += (e - s).total_seconds()
    return total


# =====================================================
# LOAD
# =====================================================

items, pres = load_data()

# email only
items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["Duration"] = (items["End DT"] - items["Start DT"]).dt.total_seconds()
items = items.dropna(subset=["Duration"])

items["Date"] = items["Start DT"].dt.date


# =====================================================
# DATE FILTER
# =====================================================

min_d = items["Date"].min()
max_d = items["Date"].max()

start, end = st.sidebar.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d),
    min_value=min_d,
    max_value=max_d
)

items = items[(items["Date"] >= start) & (items["Date"] <= end)]

ws = pd.Timestamp(start)
we = pd.Timestamp(end) + pd.Timedelta(days=1)


# =====================================================
# CAPACITY (presence)
# =====================================================

pres = pres[
    (pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)) &
    (pres["Start DT"] < we) &
    (pres["End DT"] > ws)
]

intervals = [
    x for x in (
        clip(s, e, ws, we)
        for s, e in zip(pres["Start DT"], pres["End DT"])
    )
    if x
]

available_seconds = sum_seconds(intervals)


# =====================================================
# WORKLOAD METRICS
# =====================================================

handle_seconds = items["Duration"].sum()

utilisation = handle_seconds / available_seconds if available_seconds else 0

emails_per_hour = (
    len(items) / (available_seconds / 3600)
    if available_seconds else 0
)


# =====================================================
# DAILY AGG
# =====================================================

daily = (
    items.groupby("Date")
    .agg(
        Volume=("Duration", "size"),
        AHT=("Duration", "mean")
    )
    .reset_index()
)


def daily_capacity(day):
    d_start = pd.Timestamp(day)
    d_end   = d_start + pd.Timedelta(days=1)

    iv = [
        x for x in (
            clip(s, e, d_start, d_end)
            for s, e in zip(pres["Start DT"], pres["End DT"])
        )
        if x
    ]

    return sum_seconds(iv) / 60


daily["AvailMin"] = daily["Date"].apply(daily_capacity)


# =====================================================
# UI
# =====================================================

st.title("Email Department Performance")

c1, c2, c3, c4 = st.columns(4)

c1.metric("Total Emails", f"{len(items):,}")
c2.metric("Avg Handle Time", fmt_mmss(items["Duration"].mean()))
c3.metric("Utilisation", f"{utilisation:.1%}")
c4.metric("Emails / Available Hr", f"{emails_per_hour:.1f}")

st.markdown("---")


# =====================================================
# DAILY CAPACITY VS DEMAND
# =====================================================

bars = alt.Chart(daily).mark_bar(opacity=0.3).encode(
    x="Date:T",
    y="AvailMin:Q"
)

line = alt.Chart(daily).mark_line(point=True).encode(
    x="Date:T",
    y="Volume:Q"
)

st.altair_chart(
    alt.layer(bars, line).resolve_scale(y="independent").properties(height=350),
    use_container_width=True
)


# =====================================================
# HOURLY COVERAGE
# =====================================================

st.markdown("---")
st.subheader("Hourly Coverage")

sel = st.date_input("Select day", value=end, min_value=min_d, max_value=max_d)

d_start = pd.Timestamp(sel)
d_end   = d_start + pd.Timedelta(days=1)

hours = pd.date_range(d_start, d_end, freq="h", inclusive="left")

rows = []

for h in hours:
    h_end = h + pd.Timedelta(hours=1)

    vol = len(items[(items["Start DT"] >= h) & (items["Start DT"] < h_end)])

    iv = [
        x for x in (
            clip(s, e, h, h_end)
            for s, e in zip(pres["Start DT"], pres["End DT"])
        )
        if x
    ]

    avail = sum_seconds(iv) / 60

    rows.append([h, vol, avail])

hourly = pd.DataFrame(rows, columns=["Hour", "Volume", "AvailMin"])

bars = alt.Chart(hourly).mark_bar(opacity=0.3).encode(
    x="Hour:T",
    y="AvailMin:Q"
)

line = alt.Chart(hourly).mark_line(point=True).encode(
    x="Hour:T",
    y="Volume:Q"
)

st.altair_chart(
    alt.layer(bars, line).resolve_scale(y="independent").properties(height=300),
    use_container_width=True
)
