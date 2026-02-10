import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")

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
# AUTO COLUMN DETECTION (CRITICAL FIX)
# =====================================================

def find_col(cols, keywords):
    cols_lower = {c.lower(): c for c in cols}
    for k in keywords:
        for c_lower, orig in cols_lower.items():
            if k in c_lower:
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
# -------- AHT (ItemsPT)
# =====================================================

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")

assign_dt = pd.to_datetime(
    items["Assign Date"] + " " + items["Assign Time"],
    errors="coerce",
    dayfirst=True
)

items["Date"] = assign_dt.dt.date


# =====================================================
# -------- RESPONSE TIME (ART PT)  (SAFE)
# =====================================================

open_col = find_col(art.columns, ["opened", "received", "created"])
close_col = find_col(art.columns, ["closed", "resolved", "completed"])

if open_col and close_col:

    art[open_col] = pd.to_datetime(art[open_col], errors="coerce", dayfirst=True)
    art[close_col] = pd.to_datetime(art[close_col], errors="coerce", dayfirst=True)

    art["ResponseSec"] = (art[close_col] - art[open_col]).dt.total_seconds()
    art["Date"] = art[open_col].dt.date

else:
    art["ResponseSec"] = pd.NA
    art["Date"] = pd.NaT
    st.warning("ART PT.csv does not contain recognizable opened/resolved timestamp columns.")


# =====================================================
# DATE RANGE
# =====================================================

min_d = min(items["Date"].min(), art["Date"].min())
max_d = max(items["Date"].max(), art["Date"].max())

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

items = items[(items["Date"] >= start) & (items["Date"] <= end)]
art   = art[(art["Date"] >= start) & (art["Date"] <= end)]


# =====================================================
# PRESENCE → CAPACITY
# =====================================================

def clip(s,e,ws,we):
    s2,e2=max(s,ws),min(e,we)
    return (s2,e2) if e2>s2 else None

def sum_sec(iv):
    return sum((e-s).total_seconds() for s,e in iv)

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)

ws = pd.Timestamp(start)
we = pd.Timestamp(end) + pd.Timedelta(days=1)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]

intervals = [
    x for x in (
        clip(s,e,ws,we)
        for s,e in zip(pres["Start DT"], pres["End DT"])
    ) if x
]

available_sec = sum_sec(intervals)


# =====================================================
# METRICS
# =====================================================

avg_aht = items["HandleSec"].mean()
avg_resp = art["ResponseSec"].mean()

util = items["HandleSec"].sum() / available_sec if available_sec else 0
emails_hr = len(items)/(available_sec/3600) if available_sec else 0

def fmt(sec):
    if pd.isna(sec): return "—"
    m,s = divmod(int(sec),60)
    return f"{m:02}:{s:02}"

st.title("Email Department Performance")

c1,c2,c3,c4,c5 = st.columns(5)

c1.metric("Total Emails", f"{len(items):,}")
c2.metric("Avg Handle Time", fmt(avg_aht))
c3.metric("Avg Response Time", fmt(avg_resp))
c4.metric("Utilisation", f"{util:.1%}")
c5.metric("Emails / Available Hr", f"{emails_hr:.1f}")


# =====================================================
# DAILY TREND
# =====================================================

st.markdown("---")

daily_items = items.groupby("Date").size().reset_index(name="Volume")
daily_resp  = art.groupby("Date")["ResponseSec"].mean().reset_index(name="Response")

daily = daily_items.merge(daily_resp, on="Date", how="left")

bars = alt.Chart(daily).mark_bar(opacity=0.15, color="#cbd5e1").encode(
    x="Date:T",
    y="Volume:Q"
)

line = alt.Chart(daily).mark_line(
    strokeWidth=2.5,
    color="#2563eb",
    point=True
).encode(
    x="Date:T",
    y="Response:Q"
)

st.altair_chart(
    alt.layer(bars,line).resolve_scale(y="independent"),
    use_container_width=True
)
