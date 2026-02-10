import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(layout="wide")

BASE = Path(__file__).parent

RESP_FILE = "Responded PT.csv"
PRES_FILE = "PresencePT.csv"

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
# FORMATTERS
# =====================================================

def fmt_mmss(sec):
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


def fmt_hm(sec):
    if pd.isna(sec):
        return "—"
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    return f"{h}h {m:02}m"


# =====================================================
# LOAD
# =====================================================

resp = read_csv_safe(BASE / RESP_FILE)
pres = read_csv_safe(BASE / PRES_FILE)

resp.columns = resp.columns.str.strip()
pres.columns = pres.columns.str.strip()


# =====================================================
# EXPECTED COLUMNS IN Responded PT
# (adjust names if slightly different)
# =====================================================

# timestamps
resp["ReceivedDT"] = pd.to_datetime(resp["Date/Time Received"], errors="coerce", dayfirst=True)
resp["RespondedDT"] = pd.to_datetime(resp["Date/Time Responded"], errors="coerce", dayfirst=True)

# handle time already provided
resp["HandleSec"] = pd.to_numeric(resp["Handle Time"], errors="coerce")

resp["ResponseSec"] = (resp["RespondedDT"] - resp["ReceivedDT"]).dt.total_seconds()

resp["Date"] = resp["RespondedDT"].dt.date


# =====================================================
# PRESENCE
# =====================================================

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]


# =====================================================
# DATE RANGE
# =====================================================

min_d = resp["Date"].min()
max_d = resp["Date"].max()

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

resp = resp[(resp["Date"] >= start) & (resp["Date"] <= end)]

start_ts = pd.Timestamp(start)
end_ts   = pd.Timestamp(end) + pd.Timedelta(days=1)


# =====================================================
# CAPACITY
# =====================================================

def clip(s,e,ws,we):
    s2,e2=max(s,ws),min(e,we)
    return (s2,e2) if e2>s2 else None

def sum_seconds(iv):
    return sum((e-s).total_seconds() for s,e in iv if iv)

intervals = [
    clip(s,e,start_ts,end_ts)
    for s,e in zip(pres["Start DT"], pres["End DT"])
]
intervals = [x for x in intervals if x]

available_sec = sum_seconds(intervals)


# =====================================================
# METRICS
# =====================================================

avg_aht  = resp["HandleSec"].mean()
avg_resp = resp["ResponseSec"].mean()

emails_hr = len(resp) / (available_sec/3600) if available_sec else 0
util = resp["HandleSec"].sum() / available_sec if available_sec else 0

st.title("Email Department Performance")

c1,c2,c3,c4 = st.columns(4)

c1.metric("Replies Handled", f"{len(resp):,}")
c2.metric("Avg Handle Time (AHT)", fmt_mmss(avg_aht))
c3.metric("Avg Response Time (ART)", fmt_hm(avg_resp))
c4.metric("Utilisation", f"{util:.1%}")


# =====================================================
# DAILY VOLUME TREND
# =====================================================

st.markdown("---")
st.subheader("Daily Replies")

daily = resp.groupby("Date").size().reset_index(name="Replies")

chart = alt.Chart(daily).mark_bar(color="#2563eb", opacity=0.6).encode(
    x="Date:T",
    y="Replies:Q"
)

st.altair_chart(chart, use_container_width=True)
