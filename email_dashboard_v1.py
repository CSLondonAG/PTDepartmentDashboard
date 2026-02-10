import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

RESP_FILE  = "Responded PT.csv"
ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# LOAD
# =====================================================

def read_csv_safe(path):
    if not path.exists():
        st.error(f"{path.name} missing from project folder.")
        st.stop()
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)


resp  = read_csv_safe(BASE / RESP_FILE)
items = read_csv_safe(BASE / ITEMS_FILE)
pres  = read_csv_safe(BASE / PRES_FILE)

for df in (resp, items, pres):
    df.columns = df.columns.str.strip()


# =====================================================
# ===================== ART ===========================
# =====================================================

resp["OpenedDT"] = pd.to_datetime(resp["Date/Time Opened"], errors="coerce", dayfirst=True)
resp["ReplyDT"]  = pd.to_datetime(resp["Email Message Date"], errors="coerce", dayfirst=True)

# remove reopened
resp = resp.groupby("Case ID").filter(lambda x: len(x) == 1)

resp["ARTsec"] = (resp["ReplyDT"] - resp["OpenedDT"]).dt.total_seconds()
resp["Date"]   = resp["ReplyDT"].dt.date


# =====================================================
# ===================== AHT ===========================
# EXACT COLUMN (confirmed)
# =====================================================

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")

items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

items["Date"] = items["AssignDT"].dt.date


# =====================================================
# =================== PRESENCE ========================
# =====================================================

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]


def clip(s, e, ws, we):
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None


def sum_seconds(iv):
    return sum((e - s).total_seconds() for s, e in iv if iv)


# =====================================================
# DATE RANGE (single truth for ALL datasets)
# =====================================================

min_d = resp["Date"].min()
max_d = resp["Date"].max()

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

resp  = resp[(resp["Date"] >= start) & (resp["Date"] <= end)]
items = items[(items["Date"] >= start) & (items["Date"] <= end)]

start_ts = pd.Timestamp(start)
end_ts   = pd.Timestamp(end) + pd.Timedelta(days=1)


# =====================================================
# AVAILABLE HOURS
# =====================================================

intervals = [
    clip(s, e, start_ts, end_ts)
    for s, e in zip(pres["Start DT"], pres["End DT"])
]

available_sec = sum_seconds([x for x in intervals if x])
available_hours = available_sec / 3600


# =====================================================
# HANDLE TIME (same window only — FIXED)
# =====================================================

total_handle = items["HandleSec"].sum()
avg_aht = items["HandleSec"].mean()

util = total_handle / available_sec if available_sec else 0


# =====================================================
# FORMATTERS
# =====================================================

def fmt_mmss(sec):
    if pd.isna(sec) or sec == 0:
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


def fmt_hm(sec):
    if pd.isna(sec) or sec == 0:
        return "—"
    h = int(sec)//3600
    m = (int(sec)%3600)//60
    return f"{h}h {m:02}m"


# =====================================================
# METRICS
# =====================================================

st.title("Email Department Performance")

c1,c2,c3,c4,c5 = st.columns(5)

c1.metric("Replies", len(resp))
c2.metric("Avg Response Time (ART)", fmt_hm(resp["ARTsec"].mean()))
c3.metric("Avg Handle Time (AHT)", fmt_mmss(avg_aht))
c4.metric("Available Hours", f"{available_hours:.1f}")
c5.metric("Utilisation", f"{util:.1%}")


# =====================================================
# DAILY DEMAND VS CAPACITY CHART
# =====================================================

st.markdown("---")
st.subheader("Daily Demand vs Available Hours")

daily = resp.groupby("Date").size().reset_index(name="Replies")


def hours_for_day(day):
    ds = pd.Timestamp(day)
    de = ds + pd.Timedelta(days=1)

    iv = [clip(s, e, ds, de) for s, e in zip(pres["Start DT"], pres["End DT"])]
    iv = [x for x in iv if x]

    return sum_seconds(iv) / 3600


daily["AvailableHours"] = daily["Date"].apply(hours_for_day)

bars = alt.Chart(daily).mark_bar(color="#2563eb", opacity=0.6).encode(
    x=alt.X("Date:O", title="Date"),
    y="Replies:Q"
)

hours = alt.Chart(daily).mark_bar(color="#cbd5e1", opacity=0.4).encode(
    x="Date:O",
    y="AvailableHours:Q"
)

st.altair_chart(
    alt.layer(bars, hours).resolve_scale(y="independent"),
    use_container_width=True
)
