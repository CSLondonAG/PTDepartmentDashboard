import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

RESP_FILE  = "Responded PT.csv"
ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# ---------------- LOAD ----------------

def load(path):
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)

resp  = load(BASE / RESP_FILE)
items = load(BASE / ITEMS_FILE)
pres  = load(BASE / PRES_FILE)

for df in (resp, items, pres):
    df.columns = df.columns.str.strip()


# ---------------- ART (SINGLE-INSTANCE CASES ONLY) ----------------

resp["OpenedDT"] = pd.to_datetime(resp["Date/Time Opened"], errors="coerce", dayfirst=True)
resp["ReplyDT"]  = pd.to_datetime(resp["Email Message Date"], errors="coerce", dayfirst=True)

# Filter to single instances only (excludes reopened cases)
single_instance_count = len(resp)
resp = resp.groupby("Case ID").filter(lambda x: len(x) == 1)
reopened_excluded = single_instance_count - len(resp)

resp["ARTsec"] = (resp["ReplyDT"] - resp["OpenedDT"]).dt.total_seconds()
resp["Date"]   = resp["ReplyDT"].dt.date


# ---------------- AHT (ITEMS ONLY) ----------------

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")

items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

items["Date"] = items["AssignDT"].dt.date


# ---------------- PRESENCE ----------------

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]


def clip(s, e, ws, we):
    """Clip interval (s,e) to window (ws,we)"""
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None

def sum_seconds(iv):
    """Sum duration of intervals in seconds"""
    return sum((e - s).total_seconds() for s, e in iv)


# ---------------- DATE FILTER ----------------

start, end = st.date_input(
    "Date Range",
    value=(resp["Date"].min(), resp["Date"].max())
)

resp  = resp[(resp["Date"] >= start) & (resp["Date"] <= end)]
items = items[(items["Date"] >= start) & (items["Date"] <= end)]

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)


# ---------------- HOURS + UTILIZATION ----------------

intervals = [clip(s, e, start_ts, end_ts) for s, e in zip(pres["Start DT"], pres["End DT"])]
intervals = [x for x in intervals if x]

available_sec = sum_seconds(intervals)
available_hours = available_sec / 3600

total_handle = items["HandleSec"].sum()
avg_aht = items["HandleSec"].mean()

util = total_handle / available_sec if available_sec else 0


# ---------------- FORMAT ----------------

def mmss(sec):
    """Format seconds as MM:SS"""
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"

def hm(sec):
    """Format seconds as Xh XXm"""
    if pd.isna(sec):
        return "—"
    h = int(sec) // 3600
    m = (int(sec) % 3600) // 60
    return f"{h}h {m:02}m"


# ---------------- METRICS ----------------

st.title("Email Department Performance")

st.markdown(f"**Analysis Period:** {start} to {end} | **Note:** Reopened cases excluded: {reopened_excluded}")

c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("Replies (Single-Instance)", len(resp))
c2.metric("Avg Response Time (ART)", hm(resp["ARTsec"].mean()))
c3.metric("Avg Handle Time (AHT)", mmss(avg_aht))
c4.metric("Available Hours", f"{available_hours:.1f}")
c5.metric("Utilisation", f"{util:.1%}")


# ---------------- CHART ----------------

st.markdown("---")
st.subheader("Daily Emails Received vs Items Handled vs Availability")

# Emails received (from resp data - single instances only)
daily = resp.groupby("Date").size().reset_index(name="Emails_Received")

# Items handled (from items data - all email channel items)
items_daily = items.groupby("Date").size().reset_index(name="Items_Handled")
daily = daily.merge(items_daily, on="Date", how="outer").fillna(0)
daily["Items_Handled"] = daily["Items_Handled"].astype(int)

# Available hours per day
def hours_for_day(day):
    ds = pd.Timestamp(day)
    de = ds + pd.Timedelta(days=1)
    iv = [clip(s, e, ds, de) for s, e in zip(pres["Start DT"], pres["End DT"])]
    iv = [x for x in iv if x]
    return sum_seconds(iv) / 3600

daily["Available_Hours"] = daily["Date"].apply(hours_for_day)

# Emails received bar (blue)
emails_bar = alt.Chart(daily).mark_bar(color="#3b82f6", opacity=0.85).encode(
    x=alt.X("Date:O", title="Date"),
    y=alt.Y("Emails_Received:Q", title="Count", axis=alt.Axis(orient="left", labelColor="#3b82f6")),
    tooltip=["Date:O", "Emails_Received:Q"]
).properties(
    title="Emails Received"
)

# Items handled bar (green)
handled_bar = alt.Chart(daily).mark_bar(color="#10b981", opacity=0.7).encode(
    x=alt.X("Date:O"),
    y=alt.Y("Items_Handled:Q"),
    tooltip=["Date:O", "Items_Handled:Q"]
).properties(
    title="Items Handled"
)

# Available hours line (indigo)
availability_line = alt.Chart(daily).mark_line(color="#6366f1", size=3).encode(
    x=alt.X("Date:O"),
    y=alt.Y("Available_Hours:Q", title="Available Hours", axis=alt.Axis(orient="right", labelColor="#6366f1")),
    tooltip=["Date:O", alt.Tooltip("Available_Hours:Q", format=".1f")]
).properties(
    title="Agent Availability"
)

chart = alt.layer(emails_bar, handled_bar, availability_line).resolve_scale(
    y="independent"
).properties(
    width=900,
    height=450
)

st.altair_chart(chart, use_container_width=True)

# Daily detail table
st.subheader("Daily Breakdown")
daily_display = daily.copy()
daily_display["Emails_Received"] = daily_display["Emails_Received"].astype(int)
daily_display["Items_Handled"] = daily_display["Items_Handled"].astype(int)
daily_display["Available_Hours"] = daily_display["Available_Hours"].round(1)
daily_display = daily_display.rename(columns={
    "Date": "Date",
    "Emails_Received": "📧 Emails Received",
    "Items_Handled": "✓ Items Handled",
    "Available_Hours": "⏰ Available Hours"
})
st.dataframe(
    daily_display.sort_values("Date", ascending=False),
    use_container_width=True,
    hide_index=True
)
