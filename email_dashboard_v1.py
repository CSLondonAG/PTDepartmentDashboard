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
# IO
# =====================================================

def read_csv_safe(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except Exception:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)


def to_bool(x):
    # Handles True/False, "TRUE"/"FALSE", "Yes"/"No", 1/0
    if pd.isna(x):
        return pd.NA
    if isinstance(x, bool):
        return x
    s = str(x).strip().lower()
    if s in {"true", "t", "yes", "y", "1"}:
        return True
    if s in {"false", "f", "no", "n", "0"}:
        return False
    return pd.NA


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
# HELPERS
# =====================================================

def clip(s, e, ws, we):
    if pd.isna(s) or pd.isna(e):
        return None
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None


def sum_seconds(intervals):
    total = 0.0
    for iv in intervals:
        if not iv:
            continue
        s, e = iv
        total += (e - s).total_seconds()
    return total


# =====================================================
# LOAD
# =====================================================

resp  = read_csv_safe(BASE / RESP_FILE)
items = read_csv_safe(BASE / ITEMS_FILE)
pres  = read_csv_safe(BASE / PRES_FILE)

for df in (resp, items, pres):
    df.columns = df.columns.str.strip()

# =====================================================
# RESPONDED PT (ART)
# =====================================================

# Required columns (confirmed by your file screenshots + prior listing):
# Case ID, Date/Time Opened, Email Message Date, Is Incoming, Date/Time Closed
resp["CaseID"] = resp["Case ID"].astype(str).str.strip()

resp["OpenedDT"] = pd.to_datetime(resp["Date/Time Opened"], errors="coerce", dayfirst=True)
resp["EmailDT"]  = pd.to_datetime(resp["Email Message Date"], errors="coerce", dayfirst=True)

# Agent replied = outbound email message
resp["IsIncomingBool"] = resp["Is Incoming"].apply(to_bool)

# Keep outbound rows for ART (agent replies)
agent_msgs = resp[resp["IsIncomingBool"] == False].copy()

# If the export is missing/dirty Is Incoming and yields no outbound rows, fall back to "all rows are agent messages"
if agent_msgs.empty:
    agent_msgs = resp.copy()

# Per-reply lag from OpenedDT (captures multi-response)
agent_msgs["ReplyLagSec"] = (agent_msgs["EmailDT"] - agent_msgs["OpenedDT"]).dt.total_seconds()

# Per-case first response time (FRT)
first_reply = (
    agent_msgs.dropna(subset=["OpenedDT", "EmailDT"])
    .groupby("CaseID", as_index=False)
    .agg(
        OpenedDT=("OpenedDT", "min"),
        FirstReplyDT=("EmailDT", "min"),
        Replies=("EmailDT", "count"),
        AvgReplyLagSec=("ReplyLagSec", "mean"),
    )
)
first_reply["FRTSec"] = (first_reply["FirstReplyDT"] - first_reply["OpenedDT"]).dt.total_seconds()
first_reply["OpenedDate"] = first_reply["OpenedDT"].dt.date


# =====================================================
# ITEMS PT (AHT)
# =====================================================

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["CaseID"] = items["Work Item: Name"].astype(str).str.strip()

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")

items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True
)
items["AssignDate"] = items["AssignDT"].dt.date

# Handle time at case level (sums multi-assignments; mean handle per assignment also available)
handle_by_case = (
    items.groupby("CaseID", as_index=False)
    .agg(
        TotalHandleSec=("HandleSec", "sum"),
        AvgHandleSec=("HandleSec", "mean"),
        AssignDT=("AssignDT", "min"),
    )
)


# =====================================================
# JOIN (case metrics)
# =====================================================

case = first_reply.merge(handle_by_case, on="CaseID", how="left")

# =====================================================
# PRESENCE (capacity)
# =====================================================

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)
pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)].copy()


# =====================================================
# DATE RANGE (use OpenedDate = when the case entered the system)
# =====================================================

all_dates = pd.Series(case["OpenedDate"]).dropna()
if all_dates.empty:
    st.error("No valid dates found in Responded PT.csv after parsing.")
    st.stop()

min_d = all_dates.min()
max_d = all_dates.max()

start, end = st.date_input("Date Range", value=(max_d - pd.Timedelta(days=6), max_d))

case_f = case[(case["OpenedDate"] >= start) & (case["OpenedDate"] <= end)].copy()
agent_msgs_f = agent_msgs[
    (agent_msgs["OpenedDT"].dt.date >= start) &
    (agent_msgs["OpenedDT"].dt.date <= end)
].copy()

start_ts = pd.Timestamp(start)
end_ts   = pd.Timestamp(end) + pd.Timedelta(days=1)


# =====================================================
# CAPACITY CALC (in selected range)
# =====================================================

intervals = [
    clip(s, e, start_ts, end_ts)
    for s, e in zip(pres["Start DT"], pres["End DT"])
]
intervals = [x for x in intervals if x]
available_sec = sum_seconds(intervals)


# =====================================================
# METRICS
# =====================================================

avg_frt = case_f["FRTSec"].mean()
avg_reply_lag = agent_msgs_f["ReplyLagSec"].mean()  # multi-response aware
avg_aht = case_f["AvgHandleSec"].mean()             # effort per assignment (average)
total_handle = case_f["TotalHandleSec"].sum()

util = (total_handle / available_sec) if available_sec else 0

# =====================================================
# UI
# =====================================================

st.title("Email Department Performance")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Cases (Opened)", f"{len(case_f):,}")
c2.metric("Replies (Agent)", f"{len(agent_msgs_f):,}")
c3.metric("Avg First Response (Opened → 1st Reply)", fmt_hm(avg_frt))
c4.metric("Avg Reply Lag (Opened → Each Reply)", fmt_hm(avg_reply_lag))
c5.metric("Avg Handle Time (AHT)", fmt_mmss(avg_aht))

c6, c7 = st.columns(2)
c6.metric("Total Handle Time", fmt_hm(total_handle))
c7.metric("Utilisation", f"{util:.1%}")


# =====================================================
# DAILY VOLUME (agent replies)
# =====================================================

st.markdown("---")
st.subheader("Daily Agent Replies")

agent_msgs_f["ReplyDate"] = agent_msgs_f["EmailDT"].dt.date
daily_replies = agent_msgs_f.groupby("ReplyDate").size().reset_index(name="Replies")

chart = alt.Chart(daily_replies).mark_bar(opacity=0.6, color="#2563eb").encode(
    x=alt.X("ReplyDate:T", title=""),
    y=alt.Y("Replies:Q", title="Replies"),
    tooltip=["ReplyDate:T", "Replies:Q"]
)

st.altair_chart(chart, use_container_width=True)
