import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path
import numpy as np

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

EMAIL_REC_FILE = "EmailReceivedPT.csv"
ITEMS_FILE = "ItemsPT.csv"
PRES_FILE = "PresencePT.csv"

AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# ---------------- LOAD ----------------

def load(path):
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except:
        try:
            return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)
        except:
            try:
                return pd.read_csv(path, encoding="utf-8", low_memory=False)
            except:
                return pd.read_csv(path, encoding="latin-1", low_memory=False)

email_rec = load(BASE / EMAIL_REC_FILE)
items = load(BASE / ITEMS_FILE)
pres = load(BASE / PRES_FILE)

for df in (email_rec, items, pres):
    df.columns = df.columns.str.strip()


# ---------------- EMAIL RECEIVED DATA ----------------

email_rec["OpenedDT"] = pd.to_datetime(email_rec["Date/Time Opened"], errors="coerce", dayfirst=True)
email_rec["CompletedDT"] = pd.to_datetime(email_rec["Completion Date"], errors="coerce", dayfirst=True)

email_rec["Date_Opened"] = email_rec["OpenedDT"].dt.date
email_rec["Date_Completed"] = email_rec["CompletedDT"].dt.date

# Calculate response time (only for completed items)
email_rec["ResponseTimeSec"] = (email_rec["CompletedDT"] - email_rec["OpenedDT"]).dt.total_seconds()


# ---------------- ITEMS DATA (WORK ITEMS HANDLED) ----------------

# Parse items dates
items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

items["CloseDT"] = pd.to_datetime(
    items["Close Date"].astype(str) + " " + items["Close Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")
items["Date_Assigned"] = items["AssignDT"].dt.date
items["Date_Closed"] = items["CloseDT"].dt.date

# Filter to casesChannel only
items = items[items["Service Channel: Developer Name"] == "casesChannel"].copy()


# ---------------- PRESENCE DATA ----------------

pres["StartDT"] = pd.to_datetime(
    pres["Status Start Date"].astype(str) + " " + pres["Status Start Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

pres["EndDT"] = pd.to_datetime(
    pres["Status End Date"].astype(str) + " " + pres["Status End Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)].copy()


def clip(s, e, ws, we):
    """Clip interval (s,e) to window (ws,we)"""
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None

def sum_seconds(iv):
    """Sum duration of intervals in seconds"""
    return sum((e - s).total_seconds() for s, e in iv)


# ---------------- DATE FILTER ----------------

# Calculate previous completed week (Monday-Sunday)
today = pd.Timestamp.now().date()
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - pd.Timedelta(days=days_since_sunday if days_since_sunday > 0 else 7)
week_start = last_sunday - pd.Timedelta(days=6)

default_start = max(week_start, email_rec["Date_Opened"].min())
default_end = min(last_sunday, email_rec["Date_Opened"].max())

start, end = st.date_input(
    "Date Range",
    value=(default_start, default_end),
    help="Shows previous completed week by default"
)

# Filter data by date range
email_rec_period = email_rec[
    (email_rec["Date_Opened"] >= start) & (email_rec["Date_Opened"] <= end)
].copy()

items_period = items[
    (items["Date_Closed"] >= start) & (items["Date_Closed"] <= end)
].copy()

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)


# ---------------- METRICS ----------------

# Total received: count of emails with opened date in period
total_received = email_rec_period["OpenedDT"].notna().sum()

# Total handled: count of emails with completion date in period
total_handled = email_rec_period["CompletedDT"].notna().sum()

# ART: average response time for completed emails in period
completed_emails = email_rec_period[email_rec_period["CompletedDT"].notna()]
avg_art = completed_emails["ResponseTimeSec"].mean() if len(completed_emails) > 0 else 0

# AHT: average handle time from items (for casesChannel only, closed in period)
avg_aht = items_period["HandleSec"].mean() if len(items_period) > 0 else 0

# Available hours in period
intervals = [clip(s, e, start_ts, end_ts) for s, e in zip(pres["StartDT"], pres["EndDT"])]
intervals = [x for x in intervals if x]
available_sec = sum_seconds(intervals)
available_hours = available_sec / 3600

# Utilization: total handle time of items closed in period / available seconds
total_handle_sec = items_period["HandleSec"].sum()
util = (total_handle_sec / available_sec) if available_sec > 0 else 0


# ---------------- FORMAT ----------------

def mmss(sec):
    """Format seconds as MM:SS"""
    if pd.isna(sec) or sec == 0:
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"

def hm(sec):
    """Format seconds as Xh XXm"""
    if pd.isna(sec) or sec == 0:
        return "—"
    h = int(sec) // 3600
    m = (int(sec) % 3600) // 60
    return f"{h}h {m:02}m"


# ---------------- DISPLAY METRICS ----------------

st.title("Email Department Performance")
st.markdown(f"**Period:** {start} to {end}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Emails Received", f"{total_received:,}")
c2.metric("Emails Handled", f"{total_handled:,}")
c3.metric("Avg Response Time", hm(avg_art))
c4.metric("Avg Handle Time", mmss(avg_aht))
c5.metric("Available Hours", f"{available_hours:.1f}")
c6.metric("Utilisation", f"{util:.1%}")


# ---------------- CHART ----------------

st.markdown("---")
st.subheader("Daily Emails Received vs Items Handled vs Availability")

# Daily emails received (by opened date)
daily_received = email_rec_period.groupby("Date_Opened").size().reset_index(name="Emails_Received")
daily_received = daily_received.rename(columns={"Date_Opened": "Date"})

# Daily items handled (by closed date)
daily_handled = items_period.groupby("Date_Closed").size().reset_index(name="Items_Handled")
daily_handled = daily_handled.rename(columns={"Date_Closed": "Date"})

# Merge
daily = daily_received.merge(daily_handled, on="Date", how="outer").fillna(0)
daily["Items_Handled"] = daily["Items_Handled"].astype(int)
daily["Emails_Received"] = daily["Emails_Received"].astype(int)

# Available hours per day
def hours_for_day(day):
    ds = pd.Timestamp(day)
    de = ds + pd.Timedelta(days=1)
    iv = [clip(s, e, ds, de) for s, e in zip(pres["StartDT"], pres["EndDT"])]
    iv = [x for x in intervals if x]
    return sum_seconds(iv) / 3600

daily["Available_Hours"] = daily["Date"].apply(hours_for_day)
daily = daily.sort_values("Date").reset_index(drop=True)

if len(daily) > 0:
    # Available hours bar
    availability_bar = alt.Chart(daily).mark_bar(color="#cbd5e1", opacity=0.6).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Available_Hours:Q", title="Hours", axis=alt.Axis(orient="left")),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Available_Hours:Q", format=".1f")]
    )

    # Emails received line
    emails_line = alt.Chart(daily).mark_line(color="#3b82f6", size=3, point=True).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Emails_Received:Q", title="Count", axis=alt.Axis(orient="right")),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Emails_Received:Q")]
    )

    # Items handled line
    handled_line = alt.Chart(daily).mark_line(color="#10b981", size=3, point=True).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Items_Handled:Q", title=""),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Items_Handled:Q")]
    )

    chart = alt.layer(availability_bar, emails_line, handled_line).resolve_scale(
        y="independent"
    ).properties(height=400)

    st.altair_chart(chart, use_container_width=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("🔹 <span style='color:#cbd5e1'>**Available Hours (Bar)**</span>", unsafe_allow_html=True)
    with col2:
        st.markdown("📈 <span style='color:#3b82f6'>**Emails Received (Line)**</span>", unsafe_allow_html=True)
    with col3:
        st.markdown("✓ <span style='color:#10b981'>**Items Handled (Line)**</span>", unsafe_allow_html=True)
else:
    st.warning("No data available for the selected date range")


# Daily breakdown table
st.subheader("Daily Breakdown")
daily_display = daily.copy()
daily_display["Available_Hours"] = daily_display["Available_Hours"].round(1)
daily_display = daily_display.rename(columns={
    "Date": "Date",
    "Emails_Received": "📧 Received",
    "Items_Handled": "✓ Handled",
    "Available_Hours": "⏰ Hours"
})
st.dataframe(
    daily_display.sort_values("Date", ascending=False),
    use_container_width=True,
    hide_index=True
)
