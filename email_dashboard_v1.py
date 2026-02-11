import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

EMAIL_REC_FILE = "EmailReceivedPT.csv"
PRES_FILE  = "PresencePT.csv"
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
email_rec.columns = email_rec.columns.str.strip()


# ---------------- PARSE EMAIL RECEIVED DATA ----------------

email_rec["OpenedDT"] = pd.to_datetime(email_rec["Date/Time Opened"], errors="coerce", dayfirst=True)
email_rec["CompletedDT"] = pd.to_datetime(email_rec["Completion Date"], errors="coerce", dayfirst=True)

email_rec["Date_Opened"] = email_rec["OpenedDT"].dt.date
email_rec["Date_Completed"] = email_rec["CompletedDT"].dt.date


# Load presence for availability calculation
try:
    pres = load(BASE / PRES_FILE)
    pres.columns = pres.columns.str.strip()
    pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
    pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)
    pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]
except:
    pres = None


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
# Find last Sunday
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - pd.Timedelta(days=days_since_sunday if days_since_sunday > 0 else 7)
week_start = last_sunday - pd.Timedelta(days=6)  # Previous Monday

default_start = max(week_start, email_rec["Date_Opened"].min())
default_end = min(last_sunday, email_rec["Date_Opened"].max())

start, end = st.date_input(
    "Date Range",
    value=(default_start, default_end),
    help="Shows previous completed week by default"
)

email_rec_filtered = email_rec[(email_rec["Date_Opened"] >= start) & (email_rec["Date_Opened"] <= end)].copy()

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)


# ---------------- METRICS CALCULATIONS ----------------

# Emails received: count of non-null Date/Time Opened
total_received = email_rec_filtered["OpenedDT"].notna().sum()

# Handled: count of non-null Completion Date
total_handled = email_rec_filtered["CompletedDT"].notna().sum()

# Single response cases: those with completion date (handled)
single_instance = total_handled

# Unhandled (not yet completed)
unhandled = total_received - total_handled

# Calculate ART: Average Response Time (from Opened to Completed, only for completed items)
email_rec_filtered["ARTsec"] = (email_rec_filtered["CompletedDT"] - email_rec_filtered["OpenedDT"]).dt.total_seconds()
completed_mask = email_rec_filtered["CompletedDT"].notna()
avg_art = email_rec_filtered[completed_mask]["ARTsec"].mean()

# Calculate AHT: Average Handle Time (same as ART for this data)
avg_aht = avg_art


# Available hours and utilization
available_hours = 0
util = 0

if pres is not None:
    intervals = [clip(s, e, start_ts, end_ts) for s, e in zip(pres["Start DT"], pres["End DT"])]
    intervals = [x for x in intervals if x]
    available_sec = sum_seconds(intervals)
    available_hours = available_sec / 3600
    util = (total_handled * (avg_aht / 3600)) / available_hours if available_hours > 0 else 0


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

st.markdown(f"**Period:** {start} to {end}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Total Received", total_received)
c2.metric("Total Handled", total_handled)
c3.metric("Avg Response Time", hm(avg_art))
c4.metric("Avg Handle Time", mmss(avg_aht))
c5.metric("Available Hours", f"{available_hours:.1f}")
c6.metric("Utilisation", f"{util:.1%}")


# ---------------- CHART ----------------

st.markdown("---")
st.subheader("Daily Emails Received vs Items Handled vs Availability")

# Count emails received by Date Opened
daily = email_rec_filtered.groupby("Date_Opened").size().reset_index(name="Emails_Received")
daily = daily.rename(columns={"Date_Opened": "Date"})

# Count handled by Date Completed
handled_daily = email_rec_filtered[email_rec_filtered["CompletedDT"].notna()].groupby("Date_Completed").size().reset_index(name="Items_Handled")
handled_daily = handled_daily.rename(columns={"Date_Completed": "Date"})

# Merge the two
daily = daily.merge(handled_daily, on="Date", how="outer").fillna(0)
daily["Items_Handled"] = daily["Items_Handled"].astype(int)
daily["Emails_Received"] = daily["Emails_Received"].astype(int)

# Available hours per day
def hours_for_day(day):
    if pres is None:
        return 0
    ds = pd.Timestamp(day)
    de = ds + pd.Timedelta(days=1)
    iv = [clip(s, e, ds, de) for s, e in zip(pres["Start DT"], pres["End DT"])]
    iv = [x for x in iv if x]
    return sum_seconds(iv) / 3600

daily["Available_Hours"] = daily["Date"].apply(hours_for_day)
daily = daily.sort_values("Date").reset_index(drop=True)

if len(daily) > 0:
    # Available hours bar (light gray background)
    availability_bar = alt.Chart(daily).mark_bar(color="#cbd5e1", opacity=0.6).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Available_Hours:Q", title="Hours", axis=alt.Axis(orient="left")),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Available_Hours:Q", format=".1f", title="Available Hours")]
    )

    # Emails received line (blue)
    emails_line = alt.Chart(daily).mark_line(color="#3b82f6", size=3, point=True).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Emails_Received:Q", title="Count", axis=alt.Axis(orient="right")),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Emails_Received:Q", title="Emails Received")]
    )

    # Items handled line (green)
    handled_line = alt.Chart(daily).mark_line(color="#10b981", size=3, point=True).encode(
        x=alt.X("Date:O", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Items_Handled:Q"),
        tooltip=[alt.Tooltip("Date:O", format="%a %d %b"), alt.Tooltip("Items_Handled:Q", title="Items Handled")]
    )

    chart = alt.layer(availability_bar, emails_line, handled_line).resolve_scale(
        y="independent"
    ).properties(
        height=400
    )

    st.altair_chart(chart, use_container_width=True)
    
    # Create legend below chart
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("🔹 <span style='color:#cbd5e1'>**Available Hours (Bar)**</span>", unsafe_allow_html=True)
    with col2:
        st.markdown("📈 <span style='color:#3b82f6'>**Emails Received (Line)**</span>", unsafe_allow_html=True)
    with col3:
        st.markdown("✓ <span style='color:#10b981'>**Items Handled (Line)**</span>", unsafe_allow_html=True)
else:
    st.warning("No data available for the selected date range")

# Daily detail table
st.subheader("Daily Breakdown")
daily_display = daily.copy()
daily_display = daily_display.rename(columns={
    "Date": "Date",
    "Emails_Received": "📧 Emails Received",
    "Items_Handled": "✓ Items Handled",
    "Available_Hours": "⏰ Available Hours"
})
daily_display = daily_display[["Date", "📧 Emails Received", "✓ Items Handled", "⏰ Available Hours"]]
daily_display["⏰ Available Hours"] = daily_display["⏰ Available Hours"].round(1)
st.dataframe(
    daily_display.sort_values("Date", ascending=False),
    use_container_width=True,
    hide_index=True
)

