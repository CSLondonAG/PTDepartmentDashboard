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
BUSINESS_START_HOUR = 7
BUSINESS_END_HOUR = 22

st.markdown(
    """
    <style>
      .stApp {background-color: #f7fff9;}
      div[data-testid="stMetricValue"] {color: #166534;}
      .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {color: #166534;}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load(path):
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except Exception:
        try:
            return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)
        except Exception:
            try:
                return pd.read_csv(path, encoding="utf-8", low_memory=False)
            except Exception:
                return pd.read_csv(path, encoding="latin-1", low_memory=False)


def business_seconds_between(start_dt, end_dt, start_hour=BUSINESS_START_HOUR, end_hour=BUSINESS_END_HOUR):
    """Business-time seconds between two timestamps, weekends included."""
    if pd.isna(start_dt) or pd.isna(end_dt) or end_dt <= start_dt:
        return np.nan

    start_day = start_dt.normalize()
    end_day = end_dt.normalize()
    total = 0.0

    current_day = start_day
    while current_day <= end_day:
        window_start = current_day + pd.Timedelta(hours=start_hour)
        window_end = current_day + pd.Timedelta(hours=end_hour)

        interval_start = max(start_dt, window_start)
        interval_end = min(end_dt, window_end)

        if interval_end > interval_start:
            total += (interval_end - interval_start).total_seconds()

        current_day += pd.Timedelta(days=1)

    return total


def clip(start_dt, end_dt, window_start, window_end):
    """Clip interval (start_dt, end_dt) to window (window_start, window_end)."""
    if pd.isna(start_dt) or pd.isna(end_dt):
        return None
    start_clipped, end_clipped = max(start_dt, window_start), min(end_dt, window_end)
    return (start_clipped, end_clipped) if end_clipped > start_clipped else None


def sum_seconds(intervals):
    return sum((e - s).total_seconds() for s, e in intervals)


def mmss(sec):
    if pd.isna(sec) or sec == 0:
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


def hm(sec):
    if pd.isna(sec) or sec == 0:
        return "—"
    h = int(sec) // 3600
    m = (int(sec) % 3600) // 60
    return f"{h}h {m:02}m"


# ---------------- LOAD & PREP ----------------

email_rec = load(BASE / EMAIL_REC_FILE)
items = load(BASE / ITEMS_FILE)
pres = load(BASE / PRES_FILE)

for df in (email_rec, items, pres):
    df.columns = df.columns.str.strip()

email_rec["OpenedDT"] = pd.to_datetime(email_rec["Date/Time Opened"], errors="coerce", dayfirst=True)
email_rec["CompletedDT"] = pd.to_datetime(email_rec["Completion Date"], errors="coerce", dayfirst=True)
email_rec["Date_Opened"] = email_rec["OpenedDT"].dt.date
email_rec["Date_Completed"] = email_rec["CompletedDT"].dt.date
email_rec["TargetResponseHours"] = pd.to_numeric(email_rec["Target Response (Hours)"], errors="coerce")

items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True,
)
items["CloseDT"] = pd.to_datetime(
    items["Close Date"].astype(str) + " " + items["Close Time"].astype(str),
    errors="coerce",
    dayfirst=True,
)
items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")
items["Date_Closed"] = items["CloseDT"].dt.date
items = items[items["Service Channel: Developer Name"] == "casesChannel"].copy()

pres["StartDT"] = pd.to_datetime(
    pres["Status Start Date"].astype(str) + " " + pres["Status Start Time"].astype(str),
    errors="coerce",
    dayfirst=True,
)
pres["EndDT"] = pd.to_datetime(
    pres["Status End Date"].astype(str) + " " + pres["Status End Time"].astype(str),
    errors="coerce",
    dayfirst=True,
)
pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)].copy()


# ---------------- CONTROLS ----------------

st.title("Email Department Performance")

_, refresh_col = st.columns([4, 1])
with refresh_col:
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Previous completed week (Mon-Sun)
today = pd.Timestamp.now().date()
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - pd.Timedelta(days=days_since_sunday if days_since_sunday > 0 else 7)
week_start = last_sunday - pd.Timedelta(days=6)

default_start = max(week_start, email_rec["Date_Opened"].min())
default_end = min(last_sunday, email_rec["Date_Opened"].max())

start, end = st.date_input(
    "Date Range",
    value=(default_start, default_end),
    help="Shows previous completed week by default",
)

with st.expander("Filters"):
    queue_options = sorted(items["Queue: Name"].dropna().unique().tolist())
    selected_queues = st.multiselect("Queue", queue_options, default=queue_options)

    origin_options = sorted(email_rec["Case Origin"].dropna().unique().tolist())
    selected_origins = st.multiselect("Case Origin", origin_options, default=origin_options)

    priority_options = sorted(email_rec["Priority"].dropna().unique().tolist())
    selected_priorities = st.multiselect("Priority", priority_options, default=priority_options)

    milestone_options = sorted(email_rec["Milestone"].dropna().unique().tolist())
    selected_milestones = st.multiselect("Milestone", milestone_options, default=milestone_options)


# ---------------- FILTERED DATA ----------------

email_rec_period = email_rec[
    (email_rec["Date_Opened"] >= start)
    & (email_rec["Date_Opened"] <= end)
    & (email_rec["Case Origin"].isin(selected_origins))
    & (email_rec["Priority"].isin(selected_priorities))
    & (email_rec["Milestone"].isin(selected_milestones))
].copy()

items_period = items[
    (items["Date_Closed"] >= start)
    & (items["Date_Closed"] <= end)
    & (items["Queue: Name"].isin(selected_queues))
].copy()

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)


# ---------------- METRICS ----------------

total_received = email_rec_period["OpenedDT"].notna().sum()
total_handled = items_period["CloseDT"].notna().sum()

completed_emails = email_rec_period[email_rec_period["CompletedDT"].notna()].copy()
if len(completed_emails) > 0:
    completed_emails["ResponseTimeBusinessSec"] = completed_emails.apply(
        lambda r: business_seconds_between(r["OpenedDT"], r["CompletedDT"]), axis=1
    )
    avg_art = completed_emails["ResponseTimeBusinessSec"].mean()
else:
    completed_emails["ResponseTimeBusinessSec"] = pd.Series(dtype=float)
    avg_art = 0

avg_aht = items_period["HandleSec"].mean() if len(items_period) > 0 else 0

intervals = [clip(s, e, start_ts, end_ts) for s, e in zip(pres["StartDT"], pres["EndDT"])]
intervals = [x for x in intervals if x]
available_sec = sum_seconds(intervals)
available_hours = available_sec / 3600

total_handle_sec = items_period["HandleSec"].sum()
util = (total_handle_sec / available_sec) if available_sec > 0 else 0

email_invalid_open = email_rec_period["OpenedDT"].isna().sum()
email_invalid_complete = email_rec_period["CompletedDT"].isna().sum()
items_invalid_close = items_period["CloseDT"].isna().sum()

sla_df = completed_emails[completed_emails["TargetResponseHours"].notna()].copy()
if len(sla_df) > 0:
    sla_df["ResponseBusinessHours"] = sla_df["ResponseTimeBusinessSec"] / 3600
    sla_met = (sla_df["ResponseBusinessHours"] <= sla_df["TargetResponseHours"]).sum()
    sla_missed = (sla_df["ResponseBusinessHours"] > sla_df["TargetResponseHours"]).sum()
else:
    sla_met = 0
    sla_missed = 0

backlog = email_rec[
    (email_rec["OpenedDT"] <= end_ts)
    & (email_rec["CompletedDT"].isna() | (email_rec["CompletedDT"] > end_ts))
].copy()
backlog_count = len(backlog)

if backlog_count > 0:
    backlog["OpenBusinessSecToEnd"] = backlog["OpenedDT"].apply(lambda dt: business_seconds_between(dt, end_ts))
else:
    backlog["OpenBusinessSecToEnd"] = pd.Series(dtype=float)

backlog_overdue = backlog[
    backlog["TargetResponseHours"].notna()
    & ((backlog["OpenBusinessSecToEnd"] / 3600) > backlog["TargetResponseHours"])
]
overdue_open_count = len(backlog_overdue)

business_age_hours = backlog["OpenBusinessSecToEnd"] / 3600
aging_bins = [0, 4, 24, 72, np.inf]
aging_labels = ["0-4h", "4-24h", "1-3d", "3d+"]
backlog["AgingBucket"] = pd.cut(business_age_hours, bins=aging_bins, labels=aging_labels, right=False)
aging_summary = backlog["AgingBucket"].value_counts().reindex(aging_labels, fill_value=0).reset_index()
aging_summary.columns = ["Bucket", "Count"]


# ---------------- DISPLAY ----------------

st.markdown(f"**Period:** {start} to {end}")

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Emails Received", f"{total_received:,}")
c2.metric("Work Items Handled", f"{total_handled:,}")
c3.metric("Avg Response Time (BH)", hm(avg_art))
c4.metric("Avg Handle Time", mmss(avg_aht))
c5.metric("Available Hours", f"{available_hours:.1f}")
c6.metric("Utilisation", f"{util:.1%}")

c7, c8, c9, c10 = st.columns(4)
c7.metric("SLA Met", f"{sla_met:,}")
c8.metric("SLA Missed", f"{sla_missed:,}")
c9.metric("Backlog (Open)", f"{backlog_count:,}")
c10.metric("Overdue Open Emails", f"{overdue_open_count:,}")

st.markdown("---")
st.subheader("Data Quality")
dq1, dq2, dq3 = st.columns(3)
dq1.metric("Invalid Opened Timestamps", f"{email_invalid_open:,}")
dq2.metric("Invalid Completion Timestamps", f"{email_invalid_complete:,}")
dq3.metric("Invalid Item Close Timestamps", f"{items_invalid_close:,}")

st.markdown("---")
st.subheader("Daily Emails Received vs Items Handled vs Availability")

daily_received = email_rec_period.groupby("Date_Opened").size().reset_index(name="Emails_Received")
daily_received = daily_received.rename(columns={"Date_Opened": "Date"})

daily_handled = items_period.groupby("Date_Closed").size().reset_index(name="Items_Handled")
daily_handled = daily_handled.rename(columns={"Date_Closed": "Date"})

daily = daily_received.merge(daily_handled, on="Date", how="outer").fillna(0)
daily["Date"] = pd.to_datetime(daily["Date"], errors="coerce")
daily = daily.dropna(subset=["Date"]).copy()
daily["Items_Handled"] = daily["Items_Handled"].astype(int)
daily["Emails_Received"] = daily["Emails_Received"].astype(int)


def hours_for_day(day_ts):
    ds = pd.Timestamp(day_ts).normalize()
    de = ds + pd.Timedelta(days=1)
    iv = [clip(s, e, ds, de) for s, e in zip(pres["StartDT"], pres["EndDT"])]
    iv = [x for x in iv if x]
    return sum_seconds(iv) / 3600


if len(daily) > 0:
    daily["Available_Hours"] = daily["Date"].apply(hours_for_day)
    daily = daily.sort_values("Date").reset_index(drop=True)

    availability_bar = alt.Chart(daily).mark_bar(color="#bbf7d0", opacity=0.8).encode(
        x=alt.X("Date:T", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Available_Hours:Q", title="Hours", axis=alt.Axis(orient="left")),
        tooltip=[alt.Tooltip("Date:T", format="%a %d %b"), alt.Tooltip("Available_Hours:Q", format=".1f")],
    )

    emails_line = alt.Chart(daily).mark_line(color="#15803d", size=3, point=True).encode(
        x=alt.X("Date:T", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Emails_Received:Q", title="Count", axis=alt.Axis(orient="right")),
        tooltip=[alt.Tooltip("Date:T", format="%a %d %b"), alt.Tooltip("Emails_Received:Q")],
    )

    handled_line = alt.Chart(daily).mark_line(color="#22c55e", size=3, point=True).encode(
        x=alt.X("Date:T", title="", axis=alt.Axis(labelAngle=45, format="%a %d %b")),
        y=alt.Y("Items_Handled:Q", title=""),
        tooltip=[alt.Tooltip("Date:T", format="%a %d %b"), alt.Tooltip("Items_Handled:Q")],
    )

    chart = alt.layer(availability_bar, emails_line, handled_line).resolve_scale(y="independent").properties(height=400)
    st.altair_chart(chart, use_container_width=True)
else:
    st.warning("No data available for the selected date range")

st.subheader("Day-of-Week Pattern")
if len(daily) > 0:
    dow = daily.copy()
    dow["DoW"] = dow["Date"].dt.day_name()
    ordered_dow = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    dow = dow.groupby("DoW", as_index=False)[["Emails_Received", "Items_Handled"]].mean()
    dow["DoW"] = pd.Categorical(dow["DoW"], categories=ordered_dow, ordered=True)
    dow = dow.sort_values("DoW")

    dow_long = dow.melt(
        id_vars="DoW",
        value_vars=["Emails_Received", "Items_Handled"],
        var_name="Metric",
        value_name="Average",
    )

    dow_chart = alt.Chart(dow_long).mark_bar().encode(
        x=alt.X("DoW:O", title="Day of Week"),
        y=alt.Y("Average:Q", title="Average Count"),
        color=alt.Color("Metric:N", scale=alt.Scale(range=["#15803d", "#22c55e"])),
        xOffset="Metric:N",
        tooltip=["DoW", "Metric", alt.Tooltip("Average:Q", format=".2f")],
    ).properties(height=300)

    st.altair_chart(dow_chart, use_container_width=True)

st.subheader("SLA and Backlog Aging")
col_a, col_b = st.columns(2)
with col_a:
    sla_chart_df = pd.DataFrame({"Status": ["Met", "Missed"], "Count": [sla_met, sla_missed]})
    sla_chart = alt.Chart(sla_chart_df).mark_bar().encode(
        x="Status:N",
        y="Count:Q",
        color=alt.Color("Status:N", scale=alt.Scale(range=["#22c55e", "#dc2626"])),
        tooltip=["Status", "Count"],
    )
    st.altair_chart(sla_chart, use_container_width=True)

with col_b:
    aging_chart = alt.Chart(aging_summary).mark_bar(color="#16a34a").encode(
        x=alt.X("Bucket:N", title="Business-hour Aging Bucket"),
        y=alt.Y("Count:Q", title="Open Email Count"),
        tooltip=["Bucket", "Count"],
    )
    st.altair_chart(aging_chart, use_container_width=True)

st.subheader("Daily Breakdown")
daily_display = daily.copy()
if len(daily_display) > 0:
    daily_display["Date"] = daily_display["Date"].dt.date
    daily_display["Available_Hours"] = daily_display["Available_Hours"].round(1)

daily_display = daily_display.rename(
    columns={
        "Date": "Date",
        "Emails_Received": "📧 Received",
        "Items_Handled": "✓ Handled",
        "Available_Hours": "⏰ Hours",
    }
)
st.dataframe(daily_display.sort_values("Date", ascending=False), use_container_width=True, hide_index=True)
