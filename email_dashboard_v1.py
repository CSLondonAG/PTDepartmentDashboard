import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

RESP_FILE  = "Responded PT.csv"
ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"
EMAIL_REC_FILE = "EmailReceivedPT.csv"

EMAIL_CHANNEL = "casesChannel"
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

resp  = load(BASE / RESP_FILE)
items = load(BASE / ITEMS_FILE)
pres  = load(BASE / PRES_FILE)

# Try to load email received file, fall back to resp if missing
try:
    email_rec = load(BASE / EMAIL_REC_FILE)
except FileNotFoundError:
    st.warning(f"Email received file not found. Using responded data instead.")
    email_rec = None

for df in (resp, items, pres):
    df.columns = df.columns.str.strip()

if email_rec is not None:
    email_rec.columns = email_rec.columns.str.strip()


# Keep full resp for total count, create filtered version for ART
resp_full = resp.copy()

# ---------------- ART (SINGLE-INSTANCE CASES ONLY) ----------------

resp["OpenedDT"] = pd.to_datetime(resp["Date/Time Opened"], errors="coerce", dayfirst=True)
resp["ReplyDT"]  = pd.to_datetime(resp["Email Message Date"], errors="coerce", dayfirst=True)

# Count before filtering
total_responded = len(resp_full)
resp = resp.groupby("Case ID").filter(lambda x: len(x) == 1)
single_instance = len(resp)
reopened_excluded = total_responded - single_instance

resp["ARTsec"] = (resp["ReplyDT"] - resp["OpenedDT"]).dt.total_seconds()
resp["Date"]   = resp["ReplyDT"].dt.date


# ---------------- AHT (ITEMS ONLY - ALL CASES) ----------------

items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

items["HandleSec"] = pd.to_numeric(items["Handle Time"], errors="coerce")

items["AssignDT"] = pd.to_datetime(
    items["Assign Date"].astype(str) + " " + items["Assign Time"].astype(str),
    errors="coerce",
    dayfirst=True
)

items["Date"] = items["AssignDT"].dt.date


# ---------------- EMAIL RECEIVED ----------------

email_rec["OpenedDT"] = pd.to_datetime(email_rec["Date/Time Opened"], errors="coerce", dayfirst=True)
email_rec["Date"] = email_rec["OpenedDT"].dt.date


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

# Calculate previous completed week (Monday-Sunday)
today = pd.Timestamp.now().date()
# Find last Sunday
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - pd.Timedelta(days=days_since_sunday if days_since_sunday > 0 else 7)
week_start = last_sunday - pd.Timedelta(days=6)  # Previous Monday

default_start = max(week_start, resp["Date"].min())
default_end = min(last_sunday, resp["Date"].max())

start, end = st.date_input(
    "Date Range",
    value=(default_start, default_end),
    help="Shows previous completed week by default"
)

resp  = resp[(resp["Date"] >= start) & (resp["Date"] <= end)]
items = items[(items["Date"] >= start) & (items["Date"] <= end)]
email_rec = email_rec[(email_rec["Date"] >= start) & (email_rec["Date"] <= end)]

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

st.markdown(f"**Period:** {start} to {end}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Total Responded", total_responded)
c2.metric("Single Response", single_instance)
c3.metric("Avg Response Time", hm(resp["ARTsec"].mean()))
c4.metric("Avg Handle Time", mmss(avg_aht))
c5.metric("Available Hours", f"{available_hours:.1f}")
c6.metric("Utilisation", f"{util:.1%}")


# ---------------- CHART ----------------

st.markdown("---")
st.subheader("Daily Emails Received vs Items Handled vs Availability")

# Emails received (from email_rec data if available, otherwise resp data)
if email_rec is not None:
    daily = email_rec.groupby("Date").size().reset_index(name="Emails_Received")
else:
    # Fallback: use responded data
    daily = resp.groupby("Date").size().reset_index(name="Emails_Received")

# Items handled (from items data - all email channel items)
items_daily = items.groupby("Date").size().reset_index(name="Items_Handled")
daily = daily.merge(items_daily, on="Date", how="outer").fillna(0)
daily["Items_Handled"] = daily["Items_Handled"].astype(int)
daily["Emails_Received"] = daily["Emails_Received"].astype(int)

# Available hours per day
def hours_for_day(day):
    ds = pd.Timestamp(day)
    de = ds + pd.Timedelta(days=1)
    iv = [clip(s, e, ds, de) for s, e in zip(pres["Start DT"], pres["End DT"])]
    iv = [x for x in iv if x]
    return sum_seconds(iv) / 3600

daily["Available_Hours"] = daily["Date"].apply(hours_for_day)

# Format date for display (e.g., "Mon 10 Feb")
daily["Date_Label"] = daily["Date"].apply(lambda x: pd.Timestamp(x).strftime("%a %d %b"))
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

    # Add legend manually with custom styling
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
