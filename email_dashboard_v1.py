import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Email Performance Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================
# DESIGN SYSTEM
# =====================================================

st.markdown("""
<style>
/* Base & Typography */
* {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
}

h1 {
    font-size: 32px !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em !important;
    line-height: 1.2 !important;
    color: #0f172a !important;
    margin-bottom: 8px !important;
}

.date-context {
    font-size: 14px;
    color: #64748b;
    margin-bottom: 32px;
    font-weight: 500;
}

h2, h3 {
    font-size: 18px !important;
    font-weight: 600 !important;
    letter-spacing: -0.01em !important;
    color: #0f172a !important;
    margin-top: 48px !important;
    margin-bottom: 16px !important;
}

/* Metric Cards - Premium Feel */
[data-testid="stMetricValue"] {
    font-size: 28px !important;
    font-weight: 700 !important;
    color: #0f172a !important;
    letter-spacing: -0.01em !important;
}

[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #64748b !important;
    text-transform: uppercase;
    letter-spacing: 0.03em;
}

[data-testid="metric-container"] {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px 20px 16px 20px;
    transition: all 150ms ease;
}

[data-testid="metric-container"]:hover {
    background: #ffffff;
    box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    transform: translateY(-1px);
}

/* Section Dividers */
hr {
    border: none;
    border-top: 1px solid rgba(15, 23, 42, 0.1);
    margin: 64px 0 40px 0;
}

/* Chart Containers */
[data-testid="stVegaLiteChart"] {
    border-radius: 12px;
    padding: 16px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #f8fafc;
    border-right: 1px solid #e2e8f0;
}

section[data-testid="stSidebar"] label {
    font-size: 13px !important;
    font-weight: 600 !important;
    color: #0f172a !important;
}

/* Remove Streamlit branding elements */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

BASE = Path(__file__).parent

ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"
ART_FILE   = "ART PT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE CSV LOADER
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
    """Format seconds as MM:SS for AHT"""
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


def fmt_hm(sec):
    """Format seconds as Hh MMm for response time"""
    if pd.isna(sec):
        return "—"
    sec = int(sec)
    hours = sec // 3600
    minutes = (sec % 3600) // 60
    return f"{hours}h {minutes:02}m"


# =====================================================
# TIME HELPERS
# =====================================================

def clip(s, e, ws, we):
    """Clip interval to window"""
    if pd.isna(s) or pd.isna(e):
        return None
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None


def sum_seconds(intervals):
    """Sum seconds from list of (start, end) tuples"""
    if not intervals:
        return 0
    return sum((e - s).total_seconds() for s, e in intervals)


# =====================================================
# LOAD DATA
# =====================================================

@st.cache_data
def load_all_data():
    items = read_csv_safe(BASE / ITEMS_FILE)
    pres  = read_csv_safe(BASE / PRES_FILE)
    art   = read_csv_safe(BASE / ART_FILE)
    
    for df in (items, pres, art):
        df.columns = df.columns.str.strip()
    
    return items, pres, art


items, pres, art = load_all_data()


# =====================================================
# PREPARE ITEMS (AHT)
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
# PREPARE ART (RESPONSE TIME)
# =====================================================

art["OpenedDT"] = pd.to_datetime(art["Date/Time Opened"], errors="coerce", dayfirst=True)
art["ClosedDT"] = pd.to_datetime(art["Date/Time Closed"], errors="coerce", dayfirst=True)
art["ResponseSec"] = (art["ClosedDT"] - art["OpenedDT"]).dt.total_seconds()
art["Date"] = art["OpenedDT"].dt.date


# =====================================================
# PREPARE PRESENCE
# =====================================================

pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)


# =====================================================
# DATE RANGE FILTER (SIDEBAR)
# =====================================================

all_dates = pd.concat([items["Date"].dropna(), art["Date"].dropna()])

if all_dates.empty:
    st.error("No valid timestamps in dataset.")
    st.stop()

min_d = all_dates.min()
max_d = all_dates.max()

st.sidebar.header("Filters")

start, end = st.sidebar.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d),
    min_value=min_d,
    max_value=max_d
)

# Filter datasets
items_filtered = items[(items["Date"] >= start) & (items["Date"] <= end)].copy()
art_filtered   = art[(art["Date"] >= start) & (art["Date"] <= end)].copy()

start_ts = pd.Timestamp(start)
end_ts   = pd.Timestamp(end) + pd.Timedelta(days=1)


# =====================================================
# CALCULATE CAPACITY
# =====================================================

pres_filtered = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]

intervals = [
    x for x in (
        clip(s, e, start_ts, end_ts)
        for s, e in zip(pres_filtered["Start DT"], pres_filtered["End DT"])
    ) if x
]

available_sec = sum_seconds(intervals)


# =====================================================
# CALCULATE DAILY CAPACITY
# =====================================================

def get_daily_capacity(date):
    """Get available agent hours for a single day"""
    d_start = pd.Timestamp(date)
    d_end   = d_start + pd.Timedelta(days=1)
    
    day_intervals = [
        x for x in (
            clip(s, e, d_start, d_end)
            for s, e in zip(pres_filtered["Start DT"], pres_filtered["End DT"])
        ) if x
    ]
    
    return sum_seconds(day_intervals) / 3600  # Convert to hours


# =====================================================
# METRICS
# =====================================================

avg_aht  = items_filtered["HandleSec"].mean()
avg_resp = art_filtered["ResponseSec"].mean()
util = items_filtered["HandleSec"].sum() / available_sec if available_sec else 0
emails_hr = len(items_filtered) / (available_sec / 3600) if available_sec else 0


# =====================================================
# HEADER
# =====================================================

st.title("Email Department Performance")
st.markdown(
    f'<p class="date-context">Viewing: {start.strftime("%b %d, %Y")} – {end.strftime("%b %d, %Y")}</p>',
    unsafe_allow_html=True
)


# =====================================================
# METRICS ROW
# =====================================================

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Emails Received", f"{len(art_filtered):,}")
c2.metric("Emails Handled", f"{len(items_filtered):,}")
c3.metric("Avg Handle Time", fmt_mmss(avg_aht))
c4.metric("Avg Response Time", fmt_hm(avg_resp))
c5.metric("Utilisation", f"{util:.1%}")
c6.metric("Emails / Hour", f"{emails_hr:.1f}")


# =====================================================
# DAILY VOLUME VS CAPACITY CHART
# =====================================================

st.markdown("---")
st.subheader("Daily Volume vs Agent Capacity")

# Prepare daily data - EMAILS HANDLED
daily_handled = items_filtered.groupby("Date").size().reset_index(name="Handled")

# Prepare daily data - EMAILS RECEIVED (from ART)
daily_received = art_filtered.groupby("Date").size().reset_index(name="Received")

# Get capacity for each day
date_range = pd.date_range(start, end, freq='D')
daily_capacity = pd.DataFrame({
    'Date': [d.date() for d in date_range],
    'Capacity_Hours': [get_daily_capacity(d.date()) for d in date_range]
})

# Merge all metrics
daily = daily_capacity.merge(daily_handled, on='Date', how='left')
daily = daily.merge(daily_received, on='Date', how='left')
daily['Handled'] = daily['Handled'].fillna(0)
daily['Received'] = daily['Received'].fillna(0)

# Calculate backlog indicator
daily['Backlog'] = daily['Received'] - daily['Handled']
daily['Cumulative_Backlog'] = daily['Backlog'].cumsum()

# Capacity bars (solid background showing agent hours)
capacity_bars = alt.Chart(daily).mark_bar(
    color='#94a3b8',
    opacity=0.25
).encode(
    x=alt.X('Date:O', title='Date', axis=alt.Axis(labelAngle=-45, labelFontSize=12, labelColor='#64748b', format='%b %d')),
    y=alt.Y('Capacity_Hours:Q', title='Agent Hours Available', axis=alt.Axis(labelFontSize=12, labelColor='#64748b', titleColor='#0f172a', titleFontWeight=600)),
    tooltip=[
        alt.Tooltip('Date:T', title='Date', format='%b %d, %Y'),
        alt.Tooltip('Capacity_Hours:Q', title='Agent Hours', format='.1f')
    ]
)

# Emails Received line (demand - blue solid)
received_line = alt.Chart(daily).mark_line(
    strokeWidth=3,
    color='#2563eb',
    point=alt.OverlayMarkDef(
        size=80,
        filled=True,
        color='#2563eb'
    )
).encode(
    x=alt.X('Date:O'),
    y=alt.Y('Received:Q', title='Email Volume', axis=alt.Axis(labelFontSize=12, labelColor='#64748b', titleColor='#0f172a', titleFontWeight=600)),
    tooltip=[
        alt.Tooltip('Date:T', title='Date', format='%b %d, %Y'),
        alt.Tooltip('Received:Q', title='Emails Received'),
        alt.Tooltip('Handled:Q', title='Emails Handled'),
        alt.Tooltip('Backlog:Q', title='Daily Backlog', format='+d')
    ]
)

# Emails Handled line (throughput - green solid)
handled_line = alt.Chart(daily).mark_line(
    strokeWidth=3,
    color='#10b981',
    point=alt.OverlayMarkDef(
        size=70,
        filled=True,
        color='#10b981'
    )
).encode(
    x=alt.X('Date:O'),
    y=alt.Y('Handled:Q'),
    tooltip=[
        alt.Tooltip('Date:T', title='Date', format='%b %d, %Y'),
        alt.Tooltip('Handled:Q', title='Emails Handled'),
        alt.Tooltip('Received:Q', title='Emails Received')
    ]
)

# Layer with independent Y-axes
daily_chart = alt.layer(
    capacity_bars,
    received_line,
    handled_line
).resolve_scale(
    y='independent'
).properties(
    height=400
).configure_view(
    strokeWidth=0
)

st.altair_chart(daily_chart, use_container_width=True)

# Add legend
st.markdown(
    '<div style="display: flex; gap: 24px; margin-top: -8px; margin-bottom: 16px; font-size: 13px; color: #64748b;">'
    '<span><span style="color: #94a3b8; font-weight: 700;">█</span> Agent Capacity (hours)</span>'
    '<span><span style="color: #2563eb; font-weight: 700;">●</span> Emails Received</span>'
    '<span><span style="color: #10b981; font-weight: 700;">●</span> Emails Handled</span>'
    '</div>',
    unsafe_allow_html=True
)

# Add capacity insight with proper logic
avg_capacity = daily['Capacity_Hours'].mean()
avg_received = daily['Received'].mean()
avg_handled = daily['Handled'].mean()
total_backlog = daily['Cumulative_Backlog'].iloc[-1]

# Calculate theoretical handling capacity (hours * emails per hour)
if avg_capacity > 0 and emails_hr > 0:
    theoretical_capacity = avg_capacity * emails_hr  # How many emails we COULD handle
    demand_vs_capacity_pct = (avg_received / theoretical_capacity) * 100
    
    # Key question: Are we keeping up with incoming emails?
    if avg_received > avg_handled:
        # Building backlog = understaffed
        status_color = "#ef4444"  # Red
        status = "⚠️ **Understaffed**"
        message = f"Receiving {avg_received:.0f} emails/day but only handling {avg_handled:.0f}. Backlog growing by {avg_received - avg_handled:.0f}/day."
    elif demand_vs_capacity_pct > 90:
        # Keeping up but very tight margins
        status_color = "#f59e0b"  # Amber
        status = "⚡ **Near Capacity**"
        message = f"Operating at {demand_vs_capacity_pct:.0f}% of theoretical capacity. Limited buffer for spikes."
    else:
        # Healthy buffer
        status_color = "#10b981"  # Green
        status = "✓ **Healthy Capacity**"
        message = f"Capacity buffer: {theoretical_capacity - avg_received:.0f} emails/day available"
    
    st.markdown(
        f'<div style="padding: 16px; background: {status_color}15; border-left: 4px solid {status_color}; '
        f'border-radius: 8px; margin-top: 16px;">'
        f'<p style="margin: 0; font-size: 14px; font-weight: 600; color: {status_color};">{status}</p>'
        f'<p style="margin: 4px 0 0 0; font-size: 13px; color: #64748b;">{message}</p>'
        f'<p style="margin: 4px 0 0 0; font-size: 12px; color: #94a3b8;">'
        f'Theoretical capacity: {avg_capacity:.1f} hrs/day × {emails_hr:.1f} emails/hr = {theoretical_capacity:.0f} emails | '
        f'Cumulative backlog: {total_backlog:+.0f} emails'
        f'</p></div>',
        unsafe_allow_html=True
    )


# =====================================================
# RESPONSE TIME TREND
# =====================================================

st.markdown("---")
st.subheader("Average Response Time Trend")

daily_resp = art_filtered.groupby("Date")["ResponseSec"].mean().reset_index()
daily_resp['ResponseHours'] = daily_resp['ResponseSec'] / 3600

response_chart = alt.Chart(daily_resp).mark_area(
    line={'color': '#2563eb', 'strokeWidth': 2.5},
    color=alt.Gradient(
        gradient='linear',
        stops=[
            alt.GradientStop(color='#2563eb', offset=0),
            alt.GradientStop(color='#2563eb00', offset=1)
        ],
        x1=0, x2=0, y1=0, y2=1
    ),
    opacity=0.3
).encode(
    x=alt.X('Date:T', title='Date', axis=alt.Axis(labelAngle=-45)),
    y=alt.Y('ResponseHours:Q', title='Hours'),
    tooltip=[
        alt.Tooltip('Date:T', format='%b %d, %Y'),
        alt.Tooltip('ResponseHours:Q', title='Response Time (hours)', format='.1f')
    ]
).properties(
    height=250
).configure_axis(
    labelFontSize=12,
    labelColor='#64748b',
    titleFontSize=13,
    titleColor='#0f172a',
    titleFontWeight=600,
    gridOpacity=0.08,
    domainOpacity=0.2
).configure_view(
    strokeWidth=0
)

st.altair_chart(response_chart, use_container_width=True)
