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
CASE_CAT_FILE = "CaseCatPT.csv"

AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}
BUSINESS_START_HOUR = 7
BUSINESS_END_HOUR = 22

st.markdown(
    """
    <style>
      /* ── Base ── */
      .stApp {background-color: #f8fafc;}

      /* ── Typography hierarchy ── */
      .stMarkdown h1 {color: #111827; font-size: 1.75rem; font-weight: 700; margin-bottom: 2px;}
      .stMarkdown h2 {color: #15803d; font-size: 1.1rem; font-weight: 600; margin-top: 2rem;}
      .stMarkdown h3 {color: #374151; font-size: 0.95rem; font-weight: 500;}

      /* ── Metric tiers ── */
      div[data-testid="stMetricValue"] {color: #15803d; font-size: 1.5rem; font-weight: 700;}
      div[data-testid="stMetricLabel"] {color: #6b7280; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;}
      div[data-testid="stMetric"] {background: #ffffff; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border: 1px solid #e5e7eb;}

      /* ── Button ── */
      .stButton > button {background-color: #15803d; color: white; border-radius: 8px; border: none; padding: 6px 16px; font-weight: 500; transition: background-color 0.15s ease;}
      .stButton > button:hover {background-color: #166534; color: white; border: none;}
      .stButton > button:active {background-color: #14532d; transform: scale(0.98); transition: transform 0.08s ease;}

      /* ── Focus ring ── */
      *:focus-visible {outline: 2px solid #15803d !important; outline-offset: 2px !important;}

      /* ── Chart containers ── */
      div[data-testid="stVegaLiteChart"] {border-radius: 12px; overflow: hidden;}

      /* ── Table refinement ── */
      div[data-testid="stDataFrame"] {border-radius: 8px; overflow: hidden; border: 1px solid #e5e7eb;}

      /* ── Expander ── */
      div[data-testid="stExpander"] {border: 1px solid #e5e7eb; border-radius: 8px; background: #ffffff;}

      /* ── Info/empty state ── */
      div[data-testid="stAlert"] {border-radius: 8px; border-left: 3px solid #15803d;}

      /* ── Spinner ── */
      div[data-testid="stSpinner"] > div {color: #15803d;}

      /* ── Remove default Streamlit divider weight ── */
      hr {border-color: #e5e7eb; border-width: 1px 0 0 0; margin: 1.5rem 0;}
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

with st.spinner("Loading data…"):
    email_rec = load(BASE / EMAIL_REC_FILE)
    items = load(BASE / ITEMS_FILE)
    pres = load(BASE / PRES_FILE)
    case_cat = load(BASE / CASE_CAT_FILE)

for df in (email_rec, items, pres, case_cat):
    df.columns = df.columns.str.strip()

email_rec["OpenedDT"] = pd.to_datetime(email_rec["Date/Time Opened"], errors="coerce", dayfirst=True)
email_rec["CompletedDT"] = pd.to_datetime(email_rec["Completion Date"], errors="coerce", dayfirst=True)
email_rec["Date_Opened"] = email_rec["OpenedDT"].dt.date
email_rec["Date_Completed"] = email_rec["CompletedDT"].dt.date
email_rec["TargetResponseHours"] = pd.to_numeric(email_rec["Target Response (Hours)"], errors="coerce")

case_cat["OpenedDT"] = pd.to_datetime(case_cat["Date/Time Opened"], errors="coerce", dayfirst=True)
case_cat["Date_Opened"] = case_cat["OpenedDT"].dt.date

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

title_col, refresh_col = st.columns([5, 1])
with title_col:
    st.title("Email Department Performance")
with refresh_col:
    st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
    if st.button("Refresh Data", use_container_width=True):
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

# Count non-default filters for indicator
_active_filters = sum([
    len(selected_queues) < len(queue_options),
    len(selected_origins) < len(origin_options),
    len(selected_priorities) < len(priority_options),
    len(selected_milestones) < len(milestone_options),
])
if _active_filters > 0:
    st.caption(f"⚠ {_active_filters} filter{'s' if _active_filters > 1 else ''} active — results are scoped.")


# ---------------- FILTERED DATA ----------------

email_rec_period = email_rec[
    (email_rec["Date_Opened"] >= start)
    & (email_rec["Date_Opened"] <= end)
    & (email_rec["Case Origin"].isin(selected_origins))
    & (email_rec["Priority"].isin(selected_priorities))
    & (email_rec["Milestone"].isin(selected_milestones))
].copy()

case_cat_period = case_cat[
    (case_cat["Date_Opened"] >= start)
    & (case_cat["Date_Opened"] <= end)
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

if len(completed_emails) > 0:
    closed_age_hours = completed_emails["ResponseTimeBusinessSec"] / 3600
    aging_bins = [0, 4, 24, 72, np.inf]
    aging_labels = ["0-4h", "4-24h", "1-3d", "3d+"]
    completed_emails["AgingBucket"] = pd.cut(closed_age_hours, bins=aging_bins, labels=aging_labels, right=False)
    closed_aging_summary = completed_emails["AgingBucket"].value_counts().reindex(aging_labels, fill_value=0).reset_index()
    closed_aging_summary.columns = ["Bucket", "Count"]
else:
    closed_aging_summary = pd.DataFrame({"Bucket": aging_labels, "Count": [0] * 4})


# ---------------- DISPLAY ----------------

st.markdown(
    f"<p style='color:#6b7280;margin-top:-8px;margin-bottom:20px;font-size:0.9rem;'>{start} — {end}</p>",
    unsafe_allow_html=True,
)

# ── Primary metrics (top tier) ──
c1, c2, c3 = st.columns(3)
c1.metric("Emails Received", f"{total_received:,}")
c2.metric("Work Items Handled", f"{total_handled:,}")
c3.metric("Avg Response Time (BH)", hm(avg_art))

# ── Secondary metrics (contextual) ──
st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
s1, s2, s3 = st.columns(3)
s1.metric("Avg Handle Time", mmss(avg_aht))
s2.metric("Available Hours", f"{available_hours:.1f}")
s3.metric("Utilisation", f"{util:.1%}")

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
    daily["DateLabel"] = daily["Date"].dt.strftime("%a %d %b")

st.subheader("Day-of-Week Pattern")
if len(daily) > 0:
    ordered_dow = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    dow = daily.copy()
    dow["DoW"] = dow["Date"].dt.day_name()

    dow = (
        dow.groupby("DoW", as_index=False)[["Emails_Received", "Items_Handled", "Available_Hours"]]
        .mean()
        .set_index("DoW")
        .reindex(ordered_dow, fill_value=0)
        .reset_index()
    )
    dow["DoW"] = pd.Categorical(dow["DoW"], categories=ordered_dow, ordered=True)
    dow["DoWShort"] = dow["DoW"].astype(str).str.slice(0, 3)

    color_domain = ["Emails Received", "Items Handled", "Available Hours"]
    color_range = ["#15803d", "#86efac", "#0d9488"]

    dow_counts_long = dow.melt(
        id_vars=["DoW", "DoWShort"],
        value_vars=["Emails_Received", "Items_Handled"],
        var_name="Metric",
        value_name="AverageCount",
    )
    dow_counts_long["Metric"] = dow_counts_long["Metric"].replace(
        {"Emails_Received": "Emails Received", "Items_Handled": "Items Handled"}
    )

    dow_bar = alt.Chart(dow_counts_long).mark_bar().encode(
        x=alt.X("DoWShort:N", title="Day of Week", sort=dow["DoWShort"].tolist(), axis=alt.Axis(labelAngle=0, labelPadding=6)),
        y=alt.Y("AverageCount:Q", title="Average Count", axis=alt.Axis(orient="left", format=".0f")),
        color=alt.Color("Metric:N", title="Legend", scale=alt.Scale(domain=color_domain, range=color_range)),
        xOffset="Metric:N",
        tooltip=["DoW", "Metric", alt.Tooltip("AverageCount:Q", format=",.0f")],
    )
    dow_bar_labels = alt.Chart(dow_counts_long).mark_text(dy=-8, fontSize=10).encode(
        x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
        y=alt.Y("AverageCount:Q"),
        xOffset="Metric:N",
        text=alt.Text("AverageCount:Q", format=",.0f"),
        color=alt.Color("Metric:N", scale=alt.Scale(domain=color_domain, range=color_range), legend=None),
    )

    dow_hours = dow.copy()
    dow_hours_line = alt.Chart(dow_hours).mark_line(point=alt.OverlayMarkDef(filled=True, size=70), color="#0d9488", strokeWidth=3).encode(
        x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
        y=alt.Y("Available_Hours:Q", title="Average Available Hours", axis=alt.Axis(orient="right", format=".1f")),
        tooltip=["DoW", alt.Tooltip("Available_Hours:Q", format=".1f")],
    )
    dow_hours_labels = alt.Chart(dow_hours).mark_text(dy=-10, color="#0d9488", fontSize=10).encode(
        x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
        y=alt.Y("Available_Hours:Q"),
        text=alt.Text("Available_Hours:Q", format=".1f"),
    )

    legend_source = pd.DataFrame({"Metric": color_domain, "x": [0, 0, 0], "y": [0, 0, 0]})
    legend_layer = alt.Chart(legend_source).mark_point(opacity=0).encode(
        color=alt.Color("Metric:N", title="Legend", scale=alt.Scale(domain=color_domain, range=color_range))
    )

    dow_chart = (
        alt.layer(dow_bar, dow_bar_labels, dow_hours_line, dow_hours_labels, legend_layer)
        .resolve_scale(y="independent")
        .properties(height=340)
        .configure_tooltip(theme="dark")
        .configure_view(strokeWidth=0)
    )
    st.altair_chart(dow_chart, use_container_width=True)
else:
    st.info("No daily data available for the selected date range. Try adjusting the date picker above.")

st.subheader("SLA Performance")
if closed_aging_summary["Count"].sum() > 0:
    closed_aging_bars = alt.Chart(closed_aging_summary).mark_bar(color="#15803d", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("Bucket:N", title="Business-hour Response Bucket", sort=aging_labels),
        y=alt.Y("Count:Q", title="Closed Email Count"),
        tooltip=["Bucket", "Count"],
    )
    closed_aging_labels = alt.Chart(closed_aging_summary).mark_text(dy=-10, color="#15803d", fontSize=11).encode(
        x=alt.X("Bucket:N", sort=aging_labels),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
    )
    closed_aging_chart = (
        alt.layer(closed_aging_bars, closed_aging_labels)
        .properties(height=340)
        .configure_tooltip(theme="dark")
        .configure_view(strokeWidth=0)
    )
    st.altair_chart(closed_aging_chart, use_container_width=True)
    st.caption("Closed emails grouped by business-hour response time.")
else:
    st.info("No closed emails with response time data for the selected period. Adjust the date range or check that completion timestamps are present.")

st.subheader("Case Category & Reason Breakdown")
if len(case_cat_period) > 0:
    cat_reason_summary = (
        case_cat_period.groupby(["Category", "Reason"], dropna=False)
        .size()
        .reset_index(name="Count")
    )
    cat_reason_summary["Category"] = cat_reason_summary["Category"].fillna("Unspecified")
    cat_reason_summary["Reason"] = cat_reason_summary["Reason"].fillna("Unspecified")

    cat_totals = (
        cat_reason_summary.groupby("Category", as_index=False)["Count"]
        .sum()
        .sort_values("Count", ascending=False)
        .rename(columns={"Count": "CategoryTotal"})
    )

    category_count = int(cat_totals.shape[0])
    max_categories = min(30, category_count)

    controls_col1, controls_col2 = st.columns(2)
    with controls_col1:
        top_categories = st.slider(
            "Top categories to display",
            min_value=1,
            max_value=max_categories,
            value=min(12, max_categories),
            step=1,
            help="Shows the largest categories and keeps the chart readable for high-volume datasets.",
        )
    with controls_col2:
        top_reasons = st.slider(
            "Top reasons per category",
            min_value=1,
            max_value=12,
            value=6,
            step=1,
            help="Additional reasons are grouped into 'Other'.",
        )

    selected_categories = cat_totals.head(top_categories)["Category"].tolist()
    filtered = cat_reason_summary[cat_reason_summary["Category"].isin(selected_categories)].copy()

    filtered = filtered.sort_values(["Category", "Count"], ascending=[True, False])
    filtered["ReasonRank"] = filtered.groupby("Category")["Count"].rank(method="first", ascending=False)
    filtered["ReasonCollapsed"] = np.where(filtered["ReasonRank"] <= top_reasons, filtered["Reason"], "Other")

    chart_data = (
        filtered.groupby(["Category", "ReasonCollapsed"], as_index=False)["Count"]
        .sum()
        .rename(columns={"ReasonCollapsed": "Reason"})
    )

    category_sort = cat_totals[cat_totals["Category"].isin(selected_categories)]["Category"].tolist()
    reason_totals = chart_data.groupby("Reason", as_index=False)["Count"].sum().sort_values("Count", ascending=False)
    reason_sort = reason_totals["Reason"].tolist()

    chart_data = chart_data.merge(
        chart_data.groupby("Category", as_index=False)["Count"].sum().rename(columns={"Count": "CategoryTotal"}),
        on="Category",
        how="left",
    )

    cat_bars = alt.Chart(chart_data).mark_bar().encode(
        x=alt.X("Count:Q", title="Case Count"),
        y=alt.Y("Category:N", title="Category", sort=category_sort),
        color=alt.Color("Reason:N", title="Reason", sort=reason_sort),
        order=alt.Order("Count:Q", sort="descending"),
        tooltip=["Category", "Reason", alt.Tooltip("Count:Q", format=","), alt.Tooltip("CategoryTotal:Q", format=",")],
    )

    cat_labels = alt.Chart(
        chart_data[["Category", "CategoryTotal"]].drop_duplicates()
    ).mark_text(dx=6, color="#15803d", fontSize=10).encode(
        x=alt.X("CategoryTotal:Q"),
        y=alt.Y("Category:N", sort=category_sort),
        text=alt.Text("CategoryTotal:Q", format=","),
    )

    stacked_chart = (
        alt.layer(cat_bars, cat_labels)
        .properties(height=min(max(340, len(category_sort) * 26), 600))
        .configure_tooltip(theme="dark")
        .configure_view(strokeWidth=0)
    )
    st.altair_chart(stacked_chart, use_container_width=True)
    st.caption("Top categories with reason-level distribution. Less frequent reasons grouped as 'Other'.")

    heatmap_data = chart_data.copy()
    heatmap = (
        alt.Chart(heatmap_data)
        .mark_rect()
        .encode(
            x=alt.X("Reason:N", sort=reason_sort, title="Reason"),
            y=alt.Y("Category:N", sort=category_sort, title="Category"),
            color=alt.Color("Count:Q", title="Cases", scale=alt.Scale(scheme="greens")),
            tooltip=["Category", "Reason", alt.Tooltip("Count:Q", format=",")],
        )
        .properties(height=min(max(340, len(category_sort) * 24), 600))
        .configure_tooltip(theme="dark")
        .configure_view(strokeWidth=0)
    )
    st.altair_chart(heatmap, use_container_width=True)
    st.caption("Heatmap view for scanning dense category/reason combinations.")
else:
    st.info("No case category data available for the selected period. Try widening the date range.")

st.subheader("Daily Breakdown")
daily_display = daily.copy()
if len(daily_display) > 0:
    daily_display["Date"] = daily_display["Date"].dt.date
    daily_display["Available_Hours"] = daily_display["Available_Hours"].round(1)

    daily_display = daily_display.rename(
        columns={
            "Date": "Date",
            "Emails_Received": "Received",
            "Items_Handled": "Handled",
            "Available_Hours": "Avail. Hours",
        }
    )
    st.dataframe(daily_display.sort_values("Date", ascending=True), use_container_width=True, hide_index=True)
else:
    st.info("No daily records found for the selected date range.")


st.markdown("<div style='margin-top:2rem;'></div>", unsafe_allow_html=True)
with st.expander("Data Quality", expanded=False):
    st.markdown("<p style='color:#6b7280;font-size:0.82rem;margin-bottom:8px;'>Rows excluded due to unparseable timestamps — review source data if counts are high.</p>", unsafe_allow_html=True)
    dq1, dq2, dq3 = st.columns(3)
    dq1.metric("Invalid Opened Timestamps", f"{email_invalid_open:,}")
    dq2.metric("Invalid Completion Timestamps", f"{email_invalid_complete:,}")
    dq3.metric("Invalid Item Close Timestamps", f"{items_invalid_close:,}")
