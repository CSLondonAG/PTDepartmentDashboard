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
OFFLINE_STATUSES = {"Offline"}  # extend if your export includes other offline-like values

BUSINESS_START_HOUR = 7
BUSINESS_END_HOUR = 22

st.markdown(
    """
    <style>
      .stApp {background-color: #f8fafc;}

      .stMarkdown h1 {color: #111827; font-size: 1.75rem; font-weight: 700; margin-bottom: 2px;}
      .stMarkdown h2 {color: #15803d; font-size: 1.1rem; font-weight: 600; margin-top: 2rem;}
      .stMarkdown h3 {color: #374151; font-size: 0.95rem; font-weight: 500;}

      div[data-testid="stMetricValue"] {color: #15803d; font-size: 1.5rem; font-weight: 700;}
      div[data-testid="stMetricLabel"] {color: #6b7280; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em;}
      div[data-testid="stMetric"] {background: #ffffff; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border: 1px solid #e5e7eb;}

      .stButton > button {background-color: #15803d; color: white; border-radius: 8px; border: none; padding: 6px 16px; font-weight: 500; transition: background-color 0.15s ease;}
      .stButton > button:hover {background-color: #166534; color: white; border: none;}
      .stButton > button:active {background-color: #14532d; transform: scale(0.98); transition: transform 0.08s ease;}

      *:focus-visible {outline: 2px solid #15803d !important; outline-offset: 2px !important;}

      div[data-testid="stVegaLiteChart"] {border-radius: 12px; overflow: hidden;}
      div[data-testid="stDataFrame"] {border-radius: 8px; overflow: hidden; border: 1px solid #e5e7eb;}
      div[data-testid="stExpander"] {border: 1px solid #e5e7eb; border-radius: 8px; background: #ffffff;}
      div[data-testid="stAlert"] {border-radius: 8px; border-left: 3px solid #15803d;}
      div[data-testid="stSpinner"] > div {color: #15803d;}
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


def seconds_in_window(pres_df: pd.DataFrame, window_start: pd.Timestamp, window_end: pd.Timestamp) -> float:
    """Sum presence seconds clipped to a window. Treat NaT EndDT as window_end."""
    if pres_df.empty:
        return 0.0
    ends = pres_df["EndDT"].fillna(window_end)
    intervals = [clip(s, e, window_start, window_end) for s, e in zip(pres_df["StartDT"], ends)]
    intervals = [x for x in intervals if x]
    return sum_seconds(intervals)


def _parse_name(name):
    """Reduce a name string to a (first, last) tuple for cross-file fuzzy matching.

    Handles formats:
      "First Last"          -> ("first", "last")
      "First Middle Last"   -> ("first", "last")   # middle name ignored
      "Last, First [Middle]"-> ("first", "last")   # comma-separated reversed
    """
    if not isinstance(name, str) or not name.strip():
        return ("", "")
    name = name.strip()
    if "," in name:
        last_part, first_part = name.split(",", 1)
        first = first_part.strip().split()[0].lower() if first_part.strip() else ""
        last = last_part.strip().lower()
    else:
        tokens = name.split()
        first = tokens[0].lower() if tokens else ""
        last = tokens[-1].lower() if len(tokens) > 1 else tokens[0].lower() if tokens else ""
    return (first, last)


# ---------------- LOAD & PREP ----------------

with st.spinner("Loading data…"):
    email_rec = load(BASE / EMAIL_REC_FILE)
    items = load(BASE / ITEMS_FILE)
    pres = load(BASE / PRES_FILE)
    case_cat = load(BASE / CASE_CAT_FILE)

for df in (email_rec, items, pres, case_cat):
    df.columns = df.columns.str.strip()

# Detect agent column in email_rec / case_cat for per-agent filtering
_agent_keywords = {"agent", "owner"}
_email_agent_col = next(
    (c for c in email_rec.columns if any(kw in c.lower() for kw in _agent_keywords) or c == "User: Full Name"),
    None,
)
_case_cat_agent_col = next(
    (c for c in case_cat.columns if any(kw in c.lower() for kw in _agent_keywords) or c == "User: Full Name"),
    None,
)

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

# Keep FULL presence (do not filter to available only)
pres = pres.copy()


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

filter_col1, filter_col2 = st.columns([3, 2])
with filter_col1:
    start, end = st.date_input(
        "Date Range",
        value=(default_start, default_end),
        help="Shows previous completed week by default",
    )
with filter_col2:
    _all_agents_label = "All Agents (Department)"
    _agent_pool = sorted(items["User: Full Name"].dropna().astype(str).unique().tolist())
    selected_agent = st.selectbox(
        "Agent",
        [_all_agents_label] + _agent_pool,
        index=0,
        help="Select an agent for individual performance, or keep the default for the full department view.",
    )
    is_dept_view = selected_agent == _all_agents_label

# ---------------- FILTERED DATA (DATE RANGE ONLY) ----------------

email_rec_period = email_rec[(email_rec["Date_Opened"] >= start) & (email_rec["Date_Opened"] <= end)].copy()
case_cat_period = case_cat[(case_cat["Date_Opened"] >= start) & (case_cat["Date_Opened"] <= end)].copy()
items_period = items[(items["Date_Closed"] >= start) & (items["Date_Closed"] <= end)].copy()

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)

# Apply agent filter where data supports it
if not is_dept_view:
    items_period = items_period[items_period["User: Full Name"].astype(str) == selected_agent].copy()
    _agent_key = _parse_name(selected_agent)
    if _email_agent_col:
        email_rec_period = email_rec_period[
            email_rec_period[_email_agent_col].astype(str).apply(_parse_name) == _agent_key
        ].copy()
    if _case_cat_agent_col:
        case_cat_period = case_cat_period[
            case_cat_period[_case_cat_agent_col].astype(str).apply(_parse_name) == _agent_key
        ].copy()


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

# Presence subsets (scoped to selected window for agent coverage)
pres_in_window = pres[(pres["StartDT"] < end_ts) & (pres["EndDT"].fillna(end_ts) > start_ts)].copy()

# Capture all pres names in window before agent filter (used for debug output below)
_pres_window_names = (
    pres_in_window["Created By: Full Name"].dropna().astype(str).unique()
    if not is_dept_view else []
)
_matching_pres_names: set = set()

if not is_dept_view:
    _agent_key = _parse_name(selected_agent)
    # Primary: match first + last
    _matching_pres_names = {n for n in _pres_window_names if _parse_name(n) == _agent_key}
    # Fallback: last name only (handles nickname / shortened first name)
    if not _matching_pres_names and _agent_key[1]:
        _matching_pres_names = {
            n for n in _pres_window_names
            if _parse_name(n)[1] == _agent_key[1]
        }
    pres_in_window = pres_in_window[
        pres_in_window["Created By: Full Name"].isin(_matching_pres_names)
    ].copy()

pres_avail = pres_in_window[pres_in_window["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)].copy()
pres_online = pres_in_window[~pres_in_window["Service Presence Status: Developer Name"].isin(OFFLINE_STATUSES)].copy()

available_sec = seconds_in_window(pres_avail, start_ts, end_ts)
available_hours = available_sec / 3600

online_sec = seconds_in_window(pres_online, start_ts, end_ts)
online_hours = online_sec / 3600

# Utilisation fix: align numerator to the same agent population present in Presence export.
# If Presence is missing some agents, handle time from those agents must not be included.
presence_agents = set(pres_online["Created By: Full Name"].dropna().astype(str).unique().tolist())
_pres_name_keys = {_parse_name(n) for n in presence_agents}
items_for_util = items_period[
    items_period["User: Full Name"].astype(str).apply(lambda n: _parse_name(n) in _pres_name_keys)
].copy()

total_handle_sec = items_for_util["HandleSec"].sum()
util = (total_handle_sec / online_sec) if online_sec > 0 else 0

# Coverage indicator (internal diagnostic; shown as metric)
items_agents = set(items_period["User: Full Name"].dropna().astype(str).unique().tolist())
_items_name_keys = {_parse_name(n) for n in items_agents}
covered_agents = sum(1 for k in _items_name_keys if k in _pres_name_keys)
coverage = (covered_agents / len(_items_name_keys)) if len(_items_name_keys) > 0 else 0

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
    aging_labels = ["0-4h", "4-24h", "1-3d", "3d+"]
    closed_aging_summary = pd.DataFrame({"Bucket": aging_labels, "Count": [0] * 4})


# ---------------- DISPLAY ----------------

_view_label = selected_agent if not is_dept_view else "Department"
st.markdown(
    f"<p style='color:#6b7280;margin-top:-8px;margin-bottom:20px;font-size:0.9rem;'>{start} — {end} · {_view_label}</p>",
    unsafe_allow_html=True,
)

# ── Primary metrics (top tier) ──
if is_dept_view:
    c1, c2, c3 = st.columns(3)
    c1.metric("Emails Received", f"{total_received:,}")
    c2.metric("Work Items Handled", f"{total_handled:,}")
    c3.metric("Avg Response Time (BH)", hm(avg_art))
else:
    c2, c3 = st.columns(2)
    c2.metric("Work Items Handled", f"{total_handled:,}")
    c3.metric("Avg Response Time (BH)", hm(avg_art))

# ── Secondary metrics (contextual) ──
st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
s1, s2, s3, s4 = st.columns(4)
s1.metric("Avg Handle Time", mmss(avg_aht))
s2.metric("Online Hours", f"{online_hours:.1f}")
s3.metric("Utilisation", f"{util:.1%}")
if is_dept_view:
    s4.metric("Presence Coverage", f"{coverage:.0%}")
else:
    s4.metric("In Presence Data", "Yes" if coverage > 0 else "No")

# Show name-match diagnostic inline when presence data is missing for the selected agent
if not is_dept_view and not _matching_pres_names:
    _sample = list(_pres_window_names[:8])
    _key_str = str(_parse_name(selected_agent))
    _names_str = (
        ", ".join(_sample) + (" ..." if len(_pres_window_names) > 8 else "")
        if _sample else "none - check that the date range overlaps the presence export."
    )
    st.warning("No presence data matched for agent: " + selected_agent)
    st.caption(
        "Parsed key: " + _key_str
        + " | Presence names in window (" + str(len(_pres_window_names)) + " unique): "
        + _names_str
    )
# Daily counts
daily_received = email_rec_period.groupby("Date_Opened").size().reset_index(name="Emails_Received")
daily_received = daily_received.rename(columns={"Date_Opened": "Date"})

daily_handled = items_period.groupby("Date_Closed").size().reset_index(name="Items_Handled")
daily_handled = daily_handled.rename(columns={"Date_Closed": "Date"})

daily = daily_received.merge(daily_handled, on="Date", how="outer").fillna(0)
daily["Date"] = pd.to_datetime(daily["Date"], errors="coerce")
daily = daily.dropna(subset=["Date"]).copy()
daily["Items_Handled"] = daily["Items_Handled"].astype(int)
daily["Emails_Received"] = daily["Emails_Received"].astype(int)


def hours_for_day_available(day_ts):
    ds = pd.Timestamp(day_ts).normalize()
    de = ds + pd.Timedelta(days=1)
    return seconds_in_window(pres_avail, ds, de) / 3600


if len(daily) > 0:
    daily["Available_Hours"] = daily["Date"].apply(hours_for_day_available)
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

    count_max = dow_counts_long["AverageCount"].max()
    hours_max = dow["Available_Hours"].max()
    scale_factor = count_max / hours_max if hours_max > 0 else 1
    dow["Available_Hours_Scaled"] = dow["Available_Hours"] * scale_factor

    dow_bar = alt.Chart(dow_counts_long).mark_bar().encode(
        x=alt.X(
            "DoWShort:N",
            title="Day of Week",
            sort=dow["DoWShort"].tolist(),
            axis=alt.Axis(labelAngle=0, labelPadding=6),
        ),
        y=alt.Y(
            "AverageCount:Q",
            title="Avg Count",
            axis=alt.Axis(orient="left", format=".0f", titlePadding=12),
        ),
        color=alt.Color(
            "Metric:N",
            title="Legend",
            scale=alt.Scale(domain=color_domain, range=color_range),
        ),
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

    dow_hours_line = alt.Chart(dow).mark_line(
        point=alt.OverlayMarkDef(filled=True, size=70), color="#0d9488", strokeWidth=3
    ).encode(
        x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
        y=alt.Y("Available_Hours_Scaled:Q", axis=None),
        tooltip=["DoW", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
    )

    dow_hours_labels = alt.Chart(dow).mark_text(dy=-10, color="#0d9488", fontSize=10).encode(
        x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
        y=alt.Y("Available_Hours_Scaled:Q", axis=None),
        text=alt.Text("Available_Hours:Q", format=".1f"),
    )

    dow_chart = alt.layer(dow_bar, dow_bar_labels, dow_hours_line, dow_hours_labels).properties(height=340)
    st.altair_chart(dow_chart, use_container_width=True)

    st.markdown(
        """
        <div style="display:flex;gap:24px;justify-content:center;margin-top:-8px;margin-bottom:8px;">
            <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:#374151;">
                <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#15803d;"></span>
                Emails Received
            </span>
            <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:#374151;">
                <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#86efac;"></span>
                Items Handled
            </span>
            <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:#374151;">
                <span style="display:inline-block;width:28px;height:3px;background:#0d9488;border-radius:2px;position:relative;top:0px;"></span>
                Avg Available Hours
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.info("No daily data available for the selected date range. Try adjusting the date picker above.")

st.subheader("SLA Performance")
if closed_aging_summary["Count"].sum() > 0:
    closed_aging_bars = alt.Chart(closed_aging_summary).mark_bar(
        color="#15803d", cornerRadiusTopLeft=4, cornerRadiusTopRight=4
    ).encode(
        x=alt.X("Bucket:N", title="Business-hour Response Bucket", sort=aging_labels),
        y=alt.Y("Count:Q", title="Closed Email Count"),
        tooltip=["Bucket", "Count"],
    )
    closed_aging_labels = alt.Chart(closed_aging_summary).mark_text(
        dy=-10, color="#15803d", fontSize=11
    ).encode(
        x=alt.X("Bucket:N", sort=aging_labels),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
    )
    closed_aging_chart = alt.layer(closed_aging_bars, closed_aging_labels).properties(height=340)
    st.altair_chart(closed_aging_chart, use_container_width=True)
    _sla_note = "" if is_dept_view or _email_agent_col else " Showing department-level data (no agent column detected in email export)."
    st.caption("Closed emails grouped by business-hour response time." + _sla_note)
else:
    st.info(
        "No closed emails with response time data for the selected period. Adjust the date range or check that completion timestamps are present."
    )

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
        tooltip=[
            "Category",
            "Reason",
            alt.Tooltip("Count:Q", format=","),
            alt.Tooltip("CategoryTotal:Q", format=","),
        ],
    )

    cat_labels = alt.Chart(chart_data[["Category", "CategoryTotal"]].drop_duplicates()).mark_text(
        dx=6, color="#15803d", fontSize=10
    ).encode(
        x=alt.X("CategoryTotal:Q"),
        y=alt.Y("Category:N", sort=category_sort),
        text=alt.Text("CategoryTotal:Q", format=","),
    )

    stacked_chart = alt.layer(cat_bars, cat_labels).properties(
        height=min(max(340, len(category_sort) * 26), 600)
    )
    st.altair_chart(stacked_chart, use_container_width=True)
    _cat_note = "" if is_dept_view or _case_cat_agent_col else " Showing department-level data (no agent column detected in case category export)."
    st.caption("Top categories with reason-level distribution. Less frequent reasons grouped as 'Other'." + _cat_note)

    heatmap = (
        alt.Chart(chart_data)
        .mark_rect()
        .encode(
            x=alt.X("Reason:N", sort=reason_sort, title="Reason"),
            y=alt.Y("Category:N", sort=category_sort, title="Category"),
            color=alt.Color("Count:Q", title="Cases", scale=alt.Scale(scheme="greens")),
            tooltip=["Category", "Reason", alt.Tooltip("Count:Q", format=",")],
        )
        .properties(height=min(max(340, len(category_sort) * 24), 600))
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
    if not is_dept_view:
        st.markdown("**Presence Name Matching**")
        _dbg_col1, _dbg_col2 = st.columns(2)
        with _dbg_col1:
            st.caption(f"Selected agent (items): `{selected_agent}`")
            st.caption(f"Parsed key: `{_parse_name(selected_agent)}`")
            st.caption(f"Matched pres names: {_matching_pres_names if _matching_pres_names else '⚠️ None — names do not match across files'}")
        with _dbg_col2:
            _dbg_rows = [{"Presence Name": n, "Parsed Key": str(_parse_name(n))} for n in list(_pres_window_names)[:20]]
            if _dbg_rows:
                st.dataframe(pd.DataFrame(_dbg_rows), hide_index=True, use_container_width=True)
            else:
                st.caption("No presence rows found in selected date window.")
        st.divider()
    st.markdown(
        "<p style='color:#6b7280;font-size:0.82rem;margin-bottom:8px;'>Rows excluded due to unparseable timestamps — review source data if counts are high.</p>",
        unsafe_allow_html=True,
    )
    dq1, dq2, dq3 = st.columns(3)
    dq1.metric("Invalid Opened Timestamps", f"{email_invalid_open:,}")
    dq2.metric("Invalid Completion Timestamps", f"{email_invalid_complete:,}")
    dq3.metric("Invalid Item Close Timestamps", f"{items_invalid_close:,}")
