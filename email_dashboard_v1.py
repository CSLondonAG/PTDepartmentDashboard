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

BUSINESS_START_HOUR = 8
BUSINESS_END_HOUR = 20
SLA_TARGET_HOURS = 8
SLA_TARGET_SECONDS = SLA_TARGET_HOURS * 3600
SLA_COMPLIANCE_WEIGHT = 0.60
SLA_RESPONSE_TIME_WEIGHT = 0.40

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


def calculate_email_sla_score(response_hours: pd.Series) -> dict:
    """Calculate a normalized 0-100 Email SLA Score with a genuine 60/40 split.

    Components:
      compliance_score = fraction_within_target * 100
      response_time_score = min(100, target_hours / average_response_hours * 100)
      score = 0.60 * compliance_score + 0.40 * response_time_score

    Both components are bounded to 0-100 before weighting, so compliance can
    contribute at most 60 points and response-time performance at most 40.
    """
    hours = pd.to_numeric(response_hours, errors="coerce").dropna()
    hours = hours[hours >= 0]
    total = int(len(hours))
    if total == 0:
        return {
            "eligible": 0,
            "met": 0,
            "missed": 0,
            "within_fraction": np.nan,
            "avg_response_hours": np.nan,
            "compliance_score": np.nan,
            "response_time_score": np.nan,
            "compliance_points": np.nan,
            "response_time_points": np.nan,
            "score": np.nan,
        }

    met = int((hours <= SLA_TARGET_HOURS).sum())
    within_fraction = met / total
    avg_response_hours = float(hours.mean())

    compliance_score = max(0.0, min(100.0, within_fraction * 100.0))
    if avg_response_hours <= 0:
        response_time_score = 100.0
    else:
        response_time_score = max(
            0.0,
            min(100.0, (SLA_TARGET_HOURS / avg_response_hours) * 100.0),
        )

    compliance_points = SLA_COMPLIANCE_WEIGHT * compliance_score
    response_time_points = SLA_RESPONSE_TIME_WEIGHT * response_time_score
    score = compliance_points + response_time_points

    return {
        "eligible": total,
        "met": met,
        "missed": total - met,
        "within_fraction": within_fraction,
        "avg_response_hours": avg_response_hours,
        "compliance_score": compliance_score,
        "response_time_score": response_time_score,
        "compliance_points": compliance_points,
        "response_time_points": response_time_points,
        "score": score,
    }


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
    st.caption("SLA model: 8 business hours · normalized 60/40 score · version 2026-08-06")
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
else:
    completed_emails["ResponseTimeBusinessSec"] = pd.Series(dtype=float)

# Normalized Email SLA Score with a genuine 60/40 split and an
# 8-business-hour target. Invalid response times are excluded.
valid_completed_emails = completed_emails[
    completed_emails["ResponseTimeBusinessSec"].notna()
].copy()
valid_completed_emails["ResponseTimeBusinessHours"] = (
    valid_completed_emails["ResponseTimeBusinessSec"] / 3600.0
)
valid_completed_emails["SLA_Met"] = (
    valid_completed_emails["ResponseTimeBusinessHours"] <= SLA_TARGET_HOURS
)

_period_sla = calculate_email_sla_score(valid_completed_emails["ResponseTimeBusinessHours"])
sla_eligible = _period_sla["eligible"]
sla_met = _period_sla["met"]
sla_missed = _period_sla["missed"]
sla_compliance = _period_sla["within_fraction"]
sla_avg_response_hours = _period_sla["avg_response_hours"]
sla_period_compliance_score = _period_sla["compliance_score"]
sla_period_response_time_score = _period_sla["response_time_score"]
avg_art = sla_avg_response_hours * 3600 if pd.notna(sla_avg_response_hours) else 0

# Calculate one SLA score per day, then volume-weight the daily scores across
# the selected period, as in the reference department dashboard.
_daily_sla_rows = []
if sla_eligible > 0:
    for _sla_date, _sla_group in valid_completed_emails.groupby("Date_Opened", dropna=True):
        _components = calculate_email_sla_score(_sla_group["ResponseTimeBusinessHours"])
        _daily_sla_rows.append({
            "Date": pd.to_datetime(_sla_date, errors="coerce"),
            "SLA_Eligible": _components["eligible"],
            "SLA_Met": _components["met"],
            "SLA_Missed": _components["missed"],
            "SLA_Compliance": _components["within_fraction"],
            "SLA_Avg_Response_Hours": _components["avg_response_hours"],
            "SLA_Compliance_Score": _components["compliance_score"],
            "SLA_Response_Time_Score": _components["response_time_score"],
            "SLA_Compliance_Points": _components["compliance_points"],
            "SLA_Response_Time_Points": _components["response_time_points"],
            "Email_SLA_Score": _components["score"],
        })

_daily_sla = pd.DataFrame(_daily_sla_rows)
if _daily_sla.empty:
    _daily_sla = pd.DataFrame({
        "Date": pd.Series(dtype="datetime64[ns]"),
        "SLA_Eligible": pd.Series(dtype="int64"),
        "SLA_Met": pd.Series(dtype="int64"),
        "SLA_Missed": pd.Series(dtype="int64"),
        "SLA_Compliance": pd.Series(dtype="float64"),
        "SLA_Avg_Response_Hours": pd.Series(dtype="float64"),
        "SLA_Compliance_Score": pd.Series(dtype="float64"),
        "SLA_Response_Time_Score": pd.Series(dtype="float64"),
        "SLA_Compliance_Points": pd.Series(dtype="float64"),
        "SLA_Response_Time_Points": pd.Series(dtype="float64"),
        "Email_SLA_Score": pd.Series(dtype="float64"),
    })
    email_sla_score = np.nan
    weighted_compliance_score = np.nan
    weighted_response_time_score = np.nan
else:
    _daily_sla = _daily_sla.sort_values("Date").reset_index(drop=True)
    _daily_weight = _daily_sla["SLA_Eligible"].sum()
    email_sla_score = (
        (_daily_sla["Email_SLA_Score"] * _daily_sla["SLA_Eligible"]).sum() / _daily_weight
        if _daily_weight > 0 else np.nan
    )

    # Weighted component scores are kept separately so they add back exactly
    # to the displayed headline SLA Score.
    weighted_compliance_score = (
        (_daily_sla["SLA_Compliance_Score"] * _daily_sla["SLA_Eligible"]).sum() / _daily_weight
        if _daily_weight > 0 else np.nan
    )
    weighted_response_time_score = (
        (_daily_sla["SLA_Response_Time_Score"] * _daily_sla["SLA_Eligible"]).sum() / _daily_weight
        if _daily_weight > 0 else np.nan
    )

sla_status_summary = pd.DataFrame(
    {
        "Status": [f"Met (≤{SLA_TARGET_HOURS}h)", f"Missed (>{SLA_TARGET_HOURS}h)"],
        "Count": [sla_met, sla_missed],
    }
)

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

_presence_status = (
    pres_in_window["Service Presence Status: Developer Name"]
    .fillna("")
    .astype(str)
    .str.strip()
)
_offline_mask = _presence_status.isin(OFFLINE_STATUSES)
# Treat every status beginning with "Busy" as unavailable for utilisation,
# including Busy_Lunch, Busy_Break, Busy_Meeting and future Busy_* statuses.
_busy_mask = _presence_status.str.casefold().str.startswith("busy")

pres_avail = pres_in_window[_presence_status.isin(AVAILABLE_STATUSES)].copy()
pres_online = pres_in_window[~_offline_mask].copy()
pres_util_base = pres_in_window[~_offline_mask & ~_busy_mask].copy()

available_sec = seconds_in_window(pres_avail, start_ts, end_ts)
available_hours = available_sec / 3600

# Online Hours remains all non-Offline presence time for context.
online_sec = seconds_in_window(pres_online, start_ts, end_ts)
online_hours = online_sec / 3600

# Utilisation denominator excludes Offline and all Busy* statuses.
util_presence_sec = seconds_in_window(pres_util_base, start_ts, end_ts)

# Align the handle-time numerator to agents represented in the utilisation presence base.
util_presence_agents = set(
    pres_util_base["Created By: Full Name"].dropna().astype(str).unique().tolist()
)
_util_pres_name_keys = {_parse_name(n) for n in util_presence_agents}
_util_mask = (
    items_period["User: Full Name"].astype(str)
    .apply(lambda n: _parse_name(n) in _util_pres_name_keys)
    .astype(bool)
)
items_for_util = items_period[_util_mask].copy()

total_handle_sec = items_for_util["HandleSec"].sum()
util = (total_handle_sec / util_presence_sec) if util_presence_sec > 0 else 0

email_invalid_open = email_rec_period["OpenedDT"].isna().sum()
email_invalid_complete = email_rec_period["CompletedDT"].isna().sum()
items_invalid_close = items_period["CloseDT"].isna().sum()

aging_labels = ["0-4h", "4-8h", "8-24h", "1-3d", "3d+"]
if sla_eligible > 0:
    closed_age_hours = valid_completed_emails["ResponseTimeBusinessSec"] / 3600
    aging_bins = [-np.inf, 4, SLA_TARGET_HOURS, 24, 72, np.inf]
    valid_completed_emails["AgingBucket"] = pd.cut(
        closed_age_hours,
        bins=aging_bins,
        labels=aging_labels,
        right=True,
    )
    closed_aging_summary = (
        valid_completed_emails["AgingBucket"]
        .value_counts()
        .reindex(aging_labels, fill_value=0)
        .reset_index()
    )
    closed_aging_summary.columns = ["Bucket", "Count"]
else:
    closed_aging_summary = pd.DataFrame({"Bucket": aging_labels, "Count": [0] * len(aging_labels)})


# ---------------- DISPLAY ----------------

_view_label = selected_agent if not is_dept_view else "Department"
st.markdown(
    f"<p style='color:#6b7280;margin-top:-8px;margin-bottom:20px;font-size:0.9rem;'>{start} — {end} · {_view_label}</p>",
    unsafe_allow_html=True,
)

# ── Primary metrics (top tier) ──
sla_display = f"{email_sla_score:.1f}" if pd.notna(email_sla_score) else "—"
sla_compliance_display = f"{sla_compliance:.1%}" if pd.notna(sla_compliance) else "—"
sla_help = (
    f"Daily-volume weighted 0–100 Email SLA Score: 60% from the share replied "
    f"within {SLA_TARGET_HOURS} business hours and 40% from normalized average "
    f"response time. The response component scores 100 at or below {SLA_TARGET_HOURS}h "
    f"and declines proportionally above target."
)

if is_dept_view:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Emails Received", f"{total_received:,}")
    c2.metric("Work Items Handled", f"{total_handled:,}")
    c3.metric("Avg Response Time (BH)", hm(avg_art))
    c4.metric("Email SLA Score (60/40)", sla_display, help=sla_help)
else:
    c1, c2, c3 = st.columns(3)
    c1.metric("Work Items Handled", f"{total_handled:,}")
    c2.metric("Avg Response Time (BH)", hm(avg_art))
    c3.metric("Email SLA Score (60/40)", sla_display, help=sla_help)

# ── Secondary metrics (contextual) ──
st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
s1, s2, s3, s4 = st.columns(4)
s1.metric("Avg Handle Time", mmss(avg_aht))
s2.metric(
    "Online Hours",
    f"{online_hours:.1f}",
    help="Total presence time excluding Offline status.",
)
s3.metric(
    "Available Hours",
    f"{available_hours:.1f}",
    help="Time in Available_Email_and_Web or Available_All status.",
)
s4.metric(
    "Utilisation",
    f"{util:.1%}",
    help=(
        "Email handle time divided by matched presence time after excluding "
        "Offline and all Busy* statuses."
    ),
)

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
_daily_aht = (
    items_period.groupby("Date_Closed")["HandleSec"].mean()
    .reset_index()
    .rename(columns={"Date_Closed": "Date", "HandleSec": "AvgHandleSec"})
)
_daily_aht["Date"] = pd.to_datetime(_daily_aht["Date"], errors="coerce")
daily = daily.merge(_daily_aht, on="Date", how="left")

daily = daily.merge(_daily_sla, on="Date", how="left")
daily["SLA_Eligible"] = daily["SLA_Eligible"].fillna(0).astype(int)
daily["SLA_Met"] = daily["SLA_Met"].fillna(0).astype(int)
daily["SLA_Missed"] = daily["SLA_Missed"].fillna(0).astype(int)
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

    if is_dept_view:
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
            x=alt.X("DoWShort:N", title="Day of Week", sort=dow["DoWShort"].tolist(),
                    axis=alt.Axis(labelAngle=0, labelPadding=6)),
            y=alt.Y("AverageCount:Q", title="Avg Count",
                    axis=alt.Axis(orient="left", format=".0f", titlePadding=12)),
            color=alt.Color("Metric:N", title="Legend",
                            scale=alt.Scale(domain=color_domain, range=color_range)),
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
        # Agent view: items handled bars + available hours line
        count_max = dow["Items_Handled"].max()
        hours_max = dow["Available_Hours"].max()
        scale_factor = count_max / hours_max if hours_max > 0 else 1
        dow["Available_Hours_Scaled"] = dow["Available_Hours"] * scale_factor

        agent_bar = alt.Chart(dow).mark_bar(
            color="#15803d", cornerRadiusTopLeft=4, cornerRadiusTopRight=4
        ).encode(
            x=alt.X("DoWShort:N", title="Day of Week", sort=dow["DoWShort"].tolist(),
                    axis=alt.Axis(labelAngle=0, labelPadding=6)),
            y=alt.Y("Items_Handled:Q", title="Avg Items Handled",
                    axis=alt.Axis(format=".0f", titlePadding=12)),
            tooltip=["DoW", alt.Tooltip("Items_Handled:Q", format=",.1f", title="Avg Handled")],
        )
        agent_bar_labels = alt.Chart(dow).mark_text(dy=-8, fontSize=11, color="#15803d").encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Items_Handled:Q"),
            text=alt.Text("Items_Handled:Q", format=",.0f"),
        )
        agent_hours_line = alt.Chart(dow).mark_line(
            point=alt.OverlayMarkDef(filled=True, size=70), color="#0d9488", strokeWidth=3
        ).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            tooltip=["DoW", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
        )
        agent_hours_labels = alt.Chart(dow).mark_text(dy=-10, color="#0d9488", fontSize=10).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            text=alt.Text("Available_Hours:Q", format=".1f"),
        )
        dow_chart = alt.layer(agent_bar, agent_bar_labels, agent_hours_line, agent_hours_labels).properties(height=340)
        st.altair_chart(dow_chart, use_container_width=True)
        st.markdown(
            """
            <div style="display:flex;gap:24px;justify-content:center;margin-top:-8px;margin-bottom:8px;">
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:#374151;">
                    <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#15803d;"></span>
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
if sla_eligible > 0:
    sla_m1, sla_m2, sla_m3 = st.columns(3)
    sla_m1.metric("Email SLA Score (60/40)", sla_display, help=sla_help)
    sla_m2.metric(
        f"Replied ≤{SLA_TARGET_HOURS}h",
        sla_compliance_display,
        help=f"{sla_met:,} of {sla_eligible:,} eligible completed emails met the target."
    )
    sla_m3.metric(
        "Avg Response Time (BH)",
        hm(avg_art),
        help=(
            f"Normalized response score: {weighted_response_time_score:.1f}/100. "
            f"This contributes up to 40 points to the SLA Score; averages at or below "
            f"{SLA_TARGET_HOURS} business hours receive the full response-time component."
        )
    )

    if not _daily_sla.empty:
        _sla_trend = _daily_sla.dropna(subset=["Date", "Email_SLA_Score"]).copy()

        # Explicitly show one x-axis tick per SLA date. Without this, Altair
        # can generate multiple timestamp ticks within each day and then format
        # them all with the same "%d %b" label.
        _sla_tick_dates = (
            _sla_trend["Date"]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        _trend_line = alt.Chart(_sla_trend).mark_line(
            point=alt.OverlayMarkDef(filled=True, size=70), color="#15803d", strokeWidth=3
        ).encode(
            x=alt.X(
                "Date:T",
                title="Date",
                axis=alt.Axis(
                    format="%d %b",
                    labelAngle=-45,
                    values=_sla_tick_dates,
                ),
            ),
            y=alt.Y("Email_SLA_Score:Q", title="Email SLA Score", scale=alt.Scale(domain=[0, 105])),
            tooltip=[
                alt.Tooltip("Date:T", title="Date", format="%d %b %Y"),
                alt.Tooltip("Email_SLA_Score:Q", title="SLA Score", format=".1f"),
                alt.Tooltip("SLA_Compliance:Q", title=f"Replied ≤{SLA_TARGET_HOURS}h", format=".1%"),
                alt.Tooltip("SLA_Avg_Response_Hours:Q", title="Avg Response (h)", format=".2f"),
                alt.Tooltip("SLA_Compliance_Points:Q", title="Compliance Points", format=".1f"),
                alt.Tooltip("SLA_Response_Time_Points:Q", title="Response Points", format=".1f"),
                alt.Tooltip("SLA_Eligible:Q", title="Eligible Emails", format=","),
            ],
        )
        _trend_labels = alt.Chart(_sla_trend).mark_text(
            dy=-10, color="#15803d", fontSize=10
        ).encode(
            x=alt.X("Date:T", axis=None),
            y=alt.Y("Email_SLA_Score:Q"),
            text=alt.Text("Email_SLA_Score:Q", format=".1f"),
        )
        _target_rule = alt.Chart(pd.DataFrame({"Target": [80]})).mark_rule(
            color="#dc2626", strokeDash=[5, 5]
        ).encode(y="Target:Q")
        st.altair_chart(
            alt.layer(_trend_line, _trend_labels, _target_rule).properties(height=300),
            use_container_width=True,
        )
        st.caption("Daily Email SLA Score trend. Dashed line shows a score of 80.")

    sla_status_chart = alt.Chart(sla_status_summary).mark_bar(
        cornerRadiusTopLeft=4,
        cornerRadiusTopRight=4,
    ).encode(
        x=alt.X("Status:N", title=None, sort=sla_status_summary["Status"].tolist()),
        y=alt.Y("Count:Q", title="Completed Email Count"),
        color=alt.Color(
            "Status:N",
            legend=None,
            scale=alt.Scale(
                domain=sla_status_summary["Status"].tolist(),
                range=["#15803d", "#dc2626"],
            ),
        ),
        tooltip=["Status", alt.Tooltip("Count:Q", format=",")],
    )
    sla_status_labels = alt.Chart(sla_status_summary).mark_text(
        dy=-10,
        fontSize=11,
    ).encode(
        x=alt.X("Status:N", sort=sla_status_summary["Status"].tolist()),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
        color=alt.Color(
            "Status:N",
            legend=None,
            scale=alt.Scale(
                domain=sla_status_summary["Status"].tolist(),
                range=["#15803d", "#dc2626"],
            ),
        ),
    )

    closed_aging_bars = alt.Chart(closed_aging_summary).mark_bar(
        color="#15803d", cornerRadiusTopLeft=4, cornerRadiusTopRight=4
    ).encode(
        x=alt.X("Bucket:N", title="Business-hour Response Bucket", sort=aging_labels),
        y=alt.Y("Count:Q", title="Completed Email Count"),
        tooltip=["Bucket", "Count"],
    )
    closed_aging_labels = alt.Chart(closed_aging_summary).mark_text(
        dy=-10, color="#15803d", fontSize=11
    ).encode(
        x=alt.X("Bucket:N", sort=aging_labels),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
    )

    sla_chart_col, aging_chart_col = st.columns([1, 2])
    with sla_chart_col:
        st.altair_chart(
            alt.layer(sla_status_chart, sla_status_labels).properties(height=340),
            use_container_width=True,
        )
    with aging_chart_col:
        st.altair_chart(
            alt.layer(closed_aging_bars, closed_aging_labels).properties(height=340),
            use_container_width=True,
        )

    _sla_note = "" if is_dept_view or _email_agent_col else " Showing department-level data (no agent column detected in email export)."
    st.caption(
        f"Email SLA Score is a normalized 60/40 measure. Compliance contributes up to 60 points "
        f"({SLA_COMPLIANCE_WEIGHT:.0%} × the percentage replied within {SLA_TARGET_HOURS} business hours). "
        f"Average response performance contributes up to 40 points "
        f"({SLA_RESPONSE_TIME_WEIGHT:.0%} × min(100, {SLA_TARGET_HOURS} ÷ average response hours × 100)). "
        f"Daily scores are weighted by eligible completed-email volume. Business hours are "
        f"{BUSINESS_START_HOUR:02d}:00–{BUSINESS_END_HOUR:02d}:00, weekends included."
        + _sla_note
    )
else:
    st.info(
        "No completed emails with a valid business-hour response time for the selected period. "
        "Adjust the date range or check that opened and completion timestamps are present."
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

# ── Department-level agent performance charts ──
if is_dept_view and len(items_period) > 0:
    st.subheader("Agent Performance")
    # --- Items handled per agent (horizontal) ---
    agent_handled = (
        items_period.groupby("User: Full Name").size()
        .reset_index(name="Items_Handled")
        .rename(columns={"User: Full Name": "Agent"})
    )
    agent_handled = agent_handled.sort_values("Items_Handled", ascending=True).reset_index(drop=True)

    st.markdown("**Items Handled by Agent**")
    if len(agent_handled) > 0:
        _n_handled = len(agent_handled)
        handled_bar = alt.Chart(agent_handled).mark_bar(
            color="#86efac", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_handled["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("Items_Handled:Q", title="Items Handled",
                    axis=alt.Axis(format=".0f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("Items_Handled:Q", format=",", title="Items Handled")],
        )
        handled_labels = alt.Chart(agent_handled).mark_text(
            dx=6, fontSize=11, color="#15803d", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_handled["Agent"].tolist()),
            x=alt.X("Items_Handled:Q"),
            text=alt.Text("Items_Handled:Q", format=","),
        )
        st.altair_chart(
            alt.layer(handled_bar, handled_labels).properties(height=max(260, _n_handled * 28)),
            use_container_width=True,
        )
    else:
        st.info("No items data available.")

    # --- AHT per agent (horizontal) ---
    agent_aht = (
        items_period.groupby("User: Full Name")["HandleSec"]
        .mean()
        .reset_index()
        .rename(columns={"User: Full Name": "Agent", "HandleSec": "AvgHandleSec"})
    )
    agent_aht = agent_aht[agent_aht["AvgHandleSec"].notna()].copy()
    agent_aht["AHT_minutes"] = agent_aht["AvgHandleSec"] / 60
    agent_aht["AHT_label"] = agent_aht["AvgHandleSec"].apply(mmss)
    agent_aht = agent_aht.sort_values("AvgHandleSec", ascending=True).reset_index(drop=True)

    st.markdown("**Avg Handle Time by Agent**")
    if len(agent_aht) > 0:
        _n_aht = len(agent_aht)
        aht_bar = alt.Chart(agent_aht).mark_bar(
            color="#15803d", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_aht["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("AHT_minutes:Q", title="Avg Handle Time (mins)",
                    axis=alt.Axis(format=".1f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("AHT_label:N", title="AHT (mm:ss)")],
        )
        aht_labels = alt.Chart(agent_aht).mark_text(
            dx=6, fontSize=11, color="#15803d", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_aht["Agent"].tolist()),
            x=alt.X("AHT_minutes:Q"),
            text=alt.Text("AHT_label:N"),
        )
        st.altair_chart(
            alt.layer(aht_bar, aht_labels).properties(height=max(260, _n_aht * 28)),
            use_container_width=True,
        )
    else:
        st.info("No handle time data available.")

    # --- Available hours per agent (horizontal) ---
    if not pres_avail.empty:
        _agent_avail_rows = []
        for _pres_agent, _pres_grp in pres_avail.groupby("Created By: Full Name"):
            _secs = seconds_in_window(_pres_grp, start_ts, end_ts)
            _agent_avail_rows.append({"Agent": _pres_agent, "Available_Hours": _secs / 3600})
        agent_avail_df = pd.DataFrame(_agent_avail_rows)
        agent_avail_df = agent_avail_df[agent_avail_df["Available_Hours"] > 0].sort_values(
            "Available_Hours", ascending=True
        ).reset_index(drop=True)
    else:
        agent_avail_df = pd.DataFrame(columns=["Agent", "Available_Hours"])

    st.markdown("**Available Hours by Agent**")
    if len(agent_avail_df) > 0:
        _n_avail = len(agent_avail_df)
        avail_bar = alt.Chart(agent_avail_df).mark_bar(
            color="#0d9488", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_avail_df["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("Available_Hours:Q", title="Available Hours",
                    axis=alt.Axis(format=".1f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
        )
        avail_labels = alt.Chart(agent_avail_df).mark_text(
            dx=6, fontSize=11, color="#0d9488", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_avail_df["Agent"].tolist()),
            x=alt.X("Available_Hours:Q"),
            text=alt.Text("Available_Hours:Q", format=".1f"),
        )
        st.altair_chart(
            alt.layer(avail_bar, avail_labels).properties(height=max(260, _n_avail * 28)),
            use_container_width=True,
        )
    else:
        st.info("No presence / available-hours data for the selected period.")

with st.expander("Daily Breakdown", expanded=False):
    daily_display = daily.copy()
    if len(daily_display) > 0:
        daily_display["Date"] = daily_display["Date"].dt.date
        daily_display["Available_Hours"] = daily_display["Available_Hours"].round(1)
        daily_display["Email SLA Score"] = daily_display["Email_SLA_Score"].apply(
            lambda value: f"{value:.1f}" if pd.notna(value) else "—"
        )
        daily_display[f"Replied ≤{SLA_TARGET_HOURS}h"] = daily_display["SLA_Compliance"].apply(
            lambda value: f"{value:.1%}" if pd.notna(value) else "—"
        )
        daily_display["Response Score"] = daily_display["SLA_Response_Time_Score"].apply(
            lambda value: f"{value:.1f}" if pd.notna(value) else "—"
        )
        if is_dept_view:
            daily_display = daily_display.rename(columns={
                "Emails_Received": "Received",
                "Items_Handled": "Handled",
                "Available_Hours": "Avail. Hours",
            })
            _show_cols = [
                "Date", "Received", "Handled", "Email SLA Score",
                f"Replied ≤{SLA_TARGET_HOURS}h", "Response Score", "Avail. Hours", "DateLabel"
            ]
        else:
            daily_display["AHT"] = daily_display["AvgHandleSec"].apply(mmss)
            daily_display = daily_display.rename(columns={
                "Items_Handled": "Handled",
                "Available_Hours": "Avail. Hours",
            })
            _show_cols = [
                "Date", "Handled", "AHT", "Email SLA Score",
                f"Replied ≤{SLA_TARGET_HOURS}h", "Response Score", "Avail. Hours"
            ]
        _show_cols = [c for c in _show_cols if c in daily_display.columns]
        st.dataframe(
            daily_display[_show_cols].sort_values("Date", ascending=True),
            use_container_width=True, hide_index=True
        )
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
