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

# ── KPI status colour constants ──
COLOR_GOOD = "#00FF66"
COLOR_WARN = "#FAE100"
COLOR_BAD = "#FF9300"
COLOR_NONE = "#CFE3D6"

# ── SLA colour thresholds ──
SLA_GOOD = 80
SLA_WARN = 60

BRAND_FONT = "'Proxima Nova', 'Segoe UI', system-ui, -apple-system, sans-serif"

# ── THEME CSS ──────────────────────────────────────────────────

st.markdown(
    """
    <style>
      /* ── Brand palette ── */
      :root {
        --pb-bet-green: #009A44;
        --pb-bright-green: #00FF66;
        --pb-dark-green: #003806;
        --pb-dark-grey: #111111;
        --pb-bet-yellow: #FAE100;
        --pb-bet-orange: #FF9300;
        --pb-white: #FFFFFF;

        --glass-bg: linear-gradient(145deg, rgba(255,255,255,0.18) 0%, rgba(255,255,255,0.04) 100%);
        --glass-tint: rgba(0, 56, 6, 0.38);
        --glass-border: rgba(255,255,255,0.24);
        --glass-highlight: rgba(255,255,255,0.45);
        --glass-shadow: 0 26px 50px -20px rgba(0, 20, 2, 0.70);
        --glass-blur: blur(22px) saturate(150%);

        --radius-lg: 28px;
        --radius-md: 20px;
      }

      /* ── Font faces ── */
      @font-face { font-family: 'Proxima Nova'; src: url('/app/static/fonts/ProximaNova-light.woff2') format('woff2'); font-weight: 300; font-style: normal; }
      @font-face { font-family: 'Proxima Nova'; src: url('/app/static/fonts/ProximaNova-regular.woff2') format('woff2'); font-weight: 400; font-style: normal; }
      @font-face { font-family: 'Proxima Nova'; src: url('/app/static/fonts/ProximaNova-medium.woff2') format('woff2'); font-weight: 500; font-style: normal; }
      @font-face { font-family: 'Proxima Nova'; src: url('/app/static/fonts/ProximaNova-semibold.woff2') format('woff2'); font-weight: 600; font-style: normal; }
      @font-face { font-family: 'Proxima Nova'; src: url('/app/static/fonts/ProximaNova-bold.woff2') format('woff2'); font-weight: 700; font-style: normal; }

      /* ── Backdrop ── */
      .stApp {
        background:
          radial-gradient(ellipse 60% 50% at 20% 10%, rgba(0,154,68,0.40), transparent 70%),
          radial-gradient(ellipse 50% 60% at 80% 20%, rgba(0,255,102,0.16), transparent 65%),
          radial-gradient(ellipse 70% 45% at 50% 80%, rgba(0,154,68,0.30), transparent 70%),
          radial-gradient(ellipse 40% 55% at 10% 60%, rgba(0,56,6,0.60), transparent 60%),
          radial-gradient(ellipse 55% 50% at 90% 70%, rgba(0,154,68,0.25), transparent 60%),
          radial-gradient(ellipse 45% 40% at 50% 30%, rgba(0,255,102,0.12), transparent 55%),
          linear-gradient(170deg, #003806 0%, #001a03 40%, #002d05 100%) !important;
        background-attachment: fixed !important;
        font-family: 'Proxima Nova', 'Segoe UI', system-ui, -apple-system, sans-serif;
      }

      /* ── Typography ── */
      .stMarkdown h1 { color: var(--pb-white); font-size: 1.75rem; font-weight: 500; margin-bottom: 2px; }
      .stMarkdown h2 { color: var(--pb-bet-yellow); font-size: 1.1rem; font-weight: 600; margin-top: 2rem; letter-spacing: 0.03em; }
      .stMarkdown h3 { color: rgba(255,255,255,0.80); font-size: 0.95rem; font-weight: 500; }
      .stMarkdown p, .stMarkdown li { color: rgba(255,255,255,0.85); }

      /* ── Glass surface mixin ── */
      .glass-tile,
      .glass-hero,
      div[data-testid="stMetric"],
      div[data-testid="stExpander"] details,
      div[data-testid="stVegaLiteChart"],
      div[data-testid="stArrowVegaLiteChart"] {
        background: var(--glass-tint);
        background-image: var(--glass-bg);
        border: 1px solid var(--glass-border);
        box-shadow: var(--glass-shadow), inset 0 1px 0 var(--glass-highlight);
        -webkit-backdrop-filter: var(--glass-blur);
        backdrop-filter: var(--glass-blur);
      }

      /* ── Hero header ── */
      .glass-hero {
        border-radius: var(--radius-lg);
        padding: 32px 36px 28px;
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .glass-hero::before {
        content: '';
        position: absolute; inset: 0;
        background: radial-gradient(ellipse 60% 80% at 80% 20%, rgba(0,255,102,0.12), transparent 65%);
        pointer-events: none;
      }
      .hero-kicker {
        color: var(--pb-bet-yellow);
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.18em;
        text-transform: uppercase;
        margin-bottom: 4px;
      }
      .hero-title {
        color: var(--pb-white);
        font-size: 2.2rem;
        font-weight: 500;
        line-height: 1.15;
        margin: 0 0 6px;
      }
      .hero-sub {
        color: rgba(255,255,255,0.55);
        font-size: 0.82rem;
        font-weight: 400;
      }
      .hero-icon {
        position: absolute; top: 24px; left: 28px;
        color: rgba(255,255,255,0.30);
      }

      /* ── KPI tiles ── */
      .glass-tile {
        border-radius: var(--radius-lg);
        padding: 22px 24px 18px;
        position: relative;
        overflow: hidden;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
      }
      .glass-tile:hover {
        transform: translateY(-3px);
        box-shadow: 0 32px 60px -20px rgba(0, 20, 2, 0.85), inset 0 1px 0 var(--glass-highlight);
      }
      .glass-tile::before {
        content: '';
        position: absolute;
        top: -30%; right: -20%;
        width: 60%; height: 80%;
        background: radial-gradient(circle, var(--accent, rgba(0,255,102,0.10)) 0%, transparent 70%);
        pointer-events: none;
        opacity: 0.5;
      }
      .tile-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
      .tile-icon {
        width: 36px; height: 36px;
        display: flex; align-items: center; justify-content: center;
        background: rgba(255,255,255,0.08);
        border-radius: 50%;
        color: rgba(255,255,255,0.70);
        flex-shrink: 0;
      }
      .tile-label {
        color: rgba(255,255,255,0.65);
        font-size: 0.72rem;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        flex-grow: 1;
      }
      .tile-dot {
        width: 10px; height: 10px;
        border-radius: 50%;
        box-shadow: 0 0 8px 2px var(--accent, rgba(0,255,102,0.40));
        flex-shrink: 0;
      }
      .tile-value {
        color: var(--pb-white);
        font-size: 2rem;
        font-weight: 500;
        line-height: 1.1;
        margin-bottom: 2px;
      }
      .tile-sub {
        color: rgba(255,255,255,0.45);
        font-size: 0.72rem;
        font-weight: 400;
      }

      /* ── Native metric override ── */
      div[data-testid="stMetricValue"] { color: var(--pb-white); font-size: 1.5rem; font-weight: 500; }
      div[data-testid="stMetricLabel"] { color: rgba(255,255,255,0.65); font-size: 0.72rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em; }
      div[data-testid="stMetric"] { border-radius: var(--radius-md); padding: 16px 20px; }

      /* ── Buttons ── */
      .stButton > button {
        background-color: var(--pb-bet-yellow) !important;
        color: var(--pb-dark-grey) !important;
        border-radius: 999px !important;
        border: none !important;
        padding: 8px 22px;
        font-weight: 600;
        box-shadow: 0 8px 24px -6px rgba(250,225,0,0.30);
        transition: transform 0.15s ease, box-shadow 0.15s ease, filter 0.15s ease;
      }
      .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 12px 32px -6px rgba(250,225,0,0.40);
        filter: brightness(1.05);
        color: var(--pb-dark-grey) !important;
        border: none !important;
      }
      .stButton > button:active {
        transform: scale(0.97);
        transition: transform 0.08s ease;
      }

      /* ── Tabs (pill segmented control) ── */
      div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: var(--glass-tint);
        background-image: var(--glass-bg);
        border: 1px solid var(--glass-border);
        border-radius: 999px;
        padding: 4px;
        gap: 4px;
        -webkit-backdrop-filter: var(--glass-blur);
        backdrop-filter: var(--glass-blur);
      }
      div[data-testid="stTabs"] [data-baseweb="tab"] {
        border-radius: 999px;
        padding: 6px 20px;
        color: rgba(255,255,255,0.65);
        font-weight: 500;
        background: transparent;
        border: none;
      }
      div[data-testid="stTabs"] [aria-selected="true"] {
        background: var(--pb-bet-yellow) !important;
        color: var(--pb-dark-grey) !important;
        font-weight: 600;
      }
      div[data-testid="stTabs"] [data-baseweb="tab-highlight"],
      div[data-testid="stTabs"] [data-baseweb="tab-border"] {
        display: none !important;
      }

      /* ── Inputs ── */
      div[data-testid="stDateInput"] input,
      div[data-testid="stSelectbox"] [data-baseweb="select"],
      div[data-testid="stSlider"] {
        background: rgba(255,255,255,0.10) !important;
        border: 1px solid var(--glass-border) !important;
        border-radius: 14px !important;
        color: var(--pb-white) !important;
      }

      /* ── Charts ── */
      div[data-testid="stVegaLiteChart"],
      div[data-testid="stArrowVegaLiteChart"] {
        border-radius: 24px;
        overflow: hidden;
        padding: 8px;
      }

      /* ── Vega tooltip ── */
      #vg-tooltip-element.vg-tooltip {
        background: rgba(0,56,6,0.90) !important;
        border: 1px solid rgba(255,255,255,0.20) !important;
        border-radius: 12px !important;
        color: white !important;
        font-family: 'Proxima Nova', 'Segoe UI', system-ui, sans-serif !important;
        -webkit-backdrop-filter: blur(16px);
        backdrop-filter: blur(16px);
        box-shadow: 0 12px 32px rgba(0,0,0,0.30) !important;
      }
      #vg-tooltip-element.vg-tooltip td.key { color: rgba(255,255,255,0.60) !important; }
      #vg-tooltip-element.vg-tooltip td.value { color: white !important; font-weight: 500; }

      /* ── Tables ── */
      div[data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid var(--glass-border);
        -webkit-backdrop-filter: var(--glass-blur);
        backdrop-filter: var(--glass-blur);
      }

      /* ── Expanders ── */
      div[data-testid="stExpander"] details {
        border-radius: var(--radius-md);
        border: 1px solid var(--glass-border);
      }
      div[data-testid="stExpander"] summary {
        color: rgba(255,255,255,0.80);
        font-weight: 500;
      }

      /* ── Alerts ── */
      div[data-testid="stAlert"] {
        border-radius: 18px;
        background: rgba(0,56,6,0.50);
        border: 1px solid var(--glass-border);
        color: rgba(255,255,255,0.85);
        -webkit-backdrop-filter: blur(12px);
        backdrop-filter: blur(12px);
      }

      /* ── Sidebar ── */
      section[data-testid="stSidebar"] {
        background: rgba(0,56,6,0.55) !important;
        -webkit-backdrop-filter: blur(30px) saturate(140%);
        backdrop-filter: blur(30px) saturate(140%);
        border-right: 1px solid var(--glass-border);
      }
      section[data-testid="stSidebar"] > div { background: transparent !important; }

      /* ── Focus ring ── */
      *:focus-visible { outline: 2px solid var(--pb-bet-yellow) !important; outline-offset: 2px !important; }

      /* ── Spinner ── */
      div[data-testid="stSpinner"] > div { color: var(--pb-bright-green); }

      /* ── Divider ── */
      hr { border-color: rgba(255,255,255,0.12); border-width: 1px 0 0 0; margin: 1.5rem 0; }

      /* ── Captions ── */
      .stCaption, div[data-testid="stCaptionContainer"] { color: rgba(255,255,255,0.45) !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── SVG ICON SYSTEM ──────────────────────────────────────────

_ICON_PATHS = {
    "mail": "M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z M22 6l-10 7L2 6",
    "check": "M9 12l2 2 4-4 M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z",
    "timer": "M12 6v6l4 2 M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z M10 2h4",
    "clock": "M12 6v6l4 2 M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z",
    "target": "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z M12 6a6 6 0 1 0 0 12 6 6 0 0 0 0-12z M12 10a2 2 0 1 0 0 4 2 2 0 0 0 0-4z",
    "gauge": "M12 15a1 1 0 1 0 0-2 1 1 0 0 0 0 2z M12 15l3.5-5 M5.6 10.6a8 8 0 0 1 12.8 0 M3.3 13.4a10 10 0 0 1 17.4 0",
    "bars": "M12 20V10 M18 20V4 M6 20v-4",
    "pie": "M21.21 15.89A10 10 0 1 1 8 2.83 M22 12A10 10 0 0 0 12 2v10z",
    "calendar": "M19 4H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2z M16 2v4 M8 2v4 M3 10h18 M9 16l2 2 4-4",
    "zap": "M13 2L3 14h9l-1 8 10-12h-9l1-8z",
    "spark": "M12 2l2.4 7.2L22 12l-7.6 2.8L12 22l-2.4-7.2L2 12l7.6-2.8z",
    "layers": "M12 2l10 6-10 6L2 8z M2 12l10 6 10-6 M2 16l10 6 10-6",
    "chart": "M18 20V10 M12 20V4 M6 20v-6",
}

_TILE_ICONS = {
    "Emails Received": "mail",
    "Work Items Handled": "check",
    "Avg Response Time (BH)": "clock",
    "Email SLA Score (60/40)": "target",
    "Avg Handle Time": "timer",
    "Online Hours": "gauge",
    "Available Hours": "calendar",
    "Utilisation": "pie",
    "Replied": "check",
    "Avg Response Time": "clock",
}


def _icon_svg(name, size=20):
    path_data = _ICON_PATHS.get(name, _ICON_PATHS["spark"])
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
        f'stroke="currentColor" stroke-width="1.6" stroke-linecap="round" '
        f'stroke-linejoin="round">'
        f'<path d="{path_data}"/></svg>'
    )


def _get_icon_name(label):
    for key, icon in _TILE_ICONS.items():
        if key.lower() in label.lower():
            return icon
    return "spark"


def get_sla_color(value):
    if pd.isna(value):
        return COLOR_NONE
    if value >= SLA_GOOD:
        return COLOR_GOOD
    if value >= SLA_WARN:
        return COLOR_WARN
    return COLOR_BAD


def render_custom_metric(container, title, value, tooltip="", color=COLOR_NONE, sub=""):
    icon_name = _get_icon_name(title)
    icon_html = _icon_svg(icon_name)
    sub_html = f'<div class="tile-sub">{sub}</div>' if sub else ""
    tooltip_attr = f'title="{tooltip}"' if tooltip else ""
    container.markdown(
        f"""
        <div class="glass-tile" style="--accent: {color};" {tooltip_attr}>
            <div class="tile-head">
                <div class="tile-icon">{icon_html}</div>
                <div class="tile-label">{title}</div>
                <div class="tile-dot" style="background: {color};"></div>
            </div>
            <div class="tile-value">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def themed_altair_chart(chart, **kwargs):
    """Render an Altair chart with transparent background and themed axes."""
    configured = chart.configure(
        background="transparent",
    ).configure_axis(
        labelColor="rgba(255,255,255,0.70)",
        titleColor="rgba(255,255,255,0.80)",
        gridColor="rgba(255,255,255,0.14)",
        domainColor="rgba(255,255,255,0.30)",
        tickColor="rgba(255,255,255,0.30)",
        labelFont=BRAND_FONT,
        titleFont=BRAND_FONT,
    ).configure_legend(
        labelColor="rgba(255,255,255,0.75)",
        titleColor="rgba(255,255,255,0.85)",
        labelFont=BRAND_FONT,
        titleFont=BRAND_FONT,
    ).configure_title(
        color="white",
        font=BRAND_FONT,
    ).configure_view(
        strokeWidth=0,
    )
    st.altair_chart(configured, **kwargs)


# ── DATA LOADING & HELPERS (unchanged logic) ──────────────────


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

# Hero header
_chart_icon = _icon_svg("chart", 22)
st.markdown(
    f"""
    <div class="glass-hero">
        <div class="hero-icon">{_chart_icon}</div>
        <div class="hero-kicker">CUSTOMER SERVICE</div>
        <div class="hero-title">Email Department Performance</div>
        <div class="hero-sub">SLA model: {SLA_TARGET_HOURS} business hours · normalized 60/40 score</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Refresh button
refresh_col_spacer, refresh_col = st.columns([5, 1])
with refresh_col:
    if st.button("↻ Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Previous completed week (Mon-Sun)
today = pd.Timestamp.now().date()
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - pd.Timedelta(days=days_since_sunday if days_since_sunday > 0 else 7)
week_start = last_sunday - pd.Timedelta(days=6)

_email_activity_dates = pd.concat(
    [email_rec["Date_Opened"], email_rec["Date_Completed"]],
    ignore_index=True,
).dropna()
default_start = max(week_start, _email_activity_dates.min())
default_end = min(last_sunday, _email_activity_dates.max())

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

# Received-volume metrics are attributed to the date the email was opened.
email_received_period = email_rec[
    (email_rec["Date_Opened"] >= start) & (email_rec["Date_Opened"] <= end)
].copy()

# SLA metrics are attributed to the date the email was responded to/completed.
email_responded_period = email_rec[
    email_rec["CompletedDT"].notna()
    & (email_rec["Date_Completed"] >= start)
    & (email_rec["Date_Completed"] <= end)
].copy()
case_cat_period = case_cat[(case_cat["Date_Opened"] >= start) & (case_cat["Date_Opened"] <= end)].copy()
items_period = items[(items["Date_Closed"] >= start) & (items["Date_Closed"] <= end)].copy()

start_ts = pd.Timestamp(start)
end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)

# Apply agent filter where data supports it
if not is_dept_view:
    items_period = items_period[items_period["User: Full Name"].astype(str) == selected_agent].copy()
    _agent_key = _parse_name(selected_agent)
    if _email_agent_col:
        email_received_period = email_received_period[
            email_received_period[_email_agent_col].astype(str).apply(_parse_name) == _agent_key
        ].copy()
        email_responded_period = email_responded_period[
            email_responded_period[_email_agent_col].astype(str).apply(_parse_name) == _agent_key
        ].copy()
    if _case_cat_agent_col:
        case_cat_period = case_cat_period[
            case_cat_period[_case_cat_agent_col].astype(str).apply(_parse_name) == _agent_key
        ].copy()


# ---------------- METRICS ----------------

total_received = email_received_period["OpenedDT"].notna().sum()
total_handled = items_period["CloseDT"].notna().sum()

completed_emails = email_responded_period.copy()
if len(completed_emails) > 0:
    completed_emails["ResponseTimeBusinessSec"] = completed_emails.apply(
        lambda r: business_seconds_between(r["OpenedDT"], r["CompletedDT"]), axis=1
    )
else:
    completed_emails["ResponseTimeBusinessSec"] = pd.Series(dtype=float)

# Normalized Email SLA Score
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

# Calculate one SLA score per response/completion day, then volume-weight the
# daily scores across the selected period.
_daily_sla_rows = []
if sla_eligible > 0:
    for _sla_date, _sla_group in valid_completed_emails.groupby("Date_Completed", dropna=True):
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
_busy_mask = _presence_status.str.casefold().str.startswith("busy")

pres_avail = pres_in_window[_presence_status.isin(AVAILABLE_STATUSES)].copy()
pres_online = pres_in_window[~_offline_mask].copy()
pres_util_base = pres_in_window[~_offline_mask & ~_busy_mask].copy()

available_sec = seconds_in_window(pres_avail, start_ts, end_ts)
available_hours = available_sec / 3600

online_sec = seconds_in_window(pres_online, start_ts, end_ts)
online_hours = online_sec / 3600

util_presence_sec = seconds_in_window(pres_util_base, start_ts, end_ts)

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

email_invalid_open = email_received_period["OpenedDT"].isna().sum()
email_missing_or_invalid_complete = email_received_period["CompletedDT"].isna().sum()
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


# ---------------- DISPLAY ────────────────────────────────────

_view_label = selected_agent if not is_dept_view else "Department"
st.markdown(
    f"<p style='color:rgba(255,255,255,0.45);margin-top:-8px;margin-bottom:20px;font-size:0.85rem;'>"
    f"{start} — {end} · {_view_label}</p>",
    unsafe_allow_html=True,
)

# ── Primary metrics (top tier) ──
sla_display = f"{email_sla_score:.1f}" if pd.notna(email_sla_score) else "—"
sla_compliance_display = f"{sla_compliance:.1%}" if pd.notna(sla_compliance) else "—"
sla_help = (
    f"Response-date, daily-volume weighted 0–100 Email SLA Score: 60% from the share replied "
    f"within {SLA_TARGET_HOURS} business hours and 40% from normalized average "
    f"response time."
)
sla_color = get_sla_color(email_sla_score)

if is_dept_view:
    c1, c2, c3, c4 = st.columns(4)
    render_custom_metric(c1, "Emails Received", f"{total_received:,}", color=COLOR_GOOD)
    render_custom_metric(c2, "Work Items Handled", f"{total_handled:,}", color=COLOR_GOOD)
    render_custom_metric(c3, "Avg Response Time (BH)", hm(avg_art), color=COLOR_NONE)
    render_custom_metric(c4, "Email SLA Score (60/40)", sla_display, tooltip=sla_help, color=sla_color)
else:
    c1, c2, c3 = st.columns(3)
    render_custom_metric(c1, "Work Items Handled", f"{total_handled:,}", color=COLOR_GOOD)
    render_custom_metric(c2, "Avg Response Time (BH)", hm(avg_art), color=COLOR_NONE)
    render_custom_metric(c3, "Email SLA Score (60/40)", sla_display, tooltip=sla_help, color=sla_color)

# ── Secondary metrics (contextual) ──
st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
s1, s2, s3, s4 = st.columns(4)
render_custom_metric(s1, "Avg Handle Time", mmss(avg_aht), color=COLOR_NONE)
render_custom_metric(
    s2, "Online Hours", f"{online_hours:.1f}",
    tooltip="Total presence time excluding Offline status.",
    color=COLOR_NONE,
)
render_custom_metric(
    s3, "Available Hours", f"{available_hours:.1f}",
    tooltip="Time in Available_Email_and_Web or Available_All status.",
    color=COLOR_NONE,
)
render_custom_metric(
    s4, "Utilisation", f"{util:.1%}",
    tooltip="Email handle time ÷ matched presence time (excl. Offline & Busy*).",
    color=COLOR_NONE,
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
daily_received = email_received_period.groupby("Date_Opened").size().reset_index(name="Emails_Received")
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
        color_range = ["#009A44", "#00FF66", "#FAE100"]

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
        dow_bar_labels = alt.Chart(dow_counts_long).mark_text(dy=-8, fontSize=10, color="white").encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("AverageCount:Q"),
            xOffset="Metric:N",
            text=alt.Text("AverageCount:Q", format=",.0f"),
        )
        dow_hours_line = alt.Chart(dow).mark_line(
            point=alt.OverlayMarkDef(filled=True, size=70), color="#FAE100", strokeWidth=3
        ).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            tooltip=["DoW", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
        )
        dow_hours_labels = alt.Chart(dow).mark_text(dy=-10, color="#FAE100", fontSize=10).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            text=alt.Text("Available_Hours:Q", format=".1f"),
        )
        dow_chart = alt.layer(dow_bar, dow_bar_labels, dow_hours_line, dow_hours_labels).properties(height=340)
        themed_altair_chart(dow_chart, use_container_width=True)
        st.markdown(
            """
            <div style="display:flex;gap:24px;justify-content:center;margin-top:-8px;margin-bottom:8px;">
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:rgba(255,255,255,0.70);">
                    <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#009A44;"></span>
                    Emails Received
                </span>
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:rgba(255,255,255,0.70);">
                    <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#00FF66;"></span>
                    Items Handled
                </span>
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:rgba(255,255,255,0.70);">
                    <span style="display:inline-block;width:28px;height:3px;background:#FAE100;border-radius:2px;position:relative;top:0px;"></span>
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
            color="#00FF66", cornerRadiusTopLeft=4, cornerRadiusTopRight=4
        ).encode(
            x=alt.X("DoWShort:N", title="Day of Week", sort=dow["DoWShort"].tolist(),
                    axis=alt.Axis(labelAngle=0, labelPadding=6)),
            y=alt.Y("Items_Handled:Q", title="Avg Items Handled",
                    axis=alt.Axis(format=".0f", titlePadding=12)),
            tooltip=["DoW", alt.Tooltip("Items_Handled:Q", format=",.1f", title="Avg Handled")],
        )
        agent_bar_labels = alt.Chart(dow).mark_text(dy=-8, fontSize=11, color="#00FF66").encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Items_Handled:Q"),
            text=alt.Text("Items_Handled:Q", format=",.0f"),
        )
        agent_hours_line = alt.Chart(dow).mark_line(
            point=alt.OverlayMarkDef(filled=True, size=70), color="#FAE100", strokeWidth=3
        ).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            tooltip=["DoW", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
        )
        agent_hours_labels = alt.Chart(dow).mark_text(dy=-10, color="#FAE100", fontSize=10).encode(
            x=alt.X("DoWShort:N", sort=dow["DoWShort"].tolist()),
            y=alt.Y("Available_Hours_Scaled:Q", axis=None),
            text=alt.Text("Available_Hours:Q", format=".1f"),
        )
        dow_chart = alt.layer(agent_bar, agent_bar_labels, agent_hours_line, agent_hours_labels).properties(height=340)
        themed_altair_chart(dow_chart, use_container_width=True)
        st.markdown(
            """
            <div style="display:flex;gap:24px;justify-content:center;margin-top:-8px;margin-bottom:8px;">
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:rgba(255,255,255,0.70);">
                    <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#00FF66;"></span>
                    Items Handled
                </span>
                <span style="display:flex;align-items:center;gap:6px;font-size:0.82rem;color:rgba(255,255,255,0.70);">
                    <span style="display:inline-block;width:28px;height:3px;background:#FAE100;border-radius:2px;position:relative;top:0px;"></span>
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
    render_custom_metric(sla_m1, "Email SLA Score (60/40)", sla_display, tooltip=sla_help, color=sla_color)
    render_custom_metric(
        sla_m2, f"Replied ≤{SLA_TARGET_HOURS}h", sla_compliance_display,
        tooltip=f"{sla_met:,} of {sla_eligible:,} eligible emails met the target.",
        color=get_sla_color(email_sla_score),
    )
    render_custom_metric(
        sla_m3, "Avg Response Time (BH)", hm(avg_art),
        tooltip=f"Response score: {weighted_response_time_score:.1f}/100.",
        color=COLOR_NONE,
    )

    if not _daily_sla.empty:
        _sla_trend = _daily_sla.dropna(subset=["Date", "Email_SLA_Score"]).copy()

        _sla_tick_dates = (
            _sla_trend["Date"]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        _trend_line = alt.Chart(_sla_trend).mark_line(
            point=alt.OverlayMarkDef(filled=True, size=70), color="#00FF66", strokeWidth=3
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
            dy=-10, color="#00FF66", fontSize=10
        ).encode(
            x=alt.X("Date:T", axis=None),
            y=alt.Y("Email_SLA_Score:Q"),
            text=alt.Text("Email_SLA_Score:Q", format=".1f"),
        )
        _target_rule = alt.Chart(pd.DataFrame({"Target": [80]})).mark_rule(
            color="#FF9300", strokeDash=[5, 5]
        ).encode(y="Target:Q")
        themed_altair_chart(
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
                range=["#00FF66", "#FF9300"],
            ),
        ),
        tooltip=["Status", alt.Tooltip("Count:Q", format=",")],
    )
    sla_status_labels = alt.Chart(sla_status_summary).mark_text(
        dy=-10,
        fontSize=11,
        color="white",
    ).encode(
        x=alt.X("Status:N", sort=sla_status_summary["Status"].tolist()),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
    )

    closed_aging_bars = alt.Chart(closed_aging_summary).mark_bar(
        color="#009A44", cornerRadiusTopLeft=4, cornerRadiusTopRight=4
    ).encode(
        x=alt.X("Bucket:N", title="Business-hour Response Bucket", sort=aging_labels),
        y=alt.Y("Count:Q", title="Completed Email Count"),
        tooltip=["Bucket", "Count"],
    )
    closed_aging_labels = alt.Chart(closed_aging_summary).mark_text(
        dy=-10, color="#00FF66", fontSize=11
    ).encode(
        x=alt.X("Bucket:N", sort=aging_labels),
        y=alt.Y("Count:Q"),
        text=alt.Text("Count:Q", format=","),
    )

    sla_chart_col, aging_chart_col = st.columns([1, 2])
    with sla_chart_col:
        themed_altair_chart(
            alt.layer(sla_status_chart, sla_status_labels).properties(height=340),
            use_container_width=True,
        )
    with aging_chart_col:
        themed_altair_chart(
            alt.layer(closed_aging_bars, closed_aging_labels).properties(height=340),
            use_container_width=True,
        )

    _sla_note = "" if is_dept_view or _email_agent_col else " Showing department-level data (no agent column detected in email export)."
    st.caption(
        f"Email SLA Score is a normalized 60/40 measure. Compliance contributes up to 60 points "
        f"({SLA_COMPLIANCE_WEIGHT:.0%} × the percentage replied within {SLA_TARGET_HOURS} business hours). "
        f"Average response performance contributes up to 40 points "
        f"({SLA_RESPONSE_TIME_WEIGHT:.0%} × min(100, {SLA_TARGET_HOURS} ÷ average response hours × 100)). "
        f"Daily response-date scores are weighted by eligible responded-email volume. Business hours are "
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
        color=alt.Color("Reason:N", title="Reason", sort=reason_sort,
                        scale=alt.Scale(scheme="greens")),
        order=alt.Order("Count:Q", sort="descending"),
        tooltip=[
            "Category",
            "Reason",
            alt.Tooltip("Count:Q", format=","),
            alt.Tooltip("CategoryTotal:Q", format=","),
        ],
    )

    cat_labels = alt.Chart(chart_data[["Category", "CategoryTotal"]].drop_duplicates()).mark_text(
        dx=6, color="#00FF66", fontSize=10
    ).encode(
        x=alt.X("CategoryTotal:Q"),
        y=alt.Y("Category:N", sort=category_sort),
        text=alt.Text("CategoryTotal:Q", format=","),
    )

    stacked_chart = alt.layer(cat_bars, cat_labels).properties(
        height=min(max(340, len(category_sort) * 26), 600)
    )
    themed_altair_chart(stacked_chart, use_container_width=True)
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
    themed_altair_chart(heatmap, use_container_width=True)
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
            color="#00FF66", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_handled["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("Items_Handled:Q", title="Items Handled",
                    axis=alt.Axis(format=".0f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("Items_Handled:Q", format=",", title="Items Handled")],
        )
        handled_labels = alt.Chart(agent_handled).mark_text(
            dx=6, fontSize=11, color="#00FF66", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_handled["Agent"].tolist()),
            x=alt.X("Items_Handled:Q"),
            text=alt.Text("Items_Handled:Q", format=","),
        )
        themed_altair_chart(
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
            color="#009A44", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_aht["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("AHT_minutes:Q", title="Avg Handle Time (mins)",
                    axis=alt.Axis(format=".1f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("AHT_label:N", title="AHT (mm:ss)")],
        )
        aht_labels = alt.Chart(agent_aht).mark_text(
            dx=6, fontSize=11, color="#00FF66", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_aht["Agent"].tolist()),
            x=alt.X("AHT_minutes:Q"),
            text=alt.Text("AHT_label:N"),
        )
        themed_altair_chart(
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
            color="#FAE100", cornerRadiusTopRight=4, cornerRadiusBottomRight=4
        ).encode(
            y=alt.Y("Agent:N", title=None, sort=agent_avail_df["Agent"].tolist(),
                    axis=alt.Axis(labelLimit=200, labelFontSize=12)),
            x=alt.X("Available_Hours:Q", title="Available Hours",
                    axis=alt.Axis(format=".1f", titlePadding=10)),
            tooltip=["Agent", alt.Tooltip("Available_Hours:Q", format=".1f", title="Avail. Hours")],
        )
        avail_labels = alt.Chart(agent_avail_df).mark_text(
            dx=6, fontSize=11, color="#FAE100", align="left"
        ).encode(
            y=alt.Y("Agent:N", sort=agent_avail_df["Agent"].tolist()),
            x=alt.X("Available_Hours:Q"),
            text=alt.Text("Available_Hours:Q", format=".1f"),
        )
        themed_altair_chart(
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
        "<p style='color:rgba(255,255,255,0.45);font-size:0.82rem;margin-bottom:8px;'>Rows excluded due to unparseable timestamps — review source data if counts are high.</p>",
        unsafe_allow_html=True,
    )
    dq1, dq2, dq3 = st.columns(3)
    dq1.metric("Invalid Opened Timestamps", f"{email_invalid_open:,}")
    dq2.metric(
        "Received Without Valid Completion",
        f"{email_missing_or_invalid_complete:,}",
    )
    dq3.metric("Invalid Item Close Timestamps", f"{items_invalid_close:,}")
