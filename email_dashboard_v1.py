import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(
    page_title="Email Performance Dashboard",
    layout="wide"
)

BASE = Path(__file__).parent

FILES = [
    "report1770723329010.csv",
    "report1770723306446.csv"
]

EMAIL_CHANNEL = "casesChannel"


# =====================================================
# ROBUST CSV LOADER (fixes UnicodeDecodeError on Cloud)
# =====================================================

def read_csv_safe(path: Path):
    """
    Handles:
    - UTF-8
    - Windows cp1252 (Excel/Salesforce default)
    - UTF-16 exports
    """
    try:
        return pd.read_csv(
            path,
            encoding="cp1252",
            dayfirst=True,
            parse_dates=["Start DT", "End DT"],
            low_memory=False
        )
    except UnicodeDecodeError:
        return pd.read_csv(
            path,
            encoding="utf-16",
            sep="\t",
            dayfirst=True,
            parse_dates=["Start DT", "End DT"],
            low_memory=False
        )


@st.cache_data
def load_data():
    dfs = []
    for f in FILES:
        p = BASE / f
        df = read_csv_safe(p)
        dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)

    df.columns = df.columns.str.strip()

    return df


# =====================================================
# HELPERS
# =====================================================

def fmt_mmss(sec):
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"


# =====================================================
# LOAD
# =====================================================

df = load_data()

# keep only email channel
df = df[df["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

# duration
df["Duration"] = (df["End DT"] - df["Start DT"]).dt.total_seconds()

df = df.dropna(subset=["Duration"])

df["Date"] = df["Start DT"].dt.date


# =====================================================
# SIDEBAR DATE FILTER
# =====================================================

min_d = df["Date"].min()
max_d = df["Date"].max()

start, end = st.sidebar.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d),
    min_value=min_d,
    max_value=max_d
)

df = df[(df["Date"] >= start) & (df["Date"] <= end)]


# =====================================================
# CORE METRICS
# =====================================================

total_emails = len(df)
aht = df["Duration"].mean()

daily = (
    df.groupby("Date")
      .agg(
          Volume=("Duration", "size"),
          AHT=("Duration", "mean")
      )
      .reset_index()
)

avg_per_day = daily["Volume"].mean() if not daily.empty else 0
peak_day = daily["Volume"].max() if not daily.empty else 0


# =====================================================
# UI
# =====================================================

st.title("Email Performance Dashboard")

c1, c2, c3, c4 = st.columns(4)

c1.metric("Total Emails", f"{total_emails:,}")
c2.metric("Avg Handle Time", fmt_mmss(aht))
c3.metric("Avg Emails / Day", f"{avg_per_day:.0f}")
c4.metric("Peak Day Volume", f"{peak_day:.0f}")

st.markdown("---")


# =====================================================
# CHARTS
# =====================================================

if daily.empty:
    st.info("No data for selected range.")
    st.stop()

vol_chart = (
    alt.Chart(daily)
    .mark_line(point=True)
    .encode(
        x=alt.X("Date:T", title="Date"),
        y=alt.Y("Volume:Q", title="Email Volume"),
        tooltip=["Date:T", "Volume"]
    )
    .properties(height=300)
)

aht_chart = (
    alt.Chart(daily)
    .mark_line(point=True)
    .encode(
        x=alt.X("Date:T", title="Date"),
        y=alt.Y("AHT:Q", title="Avg Handle Time (sec)"),
        tooltip=["Date:T", alt.Tooltip("AHT:Q", format=".0f")]
    )
    .properties(height=300)
)

st.altair_chart(vol_chart, use_container_width=True)
st.altair_chart(aht_chart, use_container_width=True)
