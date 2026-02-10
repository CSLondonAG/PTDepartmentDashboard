import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")

BASE = Path(__file__).parent

FILES = [
    "report1770723329010.csv",
    "report1770723306446.csv"
]

@st.cache_data
def load_data():
    dfs = []
    for f in FILES:
        df = pd.read_csv(BASE / f, dayfirst=True, parse_dates=["Start DT","End DT"])
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

df = load_data()

df = df[df["Service Channel: Developer Name"] == "casesChannel"].copy()

df["Duration"] = (df["End DT"] - df["Start DT"]).dt.total_seconds()
df["Date"] = df["Start DT"].dt.date

min_d = df["Date"].min()
max_d = df["Date"].max()

start, end = st.sidebar.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d),
    min_value=min_d,
    max_value=max_d
)

df = df[(df["Date"] >= start) & (df["Date"] <= end)]

total_emails = len(df)
aht = df["Duration"].mean()

daily = (
    df.groupby("Date")
      .agg(
          Volume=("Duration","size"),
          AHT=("Duration","mean")
      )
      .reset_index()
)

def fmt(sec):
    if pd.isna(sec):
        return "—"
    m, s = divmod(int(sec),60)
    return f"{m:02}:{s:02}"

st.title("Email Performance Dashboard")

c1,c2,c3,c4 = st.columns(4)

c1.metric("Total Emails", f"{total_emails:,}")
c2.metric("Avg Handle Time", fmt(aht))
c3.metric("Avg / Day", f"{daily['Volume'].mean():.0f}")
c4.metric("Peak Day", f"{daily['Volume'].max():.0f}")

st.markdown("---")

vol_chart = (
    alt.Chart(daily)
    .mark_line(point=True)
    .encode(x="Date:T", y="Volume:Q")
    .properties(height=300)
)

aht_chart = (
    alt.Chart(daily)
    .mark_line(point=True)
    .encode(x="Date:T", y="AHT:Q")
    .properties(height=300)
)

st.altair_chart(vol_chart, use_container_width=True)
st.altair_chart(aht_chart, use_container_width=True)
