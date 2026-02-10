import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

st.set_page_config(layout="wide")

BASE = Path(__file__).parent

ITEMS_FILE = "ItemsPT.csv"
PRES_FILE  = "PresencePT.csv"
ART_FILE   = "ART PT.csv"

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE CSV
# =====================================================

def read_csv_safe(path):
    try:
        return pd.read_csv(path, encoding="cp1252")
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t")


# =====================================================
# HELPERS
# =====================================================

def fmt(sec):
    if pd.isna(sec):
        return "—"
    m,s = divmod(int(sec),60)
    return f"{m:02}:{s:02}"

def clip(s,e,ws,we):
    if pd.isna(s) or pd.isna(e):
        return None
    s2,e2=max(s,ws),min(e,we)
    return (s2,e2) if e2>s2 else None

def sum_seconds(iv):
    return sum((e-s).total_seconds() for s,e in iv if s is not None)


def find_col(cols, keys):
    lower={c.lower():c for c in cols}
    for k in keys:
        for lc,orig in lower.items():
            if k in lc:
                return orig
    return None


# =====================================================
# LOAD
# =====================================================

items = read_csv_safe(BASE / ITEMS_FILE)
pres  = read_csv_safe(BASE / PRES_FILE)
art   = read_csv_safe(BASE / ART_FILE)

for df in (items,pres,art):
    df.columns=df.columns.str.strip()


# =====================================================
# -------- AHT (ItemsPT)
# =====================================================

items = items[items["Service Channel: Developer Name"]==EMAIL_CHANNEL]

items["HandleSec"]=pd.to_numeric(items.get("Handle Time"),errors="coerce")

assign_dt = pd.to_datetime(
    items.get("Assign Date","").astype(str)+" "+items.get("Assign Time","").astype(str),
    errors="coerce",
    dayfirst=True
)

items["Date"]=pd.to_datetime(assign_dt, errors="coerce")   # FORCE dtype


# =====================================================
# -------- RESPONSE (ART PT)
# =====================================================

open_col  = find_col(art.columns,["opened","received","created"])
close_col = find_col(art.columns,["closed","resolved","completed"])

if open_col and close_col:
    art[open_col]  = pd.to_datetime(art[open_col],errors="coerce",dayfirst=True)
    art[close_col] = pd.to_datetime(art[close_col],errors="coerce",dayfirst=True)
    art["ResponseSec"]=(art[close_col]-art[open_col]).dt.total_seconds()
    art["Date"]=pd.to_datetime(art[open_col],errors="coerce")   # FORCE dtype
else:
    art["ResponseSec"]=pd.NA
    art["Date"]=pd.NaT


# =====================================================
# DATE RANGE (dtype-safe)
# =====================================================

items_dates = items["Date"].dropna()
art_dates   = art["Date"].dropna()

all_dates = pd.concat([items_dates, art_dates])

if all_dates.empty:
    st.error("No valid timestamps in data.")
    st.stop()

min_d = all_dates.min().date()
max_d = all_dates.max().date()

start,end = st.date_input("Date Range", value=(max_d-pd.Timedelta(days=6),max_d))

start_ts = pd.Timestamp(start)
end_ts   = pd.Timestamp(end) + pd.Timedelta(days=1)


# SAFE FILTERING (timestamp vs timestamp only)
items = items[(items["Date"]>=start_ts)&(items["Date"]<end_ts)]
art   = art[(art["Date"]>=start_ts)&(art["Date"]<end_ts)]


# =====================================================
# -------- PRESENCE CAPACITY
# =====================================================

pres["Start DT"]=pd.to_datetime(pres.get("Start DT"),errors="coerce",dayfirst=True)
pres["End DT"]=pd.to_datetime(pres.get("End DT"),errors="coerce",dayfirst=True)

pres=pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]

intervals=[
    x for x in(
        clip(s,e,start_ts,end_ts)
        for s,e in zip(pres["Start DT"],pres["End DT"])
    ) if x
]

available_sec=sum_seconds(intervals)


# =====================================================
# METRICS
# =====================================================

avg_aht = items["HandleSec"].mean()
avg_resp= art["ResponseSec"].mean()

util = items["HandleSec"].sum()/available_sec if available_sec else 0
emails_hr=len(items)/(available_sec/3600) if available_sec else 0

st.title("Email Department Performance")

c1,c2,c3,c4,c5=st.columns(5)

c1.metric("Total Emails",f"{len(items):,}")
c2.metric("Avg Handle Time (AHT)",fmt(avg_aht))
c3.metric("Avg Response Time",fmt(avg_resp))
c4.metric("Utilisation",f"{util:.1%}")
c5.metric("Emails / Available Hr",f"{emails_hr:.1f}")


# =====================================================
# DAILY TREND
# =====================================================

st.markdown("---")

daily_items=items.groupby(items["Date"].dt.date).size().reset_index(name="Volume")
daily_resp=art.groupby(art["Date"].dt.date)["ResponseSec"].mean().reset_index(name="Response")

daily=daily_items.merge(daily_resp,on="Date",how="left")

bars=alt.Chart(daily).mark_bar(opacity=0.15).encode(x="Date:T",y="Volume:Q")
line=alt.Chart(daily).mark_line(strokeWidth=2.5,point=True).encode(x="Date:T",y="Response:Q")

st.altair_chart(alt.layer(bars,line).resolve_scale(y="independent"),use_container_width=True)
