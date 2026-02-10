import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

st.set_page_config(layout="wide")
BASE = Path(__file__).parent

RESP_FILE  = "Responded PT.csv"
ITEMS_FILE = "ItemsPT.csv"      # optional
PRES_FILE  = "PresencePT.csv"   # optional

EMAIL_CHANNEL = "casesChannel"
AVAILABLE_STATUSES = {"Available_Email_and_Web", "Available_All"}


# =====================================================
# SAFE FILE LOADER (cannot crash)
# =====================================================

def load_csv(path):
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path, encoding="cp1252", low_memory=False)
    except:
        return pd.read_csv(path, encoding="utf-16", sep="\t", low_memory=False)


resp  = load_csv(BASE / RESP_FILE)
items = load_csv(BASE / ITEMS_FILE)
pres  = load_csv(BASE / PRES_FILE)

for df in (resp, items, pres):
    df.columns = df.columns.str.strip()


# =====================================================
# STOP EARLY IF NO RESPONDED FILE
# =====================================================

if resp.empty:
    st.error("Responded PT.csv not found. Upload or place in project folder.")
    st.stop()


# =====================================================
# =====================  ART  =========================
# =====================================================
# EXACT RULE (as you defined)
# ART = Email Message Date (agent reply) − Date/Time Opened
# =====================================================

resp["CaseID"] = resp["Case ID"].astype(str)

resp["OpenedDT"] = pd.to_datetime(
    resp["Date/Time Opened"],
    errors="coerce",
    dayfirst=True
)

resp["ReplyDT"] = pd.to_datetime(
    resp["Email Message Date"],
    errors="coerce",
    dayfirst=True
)

# exclude reopened / multi-touch
resp = resp.groupby("CaseID").filter(lambda x: len(x) == 1)

resp["ARTsec"] = (resp["ReplyDT"] - resp["OpenedDT"]).dt.total_seconds()

resp["Date"] = resp["ReplyDT"].dt.date


# =====================================================
# =====================  AHT  =========================
# =====================================================
# From ItemsPT only (true effort)
# =====================================================

avg_aht = 0
total_handle = 0

if not items.empty:

    items = items[items["Service Channel: Developer Name"] == EMAIL_CHANNEL].copy()

    # find handle column safely
    handle_col = next(
        (c for c in items.columns if "handle" in c.lower()),
        None
    )

    if handle_col:
        items["HandleSec"] = pd.to_numeric(items[handle_col], errors="coerce")

        avg_aht = items["HandleSec"].mean()
        total_handle = items["HandleSec"].sum()


# =====================================================
# =====================  PRESENCE  ====================
# =====================================================

available_sec = 0

if not pres.empty:

    pres["Start DT"] = pd.to_datetime(pres["Start DT"], errors="coerce", dayfirst=True)
    pres["End DT"]   = pd.to_datetime(pres["End DT"], errors="coerce", dayfirst=True)

    pres = pres[pres["Service Presence Status: Developer Name"].isin(AVAILABLE_STATUSES)]

    intervals = list(zip(pres["Start DT"], pres["End DT"]))

    for s, e in intervals:
        if pd.notna(s) and pd.notna(e):
            available_sec += (e - s).total_seconds()

util = (total_handle / available_sec) if available_sec else 0


# =====================================================
# FORMATTERS
# =====================================================

def fmt_mmss(sec):
    if not sec or pd.isna(sec):
        return "—"
    m, s = divmod(int(sec), 60)
    return f"{m:02}:{s:02}"

def fmt_hm(sec):
    if not sec or pd.isna(sec):
        return "—"
    h = int(sec)//3600
    m = (int(sec)%3600)//60
    return f"{h}h {m:02}m"


# =====================================================
# DATE RANGE
# =====================================================

min_d = resp["Date"].min()
max_d = resp["Date"].max()

start, end = st.date_input(
    "Date Range",
    value=(max_d - pd.Timedelta(days=6), max_d)
)

resp = resp[(resp["Date"] >= start) & (resp["Date"] <= end)]


# =====================================================
# METRICS
# =====================================================

avg_art = resp["ARTsec"].mean()

st.title("Email Performance")

c1, c2, c3, c4 = st.columns(4)

c1.metric("Replies", f"{len(resp):,}")
c2.metric("Avg Response Time (ART)", fmt_hm(avg_art))
c3.metric("Avg Handle Time (AHT)", fmt_mmss(avg_aht))
c4.metric("Utilisation", f"{util:.1%}")


# =====================================================
# DAILY VOLUME
# =====================================================

st.markdown("---")
st.subheader("Daily Replies")

daily = resp.groupby("Date").size().reset_index(name="Replies")

chart = alt.Chart(daily).mark_bar(
    color="#2563eb",
    opacity=0.6
).encode(
    x="Date:T",
    y="Replies:Q"
)

st.altair_chart(chart, use_container_width=True)
