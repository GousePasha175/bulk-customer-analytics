import streamlit as st
import pandas as pd
from datetime import date, timedelta
from difflib import get_close_matches
import glob as _glob

st.set_page_config(page_title="AEBAS Monitoring", layout="wide")


# Navigation
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""",
        unsafe_allow_html=True)

    st.sidebar.page_link("Analytics_Excel.py", label="🔐 Login")

    _bulk = (_glob.glob("pages/Bulk Analytics.py") +
             _glob.glob("pages/*[Bb]ulk*.py"))
    if _bulk:
        st.sidebar.page_link(_bulk[0].replace("\\", "/"),
                             label="📊 Bulk Customer Analytics")

    _posb = (_glob.glob("pages/POSB Daily Report.py") +
             _glob.glob("pages/*[Pp][Oo][Ss][Bb]*.py"))
    if _posb:
        st.sidebar.page_link(_posb[0].replace("\\", "/"),
                             label="📮 POSB Daily Report")

    _aebas = (_glob.glob("pages/AEBAS Monitoring.py") +
              _glob.glob("pages/*[Aa][Ee][Bb][Aa][Ss]*.py"))
    if _aebas:
        st.sidebar.page_link(_aebas[0].replace("\\", "/"),
                             label="🖐 AEBAS Monitoring")


def normalize(x):
    if pd.isna(x):
        return ""
    x = str(x).upper()
    x = x.replace(".", "")
    x = x.replace(",", "")
    x = " ".join(x.split())
    return x.strip()


def fuzzy_match(name, candidates):
    matches = get_close_matches(name, candidates, n=1, cutoff=0.85)
    return matches[0] if matches else None


# Sidebar
with st.sidebar:
    _render_nav()
    st.markdown("---")
    st.title("🖐 AEBAS Monitoring")

    today = date.today()
    if today.weekday() == 0:
        default_date = today - timedelta(days=3)
    else:
        default_date = today - timedelta(days=1)

    report_date = st.date_input("Report Date", default_date)

    aebas_file = st.file_uploader("Upload AEBAS CSV", type=["csv"])


st.title(f"AEBAS Monitoring Report as on {report_date.strftime('%d.%m.%Y')}")

if not aebas_file:
    st.info("Upload AEBAS CSV from sidebar.")
    st.stop()


# Load master automatically
MASTER_FILE = "Master_AeBAS/AEBAS Master.xlsx"

try:
    consolidated = pd.read_excel(
        MASTER_FILE,
        sheet_name="Consolidated",
        header=None
    )
    office_master = pd.read_excel(
        MASTER_FILE,
        sheet_name="Office Master"
    )
except Exception as e:
    st.error(f"Master file error: {e}")
    st.stop()


consolidated.columns = ["office_name", "division"]


# Filter Hyderabad Region only
office_master = office_master[
    office_master["region-office-name"] == "Hyderabad City Region"
]

# Exempt BPO
office_master = office_master[
    office_master["office-type-code"] != "BPO"
]


# Read AEBAS CSV
aebas = pd.read_csv(aebas_file)
aebas.columns = [c.strip() for c in aebas.columns]

# CHANGE IF COLUMN NAME DIFFERS
AEBAS_OFFICE_COL = "Office Location"

office_master["office_norm"] = office_master["office-name"].apply(normalize)
consolidated["office_norm"] = consolidated["office_name"].apply(normalize)
aebas["office_norm"] = aebas[AEBAS_OFFICE_COL].apply(normalize)

master_names = consolidated["office_norm"].tolist()

matched = set()
for office in aebas["office_norm"].unique():
    m = fuzzy_match(office, master_names)
    if m:
        matched.add(m)

consolidated["Marked"] = consolidated["office_norm"].isin(matched)

summary = (
    consolidated.groupby("division")
    .agg(
        total=("office_name", "count"),
        marked=("Marked", "sum")
    )
    .reset_index()
)

summary["not_marked"] = summary["total"] - summary["marked"]
summary["percent"] = (
    summary["marked"] * 100 / summary["total"]
).round(2)

summary.columns = [
    "Division",
    "Total Units",
    "Marked Attendance",
    "Not Marked",
    "% Implemented AEBAS"
]

pending = consolidated[~consolidated["Marked"]].copy()
pending = pending[["division", "office_name"]]
pending.columns = ["Division", "Office"]
pending.insert(0, "Sl No", range(1, len(pending) + 1))

st.subheader("Division-wise Summary")
st.dataframe(summary, use_container_width=True, hide_index=True)

st.subheader(
    f"List of offices not marked attendance in AEBAS portal as on {report_date.strftime('%d.%m.%Y')}"
)
st.dataframe(pending, use_container_width=True, hide_index=True)
