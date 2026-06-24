import streamlit as st
import pandas as pd
import glob as _glob
from datetime import date, timedelta
from difflib import get_close_matches
from io import BytesIO

st.set_page_config(page_title="AEBAS Monitoring",layout="wide",initial_sidebar_state="expanded")


# ================= NAVIGATION =================
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
        st.sidebar.page_link(
            _bulk[0].replace("\\", "/"),
            label="📊 Bulk Customer Analytics"
        )

    _posb = (_glob.glob("pages/POSB Daily Report.py") +
             _glob.glob("pages/*[Pp][Oo][Ss][Bb]*.py"))
    if _posb:
        st.sidebar.page_link(
            _posb[0].replace("\\", "/"),
            label="📮 POSB Daily Report"
        )

    _aebas = (_glob.glob("pages/AEBAS_Monitoring.py") +
              _glob.glob("pages/*[Aa][Ee][Bb][Aa][Ss]*.py"))
    if _aebas:
        st.sidebar.page_link(
            _aebas[0].replace("\\", "/"),
            label="🖐 AEBAS Monitoring"
        )

    _dig = (_glob.glob("pages/1_Digital_Transactions.py") +
            _glob.glob("pages/*[Dd]igital*.py"))
    if _dig:
        st.sidebar.page_link(
            _dig[0].replace("\\", "/"),
            label="💻 Digital Transactions"
        )

    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)


# ================= HELPERS =================
def normalize(x):
    if pd.isna(x):
        return ""
    x = str(x).upper().strip()
    x = x.replace(".", "")
    x = x.replace(",", "")
    x = " ".join(x.split())
    return x


def fuzzy_match(name, candidates):
    match = get_close_matches(name, candidates, n=1, cutoff=0.50)
    return match[0] if match else None


def export_excel(summary_df, pending_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        pending_df.to_excel(writer, sheet_name="Not Marked", index=False)

    return output.getvalue()


# ================= SIDEBAR =================
with st.sidebar:
    _render_nav()
    st.title("🖐 AEBAS Monitoring")

    today = date.today()
    if today.weekday() == 0:  # Monday
        default_date = today - timedelta(days=3)
    else:
        default_date = today - timedelta(days=1)

    report_date = st.date_input(
        "Report Date",
        value=default_date
    )

    aebas_file = st.file_uploader(
        "Upload AEBAS CSV",
        type=["csv"]
    )


# ================= MAIN =================
st.title(f"AEBAS Monitoring Report as on {report_date.strftime('%d.%m.%Y')}")

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
    st.error(f"Unable to read Master File: {e}")
    st.stop()

if not aebas_file:
    st.info("Please upload AEBAS CSV from sidebar.")
    st.stop()

# ================= PREP MASTER =================
consolidated.columns = ["office_name", "division"]

office_master.columns = [c.strip() for c in office_master.columns]

# Hyderabad Region only
office_master = office_master[
    office_master["region-office-name"].astype(str).str.strip() == "Hyderabad City Region"
]

# Exclude Branch Offices
office_master = office_master[
    office_master["office-type-code"].astype(str).str.strip() != "BPO"
]

# Normalize
consolidated["office_norm"] = consolidated["office_name"].apply(normalize)
office_master["office_norm"] = office_master["office-name"].apply(normalize)

# Division mapping
division_map = dict(
    zip(consolidated["office_norm"], consolidated["division"])
)

office_master["division"] = office_master["office_norm"].map(division_map)

# If division missing, use Office Master division
office_master["division"] = office_master["division"].fillna(
    office_master["division-office-name"]
)

# ================= READ AEBAS =================
aebas = pd.read_csv(aebas_file)
aebas.columns = [c.strip() for c in aebas.columns]

aebas["office_norm"] = aebas["Office Location"].apply(normalize)

master_names = office_master["office_norm"].tolist()

matched = set()

for office in aebas["office_norm"].unique():
    m = fuzzy_match(office, master_names)
    if m:
        matched.add(m)

office_master["Marked"] = office_master["office_norm"].isin(matched)

# ================= SUMMARY =================
summary_df = (
    office_master.groupby("division")
    .agg(
        total=("office_norm", "count"),
        marked=("Marked", "sum")
    )
    .reset_index()
)

summary_df["not_marked"] = summary_df["total"] - summary_df["marked"]
summary_df["percent"] = (
    summary_df["marked"] * 100 / summary_df["total"]
).round(2)

summary_df.columns = [
    "Division",
    "Total Units",
    "Marked Attendance",
    "Not Marked",
    "% Implemented AEBAS"
]

# Add total row
total_row = pd.DataFrame([{
    "Division": "TOTAL HQ REGION",
    "Total Units": summary_df["Total Units"].sum(),
    "Marked Attendance": summary_df["Marked Attendance"].sum(),
    "Not Marked": summary_df["Not Marked"].sum(),
    "% Implemented AEBAS": round(
        summary_df["Marked Attendance"].sum() * 100 /
        summary_df["Total Units"].sum(), 2
    )
}])

summary_df = pd.concat([summary_df, total_row], ignore_index=True)

# ================= PENDING =================
pending_df = office_master[~office_master["Marked"]].copy()

pending_df = pending_df[["division", "office-name"]]
pending_df.columns = ["Division", "Office"]
pending_df.insert(0, "Sl No", range(1, len(pending_df) + 1))

# ================= DISPLAY =================
st.subheader(f"AEBAS Report dated {report_date.strftime('%d.%m.%Y')}")
def highlight_rows(row):
    if row["Division"] == "TOTAL HQ REGION":
        return ['background-color: #FFF2CC; font-weight: bold'] * len(row)
    elif row["% Implemented AEBAS"] < 80:
        return ['background-color: #FDEDEC'] * len(row)
    else:
        return ['background-color: #E8F8F5'] * len(row)

styled_summary = (
    summary_df.style
    .apply(highlight_rows, axis=1)
    .set_properties(**{
        'text-align': 'center'
    })
)

st.dataframe(
    styled_summary,
    use_container_width=True,
    hide_index=True
)

st.subheader(
    f"List of offices not marked attendance in AEBAS portal as on {report_date.strftime('%d.%m.%Y')}"
)
st.dataframe(
    pending_df,
    use_container_width=True,
    hide_index=True
)

# ================= DOWNLOAD =================
excel_bytes = export_excel(summary_df, pending_df)

st.download_button(
    label="⬇️ Download AEBAS Report",
    data=excel_bytes,
    file_name=f"AEBAS_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
