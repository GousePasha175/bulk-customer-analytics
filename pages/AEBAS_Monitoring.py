import streamlit as st
import pandas as pd
from datetime import date, timedelta
from io import BytesIO

st.set_page_config(page_title="AEBAS Monitoring", layout="wide")


# ============================================================
# Sidebar Nav (same style as other pages)
# ============================================================
def _render_nav():
    st.page_link("Analytics_Excel.py", label="🏠 Home")
    st.page_link("pages/Bulk_Customer_Analytics.py", label="📊 Bulk Customer Analytics")
    st.page_link("pages/POSB Daily Report.py", label="📮 POSB Daily Report")
    st.page_link("pages/Digital Transactions.py", label="💻 Digital Transactions")
    st.page_link("pages/AEBAS_Monitoring.py", label="🖐 AEBAS Monitoring")


with st.sidebar:
    _render_nav()
    st.markdown("---")
    st.title("🖐 AEBAS Monitoring")

    today = date.today()
    if today.weekday() == 0:
        default_date = today - timedelta(days=3)
    else:
        default_date = today - timedelta(days=1)

    report_date = st.date_input("Report Date", value=default_date)

    table_file = st.file_uploader("Upload Table Data CSV", type=["csv"])
    export_file = st.file_uploader("Upload AEBAS Export CSV", type=["csv"])
    master_file = st.file_uploader("Upload Master Excel", type=["xlsx"])


# ============================================================
# Helper Functions
# ============================================================
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).upper().strip()
    text = text.replace(".", "")
    text = text.replace(",", "")
    text = " ".join(text.split())
    return text


def build_excel(summary_df, pending_df, report_date):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        pending_df.to_excel(writer, sheet_name="Not Marked", index=False)

    return output.getvalue()


# ============================================================
# Main Screen
# ============================================================
st.title(f"AEBAS Monitoring Report as on {report_date.strftime('%d.%m.%Y')}")

if not all([table_file, export_file, master_file]):
    st.info("Please upload all 3 files from sidebar.")
    st.stop()


# ============================================================
# Read Files
# ============================================================
table_df = pd.read_csv(table_file)
export_df = pd.read_csv(export_file)
master_df = pd.read_excel(master_file, sheet_name="Consc")


# ============================================================
# Adjust column names if needed
# ============================================================
table_df.columns = [c.strip() for c in table_df.columns]
export_df.columns = [c.strip() for c in export_df.columns]
master_df.columns = [c.strip() for c in master_df.columns]

# CHANGE THESE IF COLUMN NAMES DIFFER
TABLE_OFFICE_COL = "office-name"
EXPORT_OFFICE_COL = "Office Location"
MASTER_OFFICE_COL = master_df.columns[0]
MASTER_DIV_COL = master_df.columns[1]


# ============================================================
# Exclude Branch Offices
# ============================================================
table_df = table_df[
    ~table_df[TABLE_OFFICE_COL].astype(str).str.endswith("B.O")
]


# ============================================================
# Normalize
# ============================================================
table_df["office_norm"] = table_df[TABLE_OFFICE_COL].apply(normalize)
export_df["office_norm"] = export_df[EXPORT_OFFICE_COL].apply(normalize)

master_df["office_norm"] = master_df[MASTER_OFFICE_COL].apply(normalize)
master_df["Division"] = master_df[MASTER_DIV_COL]


# ============================================================
# Only HQ Region offices
# ============================================================
region_offices = set(master_df["office_norm"])

table_df = table_df[table_df["office_norm"].isin(region_offices)]
export_df = export_df[export_df["office_norm"].isin(region_offices)]


# ============================================================
# Attendance Matching
# ============================================================
marked_offices = set(export_df["office_norm"].unique())

master_df["Marked"] = master_df["office_norm"].isin(marked_offices)


# ============================================================
# Summary
# ============================================================
summary_df = (
    master_df.groupby("Division")
    .agg(
        Total_Units=("office_norm", "count"),
        Marked=("Marked", "sum")
    )
    .reset_index()
)

summary_df["Not Marked"] = summary_df["Total_Units"] - summary_df["Marked"]
summary_df["% Implemented AEBAS"] = (
    summary_df["Marked"] * 100 / summary_df["Total_Units"]
).round(2)

total_row = pd.DataFrame([{
    "Division": "TOTAL HQ REGION",
    "Total_Units": summary_df["Total_Units"].sum(),
    "Marked": summary_df["Marked"].sum(),
    "Not Marked": summary_df["Not Marked"].sum(),
    "% Implemented AEBAS": round(
        summary_df["Marked"].sum() * 100 / summary_df["Total_Units"].sum(),
        2
    )
}])

summary_df = pd.concat([summary_df, total_row], ignore_index=True)


# ============================================================
# Pending Offices
# ============================================================
pending_df = master_df[~master_df["Marked"]].copy()
pending_df = pending_df[["Division", MASTER_OFFICE_COL]]
pending_df.columns = ["Division", "Office Name"]
pending_df.insert(0, "Sl No", range(1, len(pending_df) + 1))


# ============================================================
# Display Summary
# ============================================================
st.subheader(f"AEBAS Report dated {report_date.strftime('%d.%m.%Y')}")

st.dataframe(
    summary_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# Display Pending
# ============================================================
st.subheader(
    f"List of offices not marked attendance in AEBAS portal as on {report_date.strftime('%d.%m.%Y')}"
)

st.dataframe(
    pending_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# Download
# ============================================================
excel_bytes = build_excel(summary_df, pending_df, report_date)

st.download_button(
    "⬇ Download AEBAS Report",
    data=excel_bytes,
    file_name=f"AEBAS_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
