import streamlit as st
import pandas as pd
import numpy as np
import glob as _glob
import os
from datetime import date, timedelta
from io import BytesIO
import xlsxwriter
from rapidfuzz import process, fuzz

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AEBAS Monitoring",
    page_icon="🖐",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""<style>
[data-testid="stSidebarNav"]  { display:none!important; }
.block-container{padding-top:1.2rem!important;padding-bottom:0!important;}
header{visibility:hidden;height:0!important;}
[data-testid="collapsedControl"]{display:flex!important;visibility:visible!important;opacity:1!important;
    position:fixed!important;top:50%!important;left:0!important;transform:translateY(-50%)!important;
    z-index:999999!important;background-color:#2f3343!important;border-radius:0 8px 8px 0!important;
    padding:12px 7px!important;box-shadow:3px 0 8px rgba(0,0,0,0.35)!important;cursor:pointer!important;}
[data-testid="collapsedControl"] button{background:transparent!important;border:none!important;padding:0!important;}
[data-testid="collapsedControl"] svg{fill:white!important;color:white!important;}
</style>""", unsafe_allow_html=True)

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    for pat, lbl in [
        ("pages/Bulk_Analytics.py|pages/*[Bb]ulk*.py",            "\U0001f4ca Bulk Customer Analytics"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py", "\U0001f4bb Digital Transactions"),
        ("pages/Delivery_Productivity.py|pages/*[Dd]elivery*.py", "\U0001f4e6 Delivery Productivity"),
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py", "\U0001f91a AEBAS Monitoring"),
    ]:
        hits = []
        for p in pat.split("|"): hits += _glob.glob(p)
        if hits: st.sidebar.page_link(hits[0].replace("\\", "/"), label=lbl)
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

# ── Auth guard ────────────────────────────────────────────────────────────────
if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Colour palette ────────────────────────────────────────────────────────────
TITLE_BG  = "#1F3864"; TITLE_FG  = "#FFFFFF"
HDR_BG    = "#2E75B6"; HDR_FG    = "#FFFFFF"
SUB_BG    = "#9DC3E6"; SUB_FG    = "#000000"
TOTAL_BG  = "#FFF2CC"; TOTAL_FG  = "#000000"
GREEN_BG  = "#E2EFDA"; GREEN_FG  = "#000000"
AMBER_BG  = "#FFC000"; AMBER_FG  = "#000000"
RED_BG    = "#FF0000"; RED_FG    = "#FFFFFF"
NOTMK_BG  = "#FCE4D6"

# Division display order and clean names
DIV_ORDER = [
    "Hyderabad GPO",
    "Hyderabad City Division",
    "Hyderabad South East",
    "Secunderabad Division",
    "Medak Division",
    "Sangareddy Division",
    "Hyderabad Sorting Division",
    "MMS, Hyderabad",
    "PSD Hyderabad",
    "Regional Office, HQ Region",
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalise(x):
    if pd.isna(x): return ""
    import re
    x = str(x).upper().strip()
    x = re.sub(r'[.,\-/]', ' ', x)
    x = re.sub(r'\s+', ' ', x).strip()
    # Remove common suffixes that differ between datasets
    for suffix in [" SO", " S O", " HO", " H O", " DO", " SDO", " IDC",
                   " SUB DIVISION", " DIVISION", " DVN", " DN"]:
        if x.endswith(suffix):
            x = x[:-len(suffix)].strip()
            break
    return x

def fuzzy_match_offices(aebas_offices, master_offices, threshold=55):
    """
    Match AEBAS portal office names to master office names.
    Returns set of master office_norm values that are marked.
    """
    master_list = list(master_offices)
    matched = set()
    for office in aebas_offices:
        result = process.extractOne(
            office, master_list,
            scorer=fuzz.partial_ratio
        )
        if result and result[1] >= threshold:
            matched.add(result[0])
    return matched

def pct_colour(v):
    """Streamlit display colour for % value."""
    if v >= 95:  return "background-color:#70AD47;color:white;font-weight:700"
    if v >= 80:  return "background-color:#E2EFDA;font-weight:600"
    if v >= 60:  return "background-color:#FFC000;font-weight:700"
    return           "background-color:#FF0000;color:white;font-weight:700"

# ── Load master data ──────────────────────────────────────────────────────────
MASTER_PATHS = [
    "Master_AeBAS/AEBAS Master.xlsx",
    "data/AEBAS Master.xlsx",
    "AEBAS Master.xlsx",
]

@st.cache_data(show_spinner="Loading master data...")
def load_master(path):
    xl = pd.ExcelFile(path)
    # Consc sheet = consolidated office → division mapping (335 offices)
    consc = xl.parse("Consc", header=None)
    consc.columns = ["office_name", "division"]
    consc = consc.dropna(subset=["office_name"])
    consc["office_norm"] = consc["office_name"].apply(normalise)
    consc["division"]    = consc["division"].fillna("").astype(str).str.strip()
    return consc

master_path = next((p for p in MASTER_PATHS if os.path.exists(p)), None)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.title("🖐 AEBAS Monitoring")

today = date.today()
default_date = today - timedelta(days=3 if today.weekday() == 0 else 1)
report_date = st.sidebar.date_input("Report Date", value=default_date)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Upload Files")

office_master_file = st.sidebar.file_uploader(
    "Office Master (table-data CSV)",
    type=["csv"],
    help="Download from local portal — contains all offices with office-type-code"
)
aebas_file = st.sidebar.file_uploader(
    "AEBAS Export (export CSV)",
    type=["csv"],
    help="Downloaded from AEBAS portal — contains Office Location and Division/Unit"
)
master_override = st.sidebar.file_uploader(
    "AEBAS Master Excel (optional — overrides default)",
    type=["xlsx"],
    help="Excel with 'Consc' sheet mapping office names to divisions"
)

fuzzy_threshold = st.sidebar.slider(
    "Fuzzy match threshold", min_value=40, max_value=90, value=55,
    help="Lower = more lenient matching; raise if false positives appear"
)

# ── Header ────────────────────────────────────────────────────────────────────
from PIL import Image
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None
h1, h2, _ = st.columns([1, 8, 1])
with h1:
    if logo: st.image(logo, width=80)
with h2:
    st.markdown(f"""<h1 style='font-size:26px;margin-bottom:2px;color:#2f3343;font-weight:700;'>
AEBAS Monitoring Report</h1>
<p style='font-size:14px;color:#555;margin-top:0;'>
Headquarters Region – Telangana Postal Circle &nbsp;|&nbsp;
<b>Report Date: {report_date.strftime('%d.%m.%Y')}</b></p>""", unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 12px 0;border-color:#ddd;'>", unsafe_allow_html=True)

if not aebas_file:
    st.info(
        "Upload files from the sidebar to generate the report.\n\n"
        "**Required:** AEBAS Export CSV (from AEBAS portal)\n\n"
        "**Optional:** Office Master CSV (local portal) and AEBAS Master Excel"
    )
    st.stop()

# ── Load master mapping ───────────────────────────────────────────────────────
if master_override:
    try:
        xl_ov = pd.ExcelFile(master_override)
        consc = xl_ov.parse("Consc", header=None)
        consc.columns = ["office_name", "division"]
        consc = consc.dropna(subset=["office_name"])
        consc["office_norm"] = consc["office_name"].apply(normalise)
        consc["division"] = consc["division"].fillna("").astype(str).str.strip()
    except Exception as e:
        st.error(f"Could not read Master Excel: {e}"); st.stop()
elif master_path:
    consc = load_master(master_path)
else:
    st.error(
        "No AEBAS Master Excel found. Please upload it via the sidebar "
        "or place it at `Master_AeBAS/AEBAS Master.xlsx`."
    )
    st.stop()

# ── Build office list from Office Master CSV ──────────────────────────────────
if office_master_file:
    om = pd.read_csv(office_master_file)
    om.columns = [c.strip() for c in om.columns]
    # Filter: Hyderabad City Region only, exclude Branch Offices (BPO)
    om = om[om["region-office-name"].astype(str).str.strip() == "Hyderabad City Region"]
    om = om[om["office-type-code"].astype(str).str.strip() != "BPO"]
    om["office_norm"] = om["office-name"].apply(normalise)
    # Map division from Consc sheet; fall back to division-office-name
    div_map = dict(zip(consc["office_norm"], consc["division"]))
    om["division"] = om["office_norm"].map(div_map)
    om["division"] = om["division"].fillna(
        om["division-office-name"].astype(str).str.strip()
    )
    # Standardise division names
    DIV_STD = {
        "SECUNDERABAD Dvn":                "Secunderabad Division",
        "Secunderabad Division":           "Secunderabad Division",
        "Hyderabad South East":            "Hyderabad South East",
        "Hyderabad South East Division":   "Hyderabad South East",
        "Hyderabad City Division":         "Hyderabad City Division",
        "Hyderabad Sorting Division":      "Hyderabad Sorting Division",
        "Hyderabad GPO Division":          "Hyderabad GPO",
        "MMS, Hyderabad":                  "MMS, Hyderabad",
        "PSD Hyderabad":                   "PSD Hyderabad",
        "Sangareddy Division":             "Sangareddy Division",
        "Medak Division":                  "Medak Division",
        "Regional Office, HQ Region":     "Regional Office, HQ Region",
    }
    om["division"] = om["division"].map(lambda x: DIV_STD.get(str(x).strip(), str(x).strip()))
    master_offices = om
else:
    # Build from Consc sheet alone
    om_rows = []
    for _, row in consc.iterrows():
        om_rows.append({
            "office-name": row["office_name"],
            "office_norm": row["office_norm"],
            "division":    row["division"],
        })
    master_offices = pd.DataFrame(om_rows)

# ── Read AEBAS export ─────────────────────────────────────────────────────────
aebas = pd.read_csv(aebas_file)
aebas.columns = [c.strip() for c in aebas.columns]
aebas["office_norm"] = aebas["Office Location"].apply(normalise)
aebas_office_norms = aebas["office_norm"].dropna().unique().tolist()

# ── Fuzzy match ───────────────────────────────────────────────────────────────
with st.spinner("Matching AEBAS offices to master list..."):
    matched_norms = fuzzy_match_offices(
        aebas_office_norms,
        master_offices["office_norm"].tolist(),
        threshold=fuzzy_threshold
    )

master_offices["Marked"] = master_offices["office_norm"].isin(matched_norms)

# ── Summary table ─────────────────────────────────────────────────────────────
summary = (
    master_offices.groupby("division")
    .agg(
        Total=("office_norm", "count"),
        Marked=("Marked", "sum")
    )
    .reset_index()
)
summary["Not Marked"]   = summary["Total"] - summary["Marked"]
summary["% AEBAS"]      = (summary["Marked"] / summary["Total"] * 100).round(2)
summary.columns         = ["Division/Unit", "Total\nUnits", "Marked\nAttendance",
                            "Not\nMarked", "% Offices\nImplemented AEBAS"]

# Sort by % descending
summary = summary.sort_values("% Offices\nImplemented AEBAS", ascending=False).reset_index(drop=True)
summary.insert(0, "Sl. No.", range(1, len(summary) + 1))

# Total row
tot_total  = summary["Total\nUnits"].sum()
tot_marked = summary["Marked\nAttendance"].sum()
tot_row = pd.DataFrame([{
    "Sl. No.":                          "",
    "Division/Unit":                    "TOTAL HQ REGION",
    "Total\nUnits":                     int(tot_total),
    "Marked\nAttendance":               int(tot_marked),
    "Not\nMarked":                      int(tot_total - tot_marked),
    "% Offices\nImplemented AEBAS":     round(tot_marked / tot_total * 100, 2) if tot_total else 0,
}])
summary_display = pd.concat([summary, tot_row], ignore_index=True)

# ── Display Summary ───────────────────────────────────────────────────────────
st.subheader(f"AEBAS Report dated {report_date.strftime('%d.%m.%Y')}")

def style_summary(row):
    if row["Division/Unit"] == "TOTAL HQ REGION":
        return [f"background-color:{TOTAL_BG};color:{TOTAL_FG};font-weight:700"] * len(row)
    try:
        v = float(row["% Offices\nImplemented AEBAS"])
        if v >= 95:
            return ["",f"background-color:{GREEN_BG}",
                    "","","",f"background-color:#70AD47;color:white;font-weight:700"]
        elif v >= 80:
            return ["",f"background-color:{GREEN_BG}",
                    "","","",f"background-color:{GREEN_BG};font-weight:600"]
        elif v >= 60:
            return ["",f"background-color:{AMBER_BG}",
                    "","","",f"background-color:{AMBER_BG};font-weight:700"]
        else:
            return ["",f"background-color:{NOTMK_BG}",
                    "","","",f"background-color:{RED_BG};color:{RED_FG};font-weight:700"]
    except:
        return [""] * len(row)

styled = (
    summary_display.style
    .apply(style_summary, axis=1)
    .format({"% Offices\nImplemented AEBAS": "{:.2f}%"})
)
st.dataframe(styled, use_container_width=True, hide_index=True)

# ── Not Marked offices ────────────────────────────────────────────────────────
not_marked = master_offices[~master_offices["Marked"]].copy()

# Sort by division order then office name
div_order_map = {d: i for i, d in enumerate(DIV_ORDER)}
not_marked["_div_order"] = not_marked["division"].map(
    lambda x: div_order_map.get(x, 99)
)
not_marked = not_marked.sort_values(
    ["_div_order", "division", "office-name"]
).reset_index(drop=True)

# Build merged-division display (show division only on first row per group)
st.markdown("---")
st.subheader(
    f"List of offices not marked attendance in AEBAS portal as on "
    f"{report_date.strftime('%d.%m.%Y')}"
)

# Build display with merged Division column
display_rows = []
sl = 1
prev_div = None
for _, row in not_marked.iterrows():
    div_display = row["division"] if row["division"] != prev_div else ""
    display_rows.append({
        "Sl. No.":         sl,
        "Name of the\nDivision": div_display,
        "Name of the Unit/Office": row.get("office-name", row.get("office_norm", "")),
    })
    prev_div = row["division"]
    sl += 1

not_marked_display = pd.DataFrame(display_rows)

def style_notmarked(row):
    div = row["Name of the\nDivision"]
    if div and div.strip():
        return [
            "font-weight:700;background-color:#DDEBF7",
            "font-weight:700;background-color:#DDEBF7;color:#1F3864",
            "font-weight:700;background-color:#DDEBF7",
        ]
    return ["", "", "background-color:#FCE4D6"]

styled_nm = not_marked_display.style.apply(style_notmarked, axis=1)
st.dataframe(styled_nm, use_container_width=True, hide_index=True)

# ── KPI metrics ───────────────────────────────────────────────────────────────
total_offices = len(master_offices)
total_marked  = int(master_offices["Marked"].sum())
total_pending = total_offices - total_marked
overall_pct   = round(total_marked / total_offices * 100, 2) if total_offices else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Departmental Offices", total_offices)
col2.metric("Marked Attendance", total_marked)
col3.metric("Not Marked", total_pending)
col4.metric("% Implementation", f"{overall_pct}%")

# ── Excel Export ──────────────────────────────────────────────────────────────
def build_excel(summary_df, not_marked_df, report_date):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    def f(**kw):
        base = {"border": 1, "valign": "vcenter"}; base.update(kw)
        return wb.add_format(base)

    title_fmt  = f(bold=True, font_size=12, align="center",
                   bg_color=TITLE_BG, font_color=TITLE_FG)
    hdr_fmt    = f(bold=True, align="center", text_wrap=True,
                   bg_color=HDR_BG, font_color=HDR_FG)
    total_fmt  = f(bold=True, align="center",
                   bg_color=TOTAL_BG, font_color=TOTAL_FG)
    total_l    = f(bold=True, align="left",
                   bg_color=TOTAL_BG, font_color=TOTAL_FG)
    green_fmt  = f(align="center", bg_color="#70AD47", font_color="white", bold=True)
    green_l_fmt= f(align="center", bg_color=GREEN_BG)
    amber_fmt  = f(align="center", bg_color=AMBER_BG, font_color=AMBER_FG, bold=True)
    red_fmt    = f(align="center", bg_color=RED_BG, font_color=RED_FG, bold=True)
    plain_c    = f(align="center")
    plain_l    = f(align="left")
    div_hdr_f  = f(bold=True, align="left",
                   bg_color="#DDEBF7", font_color="#1F3864")
    notmk_f    = f(align="left", bg_color=NOTMK_BG)

    date_str = report_date.strftime("%d.%m.%Y")

    # ── Sheet 1: Summary ───────────────────────────────────────────────────────
    ws1 = wb.add_worksheet("Summary")
    ws1.set_column(0, 0, 6)   # Sl No
    ws1.set_column(1, 1, 30)  # Division
    ws1.set_column(2, 4, 14)  # Counts
    ws1.set_column(5, 5, 20)  # %
    ws1.set_row(0, 22); ws1.set_row(1, 36)

    ws1.merge_range(0, 0, 0, 5,
        f"AEBAS Report dated {date_str}", title_fmt)
    hdrs = ["Sl. No.", "Name of the\nDivision/Unit",
            "Total no.\nof Units*",
            "No. of offices\nmarked attendance\nin AEBAS",
            "No. of offices\nnot marked\nattendance\nin AEBAS",
            "% of offices\nimplemented\nAEBAS"]
    for ci, h in enumerate(hdrs): ws1.write(1, ci, h, hdr_fmt)

    # Data rows (exclude total row from summary_df passed in)
    data_rows = summary_df[summary_df["Division/Unit"] != "TOTAL HQ REGION"]
    ri = 2
    for _, row in data_rows.iterrows():
        v = float(row["% Offices\nImplemented AEBAS"])
        pf = green_fmt if v >= 95 else (green_l_fmt if v >= 80 else (amber_fmt if v >= 60 else red_fmt))
        ws1.write(ri, 0, row["Sl. No."],             plain_c)
        ws1.write(ri, 1, row["Division/Unit"],        plain_l)
        ws1.write(ri, 2, int(row["Total\nUnits"]),   plain_c)
        ws1.write(ri, 3, int(row["Marked\nAttendance"]), plain_c)
        ws1.write(ri, 4, int(row["Not\nMarked"]),    plain_c)
        ws1.write(ri, 5, round(v, 2),                pf)
        ri += 1

    # Total row
    tot = summary_df[summary_df["Division/Unit"] == "TOTAL HQ REGION"].iloc[0]
    ws1.write(ri, 0, "",                             total_fmt)
    ws1.write(ri, 1, "TOTAL",                        total_l)
    ws1.write(ri, 2, int(tot["Total\nUnits"]),       total_fmt)
    ws1.write(ri, 3, int(tot["Marked\nAttendance"]),total_fmt)
    ws1.write(ri, 4, int(tot["Not\nMarked"]),        total_fmt)
    ws1.write(ri, 5, round(float(tot["% Offices\nImplemented AEBAS"]), 2), total_fmt)

    # Footnote
    ws1.write(ri + 2, 1,
        "* Excludes Branch Offices (BPO). Departmental Post Offices only.",
        wb.add_format({"italic": True, "font_size": 9}))

    # ── Sheet 2: Not Marked ────────────────────────────────────────────────────
    ws2 = wb.add_worksheet("Not Marked")
    ws2.set_column(0, 0, 6)
    ws2.set_column(1, 1, 28)
    ws2.set_column(2, 2, 38)
    ws2.set_row(0, 22); ws2.set_row(1, 30)

    ws2.merge_range(0, 0, 0, 2,
        f"List of offices not marked attendance in AEBAS portal as on {date_str}",
        title_fmt)
    ws2.write(1, 0, "Sl. No.", hdr_fmt)
    ws2.write(1, 1, "Name of the Division", hdr_fmt)
    ws2.write(1, 2, "Name of the Unit/Office", hdr_fmt)

    ri2 = 2
    prev_div = None
    first_row_of_div = {}
    div_groups = {}

    for _, row in not_marked_df.iterrows():
        div = row["division"]
        oname = row.get("office-name", row.get("office_norm", ""))
        if div not in div_groups: div_groups[div] = []
        div_groups[div].append(oname)

    sl = 1
    for div in sorted(div_groups.keys(),
                      key=lambda x: div_order_map.get(x, 99)):
        offices = div_groups[div]
        start_row = ri2
        for i, oname in enumerate(offices):
            ws2.write(ri2, 0, sl, plain_c)
            ws2.write(ri2, 2, oname, notmk_f)
            ri2 += 1; sl += 1
        end_row = ri2 - 1
        # Merge division column for this group
        if start_row == end_row:
            ws2.write(start_row, 1, div, div_hdr_f)
        else:
            ws2.merge_range(start_row, 1, end_row, 1, div, div_hdr_f)

    wb.close()
    return output.getvalue()

excel_bytes = build_excel(summary_display, not_marked, report_date)
st.download_button(
    "⬇️ Download AEBAS Report (Excel)",
    data=excel_bytes,
    file_name=f"AEBAS_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ── Debug expander ────────────────────────────────────────────────────────────
with st.expander("🔍 Debug: Matching Details", expanded=False):
    st.markdown("**AEBAS portal offices not matched to any master office:**")
    unmatched_aebas = [o for o in aebas_office_norms
                       if not any(process.extractOne(o, [m], scorer=fuzz.partial_ratio)[1] >= fuzzy_threshold
                                  for m in master_offices["office_norm"].tolist()[:10])]
    st.write(f"Total unmatched: {len(unmatched_aebas)}")
    if unmatched_aebas[:20]:
        st.write(unmatched_aebas[:20])

    st.markdown("**Master offices with missing/unknown division:**")
    unknown_div = master_offices[
        master_offices["division"].isna() |
        (master_offices["division"].str.strip() == "") |
        (master_offices["division"].str.lower().str.strip() == "undefined")
    ][["office-name" if "office-name" in master_offices.columns else "office_norm", "division"]]
    st.dataframe(unknown_div, use_container_width=True, hide_index=True)
