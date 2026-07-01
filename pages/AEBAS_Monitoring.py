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
[data-testid="collapsedControl"]{display:flex!important;visibility:visible!important;
    opacity:1!important;position:fixed!important;top:50%!important;left:0!important;
    transform:translateY(-50%)!important;z-index:999999!important;
    background-color:#2f3343!important;border-radius:0 8px 8px 0!important;
    padding:12px 7px!important;box-shadow:3px 0 8px rgba(0,0,0,0.35)!important;
    cursor:pointer!important;}
[data-testid="collapsedControl"] button{background:transparent!important;
    border:none!important;padding:0!important;}
[data-testid="collapsedControl"] svg{fill:white!important;color:white!important;}
</style>""", unsafe_allow_html=True)

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown("""<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f3e0 Home")
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    for pat, lbl in [
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py","\U0001f91a AEBAS Monitoring"),
        ("pages/Bulk_Analytics.py|pages/*[Bb]ulk*.py","\U0001f4ca Bulk Customer Analytics"),
        ("pages/Delivery_Productivity.py|pages/*[Dd]elivery*.py","\U0001f4e6 Delivery Productivity"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py","\U0001f4bb Digital Transactions"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
        
        
    ]:
        hits = []
        for p in pat.split("|"): hits += _glob.glob(p)
        if hits:
            st.sidebar.page_link(hits[0].replace("\\", "/"), label=lbl)
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

# ── Auth guard ────────────────────────────────────────────────────────────────
if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Colour constants ──────────────────────────────────────────────────────────
TITLE_BG = "#1F3864"; TITLE_FG = "#FFFFFF"
HDR_BG   = "#2E75B6"; HDR_FG   = "#FFFFFF"
TOTAL_BG = "#FFF2CC"; TOTAL_FG = "#000000"
DIV_BG   = "#DDEBF7"; DIV_FG   = "#1F3864"
NOTMK_BG = "#FCE4D6"

# Canonical division names and display order
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
DIV_ORDER_MAP = {d: i for i, d in enumerate(DIV_ORDER)}

# Standardise raw division names from Consc sheet → canonical names
DIV_STD = {
    "SECUNDERABAD Dvn":          "Secunderabad Division",
    "Hyderabad South East":      "Hyderabad South East",
    "Hyderabad City Division":   "Hyderabad City Division",
    "Hyderabad Sorting Division":"Hyderabad Sorting Division",
    "Hyderabad GPO":             "Hyderabad GPO",
    "MMS, Hyderabad":            "MMS, Hyderabad",
    "PSD Hyderabad":             "PSD Hyderabad",
    "Sangareddy Division":       "Sangareddy Division",
    "Medak Division":            "Medak Division",
    "Regional Office, HQ Region":"Regional Office, HQ Region",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalise(x):
    """Normalise office name for fuzzy matching."""
    import re
    if pd.isna(x): return ""
    x = str(x).upper().strip()
    x = re.sub(r'[.,\-/]', ' ', x)
    x = re.sub(r'\s+', ' ', x).strip()
    # Strip common suffixes that differ between datasets
    for sfx in [" SUB DIVISION", " SUBDIVISION", " DIVISION", " DVN",
                " S O", " H O", " D O", " S.O", " H.O"]:
        if x.endswith(sfx):
            x = x[:-len(sfx)].strip()
    return x

@st.cache_data(show_spinner=False)
def load_master_excel(file_bytes):
    """
    Parse the AEBAS Master Excel.
    Sheet 'Consc'  → col0=office_name, col1=division  (335 offices, the master list)
    Returns DataFrame with columns: office_name, division, office_norm
    """
    xl = pd.ExcelFile(BytesIO(file_bytes))
    consc = xl.parse("Consc", header=None)
    consc.columns = ["office_name", "division"]
    consc = consc.dropna(subset=["office_name"]).copy()
    consc["office_name"] = consc["office_name"].astype(str).str.strip()
    consc["division"]    = consc["division"].astype(str).str.strip()
    # Standardise division names
    consc["division"] = consc["division"].map(
        lambda x: DIV_STD.get(x, x)
    )
    consc["office_norm"] = consc["office_name"].apply(normalise)
    return consc

@st.cache_data(show_spinner=False)
def load_aebas_export(file_bytes):
    """
    Parse the AEBAS portal export CSV.
    Key column: 'Office Location' — the offices that DID mark attendance.
    Returns list of unique normalised office names.
    """
    df = pd.read_csv(BytesIO(file_bytes))
    df.columns = [c.strip() for c in df.columns]
    offices = df["Office Location"].dropna().unique().tolist()
    return [normalise(o) for o in offices]

def match_offices(aebas_norms, master_norms, threshold):
    """
    For each AEBAS office (marked), find the closest master office.
    Returns set of master office_norm values that are marked.
    """
    master_list = list(master_norms)
    matched = set()
    for office in aebas_norms:
        result = process.extractOne(
            office, master_list, scorer=fuzz.partial_ratio
        )
        if result and result[1] >= threshold:
            matched.add(result[0])
    return matched

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(summary_df, not_marked_df, report_date):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    def f(**kw):
        base = {"border": 1, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)

    fmt_title   = f(bold=True, font_size=12, align="center",
                    bg_color=TITLE_BG, font_color=TITLE_FG)
    fmt_hdr     = f(bold=True, align="center", text_wrap=True,
                    bg_color=HDR_BG, font_color=HDR_FG)
    fmt_total   = f(bold=True, align="center",
                    bg_color=TOTAL_BG, font_color=TOTAL_FG)
    fmt_total_l = f(bold=True, align="left",
                    bg_color=TOTAL_BG, font_color=TOTAL_FG)
    fmt_green   = f(align="center", bg_color="#70AD47",
                    font_color="white", bold=True)
    fmt_lgreen  = f(align="center", bg_color="#E2EFDA", bold=True)
    fmt_amber   = f(align="center", bg_color="#FFC000", bold=True)
    fmt_red     = f(align="center", bg_color="#FF0000",
                    font_color="white", bold=True)
    fmt_c       = f(align="center")
    fmt_l       = f(align="left")
    fmt_div     = f(bold=True, align="center",
                    bg_color=DIV_BG, font_color=DIV_FG,
                    text_wrap=True)
    fmt_notmk   = f(align="left", bg_color=NOTMK_BG)
    fmt_note    = wb.add_format({"italic": True, "font_size": 9, "border": 0})

    date_str = report_date.strftime("%d.%m.%Y")

    # ── Sheet 1: Summary ──────────────────────────────────────────────────────
    ws1 = wb.add_worksheet("Summary")
    ws1.set_column(0, 0, 6)
    ws1.set_column(1, 1, 30)
    ws1.set_column(2, 4, 16)
    ws1.set_column(5, 5, 22)
    ws1.set_row(0, 24)
    ws1.set_row(1, 50)

    ws1.merge_range(0, 0, 0, 5,
        f"AEBAS Report dated {date_str}", fmt_title)
    for ci, h in enumerate([
        "Sl.\nNo.",
        "Name of the\nDivision/Unit",
        "Total no.\nof Units",
        "No. of offices\nmarked attendance\nin AEBAS",
        "No. of offices not\nmarked attendance\nin AEBAS",
        "% of offices\nimplemented\nAEBAS",
    ]):
        ws1.write(1, ci, h, fmt_hdr)

    data_rows = summary_df[summary_df["Division/Unit"] != "TOTAL HQ REGION"]
    ri = 2
    for _, row in data_rows.iterrows():
        v = float(row["% AEBAS"])
        pf = (fmt_green  if v >= 95 else
              fmt_lgreen if v >= 80 else
              fmt_amber  if v >= 60 else fmt_red)
        ws1.write(ri, 0, int(row["Sl."]),            fmt_c)
        ws1.write(ri, 1, row["Division/Unit"],        fmt_l)
        ws1.write(ri, 2, int(row["Total"]),           fmt_c)
        ws1.write(ri, 3, int(row["Marked"]),          fmt_c)
        ws1.write(ri, 4, int(row["Not Marked"]),      fmt_c)
        ws1.write(ri, 5, round(v, 2),                 pf)
        ri += 1

    # Total row
    tot = summary_df[summary_df["Division/Unit"] == "TOTAL HQ REGION"].iloc[0]
    ws1.write(ri, 0, "", fmt_total)
    ws1.write(ri, 1, "TOTAL", fmt_total_l)
    ws1.write(ri, 2, int(tot["Total"]),      fmt_total)
    ws1.write(ri, 3, int(tot["Marked"]),     fmt_total)
    ws1.write(ri, 4, int(tot["Not Marked"]), fmt_total)
    ws1.write(ri, 5, round(float(tot["% AEBAS"]), 2), fmt_total)
    ri += 2
    ws1.write(ri, 1,
        "* Departmental Post Offices only. Branch Offices (BPO) excluded.",
        fmt_note)

    # ── Sheet 2: Not Marked (merged Division column) ──────────────────────────
    ws2 = wb.add_worksheet("Not Marked")
    ws2.set_column(0, 0, 6)
    ws2.set_column(1, 1, 28)
    ws2.set_column(2, 2, 40)
    ws2.set_row(0, 24)
    ws2.set_row(1, 30)

    ws2.merge_range(0, 0, 0, 2,
        f"List of offices not marked attendance in AEBAS portal as on {date_str}",
        fmt_title)
    ws2.write(1, 0, "Sl. No.",              fmt_hdr)
    ws2.write(1, 1, "Name of the Division", fmt_hdr)
    ws2.write(1, 2, "Name of the Unit/Office", fmt_hdr)

    # Group by division preserving order
    not_marked_df = not_marked_df.copy()
    not_marked_df["_order"] = not_marked_df["division"].map(
        lambda x: DIV_ORDER_MAP.get(x, 99)
    )
    not_marked_df = not_marked_df.sort_values(
        ["_order", "division", "office_name"]
    ).reset_index(drop=True)

    ri2 = 2; sl = 1
    for div, grp in not_marked_df.groupby("division",
                                           sort=False,
                                           observed=True):
        offices = grp["office_name"].tolist()
        start = ri2
        for oname in offices:
            ws2.write(ri2, 0, sl, fmt_c)
            ws2.write(ri2, 2, oname, fmt_notmk)
            ri2 += 1; sl += 1
        # Merge division cell for this group
        if start == ri2 - 1:
            ws2.write(start, 1, div, fmt_div)
        else:
            ws2.merge_range(start, 1, ri2 - 1, 1, div, fmt_div)

    wb.close()
    return output.getvalue()

# ── Header ────────────────────────────────────────────────────────────────────
from PIL import Image
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None
h1, h2, _ = st.columns([1, 8, 1])
with h1:
    if logo: st.image(logo, width=80)
with h2:
    st.markdown(f"""
<h1 style='font-size:26px;margin-bottom:2px;color:#2f3343;font-weight:700;'>
AEBAS Monitoring Report</h1>
<p style='font-size:14px;color:#555;margin-top:0;'>
Headquarters Region – Telangana Postal Circle &nbsp;|&nbsp;
<b>Report Date: {date.today().strftime('%d.%m.%Y')}</b></p>
""", unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 12px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.title("🖐 AEBAS Monitoring")

today = date.today()
default_date = today - timedelta(days=3 if today.weekday() == 0 else 1)
report_date = st.sidebar.date_input("Report Date", value=default_date)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Upload Files")

master_file = st.sidebar.file_uploader(
    "AEBAS Master Excel",
    type=["xlsx"],
    help="Excel with 'Consc' sheet (office_name → division mapping)"
)
aebas_file = st.sidebar.file_uploader(
    "AEBAS Export CSV",
    type=["csv"],
    help="Downloaded from AEBAS portal — contains 'Office Location' column"
)

fuzzy_threshold = st.sidebar.slider(
    "Match sensitivity", 40, 90, 60,
    help="Higher = stricter matching. Lower = more lenient."
)

# ── Gate ──────────────────────────────────────────────────────────────────────
if not master_file or not aebas_file:
    st.info(
        "Upload both files from the sidebar to generate the report:\n\n"
        "1. **AEBAS Master Excel** — contains the `Consc` sheet with all "
        "335 departmental offices and their division mapping\n\n"
        "2. **AEBAS Export CSV** — downloaded from the AEBAS portal, "
        "contains the `Office Location` column of offices that marked attendance"
    )
    st.stop()

# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Loading master and AEBAS data..."):
    master_bytes = master_file.read()
    aebas_bytes  = aebas_file.read()
    master_df    = load_master_excel(master_bytes)
    aebas_norms  = load_aebas_export(aebas_bytes)

with st.spinner("Matching offices..."):
    marked_norms = match_offices(
        aebas_norms,
        master_df["office_norm"].tolist(),
        fuzzy_threshold
    )

master_df["Marked"] = master_df["office_norm"].isin(marked_norms)

# ── Summary ───────────────────────────────────────────────────────────────────
summary = (
    master_df.groupby("division")
    .agg(Total=("office_norm", "count"), Marked=("Marked", "sum"))
    .reset_index()
)
summary["Not Marked"] = summary["Total"] - summary["Marked"]
summary["% AEBAS"]    = (summary["Marked"] / summary["Total"] * 100).round(2)

# Sort by % descending
summary = summary.sort_values("% AEBAS", ascending=False).reset_index(drop=True)
summary.insert(0, "Sl.", range(1, len(summary) + 1))
summary.columns = ["Sl.", "Division/Unit", "Total", "Marked", "Not Marked", "% AEBAS"]

# Total row
tot_t = summary["Total"].sum()
tot_m = summary["Marked"].sum()
tot_row = pd.DataFrame([{
    "Sl.":          "",
    "Division/Unit":"TOTAL HQ REGION",
    "Total":         int(tot_t),
    "Marked":        int(tot_m),
    "Not Marked":    int(tot_t - tot_m),
    "% AEBAS":       round(tot_m / tot_t * 100, 2) if tot_t else 0,
}])
summary_display = pd.concat([summary, tot_row], ignore_index=True)

# ── Display Summary ───────────────────────────────────────────────────────────
st.subheader(f"AEBAS Report dated {report_date.strftime('%d.%m.%Y')}")

def _style_summary(row):
    if row["Division/Unit"] == "TOTAL HQ REGION":
        return [f"background-color:{TOTAL_BG};color:{TOTAL_FG};font-weight:700"] * len(row)
    try:
        v = float(row["% AEBAS"])
        pct_style = (
            "background-color:#70AD47;color:white;font-weight:700"  if v >= 95 else
            "background-color:#E2EFDA;font-weight:600"               if v >= 80 else
            "background-color:#FFC000;font-weight:700"               if v >= 60 else
            "background-color:#FF0000;color:white;font-weight:700"
        )
        div_style = (
            "background-color:#E2EFDA"  if v >= 80 else
            "background-color:#FFF2CC"  if v >= 60 else
            f"background-color:{NOTMK_BG}"
        )
        return ["", div_style, "", "", "", pct_style]
    except:
        return [""] * len(row)

styled = (
    summary_display.style
    .apply(_style_summary, axis=1)
    .format({"% AEBAS": "{:.2f}%"},
            subset=pd.IndexSlice[
                summary_display["Division/Unit"] != "TOTAL HQ REGION", ["% AEBAS"]
            ])
    .format({"% AEBAS": "{:.2f}%"},
            subset=pd.IndexSlice[
                summary_display["Division/Unit"] == "TOTAL HQ REGION", ["% AEBAS"]
            ])
)
st.dataframe(styled, use_container_width=True, hide_index=True)

# ── KPI metrics ───────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Offices",   int(tot_t))
c2.metric("Marked",          int(tot_m))
c3.metric("Not Marked",      int(tot_t - tot_m))
c4.metric("% Implementation",f"{round(tot_m/tot_t*100,2) if tot_t else 0}%")

# ── Not Marked table ──────────────────────────────────────────────────────────
not_marked = master_df[~master_df["Marked"]].copy()
not_marked["_order"] = not_marked["division"].map(
    lambda x: DIV_ORDER_MAP.get(x, 99)
)
not_marked = not_marked.sort_values(
    ["_order", "division", "office_name"]
).reset_index(drop=True)

st.markdown("---")
st.subheader(
    f"List of offices not marked attendance in AEBAS portal "
    f"as on {report_date.strftime('%d.%m.%Y')}"
)

# Build display with division shown only on first row of each group
rows = []; sl = 1; prev_div = None
for _, row in not_marked.iterrows():
    rows.append({
        "Sl. No.": sl,
        "Name of the Division":    row["division"] if row["division"] != prev_div else "",
        "Name of the Unit/Office": row["office_name"],
    })
    prev_div = row["division"]
    sl += 1

nm_display = pd.DataFrame(rows)

def _style_nm(row):
    div = str(row["Name of the Division"]).strip()
    if div and div != "nan":
        return [
            f"background-color:{DIV_BG};color:{DIV_FG};font-weight:700",
            f"background-color:{DIV_BG};color:{DIV_FG};font-weight:700",
            f"background-color:{DIV_BG};color:{DIV_FG};font-weight:700",
        ]
    return ["", "", f"background-color:{NOTMK_BG}"]

st.dataframe(
    nm_display.style.apply(_style_nm, axis=1),
    use_container_width=True,
    hide_index=True
)

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
excel_bytes = build_excel(summary_display, not_marked, report_date)
st.download_button(
    "⬇️ Download AEBAS Report (Excel)",
    data=excel_bytes,
    file_name=f"AEBAS_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ── Debug expander ────────────────────────────────────────────────────────────
with st.expander("🔍 Matching Debug", expanded=False):
    st.caption(f"Master offices: {len(master_df)} | "
               f"AEBAS portal offices (unique): {len(set(aebas_norms))} | "
               f"Matched: {len(marked_norms)}")
    st.markdown("**Master offices NOT matched (marked as 'Not Marked'):**")
    st.dataframe(
        not_marked[["office_name", "division"]].rename(
            columns={"office_name": "Office", "division": "Division"}
        ),
        use_container_width=True, hide_index=True
    )
