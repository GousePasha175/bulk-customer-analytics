import streamlit as st
import pandas as pd
import numpy as np
import glob as _glob
import os
import re
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
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    for pat, lbl in [
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py","\U0001f91a AEBAS Monitoring"),
        ("pages/Bulk_Analytics.py|pages/*[Bb]ulk*.py","\U0001f4ca Bulk Customer Analytics"),
        ("pages/Delivery_Productivity.py|pages/*[Dd]elivery*.py","\U0001f4e6 Delivery Productivity"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py","\U0001f4bb Digital Transactions"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
        ("pages/Sorting_Assistance.py|pages/*[Ss]orting*.py", "📮 Sorting Assistance"),
        
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

# Standardise raw division names from the master sheets → canonical names
DIV_STD = {
    "SECUNDERABAD Dvn":           "Secunderabad Division",
    "Hyderabad South East":       "Hyderabad South East",
    "Hyderabad City Division":    "Hyderabad City Division",
    "Hyderabad Sorting Division": "Hyderabad Sorting Division",
    "Hyderabad GPO":              "Hyderabad GPO",
    "Hyderabad GPO Division":     "Hyderabad GPO",
    "MMS, Hyderabad":             "MMS, Hyderabad",
    "PSD Hyderabad":              "PSD Hyderabad",
    "Sangareddy Division":        "Sangareddy Division",
    "Medak Division":             "Medak Division",
    "Regional Office, HQ Region": "Regional Office, HQ Region",
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Helpers ───────────────────────────────────────────────────────────────────
def normalise(x):
    """Normalise office name for fuzzy matching."""
    if pd.isna(x): return ""
    x = str(x).upper().strip()
    x = re.sub(r'[.,\-/]', ' ', x)
    x = re.sub(r'\s+', ' ', x).strip()
    for sfx in [" SUB DIVISION", " SUBDIVISION", " DIVISION", " DVN",
                " S O", " H O", " D O", " S.O", " H.O"]:
        if x.endswith(sfx):
            x = x[:-len(sfx)].strip()
    return x

def find_bundled_master():
    """
    Look for the bundled AEBAS master workbook anywhere near this script.
    Matches on filename containing both 'aebas' and 'master' (case-insensitive,
    regardless of underscore/space), so 'AEBAS_Master.xlsx', 'AEBAS Master.xlsx',
    'Master_AeBAS/AEBAS Master.xlsx', etc. are all found.
    """
    skip_dirs = {".git", "node_modules", "__pycache__", ".streamlit", "venv", ".venv"}
    search_roots = []
    for r in [BASE_DIR, os.path.dirname(BASE_DIR), os.getcwd()]:
        rn = os.path.normpath(r)
        if rn not in search_roots:
            search_roots.append(rn)

    for root in search_roots:
        if not os.path.isdir(root):
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in skip_dirs]
            for fn in filenames:
                low = fn.lower()
                if low.endswith(".xlsx") and "aebas" in low and "master" in low:
                    return os.path.join(dirpath, fn)
    return None

def _find_col(columns, *keywords):
    lc = {str(c).lower().strip(): c for c in columns}
    for kw in keywords:
        for l, orig in lc.items():
            if kw in l:
                return orig
    return None

def _find_division_col(columns):
    for c in columns:
        l = str(c).lower().strip()
        if "division" in l or "region" in l:
            return c
    return None

def _find_id_col(columns):
    for c in columns:
        l = str(c).lower().strip()
        if "office" in l and "id" in l:
            return c
    for c in columns:
        l = str(c).lower().strip()
        if l == "id" or l.endswith(" id") or "code" in l:
            return c
    return None

def _find_type_col(columns):
    for c in columns:
        if "type" in str(c).lower():
            return c
    return None

def _find_office_name_col(columns, exclude=()):
    cols = [c for c in columns if c not in exclude]
    for c in cols:
        l = str(c).lower().strip()
        if ("office" in l or "unit" in l) and "name" in l:
            return c
    for c in cols:
        l = str(c).lower().strip()
        if ("office" in l or "unit" in l) and "id" not in l and "type" not in l \
                and "division" not in l and "region" not in l:
            return c
    for c in cols:
        l = str(c).lower().strip()
        if "name" in l and "division" not in l and "region" not in l:
            return c
    return None

def _guess_name_col_by_content(df, used_cols):
    """Last resort: pick the unused column that looks most like free-text office names
    (mostly alphabetic, mostly unique values)."""
    candidates = [c for c in df.columns if c not in used_cols]
    best, best_score = None, -1.0
    for c in candidates:
        vals = df[c].dropna().astype(str)
        if len(vals) == 0:
            continue
        alpha_frac = vals.str.contains(r"[A-Za-z]").mean()
        uniqueness = vals.nunique() / max(len(vals), 1)
        score = alpha_frac * uniqueness
        if score > best_score:
            best_score, best = score, c
    return best

def _read_sheet(xl, preferred_name):
    if preferred_name in xl.sheet_names:
        return xl.parse(preferred_name)
    return xl.parse(xl.sheet_names[0])

def _looks_like_bo(office_name, office_type):
    """Branch Offices (BOs/BPOs) are excluded from AEBAS monitoring entirely —
    they run on GDS staff without biometric devices."""
    t = str(office_type).strip().upper()
    if t and t != "NAN":
        return t == "BPO"
    name = str(office_name).upper().replace(",", " ").strip()
    return bool(re.search(r'\bB\.?\s*O\.?$', name))

@st.cache_data(show_spinner=False)
def load_office_master(file_bytes):
    """
    Parse the 'Office Master' sheet (a.k.a. 'AEBAS Master'): the full office
    list including Branch Offices, used only as a name-matching dictionary.
    Returns columns: office_name, office_id, division, office_type, office_norm, is_bo
    """
    xl = pd.ExcelFile(BytesIO(file_bytes))
    df = _read_sheet(xl, "Office Master")
    col_div  = _find_division_col(df.columns)
    col_id   = _find_id_col(df.columns)
    col_type = _find_type_col(df.columns)
    col_name = _find_office_name_col(df.columns, exclude={c for c in [col_div, col_id, col_type] if c})
    if col_name is None:
        col_name = _guess_name_col_by_content(df, {c for c in [col_div, col_id, col_type] if c})
    if col_name is None:
        raise ValueError(
            "Office Master sheet: could not detect an office-name column. "
            f"Columns found: {list(df.columns)}"
        )
    ren = {}
    if col_div:  ren[col_div]  = "division"
    if col_id:   ren[col_id]   = "office_id"
    if col_type: ren[col_type] = "office_type"
    ren[col_name] = "office_name"
    df = df.rename(columns=ren)
    df = df.dropna(subset=["office_name"]).copy()
    df["office_name"] = df["office_name"].astype(str).str.strip()
    df = df[df["office_name"].str.len() > 0]
    if "division" not in df.columns:
        df["division"] = ""
    df["division"] = df["division"].astype(str).str.strip().map(lambda x: DIV_STD.get(x, x))
    if "office_id" in df.columns:
        df["office_id"] = pd.to_numeric(df["office_id"], errors="coerce")
    else:
        df["office_id"] = np.nan
    if "office_type" not in df.columns:
        df["office_type"] = ""
    df["office_norm"] = df["office_name"].apply(normalise)
    df["is_bo"] = df.apply(lambda r: _looks_like_bo(r["office_name"], r["office_type"]), axis=1)
    return df[["office_name", "office_id", "division", "office_type", "office_norm", "is_bo"]].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_consolidated(file_bytes):
    """
    Parse the 'Consolidated' sheet (a.k.a. 'APT Master'): the region's full
    departmental-office universe (Branch Offices already excluded).
    Returns columns: office_name, office_id, division, office_norm
    """
    xl = pd.ExcelFile(BytesIO(file_bytes))
    df = _read_sheet(xl, "Consolidated")
    col_div = _find_division_col(df.columns)
    col_id  = _find_id_col(df.columns)
    col_name = _find_office_name_col(df.columns, exclude={c for c in [col_div, col_id] if c})
    if col_name is None:
        col_name = _guess_name_col_by_content(df, {c for c in [col_div, col_id] if c})
    if col_name is None:
        raise ValueError(
            "Consolidated sheet: could not detect an office-name column. "
            f"Columns found: {list(df.columns)}"
        )
    ren = {}
    if col_div: ren[col_div] = "division"
    if col_id:  ren[col_id]  = "office_id"
    ren[col_name] = "office_name"
    df = df.rename(columns=ren)
    df = df.dropna(subset=["office_name"]).copy()
    df["office_name"] = df["office_name"].astype(str).str.strip()
    df = df[df["office_name"].str.len() > 0]
    if "division" not in df.columns:
        df["division"] = ""
    df["division"] = df["division"].astype(str).str.strip().map(lambda x: DIV_STD.get(x, x))
    if "office_id" in df.columns:
        df["office_id"] = pd.to_numeric(df["office_id"], errors="coerce")
    else:
        df["office_id"] = np.nan
    df["office_norm"] = df["office_name"].apply(normalise)
    return df[["office_name", "office_id", "division", "office_norm"]].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_aebas_export(file_bytes):
    """
    Parse the AEBAS portal export CSV.
    Returns the full dataframe plus a normalised 'office_norm' column and
    a cleaned 'Status' column ('P' / 'A' / other).
    """
    df = pd.read_csv(BytesIO(file_bytes))
    df.columns = [c.strip() for c in df.columns]
    if "Office Location" not in df.columns:
        raise ValueError("Export CSV: 'Office Location' column not found.")
    df["office_norm"] = df["Office Location"].apply(normalise)
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip().str.upper()
    else:
        df["Status"] = ""
    return df

def match_offices(query_norms, master_norms, threshold):
    """query_norms: unique list. Returns dict query_norm -> matched master_norm (or None)."""
    master_list = list(master_norms)
    out = {}
    for q in query_norms:
        if not q:
            out[q] = None
            continue
        result = process.extractOne(q, master_list, scorer=fuzz.partial_ratio)
        out[q] = result[0] if result and result[1] >= threshold else None
    return out

def match_offices_to_index(query_norms, master_df, threshold):
    """Match each unique query norm to a single row (by positional index) of master_df.
    Matching to a row index (rather than name text) avoids double-counting if two
    master rows ever share the same normalised name."""
    choices = master_df["office_norm"].tolist()
    out = {}
    for q in query_norms:
        if not q:
            out[q] = None
            continue
        result = process.extractOne(q, choices, scorer=fuzz.partial_ratio)
        out[q] = result[2] if result and result[1] >= threshold else None
    return out

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(summary_df, not_marked_df, office_wise_df, report_date):
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

    ws1.merge_range(0, 0, 0, 5, f"AEBAS Report dated {date_str}", fmt_title)
    for ci, h in enumerate([
        "Sl.\nNo.", "Name of the\nDivision/Unit", "Total no.\nof Units",
        "No. of offices\nmarked attendance\nin AEBAS",
        "No. of offices not\nmarked attendance\nin AEBAS",
        "% of offices\nimplemented\nAEBAS",
    ]):
        ws1.write(1, ci, h, fmt_hdr)

    data_rows = summary_df[summary_df["Division/Unit"] != "TOTAL HQ REGION"]
    ri = 2
    for _, row in data_rows.iterrows():
        v = float(row["% AEBAS"])
        pf = (fmt_green if v >= 95 else fmt_lgreen if v >= 80 else
              fmt_amber if v >= 60 else fmt_red)
        ws1.write(ri, 0, int(row["Sl."]), fmt_c)
        ws1.write(ri, 1, row["Division/Unit"], fmt_l)
        ws1.write(ri, 2, int(row["Total"]), fmt_c)
        ws1.write(ri, 3, int(row["Marked"]), fmt_c)
        ws1.write(ri, 4, int(row["Not Marked"]), fmt_c)
        ws1.write(ri, 5, round(v, 2), pf)
        ri += 1

    tot = summary_df[summary_df["Division/Unit"] == "TOTAL HQ REGION"].iloc[0]
    ws1.write(ri, 0, "", fmt_total)
    ws1.write(ri, 1, "TOTAL", fmt_total_l)
    ws1.write(ri, 2, int(tot["Total"]), fmt_total)
    ws1.write(ri, 3, int(tot["Marked"]), fmt_total)
    ws1.write(ri, 4, int(tot["Not Marked"]), fmt_total)
    ws1.write(ri, 5, round(float(tot["% AEBAS"]), 2), fmt_total)
    ri += 2
    ws1.write(ri, 1, "* Departmental Post Offices only. Branch Offices (BOs) excluded.", fmt_note)

    # ── Sheet 2: Not Marked (merged Division column) ──────────────────────────
    ws2 = wb.add_worksheet("Not Marked")
    ws2.set_column(0, 0, 6)
    ws2.set_column(1, 1, 28)
    ws2.set_column(2, 2, 40)
    ws2.set_row(0, 24)
    ws2.set_row(1, 30)

    ws2.merge_range(0, 0, 0, 2,
        f"List of offices not marked attendance in AEBAS portal as on {date_str}", fmt_title)
    ws2.write(1, 0, "Sl. No.", fmt_hdr)
    ws2.write(1, 1, "Name of the Division", fmt_hdr)
    ws2.write(1, 2, "Name of the Unit/Office", fmt_hdr)

    not_marked_df = not_marked_df.copy()
    not_marked_df["_order"] = not_marked_df["division"].map(lambda x: DIV_ORDER_MAP.get(x, 99))
    not_marked_df = not_marked_df.sort_values(["_order", "division", "office_name"]).reset_index(drop=True)

    ri2 = 2; sl = 1
    for div, grp in not_marked_df.groupby("division", sort=False, observed=True):
        offices = grp["office_name"].tolist()
        start = ri2
        for oname in offices:
            ws2.write(ri2, 0, sl, fmt_c)
            ws2.write(ri2, 2, oname, fmt_notmk)
            ri2 += 1; sl += 1
        if start == ri2 - 1:
            ws2.write(start, 1, div, fmt_div)
        else:
            ws2.merge_range(start, 1, ri2 - 1, 1, div, fmt_div)

    # ── Sheet 3: Office-wise Attendance ────────────────────────────────────────
    ws3 = wb.add_worksheet("Office-wise Attendance")
    ws3.set_column(0, 0, 6)
    ws3.set_column(1, 1, 28)
    ws3.set_column(2, 2, 34)
    ws3.set_column(3, 3, 14)
    ws3.set_column(4, 6, 12)
    ws3.set_row(0, 24)
    ws3.set_row(1, 34)

    ws3.merge_range(0, 0, 0, 6,
        f"Office-wise Number of Users Marked Attendance – {date_str}", fmt_title)
    for ci, h in enumerate(["Sl.\nNo.", "Name of the\nDivision", "Name of the\nOffice",
                             "Office ID", "Present", "Absent", "Total"]):
        ws3.write(1, ci, h, fmt_hdr)

    ri3 = 2; sl = 1
    for _, row in office_wise_df.iterrows():
        ws3.write(ri3, 0, sl, fmt_c)
        ws3.write(ri3, 1, row["division"], fmt_l)
        ws3.write(ri3, 2, row["office_name"], fmt_l)
        oid = row["office_id"]
        ws3.write(ri3, 3, "" if pd.isna(oid) else int(oid), fmt_c)
        ws3.write(ri3, 4, int(row["Present"]), fmt_c)
        ws3.write(ri3, 5, int(row["Absent"]), fmt_c)
        ws3.write(ri3, 6, int(row["Total"]), fmt_c)
        ri3 += 1; sl += 1

    tp, ta = office_wise_df["Present"].sum(), office_wise_df["Absent"].sum()
    ws3.write(ri3, 0, "", fmt_total)
    ws3.write(ri3, 1, "TOTAL", fmt_total_l)
    ws3.write(ri3, 2, "", fmt_total)
    ws3.write(ri3, 3, "", fmt_total)
    ws3.write(ri3, 4, int(tp), fmt_total)
    ws3.write(ri3, 5, int(ta), fmt_total)
    ws3.write(ri3, 6, int(tp + ta), fmt_total)
    ri3 += 2
    ws3.write(ri3, 1, "* Departmental Post Offices only. Branch Offices (BOs) excluded.", fmt_note)

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

# ── Sidebar: report date + required export upload ────────────────────────────
_render_nav()
st.sidebar.title("🖐 AEBAS Monitoring")

today = date.today()
default_date = today - timedelta(days=3 if today.weekday() == 0 else 1)
report_date = st.sidebar.date_input("Report Date", value=default_date)

st.sidebar.markdown("---")
aebas_file = st.sidebar.file_uploader(
    "AEBAS Export CSV",
    type=["csv"],
    help="Downloaded from AEBAS portal — contains 'Office Location' and 'Status' columns"
)

FUZZY_THRESHOLD = 60  # matching is against exact office-code-based master names, so a
                       # single sensible default is used instead of an exposed slider.

# ── Main area: optional master overrides (one row) ───────────────────────────
st.subheader("📂 Master Data")

bundled_path = find_bundled_master()
bundled_bytes = None
if bundled_path:
    with open(bundled_path, "rb") as fh:
        bundled_bytes = fh.read()

st.caption(
    f"✅ Bundled master found: `{os.path.basename(bundled_path)}`" if bundled_path
    else "ℹ️ No bundled master file found next to the app — upload both sheets below."
)

row_c1, row_c2 = st.columns(2)
with row_c1:
    office_master_override = st.file_uploader(
        "Optional: AEBAS Master (Office Master sheet) — override",
        type=["xlsx"],
        help="Full office list incl. Branch Offices, used only as the name-matching dictionary."
    )
with row_c2:
    consolidated_override = st.file_uploader(
        "Optional: APT Master (Consolidated sheet) — override",
        type=["xlsx"],
        help="Region's departmental-office universe (Branch Offices already excluded)."
    )

st.markdown("<hr style='margin:8px 0 16px 0;border-color:#eee;'>", unsafe_allow_html=True)

# ── Gate ──────────────────────────────────────────────────────────────────────
office_master_bytes = office_master_override.read() if office_master_override else bundled_bytes
consolidated_bytes = consolidated_override.read() if consolidated_override else bundled_bytes

if not aebas_file:
    st.info(
        "Upload the **AEBAS Export CSV** above to generate the reports.\n\n"
        "Master office data is loaded automatically from the bundled master workbook "
        "(Office Master + Consolidated sheets). You can override either sheet "
        "individually using the optional uploaders above."
    )
    st.stop()

missing = []
if office_master_bytes is None: missing.append("AEBAS Master (Office Master sheet)")
if consolidated_bytes is None: missing.append("APT Master (Consolidated sheet)")
if missing:
    st.error(
        "Missing master data: " + ", ".join(missing) + ". "
        "Either bundle the master workbook next to the app, or upload the missing sheet(s) above."
    )
    st.stop()

# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Loading master and AEBAS data..."):
    office_master_df = load_office_master(office_master_bytes)
    consolidated_df = load_consolidated(consolidated_bytes)
    export_df = load_aebas_export(aebas_file.read())

export_norms_unique = export_df["office_norm"].dropna().unique().tolist()
export_norms_unique = [n for n in export_norms_unique if n]

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 1 — Division-wise % (Consolidated / departmental offices only)
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Matching offices for Division-wise report..."):
    marked_norms = match_offices(export_norms_unique, consolidated_df["office_norm"].tolist(), FUZZY_THRESHOLD)
    marked_set = set(v for v in marked_norms.values() if v)

master_df = consolidated_df.copy()
master_df["Marked"] = master_df["office_norm"].isin(marked_set)

summary = (
    master_df.groupby("division")
    .agg(Total=("office_norm", "count"), Marked=("Marked", "sum"))
    .reset_index()
)
summary["Not Marked"] = summary["Total"] - summary["Marked"]
summary["% AEBAS"] = (summary["Marked"] / summary["Total"] * 100).round(2)
summary = summary.sort_values("% AEBAS", ascending=False).reset_index(drop=True)
summary.insert(0, "Sl.", range(1, len(summary) + 1))
summary.columns = ["Sl.", "Division/Unit", "Total", "Marked", "Not Marked", "% AEBAS"]

tot_t = summary["Total"].sum()
tot_m = summary["Marked"].sum()
tot_row = pd.DataFrame([{
    "Sl.": "", "Division/Unit": "TOTAL HQ REGION",
    "Total": int(tot_t), "Marked": int(tot_m), "Not Marked": int(tot_t - tot_m),
    "% AEBAS": round(tot_m / tot_t * 100, 2) if tot_t else 0,
}])
summary_display = pd.concat([summary, tot_row], ignore_index=True)

st.subheader(f"1️⃣ Division-wise % — AEBAS Report dated {report_date.strftime('%d.%m.%Y')}")

def _style_summary(row):
    if row["Division/Unit"] == "TOTAL HQ REGION":
        return [f"background-color:{TOTAL_BG};color:{TOTAL_FG};font-weight:700"] * len(row)
    try:
        v = float(row["% AEBAS"])
        pct_style = ("background-color:#70AD47;color:white;font-weight:700" if v >= 95 else
                     "background-color:#E2EFDA;font-weight:600" if v >= 80 else
                     "background-color:#FFC000;font-weight:700" if v >= 60 else
                     "background-color:#FF0000;color:white;font-weight:700")
        div_style = ("background-color:#E2EFDA" if v >= 80 else
                     "background-color:#FFF2CC" if v >= 60 else
                     f"background-color:{NOTMK_BG}")
        return ["", div_style, "", "", "", pct_style]
    except Exception:
        return [""] * len(row)

styled = (
    summary_display.style
    .apply(_style_summary, axis=1)
    .format({"% AEBAS": "{:.2f}%"})
)
st.dataframe(styled, use_container_width=True, hide_index=True)
st.caption("* Departmental Post Offices only. Branch Offices (BOs) excluded.")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Offices", int(tot_t))
c2.metric("Marked", int(tot_m))
c3.metric("Not Marked", int(tot_t - tot_m))
c4.metric("% Implementation", f"{round(tot_m/tot_t*100,2) if tot_t else 0}%")

not_marked = master_df[~master_df["Marked"]].copy()
not_marked["_order"] = not_marked["division"].map(lambda x: DIV_ORDER_MAP.get(x, 99))
not_marked = not_marked.sort_values(["_order", "division", "office_name"]).reset_index(drop=True)

st.markdown("---")
st.subheader(f"List of offices not marked attendance in AEBAS portal as on {report_date.strftime('%d.%m.%Y')}")

rows = []; sl = 1; prev_div = None
for _, row in not_marked.iterrows():
    rows.append({
        "Sl. No.": sl,
        "Name of the Division": row["division"] if row["division"] != prev_div else "",
        "Name of the Unit/Office": row["office_name"],
    })
    prev_div = row["division"]
    sl += 1
nm_display = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["Sl. No.", "Name of the Division", "Name of the Unit/Office"])

def _style_nm(row):
    div = str(row["Name of the Division"]).strip()
    if div and div != "nan":
        return [f"background-color:{DIV_BG};color:{DIV_FG};font-weight:700"] * 3
    return ["", "", f"background-color:{NOTMK_BG}"]

st.dataframe(nm_display.style.apply(_style_nm, axis=1), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# REPORT 2 — Office-wise number of users marked attendance
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.subheader(f"2️⃣ Office-wise Number of Users Marked Attendance — {report_date.strftime('%d.%m.%Y')}")

office_master_nonbo = office_master_df[~office_master_df["is_bo"]].reset_index(drop=True)

with st.spinner("Matching offices for Office-wise report..."):
    match_idx_map = match_offices_to_index(export_norms_unique, office_master_nonbo, FUZZY_THRESHOLD)

export_df["matched_idx"] = export_df["office_norm"].map(match_idx_map)
unmatched_mask = export_df["matched_idx"].isna()
matched_rows = export_df[~unmatched_mask].copy()
matched_rows["matched_idx"] = matched_rows["matched_idx"].astype(int)

if len(matched_rows):
    office_info = office_master_nonbo.loc[matched_rows["matched_idx"].values].reset_index(drop=True)
    matched_rows = matched_rows.reset_index(drop=True)
    matched_rows["m_division"] = office_info["division"]
    matched_rows["m_office_name"] = office_info["office_name"]
    matched_rows["m_office_id"] = office_info["office_id"]

    office_wise = (
        matched_rows.groupby(["m_division", "m_office_name", "m_office_id"])
        .agg(Present=("Status", lambda s: (s == "P").sum()),
             Absent=("Status", lambda s: (s == "A").sum()))
        .reset_index()
        .rename(columns={"m_division": "division", "m_office_name": "office_name", "m_office_id": "office_id"})
    )
    office_wise["Total"] = office_wise["Present"] + office_wise["Absent"]
    office_wise["_order"] = office_wise["division"].map(lambda x: DIV_ORDER_MAP.get(x, 99))
    office_wise = office_wise.sort_values(["_order", "division", "office_name"]).drop(columns="_order").reset_index(drop=True)
else:
    office_wise = pd.DataFrame(columns=["division", "office_name", "office_id", "Present", "Absent", "Total"])

ow_display = office_wise.copy()
ow_display.insert(0, "Sl.", range(1, len(ow_display) + 1))
ow_display["office_id"] = ow_display["office_id"].apply(lambda v: "" if pd.isna(v) else int(v))
ow_display = ow_display.rename(columns={
    "division": "Name of the Division", "office_name": "Name of the Office", "office_id": "Office ID"
})

st.dataframe(ow_display, use_container_width=True, hide_index=True)

oc1, oc2, oc3 = st.columns(3)
oc1.metric("Offices with Attendance Data", len(office_wise))
oc2.metric("Total Present", int(office_wise["Present"].sum()) if len(office_wise) else 0)
oc3.metric("Total Absent", int(office_wise["Absent"].sum()) if len(office_wise) else 0)
st.caption("* Departmental Post Offices only. Branch Offices (BOs) excluded.")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
excel_bytes = build_excel(summary_display, not_marked, office_wise, report_date)
st.download_button(
    "⬇️ Download AEBAS Report (Excel)",
    data=excel_bytes,
    file_name=f"AEBAS_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ── Debug expander ────────────────────────────────────────────────────────────
with st.expander("🔍 Matching Debug", expanded=False):
    st.caption(
        f"Consolidated (departmental) offices: {len(consolidated_df)} | "
        f"Office Master rows: {len(office_master_df)} (non-BO: {len(office_master_nonbo)}) | "
        f"AEBAS export unique locations: {len(set(export_norms_unique))} | "
        f"Matched to Consolidated: {len(marked_set)} | "
        f"Matched to Office Master (non-BO): {export_df['matched_idx'].notna().sum()} export rows"
    )
    st.markdown("**Consolidated offices NOT matched (shown as 'Not Marked' above):**")
    st.dataframe(
        not_marked[["office_name", "division"]].rename(columns={"office_name": "Office", "division": "Division"}),
        use_container_width=True, hide_index=True
    )
    unmatched_locs = export_df.loc[unmatched_mask, "Office Location"].dropna().unique().tolist()
    st.markdown(f"**Export locations that did not match any non-BO office ({len(unmatched_locs)}):**")
    st.dataframe(pd.DataFrame({"Unmatched Office Location": unmatched_locs}),
                 use_container_width=True, hide_index=True)
