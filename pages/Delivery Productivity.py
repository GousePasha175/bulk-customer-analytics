import streamlit as st
import pandas as pd
import numpy as np
import glob as _glob
import io
import os
import re
from datetime import date
import xlsxwriter
from PIL import Image

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f3e0 Home")
    for pat, lbl in [
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py","\U0001f91a AEBAS Monitoring"),
        ("pages/Bulk Analytics.py|pages/*[Bb]ulk*.py",           "\U0001f4ca Bulk Customer Analytics"),
        ("pages/Delivery Productivity.py|pages/*[Dd]elivery*.py", "\U0001f4e6 Delivery Productivity"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py", "\U0001f4bb Digital Transactions"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
        ("pages/Sorting_Assistance.py|pages/*[Ss]orting*.py", "📮 Sorting Assistance"),
    ]:
        hits = []
        for p in pat.split("|"): hits += _glob.glob(p)
        if hits: st.sidebar.page_link(hits[0].replace("\\", "/"), label=lbl)
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

st.set_page_config(
    page_title="Delivery Transit Analysis (D+0)",
    page_icon="\U0001f4e6",
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
.dtx-table{border-collapse:collapse;width:100%;font-size:13.5px;font-family:Arial,Helvetica,sans-serif;margin-bottom:6px;}
.dtx-table th{border:1px solid #444;background:#EDEDED;font-weight:700;text-align:center;padding:6px 6px;color:#1a1a1a;}
.dtx-table td{border:1px solid #999;text-align:center;padding:5px 6px;color:#1a1a1a;}
.dtx-table td.name{text-align:left;font-weight:600;}
#.dtx-title{background:linear-gradient(180deg,#c1615c,#a94f4a);color:#fff;font-weight:700;
 #   text-align:center;padding:10px 6px;font-size:15.5px;border:1px solid #444;letter-spacing:.2px;}
.dtx-title{
    background:#003366;   /* Dark Blue */
    color:#FFFFFF;        /* White text */
    font-weight:700;
    text-align:center;
    padding:10px 6px;
    font-size:15.5px;
    border:1px solid #444;
    letter-spacing:.2px;
}
.dtx-total{background:#FFFF00!important;font-weight:700!important;}
</style>""", unsafe_allow_html=True)

if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Column mapping / constants ────────────────────────────────────────────────
def _canon(col):
    c = str(col).strip().lower()
    return re.sub(r'[^a-z0-9]', '', c)

COL_MAP = {
    "officeid": "office_id", "officename": "office_name", "officetype": "office_type",
    "productcode": "product_code", "productcategory": "product_category",
    "customerid": "customer_id", "contractid": "contract_id",
    "received": "received", "samedayinvoiced": "invoiced",
    "d0delivered": "d0_delivered", "d0redirected": "d0_redirected", "d0returned": "d0_returned",
}
REQUIRED = ["office_id", "office_name", "received", "invoiced",
            "d0_delivered", "d0_redirected", "d0_returned"]
METRICS = ["received", "invoiced", "d0_delivered", "d0_redirected", "d0_returned"]

HEADERS = ["Name of the Division", "Received", "Invoiced", "D+0 Delivered", "D+0 Redirected",
           "D+0 Returned", "D+O (Delivered+Retuned+Redirected)",
           "%D+O (Received)", "%D+O (Invoiced)", "% D+0 successful delivery"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def strip_div_word(name):
    """Column header already says 'Division' — drop the trailing word from each name."""
    return re.sub(r'\s*Division\s*$', '', str(name), flags=re.IGNORECASE).strip()

def indian_number(v):
    """Format an integer using the Indian numbering system, e.g. 110000000 -> 11,00,00,000."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    n = int(round(v))
    neg = n < 0
    n = abs(n)
    s = str(n)
    if len(s) <= 3:
        res = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        res = ",".join(parts) + "," + last3
    return ("-" if neg else "") + res

def drop_empty_bo_rows(df):
    """Remove offices that have no BO presence at all in this file (e.g. Hyderabad GPO)."""
    mask = ~(df["received"].isna() | (df["received"] == 0))
    return df[mask].reset_index(drop=True)

def read_transit_csv(uploaded_file):
    """Read + canonicalise a D+0 transit CSV. Returns (aggregated_df, category_list)."""
    try:
        raw = pd.read_csv(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        raw = pd.read_csv(uploaded_file, sep=None, engine="python")
    if raw.shape[1] == 1:
        uploaded_file.seek(0)
        raw = pd.read_csv(uploaded_file, sep=None, engine="python")

    ren = {c: COL_MAP[_canon(c)] for c in raw.columns if _canon(c) in COL_MAP}
    df = raw.rename(columns=ren)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")
    if "product_category" not in df.columns:
        df["product_category"] = "ALL"

    df["office_id"] = df["office_id"].astype(str).str.strip()
    df["office_name"] = df["office_name"].astype(str).str.strip()
    for c in METRICS:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    agg = df.groupby(["office_id", "office_name"], as_index=False)[METRICS].sum()
    agg["_sortkey"] = pd.to_numeric(agg["office_id"], errors="coerce")
    agg = agg.sort_values("_sortkey").drop(columns="_sortkey").reset_index(drop=True)
    categories = sorted(set(df["product_category"].dropna().astype(str).str.strip()))
    return agg, categories

def reindex_to_master(df, master):
    return master.merge(df, on=["office_id", "office_name"], how="left")

def add_metrics(df):
    """Add D+O and the three % columns. Missing denominators -> None (rendered as #DIV/0!)."""
    df = df.copy()
    df["d_plus_0"] = df[["d0_delivered", "d0_redirected", "d0_returned"]].sum(axis=1, min_count=0)

    def pctcol(num, den):
        out = []
        for n, d in zip(num, den):
            out.append(None if (pd.isna(d) or d == 0) else round(n / d * 100, 2))
        return out

    df["pct_received"] = pctcol(df["d_plus_0"], df["received"])
    df["pct_invoiced"] = pctcol(df["d_plus_0"], df["invoiced"])
    df["pct_success"] = pctcol(df["d0_delivered"], df["received"])
    return df

def total_row(df):
    s = df[METRICS + ["d_plus_0"]].sum(numeric_only=True)
    out = dict(s)
    out["pct_received"] = round(out["d_plus_0"] / out["received"] * 100, 2) if out["received"] else None
    out["pct_invoiced"] = round(out["d_plus_0"] / out["invoiced"] * 100, 2) if out["invoiced"] else None
    out["pct_success"] = round(out["d0_delivered"] / out["received"] * 100, 2) if out["received"] else None
    return out

def make_master(*dfs):
    parts = [d[["office_id", "office_name"]] for d in dfs if d is not None]
    m = pd.concat(parts).drop_duplicates()
    m["_sortkey"] = pd.to_numeric(m["office_id"], errors="coerce")
    return m.sort_values("_sortkey").drop(columns="_sortkey").reset_index(drop=True)

def combine_sum(df_a, df_b, master):
    """Element-wise sum of two category files (e.g. Speed Parcel + Regd. Parcel), aligned to master."""
    a = reindex_to_master(df_a, master)[METRICS].fillna(0) if df_a is not None else pd.DataFrame(0, index=master.index, columns=METRICS)
    b = reindex_to_master(df_b, master)[METRICS].fillna(0) if df_b is not None else pd.DataFrame(0, index=master.index, columns=METRICS)
    out = master.copy()
    for m in METRICS:
        out[m] = a[m].values + b[m].values
    return out

def subtract(df_all, df_bo, master):
    """Excluding-BO = All Offices − BOs (missing BO office treated as 0)."""
    a = reindex_to_master(df_all, master)[METRICS].fillna(0)
    b = reindex_to_master(df_bo, master)[METRICS].fillna(0)
    out = master.copy()
    for m in METRICS:
        out[m] = a[m].values - b[m].values
    return out

def validate_category(categories, expected_keywords, label):
    if not categories:
        return
    found_ok = any(any(k.lower() in c.lower() for k in expected_keywords) for c in categories)
    if not found_ok:
        st.sidebar.warning(
            f"⚠️ {label}: expected category containing {', '.join(expected_keywords)}, "
            f"but file shows: {', '.join(categories)}. Report will still be generated."
        )

# ── HTML rendering ────────────────────────────────────────────────────────────
def _fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return indian_number(v)

def _fmt_pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "#DIV/0!"
    return f"{v:.2f}%"

def render_html_table(title, df):
    rows_html = ""
    for _, r in df.iterrows():
        name = strip_div_word(r["office_name"])
        rows_html += (
            "<tr>"
            f"<td class='name'>{name}</td>"
            f"<td>{_fmt_num(r['received'])}</td>"
            f"<td>{_fmt_num(r['invoiced'])}</td>"
            f"<td>{_fmt_num(r['d0_delivered'])}</td>"
            f"<td>{_fmt_num(r['d0_redirected'])}</td>"
            f"<td>{_fmt_num(r['d0_returned'])}</td>"
            f"<td>{_fmt_num(r['d_plus_0'])}</td>"
            f"<td>{_fmt_pct(r['pct_received'])}</td>"
            f"<td>{_fmt_pct(r['pct_invoiced'])}</td>"
            f"<td>{_fmt_pct(r['pct_success'])}</td>"
            "</tr>"
        )
    tot = total_row(df)
    rows_html += (
        "<tr class='dtx-total'>"
        "<td class='name'>Total</td>"
        f"<td>{_fmt_num(tot['received'])}</td>"
        f"<td>{_fmt_num(tot['invoiced'])}</td>"
        f"<td>{_fmt_num(tot['d0_delivered'])}</td>"
        f"<td>{_fmt_num(tot['d0_redirected'])}</td>"
        f"<td>{_fmt_num(tot['d0_returned'])}</td>"
        f"<td>{_fmt_num(tot['d_plus_0'])}</td>"
        f"<td>{_fmt_pct(tot['pct_received'])}</td>"
        f"<td>{_fmt_pct(tot['pct_invoiced'])}</td>"
        f"<td>{_fmt_pct(tot['pct_success'])}</td>"
        "</tr>"
    )
    thead = "".join(f"<th>{h}</th>" for h in HEADERS)
    html = f"""
    <table class="dtx-table">
      <tr><td colspan="10" class="dtx-title">{title}</td></tr>
      <tr>{thead}</tr>
      {rows_html}
    </table>
    """
    return html

# ── Excel sheet writer ────────────────────────────────────────────────────────
def write_excel_sheet(wb, sheet_name, title, df):
    ws = wb.add_worksheet(sheet_name[:31])
    ws.set_column(0, 0, 26)
    ws.set_column(1, 9, 13)

    fmt_title = wb.add_format({"bold": True, "font_size": 12, "align": "center", "valign": "vcenter",
                                "bg_color": "#A94F4A", "font_color": "#FFFFFF", "border": 1})
    fmt_hdr = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True,
                              "bg_color": "#EDEDED", "border": 1})
    fmt_name = wb.add_format({"align": "left", "border": 1})
    fmt_num = wb.add_format({"align": "center", "border": 1})
    fmt_pct = wb.add_format({"align": "center", "border": 1})
    fmt_blank = wb.add_format({"align": "center", "border": 1})
    fmt_tot_name = wb.add_format({"align": "left", "border": 1, "bold": True, "bg_color": "#FFFF00"})
    fmt_tot_num = wb.add_format({"align": "center", "border": 1, "bold": True, "bg_color": "#FFFF00"})
    fmt_tot_pct = wb.add_format({"align": "center", "border": 1, "bold": True, "bg_color": "#FFFF00"})

    ws.merge_range(0, 0, 0, 9, title, fmt_title)
    for ci, h in enumerate(HEADERS):
        ws.write(1, ci, h, fmt_hdr)

    def wnum(row, col, v, fmt):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            ws.write_blank(row, col, None, fmt_blank)
        else:
            ws.write(row, col, indian_number(v), fmt)

    def wpct(row, col, v, fmt):
        ws.write(row, col, _fmt_pct(v), fmt)

    ri = 2
    for _, r in df.iterrows():
        ws.write(ri, 0, strip_div_word(r["office_name"]), fmt_name)
        wnum(ri, 1, r["received"], fmt_num)
        wnum(ri, 2, r["invoiced"], fmt_num)
        wnum(ri, 3, r["d0_delivered"], fmt_num)
        wnum(ri, 4, r["d0_redirected"], fmt_num)
        wnum(ri, 5, r["d0_returned"], fmt_num)
        wnum(ri, 6, r["d_plus_0"], fmt_num)
        wpct(ri, 7, r["pct_received"], fmt_pct)
        wpct(ri, 8, r["pct_invoiced"], fmt_pct)
        wpct(ri, 9, r["pct_success"], fmt_pct)
        ri += 1

    tot = total_row(df)
    ws.write(ri, 0, "Total", fmt_tot_name)
    ws.write(ri, 1, indian_number(tot["received"]), fmt_tot_num)
    ws.write(ri, 2, indian_number(tot["invoiced"]), fmt_tot_num)
    ws.write(ri, 3, indian_number(tot["d0_delivered"]), fmt_tot_num)
    ws.write(ri, 4, indian_number(tot["d0_redirected"]), fmt_tot_num)
    ws.write(ri, 5, indian_number(tot["d0_returned"]), fmt_tot_num)
    ws.write(ri, 6, indian_number(tot["d_plus_0"]), fmt_tot_num)
    ws.write(ri, 7, _fmt_pct(tot["pct_received"]), fmt_tot_pct)
    ws.write(ri, 8, _fmt_pct(tot["pct_invoiced"]), fmt_tot_pct)
    ws.write(ri, 9, _fmt_pct(tot["pct_success"]), fmt_tot_pct)

def build_master_workbook(sheets):
    """sheets: list of (sheet_name, title, df)"""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    for sheet_name, title, df in sheets:
        write_excel_sheet(wb, sheet_name, title, df)
    wb.close()
    return output.getvalue()

# ── Header ────────────────────────────────────────────────────────────────────
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None
h1, h2, _ = st.columns([1, 8, 1])
with h1:
    if logo: st.image(logo, width=90)
with h2:
    st.markdown("""<h1 style='font-size:28px;margin-bottom:2px;color:#2f3343;
font-weight:700;padding-top:4px;'>Delivery Transit Analysis (D+0)</h1>
<p style='font-size:15px;color:#555;margin-top:0;'>
Headquarters Region – Telangana Postal Circle</p>""", unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.header("Report Period")
c1, c2 = st.sidebar.columns(2)
from_date = c1.date_input("From", value=date.today())
to_date = c2.date_input("To", value=date.today())
period_str = f"{from_date.strftime('%d.%m.%Y')} to {to_date.strftime('%d.%m.%Y')}"

st.sidebar.markdown("---")
st.sidebar.header("Upload Report Files")

st.sidebar.subheader("1️⃣ All Products")
allprod_all_file = st.sidebar.file_uploader("All Offices file", type=["csv", "tsv"], key="allprod_all")
allprod_bo_file = st.sidebar.file_uploader("BOs file (optional)", type=["csv", "tsv"], key="allprod_bo")

st.sidebar.markdown("---")
st.sidebar.subheader("2️⃣ Speed Post Document (Speed Letter)")
spd_all_file = st.sidebar.file_uploader("All Offices file", type=["csv", "tsv"], key="spd_all")
spd_bo_file = st.sidebar.file_uploader("BOs file (optional)", type=["csv", "tsv"], key="spd_bo")

st.sidebar.markdown("---")
st.sidebar.subheader("3️⃣ Consolidated Parcels (Speed Parcel + Regd. Parcel)")
spp_all_file = st.sidebar.file_uploader("Speed Parcel – All Offices", type=["csv", "tsv"], key="spp_all")
rp_all_file = st.sidebar.file_uploader("Regd. Parcel – All Offices", type=["csv", "tsv"], key="rp_all")
spp_bo_file = st.sidebar.file_uploader("Speed Parcel – BOs (optional)", type=["csv", "tsv"], key="spp_bo")
rp_bo_file = st.sidebar.file_uploader("Regd. Parcel – BOs (optional)", type=["csv", "tsv"], key="rp_bo")

# ── Read files ────────────────────────────────────────────────────────────────
def safe_read(f, label):
    if f is None:
        return None, []
    try:
        df, cats = read_transit_csv(f)
        st.sidebar.success(f"✅ {label}: {len(df)} offices")
        return df, cats
    except Exception as e:
        st.sidebar.error(f"❌ {label}: {e}")
        return None, []

allprod_all, cat_ap_all = safe_read(allprod_all_file, "All Products / All Offices")
allprod_bo, cat_ap_bo = safe_read(allprod_bo_file, "All Products / BOs")
spd_all, cat_spd_all = safe_read(spd_all_file, "Speed Post Doc / All Offices")
spd_bo, cat_spd_bo = safe_read(spd_bo_file, "Speed Post Doc / BOs")
spp_all, cat_spp_all = safe_read(spp_all_file, "Speed Parcel / All Offices")
rp_all, cat_rp_all = safe_read(rp_all_file, "Regd. Parcel / All Offices")
spp_bo, cat_spp_bo = safe_read(spp_bo_file, "Speed Parcel / BOs")
rp_bo, cat_rp_bo = safe_read(rp_bo_file, "Regd. Parcel / BOs")

if spd_all is not None: validate_category(cat_spd_all, ["speed letter"], "Speed Post Doc / All Offices")
if spd_bo is not None: validate_category(cat_spd_bo, ["speed letter"], "Speed Post Doc / BOs")
if spp_all is not None: validate_category(cat_spp_all, ["speed parcel"], "Speed Parcel / All Offices")
if spp_bo is not None: validate_category(cat_spp_bo, ["speed parcel"], "Speed Parcel / BOs")
if rp_all is not None: validate_category(cat_rp_all, ["registered parcel", "regd"], "Regd. Parcel / All Offices")
if rp_bo is not None: validate_category(cat_rp_bo, ["registered parcel", "regd"], "Regd. Parcel / BOs")

any_file = any(x is not None for x in
               [allprod_all, allprod_bo, spd_all, spd_bo, spp_all, rp_all, spp_bo, rp_bo])

if not any_file:
    st.info(
        "Upload report files from the sidebar to generate D+0 Transit Analysis reports.  \n\n"
        "**1) All Products** — upload the All Offices file, and optionally the BOs file "
        "(to also generate the BOs and Excluding-BOs reports).  \n"
        "**2) Speed Post Document** — same pattern, for files where Product Category = *Speed Letter*.  \n"
        "**3) Consolidated Parcels (Speed Parcel + India Post/Regd. Parcel)** — upload up to 4 files: "
        "Speed Parcel & Regd. Parcel, each for All Offices and (optionally) BOs. The two are summed together, "
        "and Excluding-BOs = combined All Offices − combined BOs."
    )
    st.stop()

st.markdown(f"""
<div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
border-radius:6px;margin-bottom:16px;font-size:15px;'>
<b>Report Period:</b> {period_str}
</div>
""", unsafe_allow_html=True)

SORT_OPTIONS = {
    "No sorting (office order)": None,
    "% D+0 Successful Delivery — High to Low": ("pct_success", False),
    "% D+0 Successful Delivery — Low to High": ("pct_success", True),
    "%D+O (Received) — High to Low": ("pct_received", False),
    "%D+O (Received) — Low to High": ("pct_received", True),
    "%D+O (Invoiced) — High to Low": ("pct_invoiced", False),
    "%D+O (Invoiced) — Low to High": ("pct_invoiced", True),
}
sort_choice = st.selectbox("Sort divisions by", list(SORT_OPTIONS.keys()), index=0)

def apply_sort(df):
    """Sort a table's division rows by the chosen % column (Total row is computed separately)."""
    key = SORT_OPTIONS[sort_choice]
    if key is None:
        return df
    col, ascending = key
    return df.sort_values(col, ascending=ascending, na_position="last").reset_index(drop=True)

st.markdown("<hr style='margin:2px 0 14px 0;border-color:#eee;'>", unsafe_allow_html=True)

excel_sheets = []

# ════════════════════════════════════════════════════════════════════════════
# SECTION 1 — ALL PRODUCTS
# ════════════════════════════════════════════════════════════════════════════
if allprod_all is not None:
    st.subheader("📦 All Products")
    master_ap = make_master(allprod_all, allprod_bo)

    df_ap_all = add_metrics(reindex_to_master(allprod_all, master_ap))
    df_ap_all = apply_sort(df_ap_all)
    title = f"Delivery Transit Analysis All Products -{period_str}"
    st.markdown(render_html_table(title, df_ap_all), unsafe_allow_html=True)
    excel_sheets.append(("AllProducts_AllOffices", title, df_ap_all))

    if allprod_bo is not None:
        df_ap_bo = add_metrics(reindex_to_master(allprod_bo, master_ap))
        df_ap_bo_display = apply_sort(drop_empty_bo_rows(df_ap_bo))
        title_bo = f"Delivery Transit Analysis All Products-B.O dated {period_str}"
        st.markdown(render_html_table(title_bo, df_ap_bo_display), unsafe_allow_html=True)
        excel_sheets.append(("AllProducts_BOs", title_bo, df_ap_bo_display))

        df_ap_excl = add_metrics(subtract(allprod_all, allprod_bo, master_ap))
        df_ap_excl = apply_sort(df_ap_excl)
        title_excl = f"Delivery Transit Analysis All Products (Excl. B.O) dated {period_str}"
        st.markdown(render_html_table(title_excl, df_ap_excl), unsafe_allow_html=True)
        excel_sheets.append(("AllProducts_ExclBOs", title_excl, df_ap_excl))
    else:
        st.caption("Upload the BOs file above to also generate the BOs and Excluding-BOs reports for All Products.")

# ════════════════════════════════════════════════════════════════════════════
# SECTION 2 — SPEED POST DOCUMENT
# ════════════════════════════════════════════════════════════════════════════
if spd_all is not None or spd_bo is not None:
    st.markdown("---")
    st.subheader("✉️ Speed Post Document (Speed Letter)")

    if spd_all is not None:
        master_spd = make_master(spd_all, spd_bo)
        df_spd_all = add_metrics(reindex_to_master(spd_all, master_spd))
        df_spd_all = apply_sort(df_spd_all)
        title = f"Delivery Transit Analysis Speed Post Document dated {period_str}"
        st.markdown(render_html_table(title, df_spd_all), unsafe_allow_html=True)
        excel_sheets.append(("SpeedDoc_AllOffices", title, df_spd_all))

        if spd_bo is not None:
            df_spd_bo = add_metrics(reindex_to_master(spd_bo, master_spd))
            df_spd_bo_display = apply_sort(drop_empty_bo_rows(df_spd_bo))
            title_bo = f"Delivery Transit Analysis Speed Post Document-B.O dated {period_str}"
            st.markdown(render_html_table(title_bo, df_spd_bo_display), unsafe_allow_html=True)
            excel_sheets.append(("SpeedDoc_BOs", title_bo, df_spd_bo_display))

            df_spd_excl = add_metrics(subtract(spd_all, spd_bo, master_spd))
            df_spd_excl = apply_sort(df_spd_excl)
            title_excl = f"Delivery Transit Analysis Speed Post Document (Excl. B.O) dated {period_str}"
            st.markdown(render_html_table(title_excl, df_spd_excl), unsafe_allow_html=True)
            excel_sheets.append(("SpeedDoc_ExclBOs", title_excl, df_spd_excl))
        else:
            st.caption("Upload the BOs file above to also generate the BOs and Excluding-BOs reports for Speed Post Document.")
    else:
        st.warning("Upload the All Offices file for Speed Post Document to generate this section (the BOs file alone cannot build the Excluding-BOs report).")
        master_spd = make_master(spd_bo)
        df_spd_bo = add_metrics(reindex_to_master(spd_bo, master_spd))
        df_spd_bo_display = apply_sort(drop_empty_bo_rows(df_spd_bo))
        title_bo = f"Delivery Transit Analysis Speed Post Document-B.O dated {period_str}"
        st.markdown(render_html_table(title_bo, df_spd_bo_display), unsafe_allow_html=True)
        excel_sheets.append(("SpeedDoc_BOs", title_bo, df_spd_bo_display))

# ════════════════════════════════════════════════════════════════════════════
# SECTION 3 — CONSOLIDATED PARCELS (Speed Parcel + India Post/Regd. Parcel)
# ════════════════════════════════════════════════════════════════════════════
if spp_all is not None or rp_all is not None or spp_bo is not None or rp_bo is not None:
    st.markdown("---")
    st.subheader("📮 Consolidated Parcels (Speed Parcel + India Post Parcel)")

    have_all = spp_all is not None or rp_all is not None
    have_bo = spp_bo is not None or rp_bo is not None

    if have_all:
        master_pc = make_master(spp_all, rp_all, spp_bo, rp_bo)
        df_pc_all = add_metrics(combine_sum(spp_all, rp_all, master_pc))
        title = f"Delivery Transit Analysis Consolidated Parcels (SPP+IPP R) dated {period_str}"
        st.markdown(render_html_table(title, apply_sort(df_pc_all)), unsafe_allow_html=True)
        excel_sheets.append(("Parcels_AllOffices", title, apply_sort(df_pc_all)))

        if have_bo:
            df_pc_bo = add_metrics(combine_sum(spp_bo, rp_bo, master_pc))

            df_pc_excl = add_metrics(subtract(
                df_pc_all.rename(columns={}), df_pc_bo.rename(columns={}), master_pc
            ))

            df_pc_bo_display = apply_sort(drop_empty_bo_rows(df_pc_bo))
            title_bo = f"Delivery Transit Analysis Consolidated Parcels( SPP+IPP R)-B.O dated  {period_str}"
            st.markdown(render_html_table(title_bo, df_pc_bo_display), unsafe_allow_html=True)
            excel_sheets.append(("Parcels_BOs", title_bo, df_pc_bo_display))

            df_pc_excl = apply_sort(df_pc_excl)
            title_excl = f"Delivery Transit Analysis Consolidated Parcels( SPP+IPP R) (Excl. B.O) dated {period_str}"
            st.markdown(render_html_table(title_excl, df_pc_excl), unsafe_allow_html=True)
            excel_sheets.append(("Parcels_ExclBOs", title_excl, df_pc_excl))
        else:
            st.caption("Upload the BOs files above (Speed Parcel and/or Regd. Parcel) to also generate the BOs and Excluding-BOs reports.")
    else:
        st.warning("Upload at least one All Offices file (Speed Parcel and/or Regd. Parcel) to generate this section.")

# ── Download master workbook ──────────────────────────────────────────────────
if excel_sheets:
    st.markdown("---")
    excel_bytes = build_master_workbook(excel_sheets)
    st.download_button(
        "⬇️ Download All Reports (Master Excel Workbook)",
        data=excel_bytes,
        file_name=f"Delivery_Transit_Analysis_{from_date.strftime('%d%m%Y')}_{to_date.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
