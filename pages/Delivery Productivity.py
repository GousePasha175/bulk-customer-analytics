import streamlit as st
import pandas as pd
import numpy as np
import glob as _glob
import io
import os
from datetime import date, datetime
import xlsxwriter

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    for pat, lbl in [
        ("pages/Bulk Analytics.py|pages/*[Bb]ulk*.py",           "\U0001f4ca Bulk Customer Analytics"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py", "\U0001f4bb Digital Transactions"),
        ("pages/Delivery Productivity.py|pages/*[Dd]elivery*.py", "\U0001f4e6 Delivery Productivity"),
    ]:
        hits = []
        for p in pat.split("|"): hits += _glob.glob(p)
        if hits: st.sidebar.page_link(hits[0].replace("\\", "/"), label=lbl)
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

from PIL import Image

st.set_page_config(
    page_title="Delivery Productivity Report",
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
</style>""", unsafe_allow_html=True)

if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Constants ─────────────────────────────────────────────────────────────────
PROD_COLS  = ["office-id","office-name","postman-count","invoice-count",
              "delivery-count","deposit-count","return-count",
              "redirection-count","beat-diversion-count"]
STAT_COLS  = ["office-id","office-name","product-name",
              "rec-count","del-count","ret-count","red-count"]

PROD_LABELS = {
    "postman-count":      "Postmen",
    "invoice-count":      "Invoiced",
    "delivery-count":     "Delivered",
    "deposit-count":      "Deposited",
    "return-count":       "Returned",
    "redirection-count":  "Redirected",
    "beat-diversion-count":"Beat Diversion",
}
STAT_LABELS = {
    "rec-count":  "Total Received",
    "del-count":  "Delivered",
    "ret-count":  "Returned",
    "red-count":  "Redirected",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def pct(num, den):
    if den and den > 0: return round(num / den * 100, 1)
    return 0.0

def read_csv_flexible(f):
    """Read uploaded CSV or plain file regardless of separator."""
    try:
        return pd.read_csv(f)
    except Exception:
        f.seek(0)
        return pd.read_csv(f, sep=None, engine="python")

def detect_file_type(df):
    """Return 'productivity' or 'statistics' based on columns present."""
    cols = [c.lower().strip() for c in df.columns]
    if "postman-count" in cols or "invoice-count" in cols:
        return "productivity"
    if "rec-count" in cols or "product-name" in cols:
        return "statistics"
    return "unknown"

def normalise_prod(df):
    df = df.copy()
    df.columns = [c.lower().strip() for c in df.columns]
    for col in PROD_COLS:
        if col not in df.columns: df[col] = 0
    df["office-id"] = df["office-id"].astype(str).str.strip()
    df["office-name"] = df["office-name"].astype(str).str.strip()
    # Remove summary rows (office-id = 0 or starts with 0 and has 'summary' in name)
    df = df[~df["office-name"].str.lower().str.contains("summary")]
    df = df[df["office-id"] != "0"]
    for c in PROD_COLS[2:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df

def normalise_stat(df):
    df = df.copy()
    df.columns = [c.lower().strip() for c in df.columns]
    for col in STAT_COLS:
        if col not in df.columns: df[col] = 0
    df["office-id"] = df["office-id"].astype(str).str.strip()
    df["office-name"] = df["office-name"].astype(str).str.strip()
    df = df[~df["office-name"].str.lower().str.contains("summary")]
    df = df[df["office-id"] != "0"]
    for c in STAT_COLS[3:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    # Keep only All Products rows if multiple products
    if "product-name" in df.columns:
        all_prod = df[df["product-name"].str.lower().str.contains("all")]
        if not all_prod.empty: df = all_prod
    return df

def strip_div(name):
    import re
    return re.sub(r'\s*(division|region|circle)\s*$', '', str(name), flags=re.IGNORECASE).strip()

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(stat_df, prod_today, prod_yest, report_date, merged):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    def f(**kw):
        base = {"border": 1, "valign": "vcenter"}; base.update(kw)
        return wb.add_format(base)

    title_fmt  = f(bold=True, font_size=12, align="center", bg_color="#1F3864", font_color="#FFFFFF")
    hdr_fmt    = f(bold=True, align="center", text_wrap=True, bg_color="#2E75B6", font_color="#FFFFFF")
    sub_fmt    = f(bold=True, align="center", bg_color="#9DC3E6")
    data_fmt   = f(align="center")
    left_fmt   = f(align="left")
    total_fmt  = f(bold=True, align="center", bg_color="#FFF2CC")
    total_l    = f(bold=True, align="left",   bg_color="#FFF2CC")
    pct_good   = f(align="center", bg_color="#E2EFDA")
    pct_warn   = f(align="center", bg_color="#FFC000")
    pct_bad    = f(align="center", bg_color="#FCE4D6")

    def pct_fmt(wb_obj, v):
        if v >= 80: return pct_good
        if v >= 60: return pct_warn
        return pct_bad

    date_str = report_date.strftime("%d.%m.%Y")

    # ── Sheet 1: Statistics ───────────────────────────────────────────────────
    if stat_df is not None and not stat_df.empty:
        ws = wb.add_worksheet("Delivery Statistics")
        ws.set_column(0, 0, 5); ws.set_column(1, 1, 28)
        ws.set_column(2, 10, 14)
        ws.merge_range(0, 0, 0, 7,
            f"Delivery Statistics Report — {date_str}", title_fmt)
        hdrs = ["Sl.", "Office / Division", "Total Received",
                "Delivered", "Del %", "Returned", "Ret %",
                "Redirected", "Pending"]
        for ci, h in enumerate(hdrs): ws.write(1, ci, h, hdr_fmt)
        ri = 2
        for sl, (_, row) in enumerate(stat_df.iterrows(), 1):
            rec = int(row["rec-count"]); dl = int(row["del-count"])
            rt  = int(row["ret-count"]); rd = int(row["red-count"])
            pend = max(0, rec - dl - rt - rd)
            dp = pct(dl, rec); rp = pct(rt, rec)
            ws.write(ri, 0, sl, data_fmt)
            ws.write(ri, 1, strip_div(row["office-name"]), left_fmt)
            ws.write(ri, 2, rec,  data_fmt)
            ws.write(ri, 3, dl,   data_fmt)
            ws.write(ri, 4, f"{dp}%", pct_fmt(wb, dp))
            ws.write(ri, 5, rt,   data_fmt)
            ws.write(ri, 6, f"{rp}%", data_fmt)
            ws.write(ri, 7, rd,   data_fmt)
            ws.write(ri, 8, pend, data_fmt)
            ri += 1
        # Total row
        totals = stat_df[["rec-count","del-count","ret-count","red-count"]].sum()
        ws.write(ri, 0, "",   total_fmt)
        ws.write(ri, 1, "TOTAL", total_l)
        ws.write(ri, 2, int(totals["rec-count"]),  total_fmt)
        ws.write(ri, 3, int(totals["del-count"]),  total_fmt)
        dp_t = pct(int(totals["del-count"]), int(totals["rec-count"]))
        ws.write(ri, 4, f"{dp_t}%", total_fmt)
        ws.write(ri, 5, int(totals["ret-count"]),  total_fmt)
        rp_t = pct(int(totals["ret-count"]), int(totals["rec-count"]))
        ws.write(ri, 6, f"{rp_t}%", total_fmt)
        ws.write(ri, 7, int(totals["red-count"]),  total_fmt)
        pend_t = max(0, int(totals["rec-count"]) - int(totals["del-count"])
                       - int(totals["ret-count"]) - int(totals["red-count"]))
        ws.write(ri, 8, pend_t, total_fmt)

    # ── Sheet 2: Productivity ─────────────────────────────────────────────────
    if prod_today is not None and not prod_today.empty:
        ws2 = wb.add_worksheet("Delivery Productivity")
        ws2.set_column(0, 0, 5); ws2.set_column(1, 1, 28)
        ws2.set_column(2, 12, 13)
        ws2.merge_range(0, 0, 0, 11,
            f"Delivery Productivity Report — {date_str}", title_fmt)
        hdrs2 = ["Sl.", "Office / Division", "Postmen",
                 "Invoiced", "Delivered", "Del %",
                 "Deposited", "Returned", "Ret %",
                 "Redirected", "Beat Div.", "Articles/Postman"]
        for ci, h in enumerate(hdrs2): ws2.write(1, ci, h, hdr_fmt)
        ri = 2
        for sl, (_, row) in enumerate(prod_today.iterrows(), 1):
            pm = int(row["postman-count"]); inv = int(row["invoice-count"])
            dl = int(row["delivery-count"]); dep = int(row["deposit-count"])
            rt = int(row["return-count"]);   rd  = int(row["redirection-count"])
            bd = int(row["beat-diversion-count"])
            dp = pct(dl, inv); rp = pct(rt, inv)
            app = round(inv / pm, 1) if pm > 0 else 0
            ws2.write(ri, 0, sl, data_fmt)
            ws2.write(ri, 1, strip_div(row["office-name"]), left_fmt)
            ws2.write(ri, 2, pm,  data_fmt); ws2.write(ri, 3, inv, data_fmt)
            ws2.write(ri, 4, dl,  data_fmt); ws2.write(ri, 5, f"{dp}%", pct_fmt(wb, dp))
            ws2.write(ri, 6, dep, data_fmt); ws2.write(ri, 7, rt,  data_fmt)
            ws2.write(ri, 8, f"{rp}%", data_fmt)
            ws2.write(ri, 9, rd,  data_fmt); ws2.write(ri, 10, bd,  data_fmt)
            ws2.write(ri, 11, app, data_fmt)
            ri += 1
        totals = prod_today[["postman-count","invoice-count","delivery-count",
                              "deposit-count","return-count","redirection-count",
                              "beat-diversion-count"]].sum()
        ws2.write(ri, 0, "", total_fmt); ws2.write(ri, 1, "TOTAL", total_l)
        ws2.write(ri, 2, int(totals["postman-count"]),  total_fmt)
        ws2.write(ri, 3, int(totals["invoice-count"]),  total_fmt)
        ws2.write(ri, 4, int(totals["delivery-count"]), total_fmt)
        dp_t = pct(int(totals["delivery-count"]), int(totals["invoice-count"]))
        ws2.write(ri, 5, f"{dp_t}%", total_fmt)
        ws2.write(ri, 6, int(totals["deposit-count"]),  total_fmt)
        ws2.write(ri, 7, int(totals["return-count"]),   total_fmt)
        rp_t = pct(int(totals["return-count"]), int(totals["invoice-count"]))
        ws2.write(ri, 8, f"{rp_t}%", total_fmt)
        ws2.write(ri, 9, int(totals["redirection-count"]),   total_fmt)
        ws2.write(ri, 10, int(totals["beat-diversion-count"]), total_fmt)
        app_t = round(int(totals["invoice-count"]) / int(totals["postman-count"]), 1) \
            if int(totals["postman-count"]) > 0 else 0
        ws2.write(ri, 11, app_t, total_fmt)

    # ── Sheet 3: Carryforward Analysis ───────────────────────────────────────
    if prod_today is not None and prod_yest is not None:
        ws3 = wb.add_worksheet("Carryforward Analysis")
        ws3.set_column(0, 0, 5); ws3.set_column(1, 1, 28)
        ws3.set_column(2, 10, 16)
        ws3.merge_range(0, 0, 0, 9,
            f"Carryforward Analysis — Yesterday deposits vs Today load", title_fmt)
        hdrs3 = ["Sl.", "Office / Division",
                 "Yest. Deposited (Carryforward)",
                 "Today Invoiced", "Today Received (Stat.)",
                 "Carryforward %\n(of Today Invoiced)",
                 "Today Delivered", "Today Deposited (New CF)",
                 "Today Returned", "Net Clearance %"]
        for ci, h in enumerate(hdrs3): ws3.write(1, ci, h, hdr_fmt)
        ri = 2
        for sl, (_, row) in enumerate(merged.iterrows(), 1):
            cf   = int(row.get("yest_deposit", 0))
            inv  = int(row.get("invoice-count_today", 0))
            rec  = int(row.get("rec-count", 0))
            dl   = int(row.get("delivery-count_today", 0))
            dep  = int(row.get("deposit-count_today", 0))
            rt   = int(row.get("return-count_today", 0))
            cf_p = pct(cf, inv)
            nc_p = pct(dl, rec) if rec > 0 else pct(dl, inv)
            ws3.write(ri, 0, sl, data_fmt)
            ws3.write(ri, 1, strip_div(row["office-name"]), left_fmt)
            ws3.write(ri, 2, cf,   data_fmt)
            ws3.write(ri, 3, inv,  data_fmt)
            ws3.write(ri, 4, rec,  data_fmt)
            ws3.write(ri, 5, f"{cf_p}%", data_fmt)
            ws3.write(ri, 6, dl,   data_fmt)
            ws3.write(ri, 7, dep,  data_fmt)
            ws3.write(ri, 8, rt,   data_fmt)
            ws3.write(ri, 9, f"{nc_p}%", pct_fmt(wb, nc_p))
            ri += 1

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
font-weight:700;padding-top:4px;'>Delivery Productivity Report</h1>
<p style='font-size:15px;color:#555;margin-top:0;'>
Headquarters Region – Telangana Postal Circle</p>""", unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.header("Upload Report Files")
st.sidebar.caption(
    "Upload Division-wise OR Office-wise files — detected automatically from the Office Name column."
)

report_date = st.sidebar.date_input("Report Date (Today)", value=date.today())
yest_date   = st.sidebar.date_input("Previous Day Date",   value=date.today())

st.sidebar.markdown("---")
st.sidebar.subheader("Today's Files")
stat_file       = st.sidebar.file_uploader(
    f"Delivery Statistics ({report_date.strftime('%d.%m.%Y')})",
    type=["csv","tsv"], key="stat_today"
)
prod_today_file = st.sidebar.file_uploader(
    f"Delivery Productivity ({report_date.strftime('%d.%m.%Y')})",
    type=["csv","tsv"], key="prod_today"
)

st.sidebar.markdown("---")
st.sidebar.subheader("Previous Day's File")
prod_yest_file  = st.sidebar.file_uploader(
    f"Delivery Productivity ({yest_date.strftime('%d.%m.%Y')}) — for Carryforward",
    type=["csv","tsv"], key="prod_yest"
)

# ── Parse uploaded files ──────────────────────────────────────────────────────
stat_df = prod_today = prod_yest = None

if stat_file:
    raw = read_csv_flexible(stat_file)
    ft  = detect_file_type(raw)
    if ft == "statistics":
        stat_df = normalise_stat(raw)
        st.sidebar.success(f"✅ Statistics: {len(stat_df)} rows")
    else:
        st.sidebar.error("Statistics file not recognised — check columns.")

if prod_today_file:
    raw = read_csv_flexible(prod_today_file)
    ft  = detect_file_type(raw)
    if ft == "productivity":
        prod_today = normalise_prod(raw)
        st.sidebar.success(f"✅ Productivity (today): {len(prod_today)} rows")
    else:
        st.sidebar.error("Today's Productivity file not recognised.")

if prod_yest_file:
    raw = read_csv_flexible(prod_yest_file)
    ft  = detect_file_type(raw)
    if ft == "productivity":
        prod_yest = normalise_prod(raw)
        st.sidebar.success(f"✅ Productivity (yesterday): {len(prod_yest)} rows")
    else:
        st.sidebar.error("Yesterday's Productivity file not recognised.")

# ── Nothing uploaded yet ──────────────────────────────────────────────────────
if stat_df is None and prod_today is None:
    st.info(
        "Upload files from the sidebar to generate the report.  \n"
        "- **Delivery Statistics** = total mail received at office (superset)  \n"
        "- **Delivery Productivity** = mail invoiced to postmen for beat delivery  \n"
        "- **Previous day Productivity** = yesterday's deposits carried forward into today"
    )
    st.stop()

# ── Build merged carryforward table ──────────────────────────────────────────
merged = pd.DataFrame()
if prod_today is not None and prod_yest is not None:
    yest_cf = prod_yest[["office-id","office-name","deposit-count"]].copy()
    yest_cf = yest_cf.rename(columns={"deposit-count": "yest_deposit"})
    merged  = prod_today.merge(yest_cf[["office-id","yest_deposit"]],
                               on="office-id", how="left")
    merged["yest_deposit"] = merged["yest_deposit"].fillna(0).astype(int)
    # Rename today's cols to avoid ambiguity when stat also merged
    merged = merged.rename(columns={
        "invoice-count":  "invoice-count_today",
        "delivery-count": "delivery-count_today",
        "deposit-count":  "deposit-count_today",
        "return-count":   "return-count_today",
    })
    if stat_df is not None:
        merged = merged.merge(
            stat_df[["office-id","rec-count","del-count","ret-count","red-count"]],
            on="office-id", how="left"
        )
        merged[["rec-count","del-count","ret-count","red-count"]] = \
            merged[["rec-count","del-count","ret-count","red-count"]].fillna(0).astype(int)

# ── Display date banner ───────────────────────────────────────────────────────
st.markdown(f"""
<div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
border-radius:6px;margin-bottom:16px;font-size:15px;'>
<b>Report Date:</b> {report_date.strftime('%d.%m.%Y')}
&nbsp;&nbsp;|&nbsp;&nbsp;
<b>Files loaded:</b>
{'✅ Statistics' if stat_df is not None else '—'} &nbsp;
{'✅ Productivity' if prod_today is not None else '—'} &nbsp;
{'✅ Previous Day' if prod_yest is not None else '—'}
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DELIVERY STATISTICS
# ════════════════════════════════════════════════════════════════════════════
if stat_df is not None:
    st.subheader(f"📊 Section 1 — Delivery Statistics ({report_date.strftime('%d.%m.%Y')})")
    st.caption("Total mail received at office — includes counter deliveries, missent, RTS, and carryforward")

    disp = stat_df.copy()
    disp["Office"] = disp["office-name"].apply(strip_div)
    disp["Total Received"]  = disp["rec-count"]
    disp["Delivered"]       = disp["del-count"]
    disp["Del %"]           = disp.apply(lambda r: f"{pct(r['del-count'],r['rec-count'])}%", axis=1)
    disp["Returned"]        = disp["ret-count"]
    disp["Ret %"]           = disp.apply(lambda r: f"{pct(r['ret-count'],r['rec-count'])}%", axis=1)
    disp["Redirected"]      = disp["red-count"]
    disp["Pending"]         = (disp["rec-count"] - disp["del-count"]
                                - disp["ret-count"] - disp["red-count"]).clip(lower=0)

    show_cols = ["Office","Total Received","Delivered","Del %",
                 "Returned","Ret %","Redirected","Pending"]

    # Totals row
    tot = disp[["Total Received","Delivered","Returned","Redirected","Pending"]].sum()
    tot_row = {
        "Office": "**TOTAL**",
        "Total Received": int(tot["Total Received"]),
        "Delivered":      int(tot["Delivered"]),
        "Del %":          f"{pct(int(tot['Delivered']),int(tot['Total Received']))}%",
        "Returned":       int(tot["Returned"]),
        "Ret %":          f"{pct(int(tot['Returned']),int(tot['Total Received']))}%",
        "Redirected":     int(tot["Redirected"]),
        "Pending":        int(tot["Pending"]),
    }
    display_df = pd.concat([disp[show_cols],
                             pd.DataFrame([tot_row])], ignore_index=True)

    def _style_stat(row):
        if row["Office"] == "**TOTAL**":
            return ["background-color:#1F3864;color:white;font-weight:700"] * len(row)
        return [""] * len(row)

    def _style_del_pct(val):
        try:
            v = float(str(val).replace("%",""))
            if v >= 80: return "background-color:#70AD47;color:white;font-weight:600"
            if v >= 60: return "background-color:#FFC000;font-weight:600"
            return "background-color:#FF0000;color:white;font-weight:600"
        except: return ""

    styled = (display_df.style
              .apply(_style_stat, axis=1)
              .map(_style_del_pct, subset=["Del %"]))
    st.dataframe(styled, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 2 — DELIVERY PRODUCTIVITY
# ════════════════════════════════════════════════════════════════════════════
if prod_today is not None:
    st.markdown("---")
    st.subheader(f"📬 Section 2 — Delivery Productivity ({report_date.strftime('%d.%m.%Y')})")
    st.caption("Mail formally invoiced to postmen for beat delivery")

    disp2 = prod_today.copy()
    disp2["Office"]          = disp2["office-name"].apply(strip_div)
    disp2["Postmen"]         = disp2["postman-count"]
    disp2["Invoiced"]        = disp2["invoice-count"]
    disp2["Delivered"]       = disp2["delivery-count"]
    disp2["Del %"]           = disp2.apply(lambda r: f"{pct(r['delivery-count'],r['invoice-count'])}%", axis=1)
    disp2["Deposited"]       = disp2["deposit-count"]
    disp2["Returned"]        = disp2["return-count"]
    disp2["Ret %"]           = disp2.apply(lambda r: f"{pct(r['return-count'],r['invoice-count'])}%", axis=1)
    disp2["Redirected"]      = disp2["redirection-count"]
    disp2["Beat Div."]       = disp2["beat-diversion-count"]
    disp2["Art./Postman"]    = disp2.apply(
        lambda r: round(r["invoice-count"]/r["postman-count"],1) if r["postman-count"]>0 else 0, axis=1)

    show2 = ["Office","Postmen","Invoiced","Delivered","Del %",
             "Deposited","Returned","Ret %","Redirected","Beat Div.","Art./Postman"]

    tot2 = disp2[["Postmen","Invoiced","Delivered","Deposited",
                  "Returned","Redirected","Beat Div."]].sum()
    tot2_row = {
        "Office":       "**TOTAL**",
        "Postmen":      int(tot2["Postmen"]),
        "Invoiced":     int(tot2["Invoiced"]),
        "Delivered":    int(tot2["Delivered"]),
        "Del %":        f"{pct(int(tot2['Delivered']),int(tot2['Invoiced']))}%",
        "Deposited":    int(tot2["Deposited"]),
        "Returned":     int(tot2["Returned"]),
        "Ret %":        f"{pct(int(tot2['Returned']),int(tot2['Invoiced']))}%",
        "Redirected":   int(tot2["Redirected"]),
        "Beat Div.":    int(tot2["Beat Div."]),
        "Art./Postman": round(int(tot2["Invoiced"])/int(tot2["Postmen"]),1)
                        if int(tot2["Postmen"])>0 else 0,
    }
    display_df2 = pd.concat([disp2[show2], pd.DataFrame([tot2_row])], ignore_index=True)

    styled2 = (display_df2.style
               .apply(_style_stat, axis=1)
               .map(_style_del_pct, subset=["Del %"]))
    st.dataframe(styled2, use_container_width=True, hide_index=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 3 — CARRYFORWARD ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
if not merged.empty:
    st.markdown("---")
    st.subheader("🔄 Section 3 — Carryforward Analysis")
    st.caption(
        f"Yesterday's ({yest_date.strftime('%d.%m.%Y')}) deposited articles carried forward "
        f"into today's ({report_date.strftime('%d.%m.%Y')}) workload"
    )

    disp3 = merged.copy()
    disp3["Office"] = disp3["office-name"].apply(strip_div)
    disp3["Yest. Deposited\n(Carryforward)"] = disp3["yest_deposit"]
    disp3["Today\nInvoiced"]  = disp3["invoice-count_today"]
    has_stat = "rec-count" in disp3.columns
    disp3["Today\nReceived"]  = disp3["rec-count"] if has_stat else disp3["invoice-count_today"]
    disp3["CF %\n(of Invoiced)"] = disp3.apply(
        lambda r: f"{pct(r['yest_deposit'], r['invoice-count_today'])}%", axis=1)
    disp3["Today\nDelivered"]  = disp3["delivery-count_today"]
    disp3["New Deposit\n(Tomorrow CF)"] = disp3["deposit-count_today"]
    disp3["Today\nReturned"]   = disp3["return-count_today"]
    base_for_clearance = "rec-count" if has_stat else "invoice-count_today"
    disp3["Net Clearance %"]   = disp3.apply(
        lambda r: f"{pct(r['delivery-count_today'], r[base_for_clearance])}%", axis=1)

    show3 = ["Office",
             "Yest. Deposited\n(Carryforward)","Today\nInvoiced","Today\nReceived",
             "CF %\n(of Invoiced)","Today\nDelivered",
             "New Deposit\n(Tomorrow CF)","Today\nReturned","Net Clearance %"]

    def _style_cf(row):
        if row["Office"] == "**TOTAL**":
            return ["background-color:#1F3864;color:white;font-weight:700"] * len(row)
        return [""] * len(row)

    def _style_clearance(val):
        try:
            v = float(str(val).replace("%",""))
            if v >= 75: return "background-color:#70AD47;color:white;font-weight:600"
            if v >= 55: return "background-color:#FFC000;font-weight:600"
            return "background-color:#FF0000;color:white;font-weight:600"
        except: return ""

    # Total row
    num_cols = ["yest_deposit","invoice-count_today","delivery-count_today",
                "deposit-count_today","return-count_today"]
    if has_stat: num_cols.append("rec-count")
    tots3 = disp3[num_cols].sum()
    base_tot = int(tots3.get("rec-count", tots3["invoice-count_today"]))
    tot3_row = {
        "Office": "**TOTAL**",
        "Yest. Deposited\n(Carryforward)": int(tots3["yest_deposit"]),
        "Today\nInvoiced":  int(tots3["invoice-count_today"]),
        "Today\nReceived":  base_tot,
        "CF %\n(of Invoiced)": f"{pct(int(tots3['yest_deposit']), int(tots3['invoice-count_today']))}%",
        "Today\nDelivered":  int(tots3["delivery-count_today"]),
        "New Deposit\n(Tomorrow CF)": int(tots3["deposit-count_today"]),
        "Today\nReturned":   int(tots3["return-count_today"]),
        "Net Clearance %":   f"{pct(int(tots3['delivery-count_today']), base_tot)}%",
    }
    display_df3 = pd.concat([disp3[show3], pd.DataFrame([tot3_row])], ignore_index=True)
    styled3 = (display_df3.style
               .apply(_style_cf, axis=1)
               .map(_style_clearance, subset=["Net Clearance %"]))
    st.dataframe(styled3, use_container_width=True, hide_index=True)

    # Key insights
    st.markdown("#### 📌 Key Insights")
    total_cf    = int(merged["yest_deposit"].sum())
    total_inv   = int(merged["invoice-count_today"].sum())
    total_new_dep = int(merged["deposit-count_today"].sum())
    cf_pct      = pct(total_cf, total_inv)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Yesterday's Carryforward", f"{total_cf:,}")
    col2.metric("Carryforward as % of Today's Load", f"{cf_pct}%")
    col3.metric("Today's New Deposits (Tomorrow CF)", f"{total_new_dep:,}")
    delta_cf = total_new_dep - total_cf
    col4.metric("Change in Carryforward", f"{abs(delta_cf):,}",
                delta=f"{'↑ Increasing' if delta_cf > 0 else '↓ Decreasing'}",
                delta_color="inverse")

# ── Download Excel ────────────────────────────────────────────────────────────
st.markdown("---")
excel_bytes = build_excel(stat_df, prod_today, prod_yest, report_date, merged)
st.download_button(
    "⬇️ Download Consolidated Report (Excel)",
    data=excel_bytes,
    file_name=f"Delivery_Report_{report_date.strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
