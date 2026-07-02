import streamlit as st
import glob as _glob

import glob as _glob

def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'></p>
        </div>""", unsafe_allow_html=True)
    
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f3e0 Home")
    
    _aebas = (_glob.glob("pages/AEBAS_Monitoring.py") +
              _glob.glob("pages/*[Aa][Ee][Bb][Aa][Ss]*.py"))
    if _aebas:
        st.sidebar.page_link(_aebas[0].replace("\\", "/"), label="\U0001f91a AEBAS Monitoring")
    
    _bulk = (_glob.glob("pages/Bulk Analytics.py") +
             _glob.glob("pages/*[Bb]ulk*.py"))
    if _bulk:
        st.sidebar.page_link(_bulk[0].replace("\\", "/"), label="\U0001f4ca Bulk Customer Analytics")
    
    _del = (_glob.glob("pages/Delivery Productivity.py") +
            _glob.glob("pages/*[Dd]elivery*.py"))
    if _del:
        st.sidebar.page_link(_del[0].replace("\\", "/"), label="\U0001f4e6 Delivery Productivity")
    
    _dig = (_glob.glob("pages/1_Digital_Transactions.py") +
            _glob.glob("pages/*[Dd]igital*.py"))
    if _dig:
        st.sidebar.page_link(_dig[0].replace("\\", "/"), label="\U0001f4bb Digital Transactions")
    
    _posb = (_glob.glob("pages/POSB Daily Report.py") +
             _glob.glob("pages/*[Pp][Oo][Ss][Bb]*.py"))
    if _posb:
        st.sidebar.page_link(_posb[0].replace("\\", "/"), label="\U0001f4ee POSB Daily Report")
    
    _sort = (_glob.glob("pages/Sorting Application.py") +
             _glob.glob("pages/*[Ss]orting*.py"))
    if _sort:
        st.sidebar.page_link(_sort[0].replace("\\", "/"),label="📮 Sorting Assistance")
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)
import pandas as pd
import io, os, re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from PIL import Image

st.set_page_config(page_title="Digital Transactions", layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
[data-testid="stSidebarNav"] { display: none !important; }
[data-testid="collapsedControl"] {
    display: flex !important; visibility: visible !important; opacity: 1 !important;
    position: fixed !important; top: 50% !important; left: 0px !important;
    transform: translateY(-50%) !important; z-index: 999999 !important;
    background-color: #2f3343 !important; border-radius: 0 8px 8px 0 !important;
    padding: 12px 7px !important; box-shadow: 3px 0 8px rgba(0,0,0,0.35) !important;
    cursor: pointer !important;
}
[data-testid="collapsedControl"] button { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="collapsedControl"] svg { fill: white !important; color: white !important; }
</style>""", unsafe_allow_html=True)


if "authenticated" not in st.session_state or not st.session_state.authenticated:
    st.warning("Please login from the main page.")
    st.stop()

st.markdown("""
<style>
.block-container{padding-top:1rem !important;padding-bottom:0 !important;}
header{visibility:hidden;height:0px !important;}
.rpt-wrap{overflow-x:auto;}
.rpt{border-collapse:collapse;font-size:12.5px;width:100%;}
.rpt caption{font-size:15px;font-weight:700;color:#c00;background:#FFD700;
             padding:8px;text-align:center;caption-side:top;}
.rpt th{background:#2f3343;color:#fff;border:1px solid #555;
        padding:5px 7px;text-align:center;white-space:nowrap;}
.rpt th.grp{background:#1a1f2e;}
.rpt td{border:1px solid #bbb;padding:4px 7px;text-align:center;white-space:nowrap;}
.rpt td.lft{text-align:left !important;}
.rpt tr.rtot td{background:#b8d4f0 !important;font-weight:700;}
.rpt tr.gtot td{background:#FFD700 !important;font-weight:700;}

/* ── Sidebar collapse fix: re-open arrow always visible ── */
[data-testid="collapsedControl"] {
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
    position: fixed !important;
    top: 50% !important;
    left: 0px !important;
    transform: translateY(-50%) !important;
    z-index: 999999 !important;
    background-color: #2f3343 !important;
    border-radius: 0 8px 8px 0 !important;
    padding: 12px 7px !important;
    box-shadow: 3px 0 8px rgba(0,0,0,0.35) !important;
    cursor: pointer !important;
}
[data-testid="collapsedControl"] button {
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
}
[data-testid="collapsedControl"] svg {
    fill: white !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# ---- Logo / Header ----
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None
hl, hc, hr = st.columns([1,8,1])
with hl:
    if logo: st.image(logo, width=90)
with hc:
    st.markdown("""
    <h1 style='font-size:24px;margin-bottom:2px;color:#2f3343;font-weight:700;padding-top:4px;'>
    Digital Transaction Status Report</h1>
    <p style='font-size:13px;color:#555;margin:0;'>
    Headquarter Region – Telangana Postal Circle</p>
    """, unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 8px;border-color:#ddd;'>", unsafe_allow_html=True)

# ========== CONSTANTS ==========
REGION_MAP = {
    30530001:"Hyderabad Region", 30530002:"Hyderabad Region",
    30530003:"Hyderabad Region", 30530004:"Hyderabad Region",
    30530005:"Hyderabad Region", 30530006:"Hyderabad Region",
    30530007:"Hyderabad Region", 30530008:"Hyderabad Region",
    30530009:"Hyderabad Region", 30530010:"Hyderabad Region",
    30530011:"Hyderabad Region", 30600001:"Hyderabad Region",
    30530012:"Headquarters Region", 30530013:"Headquarters Region",
    30530014:"Headquarters Region", 30530015:"Headquarters Region",
    30530016:"Headquarters Region", 30530017:"Headquarters Region",
    30600002:"Headquarters Region",
    # Region-level summary IDs (pre-aggregated reports)
    30610001:"Headquarters Region",
    30610002:"Hyderabad Region",
}
REGION_ORDER = ["Headquarters Region", "Hyderabad Region", "Other"]

# IDs that represent region-level totals (not individual divisions)
REGION_SUMMARY_IDS = {30610001, 30610002}

# ========== HELPERS ==========
def strip_div(name):
    return re.sub(r'\s*Division\s*$', '', str(name), flags=re.IGNORECASE).strip()

def indian_num(val):
    try: v = int(round(float(val)))
    except: return str(val)
    if v < 0: return '-' + indian_num(-v)
    s = str(abs(v))
    if len(s) <= 3: return s
    result = s[-3:]; s = s[:-3]
    while s: result = s[-2:] + ',' + result; s = s[:-2]
    return result

def indian_amt(val):
    try: v = float(val)
    except: return str(val)
    return '₹' + indian_num(int(round(v)))

def n(row, col):
    v = pd.to_numeric(row.get(col, 0), errors='coerce')
    return 0 if pd.isna(v) else v

def clr_dig(pct, use_color=True):
    if not use_color: return "#ffffff"
    p = float(pct)
    if p >= 60: return "#90EE90"
    if p >= 40: return "#FFFACD"
    return "#FF9999"

def clr_cod(pct, use_color=True):
    if not use_color: return "#ffffff"
    p = float(pct)
    if p >= 50: return "#90EE90"
    if p >= 30: return "#FFFACD"
    return "#FF9999"

def th(label, cs=1, rs=1, cls=""):
    a = f' colspan="{cs}"' if cs > 1 else ''
    b = f' rowspan="{rs}"' if rs > 1 else ''
    c = f' class="{cls}"' if cls else ''
    return f'<th{a}{b}{c}>{label}</th>'

def gv(d, k):
    """Safe getter — avoids pd.Series.name attribute collision"""
    if isinstance(d, dict):
        return d.get(k, 0)
    try:
        return d[k]   # works for both Series and dict-like
    except (KeyError, TypeError):
        return 0

def rsum(rdf):
    skip = {"oid","pct"}  # skip these even if numeric
    num_cols = [c for c in rdf.columns 
                if pd.api.types.is_numeric_dtype(rdf[c]) and c not in skip]
    s = {k: float(rdf[k].sum()) for k in num_cols}
    tc = s.get("tot_c", 0); dc = s.get("dig_c", 0)
    s["pct"] = round(dc / tc * 100, 2) if tc > 0 else 0.0
    return s

def region_rows(df, col_region, col_sort):
    for region in REGION_ORDER:
        sub = df[df[col_region] == region]
        if not sub.empty:
            yield region, sub.sort_values(col_sort, ascending=False)

# ========== PROCESS BOOKING ==========
def process_booking(df_raw):
    rows = []
    for _, r in df_raw.iterrows():
        oid = r.get("Office ID")
        if pd.isna(oid): continue
        oid_i = int(float(str(oid)))
        region = REGION_MAP.get(oid_i, "Other")
        # Always use Office Name string — guard against numeric leakage
        raw_name = str(r.get("Office Name", "")).strip()
        try:
            float(raw_name); raw_name = f"Office {oid_i}"
        except: pass
        name = strip_div(raw_name)

        cash_c = n(r,"Cash (Cnt)");               cash_a = n(r,"Cash (Amt)")
        dqr_c  = n(r,"DQR Scan (Cnt)");           dqr_a  = n(r,"DQR Scan (Amt)")
        pc_c   = n(r,"SBIPOS-CARD (Cnt)");        pc_a   = n(r,"SBIPOS-CARD (Amt)")
        pb_c   = n(r,"SBIPOS BHARATQR (Cnt)");    pb_a   = n(r,"SBIPOS BHARATQR (Amt)")
        eb_c   = n(r,"SBIEPAY BHARATQR (Cnt)");   eb_a   = n(r,"SBIEPAY BHARATQR (Amt)")
        eu_c   = n(r,"SBIEPAY UPI (Cnt)");        eu_a   = n(r,"SBIEPAY UPI (Amt)")
        ec_c   = n(r,"SBIEPAY Credit Card (Cnt)");ec_a   = n(r,"SBIEPAY Credit Card (Amt)")
        ed_c   = n(r,"SBIEPAY Debit Card (Cnt)"); ed_a   = n(r,"SBIEPAY Debit Card (Amt)")

        pos_c = pc_c+pb_c;            pos_a = pc_a+pb_a
        pay_c = eb_c+eu_c+ec_c+ed_c;  pay_a = eb_a+eu_a+ec_a+ed_a
        dig_c = dqr_c+pos_c+pay_c;    dig_a = dqr_a+pos_a+pay_a
        tot_c = cash_c+dig_c;          tot_a = cash_a+dig_a
        pct   = round(dig_c/tot_c*100,2) if tot_c > 0 else 0.0

        rows.append(dict(
            oid=oid_i, region=region, name=name,
            cash_c=cash_c, cash_a=cash_a,
            dqr_c=dqr_c,   dqr_a=dqr_a,
            pc_c=pc_c, pc_a=pc_a, pb_c=pb_c, pb_a=pb_a,
            pos_c=pos_c, pos_a=pos_a,
            eb_c=eb_c, eb_a=eb_a, eu_c=eu_c, eu_a=eu_a,
            ec_c=ec_c, ec_a=ec_a, ed_c=ed_c, ed_a=ed_a,
            pay_c=pay_c, pay_a=pay_a,
            dig_c=dig_c, dig_a=dig_a,
            tot_c=tot_c, tot_a=tot_a, pct=pct,
        ))
    return pd.DataFrame(rows)

# ========== PROCESS COD ==========
def process_cod(df_raw):
    df = df_raw.copy()
    cols = [c.lower().strip() for c in df.columns]

    # Detect office-id column (handle variations)
    id_col = next((c for c in df.columns if c.lower().strip() in
                   ('office-id','office_id','office id')), None)
    # Accept "Division Name", "Division_Name", "Office Name" etc.
    _name_kws = ('office-name','office_name','office name',
                 'division name','division_name','div name','division')
    name_col = next((c for c in df.columns if c.lower().strip() in _name_kws), None)
    dig_col  = next((c for c in df.columns if 'digital' in c.lower() and 'count' in c.lower()), None)
    cash_col = next((c for c in df.columns if 'cash' in c.lower() and 'count' in c.lower()), None)

    if not id_col:
        raise ValueError(
            f"Could not find Office ID column in uploaded file. "
            f"Found columns: {list(df.columns[:6])}. "
            f"Please make sure you selected the correct report type (COD Digital Transactions)."
        )
    if not dig_col or not cash_col:
        raise ValueError(
            f"Could not find digital/cash count columns. "
            f"Found: {list(df.columns)}. "
            f"Please ensure you uploaded the COD report (not the Booking Payment report)."
        )

    df['oid_i']  = df[id_col].apply(lambda x: int(float(str(x))) if pd.notna(x) else None)
    df = df.dropna(subset=['oid_i'])
    df['oid_i']  = df['oid_i'].astype(int)
    df['region'] = df['oid_i'].map(REGION_MAP).fillna("Other")
    df['name']   = df[name_col].apply(lambda x: strip_div(str(x))) if name_col else df['oid_i'].astype(str)
    agg = df.groupby(['oid_i','region','name'], as_index=False).agg(
        digital=(dig_col,'sum'),
        cash=(cash_col,'sum')
    )
    agg['total_cod'] = agg['digital'] + agg['cash']
    agg['pct'] = agg.apply(
        lambda r: round(r['digital']/r['total_cod']*100, 2) if r['total_cod'] > 0 else 0.0,
        axis=1)
    # Flag if this is a region-summary file (all IDs are region-level, not division-level)
    agg['is_region_summary'] = agg['oid_i'].isin(REGION_SUMMARY_IDS)
    return agg

# ========== HTML TABLE — BOOKING ==========
def booking_html(df, view, show_region, date_str, use_color, total_label, full_df=None):
    tbl_id = {"Count":"tbl-cnt","Amount":"tbl-amt","Combined":"tbl-comb"}[view]

    if view == "Count":
        h1  = th("Sl.",rs=2)+th("Division",rs=2)
        h1 += th("Cash",rs=2)+th("Digital QR (APT)",rs=2)
        h1 += th("SBI POS Transactions",cs=3,cls="grp")
        h1 += th("SBI ePAY Transactions",cs=5,cls="grp")
        h1 += th("Non-Digital",rs=2)+th("Digital",rs=2)+th("Total",rs=2)
        h1 += th("% of Digital Trnx",rs=2)
        h2  = th("Card")+th("Bharat QR")+th("Total")
        h2 += th("Bharat QR")+th("UPI")+th("Credit Card")+th("Debit Card")+th("Total")
    elif view == "Amount":
        h1  = th("Sl.",rs=2)+th("Division",rs=2)
        h1 += th("Cash (₹)",rs=2)+th("Digital QR (APT) (₹)",rs=2)
        h1 += th("SBI POS Transactions (₹)",cs=3,cls="grp")
        h1 += th("SBI ePAY Transactions (₹)",cs=5,cls="grp")
        h1 += th("Non-Digital (₹)",rs=2)+th("Digital (₹)",rs=2)+th("Total (₹)",rs=2)
        h1 += th("% of Digital Trnx",rs=2)
        h2  = th("Card")+th("Bharat QR")+th("Total")
        h2 += th("Bharat QR")+th("UPI")+th("Credit Card")+th("Debit Card")+th("Total")
    else:
        h1  = th("Sl.",rs=2)+th("Division",rs=2)
        h1 += th("Cash",cs=2,cls="grp")+th("Digital QR (APT)",cs=2,cls="grp")
        h1 += th("SBI POS Transactions",cs=6,cls="grp")
        h1 += th("SBI ePAY Transactions",cs=10,cls="grp")
        h1 += th("Non-Digital",cs=2,cls="grp")+th("Digital",cs=2,cls="grp")
        h1 += th("Total",cs=2,cls="grp")+th("% of Digital Trnx",cs=2,cls="grp")
        h2  = th("Cnt")+th("₹")+th("Cnt")+th("₹")
        h2 += th("Card(Cnt)")+th("Card(₹)")+th("BQR(Cnt)")+th("BQR(₹)")+th("Total(Cnt)")+th("Total(₹)")
        h2 += th("BQR(Cnt)")+th("BQR(₹)")+th("UPI(Cnt)")+th("UPI(₹)")+th("CC(Cnt)")+th("CC(₹)")+th("DC(Cnt)")+th("DC(₹)")+th("Total(Cnt)")+th("Total(₹)")
        h2 += th("Cnt")+th("₹")+th("Cnt")+th("₹")+th("Cnt")+th("₹")+th("Cnt")+th("₹")

    cap = {"Count":"Count/Transactions","Amount":"Amount (₹)","Combined":"Combined"}[view]
    html = f'<div class="rpt-wrap" id="{tbl_id}"><table class="rpt">'
    html += f'<caption>Digital Transaction Status – {cap} — {date_str}</caption>'
    html += f'<thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'

    def cells(d):
        pv = float(gv(d,"pct"))
        t = ""
        if view == "Count":
            t += f'<td>{indian_num(gv(d,"cash_c"))}</td>'
            t += f'<td>{indian_num(gv(d,"dqr_c"))}</td>'
            t += f'<td>{indian_num(gv(d,"pc_c"))}</td><td>{indian_num(gv(d,"pb_c"))}</td>'
            t += f'<td><b>{indian_num(gv(d,"pos_c"))}</b></td>'
            t += f'<td>{indian_num(gv(d,"eb_c"))}</td><td>{indian_num(gv(d,"eu_c"))}</td>'
            t += f'<td>{indian_num(gv(d,"ec_c"))}</td><td>{indian_num(gv(d,"ed_c"))}</td>'
            t += f'<td><b>{indian_num(gv(d,"pay_c"))}</b></td>'
            t += f'<td>{indian_num(gv(d,"cash_c"))}</td>'
            t += f'<td>{indian_num(gv(d,"dig_c"))}</td>'
            t += f'<td>{indian_num(gv(d,"tot_c"))}</td>'
            t += f'<td><b>{pv:.2f}%</b></td>'
        elif view == "Amount":
            t += f'<td>{indian_amt(gv(d,"cash_a"))}</td>'
            t += f'<td>{indian_amt(gv(d,"dqr_a"))}</td>'
            t += f'<td>{indian_amt(gv(d,"pc_a"))}</td><td>{indian_amt(gv(d,"pb_a"))}</td>'
            t += f'<td><b>{indian_amt(gv(d,"pos_a"))}</b></td>'
            t += f'<td>{indian_amt(gv(d,"eb_a"))}</td><td>{indian_amt(gv(d,"eu_a"))}</td>'
            t += f'<td>{indian_amt(gv(d,"ec_a"))}</td><td>{indian_amt(gv(d,"ed_a"))}</td>'
            t += f'<td><b>{indian_amt(gv(d,"pay_a"))}</b></td>'
            t += f'<td>{indian_amt(gv(d,"cash_a"))}</td>'
            t += f'<td>{indian_amt(gv(d,"dig_a"))}</td>'
            t += f'<td>{indian_amt(gv(d,"tot_a"))}</td>'
            tc=float(gv(d,"tot_c")); dc=float(gv(d,"dig_c"))
            t += f'<td><b>{round(dc/tc*100,2) if tc>0 else 0:.2f}%</b></td>'
        else:
            t += f'<td>{indian_num(gv(d,"cash_c"))}</td><td>{indian_amt(gv(d,"cash_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"dqr_c"))}</td><td>{indian_amt(gv(d,"dqr_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"pc_c"))}</td><td>{indian_amt(gv(d,"pc_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"pb_c"))}</td><td>{indian_amt(gv(d,"pb_a"))}</td>'
            t += f'<td><b>{indian_num(gv(d,"pos_c"))}</b></td><td><b>{indian_amt(gv(d,"pos_a"))}</b></td>'
            t += f'<td>{indian_num(gv(d,"eb_c"))}</td><td>{indian_amt(gv(d,"eb_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"eu_c"))}</td><td>{indian_amt(gv(d,"eu_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"ec_c"))}</td><td>{indian_amt(gv(d,"ec_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"ed_c"))}</td><td>{indian_amt(gv(d,"ed_a"))}</td>'
            t += f'<td><b>{indian_num(gv(d,"pay_c"))}</b></td><td><b>{indian_amt(gv(d,"pay_a"))}</b></td>'
            t += f'<td>{indian_num(gv(d,"cash_c"))}</td><td>{indian_amt(gv(d,"cash_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"dig_c"))}</td><td>{indian_amt(gv(d,"dig_a"))}</td>'
            t += f'<td>{indian_num(gv(d,"tot_c"))}</td><td>{indian_amt(gv(d,"tot_a"))}</td>'
            tc=float(gv(d,"tot_c")); dc=float(gv(d,"dig_c"))
            ta=float(gv(d,"tot_a")); da=float(gv(d,"dig_a"))
            pc=round(dc/tc*100,2) if tc>0 else 0.0
            pa=round(da/ta*100,2) if ta>0 else 0.0
            t += f'<td><b>{pc:.2f}%</b></td><td><b>{pa:.2f}%</b></td>'
        return t

    sl = 1
    # Only show division rows for HQR (df is already filtered to HQR only or full)
    # Subtotals are handled entirely in the footer section below
    for region, rdf in region_rows(df, "region", "pct"):
        for _, row in rdf.iterrows():
            bg = clr_dig(row["pct"], use_color)
            pre = f'<td>{sl}</td>'
            pre += f'<td class="lft">{row["name"]}</td>'
            html += f'<tr style="background:{bg};">{pre}{cells(row)}</tr>\n'
            sl += 1

    if full_df is not None:
        # Multi-region: blue HQR subtotal + blue HYD subtotal + gold Circle total
        hqr_s = rsum(df[df["region"]=="Headquarters Region"]); hqr_s["name"]="Headquarters Region"
        html += f'<tr class="rtot"><td></td><td class="lft"><b>Headquarters Region</b></td>{cells(hqr_s)}</tr>\n'
        hyd_local = full_df[full_df["region"]=="Hyderabad Region"]
        if not hyd_local.empty:
            hyd_s = rsum(hyd_local); hyd_s["name"]="Hyderabad Region"
            html += f'<tr class="rtot"><td></td><td class="lft"><b>Hyderabad Region</b></td>{cells(hyd_s)}</tr>\n'
        gt = rsum(full_df)
    else:
        # Single region: no blue subtotal — only gold total
        gt = rsum(df)

    html += f'<tr class="gtot"><td></td><td class="lft"><b>{total_label}</b></td>{cells(gt)}</tr>\n'
    html += "</tbody></table></div>"
    return html

# ========== HTML TABLE — COD ==========
def cod_html(df, show_region, date_str, use_color, total_label, full_df=None):
    h  = th("Sl.") + th("Division")
    h += th("Total COD Delivered") + th("Digital Trnx.") + th("Cash Trnx.") + th("% Digital")

    html  = '<div class="rpt-wrap" id="tbl-cod"><table class="rpt">'
    html += f'<caption>Status of COD Digital Transactions — {date_str}</caption>'
    html += f'<thead><tr>{h}</tr></thead><tbody>'

    sl = 1
    # Division rows only — subtotals handled in footer below
    for region, rdf in region_rows(df, "region", "pct"):
        for _, row in rdf.iterrows():
            bg = clr_cod(row["pct"], use_color)
            pre = f'<td>{sl}</td>'
            pre += f'<td class="lft">{row["name"]}</td>'
            pre += f'<td>{indian_num(row["total_cod"])}</td>'
            pre += f'<td>{indian_num(row["digital"])}</td>'
            pre += f'<td>{indian_num(row["cash"])}</td>'
            pre += f'<td><b>{row["pct"]:.2f}%</b></td>'
            html += f'<tr style="background:{bg};">{pre}</tr>\n'
            sl += 1

    is_summary = "is_region_summary" in df.columns and df["is_region_summary"].all()

    if full_df is not None and not is_summary:
        # Multi-region division file: blue HQR + blue HYD + gold total
        hqr_c = df[df["region"]=="Headquarters Region"]
        tc_h=hqr_c["total_cod"].sum(); dc_h=hqr_c["digital"].sum(); cc_h=hqr_c["cash"].sum()
        pp_h=round(dc_h/tc_h*100,2) if tc_h>0 else 0.0
        html += (f'<tr class="rtot"><td></td><td class="lft"><b>Headquarters Region</b></td>'
                 f'<td>{indian_num(tc_h)}</td><td>{indian_num(dc_h)}</td>'
                 f'<td>{indian_num(cc_h)}</td><td><b>{pp_h:.2f}%</b></td></tr>\n')
        hyd_c = full_df[full_df["region"]=="Hyderabad Region"]
        if not hyd_c.empty:
            tc_y=hyd_c["total_cod"].sum(); dc_y=hyd_c["digital"].sum(); cc_y=hyd_c["cash"].sum()
            pp_y=round(dc_y/tc_y*100,2) if tc_y>0 else 0.0
            html += (f'<tr class="rtot"><td></td><td class="lft"><b>Hyderabad Region</b></td>'
                     f'<td>{indian_num(tc_y)}</td><td>{indian_num(dc_y)}</td>'
                     f'<td>{indian_num(cc_y)}</td><td><b>{pp_y:.2f}%</b></td></tr>\n')
        tc=full_df["total_cod"].sum(); dc=full_df["digital"].sum(); cc=full_df["cash"].sum()
    else:
        # Single region or region-summary: no blue subtotal, just gold total
        tc=df["total_cod"].sum(); dc=df["digital"].sum(); cc=df["cash"].sum()

    pg=round(dc/tc*100,2) if tc>0 else 0.0
    html += (f'<tr class="gtot"><td></td><td class="lft"><b>{total_label}</b></td>'
             f'<td>{indian_num(tc)}</td><td>{indian_num(dc)}</td>'
             f'<td>{indian_num(cc)}</td><td><b>{pg:.2f}%</b></td></tr>\n')
    html += "</tbody></table></div>"
    return html

# ========== SERVER-SIDE PNG ==========
def df_to_png(df_display, title, col_colors, header_color="#2f3343"):
    """Render a DataFrame as a PNG using matplotlib — tight layout, no wasted space"""
    import matplotlib.pyplot as plt

    nrows, ncols = df_display.shape
    # Scale figure tightly to content
    col_w = 1.3
    row_h = 0.32
    title_h = 0.45
    fig_w = max(ncols * col_w, 8)
    fig_h = (nrows + 1) * row_h + title_h + 0.2  # +1 for header row

    fig = plt.figure(figsize=(fig_w, fig_h))
    # Title axes at top
    ax_title = fig.add_axes([0, 1 - title_h/fig_h, 1, title_h/fig_h])
    ax_title.axis('off')
    ax_title.set_facecolor('#FFD700')
    ax_title.text(0.5, 0.5, title, ha='center', va='center',
                  fontsize=10, fontweight='bold', color='#cc0000',
                  transform=ax_title.transAxes)

    # Table axes below title
    tbl_frac = (nrows + 1) * row_h / fig_h
    ax = fig.add_axes([0, 0, 1, tbl_frac])
    ax.axis('off')

    tbl = ax.table(
        cellText=df_display.values,
        colLabels=df_display.columns,
        loc='center',
        cellLoc='center',
        bbox=[0, 0, 1, 1]
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.auto_set_column_width(col=list(range(ncols)))
    # Scale row heights to be tight
    for (row_idx, col_idx), cell in tbl.get_celld().items():
        cell.set_height(row_h / fig_h)

    for ci in range(ncols):
        cell = tbl[0, ci]
        cell.set_facecolor(header_color)
        cell.set_text_props(color='white', fontweight='bold')
        cell.set_edgecolor('#555555')

    for ri in range(nrows):
        bg = col_colors[ri] if ri < len(col_colors) else "#ffffff"
        for ci in range(ncols):
            cell = tbl[ri+1, ci]
            cell.set_facecolor(bg)
            cell.set_edgecolor('#bbbbbb')

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', pad_inches=0.05)
    buf.seek(0)
    plt.close(fig)
    return buf.getvalue()

def booking_png(df, view, show_region, date_str, use_color, total_label, full_df=None):
    """Build PNG: division rows + HQR subtotal + HYD subtotal (if multi) + Circle/HQR total"""
    rows_data = []; colors = []
    sl = 1

    def s_row(s, label, view):
        r = {"Sl":"","Division":label}
        if view=="Count":
            r.update({"Cash":indian_num(s.get("cash_c",0)),"DQR":indian_num(s.get("dqr_c",0)),
                       "POS Card":indian_num(s.get("pc_c",0)),"POS BQR":indian_num(s.get("pb_c",0)),"POS Tot":indian_num(s.get("pos_c",0)),
                       "ePAY BQR":indian_num(s.get("eb_c",0)),"UPI":indian_num(s.get("eu_c",0)),"CC":indian_num(s.get("ec_c",0)),"DC":indian_num(s.get("ed_c",0)),"ePAY Tot":indian_num(s.get("pay_c",0)),
                       "Non-Dig":indian_num(s.get("cash_c",0)),"Digital":indian_num(s.get("dig_c",0)),"Total":indian_num(s.get("tot_c",0)),"% Digital":f'{s.get("pct",0):.2f}%'})
        elif view=="Amount":
            r.update({"Cash(Rs.)":indian_amt(s.get("cash_a",0)),"DQR(Rs.)":indian_amt(s.get("dqr_a",0)),
                       "POS Card(Rs.)":indian_amt(s.get("pc_a",0)),"POS BQR(Rs.)":indian_amt(s.get("pb_a",0)),"POS Tot(Rs.)":indian_amt(s.get("pos_a",0)),
                       "BQR(Rs.)":indian_amt(s.get("eb_a",0)),"UPI(Rs.)":indian_amt(s.get("eu_a",0)),"CC(Rs.)":indian_amt(s.get("ec_a",0)),"DC(Rs.)":indian_amt(s.get("ed_a",0)),"ePAY Tot(Rs.)":indian_amt(s.get("pay_a",0)),
                       "Non-Dig(Rs.)":indian_amt(s.get("cash_a",0)),"Digital(Rs.)":indian_amt(s.get("dig_a",0)),"Total(Rs.)":indian_amt(s.get("tot_a",0)),"% Digital":f'{s.get("pct",0):.2f}%'})
        else:
            tc2=float(s.get("tot_c",0)); dc2=float(s.get("dig_c",0))
            r.update({"Cash":indian_num(s.get("cash_c",0)),"DQR":indian_num(s.get("dqr_c",0)),
                       "POS Tot":indian_num(s.get("pos_c",0)),"ePAY Tot":indian_num(s.get("pay_c",0)),
                       "Digital":indian_num(s.get("dig_c",0)),"Digital(Rs.)":indian_amt(s.get("dig_a",0)),
                       "Total":indian_num(s.get("tot_c",0)),"Total(Rs.)":indian_amt(s.get("tot_a",0)),
                       "% Digital":f'{round(dc2/tc2*100,2) if tc2>0 else 0:.2f}%'})
        return r

    # Division rows — only HQR (df is already filtered)
    for _, row in df[df["region"]=="Headquarters Region"].sort_values("pct",ascending=False).iterrows():
        r = {"Sl":sl,"Division":row["name"]}
        if view=="Count":
            r.update({"Cash":indian_num(row["cash_c"]),"DQR":indian_num(row["dqr_c"]),
                       "POS Card":indian_num(row["pc_c"]),"POS BQR":indian_num(row["pb_c"]),"POS Tot":indian_num(row["pos_c"]),
                       "ePAY BQR":indian_num(row["eb_c"]),"UPI":indian_num(row["eu_c"]),"CC":indian_num(row["ec_c"]),"DC":indian_num(row["ed_c"]),"ePAY Tot":indian_num(row["pay_c"]),
                       "Non-Dig":indian_num(row["cash_c"]),"Digital":indian_num(row["dig_c"]),"Total":indian_num(row["tot_c"]),"% Digital":f'{row["pct"]:.2f}%'})
        elif view=="Amount":
            r.update({"Cash(Rs.)":indian_amt(row["cash_a"]),"DQR(Rs.)":indian_amt(row["dqr_a"]),
                       "POS Card(Rs.)":indian_amt(row["pc_a"]),"POS BQR(Rs.)":indian_amt(row["pb_a"]),"POS Tot(Rs.)":indian_amt(row["pos_a"]),
                       "BQR(Rs.)":indian_amt(row["eb_a"]),"UPI(Rs.)":indian_amt(row["eu_a"]),"CC(Rs.)":indian_amt(row["ec_a"]),"DC(Rs.)":indian_amt(row["ed_a"]),"ePAY Tot(Rs.)":indian_amt(row["pay_a"]),
                       "Non-Dig(Rs.)":indian_amt(row["cash_a"]),"Digital(Rs.)":indian_amt(row["dig_a"]),"Total(Rs.)":indian_amt(row["tot_a"]),"% Digital":f'{row["pct"]:.2f}%'})
        else:
            r.update({"Cash":indian_num(row["cash_c"]),"DQR":indian_num(row["dqr_c"]),
                       "POS Tot":indian_num(row["pos_c"]),"ePAY Tot":indian_num(row["pay_c"]),
                       "Digital":indian_num(row["dig_c"]),"Digital(Rs.)":indian_amt(row["dig_a"]),
                       "Total":indian_num(row["tot_c"]),"Total(Rs.)":indian_amt(row["tot_a"]),"% Digital":f'{row["pct"]:.2f}%'})
        rows_data.append(r); colors.append(clr_dig(row["pct"],use_color)); sl+=1

    if full_df is not None:
        # Multi-region: blue HQR + blue HYD + gold total
        hqr_s = rsum(df[df["region"]=="Headquarters Region"])
        rows_data.append(s_row(hqr_s,"Headquarters Region",view)); colors.append("#b8d4f0")
        hyd_df2 = full_df[full_df["region"]=="Hyderabad Region"]
        if not hyd_df2.empty:
            hyd_s2 = rsum(hyd_df2)
            rows_data.append(s_row(hyd_s2,"Hyderabad Region",view)); colors.append("#b8d4f0")
        gt2 = rsum(full_df)
        rows_data.append(s_row(gt2,total_label,view)); colors.append("#FFD700")
    else:
        # Single region: no blue subtotal, only gold total
        gt2 = rsum(df)
        rows_data.append(s_row(gt2,total_label,view)); colors.append("#FFD700")

    dfd = pd.DataFrame(rows_data)
    cap = {"Count":"Count/Transactions","Amount":"Amount (Rs.)","Combined":"Combined"}[view]
    return df_to_png(dfd, f"Digital Transaction Status – {cap} — {date_str}", colors)

def cod_png(df, show_region, date_str, use_color, total_label, full_df=None):
    rows_data = []; colors = []; sl = 1

    # Division rows — HQR only (df is already filtered)
    for _, row in df[df["region"]=="Headquarters Region"].sort_values("pct",ascending=False).iterrows():
        rows_data.append({"Sl":sl,"Division":row["name"],
                           "Total COD":indian_num(row["total_cod"]),"Digital":indian_num(row["digital"]),
                           "Cash":indian_num(row["cash"]),"% Digital":f'{row["pct"]:.2f}%'})
        colors.append(clr_cod(row["pct"],use_color)); sl+=1

    is_summary_png = "is_region_summary" in df.columns and df["is_region_summary"].all()

    if full_df is not None and not is_summary_png:
        # Multi-region division file: blue HQR + blue HYD + gold total
        hqr_c = df[df["region"]=="Headquarters Region"]
        tc_h=hqr_c["total_cod"].sum(); dc_h=hqr_c["digital"].sum(); cc_h=hqr_c["cash"].sum()
        pp_h=round(dc_h/tc_h*100,2) if tc_h>0 else 0.0
        rows_data.append({"Sl":"","Division":"Headquarters Region","Total COD":indian_num(tc_h),
                           "Digital":indian_num(dc_h),"Cash":indian_num(cc_h),"% Digital":f'{pp_h:.2f}%'})
        colors.append("#b8d4f0")
        hyd_c = full_df[full_df["region"]=="Hyderabad Region"]
        if not hyd_c.empty:
            tc_y=hyd_c["total_cod"].sum(); dc_y=hyd_c["digital"].sum(); cc_y=hyd_c["cash"].sum()
            pp_y=round(dc_y/tc_y*100,2) if tc_y>0 else 0.0
            rows_data.append({"Sl":"","Division":"Hyderabad Region","Total COD":indian_num(tc_y),
                               "Digital":indian_num(dc_y),"Cash":indian_num(cc_y),"% Digital":f'{pp_y:.2f}%'})
            colors.append("#b8d4f0")
        tc=full_df["total_cod"].sum(); dc=full_df["digital"].sum(); cc=full_df["cash"].sum()
    else:
        # Single region or summary: no blue subtotal
        tc=df["total_cod"].sum(); dc=df["digital"].sum(); cc=df["cash"].sum()

    pg=round(dc/tc*100,2) if tc>0 else 0.0
    rows_data.append({"Sl":"","Division":total_label,"Total COD":indian_num(tc),
                       "Digital":indian_num(dc),"Cash":indian_num(cc),"% Digital":f'{pg:.2f}%'})
    colors.append("#FFD700")

    dfd = pd.DataFrame(rows_data)
    return df_to_png(dfd, f"Status of COD Digital Transactions — {date_str}", colors)

# ========== EXCEL BUILDER ==========
def build_excel(df, view, show_region, date_str, use_color, total_label, full_df=None, mode="booking"):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        def mkfmt(**kw):
            base={"border":1,"align":"center","valign":"vcenter","text_wrap":True}
            base.update(kw); return wb.add_format(base)
        hdr  = mkfmt(bold=True,bg_color="#2f3343",font_color="#FFFFFF")
        grpf = mkfmt(bold=True,bg_color="#1a1f2e",font_color="#FFFFFF")
        cell = mkfmt(); lft=mkfmt(align="left"); boldf=mkfmt(bold=True)
        rtot = mkfmt(bold=True,bg_color="#b8d4f0")
        rlft = mkfmt(bold=True,bg_color="#b8d4f0",align="left")
        gtot = mkfmt(bold=True,bg_color="#FFD700")
        glft = mkfmt(bold=True,bg_color="#FFD700",align="left")
        def cf_dig(pct):
            if not use_color: return cell
            p=float(pct)
            return mkfmt(bg_color="#90EE90" if p>=60 else ("#FFFACD" if p>=40 else "#FF9999"))
        def cf_dig_l(pct):
            if not use_color: return lft
            p=float(pct)
            return mkfmt(bg_color="#90EE90" if p>=60 else ("#FFFACD" if p>=40 else "#FF9999"),align="left")
        def cf_cod(pct):
            if not use_color: return cell
            p=float(pct)
            return mkfmt(bg_color="#90EE90" if p>=50 else ("#FFFACD" if p>=30 else "#FF9999"))
        def cf_cod_l(pct):
            if not use_color: return lft
            p=float(pct)
            return mkfmt(bg_color="#90EE90" if p>=50 else ("#FFFACD" if p>=30 else "#FF9999"),align="left")

        ws_name = {"Count":"Count","Amount":"Amount","Combined":"Combined","COD":"COD"}[view]
        ws = wb.add_worksheet(ws_name); writer.sheets[ws_name] = ws

        # Title row
        title_fmt = mkfmt(bold=True,bg_color="#FFD700",font_color="#cc0000",font_size=12,align="center")
        report_titles = {
            "Count": f"Digital Transaction Status – Count/Transactions — {date_str}",
            "Amount": f"Digital Transaction Status – Amount (Rs.) — {date_str}",
            "Combined": f"Digital Transaction Status – Combined — {date_str}",
            "COD": f"Status of COD Digital Transactions — {date_str}",
        }
        ws.set_row(0, 22)

        if mode == "cod":
            cols = [("Sl.","sl"),("Division","name"),
                    ("Total COD Delivered","total_cod"),
                    ("Digital Trnx.","digital"),
                    ("Cash Trnx.","cash"),("% Digital","pct")]
            if show_region: cols.insert(1,("Region","region"))
            ncols_cod = len(cols)
            ws.merge_range(0,0,0,ncols_cod-1, report_titles.get(view,""), title_fmt)
            ws.set_row(1,28)
            for ci,(lbl,_) in enumerate(cols): ws.write(1,ci,lbl,hdr)
            dr=2; sl=1
            # Division rows only — subtotals in footer below
            for region,rdf in region_rows(df,"region","pct"):
                for _,row in rdf.iterrows():
                    pct_v=row["pct"]
                    for ci,(lbl,key) in enumerate(cols):
                        f=cf_cod(pct_v); fl=cf_cod_l(pct_v)
                        if key=="sl":        ws.write(dr,ci,sl,f)
                        elif key=="region":  ws.write(dr,ci,row["region"],f)
                        elif key=="name":    ws.write(dr,ci,row["name"],fl)
                        elif key=="pct":     ws.write(dr,ci,f"{pct_v:.2f}%",f)
                        else:                ws.write(dr,ci,int(row[key]),f)
                    sl+=1; dr+=1
            is_summary_xl = "is_region_summary" in df.columns and df["is_region_summary"].all()

            if full_df is not None and not is_summary_xl:
                # Multi-region: blue HQR + blue HYD + gold total
                hqr_xl = df[df["region"]=="Headquarters Region"] if "region" in df.columns else df
                tc_hqr=hqr_xl["total_cod"].sum(); dc_hqr=hqr_xl["digital"].sum(); cc_hqr=hqr_xl["cash"].sum()
                ph_hqr=round(dc_hqr/tc_hqr*100,2) if tc_hqr>0 else 0.0
                for ci,(lbl,key) in enumerate(cols):
                    if key=="sl": ws.write(dr,ci,"",rtot)
                    elif key in ("region",): ws.write(dr,ci,"",rtot)
                    elif key=="name": ws.write(dr,ci,"Headquarters Region",rlft)
                    elif key=="total_cod": ws.write(dr,ci,int(tc_hqr),rtot)
                    elif key=="digital": ws.write(dr,ci,int(dc_hqr),rtot)
                    elif key=="cash": ws.write(dr,ci,int(cc_hqr),rtot)
                    elif key=="pct": ws.write(dr,ci,f"{ph_hqr:.2f}%",rtot)
                dr+=1
                hyd_xl = full_df[full_df["region"]=="Hyderabad Region"]
                if not hyd_xl.empty:
                    tc_h=hyd_xl["total_cod"].sum(); dc_h=hyd_xl["digital"].sum(); cc_h=hyd_xl["cash"].sum()
                    ph_h=round(dc_h/tc_h*100,2) if tc_h>0 else 0.0
                    for ci,(lbl,key) in enumerate(cols):
                        if key=="sl": ws.write(dr,ci,"",rtot)
                        elif key in ("region",): ws.write(dr,ci,"",rtot)
                        elif key=="name": ws.write(dr,ci,"Hyderabad Region",rlft)
                        elif key=="total_cod": ws.write(dr,ci,int(tc_h),rtot)
                        elif key=="digital": ws.write(dr,ci,int(dc_h),rtot)
                        elif key=="cash": ws.write(dr,ci,int(cc_h),rtot)
                        elif key=="pct": ws.write(dr,ci,f"{ph_h:.2f}%",rtot)
                    dr+=1
                tc2=full_df["total_cod"].sum(); dc2=full_df["digital"].sum(); cc2=full_df["cash"].sum()
            else:
                # Single region or summary: no blue subtotal
                tc2=df["total_cod"].sum(); dc2=df["digital"].sum(); cc2=df["cash"].sum()
            pg=round(dc2/tc2*100,2) if tc2>0 else 0.0
            for ci,(lbl,key) in enumerate(cols):
                if key=="sl":           ws.write(dr,ci,"",gtot)
                elif key=="region":     ws.write(dr,ci,"",gtot)
                elif key=="name":       ws.write(dr,ci,total_label,glft)
                elif key=="total_cod":  ws.write(dr,ci,int(tc2),gtot)
                elif key=="digital":    ws.write(dr,ci,int(dc2),gtot)
                elif key=="cash":       ws.write(dr,ci,int(cc2),gtot)
                elif key=="pct":        ws.write(dr,ci,f"{pg:.2f}%",gtot)
            max_n = max((len(str(r)) for r in df["name"]), default=20)
            for ci, (lbl, key) in enumerate(cols):
                if key == "name": ws.set_column(ci, ci, min(max_n + 3, 35))
                elif key == "region": ws.set_column(ci, ci, 22)
                elif key == "pct": ws.set_column(ci, ci, 14)
                else: ws.set_column(ci, ci, max(len(str(lbl)) + 2, 12))
        else:
            # Build column spec
            fixed=[]
            fixed.append({"r1":"Sl.","r2":"","k":"sl","t":"c"})
            if show_region: fixed.append({"r1":"Region","r2":"","k":"region","t":"l"})
            fixed.append({"r1":"Division","r2":"","k":"name","t":"l"})

            dc_list=[]
            def ac(r1,r2,k,t,grp=False): dc_list.append({"r1":r1,"r2":r2,"k":k,"t":t,"grp":grp})

            if view=="Count":
                ac("Cash","","cash_c","c"); ac("Digital QR (APT)","","dqr_c","c")
                ac("SBI POS","Card","pc_c","c",True); ac("SBI POS","Bharat QR","pb_c","c",True); ac("SBI POS","Total","pos_c","b",True)
                ac("SBI ePAY","Bharat QR","eb_c","c",True); ac("SBI ePAY","UPI","eu_c","c",True)
                ac("SBI ePAY","Credit Card","ec_c","c",True); ac("SBI ePAY","Debit Card","ed_c","c",True); ac("SBI ePAY","Total","pay_c","b",True)
                ac("Non-Digital","","cash_c","c"); ac("Digital","","dig_c","c"); ac("Total","","tot_c","c"); ac("% of Digital Trnx","","pct","p")
            elif view=="Amount":
                ac("Cash (₹)","","cash_a","a"); ac("Digital QR (₹)","","dqr_a","a")
                ac("SBI POS (₹)","Card","pc_a","a",True); ac("SBI POS (₹)","Bharat QR","pb_a","a",True); ac("SBI POS (₹)","Total","pos_a","b",True)
                ac("SBI ePAY (₹)","Bharat QR","eb_a","a",True); ac("SBI ePAY (₹)","UPI","eu_a","a",True)
                ac("SBI ePAY (₹)","Credit Card","ec_a","a",True); ac("SBI ePAY (₹)","Debit Card","ed_a","a",True); ac("SBI ePAY (₹)","Total","pay_a","b",True)
                ac("Non-Digital (₹)","","cash_a","a"); ac("Digital (₹)","","dig_a","a"); ac("Total (₹)","","tot_a","a"); ac("% of Digital Trnx","","pct","p")
            else:
                ac("Cash","Cnt","cash_c","c",True); ac("Cash","₹","cash_a","a",True)
                ac("Digital QR","Cnt","dqr_c","c",True); ac("Digital QR","₹","dqr_a","a",True)
                ac("SBI POS","Card(Cnt)","pc_c","c",True); ac("SBI POS","Card(₹)","pc_a","a",True)
                ac("SBI POS","BQR(Cnt)","pb_c","c",True); ac("SBI POS","BQR(₹)","pb_a","a",True)
                ac("SBI POS","Total(Cnt)","pos_c","b",True); ac("SBI POS","Total(₹)","pos_a","b",True)
                ac("SBI ePAY","BQR(Cnt)","eb_c","c",True); ac("SBI ePAY","BQR(₹)","eb_a","a",True)
                ac("SBI ePAY","UPI(Cnt)","eu_c","c",True); ac("SBI ePAY","UPI(₹)","eu_a","a",True)
                ac("SBI ePAY","CC(Cnt)","ec_c","c",True); ac("SBI ePAY","CC(₹)","ec_a","a",True)
                ac("SBI ePAY","DC(Cnt)","ed_c","c",True); ac("SBI ePAY","DC(₹)","ed_a","a",True)
                ac("SBI ePAY","Total(Cnt)","pay_c","b",True); ac("SBI ePAY","Total(₹)","pay_a","b",True)
                ac("Non-Digital","Cnt","cash_c","c",True); ac("Non-Digital","₹","cash_a","a",True)
                ac("Digital","Cnt","dig_c","c",True); ac("Digital","₹","dig_a","a",True)
                ac("Total","Cnt","tot_c","c",True); ac("Total","₹","tot_a","a",True)
                ac("% Digital","Cnt","pct","p",True); ac("% Digital","₹","pct_a","p",True)

            all_cols = fixed + dc_list

            # Title row 0 (merged across all cols), headers start at row 1
            total_cols_count = len(fixed) + len(dc_list)
            ws.merge_range(0,0,0,total_cols_count-1, report_titles.get(view,""), title_fmt)
            ws.set_row(0,22); ws.set_row(1,28); ws.set_row(2,22)

            # Write fixed merged headers (now at rows 1+2 instead of 0+1)
            for ci,col in enumerate(fixed):
                ws.merge_range(1,ci,2,ci,col["r1"],hdr)

            base = len(fixed)
            # Write row-2 sub-headers
            for ci,col in enumerate(dc_list):
                abs_ci = base+ci
                if col["r2"]=="":
                    ws.merge_range(1,abs_ci,2,abs_ci,col["r1"],grpf if col.get("grp") else hdr)
                else:
                    ws.write(2,abs_ci,col["r2"],hdr)

            # Merge row-1 group headers
            ci2=0
            while ci2 < len(dc_list):
                col=dc_list[ci2]
                if col["r2"]=="": ci2+=1; continue
                r1=col["r1"]; ss=base+ci2; se=ss; j=ci2+1
                while j<len(dc_list) and dc_list[j]["r1"]==r1 and dc_list[j]["r2"]!="":
                    se=base+j; j+=1
                f2=grpf if col.get("grp") else hdr
                if ss==se: ws.write(1,ss,r1,f2)
                else: ws.merge_range(1,ss,1,se,r1,f2)
                ci2=j

            dr=3; sl=1
            def write_row(ri, d, pct_v, is_tot=False, tot_fmt=None, tot_lft=None):
                for ci3,col in enumerate(all_cols):
                    k=col["k"]; t2=col["t"]
                    f=cf_dig(pct_v) if not is_tot else (tot_fmt or rtot)
                    fl=cf_dig_l(pct_v) if not is_tot else (tot_lft or rlft)
                    if k=="sl":
                        ws.write(ri,ci3,"" if is_tot else sl, f)
                    elif k=="region":
                        ws.write(ri,ci3,"" if is_tot else gv(d,"region"), f)
                    elif k=="name":
                        # FIX: always use string name, never numeric
                        val = gv(d,"name")
                        try: float(str(val)); val=f"Office {gv(d,'oid')}"
                        except: pass
                        ws.write(ri,ci3,str(val),fl)
                    elif k=="pct":
                        tc3=float(gv(d,"tot_c")); dc3=float(gv(d,"dig_c"))
                        ws.write(ri,ci3,f"{round(dc3/tc3*100,2) if tc3>0 else 0:.2f}%",f)
                    elif k=="pct_a":
                        ta3=float(gv(d,"tot_a")); da3=float(gv(d,"dig_a"))
                        ws.write(ri,ci3,f"{round(da3/ta3*100,2) if ta3>0 else 0:.2f}%",f)
                    else:
                        val=gv(d,k)
                        ws.write(ri,ci3,int(val) if t2 in ("c","b") else round(float(val),2),
                                 boldf if t2=="b" and not is_tot else f)

            # Division rows only — subtotals in footer below
            for region,rdf in region_rows(df,"region","pct"):
                for _,row in rdf.iterrows():
                    write_row(dr, row, row["pct"]); sl+=1; dr+=1

            if full_df is not None:
                # Multi-region: blue HQR + blue HYD + gold total
                hqr_xl2 = df[df["region"]=="Headquarters Region"] if "region" in df.columns else df
                hqr_s_xl = rsum(hqr_xl2); hqr_s_xl["name"]="Headquarters Region"; hqr_s_xl["region"]=""; hqr_s_xl["oid"]=0
                write_row(dr, hqr_s_xl, 0, is_tot=True, tot_fmt=rtot, tot_lft=rlft); dr+=1
                hyd_xl2 = full_df[full_df["region"]=="Hyderabad Region"]
                if not hyd_xl2.empty:
                    hyd_s_xl = rsum(hyd_xl2); hyd_s_xl["name"]="Hyderabad Region"; hyd_s_xl["region"]=""; hyd_s_xl["oid"]=0
                    write_row(dr, hyd_s_xl, 0, is_tot=True, tot_fmt=rtot, tot_lft=rlft); dr+=1
                gt_src_xl = full_df
            else:
                # Single region: no blue subtotal, only gold total
                gt_src_xl = df
            gt=rsum(gt_src_xl); gt["name"]=total_label; gt["region"]=""; gt["oid"]=0
            write_row(dr, gt, 0, is_tot=True, tot_fmt=gtot, tot_lft=glft)

            max_n = max((len(str(r)) for r in df["name"]), default=20)
            for ci3, col in enumerate(all_cols):
                k = col["k"]
                r1_len = len(str(col.get("r1","")))
                r2_len = len(str(col.get("r2","")))
                if k == "name":
                    ws.set_column(ci3, ci3, min(max_n + 3, 35))
                elif k == "region":
                    ws.set_column(ci3, ci3, 22)
                elif k in ("pct","pct_a"):
                    ws.set_column(ci3, ci3, 18)
                else:
                    ws.set_column(ci3, ci3, max(r1_len + 2, r2_len + 2, 10))

    out.seek(0); return out.getvalue()

# ========== SIDEBAR ==========
_render_nav()
st.sidebar.header("Upload Reports")
report_type = st.sidebar.radio("Report Type",
    ["Digital Transactions","COD Digital Transactions"], index=0)
uploaded    = st.sidebar.file_uploader("Upload CSV Report", type=["csv"])
report_date = st.sidebar.date_input("Report Date")
use_color   = st.sidebar.checkbox("Colour Coding", value=True,
    help="Uncheck for black & white output")
show_hyd_detail = st.sidebar.checkbox(
    "Show Hyderabad Region division detail",
    value=True,
    help="Uncheck to show only HQR division rows + region subtotals + circle total"
)

# ========== COMPACT VIEW BUILDER ==========
def apply_hyd_filter(df, show_hyd_detail):
    """
    Returns (df_display, full_df_or_None)
    df_display = HQR rows only when show_hyd_detail=False AND multi-region
    full_df    = original full df (for computing HYD subtotal + Circle total)
                 None when showing all detail or single region
    """
    regions_in_data = [r for r in REGION_ORDER if r in df["region"].unique()]
    is_multi = len(regions_in_data) > 1

    if not is_multi:
        # Single region — show all, no HYD subtotal needed
        return df, None

    if show_hyd_detail:
        # Multi-region, full detail — pass full_df so footers can compute correct totals
        return df, df

    # Multi-region, hide HYD detail — show only HQR rows
    hqr_df = df[df["region"] == "Headquarters Region"].copy()
    return hqr_df, df

# ========== MAIN ==========
if uploaded:
    df_raw   = pd.read_csv(uploaded)
    date_str = report_date.strftime("%d.%m.%Y")
    date_fn  = report_date.strftime("%d_%m_%Y")

    # Date warning — shown always so user can verify
    from datetime import date as _date, datetime as _datetime
    today = _date.today()
    # Normalise report_date in case Streamlit returns datetime instead of date
    rdate = report_date if isinstance(report_date, _date) and not isinstance(report_date, _datetime) else report_date.date()
    if rdate != today:
        st.warning(
            f"⚠️ **Report date set to {date_str}** — today is **{today.strftime('%d.%m.%Y')}**. "
            f"Please confirm this is correct before downloading."
        )
    else:
        st.info(f"📅 Report date: **{date_str}**")

    if report_type == "Digital Transactions":
        try:
            df = process_booking(df_raw)
        except Exception as e:
            st.error(f"❌ Could not read Digital Transactions file: {e}")
            st.stop()
        if df.empty: st.warning("No valid data found. Check that the Booking Payment Report is uploaded."); st.stop()
        regions     = df["region"].unique().tolist()
        multi_region = len([r for r in REGION_ORDER if r in regions]) > 1
        show_region = False  # Region column always hidden — grouping shown via subtotals
        total_label = "Telangana Circle" if multi_region else "Headquarters Region"

        if use_color:
            st.markdown("""<div style='display:flex;gap:14px;margin-bottom:8px;font-size:13px;'>
            <span style='background:#90EE90;padding:3px 10px;border-radius:4px;'>≥ 60% Digital</span>
            <span style='background:#FFFACD;padding:3px 10px;border-radius:4px;'>40–59% Digital</span>
            <span style='background:#FF9999;padding:3px 10px;border-radius:4px;'>< 40% Digital</span>
            </div>""", unsafe_allow_html=True)

        tab1, tab2, tab3 = st.tabs(["📊 Count / Transactions","💰 Amount (₹)","📋 Combined"])

        df_display, full_df_ref = apply_hyd_filter(df, show_hyd_detail)

        for tab, view, fn_sfx in [
            (tab1,"Count",f"Digital_Transactions_Count_{date_fn}"),
            (tab2,"Amount",f"Digital_Transactions_Amount_{date_fn}"),
            (tab3,"Combined",f"Digital_Transactions_{date_fn}"),
        ]:
            with tab:
                st.markdown(booking_html(df_display,view,show_region,date_str,use_color,total_label,full_df_ref),
                            unsafe_allow_html=True)
                png_bytes = booking_png(df_display,view,show_region,date_str,use_color,total_label,full_df_ref)
                st.download_button(f"📷 Download as Image ({view})", png_bytes,
                    file_name=f"{fn_sfx}.png", mime="image/png")
                xl = build_excel(df_display,view,show_region,date_str,use_color,total_label,full_df=full_df_ref)
                st.download_button(f"⬇ Download Excel ({view})", xl,
                    file_name=f"{fn_sfx}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:  # COD
        try:
            df = process_cod(df_raw)
        except ValueError as e:
            st.error(f"❌ {e}")
            st.stop()
        if df.empty: st.warning("No valid data found. Check that the correct COD report is uploaded."); st.stop()
        regions     = df["region"].unique().tolist()
        multi_region = len([r for r in REGION_ORDER if r in regions]) > 1
        show_region = False  # Region column always hidden — grouping shown via subtotals
        total_label = "Telangana Circle" if multi_region else "Headquarters Region"

        if use_color:
            st.markdown("""<div style='display:flex;gap:14px;margin-bottom:8px;font-size:13px;'>
            <span style='background:#90EE90;padding:3px 10px;border-radius:4px;'>≥ 50% Digital</span>
            <span style='background:#FFFACD;padding:3px 10px;border-radius:4px;'>30–49% Digital</span>
            <span style='background:#FF9999;padding:3px 10px;border-radius:4px;'>< 30% Digital</span>
            </div>""", unsafe_allow_html=True)

        df_display_cod, full_df_cod = apply_hyd_filter(df, show_hyd_detail)
        st.markdown(cod_html(df_display_cod,show_region,date_str,use_color,total_label,full_df_cod),
                    unsafe_allow_html=True)
        fn = f"COD_Digital_Transactions_{date_fn}"
        png_bytes = cod_png(df_display_cod,show_region,date_str,use_color,total_label,full_df_cod)
        st.download_button("📷 Download as Image", png_bytes,
            file_name=f"{fn}.png", mime="image/png")
        xl = build_excel(df_display_cod,"COD",show_region,date_str,use_color,total_label,full_df=full_df_cod,mode="cod")
        st.download_button("⬇ Download Excel (COD)", xl,
            file_name=f"{fn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Select report type and upload a CSV from the sidebar to begin.")
