import streamlit as st
import pandas as pd
import io, os, re
from PIL import Image

st.set_page_config(
    page_title="Digital Transactions",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
}

# ========== HELPERS ==========
def strip_division(name):
    """Remove trailing ' Division' (case-insensitive)"""
    return re.sub(r'\s*Division\s*$', '', str(name), flags=re.IGNORECASE).strip()

def indian_num(val):
    """Format number in Indian numbering system: 1,00,00,000"""
    try:
        v = int(round(float(val)))
    except:
        return str(val)
    if v < 0:
        return '-' + indian_num(-v)
    s = str(abs(v))
    if len(s) <= 3:
        return s
    # last 3 digits, then groups of 2
    result = s[-3:]
    s = s[:-3]
    while s:
        result = s[-2:] + ',' + result
        s = s[:-2]
    return result

def indian_amt(val):
    try:
        v = float(val)
    except:
        return str(val)
    # Format with Indian grouping, 0 decimal places
    integer_part = int(round(v))
    return '₹' + indian_num(integer_part)

def n(row, col):
    v = pd.to_numeric(row.get(col, 0), errors='coerce')
    return 0 if pd.isna(v) else v

def clr(pct, use_color=True):
    if not use_color:
        return "#ffffff"
    p = float(pct)
    if p >= 60: return "#90EE90"
    if p >= 40: return "#FFFACD"
    return "#FF9999"

# ========== PROCESS ==========
def process(df_raw):
    rows = []
    for _, r in df_raw.iterrows():
        oid = r.get("Office ID")
        if pd.isna(oid): continue
        oid_i  = int(float(str(oid)))
        region = REGION_MAP.get(oid_i, "Other")
        name   = strip_division(str(r.get("Office Name","")))

        cash_c = n(r,"Cash (Cnt)");              cash_a = n(r,"Cash (Amt)")
        dqr_c  = n(r,"DQR Scan (Cnt)");          dqr_a  = n(r,"DQR Scan (Amt)")
        pc_c   = n(r,"SBIPOS-CARD (Cnt)");       pc_a   = n(r,"SBIPOS-CARD (Amt)")
        pb_c   = n(r,"SBIPOS BHARATQR (Cnt)");   pb_a   = n(r,"SBIPOS BHARATQR (Amt)")
        eb_c   = n(r,"SBIEPAY BHARATQR (Cnt)");  eb_a   = n(r,"SBIEPAY BHARATQR (Amt)")
        eu_c   = n(r,"SBIEPAY UPI (Cnt)");        eu_a   = n(r,"SBIEPAY UPI (Amt)")
        ec_c   = n(r,"SBIEPAY Credit Card (Cnt)");ec_a  = n(r,"SBIEPAY Credit Card (Amt)")
        ed_c   = n(r,"SBIEPAY Debit Card (Cnt)"); ed_a   = n(r,"SBIEPAY Debit Card (Amt)")

        pos_c = pc_c+pb_c;           pos_a = pc_a+pb_a
        pay_c = eb_c+eu_c+ec_c+ed_c; pay_a = eb_a+eu_a+ec_a+ed_a
        dig_c = dqr_c+pos_c+pay_c;   dig_a = dqr_a+pos_a+pay_a
        tot_c = cash_c+dig_c;         tot_a = cash_a+dig_a
        pct   = round(dig_c/tot_c*100,2) if tot_c>0 else 0.0

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

def rsum(rdf):
    s = {k: rdf[k].sum() for k in rdf.columns if pd.api.types.is_numeric_dtype(rdf[k])}
    tc = s.get('tot_c',0); dc = s.get('dig_c',0)
    s['pct'] = round(dc/tc*100,2) if tc>0 else 0.0
    return s

def th(label, cs=1, rs=1, cls=""):
    a = f' colspan="{cs}"' if cs>1 else ''
    b = f' rowspan="{rs}"' if rs>1 else ''
    c = f' class="{cls}"' if cls else ''
    return f'<th{a}{b}{c}>{label}</th>'

def g(d, k):
    return d.get(k,0) if isinstance(d,dict) else getattr(d,k,0)

# ========== COUNT TABLE ==========
def build_count(df, show_region, date_str, use_color):
    h1  = th("Sl.",rs=2) + (th("Region",rs=2) if show_region else "") + th("Division",rs=2)
    h1 += th("Cash",rs=2)
    h1 += th("Digital QR (APT)",rs=2)
    h1 += th("SBI POS Transactions",cs=3,cls="grp")
    h1 += th("SBI ePAY Transactions",cs=5,cls="grp")
    h1 += th("Non-Digital",rs=2) + th("Digital",rs=2) + th("Total",rs=2)
    h1 += th("% of Digital Trnx",rs=2)

    h2  = th("Card") + th("Bharat QR") + th("Total")
    h2 += th("Bharat QR") + th("UPI") + th("Credit Card") + th("Debit Card") + th("Total")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def tds(d, label=None, rcls=""):
        t = ""
        if label is not None:
            span = 2 if show_region else 1
            t += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        t += f'<td>{indian_num(g(d,"cash_c"))}</td>'
        t += f'<td>{indian_num(g(d,"dqr_c"))}</td>'
        t += f'<td>{indian_num(g(d,"pc_c"))}</td>'
        t += f'<td>{indian_num(g(d,"pb_c"))}</td>'
        t += f'<td><b>{indian_num(g(d,"pos_c"))}</b></td>'
        t += f'<td>{indian_num(g(d,"eb_c"))}</td>'
        t += f'<td>{indian_num(g(d,"eu_c"))}</td>'
        t += f'<td>{indian_num(g(d,"ec_c"))}</td>'
        t += f'<td>{indian_num(g(d,"ed_c"))}</td>'
        t += f'<td><b>{indian_num(g(d,"pay_c"))}</b></td>'
        t += f'<td>{indian_num(g(d,"cash_c"))}</td>'
        t += f'<td>{indian_num(g(d,"dig_c"))}</td>'
        t += f'<td>{indian_num(g(d,"tot_c"))}</td>'
        t += f'<td><b>{float(g(d,"pct")):.2f}%</b></td>'
        rc = f' class="{rcls}"' if rcls else ''
        return f'<tr{rc}>{t}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-cnt">
    <table class="rpt">
    <caption>Digital Transaction Status – Count/Transactions — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct",ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"], use_color)
            pre = f'<td>{sl}</td>'
            if show_region: pre += f'<td>{row["region"]}</td>'
            pre += f'<td class="lft">{row["name"]}</td>'
            # data cells without wrapping tr
            inner = tds(row)[4:-6]
            html += f'<tr style="background:{bg};">{pre}{inner}</tr>\n'
            sl += 1
        s = rsum(rdf)
        # region label only (no "TOTAL —")
        html += tds(s, label=region, rcls="rtot")

    gt = rsum(df)
    html += tds(gt, label="Total", rcls="gtot")
    html += "</tbody></table></div>"
    return html

# ========== AMOUNT TABLE ==========
def build_amount(df, show_region, date_str, use_color):
    h1  = th("Sl.",rs=2) + (th("Region",rs=2) if show_region else "") + th("Division",rs=2)
    h1 += th("Cash (₹)",rs=2)
    h1 += th("Digital QR (APT) (₹)",rs=2)
    h1 += th("SBI POS Transactions (₹)",cs=3,cls="grp")
    h1 += th("SBI ePAY Transactions (₹)",cs=5,cls="grp")
    h1 += th("Non-Digital (₹)",rs=2) + th("Digital (₹)",rs=2) + th("Total (₹)",rs=2)
    h1 += th("% of Digital Trnx",rs=2)

    h2  = th("Card") + th("Bharat QR") + th("Total")
    h2 += th("Bharat QR") + th("UPI") + th("Credit Card") + th("Debit Card") + th("Total")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def tds(d, label=None, rcls=""):
        t = ""
        if label is not None:
            span = 2 if show_region else 1
            t += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        t += f'<td>{indian_amt(g(d,"cash_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"dqr_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"pc_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"pb_a"))}</td>'
        t += f'<td><b>{indian_amt(g(d,"pos_a"))}</b></td>'
        t += f'<td>{indian_amt(g(d,"eb_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"eu_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"ec_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"ed_a"))}</td>'
        t += f'<td><b>{indian_amt(g(d,"pay_a"))}</b></td>'
        t += f'<td>{indian_amt(g(d,"cash_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"dig_a"))}</td>'
        t += f'<td>{indian_amt(g(d,"tot_a"))}</td>'
        tc = float(g(d,"tot_c")); dc = float(g(d,"dig_c"))
        pct = round(dc/tc*100,2) if tc>0 else 0.0
        t += f'<td><b>{pct:.2f}%</b></td>'
        rc = f' class="{rcls}"' if rcls else ''
        return f'<tr{rc}>{t}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-amt">
    <table class="rpt">
    <caption>Digital Transaction Status – Amount (₹) — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct",ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"], use_color)
            pre = f'<td>{sl}</td>'
            if show_region: pre += f'<td>{row["region"]}</td>'
            pre += f'<td class="lft">{row["name"]}</td>'
            inner = tds(row)[4:-6]
            html += f'<tr style="background:{bg};">{pre}{inner}</tr>\n'
            sl += 1
        s = rsum(rdf)
        html += tds(s, label=region, rcls="rtot")

    gt = rsum(df)
    html += tds(gt, label="Total", rcls="gtot")
    html += "</tbody></table></div>"
    return html

# ========== COMBINED TABLE ==========
def build_combined(df, show_region, date_str, use_color):
    h1  = th("Sl.",rs=2) + (th("Region",rs=2) if show_region else "") + th("Division",rs=2)
    h1 += th("Cash",cs=2,cls="grp")
    h1 += th("Digital QR (APT)",cs=2,cls="grp")
    h1 += th("SBI POS Transactions",cs=6,cls="grp")
    h1 += th("SBI ePAY Transactions",cs=10,cls="grp")
    h1 += th("Non-Digital",cs=2,cls="grp")
    h1 += th("Digital",cs=2,cls="grp")
    h1 += th("Total",cs=2,cls="grp")
    h1 += th("% of Digital Trnx",cs=2,cls="grp")

    h2  = th("Cnt") + th("₹")
    h2 += th("Cnt") + th("₹")
    h2 += th("Card(Cnt)") + th("Card(₹)") + th("Bharat QR(Cnt)") + th("Bharat QR(₹)") + th("Total(Cnt)") + th("Total(₹)")
    h2 += th("BQR(Cnt)") + th("BQR(₹)") + th("UPI(Cnt)") + th("UPI(₹)") + th("CC(Cnt)") + th("CC(₹)") + th("DC(Cnt)") + th("DC(₹)") + th("Total(Cnt)") + th("Total(₹)")
    h2 += th("Cnt") + th("₹")
    h2 += th("Cnt") + th("₹")
    h2 += th("Cnt") + th("₹")
    h2 += th("Cnt") + th("₹")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def tds(d, label=None, rcls=""):
        t = ""
        if label is not None:
            span = 2 if show_region else 1
            t += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        t += f'<td>{indian_num(g(d,"cash_c"))}</td><td>{indian_amt(g(d,"cash_a"))}</td>'
        t += f'<td>{indian_num(g(d,"dqr_c"))}</td><td>{indian_amt(g(d,"dqr_a"))}</td>'
        t += f'<td>{indian_num(g(d,"pc_c"))}</td><td>{indian_amt(g(d,"pc_a"))}</td>'
        t += f'<td>{indian_num(g(d,"pb_c"))}</td><td>{indian_amt(g(d,"pb_a"))}</td>'
        t += f'<td><b>{indian_num(g(d,"pos_c"))}</b></td><td><b>{indian_amt(g(d,"pos_a"))}</b></td>'
        t += f'<td>{indian_num(g(d,"eb_c"))}</td><td>{indian_amt(g(d,"eb_a"))}</td>'
        t += f'<td>{indian_num(g(d,"eu_c"))}</td><td>{indian_amt(g(d,"eu_a"))}</td>'
        t += f'<td>{indian_num(g(d,"ec_c"))}</td><td>{indian_amt(g(d,"ec_a"))}</td>'
        t += f'<td>{indian_num(g(d,"ed_c"))}</td><td>{indian_amt(g(d,"ed_a"))}</td>'
        t += f'<td><b>{indian_num(g(d,"pay_c"))}</b></td><td><b>{indian_amt(g(d,"pay_a"))}</b></td>'
        t += f'<td>{indian_num(g(d,"cash_c"))}</td><td>{indian_amt(g(d,"cash_a"))}</td>'
        t += f'<td>{indian_num(g(d,"dig_c"))}</td><td>{indian_amt(g(d,"dig_a"))}</td>'
        t += f'<td>{indian_num(g(d,"tot_c"))}</td><td>{indian_amt(g(d,"tot_a"))}</td>'
        tc = float(g(d,"tot_c")); dc = float(g(d,"dig_c"))
        ta = float(g(d,"tot_a")); da = float(g(d,"dig_a"))
        pc = round(dc/tc*100,2) if tc>0 else 0.0
        pa = round(da/ta*100,2) if ta>0 else 0.0
        t += f'<td><b>{pc:.2f}%</b></td><td><b>{pa:.2f}%</b></td>'
        rc = f' class="{rcls}"' if rcls else ''
        return f'<tr{rc}>{t}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-comb">
    <table class="rpt">
    <caption>Digital Transaction Status – Combined — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct",ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"], use_color)
            pre = f'<td>{sl}</td>'
            if show_region: pre += f'<td>{row["region"]}</td>'
            pre += f'<td class="lft">{row["name"]}</td>'
            inner = tds(row)[4:-6]
            html += f'<tr style="background:{bg};">{pre}{inner}</tr>\n'
            sl += 1
        s = rsum(rdf)
        html += tds(s, label=region, rcls="rtot")

    gt = rsum(df)
    html += tds(gt, label="Total", rcls="gtot")
    html += "</tbody></table></div>"
    return html

# ========== EXCEL BUILDER ==========
def build_excel(df, view, show_region, date_str, use_color):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        # ---- formats ----
        hdr_fmt  = wb.add_format({"bold":True,"bg_color":"#2f3343","font_color":"#FFFFFF",
                                   "border":1,"align":"center","valign":"vcenter","text_wrap":True})
        grp_fmt  = wb.add_format({"bold":True,"bg_color":"#1a1f2e","font_color":"#FFFFFF",
                                   "border":1,"align":"center","valign":"vcenter","text_wrap":True})
        cell_fmt = wb.add_format({"border":1,"align":"center","valign":"vcenter"})
        lft_fmt  = wb.add_format({"border":1,"align":"left","valign":"vcenter"})
        bold_fmt = wb.add_format({"border":1,"align":"center","valign":"vcenter","bold":True})
        rtot_fmt = wb.add_format({"border":1,"align":"center","valign":"vcenter",
                                   "bold":True,"bg_color":"#b8d4f0"})
        rtot_lft = wb.add_format({"border":1,"align":"left","valign":"vcenter",
                                   "bold":True,"bg_color":"#b8d4f0"})
        gtot_fmt = wb.add_format({"border":1,"align":"center","valign":"vcenter",
                                   "bold":True,"bg_color":"#FFD700"})
        gtot_lft = wb.add_format({"border":1,"align":"left","valign":"vcenter",
                                   "bold":True,"bg_color":"#FFD700"})

        def color_fmt(pct, base=None):
            if not use_color: return base or cell_fmt
            p = float(pct)
            bg = "#90EE90" if p>=60 else ("#FFFACD" if p>=40 else "#FF9999")
            return wb.add_format({"border":1,"align":"center","valign":"vcenter","bg_color":bg})
        def color_lft(pct):
            if not use_color: return lft_fmt
            p = float(pct)
            bg = "#90EE90" if p>=60 else ("#FFFACD" if p>=40 else "#FF9999")
            return wb.add_format({"border":1,"align":"left","valign":"vcenter","bg_color":bg})

        is_cnt = view in ("Count","Combined")
        is_amt = view in ("Amount","Combined")
        ws_name = {"Count":"Count_Transactions","Amount":"Amount","Combined":"Combined"}[view]
        ws = wb.add_worksheet(ws_name)
        writer.sheets[ws_name] = ws

        region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                        if r in df["region"].unique()]

        # ---- build header & column list ----
        # Each entry: (header_row1_label, header_row2_label, data_key, fmt_type)
        # fmt_type: 'c'=count, 'a'=amount, 'b'=bold, 'p'=percent, 'l'=left
        cols = []  # list of dicts

        fixed = []
        fixed.append({"r1":"Sl.","r2":"","key":"sl","t":"c"})
        if show_region:
            fixed.append({"r1":"Region","r2":"","key":"region","t":"l"})
        fixed.append({"r1":"Division","r2":"","key":"name","t":"l"})

        def add(r1,r2,key,t,grp=False):
            cols.append({"r1":r1,"r2":r2,"key":key,"t":t,"grp":grp})

        if is_cnt and not is_amt:
            add("Cash","",          "cash_c","c")
            add("Digital QR (APT)","","dqr_c","c")
            add("SBI POS","Card",   "pc_c","c",grp=True)
            add("SBI POS","Bharat QR","pb_c","c",grp=True)
            add("SBI POS","Total",  "pos_c","b",grp=True)
            add("SBI ePAY","Bharat QR","eb_c","c",grp=True)
            add("SBI ePAY","UPI",   "eu_c","c",grp=True)
            add("SBI ePAY","Credit Card","ec_c","c",grp=True)
            add("SBI ePAY","Debit Card","ed_c","c",grp=True)
            add("SBI ePAY","Total", "pay_c","b",grp=True)
            add("Non-Digital","",   "cash_c","c")
            add("Digital","",       "dig_c","c")
            add("Total","",         "tot_c","c")
            add("% of Digital Trnx","","pct","p")

        elif is_amt and not is_cnt:
            add("Cash (₹)","",       "cash_a","a")
            add("Digital QR (₹)","", "dqr_a","a")
            add("SBI POS (₹)","Card","pc_a","a",grp=True)
            add("SBI POS (₹)","Bharat QR","pb_a","a",grp=True)
            add("SBI POS (₹)","Total","pos_a","b",grp=True)
            add("SBI ePAY (₹)","Bharat QR","eb_a","a",grp=True)
            add("SBI ePAY (₹)","UPI","eu_a","a",grp=True)
            add("SBI ePAY (₹)","Credit Card","ec_a","a",grp=True)
            add("SBI ePAY (₹)","Debit Card","ed_a","a",grp=True)
            add("SBI ePAY (₹)","Total","pay_a","b",grp=True)
            add("Non-Digital (₹)","","cash_a","a")
            add("Digital (₹)","",    "dig_a","a")
            add("Total (₹)","",      "tot_a","a")
            add("% of Digital Trnx","","pct","p")

        else:  # combined
            add("Cash","Cnt",        "cash_c","c",grp=True)
            add("Cash","₹",          "cash_a","a",grp=True)
            add("Digital QR","Cnt",  "dqr_c","c",grp=True)
            add("Digital QR","₹",    "dqr_a","a",grp=True)
            add("SBI POS","Card(Cnt)","pc_c","c",grp=True)
            add("SBI POS","Card(₹)", "pc_a","a",grp=True)
            add("SBI POS","BQR(Cnt)","pb_c","c",grp=True)
            add("SBI POS","BQR(₹)",  "pb_a","a",grp=True)
            add("SBI POS","Total(Cnt)","pos_c","b",grp=True)
            add("SBI POS","Total(₹)","pos_a","b",grp=True)
            add("SBI ePAY","BQR(Cnt)","eb_c","c",grp=True)
            add("SBI ePAY","BQR(₹)", "eb_a","a",grp=True)
            add("SBI ePAY","UPI(Cnt)","eu_c","c",grp=True)
            add("SBI ePAY","UPI(₹)", "eu_a","a",grp=True)
            add("SBI ePAY","CC(Cnt)", "ec_c","c",grp=True)
            add("SBI ePAY","CC(₹)",   "ec_a","a",grp=True)
            add("SBI ePAY","DC(Cnt)", "ed_c","c",grp=True)
            add("SBI ePAY","DC(₹)",   "ed_a","a",grp=True)
            add("SBI ePAY","Total(Cnt)","pay_c","b",grp=True)
            add("SBI ePAY","Total(₹)","pay_a","b",grp=True)
            add("Non-Digital","Cnt", "cash_c","c",grp=True)
            add("Non-Digital","₹",   "cash_a","a",grp=True)
            add("Digital","Cnt",     "dig_c","c",grp=True)
            add("Digital","₹",       "dig_a","a",grp=True)
            add("Total","Cnt",       "tot_c","c",grp=True)
            add("Total","₹",         "tot_a","a",grp=True)
            add("% Digital","Cnt",   "pct","p",grp=True)
            add("% Digital","₹",     "pct_a","p",grp=True)

        all_cols = fixed + cols
        ncols = len(all_cols)

        # ---- write header rows ----
        # Row 0: group/merged headers
        # Row 1: sub-headers
        ws.set_row(0, 30)
        ws.set_row(1, 25)

        # For fixed columns: merge rows 0 and 1
        for ci, col in enumerate(fixed):
            ws.merge_range(0, ci, 1, ci, col["r1"], hdr_fmt)

        # For data cols: write row 0 group labels, row 1 sub-labels
        base = len(fixed)
        prev_r1 = None
        r1_start = base
        for ci, col in enumerate(cols):
            abs_ci = base + ci
            r1_label = col["r1"]
            r2_label = col["r2"]
            fmt = grp_fmt if col.get("grp") else hdr_fmt

            if r2_label == "":
                # Single header — merge row 0 and 1
                ws.merge_range(0, abs_ci, 1, abs_ci, r1_label, fmt)
            else:
                # Write sub-header in row 1
                ws.write(1, abs_ci, r2_label, hdr_fmt)
                # Group header in row 0 handled after loop via merge

        # Merge row-0 group headers
        ci = 0
        while ci < len(cols):
            col = cols[ci]
            if col["r2"] == "":
                ci += 1
                continue
            # find span of same r1
            r1 = col["r1"]
            span_start = base + ci
            span_end   = span_start
            j = ci + 1
            while j < len(cols) and cols[j]["r1"] == r1 and cols[j]["r2"] != "":
                span_end = base + j
                j += 1
            fmt = grp_fmt if col.get("grp") else hdr_fmt
            if span_start == span_end:
                ws.write(0, span_start, r1, fmt)
            else:
                ws.merge_range(0, span_start, 0, span_end, r1, fmt)
            ci = j

        # ---- write data rows ----
        data_row_start = 2
        dr = data_row_start
        sl = 1

        def write_row(ws, row_i, d, pct_val, fmt_fn, lft_fn, is_total=False):
            for ci2, col in enumerate(all_cols):
                key = col["key"]
                t   = col["t"]
                if key == "sl":
                    ws.write(row_i, ci2, "" if is_total else sl, fmt_fn(pct_val))
                elif key == "region":
                    ws.write(row_i, ci2, g(d,"region") if not is_total else "", lft_fn(pct_val) if not is_total else lft_fn(pct_val))
                elif key == "name":
                    label = g(d,"name") if not is_total else d.get("_label","")
                    ws.write(row_i, ci2, label, lft_fn(pct_val))
                elif key == "pct":
                    v = float(g(d,"pct") if "pct" in d else 0)
                    ws.write(row_i, ci2, f"{v:.2f}%", fmt_fn(pct_val))
                elif key == "pct_a":
                    ta = float(g(d,"tot_a")); da = float(g(d,"dig_a"))
                    pa = round(da/ta*100,2) if ta>0 else 0.0
                    ws.write(row_i, ci2, f"{pa:.2f}%", fmt_fn(pct_val))
                else:
                    val = g(d, key)
                    if t in ("c","b"):
                        ws.write(row_i, ci2, int(val), bold_fmt if t=="b" else fmt_fn(pct_val))
                    else:
                        ws.write(row_i, ci2, round(float(val),2), fmt_fn(pct_val))

        for region in region_order:
            rdf = df[df["region"]==region].sort_values("pct",ascending=False)
            for _, row in rdf.iterrows():
                pct_v = row["pct"]
                write_row(ws, dr, row, pct_v,
                          lambda p: color_fmt(p),
                          lambda p: color_lft(p))
                sl += 1; dr += 1
            # region total row
            s = rsum(rdf)
            s["_label"] = region
            s["name"]   = region
            s["region"] = ""
            for ci2, col in enumerate(all_cols):
                key = col["key"]
                if key == "sl":   ws.write(dr, ci2, "", rtot_fmt)
                elif key in ("region",): ws.write(dr, ci2, "", rtot_fmt)
                elif key == "name": ws.write(dr, ci2, region, rtot_lft)
                elif key == "pct":
                    tc2 = float(s.get("tot_c",0)); dc2 = float(s.get("dig_c",0))
                    ws.write(dr, ci2, f"{round(dc2/tc2*100,2) if tc2>0 else 0:.2f}%", rtot_fmt)
                elif key == "pct_a":
                    ta2 = float(s.get("tot_a",0)); da2 = float(s.get("dig_a",0))
                    ws.write(dr, ci2, f"{round(da2/ta2*100,2) if ta2>0 else 0:.2f}%", rtot_fmt)
                else:
                    val = s.get(key,0)
                    t   = all_cols[ci2]["t"]
                    ws.write(dr, ci2, int(val) if t in ("c","b") else round(float(val),2), rtot_fmt)
            dr += 1

        # grand total
        gt = rsum(df)
        gt["_label"] = "Total"
        for ci2, col in enumerate(all_cols):
            key = col["key"]
            if key == "sl":    ws.write(dr, ci2, "", gtot_fmt)
            elif key == "region": ws.write(dr, ci2, "", gtot_fmt)
            elif key == "name": ws.write(dr, ci2, "Total", gtot_lft)
            elif key == "pct":
                tc2 = float(gt.get("tot_c",0)); dc2 = float(gt.get("dig_c",0))
                ws.write(dr, ci2, f"{round(dc2/tc2*100,2) if tc2>0 else 0:.2f}%", gtot_fmt)
            elif key == "pct_a":
                ta2 = float(gt.get("tot_a",0)); da2 = float(gt.get("dig_a",0))
                ws.write(dr, ci2, f"{round(da2/ta2*100,2) if ta2>0 else 0:.2f}%", gtot_fmt)
            else:
                val = gt.get(key,0)
                t   = all_cols[ci2]["t"]
                ws.write(dr, ci2, int(val) if t in ("c","b") else round(float(val),2), gtot_fmt)
        dr += 1

        # ---- auto column width ----
        for ci2, col in enumerate(all_cols):
            # estimate width from header and data
            max_len = max(len(str(col["r1"])), len(str(col["r2"])))
            if col["key"] == "name":
                max_len = max(max_len, df["name"].str.len().max() or 20)
            ws.set_column(ci2, ci2, min(max_len + 4, 30))

    out.seek(0)
    return out.getvalue()

# ========== IMAGE DOWNLOAD ==========
IMG_SCRIPT = """
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script>
function downloadTableImage(tableId, filename) {
    var el = document.getElementById(tableId);
    if (!el) { alert('Table not found — please wait for it to fully load then try again.'); return; }
    html2canvas(el, {scale: 2, useCORS: true, backgroundColor: '#ffffff'}).then(function(canvas) {
        var link = document.createElement('a');
        link.download = filename + '.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });
}
</script>
"""

def img_btn(table_id, filename, label="📷 Download as Image"):
    return f"""
    {IMG_SCRIPT}
    <button onclick="downloadTableImage('{table_id}','{filename}')"
     style="background:#2f3343;color:white;border:none;padding:8px 18px;
            border-radius:6px;cursor:pointer;font-size:14px;margin:6px 4px 0 0;">
     {label}
    </button>"""

# ========== SIDEBAR ==========
st.sidebar.header("Upload Report")
uploaded    = st.sidebar.file_uploader("Upload Booking Payment Report (CSV)", type=["csv"])
report_date = st.sidebar.date_input("Report Date")
use_color   = st.sidebar.checkbox("Colour Coding", value=True,
                  help="Uncheck for black & white output (both image and Excel)")

# ========== MAIN ==========
if uploaded:
    df_raw = pd.read_csv(uploaded)
    df     = process(df_raw)

    if df.empty:
        st.warning("No valid data found.")
        st.stop()

    regions     = df["region"].unique().tolist()
    show_region = len(regions) > 1
    date_str    = report_date.strftime("%d.%m.%Y")
    # For filenames: DD_MM_YYYY
    date_fn     = report_date.strftime("%d_%m_%Y")

    # Legend
    if use_color:
        st.markdown("""
        <div style='display:flex;gap:14px;margin-bottom:8px;font-size:13px;'>
            <span style='background:#90EE90;padding:3px 10px;border-radius:4px;'>≥ 60% Digital</span>
            <span style='background:#FFFACD;padding:3px 10px;border-radius:4px;'>40–59% Digital</span>
            <span style='background:#FF9999;padding:3px 10px;border-radius:4px;'>< 40% Digital</span>
        </div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Count / Transactions", "💰 Amount (₹)", "📋 Combined"])

    with tab1:
        cnt_html = build_count(df, show_region, date_str, use_color)
        st.markdown(cnt_html, unsafe_allow_html=True)
        st.markdown(img_btn("tbl-cnt",
            f"Digital_Transactions_Count_{date_fn}"), unsafe_allow_html=True)
        cnt_xl = build_excel(df, "Count", show_region, date_str, use_color)
        st.download_button("⬇ Download Excel (Count)",
            cnt_xl,
            file_name=f"Digital_Transactions_Count_{date_fn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        amt_html = build_amount(df, show_region, date_str, use_color)
        st.markdown(amt_html, unsafe_allow_html=True)
        st.markdown(img_btn("tbl-amt",
            f"Digital_Transactions_Amount_{date_fn}"), unsafe_allow_html=True)
        amt_xl = build_excel(df, "Amount", show_region, date_str, use_color)
        st.download_button("⬇ Download Excel (Amount)",
            amt_xl,
            file_name=f"Digital_Transactions_Amount_{date_fn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab3:
        comb_html = build_combined(df, show_region, date_str, use_color)
        st.markdown(comb_html, unsafe_allow_html=True)
        st.markdown(img_btn("tbl-comb",
            f"Digital_Transactions_{date_fn}"), unsafe_allow_html=True)
        comb_xl = build_excel(df, "Combined", show_region, date_str, use_color)
        st.download_button("⬇ Download Excel (Combined)",
            comb_xl,
            file_name=f"Digital_Transactions_{date_fn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Upload a Booking Payment Report CSV from the sidebar to begin.")
