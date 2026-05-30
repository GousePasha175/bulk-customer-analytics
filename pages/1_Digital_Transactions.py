import streamlit as st
import pandas as pd
import io, os, base64
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
.block-container { padding-top:1rem !important; padding-bottom:0 !important; }
header { visibility:hidden; height:0px !important; }

/* ---- Report table ---- */
.rpt-wrap { overflow-x:auto; }
.rpt { border-collapse:collapse; font-size:12.5px; width:100%; }
.rpt caption {
    font-size:15px; font-weight:700; color:#c00;
    background:#FFD700; padding:8px; text-align:center;
    caption-side:top;
}
.rpt th {
    background:#2f3343; color:#fff;
    border:1px solid #555; padding:5px 7px;
    text-align:center; white-space:nowrap;
}
.rpt th.grp  { background:#1a1f2e; }
.rpt th.grp2 { background:#263050; }
.rpt td {
    border:1px solid #bbb; padding:4px 7px;
    text-align:center; white-space:nowrap;
}
.rpt td.lft  { text-align:left !important; }
.rpt tr.rtot td { background:#b8d4f0 !important; font-weight:700; }
.rpt tr.gtot td { background:#FFD700 !important; font-weight:700; }
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
    <p style='font-size:13px;color:#555;margin:0;'>Headquarter Region – Telangana Postal Circle</p>
    """, unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 8px;border-color:#ddd;'>", unsafe_allow_html=True)

# ========== REGION MAP ==========
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

# ========== COLUMN DEFS ==========
DQR_CNT     = "DQR Scan (Cnt)"
DQR_AMT     = "DQR Scan (Amt)"
POS_CARD_C  = "SBIPOS-CARD (Cnt)";       POS_CARD_A  = "SBIPOS-CARD (Amt)"
POS_BQR_C   = "SBIPOS BHARATQR (Cnt)";   POS_BQR_A   = "SBIPOS BHARATQR (Amt)"
PAY_BQR_C   = "SBIEPAY BHARATQR (Cnt)";  PAY_BQR_A   = "SBIEPAY BHARATQR (Amt)"
PAY_UPI_C   = "SBIEPAY UPI (Cnt)";       PAY_UPI_A   = "SBIEPAY UPI (Amt)"
PAY_CC_C    = "SBIEPAY Credit Card (Cnt)";PAY_CC_A    = "SBIEPAY Credit Card (Amt)"
PAY_DC_C    = "SBIEPAY Debit Card (Cnt)"; PAY_DC_A    = "SBIEPAY Debit Card (Amt)"

def n(row, col):
    v = pd.to_numeric(row.get(col, 0), errors='coerce')
    return 0 if pd.isna(v) else v

def clr(pct):
    p = float(pct)
    if p >= 60: return "#90EE90"
    if p >= 40: return "#FFFACD"
    return "#FF9999"

def process(df_raw):
    rows = []
    for _, r in df_raw.iterrows():
        oid = r.get("Office ID")
        if pd.isna(oid): continue
        oid_i  = int(float(str(oid)))
        region = REGION_MAP.get(oid_i, "Other")
        name   = str(r.get("Office Name","")).strip()

        cash_c = n(r,"Cash (Cnt)");   cash_a = n(r,"Cash (Amt)")
        dqr_c  = n(r,DQR_CNT);       dqr_a  = n(r,DQR_AMT)
        pc_c   = n(r,POS_CARD_C);    pc_a   = n(r,POS_CARD_A)
        pb_c   = n(r,POS_BQR_C);     pb_a   = n(r,POS_BQR_A)
        eb_c   = n(r,PAY_BQR_C);     eb_a   = n(r,PAY_BQR_A)
        eu_c   = n(r,PAY_UPI_C);     eu_a   = n(r,PAY_UPI_A)
        ec_c   = n(r,PAY_CC_C);      ec_a   = n(r,PAY_CC_A)
        ed_c   = n(r,PAY_DC_C);      ed_a   = n(r,PAY_DC_A)

        pos_c  = pc_c+pb_c;          pos_a  = pc_a+pb_a
        pay_c  = eb_c+eu_c+ec_c+ed_c;pay_a  = eb_a+eu_a+ec_a+ed_a
        dig_c  = dqr_c+pos_c+pay_c;  dig_a  = dqr_a+pos_a+pay_a
        tot_c  = cash_c+dig_c;       tot_a  = cash_a+dig_a
        pct    = round(dig_c/tot_c*100,2) if tot_c>0 else 0.0

        rows.append(dict(
            oid=oid_i, region=region, name=name,
            cash_c=int(cash_c), cash_a=round(cash_a,2),
            dqr_c=int(dqr_c),   dqr_a=round(dqr_a,2),
            pc_c=int(pc_c),  pc_a=round(pc_a,2),
            pb_c=int(pb_c),  pb_a=round(pb_a,2),
            pos_c=int(pos_c),pos_a=round(pos_a,2),
            eb_c=int(eb_c),  eb_a=round(eb_a,2),
            eu_c=int(eu_c),  eu_a=round(eu_a,2),
            ec_c=int(ec_c),  ec_a=round(ec_a,2),
            ed_c=int(ed_c),  ed_a=round(ed_a,2),
            pay_c=int(pay_c),pay_a=round(pay_a,2),
            dig_c=int(dig_c),dig_a=round(dig_a,2),
            tot_c=int(tot_c),tot_a=round(tot_a,2),
            pct=pct,
        ))
    return pd.DataFrame(rows)

# ========== TABLE BUILDERS ==========

def th(label, cs=1, rs=1, cls=""):
    a = f' colspan="{cs}"' if cs>1 else ''
    b = f' rowspan="{rs}"' if rs>1 else ''
    c = f' class="{cls}"' if cls else ''
    return f'<th{a}{b}{c}>{label}</th>'

def fmt_c(v): return f'{int(v):,}'
def fmt_a(v): return f'₹{float(v):,.0f}'

def region_sum(rdf):
    s = {k: rdf[k].sum() for k in rdf.columns if rdf[k].dtype in ['int64','float64','int32']}
    s['tot_c'] = s.get('tot_c',0)
    s['dig_c'] = s.get('dig_c',0)
    s['pct']   = round(s['dig_c']/s['tot_c']*100,2) if s.get('tot_c',0)>0 else 0.0
    return s

# ---- COUNT TABLE ----
def build_count(df, show_region, date_str):
    # Header structure:
    # Row1: Sl | [Region] | Division | Cash | Digital QR(APT) | SBI POS Txn(span 3) | SBI ePAY Txn(span 5) | Non-Digital | Digital | Total | % of Digital Trnx
    # Row2:                            (merged) | (merged)    | Card | Bharat QR | Total | BQR | UPI | CC | DC | Total |

    sr = 2  # default rowspan for single-cell headers
    h1 = th("Sl.", rs=sr) + (th("Region", rs=sr) if show_region else "") + th("Division", rs=sr)
    h1 += th("Cash", rs=sr)
    h1 += th("Digital QR<br>(APT)", rs=sr)
    h1 += th("SBI POS Transactions", cs=3, cls="grp")
    h1 += th("SBI ePAY Transactions", cs=5, cls="grp")
    h1 += th("Non-Digital", rs=sr)
    h1 += th("Digital", rs=sr)
    h1 += th("Total", rs=sr)
    h1 += th("% of Digital Trnx", rs=sr)

    h2 = th("Card") + th("Bharat QR") + th("Total")
    h2 += th("Bharat QR") + th("UPI") + th("Credit Card") + th("Debit Card") + th("Total")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def data_row(d, label=None, row_cls=""):
        tds = ""
        if label is not None:
            span = 2 if show_region else 1
            tds += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        if is_dict := isinstance(d, dict):
            g = lambda k: d.get(k,0)
        else:
            g = lambda k: getattr(d, k, 0)
        tds += f'<td>{fmt_c(g("cash_c"))}</td>'
        tds += f'<td>{fmt_c(g("dqr_c"))}</td>'
        tds += f'<td>{fmt_c(g("pc_c"))}</td>'
        tds += f'<td>{fmt_c(g("pb_c"))}</td>'
        tds += f'<td><b>{fmt_c(g("pos_c"))}</b></td>'
        tds += f'<td>{fmt_c(g("eb_c"))}</td>'
        tds += f'<td>{fmt_c(g("eu_c"))}</td>'
        tds += f'<td>{fmt_c(g("ec_c"))}</td>'
        tds += f'<td>{fmt_c(g("ed_c"))}</td>'
        tds += f'<td><b>{fmt_c(g("pay_c"))}</b></td>'
        tds += f'<td>{fmt_c(g("cash_c"))}</td>'
        tds += f'<td>{fmt_c(g("dig_c"))}</td>'
        tds += f'<td>{fmt_c(g("tot_c"))}</td>'
        tds += f'<td><b>{float(g("pct")):.2f}%</b></td>'
        rc = f' class="{row_cls}"' if row_cls else ''
        return f'<tr{rc}>{tds}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-cnt">
    <table class="rpt">
    <caption>Digital Transaction Status – Count/Transactions — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct", ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"])
            pre = f'<td>{sl}</td>' + (f'<td>{row["region"]}</td>' if show_region else "") + f'<td class="lft">{row["name"]}</td>'
            html += f'<tr style="background:{bg};">{pre}{data_row(row, row_cls="")[4:-6]}</tr>\n'
            sl += 1
        s = region_sum(rdf)
        html += data_row(s, label=f"TOTAL — {region}", row_cls="rtot")

    gt = region_sum(df)
    html += data_row(gt, label="GRAND TOTAL", row_cls="gtot")
    html += "</tbody></table></div>"
    return html

# ---- AMOUNT TABLE ----
def build_amount(df, show_region, date_str):
    sr = 2
    h1 = th("Sl.", rs=sr) + (th("Region", rs=sr) if show_region else "") + th("Division", rs=sr)
    h1 += th("Cash (₹)", rs=sr)
    h1 += th("Digital QR<br>(APT) (₹)", rs=sr)
    h1 += th("SBI POS Transactions (₹)", cs=3, cls="grp")
    h1 += th("SBI ePAY Transactions (₹)", cs=5, cls="grp")
    h1 += th("Non-Digital (₹)", rs=sr)
    h1 += th("Digital (₹)", rs=sr)
    h1 += th("Total (₹)", rs=sr)
    h1 += th("% of Digital Trnx", rs=sr)

    h2 = th("Card") + th("Bharat QR") + th("Total")
    h2 += th("Bharat QR") + th("UPI") + th("Credit Card") + th("Debit Card") + th("Total")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def data_row(d, label=None, row_cls=""):
        tds = ""
        if label is not None:
            span = 2 if show_region else 1
            tds += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        g = lambda k: d.get(k,0) if isinstance(d,dict) else getattr(d,k,0)
        tds += f'<td>{fmt_a(g("cash_a"))}</td>'
        tds += f'<td>{fmt_a(g("dqr_a"))}</td>'
        tds += f'<td>{fmt_a(g("pc_a"))}</td>'
        tds += f'<td>{fmt_a(g("pb_a"))}</td>'
        tds += f'<td><b>{fmt_a(g("pos_a"))}</b></td>'
        tds += f'<td>{fmt_a(g("eb_a"))}</td>'
        tds += f'<td>{fmt_a(g("eu_a"))}</td>'
        tds += f'<td>{fmt_a(g("ec_a"))}</td>'
        tds += f'<td>{fmt_a(g("ed_a"))}</td>'
        tds += f'<td><b>{fmt_a(g("pay_a"))}</b></td>'
        tds += f'<td>{fmt_a(g("cash_a"))}</td>'
        tds += f'<td>{fmt_a(g("dig_a"))}</td>'
        tds += f'<td>{fmt_a(g("tot_a"))}</td>'
        tc = float(g("tot_c")); dc = float(g("dig_c"))
        pct = round(dc/tc*100,2) if tc>0 else 0.0
        tds += f'<td><b>{pct:.2f}%</b></td>'
        rc = f' class="{row_cls}"' if row_cls else ''
        return f'<tr{rc}>{tds}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-amt">
    <table class="rpt">
    <caption>Digital Transaction Status – Amount (₹) — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct", ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"])
            pre = f'<td>{sl}</td>' + (f'<td>{row["region"]}</td>' if show_region else "") + f'<td class="lft">{row["name"]}</td>'
            html += f'<tr style="background:{bg};">{pre}{data_row(row, row_cls="")[4:-6]}</tr>\n'
            sl += 1
        s = region_sum(rdf)
        html += data_row(s, label=f"TOTAL — {region}", row_cls="rtot")

    gt = region_sum(df)
    html += data_row(gt, label="GRAND TOTAL", row_cls="gtot")
    html += "</tbody></table></div>"
    return html

# ---- COMBINED TABLE ----
def build_combined(df, show_region, date_str):
    # Structure:
    # Row1: Sl|[Reg]|Div | Cash(2) | Dig QR(2) | SBI POS(6) | SBI ePAY(10) | Non-Digital(2) | Digital(2) | Total(2) | %Digital(2)
    # Row2:             | Cnt|₹   | Cnt|₹      | Card Cnt|₹ BQR Cnt|₹ Tot Cnt|₹ | BQR Cnt|₹ UPI Cnt|₹ CC Cnt|₹ DC Cnt|₹ Tot Cnt|₹ | Cnt|₹ | Cnt|₹ | Cnt|₹ | Cnt|₹

    sr = 2
    h1 = th("Sl.", rs=sr) + (th("Region", rs=sr) if show_region else "") + th("Division", rs=sr)
    h1 += th("Cash", cs=2, cls="grp")
    h1 += th("Digital QR (APT)", cs=2, cls="grp")
    h1 += th("SBI POS Transactions", cs=6, cls="grp")
    h1 += th("SBI ePAY Transactions", cs=10, cls="grp")
    h1 += th("Non-Digital", cs=2, cls="grp")
    h1 += th("Digital", cs=2, cls="grp")
    h1 += th("Total", cs=2, cls="grp")
    h1 += th("% of Digital Trnx", cs=2, cls="grp")

    h2 = th("Cnt") + th("₹")               # Cash
    h2 += th("Cnt") + th("₹")              # DQR
    h2 += th("Card(Cnt)") + th("Card(₹)") + th("Bharat QR(Cnt)") + th("Bharat QR(₹)") + th("Total(Cnt)") + th("Total(₹)")  # POS
    h2 += th("BQR(Cnt)") + th("BQR(₹)") + th("UPI(Cnt)") + th("UPI(₹)") + th("CC(Cnt)") + th("CC(₹)") + th("DC(Cnt)") + th("DC(₹)") + th("Total(Cnt)") + th("Total(₹)")  # ePAY
    h2 += th("Cnt") + th("₹")              # Non-digital
    h2 += th("Cnt") + th("₹")              # Digital
    h2 += th("Cnt") + th("₹")              # Total
    h2 += th("Cnt") + th("₹")              # % (Cnt-based and Amt-based)

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["region"].unique()]

    def data_row(d, label=None, row_cls=""):
        tds = ""
        if label is not None:
            span = 2 if show_region else 1
            tds += f'<td colspan="{span}"></td><td class="lft"><b>{label}</b></td>'
        g = lambda k: d.get(k,0) if isinstance(d,dict) else getattr(d,k,0)
        tds += f'<td>{fmt_c(g("cash_c"))}</td><td>{fmt_a(g("cash_a"))}</td>'
        tds += f'<td>{fmt_c(g("dqr_c"))}</td><td>{fmt_a(g("dqr_a"))}</td>'
        tds += f'<td>{fmt_c(g("pc_c"))}</td><td>{fmt_a(g("pc_a"))}</td>'
        tds += f'<td>{fmt_c(g("pb_c"))}</td><td>{fmt_a(g("pb_a"))}</td>'
        tds += f'<td><b>{fmt_c(g("pos_c"))}</b></td><td><b>{fmt_a(g("pos_a"))}</b></td>'
        tds += f'<td>{fmt_c(g("eb_c"))}</td><td>{fmt_a(g("eb_a"))}</td>'
        tds += f'<td>{fmt_c(g("eu_c"))}</td><td>{fmt_a(g("eu_a"))}</td>'
        tds += f'<td>{fmt_c(g("ec_c"))}</td><td>{fmt_a(g("ec_a"))}</td>'
        tds += f'<td>{fmt_c(g("ed_c"))}</td><td>{fmt_a(g("ed_a"))}</td>'
        tds += f'<td><b>{fmt_c(g("pay_c"))}</b></td><td><b>{fmt_a(g("pay_a"))}</b></td>'
        tds += f'<td>{fmt_c(g("cash_c"))}</td><td>{fmt_a(g("cash_a"))}</td>'
        tds += f'<td>{fmt_c(g("dig_c"))}</td><td>{fmt_a(g("dig_a"))}</td>'
        tds += f'<td>{fmt_c(g("tot_c"))}</td><td>{fmt_a(g("tot_a"))}</td>'
        tc = float(g("tot_c")); dc = float(g("dig_c"))
        ta = float(g("tot_a")); da = float(g("dig_a"))
        pct_c = round(dc/tc*100,2) if tc>0 else 0.0
        pct_a = round(da/ta*100,2) if ta>0 else 0.0
        tds += f'<td><b>{pct_c:.2f}%</b></td><td><b>{pct_a:.2f}%</b></td>'
        rc = f' class="{row_cls}"' if row_cls else ''
        return f'<tr{rc}>{tds}</tr>\n'

    html = f'''<div class="rpt-wrap" id="tbl-comb">
    <table class="rpt">
    <caption>Digital Transaction Status – Combined (Count + Amount) — {date_str}</caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>'''

    sl = 1
    for region in region_order:
        rdf = df[df["region"]==region].sort_values("pct", ascending=False)
        for _, row in rdf.iterrows():
            bg = clr(row["pct"])
            pre = f'<td>{sl}</td>' + (f'<td>{row["region"]}</td>' if show_region else "") + f'<td class="lft">{row["name"]}</td>'
            html += f'<tr style="background:{bg};">{pre}{data_row(row, row_cls="")[4:-6]}</tr>\n'
            sl += 1
        s = region_sum(rdf)
        html += data_row(s, label=f"TOTAL — {region}", row_cls="rtot")

    gt = region_sum(df)
    html += data_row(gt, label="GRAND TOTAL", row_cls="gtot")
    html += "</tbody></table></div>"
    return html

# ========== IMAGE DOWNLOAD HELPER ==========
def html_to_download_btn(html_content, filename, label):
    """Render a JS-based screenshot button using html2canvas"""
    encoded = base64.b64encode(html_content.encode()).decode()
    btn_html = f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <button onclick="
        var el = document.getElementById('{filename}');
        html2canvas(el, {{scale:2, useCORS:true}}).then(function(canvas){{
            var a = document.createElement('a');
            a.download = '{filename}.png';
            a.href = canvas.toDataURL('image/png');
            a.click();
        }});
    " style="background:#2f3343;color:white;border:none;padding:8px 18px;
             border-radius:6px;cursor:pointer;font-size:14px;margin-top:6px;">
    📷 Download as Image
    </button>
    """
    return btn_html

# ========== SIDEBAR ==========
st.sidebar.header("Upload Report")
uploaded    = st.sidebar.file_uploader("Upload Booking Payment Report (CSV)", type=["csv"])
report_date = st.sidebar.date_input("Report Date")

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

    # Legend
    st.markdown("""
    <div style='display:flex;gap:14px;margin-bottom:8px;font-size:13px;'>
        <span style='background:#90EE90;padding:3px 10px;border-radius:4px;'>≥ 60% Digital</span>
        <span style='background:#FFFACD;padding:3px 10px;border-radius:4px;'>40–59% Digital</span>
        <span style='background:#FF9999;padding:3px 10px;border-radius:4px;'>< 40% Digital</span>
    </div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Count / Transactions", "💰 Amount (₹)", "📋 Combined"])

    with tab1:
        cnt_html = build_count(df, show_region, date_str)
        st.markdown(cnt_html, unsafe_allow_html=True)
        st.markdown(html_to_download_btn(cnt_html, "tbl-cnt", "Count"), unsafe_allow_html=True)

    with tab2:
        amt_html = build_amount(df, show_region, date_str)
        st.markdown(amt_html, unsafe_allow_html=True)
        st.markdown(html_to_download_btn(amt_html, "tbl-amt", "Amount"), unsafe_allow_html=True)

    with tab3:
        comb_html = build_combined(df, show_region, date_str)
        st.markdown(comb_html, unsafe_allow_html=True)
        st.markdown(html_to_download_btn(comb_html, "tbl-comb", "Combined"), unsafe_allow_html=True)

    # Excel download
    st.markdown("<br>", unsafe_allow_html=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        export = df[["oid","region","name","cash_c","cash_a",
                     "dqr_c","dqr_a","pc_c","pc_a","pb_c","pb_a","pos_c","pos_a",
                     "eb_c","eb_a","eu_c","eu_a","ec_c","ec_a","ed_c","ed_a",
                     "pay_c","pay_a","dig_c","dig_a","tot_c","tot_a","pct"]].copy()
        export.columns = ["Office ID","Region","Division","Cash Cnt","Cash Amt",
                          "DQR Cnt","DQR Amt","POS Card Cnt","POS Card Amt",
                          "POS BQR Cnt","POS BQR Amt","SBI POS Total Cnt","SBI POS Total Amt",
                          "ePAY BQR Cnt","ePAY BQR Amt","ePAY UPI Cnt","ePAY UPI Amt",
                          "ePAY CC Cnt","ePAY CC Amt","ePAY DC Cnt","ePAY DC Amt",
                          "SBI ePAY Total Cnt","SBI ePAY Total Amt",
                          "Digital Cnt","Digital Amt","Total Cnt","Total Amt","% Digital"]
        if not show_region:
            export = export.drop(columns=["Region"])
        export.to_excel(writer, index=False, sheet_name="Digital Txn")
        wb = writer.book; ws = writer.sheets["Digital Txn"]
        g = wb.add_format({"bg_color":"#90EE90"})
        y = wb.add_format({"bg_color":"#FFFACD"})
        rd= wb.add_format({"bg_color":"#FF9999"})
        for i, pct in enumerate(df["pct"], start=1):
            ws.set_row(i, None, g if pct>=60 else (y if pct>=40 else rd))
        for i in range(len(export.columns)):
            ws.set_column(i,i,18)

    st.download_button("⬇ Download Excel Report", out.getvalue(),
        file_name=f"digital_txn_{date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Upload a Booking Payment Report CSV from the sidebar to begin.")
