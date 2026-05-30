import streamlit as st
import pandas as pd
import numpy as np
import io
from PIL import Image
import os

st.set_page_config(page_title="Digital Transactions", layout="wide", initial_sidebar_state="expanded")

if "authenticated" not in st.session_state or not st.session_state.authenticated:
    st.warning("Please login from the main page.")
    st.stop()

st.markdown("""
<style>
.block-container { padding-top: 1rem !important; padding-bottom: 0rem !important; }
header { visibility: hidden; height: 0px !important; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th {
    background-color: #2f3343; color: white;
    text-align: center !important; padding: 6px 8px;
    border: 1px solid #555; white-space: nowrap;
}
th.grp { background-color: #1a1f2e; }
td { text-align: center; padding: 5px 8px; border: 1px solid #ccc; white-space: nowrap; }
td.left { text-align: left !important; }
tr.region-total td { background-color: #b8d4f0 !important; font-weight: 700; }
tr.grand-total td { background-color: #FFD700 !important; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None
h_l, h_c, h_r = st.columns([1, 8, 1])
with h_l:
    if logo: st.image(logo, width=90)
with h_c:
    st.markdown("""
    <h1 style='font-size:26px;margin-bottom:2px;color:#2f3343;font-weight:700;padding-top:4px;'>
    Digital Transaction Status Report</h1>
    <p style='font-size:14px;color:#555;margin-top:0;'>Headquarter Region - Telangana Postal Circle</p>
    """, unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

REGION_MAP = {
    30530001:"Hyderabad Region", 30530002:"Hyderabad Region", 30530003:"Hyderabad Region",
    30530004:"Hyderabad Region", 30530005:"Hyderabad Region", 30530006:"Hyderabad Region",
    30530007:"Hyderabad Region", 30530008:"Hyderabad Region", 30530009:"Hyderabad Region",
    30530010:"Hyderabad Region", 30530011:"Hyderabad Region", 30600001:"Hyderabad Region",
    30530012:"Headquarters Region", 30530013:"Headquarters Region", 30530014:"Headquarters Region",
    30530015:"Headquarters Region", 30530016:"Headquarters Region", 30530017:"Headquarters Region",
    30600002:"Headquarters Region",
}

DQR_CNT     = ["DQR Scan (Cnt)"]
SBIPOS_CNT  = ["SBIPOS-CARD (Cnt)", "SBIPOS BHARATQR (Cnt)"]
SBIEPAY_CNT = ["SBIEPAY BHARATQR (Cnt)", "SBIEPAY UPI (Cnt)",
               "SBIEPAY Credit Card (Cnt)", "SBIEPAY Debit Card (Cnt)"]
DQR_AMT     = ["DQR Scan (Amt)"]
SBIPOS_AMT  = ["SBIPOS-CARD (Amt)", "SBIPOS BHARATQR (Amt)"]
SBIEPAY_AMT = ["SBIEPAY BHARATQR (Amt)", "SBIEPAY UPI (Amt)",
               "SBIEPAY Credit Card (Amt)", "SBIEPAY Debit Card (Amt)"]

def s(row, cols):
    t = 0
    for c in cols:
        v = pd.to_numeric(row.get(c, 0), errors='coerce')
        t += 0 if pd.isna(v) else v
    return t

def process(df_raw):
    rows = []
    for _, r in df_raw.iterrows():
        oid = r.get("Office ID")
        if pd.isna(oid): continue
        oid_i = int(float(str(oid)))
        region = REGION_MAP.get(oid_i, "Other")
        name   = str(r.get("Office Name","")).strip()

        cash     = pd.to_numeric(r.get("Cash (Cnt)",0), errors='coerce') or 0
        cash_amt = pd.to_numeric(r.get("Cash (Amt)",0), errors='coerce') or 0

        dqr_vals      = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in DQR_CNT}
        sbipos_vals   = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in SBIPOS_CNT}
        sbiepay_vals  = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in SBIEPAY_CNT}
        dqr_avals     = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in DQR_AMT}
        sbipos_avals  = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in SBIPOS_AMT}
        sbiepay_avals = {c: (pd.to_numeric(r.get(c,0),errors='coerce') or 0) for c in SBIEPAY_AMT}

        dqr    = sum(dqr_vals.values())
        sbipos = sum(sbipos_vals.values())
        sbiepay= sum(sbiepay_vals.values())
        dqr_a  = sum(dqr_avals.values())
        sbipos_a = sum(sbipos_avals.values())
        sbiepay_a= sum(sbiepay_avals.values())

        dig   = dqr + sbipos + sbiepay
        dig_a = dqr_a + sbipos_a + sbiepay_a
        total = cash + dig
        total_a = cash_amt + dig_a
        pct = round(dig/total*100, 2) if total > 0 else 0.0

        row_d = {
            "Office ID": oid_i, "Region": region, "Division": name,
            "Cash (Cnt)": int(cash), "Cash (Amt)": round(cash_amt,2),
            "DQR Total (Cnt)": int(dqr),     "DQR Total (Amt)": round(dqr_a,2),
            "SBIPOS Total (Cnt)": int(sbipos),"SBIPOS Total (Amt)": round(sbipos_a,2),
            "SBIePay Total (Cnt)": int(sbiepay),"SBIePay Total (Amt)": round(sbiepay_a,2),
            "Non-Digital (Cnt)": int(cash),
            "Digital Total (Cnt)": int(dig),  "Digital Total (Amt)": round(dig_a,2),
            "Grand Total (Cnt)": int(total),  "Grand Total (Amt)": round(total_a,2),
            "% Digital": pct,
        }
        for c,v in dqr_vals.items():    row_d[f"sub_cnt_{c}"] = int(v)
        for c,v in sbipos_vals.items(): row_d[f"sub_cnt_{c}"] = int(v)
        for c,v in sbiepay_vals.items():row_d[f"sub_cnt_{c}"] = int(v)
        for c,v in dqr_avals.items():    row_d[f"sub_amt_{c}"] = round(v,2)
        for c,v in sbipos_avals.items(): row_d[f"sub_amt_{c}"] = round(v,2)
        for c,v in sbiepay_avals.items():row_d[f"sub_amt_{c}"] = round(v,2)
        rows.append(row_d)
    return pd.DataFrame(rows)

def clr(pct):
    p = float(pct)
    if p >= 60: return "#90EE90"
    elif p >= 40: return "#FFFACD"
    return "#FF9999"

def th(label, colspan=1, rowspan=1, cls=""):
    cs = f' colspan="{colspan}"' if colspan>1 else ''
    rs = f' rowspan="{rowspan}"' if rowspan>1 else ''
    c  = f' class="{cls}"' if cls else ''
    return f'<th{cs}{rs}{c}>{label}</th>'

def build_table(df, view, show_region, date_str):
    is_cnt = view in ("Count","Combined")
    is_amt = view in ("Amount","Combined")

    region_order = [r for r in ["Hyderabad Region","Headquarters Region","Other"]
                    if r in df["Region"].unique()]

    # --- header row 1 ---
    h1 = ""
    fixed = ["Sl.","Region","Division"] if show_region else ["Sl.","Division"]
    for f in fixed: h1 += th(f, rowspan=2)

    cash_span = (1 if is_cnt else 0) + (1 if is_amt else 0)
    dqr_span  = (1 if is_cnt else 0) + (1 if is_amt else 0) + (1 if is_cnt else 0) + (1 if is_amt else 0)
    pos_span  = (2 if is_cnt else 0) + (2 if is_amt else 0) + (1 if is_cnt else 0) + (1 if is_amt else 0)
    pay_span  = (4 if is_cnt else 0) + (4 if is_amt else 0) + (1 if is_cnt else 0) + (1 if is_amt else 0)
    sum_span  = (1 if is_cnt else 0)*3 + (1 if is_amt else 0)*2 + (1 if is_cnt else 0)

    h1 += th("Cash", colspan=cash_span, cls="grp")
    h1 += th("Digital QR (APT)", colspan=dqr_span, cls="grp")
    h1 += th("SBI POS Transactions", colspan=pos_span, cls="grp")
    h1 += th("SBI ePAY Transactions", colspan=pay_span, cls="grp")
    if is_cnt: h1 += th("Non-Digital",rowspan=2)
    if is_cnt: h1 += th("Digital Total",rowspan=2)
    if is_amt: h1 += th("Digital Total (₹)",rowspan=2)
    if is_cnt: h1 += th("Grand Total",rowspan=2)
    if is_amt: h1 += th("Grand Total (₹)",rowspan=2)
    if is_cnt: h1 += th("% Digital",rowspan=2)

    # --- header row 2 ---
    h2 = ""
    if is_cnt: h2 += th("Cnt")
    if is_amt: h2 += th("₹")
    # DQR
    if is_cnt: h2 += th("DQR Scan")
    if is_amt: h2 += th("DQR Scan (₹)")
    if is_cnt: h2 += th("<b>Total</b>")
    if is_amt: h2 += th("<b>Total (₹)</b>")
    # SBIPOS
    if is_cnt: h2 += th("CARD"); h2 += th("BHARAT QR")
    if is_amt: h2 += th("CARD (₹)"); h2 += th("BHARAT QR (₹)")
    if is_cnt: h2 += th("<b>Total</b>")
    if is_amt: h2 += th("<b>Total (₹)</b>")
    # SBIePay
    if is_cnt: h2 += th("BHARAT QR"); h2 += th("UPI"); h2 += th("Credit Card"); h2 += th("Debit Card")
    if is_amt: h2 += th("BHARAT QR (₹)"); h2 += th("UPI (₹)"); h2 += th("Credit Card (₹)"); h2 += th("Debit Card (₹)")
    if is_cnt: h2 += th("<b>Total</b>")
    if is_amt: h2 += th("<b>Total (₹)</b>")

    html = f"""<div style='overflow-x:auto;'>
    <table>
    <caption style='font-size:15px;font-weight:700;color:#c00;padding:8px;
        background:#FFD700;caption-side:top;text-align:center;'>
        Digital Transaction Status Report — {date_str}
    </caption>
    <thead><tr>{h1}</tr><tr>{h2}</tr></thead><tbody>"""

    def data_tds(row, is_region_total=False):
        tds = ""
        if is_cnt: tds += f'<td>{int(row.get("Cash (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("Cash (Amt)",0):,.2f}</td>'
        # DQR
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_DQR Scan (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_DQR Scan (Amt)",0):,.2f}</td>'
        if is_cnt: tds += f'<td><b>{int(row.get("DQR Total (Cnt)",0)):,}</b></td>'
        if is_amt: tds += f'<td><b>{row.get("DQR Total (Amt)",0):,.2f}</b></td>'
        # SBIPOS
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIPOS-CARD (Cnt)",0)):,}</td>'
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIPOS BHARATQR (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIPOS-CARD (Amt)",0):,.2f}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIPOS BHARATQR (Amt)",0):,.2f}</td>'
        if is_cnt: tds += f'<td><b>{int(row.get("SBIPOS Total (Cnt)",0)):,}</b></td>'
        if is_amt: tds += f'<td><b>{row.get("SBIPOS Total (Amt)",0):,.2f}</b></td>'
        # SBIePay
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIEPAY BHARATQR (Cnt)",0)):,}</td>'
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIEPAY UPI (Cnt)",0)):,}</td>'
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIEPAY Credit Card (Cnt)",0)):,}</td>'
        if is_cnt: tds += f'<td>{int(row.get("sub_cnt_SBIEPAY Debit Card (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIEPAY BHARATQR (Amt)",0):,.2f}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIEPAY UPI (Amt)",0):,.2f}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIEPAY Credit Card (Amt)",0):,.2f}</td>'
        if is_amt: tds += f'<td>{row.get("sub_amt_SBIEPAY Debit Card (Amt)",0):,.2f}</td>'
        if is_cnt: tds += f'<td><b>{int(row.get("SBIePay Total (Cnt)",0)):,}</b></td>'
        if is_amt: tds += f'<td><b>{row.get("SBIePay Total (Amt)",0):,.2f}</b></td>'
        # Summary
        if is_cnt: tds += f'<td>{int(row.get("Non-Digital (Cnt)",0)):,}</td>'
        if is_cnt: tds += f'<td>{int(row.get("Digital Total (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("Digital Total (Amt)",0):,.2f}</td>'
        if is_cnt: tds += f'<td>{int(row.get("Grand Total (Cnt)",0)):,}</td>'
        if is_amt: tds += f'<td>{row.get("Grand Total (Amt)",0):,.2f}</td>'
        gt = row.get("Grand Total (Cnt)",0)
        dig = row.get("Digital Total (Cnt)",0)
        pct = round(dig/gt*100,2) if gt > 0 else 0.0
        if is_region_total:
            if is_cnt: tds += f'<td><b>{pct:.2f}%</b></td>'
        else:
            if is_cnt: tds += f'<td><b>{row.get("% Digital",0):.2f}%</b></td>'
        return tds

    sl = 1
    for region in region_order:
        rdf = df[df["Region"]==region].sort_values("% Digital", ascending=False)
        for _, row in rdf.iterrows():
            bg  = clr(row["% Digital"])
            tds = f'<td>{sl}</td>'
            if show_region: tds += f'<td>{row["Region"]}</td>'
            tds += f'<td class="left">{row["Division"]}</td>'
            tds += data_tds(row)
            html += f'<tr style="background:{bg};">{tds}</tr>\n'
            sl += 1

        # Region total
        rt = rdf.sum(numeric_only=True).to_dict()
        tds = f'<td colspan="{2 if show_region else 1}"></td>'
        tds += f'<td class="left"><b>TOTAL — {region}</b></td>'
        tds += data_tds(rt, is_region_total=True)
        html += f'<tr class="region-total">{tds}</tr>\n'

    # Grand total
    gt_row = df.sum(numeric_only=True).to_dict()
    tds = f'<td colspan="{2 if show_region else 1}"></td>'
    tds += f'<td class="left"><b>GRAND TOTAL</b></td>'
    tds += data_tds(gt_row, is_region_total=True)
    html += f'<tr class="grand-total">{tds}</tr>\n'
    html += "</tbody></table></div>"
    return html

# ---- Sidebar ----
st.sidebar.header("Upload Report")
uploaded    = st.sidebar.file_uploader("Upload Booking Payment Report (CSV)", type=["csv"])
report_date = st.sidebar.date_input("Report Date")

if uploaded:
    df_raw = pd.read_csv(uploaded)
    df     = process(df_raw)

    if df.empty:
        st.warning("No valid data found.")
        st.stop()

    regions     = df["Region"].unique().tolist()
    show_region = len(regions) > 1
    date_str    = report_date.strftime("%d.%m.%Y")

    st.markdown("""
    <div style='display:flex;gap:16px;margin-bottom:8px;font-size:13px;'>
        <span style='background:#90EE90;padding:3px 10px;border-radius:4px;'>≥ 60% Digital</span>
        <span style='background:#FFFACD;padding:3px 10px;border-radius:4px;'>40–59% Digital</span>
        <span style='background:#FF9999;padding:3px 10px;border-radius:4px;'>< 40% Digital</span>
    </div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Count", "💰 Amount", "📋 Combined"])
    with tab1: st.markdown(build_table(df,"Count",   show_region, date_str), unsafe_allow_html=True)
    with tab2: st.markdown(build_table(df,"Amount",  show_region, date_str), unsafe_allow_html=True)
    with tab3: st.markdown(build_table(df,"Combined",show_region, date_str), unsafe_allow_html=True)

    # Excel download
    st.markdown("<br>", unsafe_allow_html=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        ecols = ["Office ID","Division","Cash (Cnt)","DQR Total (Cnt)","SBIPOS Total (Cnt)",
                 "SBIePay Total (Cnt)","Non-Digital (Cnt)","Digital Total (Cnt)",
                 "Grand Total (Cnt)","% Digital",
                 "DQR Total (Amt)","SBIPOS Total (Amt)","SBIePay Total (Amt)",
                 "Digital Total (Amt)","Grand Total (Amt)"]
        if show_region: ecols.insert(1,"Region")
        ecols = [c for c in ecols if c in df.columns]
        df[ecols].to_excel(writer, index=False, sheet_name="Digital Txn")
        wb = writer.book; ws = writer.sheets["Digital Txn"]
        g = wb.add_format({"bg_color":"#90EE90"})
        y = wb.add_format({"bg_color":"#FFFACD"})
        rd= wb.add_format({"bg_color":"#FF9999"})
        pi = ecols.index("% Digital")
        for i, pct in enumerate(df["% Digital"], start=1):
            ws.set_row(i, None, g if pct>=60 else (y if pct>=40 else rd))
        for i in range(len(ecols)): ws.set_column(i,i,20)

    st.download_button("⬇ Download Excel Report", out.getvalue(),
        file_name=f"digital_txn_{date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload a Booking Payment Report CSV from the sidebar to begin.")
