import streamlit as st
import pandas as pd
import numpy as np
import calendar
import io
import os
import re
import glob as _glob

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    _bulk = (_glob.glob("pages/Bulk_Analytics.py") + _glob.glob("pages/*[Bb]ulk*.py"))
    if _bulk: st.sidebar.page_link(_bulk[0].replace("\\","/"), label="\U0001f4ca Bulk Customer Analytics")
    _posb = (_glob.glob("pages/POSB Daily Report.py") + _glob.glob("pages/*[Pp][Oo][Ss][Bb]*.py"))
    if _posb: st.sidebar.page_link(_posb[0].replace("\\","/"), label="\U0001f4ee POSB Daily Report")
    _dig  = (_glob.glob("pages/1_Digital_Transactions.py") + _glob.glob("pages/*[Dd]igital*.py"))
    if _dig:  st.sidebar.page_link(_dig[0].replace("\\","/"),  label="\U0001f4bb Digital Transactions")
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

from PIL import Image

st.set_page_config(page_title="Bulk Customer Analytics", page_icon="\U0001f4ca", layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
[data-testid="stSidebarNav"]  { display: none !important; }
.block-container { padding-top: 1.2rem !important; padding-bottom: 0rem !important; }
header { visibility: hidden; height: 0px !important; }
[data-testid="collapsedControl"] {
    display: flex !important; visibility: visible !important; opacity: 1 !important;
    position: fixed !important; top: 50% !important; left: 0px !important;
    transform: translateY(-50%) !important; z-index: 999999 !important;
    background-color: #2f3343 !important; border-radius: 0 8px 8px 0 !important;
    padding: 12px 7px !important; box-shadow: 3px 0 8px rgba(0,0,0,0.35) !important; cursor: pointer !important;
}
[data-testid="collapsedControl"] button { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="collapsedControl"] svg    { fill: white !important; color: white !important; }
</style>""", unsafe_allow_html=True)

if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar to log in.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Palette ───────────────────────────────────────────────────────────────────
PALETTE = {
    "title_bg":"#1F3864","title_fg":"#FFFFFF","header_bg":"#2E75B6","header_fg":"#FFFFFF",
    "sub_hdr_bg":"#9DC3E6","sub_hdr_fg":"#000000",
    "excellent":"#70AD47","excellent_fg":"#FFFFFF","normal":"#FFFF00","normal_fg":"#000000",
    "warning":"#FFC000","warning_fg":"#000000","critical":"#FF0000","critical_fg":"#FFFFFF",
    "no_hist":"#D3D3D3","no_hist_fg":"#000000","total_bg":"#FFF2CC","total_fg":"#000000",
}
MASTER_FOLDER = "master"
STATUS_ORDER  = ["Excellent","Normal","Warning","Critical","No Historical Data"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def format_indian(n):
    try: n=int(round(n))
    except: return str(n)
    neg=n<0; n=abs(n); s=str(n)
    if len(s)<=3: return ("-" if neg else "")+s
    last3=s[-3:]; rest=s[:-3]; parts=[]
    while len(rest)>2: parts.append(rest[-2:]); rest=rest[:-2]
    if rest: parts.append(rest)
    parts.reverse()
    return ("-" if neg else "")+",".join(parts)+","+last3

def parse_dates_robust(series):
    for fmt in ["%d-%m-%Y","%d/%m/%Y","%d-%m-%y","%d/%m/%y","%d %m %Y","%d %b %Y",
                "%d-%b-%Y","%d/%b/%Y","%Y-%m-%d","%m/%d/%Y","%m-%d-%Y"]:
        try: return pd.to_datetime(series, format=fmt, errors='raise')
        except: continue
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

def classify(variance, threshold):
    if pd.isna(variance): return "No Historical Data"
    if variance>=threshold: return "Excellent"
    elif variance>=0: return "Normal"
    elif variance>=-threshold: return "Warning"
    else: return "Critical"

def color_status(val):
    return {"Excellent":f"background-color:{PALETTE['excellent']};color:{PALETTE['excellent_fg']}",
            "Normal":f"background-color:{PALETTE['normal']};color:{PALETTE['normal_fg']}",
            "Warning":f"background-color:{PALETTE['warning']};color:{PALETTE['warning_fg']}",
            "Critical":f"background-color:{PALETTE['critical']};color:{PALETTE['critical_fg']}",
            "No Historical Data":f"background-color:{PALETTE['no_hist']};color:{PALETTE['no_hist_fg']}"}.get(val,"")

def detect_daily_columns(df):
    cid=cname=rev=trf=sdt=edt=None
    for col in df.columns:
        c=str(col).strip().lower()
        if "customer id" in c: cid=col
        elif "customer name" in c: cname=col
        elif "amount" in c or "revenue" in c: rev=col
        elif "article" in c or "traffic" in c: trf=col
        elif "start" in c: sdt=col
        elif "end" in c: edt=col
    return cid,cname,rev,trf,sdt,edt

def get_fy_year(year, month):
    """Return FY label for a given year/month (Indian FY Apr–Mar)."""
    return year if month >= 4 else year - 1

def load_master_from_folder():
    """Load all CSVs from master/ folder into a unified pivoted DataFrame."""
    csv_files = sorted(_glob.glob(f"{MASTER_FOLDER}/*.csv")+_glob.glob(f"{MASTER_FOLDER}/*.CSV"))
    if not csv_files: return pd.DataFrame(), []
    all_pieces=[]; loaded=[]
    for fp in csv_files:
        try:
            df=pd.read_csv(fp)
            cid_col,name_col,rev_col,trf_col,start_col,_=detect_daily_columns(df)
            if not all([cid_col,rev_col,start_col]): continue
            df[start_col]=parse_dates_robust(df[start_col])
            df["_year"]=df[start_col].dt.year; df["_month"]=df[start_col].dt.month
            df["_cid"]=df[cid_col].astype(str).str.replace(".0","",regex=False).str.strip()
            df["_name"]=df[name_col].astype(str) if name_col else ""
            df["_rev"]=pd.to_numeric(df[rev_col],errors="coerce").fillna(0)
            df["_trf"]=pd.to_numeric(df[trf_col],errors="coerce").fillna(0) if trf_col else 0
            all_pieces.append(df[["_cid","_name","_year","_month","_rev","_trf"]])
            y,m=int(df["_year"].iloc[0]),int(df["_month"].iloc[0])
            loaded.append(f"{calendar.month_name[m]} {y}")
        except: continue
    if not all_pieces: return pd.DataFrame(), []
    combined=pd.concat(all_pieces,ignore_index=True)
    grouped=(combined.groupby(["_cid","_year","_month"])
             .agg(_name=("_name","first"),Revenue=("_rev","sum"),Traffic=("_trf","sum"))
             .reset_index())
    name_map=combined.groupby("_cid")["_name"].first().to_dict()
    master={}
    for _,row in grouped.iterrows():
        cid=row["_cid"]; key=f"{int(row['_year'])}-{int(row['_month']):02d}"
        if cid not in master: master[cid]={"CUSTOMER ID":cid,"CUSTOMER NAME":str(name_map.get(cid,""))}
        master[cid][f"{key} REVENUE"]=row["Revenue"]; master[cid][f"{key} TRAFFIC"]=row["Traffic"]
    return pd.DataFrame(list(master.values())), loaded

def parse_master_excel(file):
    xl=pd.ExcelFile(file); master={}
    for sheet in xl.sheet_names:
        raw=pd.read_excel(file,sheet_name=sheet,header=None)
        id_ci=nm_ci=None
        has_sub=any(str(v).strip().upper()=='TRAFFIC' for v in raw.iloc[1].tolist())
        ds=2 if has_sub else 1
        for scan in range(min(3,len(raw))):
            for ci,val in enumerate(raw.iloc[scan]):
                s=str(val).strip().lower()
                if ('customer id' in s or 'cust id' in s) and id_ci is None: id_ci=ci
                if ('customer name' in s or 'cutomer name' in s) and nm_ci is None: nm_ci=ci
        if id_ci is None: continue
        mm={}
        if has_sub:
            dr=raw.iloc[0]; sr=raw.iloc[1]; cd=None
            for ci in range(len(dr)):
                v=dr.iloc[ci]
                if hasattr(v,'year'): cd=(v.year,v.month); mm[cd]={'t':None,'r':None}
                if cd:
                    sub=str(sr.iloc[ci]).strip().upper()
                    if sub=='TRAFFIC': mm[cd]['t']=ci
                    elif sub in ('REVENUE','REV'): mm[cd]['r']=ci
        else:
            ma={'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,'jan':1,'feb':2,'mar':3}
            for ci,val in enumerate(raw.iloc[0].tolist()):
                s=str(val).strip().lower()
                for ab,mo in ma.items():
                    if ab in s:
                        ym=re.search(r'(\d{2})',s)
                        if ym:
                            yr=2000+int(ym.group(1)); k=(yr,mo)
                            if k not in mm: mm[k]={'t':None,'r':None}
                            if 'traf' in s or 'trf' in s: mm[k]['t']=ci
                            elif 'rev' in s: mm[k]['r']=ci
                        break
        if not mm: continue
        for ri in range(ds,len(raw)):
            row=raw.iloc[ri]; cv=row.iloc[id_ci]
            if pd.isna(cv): continue
            s=str(cv).strip().lower()
            if s in ('','nan','total','grand total','sl no','sno'): continue
            cid=str(cv).replace('.0','').strip(); cn=''
            if nm_ci is not None and nm_ci<len(row):
                cn=str(row.iloc[nm_ci]).strip()
                if cn.lower() in ('nan',''): cn=''
            if cid not in master: master[cid]={'CUSTOMER ID':cid,'CUSTOMER NAME':cn}
            for (yr,mo),cols in mm.items():
                tk=f"{yr}-{mo:02d} TRAFFIC"; rk=f"{yr}-{mo:02d} REVENUE"; tv=rv=0
                if cols['t'] is not None and cols['t']<len(row):
                    tv=pd.to_numeric(row.iloc[cols['t']],errors='coerce'); tv=0 if pd.isna(tv) else tv
                if cols['r'] is not None and cols['r']<len(row):
                    rv=pd.to_numeric(row.iloc[cols['r']],errors='coerce'); rv=0 if pd.isna(rv) else rv
                master[cid][tk]=master[cid].get(tk,0)+tv; master[cid][rk]=master[cid].get(rk,0)+rv
    return pd.DataFrame(list(master.values())) if master else pd.DataFrame()

def parse_master_from_daily_format(file):
    try: df=pd.read_csv(file)
    except: return pd.DataFrame()
    cid_col,name_col,rev_col,trf_col,start_col,_=detect_daily_columns(df)
    if not all([cid_col,rev_col,trf_col,start_col]): return pd.DataFrame()
    df[start_col]=parse_dates_robust(df[start_col])
    df["_year"]=df[start_col].dt.year; df["_month"]=df[start_col].dt.month
    df["_cid"]=df[cid_col].astype(str).str.replace(".0","",regex=False).str.strip()
    grouped=(df.groupby(["_cid","_year","_month"]).agg(Revenue=(rev_col,"sum"),Traffic=(trf_col,"sum")).reset_index())
    nm=df.groupby("_cid")[name_col].first().to_dict() if name_col else {}
    master={}
    for _,row in grouped.iterrows():
        cid=row["_cid"]; k=f"{int(row['_year'])}-{int(row['_month']):02d}"
        if cid not in master: master[cid]={"CUSTOMER ID":cid,"CUSTOMER NAME":str(nm.get(cid,""))}
        master[cid][f"{k} REVENUE"]=row["Revenue"]; master[cid][f"{k} TRAFFIC"]=row["Traffic"]
    return pd.DataFrame(list(master.values())) if master else pd.DataFrame()

def get_month_totals(hist_df, year, month):
    """Return (rev, trf) for a specific year-month from master df, or (None,None)."""
    rk=f"{year}-{month:02d} REVENUE"; tk=f"{year}-{month:02d} TRAFFIC"
    if rk not in hist_df.columns: return None, None
    return rk, tk

def find_highest_month(hist_df, exclude_key=None, fy_filter=None):
    """
    Find the month with highest total revenue across all customers.
    fy_filter: 'last' = last FY only, 'current' = current FY only, None = all
    exclude_key: 'YYYY-MM' string to exclude (the uploaded period's month)
    Returns (year, month, label) or (None,None,None)
    """
    rev_cols=[c for c in hist_df.columns if c.endswith(" REVENUE")]
    best_rev=-1; best_yr=best_mo=None
    for col in rev_cols:
        m=re.match(r'(\d{4})-(\d{2}) REVENUE',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        key=f"{yr}-{mo:02d}"
        if exclude_key and key==exclude_key: continue
        fy=get_fy_year(yr,mo)
        if fy_filter=='last' and fy!=_LAST_FY: continue
        if fy_filter=='current' and fy!=_CURRENT_FY: continue
        total=pd.to_numeric(hist_df[col],errors='coerce').sum()
        if total>best_rev: best_rev=total; best_yr=yr; best_mo=mo
    if best_yr is None: return None,None,None
    return best_yr,best_mo,f"{calendar.month_name[best_mo]} {best_yr}"

def compute_variance(current_val, expected_val, sd_pct):
    if expected_val is None or expected_val==0: return np.nan, "No Historical Data"
    v=((current_val-expected_val)/expected_val)*100
    return round(v,2), classify(v, sd_pct)

def _make_excel_formats(workbook):
    def f(**kw):
        base={"border":1,"valign":"vcenter"}; base.update(kw); return workbook.add_format(base)
    return {
        "title":     f(bold=True,font_size=13,align="center",  bg_color=PALETTE["title_bg"],  font_color=PALETTE["title_fg"]),
        "header":    f(bold=True,align="center",text_wrap=True,bg_color=PALETTE["header_bg"], font_color=PALETTE["header_fg"]),
        "sub_hdr":   f(bold=True,align="center",               bg_color=PALETTE["sub_hdr_bg"],font_color=PALETTE["sub_hdr_fg"]),
        "grp_hdr":   f(bold=True,font_size=12,                 bg_color=PALETTE["title_bg"],  font_color=PALETTE["title_fg"]),
        "excellent": f(bg_color=PALETTE["excellent"],font_color=PALETTE["excellent_fg"]),
        "normal":    f(bg_color=PALETTE["normal"],   font_color=PALETTE["normal_fg"]),
        "warning":   f(bg_color=PALETTE["warning"],  font_color=PALETTE["warning_fg"]),
        "critical":  f(bg_color=PALETTE["critical"], font_color=PALETTE["critical_fg"]),
        "no_hist":   f(bg_color=PALETTE["no_hist"],  font_color=PALETTE["no_hist_fg"]),
        "total":     f(bold=True,align="center",bg_color=PALETTE["total_bg"],font_color=PALETTE["total_fg"]),
        "total_l":   f(bold=True,align="left",  bg_color=PALETTE["total_bg"],font_color=PALETTE["total_fg"]),
        "plain":     f(),
        "plain_l":   f(align="left"),
    }

def write_grouped_sheet(writer, df, sheet_name, fmts, status_col="Revenue Status"):
    wb=writer.book; ws=wb.add_worksheet(sheet_name); writer.sheets[sheet_name]=ws
    cols=list(df.columns); cw=22; cr=0
    rev_ci=cols.index("Revenue Status") if "Revenue Status" in cols else None
    trf_ci=cols.index("Traffic Status") if "Traffic Status" in cols else None
    for status in STATUS_ORDER:
        grp=df[df[status_col]==status]
        if grp.empty: continue
        ws.merge_range(cr,0,cr,len(cols)-1,f"{status}  ({len(grp)})",fmts["grp_hdr"]); cr+=1
        for ci,col in enumerate(cols): ws.write(cr,ci,col,fmts["sub_hdr"]); ws.set_column(ci,ci,cw)
        cr+=1
        for _,row in grp.iterrows():
            for ci,col in enumerate(cols):
                val=row[col]
                if isinstance(val,float) and np.isnan(val): val=""
                cfmt=fmts["plain"]
                if ci==rev_ci or ci==trf_ci:
                    cfmt=fmts.get(str(val).lower().replace(" ","_").replace("no_historical_data","no_hist"),fmts["plain"])
                ws.write(cr,ci,val,cfmt)
            cr+=1
        cr+=1

# ── Load logo ─────────────────────────────────────────────────────────────────
logo_path="assets/logo.png"
logo=Image.open(logo_path) if os.path.exists(logo_path) else None

# ── Header ────────────────────────────────────────────────────────────────────
h_l,h_c,h_r=st.columns([1,8,1])
with h_l:
    if logo: st.image(logo,width=90)
with h_c:
    st.markdown("""<h1 style='font-size:28px;margin-bottom:2px;color:#2f3343;font-weight:700;padding-top:4px;'>
    Bulk Customer Business Analytics</h1>
    <p style='font-size:15px;color:#555;margin-top:0;'>Headquarters Region – Telangana Postal Circle</p>
    """,unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>",unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.header("Upload Files")
daily_file  = st.sidebar.file_uploader("Daily / Period File (CSV)", type=["csv"])
master_file = st.sidebar.file_uploader("Master Data File (optional – overrides default)", type=["xlsx","xls","csv"])

_xl_cands     = _glob.glob("data/[Mm]aster.xlsx")+_glob.glob("data/[Mm]aster.xls")
DEFAULT_MASTER = _xl_cands[0] if _xl_cands else None
_folder_csvs   = sorted(_glob.glob(f"{MASTER_FOLDER}/*.csv")+_glob.glob(f"{MASTER_FOLDER}/*.CSV"))

if master_file:        st.sidebar.success("✅ Using uploaded master file")
elif _folder_csvs:     st.sidebar.success(f"📂 Master folder: {len(_folder_csvs)} CSV(s) found")
elif DEFAULT_MASTER:   st.sidebar.info("📂 Using default master Excel")
else:                  st.sidebar.warning("⚠️ No master data found.")

sd_percent = st.sidebar.slider("Deviation %", min_value=1, max_value=50, value=10)
show_mode  = st.sidebar.radio("Filter Records",
    ["All records (mark unmatched as 'No Historical Data')",
     "Only records present in Master Data"], index=0)

# ── Comparison mode ───────────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.subheader("📊 Comparison Mode")
comparison_base = st.sidebar.radio(
    "Base comparison against:",
    ["Previous month (same period)",
     "Last FY corresponding month",
     "Average of selected base"], index=0,
    help="Selects which historical period to compare against.")

highest_mode = st.sidebar.radio(
    "Also compare against highest month:",
    ["None",
     "Highest in last FY",
     "Highest in current FY (excl. uploaded period)"],
    index=0,
    help="Adds a second comparison column showing performance vs best month.")

use_average_fallback = st.sidebar.checkbox(
    "If base period unavailable, fall back to average of available months in same FY",
    value=True)

# ── Load master ───────────────────────────────────────────────────────────────
folder_master_df=pd.DataFrame(); available_months=[]
if not master_file and _folder_csvs:
    folder_master_df, available_months = load_master_from_folder()
    if available_months:
        st.sidebar.caption(f"Months in master: {', '.join(available_months)}")

# ── MAIN PROCESS ──────────────────────────────────────────────────────────────
if not daily_file:
    st.info("Please upload the Daily / Period CSV file in the sidebar to begin.")
    st.stop()

if not master_file and not _folder_csvs and not DEFAULT_MASTER:
    st.info("Please also upload the Master Data file in the sidebar.")
    st.stop()

# Load daily
daily_df=pd.read_csv(daily_file)
cid_col,cname_col,rev_col,trf_col,sdt_col,edt_col=detect_daily_columns(daily_df)
missing=[n for n,v in [("Customer ID",cid_col),("Customer Name",cname_col),
                        ("Revenue/Amount",rev_col),("Traffic/Articles",trf_col),
                        ("Start Date",sdt_col)] if v is None]
if missing: st.error(f"Could not detect columns: {', '.join(missing)}"); st.stop()

daily_df[sdt_col]=parse_dates_robust(daily_df[sdt_col])
daily_df[edt_col]=parse_dates_robust(daily_df[edt_col]) if edt_col else daily_df[sdt_col]

upload_start=pd.to_datetime(daily_df[sdt_col].iloc[0],errors="coerce")
upload_end  =pd.to_datetime(daily_df[edt_col].iloc[0],errors="coerce")
uploaded_days=(upload_end-upload_start).days+1
upload_year=upload_start.year; upload_month=upload_start.month
upload_key=f"{upload_year}-{upload_month:02d}"

_CURRENT_FY = get_fy_year(upload_year, upload_month)
_LAST_FY    = _CURRENT_FY - 1

# Load historical
with st.spinner("Loading master data..."):
    if master_file:
        sn=master_file.name if hasattr(master_file,"name") else str(master_file)
        if sn.lower().endswith(".csv"):
            hist_df=parse_master_from_daily_format(master_file)
            if hist_df.empty: master_file.seek(0); hist_df=pd.read_csv(master_file)
        else: hist_df=parse_master_excel(master_file)
    elif not folder_master_df.empty: hist_df=folder_master_df
    elif DEFAULT_MASTER and os.path.exists(DEFAULT_MASTER): hist_df=parse_master_excel(DEFAULT_MASTER)
    else: st.error("No master data available."); st.stop()

if hist_df.empty: st.error("Could not read master data."); st.stop()

# Detect CUSTOMER ID col in master
hist_cid_col=next((c for c in hist_df.columns if "CUSTOMER ID" in str(c).upper()), None)
if not hist_cid_col: st.error("CUSTOMER ID column not found in master."); st.stop()
hist_df["CLEAN_ID"]=hist_df[hist_cid_col].astype(str).str.replace(".0","",regex=False).str.strip()

# ── Determine base comparison year/month ─────────────────────────────────────
if comparison_base=="Previous month (same period)":
    cmp_mo=upload_month-1 if upload_month>1 else 12
    cmp_yr=upload_year    if upload_month>1 else upload_year-1
    cmp_label=f"{calendar.month_name[cmp_mo]} {cmp_yr} (previous month)"
elif comparison_base=="Last FY corresponding month":
    cmp_yr,cmp_mo=upload_year-1,upload_month
    cmp_label=f"{calendar.month_name[cmp_mo]} {cmp_yr} (last FY same month)"
else:  # Average
    cmp_yr,cmp_mo=None,None
    cmp_label="Average of all available months"

days_in_cmp=calendar.monthrange(cmp_yr,cmp_mo)[1] if cmp_yr else 30

# Check availability
base_rev_col=base_trf_col=None
if cmp_yr:
    base_rev_col=f"{cmp_yr}-{cmp_mo:02d} REVENUE" if f"{cmp_yr}-{cmp_mo:02d} REVENUE" in hist_df.columns else None
    base_trf_col=f"{cmp_yr}-{cmp_mo:02d} TRAFFIC" if f"{cmp_yr}-{cmp_mo:02d} TRAFFIC" in hist_df.columns else None

has_base=(base_rev_col is not None)
if cmp_yr and not has_base:
    if comparison_base=="Previous month (same period)":
        fb_label=f"current FY (FY {_CURRENT_FY}-{str(_CURRENT_FY+1)[-2:]})"
        fb_filter='current'
    else:
        fb_label=f"last FY (FY {_LAST_FY}-{str(_LAST_FY+1)[-2:]})"
        fb_filter='last'
    st.warning(f"⚠️ **{cmp_label}** not found in master. "
               f"{'Average fallback enabled — will use available months in ' + fb_label if use_average_fallback else 'Showing as No Historical Data.'}")

# ── Highest month analysis ────────────────────────────────────────────────────
high_last_yr,high_last_mo,high_last_lbl=None,None,None
high_curr_yr,high_curr_mo,high_curr_lbl=None,None,None
if highest_mode=="Highest in last FY":
    high_last_yr,high_last_mo,high_last_lbl=find_highest_month(hist_df,exclude_key=upload_key,fy_filter='last')
    if not high_last_yr: st.info("No data found in last FY for highest month comparison.")
elif highest_mode=="Highest in current FY (excl. uploaded period)":
    high_curr_yr,high_curr_mo,high_curr_lbl=find_highest_month(hist_df,exclude_key=upload_key,fy_filter='current')
    if not high_curr_yr: st.info("No data found in current FY for highest month comparison.")

# ── KPIs ──────────────────────────────────────────────────────────────────────
total_revenue  =pd.to_numeric(daily_df[rev_col],errors="coerce").sum()
total_traffic  =pd.to_numeric(daily_df[trf_col],errors="coerce").sum()
total_customers=daily_df[cid_col].nunique()

info_parts=[
    f"**Period:** {upload_start.strftime('%d %b %Y')} → {upload_end.strftime('%d %b %Y')}",
    f"**Days:** {uploaded_days}",
    f"**Base comparison:** {cmp_label}",
]
if high_last_lbl:  info_parts.append(f"**Highest last FY:** {high_last_lbl}")
if high_curr_lbl:  info_parts.append(f"**Highest current FY:** {high_curr_lbl}")
st.markdown(f"""<div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
border-radius:6px;margin-bottom:12px;font-size:15px;'>{"&nbsp;&nbsp;|&nbsp;&nbsp;".join(info_parts)}</div>""",
unsafe_allow_html=True)

c1,c2,c3=st.columns(3)
c1.metric("Total Revenue",  f"₹ {format_indian(total_revenue)}")
c2.metric("Total Traffic",  format_indian(total_traffic))
c3.metric("Total Customers",total_customers)
st.markdown("<hr style='margin:8px 0;border-color:#eee;'>",unsafe_allow_html=True)

use_avg_analysis=st.checkbox("Show average-based deep analysis for No Historical Data customers")

# ── Analytics loop ────────────────────────────────────────────────────────────
all_rev_cols=[c for c in hist_df.columns if c.endswith(" REVENUE")]
all_trf_cols=[c for c in hist_df.columns if c.endswith(" TRAFFIC")]

def avg_for_filter(hist_row, fy_filter):
    rvs=[]; tvs=[]
    for col in all_rev_cols:
        m=re.match(r'(\d{4})-(\d{2}) REVENUE',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        fy=get_fy_year(yr,mo)
        if fy_filter=='current' and fy!=_CURRENT_FY: continue
        if fy_filter=='last'    and fy!=_LAST_FY:    continue
        v=pd.to_numeric(hist_row[col],errors='coerce')
        if pd.notna(v) and v>0: rvs.append(v)
    for col in all_trf_cols:
        m=re.match(r'(\d{4})-(\d{2}) TRAFFIC',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        fy=get_fy_year(yr,mo)
        if fy_filter=='current' and fy!=_CURRENT_FY: continue
        if fy_filter=='last'    and fy!=_LAST_FY:    continue
        v=pd.to_numeric(hist_row[col],errors='coerce')
        if pd.notna(v) and v>0: tvs.append(v)
    return (np.mean(rvs) if rvs else None), (np.mean(tvs) if tvs else None), len(rvs), len(tvs)

results=[]; no_hist=[]; avg_results=[]

for _,row in daily_df.iterrows():
    cid  =str(row[cid_col]).replace(".0","").strip()
    cname=str(row[cname_col])
    rev  =pd.to_numeric(row[rev_col],errors="coerce")
    trf  =pd.to_numeric(row[trf_col],errors="coerce")
    hm   =hist_df[hist_df["CLEAN_ID"]==cid]

    if hm.empty and show_mode=="Only records present in Master Data": continue
    if hm.empty:
        no_hist.append({"Customer ID":cid,"Customer Name":cname,
                        "Actual Revenue":round(rev) if not pd.isna(rev) else 0,
                        "Actual Traffic":round(trf) if not pd.isna(trf) else 0,
                        "Revenue Status":"No Historical Data","Traffic Status":"No Historical Data"})
        continue

    hr=hm.iloc[0]

    # ── Base expected values ──────────────────────────────────────────────────
    exp_rev=exp_trf=None; base_used=cmp_label; base_months_rev=base_months_trf=None

    if comparison_base in ["Previous month (same period)","Last FY corresponding month"]:
        if has_base:
            mr=pd.to_numeric(hr.get(base_rev_col,0),errors='coerce'); mr=0 if pd.isna(mr) else mr
            mt=pd.to_numeric(hr.get(base_trf_col,0),errors='coerce'); mt=0 if pd.isna(mt) else mt
            if mr>0: exp_rev=(mr/days_in_cmp)*uploaded_days
            if mt>0: exp_trf=(mt/days_in_cmp)*uploaded_days
        # Fallback to FY average if base unavailable or zero
        if (exp_rev is None) and use_average_fallback:
            fb='current' if comparison_base=="Previous month (same period)" else 'last'
            avg_r,avg_t,nm_r,nm_t=avg_for_filter(hr,fb)
            if avg_r: exp_rev=(avg_r/30.44)*uploaded_days; base_months_rev=nm_r
            if avg_t: exp_trf=(avg_t/30.44)*uploaded_days; base_months_trf=nm_t
            base_used=(f"FY avg ({nm_r} months rev)" if nm_r else "No Historical Data")
    else:  # Average of all available months
        avg_r,avg_t,nm_r,nm_t=avg_for_filter(hr,None)
        if avg_r: exp_rev=(avg_r/30.44)*uploaded_days; base_months_rev=nm_r
        if avg_t: exp_trf=(avg_t/30.44)*uploaded_days; base_months_trf=nm_t
        base_used=f"Average ({nm_r} months)" if nm_r else "No Historical Data"

    rv,rs=compute_variance(rev,exp_rev,sd_percent)
    tv,ts=compute_variance(trf,exp_trf,sd_percent)

    rec={
        "Customer ID":cid,"Customer Name":cname,
        "Actual Revenue":round(rev) if not pd.isna(rev) else 0,
        "Expected Revenue":round(exp_rev) if exp_rev else "",
        "Revenue Variance %":rv if not pd.isna(rv) else "","Revenue Status":rs,
        "Actual Traffic":round(trf) if not pd.isna(trf) else 0,
        "Expected Traffic":round(exp_trf) if exp_trf else "",
        "Traffic Variance %":tv if not pd.isna(tv) else "","Traffic Status":ts,
        "Comparison Base":base_used,
    }

    # ── Highest month comparison ──────────────────────────────────────────────
    if highest_mode=="Highest in last FY" and high_last_yr:
        h_rk=f"{high_last_yr}-{high_last_mo:02d} REVENUE"; h_tk=f"{high_last_yr}-{high_last_mo:02d} TRAFFIC"
        h_days=calendar.monthrange(high_last_yr,high_last_mo)[1]
        hm_rev=pd.to_numeric(hr.get(h_rk,0),errors='coerce'); hm_rev=0 if pd.isna(hm_rev) else hm_rev
        hm_trf=pd.to_numeric(hr.get(h_tk,0),errors='coerce'); hm_trf=0 if pd.isna(hm_trf) else hm_trf
        exp_h_rev=(hm_rev/h_days)*uploaded_days if hm_rev else None
        exp_h_trf=(hm_trf/h_days)*uploaded_days if hm_trf else None
        hv,hs=compute_variance(rev,exp_h_rev,sd_percent)
        htv,hts=compute_variance(trf,exp_h_trf,sd_percent)
        rec[f"Exp Rev vs {high_last_lbl}"]=round(exp_h_rev) if exp_h_rev else ""
        rec[f"Rev Status vs {high_last_lbl}"]=hs
        rec[f"Exp Trf vs {high_last_lbl}"]=round(exp_h_trf) if exp_h_trf else ""
        rec[f"Trf Status vs {high_last_lbl}"]=hts

    elif highest_mode=="Highest in current FY (excl. uploaded period)" and high_curr_yr:
        h_rk=f"{high_curr_yr}-{high_curr_mo:02d} REVENUE"; h_tk=f"{high_curr_yr}-{high_curr_mo:02d} TRAFFIC"
        h_days=calendar.monthrange(high_curr_yr,high_curr_mo)[1]
        hm_rev=pd.to_numeric(hr.get(h_rk,0),errors='coerce'); hm_rev=0 if pd.isna(hm_rev) else hm_rev
        hm_trf=pd.to_numeric(hr.get(h_tk,0),errors='coerce'); hm_trf=0 if pd.isna(hm_trf) else hm_trf
        exp_h_rev=(hm_rev/h_days)*uploaded_days if hm_rev else None
        exp_h_trf=(hm_trf/h_days)*uploaded_days if hm_trf else None
        hv,hs=compute_variance(rev,exp_h_rev,sd_percent)
        htv,hts=compute_variance(trf,exp_h_trf,sd_percent)
        rec[f"Exp Rev vs {high_curr_lbl}"]=round(exp_h_rev) if exp_h_rev else ""
        rec[f"Rev Status vs {high_curr_lbl}"]=hs
        rec[f"Exp Trf vs {high_curr_lbl}"]=round(exp_h_trf) if exp_h_trf else ""
        rec[f"Trf Status vs {high_curr_lbl}"]=hts

    results.append(rec)

result_df =pd.DataFrame(results)
no_hist_df=pd.DataFrame(no_hist)
count_no_hist=len(no_hist)

# ── Display ───────────────────────────────────────────────────────────────────
st.subheader("Customer Analytics")

status_style_cols=[c for c in (result_df.columns if not result_df.empty else [])
                   if "Status" in c]

for status in STATUS_ORDER[:4]:
    if not result_df.empty:
        grp=result_df[result_df["Revenue Status"]==status]
        if not grp.empty:
            st.markdown(f"### {status} ({len(grp)})")
            st.dataframe(grp.style.map(color_status,subset=status_style_cols),
                         use_container_width=True,hide_index=True)

if count_no_hist>0:
    st.markdown(f"### No Historical Data ({count_no_hist})")
    st.dataframe(no_hist_df,use_container_width=True,hide_index=True)

# ── Average-based deep analysis ───────────────────────────────────────────────
if use_avg_analysis and count_no_hist>0:
    st.markdown("---"); st.subheader("Average-Based Analysis — No Historical Data Customers")
    avg_rows=[]
    for _,row in no_hist_df.iterrows():
        cid=str(row["Customer ID"]); cname=str(row["Customer Name"])
        rev_v=float(row["Actual Revenue"]); trf_v=float(row["Actual Traffic"])
        hm=hist_df[hist_df["CLEAN_ID"]==cid]
        if hm.empty: continue
        hr=hm.iloc[0]
        avg_r,avg_t,nm_r,nm_t=avg_for_filter(hr,None)
        if not avg_r and not avg_t: continue
        exp_r=(avg_r/30.44)*uploaded_days if avg_r else None
        exp_t=(avg_t/30.44)*uploaded_days if avg_t else None
        rv,rs=compute_variance(rev_v,exp_r,sd_percent)
        tv,ts=compute_variance(trf_v,exp_t,sd_percent)
        avg_rows.append({"Customer ID":cid,"Customer Name":cname,
                         "Actual Revenue":round(rev_v),"Expected Revenue":round(exp_r) if exp_r else "",
                         "Revenue Variance %":rv if not pd.isna(rv) else "","Revenue Status":rs,
                         "Actual Traffic":round(trf_v),"Expected Traffic":round(exp_t) if exp_t else "",
                         "Traffic Variance %":tv if not pd.isna(tv) else "","Traffic Status":ts,
                         "Months Averaged (Rev)":nm_r,"Months Averaged (Trf)":nm_t})
    avg_df=pd.DataFrame(avg_rows)
    if not avg_df.empty:
        for status in STATUS_ORDER:
            grp=avg_df[avg_df["Revenue Status"]==status]
            if not grp.empty:
                st.markdown(f"### {status} ({len(grp)})")
                st.dataframe(grp.style.map(color_status,subset=["Revenue Status","Traffic Status"]),
                             use_container_width=True,hide_index=True)

# ── Excel download ────────────────────────────────────────────────────────────
out=io.BytesIO()
all_df=(pd.concat([result_df,no_hist_df],ignore_index=True)
        if not no_hist_df.empty else result_df.copy())

with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
    fmts=_make_excel_formats(writer.book)
    if not all_df.empty:
        write_grouped_sheet(writer,all_df,"Customer Analytics",fmts)
    if use_avg_analysis and 'avg_df' in dir() and not avg_df.empty:
        write_grouped_sheet(writer,avg_df,"Avg-Based Analysis",fmts)

st.download_button("⬇ Download Excel Report",out.getvalue(),
                   file_name="analytics_report.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
