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
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'></p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Home")
    for pat,lbl in [("pages/Bulk Analytics.py|pages/*[Bb]ulk*.py","\U0001f4ca Bulk Customer Analytics"),
                    ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","\U0001f4ee POSB Daily Report"),
                    ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py","\U0001f4bb Digital Transactions"),
                    ("pages/Delivery Productivity.py|pages/*[Dd]elivery*.py", "\U0001f4e6 Delivery Productivity")]:
        hits=[]
        for p in pat.split("|"): hits+=_glob.glob(p)
        if hits: st.sidebar.page_link(hits[0].replace("\\","/"),label=lbl)
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>",unsafe_allow_html=True)

from PIL import Image

st.set_page_config(page_title="Bulk Customer Analytics",page_icon="\U0001f4ca",layout="wide",initial_sidebar_state="expanded")
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
</style>""",unsafe_allow_html=True)

if not st.session_state.get("authenticated",False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar to log in.")
    with st.sidebar: _render_nav()
    st.stop()

# ── Constants ─────────────────────────────────────────────────────────────────
PALETTE={"title_bg":"#1F3864","title_fg":"#FFFFFF","header_bg":"#2E75B6","header_fg":"#FFFFFF",
         "sub_hdr_bg":"#9DC3E6","sub_hdr_fg":"#000000","excellent":"#70AD47","excellent_fg":"#FFFFFF",
         "normal":"#FFFF00","normal_fg":"#000000","warning":"#FFC000","warning_fg":"#000000",
         "critical":"#FF0000","critical_fg":"#FFFFFF","no_hist":"#D3D3D3","no_hist_fg":"#000000",
         "total_bg":"#FFF2CC","total_fg":"#000000"}
STATUS_ORDER=["Excellent","Normal","Warning","Critical","No Historical Data"]
# CSVs live in data/ folder (named "Bulk Month Year.csv")
DATA_FOLDER="data"

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_indian(n):
    try: n=int(round(n))
    except: return str(n)
    neg=n<0; n=abs(n); s=str(n)
    if len(s)<=3: return("-"if neg else"")+s
    last3=s[-3:]; rest=s[:-3]; parts=[]
    while len(rest)>2: parts.append(rest[-2:]); rest=rest[:-2]
    if rest: parts.append(rest)
    parts.reverse()
    return("-"if neg else"")+",".join(parts)+","+last3

def parse_dates(series):
    for fmt in ["%d-%m-%Y","%d/%m/%Y","%Y-%m-%d","%d-%m-%y","%d/%m/%y",
                "%d %b %Y","%d-%b-%Y","%m/%d/%Y"]:
        try: return pd.to_datetime(series,format=fmt,errors='raise')
        except: continue
    return pd.to_datetime(series,dayfirst=True,errors='coerce')

def classify(v,thr):
    if pd.isna(v): return "No Historical Data"
    if v>=thr: return "Excellent"
    if v>=0:   return "Normal"
    if v>=-thr:return "Warning"
    return "Critical"

def color_status(val):
    return {"Excellent":f"background-color:{PALETTE['excellent']};color:{PALETTE['excellent_fg']}",
            "Normal":f"background-color:{PALETTE['normal']};color:{PALETTE['normal_fg']}",
            "Warning":f"background-color:{PALETTE['warning']};color:{PALETTE['warning_fg']}",
            "Critical":f"background-color:{PALETTE['critical']};color:{PALETTE['critical_fg']}",
            "No Historical Data":f"background-color:{PALETTE['no_hist']};color:{PALETTE['no_hist_fg']}"}.get(val,"")

def detect_cols(df):
    cid=cn=rev=trf=sd=ed=None
    for col in df.columns:
        c=str(col).strip().lower()
        if "customer id" in c: cid=col
        elif "customer name" in c: cn=col
        elif "amount" in c or "revenue" in c: rev=col
        elif "article" in c or "traffic" in c: trf=col
        elif "start" in c: sd=col
        elif "end" in c: ed=col
    return cid,cn,rev,trf,sd,ed

def get_fy(year,month):
    return year if month>=4 else year-1

# ── Load all CSVs from data/ folder ─────────────────────────────────────────
def load_master():
    """
    Reads all CSVs from data/ folder (named 'Bulk Month Year.csv').
    Also falls back to master/ if data/ has none.
    Returns pivoted DataFrame keyed by CUSTOMER ID with columns YYYY-MM REVENUE, YYYY-MM TRAFFIC.
    """
    files=sorted(
        _glob.glob(f"{DATA_FOLDER}/[Bb]ulk*.csv") +
        _glob.glob(f"{DATA_FOLDER}/[Bb]ulk*.CSV") +
        _glob.glob("master/[Bb]ulk*.csv") +
        _glob.glob("master/[Bb]ulk*.CSV")
    )
    if not files: return pd.DataFrame(), []
    pieces=[]; periods=[]
    for fp in files:
        try:
            df=pd.read_csv(fp)
            cid,cn,rev,trf,sd,ed=detect_cols(df)
            if not all([cid,rev,sd]): continue
            df[sd]=parse_dates(df[sd])
            yr=int(df[sd].dt.year.iloc[0]); mo=int(df[sd].dt.month.iloc[0])
            df["_cid"]=df[cid].astype(str).str.replace(".0","",regex=False).str.strip()
            df["_cn"] =df[cn].astype(str) if cn else ""
            df["_rev"]=pd.to_numeric(df[rev],errors="coerce").fillna(0)
            df["_trf"]=pd.to_numeric(df[trf],errors="coerce").fillna(0) if trf else 0
            grp=df.groupby("_cid").agg(_cn=("_cn","first"),rev=("_rev","sum"),trf=("_trf","sum")).reset_index()
            grp["_yr"]=yr; grp["_mo"]=mo
            pieces.append(grp)
            periods.append((yr,mo))
        except Exception as e:
            continue
    if not pieces: return pd.DataFrame(), []
    combined=pd.concat(pieces,ignore_index=True)
    # Build pivoted master
    master={}
    name_map={}
    for _,row in combined.iterrows():
        cid=row["_cid"]; key=f"{int(row['_yr'])}-{int(row['_mo']):02d}"
        if cid not in master: master[cid]={"CUSTOMER ID":cid,"CUSTOMER NAME":str(row["_cn"])}
        master[cid][f"{key} REVENUE"]=round(float(row["rev"]),2)
        master[cid][f"{key} TRAFFIC"]=int(round(float(row["trf"])))
    df_out=pd.DataFrame(list(master.values()))
    labels=[f"{calendar.month_name[m]} {y}" for y,m in sorted(set(periods))]
    return df_out, labels

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
                if ('customer name' in s) and nm_ci is None: nm_ci=ci
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
            ma={'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,
                'oct':10,'nov':11,'dec':12,'jan':1,'feb':2,'mar':3}
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
                tk=f"{yr}-{mo:02d} TRAFFIC"; rk=f"{yr}-{mo:02d} REVENUE"
                tv=rv=0
                if cols['t'] is not None and cols['t']<len(row):
                    tv=pd.to_numeric(row.iloc[cols['t']],errors='coerce'); tv=0 if pd.isna(tv) else tv
                if cols['r'] is not None and cols['r']<len(row):
                    rv=pd.to_numeric(row.iloc[cols['r']],errors='coerce'); rv=0 if pd.isna(rv) else rv
                master[cid][tk]=master[cid].get(tk,0)+tv
                master[cid][rk]=master[cid].get(rk,0)+rv
    return pd.DataFrame(list(master.values())) if master else pd.DataFrame()

def _make_fmts(wb):
    def f(**kw):
        b={"border":1,"valign":"vcenter"}; b.update(kw); return wb.add_format(b)
    return {
        "grp_hdr": f(bold=True,font_size=12,bg_color=PALETTE["title_bg"],font_color=PALETTE["title_fg"]),
        "sub_hdr": f(bold=True,align="center",bg_color=PALETTE["sub_hdr_bg"],font_color=PALETTE["sub_hdr_fg"]),
        "excellent":f(bg_color=PALETTE["excellent"],font_color=PALETTE["excellent_fg"]),
        "normal":   f(bg_color=PALETTE["normal"],   font_color=PALETTE["normal_fg"]),
        "warning":  f(bg_color=PALETTE["warning"],  font_color=PALETTE["warning_fg"]),
        "critical": f(bg_color=PALETTE["critical"], font_color=PALETTE["critical_fg"]),
        "no_hist":  f(bg_color=PALETTE["no_hist"],  font_color=PALETTE["no_hist_fg"]),
        "plain":    f(), "plain_l": f(align="left"),
    }

def write_sheet(writer,df,name,fmts,status_col="Revenue Status"):
    wb=writer.book; ws=wb.add_worksheet(name); writer.sheets[name]=ws
    cols=list(df.columns); cr=0; cw=22
    st_cols=[i for i,c in enumerate(cols) if "Status" in c]
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
                if ci in st_cols:
                    cfmt=fmts.get(str(val).lower().replace(" ","_").replace("no_historical_data","no_hist"),fmts["plain"])
                ws.write(cr,ci,val,cfmt)
            cr+=1
        cr+=1

# ── Logo ──────────────────────────────────────────────────────────────────────
logo=Image.open("assets/logo.png") if os.path.exists("assets/logo.png") else None

# ── Header ────────────────────────────────────────────────────────────────────
h1,h2,_=st.columns([1,8,1])
with h1:
    if logo: st.image(logo,width=90)
with h2:
    st.markdown("""<h1 style='font-size:28px;margin-bottom:2px;color:#2f3343;font-weight:700;padding-top:4px;'>
Bulk Customer Business Analytics</h1>
<p style='font-size:15px;color:#555;margin-top:0;'>Headquarters Region – Telangana Postal Circle</p>""",
    unsafe_allow_html=True)
st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>",unsafe_allow_html=True)

# ── Scan data/ folder for monthly CSVs ───────────────────────────────────────
_fcsvs = sorted(
    _glob.glob("data/[Bb]ulk*.csv") + _glob.glob("data/[Bb]ulk*.CSV") +
    _glob.glob("master/[Bb]ulk*.csv") + _glob.glob("master/[Bb]ulk*.CSV")
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
_render_nav()
st.sidebar.header("Upload Files")
daily_file = st.sidebar.file_uploader("Daily / Period File (CSV)", type=["csv"])

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Upload Master Data (optional)")
st.sidebar.caption(
    "Accepts any Excel or CSV. App auto-detects columns for "
    "Customer ID/Name, Traffic, Revenue, and Period."
)
master_file = st.sidebar.file_uploader(
    "Upload Master File",
    type=["xlsx", "xls", "xlsm", "csv", "tsv"],
    help="Excel (any format) or CSV/TSV. Multi-sheet Excel supported."
)

if master_file:    st.sidebar.success(f"✅ Uploaded: {master_file.name}")
elif _fcsvs:       st.sidebar.success(f"📂 data/ folder: {len(_fcsvs)} monthly CSV(s) found")
else:              st.sidebar.warning("⚠️ No master data. Add Bulk CSVs to data/ folder or upload a file.")

sd_pct = st.sidebar.slider("Deviation % threshold", min_value=1, max_value=50, value=10)

st.sidebar.markdown("---")
st.sidebar.subheader("📊 Comparison Options")
st.sidebar.caption("Default: **Previous month**. Check below for more comparisons.")
cmp_last_fy = st.sidebar.checkbox(
    "Also compare: Last FY same month", value=False,
    help="E.g. May 2026 vs May 2025"
)
cmp_highest = st.sidebar.checkbox(
    "Also compare: Highest month per customer", value=False,
    help="Each customer's best month from all available data"
)
show_avg_deep = st.sidebar.checkbox(
    "Analyse No Historical Data using average", value=False,
    help="Last FY average first; falls back to all-months average if needed."
)


def _smart_detect_cols(df):
    """
    Flexible column detector — works on any file format.
    Returns (cid_col, name_col, rev_col, trf_col, start_col, end_col).
    """
    cid = cname = rev = trf = sd = ed = None
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    cid_kw   = ['customer id','cust id','customerid','custid','client id','acct no','account no']
    name_kw  = ['customer name','cust name','client name','cutomer name','customer  name',
                'name','party name','firm name']
    rev_kw   = ['total amount','total amt','revenue','amount','rev','turnover','value',
                'billing','invoice amount','net amount']
    trf_kw   = ['article count','articles','traffic','article','count','qty',
                'quantity','items','pieces','shipments','bookings']
    start_kw = ['start date','start','from date','period start','month','date',
                'booking date','period']
    end_kw   = ['end date','end','to date','period end']

    def first_match(kws, skip=None):
        for kw in kws:
            for col, cl in cols_lower.items():
                if col == skip: continue
                if kw in cl: return col
        return None

    cid   = first_match(cid_kw)
    cname = first_match(name_kw, skip=cid)
    rev   = first_match(rev_kw)
    trf   = first_match(trf_kw)
    sd    = first_match(start_kw)
    ed    = first_match(end_kw)

    # Fallback by dtype if rev/trf not found
    if rev is None or trf is None:
        num_cols = df.select_dtypes(include='number').columns.tolist()
        num_cols = [c for c in num_cols if c not in (cid, rev, trf)]
        if num_cols:
            means = {c: df[c].mean() for c in num_cols}
            sm = sorted(means, key=means.get, reverse=True)
            if rev is None and sm: rev = sm[0]
            if trf is None and len(sm) > 1: trf = sm[1]

    return cid, cname, rev, trf, sd, ed


def _read_any_file(uploaded_file):
    """Read any uploaded file — CSV/TSV or Excel (single or multi-sheet)."""
    name = uploaded_file.name.lower() if hasattr(uploaded_file, 'name') else ''
    uploaded_file.seek(0)
    if name.endswith('.csv') or name.endswith('.tsv'):
        sep = '\t' if name.endswith('.tsv') else ','
        try:
            return pd.read_csv(uploaded_file, sep=sep)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=None, engine='python')
    # Excel
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        if len(xl.sheet_names) == 1:
            return xl.parse(xl.sheet_names[0])
        # Multi-sheet: concat sheets that have recognisable ID+revenue cols
        pieces = []
        for sheet in xl.sheet_names:
            try:
                df_s = xl.parse(sheet)
                cid_, _, rev_, _, _, _ = _smart_detect_cols(df_s)
                if cid_ and rev_: pieces.append(df_s)
            except Exception:
                continue
        return pd.concat(pieces, ignore_index=True) if pieces else xl.parse(xl.sheet_names[0])
    except Exception:
        return None


def _pivot_to_master(raw_df):
    """
    Convert any flat/wide file into the standard pivoted master DataFrame
    with columns YYYY-MM REVENUE and YYYY-MM TRAFFIC keyed by CUSTOMER ID.
    Returns (pivoted_df, available_month_labels).
    """
    import re as _re
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(), []

    cid_c, cn_c, rev_c, trf_c, sd_c, ed_c = _smart_detect_cols(raw_df)

    # Case 1: Long format (each row = one customer + one period)
    if cid_c and rev_c and sd_c:
        raw_df = raw_df.copy()
        raw_df[sd_c] = parse_dates(raw_df[sd_c])
        raw_df["_yr"]  = raw_df[sd_c].dt.year
        raw_df["_mo"]  = raw_df[sd_c].dt.month
        raw_df["_cid"] = raw_df[cid_c].astype(str).str.replace(".0","",regex=False).str.strip()
        raw_df["_rev"] = pd.to_numeric(raw_df[rev_c], errors="coerce").fillna(0)
        raw_df["_trf"] = pd.to_numeric(raw_df[trf_c], errors="coerce").fillna(0) if trf_c else 0
        raw_df["_cn"]  = raw_df[cn_c].astype(str) if cn_c else ""

        grouped = (raw_df.groupby(["_cid","_yr","_mo"])
                   .agg(_cn=("_cn","first"), rev=("_rev","sum"), trf=("_trf","sum"))
                   .reset_index())
        name_map = raw_df.groupby("_cid")["_cn"].first().to_dict()

        master_d = {}
        for _, row in grouped.iterrows():
            cid  = row["_cid"]
            key  = f"{int(row['_yr'])}-{int(row['_mo']):02d}"
            if cid not in master_d:
                master_d[cid] = {"CUSTOMER ID": cid,
                                  "CUSTOMER NAME": str(name_map.get(cid,""))}
            master_d[cid][f"{key} REVENUE"] = row["rev"]
            master_d[cid][f"{key} TRAFFIC"] = row["trf"]

        result = pd.DataFrame(list(master_d.values()))
        all_mo = sorted({(int(r["_yr"]),int(r["_mo"])) for _,r in grouped.iterrows()})
        labels = [f"{calendar.month_name[m]} {y}" for y,m in all_mo]
        return result, labels

    # Case 2: Wide/pivot format (month columns)
    if cid_c:
        mon_map = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
                   'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
        master_d = {}; labels_set = set()
        for _, row in raw_df.iterrows():
            cid = str(row[cid_c]).replace(".0","").strip()
            try: int(cid)
            except: continue
            cname = str(row[cn_c]).strip() if cn_c and cn_c in raw_df.columns else ""
            if cname.lower() == 'nan': cname = ""
            if cid not in master_d:
                master_d[cid] = {"CUSTOMER ID": cid, "CUSTOMER NAME": cname}
            for col in raw_df.columns:
                if col in (cid_c, cn_c): continue
                cl = str(col).strip().lower()
                m = _re.search(r'([a-z]{3})[\s\-]?(\d{2,4})', cl)
                if not m:
                    m2 = _re.search(r'(\d{4})[\-/](\d{1,2})', cl)
                    if m2: yr, mo = int(m2.group(1)), int(m2.group(2))
                    else:  continue
                else:
                    mo  = mon_map.get(m.group(1))
                    if not mo: continue
                    yr_s = m.group(2)
                    yr   = 2000+int(yr_s) if len(yr_s)==2 else int(yr_s)
                key = f"{yr}-{mo:02d}"
                val = pd.to_numeric(row[col], errors="coerce")
                if pd.isna(val): continue
                # Classify traffic vs revenue
                if any(k in cl for k in ['traf','art','count','qty','piece','book']):
                    field = "TRAFFIC"
                elif any(k in cl for k in ['rev','amt','amount','val','turn','bill']):
                    field = "REVENUE"
                else:
                    field = "REVENUE" if float(val) > 5000 else "TRAFFIC"
                master_d[cid][f"{key} {field}"] = master_d[cid].get(f"{key} {field}", 0) + float(val)
                labels_set.add((yr, mo))

        result = pd.DataFrame(list(master_d.values()))
        labels = [f"{calendar.month_name[m]} {y}" for y,m in sorted(labels_set)]
        return result, labels

    return raw_df, []


# ── Load master ───────────────────────────────────────────────────────────────
hist_df = pd.DataFrame(); avail_months = []

with st.spinner("Loading master data..."):
    if master_file:
        raw = _read_any_file(master_file)
        if raw is not None and not raw.empty:
            cid_d, cn_d, rev_d, trf_d, sd_d, _ = _smart_detect_cols(raw)
            with st.sidebar.expander("🔍 Detected columns in uploaded file", expanded=True):
                st.markdown(
                    f"| Field | Detected Column |\n|---|---|\n"
                    f"| Customer ID | `{cid_d or 'not found'}` |\n"
                    f"| Customer Name | `{cn_d or 'not found'}` |\n"
                    f"| Revenue / Amount | `{rev_d or 'not found'}` |\n"
                    f"| Traffic / Articles | `{trf_d or 'not found'}` |\n"
                    f"| Start Date / Period | `{sd_d or 'not found'}` |\n"
                )
                st.caption("All columns: " + ", ".join(f"`{c}`" for c in raw.columns[:20]))
            hist_df, avail_months = _pivot_to_master(raw)
        if hist_df.empty:
            st.sidebar.error("Could not parse uploaded file. Check column names above.")
    elif _fcsvs:
        hist_df, avail_months = load_master()

if avail_months:
    st.sidebar.caption(f"Months loaded: {', '.join(avail_months)}")

# ── Gate on daily file ────────────────────────────────────────────────────────
if not daily_file:
    st.info("Please upload the Daily / Period CSV file in the sidebar to begin.")
    st.stop()
if hist_df.empty:
    st.info(
        "No master data loaded.  \n"
        "- Place monthly CSVs named **'Bulk Month Year.csv'** in the **data/** folder, or  \n"
        "- Upload a master file using the sidebar uploader above."
    )
    st.stop()



# ── Parse daily ───────────────────────────────────────────────────────────────
daily_df=pd.read_csv(daily_file)
cid_c,cn_c,rev_c,trf_c,sd_c,ed_c=detect_cols(daily_df)
miss=[n for n,v in [("Customer ID",cid_c),("Revenue/Amount",rev_c),("Start Date",sd_c)] if v is None]
if miss: st.error(f"Could not detect columns in daily file: {', '.join(miss)}"); st.stop()

daily_df[sd_c]=parse_dates(daily_df[sd_c])
if ed_c: daily_df[ed_c]=parse_dates(daily_df[ed_c])
upload_start=pd.to_datetime(daily_df[sd_c].iloc[0],errors="coerce")
upload_end  =pd.to_datetime(daily_df[ed_c].iloc[0],errors="coerce") if ed_c else upload_start
up_days=(upload_end-upload_start).days+1
up_yr=upload_start.year; up_mo=upload_start.month
up_key=f"{up_yr}-{up_mo:02d}"
cur_fy=get_fy(up_yr,up_mo); last_fy=cur_fy-1

# ── Clean master ID column ────────────────────────────────────────────────────
hist_cid_col=next((c for c in hist_df.columns if "CUSTOMER ID" in str(c).upper()),None)
if not hist_cid_col: st.error("CUSTOMER ID column not found in master."); st.stop()
hist_df["_CID"]=hist_df[hist_cid_col].astype(str).str.replace(".0","",regex=False).str.strip()

all_rev_cols=[c for c in hist_df.columns if c.endswith(" REVENUE")]
all_trf_cols=[c for c in hist_df.columns if c.endswith(" TRAFFIC")]

# ── Default comparison: PREVIOUS MONTH ───────────────────────────────────────
prev_mo   = up_mo-1 if up_mo>1 else 12
prev_yr   = up_yr   if up_mo>1 else up_yr-1
prev_rk   = f"{prev_yr}-{prev_mo:02d} REVENUE"
prev_tk   = f"{prev_yr}-{prev_mo:02d} TRAFFIC"
prev_label= f"{calendar.month_name[prev_mo]} {prev_yr}"
has_prev  = (prev_rk in hist_df.columns)
days_prev = calendar.monthrange(prev_yr, prev_mo)[1]

# ── Last FY same month ────────────────────────────────────────────────────────
lfy_yr    = up_yr - 1
lfy_mo    = up_mo
lfy_rk    = f"{lfy_yr}-{lfy_mo:02d} REVENUE"
lfy_tk    = f"{lfy_yr}-{lfy_mo:02d} TRAFFIC"
lfy_label = f"{calendar.month_name[lfy_mo]} {lfy_yr}"
has_lfy   = (lfy_rk in hist_df.columns)
days_lfy  = calendar.monthrange(lfy_yr, lfy_mo)[1]

# ── Variance helpers ──────────────────────────────────────────────────────────
def var(actual, expected):
    if expected is None or expected==0: return np.nan
    return int(round(((actual-expected)/expected)*100))

def expected_from_month(hist_row, rk, tk, period_days, month_days):
    """Scale monthly figure to uploaded period length."""
    rv=pd.to_numeric(hist_row.get(rk,0),errors='coerce'); rv=0 if pd.isna(rv) else rv
    tv=pd.to_numeric(hist_row.get(tk,0),errors='coerce'); tv=0 if pd.isna(tv) else tv
    exp_r=(rv/month_days)*period_days if rv>0 else None
    exp_t=(tv/month_days)*period_days if tv>0 else None
    return exp_r, exp_t

def avg_expected(hist_row, period_days, fy_filter=None):
    """Average across all months optionally filtered by FY."""
    rvs=[]; tvs=[]
    for col in all_rev_cols:
        m=re.match(r'(\d{4})-(\d{2}) REVENUE',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        if fy_filter=='last'    and get_fy(yr,mo)!=last_fy: continue
        if fy_filter=='current' and get_fy(yr,mo)!=cur_fy:  continue
        v=pd.to_numeric(hist_row.get(col,0),errors='coerce')
        if pd.notna(v) and v>0: rvs.append(v)
    for col in all_trf_cols:
        m=re.match(r'(\d{4})-(\d{2}) TRAFFIC',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        if fy_filter=='last'    and get_fy(yr,mo)!=last_fy: continue
        if fy_filter=='current' and get_fy(yr,mo)!=cur_fy:  continue
        v=pd.to_numeric(hist_row.get(col,0),errors='coerce')
        if pd.notna(v) and v>0: tvs.append(v)
    if not rvs and not tvs: return None, None, 0, 0
    avg_r=np.mean(rvs) if rvs else None; avg_t=np.mean(tvs) if tvs else None
    exp_r=(avg_r/30.44)*period_days if avg_r else None
    exp_t=(avg_t/30.44)*period_days if avg_t else None
    return exp_r, exp_t, len(rvs), len(tvs)

def highest_month_for(hist_row):
    """Find the month with highest revenue for this specific customer."""
    best_rev=-1; best_yr=best_mo=None
    for col in all_rev_cols:
        m=re.match(r'(\d{4})-(\d{2}) REVENUE',col)
        if not m: continue
        yr,mo=int(m.group(1)),int(m.group(2))
        if f"{yr}-{mo:02d}"==up_key: continue  # exclude uploaded period
        v=pd.to_numeric(hist_row.get(col,0),errors='coerce')
        if pd.notna(v) and float(v)>best_rev: best_rev=float(v); best_yr=yr; best_mo=mo
    return best_yr, best_mo

# ── KPIs ──────────────────────────────────────────────────────────────────────
total_rev  =pd.to_numeric(daily_df[rev_c],errors="coerce").sum()
total_trf  =pd.to_numeric(daily_df[trf_c],errors="coerce").sum() if trf_c else 0
total_custs=daily_df[cid_c].nunique()

# Info banner — use HTML <b> tags (markdown ** won't render inside html div)
banner_parts=[
    f"<b>Period:</b> {upload_start.strftime('%d %b %Y')} → {upload_end.strftime('%d %b %Y')}",
    f"<b>Days:</b> {up_days}",
    f"<b>Default comparison:</b> {prev_label} (previous month)",
]
if not has_prev:
    banner_parts.append(f"<span style='color:#c00;'>⚠ {prev_label} not in master — fallback to last FY average</span>")
if cmp_last_fy:
    if has_lfy: banner_parts.append(f"<b>+ Last FY same month:</b> {lfy_label}")
    else:       banner_parts.append(f"<span style='color:#c00;'>⚠ {lfy_label} not in master</span>")
if cmp_highest:    banner_parts.append("<b>+ Highest month per customer</b>")
st.markdown(f"""<div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
border-radius:6px;margin-bottom:12px;font-size:15px;'>{"&nbsp;&nbsp;|&nbsp;&nbsp;".join(banner_parts)}</div>""",
unsafe_allow_html=True)

c1,c2,c3=st.columns(3)
c1.metric("Total Revenue",  f"₹ {fmt_indian(total_rev)}")
c2.metric("Total Traffic",  fmt_indian(total_trf))
c3.metric("Total Customers",total_custs)
st.markdown("<hr style='margin:8px 0;border-color:#eee;'>",unsafe_allow_html=True)

# ── Analytics loop ────────────────────────────────────────────────────────────
results=[]; no_hist_list=[]

for _,row in daily_df.iterrows():
    cid =str(row[cid_c]).replace(".0","").strip()
    cnam=str(row[cn_c]) if cn_c else cid
    rev =pd.to_numeric(row[rev_c],errors="coerce"); rev=0 if pd.isna(rev) else rev
    trf =pd.to_numeric(row[trf_c],errors="coerce") if trf_c else 0; trf=0 if pd.isna(trf) else trf

    hm=hist_df[hist_df["_CID"]==cid]

    # ── No match in master at all ────────────────────────────────────
    if hm.empty:
        no_hist_list.append({"Customer ID":cid,"Customer Name":cnam,
            "Actual Revenue":round(rev),"Actual Traffic":round(trf),
            "Revenue Status":"No Historical Data","Traffic Status":"No Historical Data"})
        continue

    hr=hm.iloc[0]

    # ── Default: previous month; fallback to last-FY average ─────────────────
    # Use FIXED column names so the table structure is always consistent.
    # The actual source used is recorded in "Comparison Used".
    exp_r_def=exp_t_def=None; def_used=prev_label; def_src=prev_label

    if has_prev:
        exp_r_def,exp_t_def=expected_from_month(hr,prev_rk,prev_tk,up_days,days_prev)
        def_src=prev_label

    # If revenue OR traffic is still None, try last-FY average
    if exp_r_def is None and exp_t_def is None:
        exp_r_def,exp_t_def,n_r,n_t=avg_expected(hr,up_days,'last')
        if n_r or n_t:
            def_src=f"Last FY avg ({max(n_r,n_t)} mo)"

    # Ultimate fallback: all available months
    if exp_r_def is None and exp_t_def is None:
        exp_r_def,exp_t_def,n_r,n_t=avg_expected(hr,up_days,None)
        if n_r or n_t:
            def_src=f"All months avg ({max(n_r,n_t)} mo)"

    # Independent revenue/traffic fallbacks:
    # If only one of rev/trf is None, try to fill it separately
    if exp_r_def is None and exp_t_def is not None:
        tmp_r,_,nr2,_=avg_expected(hr,up_days,'last')
        if tmp_r is None: tmp_r,_,nr2,_=avg_expected(hr,up_days,None)
        exp_r_def=tmp_r
    if exp_t_def is None and exp_r_def is not None:
        _,tmp_t,_,nt2=avg_expected(hr,up_days,'last')
        if tmp_t is None: _,tmp_t,_,nt2=avg_expected(hr,up_days,None)
        exp_t_def=tmp_t

    # Variance — only compute when expected value exists
    rv_def = var(rev, exp_r_def) if exp_r_def else np.nan
    tv_def = var(trf, exp_t_def) if exp_t_def else np.nan
    rs_def = classify(rv_def, sd_pct)
    ts_def = classify(tv_def, sd_pct)

    # Fixed column names keyed to prev_label (always "Expected Rev (April 2026)" etc.)
    rec={"Customer ID":cid,"Customer Name":cnam,
         "Actual Revenue":round(rev),"Actual Traffic":round(trf),
         f"Expected Rev ({prev_label})":int(round(exp_r_def)) if exp_r_def else "",
         "Revenue Variance %":int(round(rv_def)) if not pd.isna(rv_def) else "",
         "Revenue Status":rs_def,
         f"Expected Trf ({prev_label})":int(round(exp_t_def)) if exp_t_def else "",
         "Traffic Variance %":int(round(tv_def)) if not pd.isna(tv_def) else "",
         "Traffic Status":ts_def,
         "Comparison Used":def_src}

    # ── Optional: last FY same month ──────────────────────────────────────────
    if cmp_last_fy:
        exp_r_lfy=exp_t_lfy=None; lfy_used=lfy_label
        if has_lfy:
            exp_r_lfy,exp_t_lfy=expected_from_month(hr,lfy_rk,lfy_tk,up_days,days_lfy)
        if exp_r_lfy is None:
            # Fallback: any last FY month average
            exp_r_lfy,exp_t_lfy,n_r,n_t=avg_expected(hr,up_days,'last')
            lfy_used=(f"Last FY avg ({n_r} mo)" if n_r else "No Historical Data")
        rv_lfy=var(rev,exp_r_lfy); tv_lfy=var(trf,exp_t_lfy)
        rec[f"Exp Rev ({lfy_label})"]=int(round(exp_r_lfy)) if exp_r_lfy else ""
        rec[f"Rev Var % ({lfy_label})"]=int(round(rv_lfy)) if not pd.isna(rv_lfy) else ""
        rec[f"Rev Status ({lfy_label})"]=classify(rv_lfy,sd_pct)
        rec[f"Exp Trf ({lfy_label})"]=int(round(exp_t_lfy)) if exp_t_lfy else ""
        rec[f"Trf Var % ({lfy_label})"]=int(round(tv_lfy)) if not pd.isna(tv_lfy) else ""
        rec[f"Trf Status ({lfy_label})"]=classify(tv_lfy,sd_pct)

    # ── Optional: highest month per customer ──────────────────────
    if cmp_highest:
        h_yr,h_mo=highest_month_for(hr)
        if h_yr:
            h_rk=f"{h_yr}-{h_mo:02d} REVENUE"; h_tk=f"{h_yr}-{h_mo:02d} TRAFFIC"
            h_days=calendar.monthrange(h_yr,h_mo)[1]
            h_lbl=f"{calendar.month_name[h_mo]} {h_yr}"
            exp_r_h,exp_t_h=expected_from_month(hr,h_rk,h_tk,up_days,h_days)
            rv_h=var(rev,exp_r_h); tv_h=var(trf,exp_t_h)
            rec[f"Highest month"]=h_lbl
            rec[f"Exp Rev (Highest)"]=int(round(exp_r_h)) if exp_r_h else ""
            rec[f"Rev Var % (Highest)"]=int(round(rv_h)) if not pd.isna(rv_h) else ""
            rec[f"Rev Status (Highest)"]=classify(rv_h,sd_pct)
            rec[f"Exp Trf (Highest)"]=int(round(exp_t_h)) if exp_t_h else ""
            rec[f"Trf Var % (Highest)"]=int(round(tv_h)) if not pd.isna(tv_h) else ""
            rec[f"Trf Status (Highest)"]=classify(tv_h,sd_pct)
        else:
            rec["Highest month"]="No data"

    results.append(rec)

result_df =pd.DataFrame(results)
no_hist_df=pd.DataFrame(no_hist_list)


def _fmt_val(val, col):
    """Format a single cell value for display — integers clean, variance 1dp."""
    if val == "" or val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    try:
        f = float(val)
        if "Variance" in col or (col.endswith("%") and "Status" not in col):
            return int(round(f))
        if any(k in col for k in ["Revenue","Traffic","Expected","Actual"]):
            return int(round(f))
        return val
    except (ValueError, TypeError):
        return val


def _clean_df(df):
    """
    Apply formatting cell-by-cell via applymap so dtype stays object
    (preserving empty strings) but displayed values are clean ints / 1dp floats.
    This avoids the Int64/column_config conflict that suppresses Styler colours.
    """
    if df.empty: return df
    df = df.copy()
    for col in df.columns:
        if "Status" in col or "Customer" in col or "Comparison" in col                 or "month" in col.lower() or "ID" in col:
            continue
        df[col] = df[col].apply(lambda v: _fmt_val(v, col))
    return df


def _style_status(val):
    """Cell-level colour for any Status column."""
    return color_status(val)


def _show_df(df, status_cols):
    """Display a cleaned DataFrame with colour coding on all Status columns."""
    df = _clean_df(df)
    # Detect ALL status columns present in this specific df
    active_status_cols = [c for c in df.columns if "Status" in c]
    styled = (
        df.style
          .map(_style_status, subset=active_status_cols)
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ── Status columns list ───────────────────────────────────────────────────────
status_cols=[c for c in (result_df.columns if not result_df.empty else []) if "Status" in c]

# ── Display ───────────────────────────────────────────────────────────────────
st.subheader("Customer Analytics")

# Merge any "No Historical Data" rows from result_df into no_hist_df
# (customers matched in master but with no usable data for any period)
if not result_df.empty:
    nhd_in_results = result_df[result_df["Revenue Status"] == "No Historical Data"]
    if not nhd_in_results.empty:
        # Keep only basic cols to unify with no_hist_df
        nhd_basic = nhd_in_results[["Customer ID","Customer Name",
                                     "Actual Revenue","Actual Traffic",
                                     "Revenue Status","Traffic Status"]].copy()
        no_hist_df = pd.concat([no_hist_df, nhd_basic], ignore_index=True)
        result_df  = result_df[result_df["Revenue Status"] != "No Historical Data"]

# Total count check
total_shown = (len(result_df) if not result_df.empty else 0) + len(no_hist_df)

for status in STATUS_ORDER[:4]:
    if not result_df.empty:
        grp = result_df[result_df["Revenue Status"] == status]
        if not grp.empty:
            st.markdown(f"### {status} ({len(grp)})")
            _show_df(grp, status_cols)

if not no_hist_df.empty:
    st.markdown(f"### No Historical Data ({len(no_hist_df)})")
    _show_df(no_hist_df, [])

# ── Average-based deep analysis ───────────────────────────────────────────────
if show_avg_deep and not no_hist_df.empty:
    st.markdown("---"); st.subheader("Average-Based Analysis — No Historical Data Customers")
    avg_rows=[]
    for _,row in no_hist_df.iterrows():
        cid=str(row["Customer ID"]); cnam=str(row["Customer Name"])
        rev_v=float(row["Actual Revenue"]); trf_v=float(row["Actual Traffic"])
        hm=hist_df[hist_df["_CID"]==cid]
        if hm.empty: continue
        # Try last FY average first; then fall back to all months
        exp_r,exp_t,n_r,n_t=avg_expected(hm.iloc[0],up_days,'last')
        if not exp_r:
            exp_r,exp_t,n_r,n_t=avg_expected(hm.iloc[0],up_days,None)
        if not exp_r and not exp_t: continue
        rv=var(rev_v,exp_r); tv=var(trf_v,exp_t)
        avg_rows.append({"Customer ID":cid,"Customer Name":cnam,
            "Actual Revenue":int(round(rev_v)),"Expected Revenue":int(round(exp_r)) if exp_r else "",
            "Revenue Variance %":int(round(rv)) if not pd.isna(rv) else "","Revenue Status":classify(rv,sd_pct),
            "Actual Traffic":int(round(trf_v)),"Expected Traffic":int(round(exp_t)) if exp_t else "",
            "Traffic Variance %":int(round(tv)) if not pd.isna(tv) else "","Traffic Status":classify(tv,sd_pct),
            "Months Avg (Rev)":n_r,"Months Avg (Trf)":n_t})
    avg_df=pd.DataFrame(avg_rows)
    if not avg_df.empty:
        avg_sc=[c for c in avg_df.columns if "Status" in c]
        for status in STATUS_ORDER:
            grp=avg_df[avg_df["Revenue Status"]==status]
            if not grp.empty:
                st.markdown(f"### {status} ({len(grp)})")
                _show_df(grp, avg_sc)

# ── Excel download ────────────────────────────────────────────────────────────
out=io.BytesIO()
all_df=pd.concat([result_df,no_hist_df],ignore_index=True) if not no_hist_df.empty else result_df.copy()
with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
    fmts=_make_fmts(writer.book)
    if not all_df.empty:
        write_sheet(writer,all_df,"Customer Analytics",fmts)
    if show_avg_deep and 'avg_df' in dir() and not avg_df.empty:
        write_sheet(writer,avg_df,"Avg-Based Analysis",fmts)
st.download_button("⬇ Download Excel Report",out.getvalue(),
    file_name="analytics_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
