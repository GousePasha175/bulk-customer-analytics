import streamlit as st
import pandas as pd
import numpy as np
import calendar
import io
import os
import re
import glob as _glob
import datetime
from PIL import Image

# ==========================================
# PAGE CONFIGURATION & INTERFACE LOOK
# ==========================================
st.set_page_config(
    page_title="Analytics (Business and Operations)",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.block-container { padding-top: 1.2rem !important; padding-bottom: 0rem !important; }
header { visibility: hidden; height: 0px !important; }
.stTextInput input { height: 44px; border-radius: 8px; font-size: 16px; }
div.stButton > button {
    background-color: #ff4b4b; color: white; border: none;
    border-radius: 8px; width: 100%; height: 48px;
    font-size: 20px; font-weight: 600; margin-top: 8px;
}
div.stButton > button:hover { background-color: #e23d3d; color: white; }
.metric-card-header {
    background-color: #2f3343; color: white; padding: 10px; 
    border-top-left-radius: 8px; border-top-right-radius: 8px;
    font-weight: 600; font-size: 14px; text-align: center;
}
.metric-card-body {
    background-color: #f8f9fa; border: 1px solid #e9ecef;
    border-top: none; padding: 15px; text-align: center;
    border-bottom-left-radius: 8px; border-bottom-right-radius: 8px;
    font-size: 20px; font-weight: bold; color: #ff4b4b;
    min-height: 65px; display: flex; align-items: center; justify-content: center;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# HELPER DATA CONVERSION UTILITIES
# ==========================================
def format_indian(n):
    """Formats raw values into the standard Indian numbering system"""
    try: n = int(round(n))
    except (ValueError, TypeError): return str(n)
    is_negative = n < 0
    n = abs(n)
    s = str(n)
    if len(s) <= 3:
        result = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.append(rest[-2:])
            rest = rest[:-2]
        if rest: parts.append(rest)
        parts.reverse()
        result = ",".join(parts) + "," + last3
    return ("-" if is_negative else "") + result

def classify(variance, threshold):
    if pd.isna(variance): return "No Historical Data"
    if variance >= threshold: return "Excellent"
    elif variance >= 0: return "Normal"
    elif variance >= -threshold: return "Warning"
    else: return "Critical"

def color_status(val):
    colors = {
        "Excellent": "background-color: #90EE90; color: black;",
        "Normal": "background-color: #FFFACD; color: black;",
        "Warning": "background-color: #FFD580; color: black;",
        "Critical": "background-color: #FF7F7F; color: black;",
        "No Historical Data": "background-color: #D3D3D3; color: black;",
    }
    return colors.get(val, "")

# ==========================================
# CORE LEGACY HIERARCHICAL FILE PARSER
# ==========================================
def parse_master_dataframe(source_input, is_path=False):
    """
    Parses legacy multi-sheet transactional database dumps natively, 
    accounting for nested variable header blocks and complex layout formatting styles.
    """
    consolidated = {}
    short_months = [calendar.month_abbr[i].lower() for i in range(1, 13)]
    full_months = [calendar.month_name[i].lower() for i in range(1, 13)]
    
    try:
        if is_path:
            df_raw = pd.read_csv(source_input, nrows=5, header=None)
        else:
            df_raw = pd.read_csv(source_input, nrows=5, header=None)
            source_input.seek(0)

        is_iso_date = False
        for r_idx in [0, 1]:
            if r_idx < len(df_raw) and any(re.match(r'\d{4}-\d{2}-\d{2}', str(x).strip()) for x in df_raw.iloc[r_idx]):
                is_iso_date = True
                break

        if not is_path: source_input.seek(0)

        # Strategy A: ISO Multi-Header Rows Detection (Jeedimetla, Bollaram, BNPL Formats)
        if is_iso_date:
            top_headers = pd.read_csv(source_input, nrows=1, header=None).iloc[0].ffill().tolist()
            if not is_path: source_input.seek(0)
            sub_headers = pd.read_csv(source_input, skiprows=1, nrows=1, header=None).iloc[0].tolist()
            if not is_path: source_input.seek(0)
            df = pd.read_csv(source_input, skiprows=2, header=None)
            
            col_names = []
            for t, s in zip(top_headers, sub_headers):
                t_s, s_s = str(t).strip(), str(s).strip()
                if re.match(r'\d{4}-\d{2}-\d{2}', t_s):
                    col_names.append(f"{t_s[:7]} {s_s.upper()}")
                else:
                    col_names.append(s_s if s_s and s_s != "nan" else t_s)
            df.columns = col_names
            
            cid_col = [c for c in df.columns if "CUSTOMER ID" in str(c).upper() or "CUST ID" in str(c).upper()][0]
            for _, row in df.iterrows():
                raw_id = str(row[cid_col]).split('.')[0].strip()
                if raw_id.lower() in ('', 'nan', 'total', 'grand total'): continue
                if raw_id not in consolidated: consolidated[raw_id] = {"CUSTOMER ID": raw_id}
                
                for col in df.columns:
                    if "TRAFFIC" in str(col) or "REVENUE" in str(col):
                        parts = str(col).split()
                        if len(parts) >= 2 and re.match(r'\d{4}-\d{2}', parts[0]):
                            p_key = parts[0]
                            metric = parts[1].upper()
                            val = pd.to_numeric(row[col], errors='coerce') or 0
                            consolidated[raw_id][f"{p_key} {metric}"] = val
                            y, m = map(int, p_key.split('-'))
                            consolidated[raw_id][f"{p_key} DAYS"] = calendar.monthrange(y, m)[1]

        # Strategy B: Flat Text Attribute Matrix Layout (Autonagar Format)
        else:
            df = pd.read_csv(source_input)
            cid_col = None
            for c in df.columns:
                if str(c).strip().lower() in ["cust id", "customer id"]: cid_col = c
            if cid_col:
                for _, row in df.iterrows():
                    raw_id = str(row[cid_col]).split('.')[0].strip()
                    if raw_id.lower() in ('', 'nan', 'total', 'grand total'): continue
                    if raw_id not in consolidated: consolidated[raw_id] = {"CUSTOMER ID": raw_id}
                    
                    for col in df.columns:
                        col_s = str(col).strip()
                        m_match = re.search(r'([A-Za-z]{3})[- ]*(\d{2})', col_s)
                        if m_match:
                            m_nam, y_short = m_match.group(1).lower(), m_match.group(2)
                            if m_nam in short_months:
                                m_num = short_months.index(m_nam) + 1
                                p_key = f"20{y_short}-{m_num:02d}"
                                metric = "REVENUE" if "REV" in col_s.upper() else "TRAFFIC"
                                val = pd.to_numeric(row[col], errors='coerce') or 0
                                consolidated[raw_id][f"{p_key} {metric}"] = val
                                consolidated[raw_id][f"{p_key} DAYS"] = calendar.monthrange(int(f"20{y_short}"), m_num)[1]
    except Exception:
        pass
        
    return pd.DataFrame(list(consolidated.values()))

# ==========================================
# EXCEL SPREADSHEET WRITER ENGINE
# ==========================================
def write_grouped_sheet(writer, df, sheet_name, workbook, formats, status_col="Revenue Status"):
    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]
    header_format = workbook.add_format({"bold": True, "font_size": 12, "bg_color": "#2f3343", "font_color": "#FFFFFF", "border": 1})
    col_header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "border": 1})
    plain_fmt = workbook.add_format({"border": 1})

    ws = writer.book.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    cols = list(df.columns)
    current_row = 0

    for status in status_order:
        grp = df[df[status_col] == status]
        if grp.empty: continue

        ws.merge_range(current_row, 0, current_row, len(cols) - 1, f"{status} ({len(grp)})", header_format)
        current_row += 1
        for ci, col in enumerate(cols):
            ws.write(current_row, ci, col, col_header_fmt)
            ws.set_column(ci, ci, 20)
        current_row += 1

        rev_ci = cols.index("Revenue Status") if "Revenue Status" in cols else None
        trf_ci = cols.index("Traffic Status") if "Traffic Status" in cols else None

        for _, data_row in grp.iterrows():
            for ci, col in enumerate(cols):
                val = data_row[col]
                if isinstance(val, float) and np.isnan(val): val = ""
                cell_fmt = plain_fmt
                if ci in (rev_ci, trf_ci): cell_fmt = formats.get(str(val), plain_fmt)
                ws.write(current_row, ci, val, cell_fmt)
            current_row += 1
        current_row += 1

# ==========================================
# USER ROUTING SECURITY VERIFICATION
# ==========================================
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<style>[data-testid='stSidebar'] { display: none !important; }</style>", unsafe_allow_html=True)
    hl, hc, hr = st.columns([1, 3, 1])
    with hc:
        cl, cr = st.columns([1, 4])
        with cl:
            if logo: st.image(logo, width=100)
        with cr:
            st.markdown("<h1 style='font-size:24px; color:#2f3343;'>Analytics (Business & Operations)</h1><p>Telangana Circle</p>", unsafe_allow_html=True)
        with st.form("login"):
            st.text_input("Username", key="usr")
            st.text_input("Password", type="password", key="pwd")
            if st.form_submit_button("Submit", use_container_width=True):
                if st.session_state.usr == "admin" and st.session_state.pwd == "HQR@2026":
                    st.session_state.authenticated = True
                    st.rerun()
                else: st.error("Invalid credentials")
    st.stop()

# --- Main App Heading Canvas Canvas ---
hl, hc, hr = st.columns([1, 8, 1])
if logo: hl.image(logo, width=90)
hc.markdown("<h1 style='margin:0; color:#2f3343;'>Dynamic Customer Cross-Comparison Engine</h1>", unsafe_allow_html=True)
st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

# ==========================================
# SIDEBAR SETUP CONFIGURATION
# ==========================================
st.sidebar.header("Configuration Panel")

# Optional Custom Override File Tool
opt_master_file = st.sidebar.file_uploader("Upload Master Data File (Optional Override)", type=["csv"])
daily_file = st.sidebar.file_uploader("Upload Target Evaluation File (Mandatory CSV)", type=["csv"])

deviation_th = st.sidebar.slider("Acceptable Deviation %", 1, 50, 10)
show_mode = st.sidebar.radio("Records View Filter", ["All Records", "Only records matching Master"])
use_average_history = st.sidebar.checkbox("Fallback to Cumulative Averages for Missing Target Slices", value=True)

# Native Master Sync Pipeline
master_db = pd.DataFrame()
if opt_master_file:
    master_db = parse_master_dataframe(opt_master_file, is_path=False)
elif os.path.exists("master"):
    local_csv_logs = _glob.glob(os.path.join("master", "*.csv"))
    chunk_frames = []
    for f_path in local_csv_logs:
        if os.path.getsize(f_path) > 0:
            parsed_chunk = parse_master_dataframe(f_path, is_path=True)
            if not parsed_chunk.empty:
                chunk_frames.append(parsed_chunk)
    if chunk_frames:
        master_db = pd.concat(chunk_frames, ignore_index=True).drop_duplicates(subset=["CUSTOMER ID"])

if master_db.empty:
    st.sidebar.error("⚠️ Local master file directory data repository is empty.")
    st.info("Place historical snapshot logs inside your local 'master/' workspace directory to enable processing profiles.")
    st.stop()

# Chronological Parsing Subsystem
available_periods = sorted(list(set([c.split()[0] for c in master_db.columns if "REVENUE" in c])))
month_names_mapping = {f"{i:02d}": calendar.month_name[i] for i in range(1, 13)}

def format_period_label(p):
    try:
        y, m = p.split('-')
        return f"{month_names_mapping.get(m, m)} {y}"
    except: return p

# Strategic Temporal Targets Base Selectors
baseline_modes = [
    "Previous Year Corresponding Month",
    "Last Month (MoM Preceding Target Slice)",
    "2-Month Rolling Historical Average Window",
    "3-Month Rolling Historical Average Window",
    "4-Month Rolling Historical Average Window",
    "5-Month Rolling Historical Average Window",
    "Global Consolidated Database Average Day Matrix"
]
for period in available_periods:
    baseline_modes.append(f"Static Custom Timeline Snapshot: {format_period_label(period)}")

selected_baseline_strategy = st.sidebar.selectbox("Select Baseline Strategy Rule Target", baseline_modes)

# ==========================================
# MATHEMATICAL ANALYTICAL COMPUTATION PIPELINE
# ==========================================
if daily_file:
    daily_df = pd.read_csv(daily_file)
    
    # Target Parsing Verification Engine
    cid_col = cname_col = rev_col = traf_col = sd_col = ed_col = None
    for col in daily_df.columns:
        c_l = str(col).strip().lower()
        if "customer id" in c_l or "cust id" in c_l: cid_col = col
        elif "customer name" in c_l or "cutomer name" in c_l: cname_col = col
        elif "revenue" in c_l or "amount" in c_l: rev_col = col
        elif "traffic" in c_l or "article" in c_l: traf_col = col
        elif "start" in c_l: sd_col = col
        elif "end" in c_l: ed_col = col

    missing_keys = [k for k, v in [("ID", cid_col), ("Name", cname_col), ("Revenue", rev_col), ("Traffic", traf_col)] if v is None]
    if missing_keys:
        st.error(f"Target upload validation structure failure. Columns mapping error for attributes: {missing_keys}")
        st.stop()

    # Timeline Normalization Framework
    if sd_col and ed_col:
        u_start = pd.to_datetime(daily_df[sd_col].iloc[0], format="%d/%m/%Y", errors="coerce")
        u_end = pd.to_datetime(daily_df[ed_col].iloc[0], format="%d/%m/%Y", errors="coerce")
        if pd.notna(u_start) and pd.notna(u_end):
            u_days = (u_end - u_start).days + 1
            analysis_period_str = f"{u_start.strftime('%B %Y')}"
            target_range_str = f"{u_start.strftime('%d/%m/%Y')} to {
