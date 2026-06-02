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

        # Strategy A: ISO Multi-Header Rows Detection
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

        # Strategy B: Flat Text Attribute Matrix Layout
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

# ==========================================
# CORE WORKSPACE INTERFACE (AUTHENTICATED)
# ==========================================
hl, hc, hr = st.columns([1, 8, 1])
if logo: hl.image(logo, width=90)
hc.markdown("<h1 style='margin:0; color:#2f3343;'>Dynamic Customer Cross-Comparison Engine</h1>", unsafe_allow_html=True)
st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

# Side Panels Configurations
st.sidebar.header("Configuration Panel")
opt_master_file = st.sidebar.file_uploader("Upload Master Data File (Optional Override)", type=["csv"])
daily_file = st.sidebar.file_uploader("Upload Target Evaluation File (Mandatory CSV)", type=["csv"])

deviation_th = st.sidebar.slider("Acceptable Deviation %", 1, 50, 10)
show_mode = st.sidebar.radio("Records View Filter", ["All Records", "Only records matching Master"])
use_average_history = st.sidebar.checkbox("Fallback to Cumulative Averages for Missing Target Slices", value=True)

# Safe Lazy Evaluation Master Sync Pipeline
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

# Safe Guard-Rail to prevent interface crashes if folder asset files are missing
if master_db.empty:
    st.sidebar.error("⚠️ Local master file directory data repository is empty.")
    st.info("Please place your historical tracking database logs inside your local 'master/' directory folder, or upload a custom Master file override to begin.")
    st.stop()

# Chronological Parsing Subsystem
available_periods = sorted(list(set([c.split()[0] for c in master_db.columns if "REVENUE" in c])))
month_names_mapping = {f"{i:02d}": calendar.month_name[i] for i in range(1, 13)}

def format_period_label(p):
    try:
        y, m = p.split('-')
        return f"{month_names_mapping.get(m, m)} {y}"
    except: return p

# Strategic Temporal Targets Selectors
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

    # Timeline Normalization Engine
    u_days = 30
    analysis_period_str = "Custom Evaluation Base"
    target_range_str = "Variable Selection Framework"
    active_month_key = prev_year_month_key = preceding_month_key = None

    if sd_col and ed_col and len(daily_df) > 0:
        raw_start = str(daily_df[sd_col].iloc[0]).strip()
        raw_end = str(daily_df[ed_col].iloc[0]).strip()
        
        u_start = pd.to_datetime(raw_start, dayfirst=True, errors="coerce")
        u_end = pd.to_datetime(raw_end, dayfirst=True, errors="coerce")
        
        if pd.notna(u_start) and pd.notna(u_end):
            u_days = (u_end - u_start).days + 1
            analysis_period_str = f"{u_start.strftime('%B %Y')}"
            target_range_str = f"{u_start.strftime('%d/%m/%Y')} to {u_end.strftime('%d/%m/%Y')}"
            active_month_key = f"{u_start.year}-{u_start.month:02d}"
            prev_year_month_key = f"{u_start.year - 1}-{u_start.month:02d}"
            preceding_month_key = f"{u_start.year - 1}-12" if u_start.month == 1 else f"{u_start.year}-{u_start.month - 1:02d}"

    # KPI Interface Header Banner Panel
    st.markdown("### 📊 Operational Execution Summary Context")
    grid1, grid2, grid3, grid4 = st.columns(4)
    with grid1:
        st.markdown("<div class='metric-card-header'>ANALYSIS PERIOD</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-card-body'>{analysis_period_str}</div>", unsafe_allow_html=True)
    with grid2:
        st.markdown("<div class='metric-card-header'>TARGET PERIOD RANGE</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-card-body' style='font-size:16px; color:#2f3343;'>{target_range_str}</div>", unsafe_allow_html=True)
    with grid3:
        st.markdown("<div class='metric-card-header'>TARGET PERIOD AVERAGE NUMBER OF CALENDAR DAYS</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-card-body'>{u_days} Days</div>", unsafe_allow_html=True)
    with grid4:
        st.markdown("<div class='metric-card-header'>COMPARISON BASELINE STRATEGY LAYER</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-card-body' style='font-size:13px; color:#1a73e8;'>{selected_baseline_strategy.split(':')[0]}</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Cross Evaluation Engine Pipeline Loop Execution
    eval_records = []
    avg_records_pool = [] # Sub-pool tracking table
    master_db["CLEAN_ID"] = master_db["CUSTOMER ID"].astype(str).str.strip()

    for _, row in daily_df.iterrows():
        cust_id = str(row[cid_col]).split('.')[0].strip()
        if cust_id.lower() in ('', 'nan', 'total', 'grand total'): continue
        cust_name = row[cname_col] if cname_col and pd.notna(row[cname_col]) else "Unknown Account Profile"
        
        act_rev = pd.to_numeric(row[rev_col], errors='coerce') or 0
        act_trf = pd.to_numeric(row[traf_col], errors='coerce') or 0
        
        match_hist = master_db[master_db["CLEAN_ID"] == cust_id]
        if match_hist.empty:
            if show_mode == "All Records":
                eval_records.append({
                    "Customer ID": cust_id, "Customer Name": cust_name,
                    "Actual Revenue": round(act_rev), "Expected Revenue (Pro-Rata)": "", "Revenue Variance %": "", "Revenue Status": "No Historical Data",
                    "Actual Traffic": round(act_trf), "Expected Traffic (Pro-Rata)": "", "Traffic Variance %": "", "Traffic Status": "No Historical Data"
                })
            continue

        hist_row = match_hist.iloc[0]
        exp_rev = exp_trf = 0
        valid_comparison = False
        fallback_used = False
        target_keys_pool = []

        if "Previous Year Corresponding Month" in selected_baseline_strategy:
            if prev_year_month_key: target_keys_pool = [prev_year_month_key]
        elif "Last Month" in selected_baseline_strategy:
            if preceding_month_key: target_keys_pool = [preceding_month_key]
        elif "Rolling Historical Average Window" in selected_baseline_strategy:
            try:
                num_months = int(selected_baseline_strategy.split("-")[0])
                if active_month_key in available_periods:
                    curr_idx = available_periods.index(active_month_key)
                    start_idx = max(0, curr_idx - num_months)
                    target_keys_pool = available_periods[start_idx:curr_idx]
                else:
                    target_keys_pool = available_periods[-num_months:]
            except: pass
        elif "Static Custom Timeline Snapshot:" in selected_baseline_strategy:
            raw_lbl = selected_baseline_strategy.replace("Static Custom Timeline Snapshot: ", "").strip()
            for p in available_periods:
                if format_period_label(p) == raw_lbl:
                    target_keys_pool = [p]
                    break

        if target_keys_pool and not "Global Consolidated Database Average Day Matrix" in selected_baseline_strategy:
            r_slices = [pd.to_numeric(hist_row.get(f"{k} REVENUE"), errors='coerce') or 0 for k in target_keys_pool if f"{k} REVENUE" in master_db.columns]
            t_slices = [pd.to_numeric(hist_row.get(f"{k} TRAFFIC"), errors='coerce') or 0 for k in target_keys_pool if f"{k} TRAFFIC" in master_db.columns]
            d_slices = [pd.to_numeric(hist_row.get(f"{k} DAYS"), errors='coerce') or 30 for k in target_keys_pool if f"{k} DAYS" in master_db.columns]
            
            if sum(r_slices) > 0 or sum(t_slices) > 0:
                daily_rev_rates = [r / d if d > 0 else 0 for r, d in zip(r_slices, d_slices) if r > 0]
                daily_trf_rates = [t / d if d > 0 else 0 for t, d in zip(t_slices, d_slices) if t > 0]
                
                exp_rev = (np.mean(daily_rev_rates) if daily_rev_rates else 0) * u_days
                exp_trf = (np.mean(daily_trf_rates) if daily_trf_rates else 0) * u_days
                valid_comparison = True

        # Fallback Subroutine Strategy Matrix
        if not valid_comparison and (use_average_history or "Global Consolidated Database Average Day Matrix" in selected_baseline_strategy):
            rev_cols = [c for c in master_db.columns if "REVENUE" in c]
            trf_cols = [c for c in master_db.columns if "TRAFFIC" in c]
            
            r_vals = [pd.to_numeric(hist_row[c], errors='coerce') or 0 for c in rev_cols if (pd.to_numeric(hist_row[c], errors='coerce') or 0) > 0]
            t_vals = [pd.to_numeric(hist_row[c], errors='coerce') or 0 for c in trf_cols if (pd.to_numeric(hist_row[c], errors='coerce') or 0) > 0]
            
            if r_vals or t_vals:
                exp_rev = ((np.mean(r_vals) if r_vals else 0) / 30.44) * u_days
                exp_trf = ((np.mean(t_vals) if t_vals else 0) / 30.44) * u_days
                valid_comparison = True
                fallback_used = True

        if valid_comparison:
            r_var = (((act_rev - exp_rev) / exp_rev) * 100) if exp_rev > 0 else np.nan
            t_var = (((act_trf - exp_trf) / exp_trf) * 100) if exp_trf > 0 else np.nan
            
            rec = {
                "Customer ID": cust_id, "Customer Name": cust_name,
                "Actual Revenue": round(act_rev), "Expected Revenue (Pro-Rata)": round(exp_rev), 
                "Revenue Variance %": round(r_var, 2) if not pd.isna(r_var) else "", "Revenue Status": classify(r_var, deviation_th),
                "Actual Traffic": round(act_trf), "Expected Traffic (Pro-Rata)": round(exp_trf), 
                "Traffic Variance %": round(t_var, 2) if not pd.isna(t_var) else "", "Traffic Status": classify(t_var, deviation_th)
            }
            eval_records.append(rec)
            if fallback_used:
                avg_records_pool.append(rec)
        else:
            if show_mode == "All Records":
                eval_records.append({
                    "Customer ID": cust_id, "Customer Name": cust_name,
                    "Actual Revenue": round(act_rev), "Expected Revenue (Pro-Rata)": "", "Revenue Variance %": "", "Revenue Status": "No Historical Data",
                    "Actual Traffic": round(act_trf), "Expected Traffic (Pro-Rata)": "", "Traffic Variance %": "", "Traffic Status": "No Historical Data"
                })

    out_df = pd.DataFrame(eval_records)
    avg_history_df = pd.DataFrame(avg_records_pool) # CRITICAL FIX: Safe initialization to avoid NameErrors

    if out_df.empty:
        st.warning("No tracking performance values discovered matching criteria bounds filters.")
        st.stop()

    st.subheader("Performance Status Classifications Grid")
    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]
    
    for status in status_order:
        sub_grp = out_df[out_df["Revenue Status"] == status]
        if not sub_grp.empty:
            st.markdown(f"#### {status} Status Segment — ({len(sub_grp)} Customer Accounts Managed)")
            st.dataframe(
                sub_grp.style.applymap(color_status, subset=["Revenue Status", "Traffic Status"]),
                use_container_width=True, hide_index=True
            )

    # Excel Output Direct Spreadsheet Generation Block
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        status_formats = {
            "Excellent": workbook.add_format({"bg_color": "#90EE90", "border": 1}),
            "Normal": workbook.add_format({"bg_color": "#FFFACD", "border": 1}),
            "Warning": workbook.add_format({"bg_color": "#FFD580", "border": 1}),
            "Critical": workbook.add_format({"bg_color": "#FF7F7F", "border": 1}),
            "No Historical Data": workbook.add_format({"bg_color": "#D3D3D3", "border": 1})
        }
        
        # Sheet 1: Core Analysis
        write_grouped_sheet(writer, out_df, sheet_name="Cross Analysis Report", workbook=workbook, formats=status_formats)
        
        # Sheet 2: Average-Based Analysis Fallback Sheet (Conditional verification completely fixed)
        if use_average_history and not avg_history_df.empty:
            write_grouped_sheet(writer, avg_history_df, sheet_name="Avg-Based Analysis", workbook=workbook, formats=status_formats)

    st.download_button(
        label="⬇ Download Cross Comparison Analytical Report (Excel)",
        data=output.getvalue(),
        file_name=f"Cross_Comparison_Report_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.
