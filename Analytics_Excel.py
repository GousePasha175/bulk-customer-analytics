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

# ==========================
# PAGE CONFIG & CSS
# ==========================
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
</style>
""", unsafe_allow_html=True)

# ==========================
# CORE PARSING & MAPPING LOGIC
# ==========================
def format_indian(n):
    """Format a number in Indian numbering system: 1,38,13,220"""
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
        "Excellent": "background-color: #90EE90",
        "Normal": "background-color: #FFFACD",
        "Warning": "background-color: #FFD580",
        "Critical": "background-color: #FF7F7F",
        "No Historical Data": "background-color: #D3D3D3",
    }
    return colors.get(val, "")

def parse_legacy_and_new_masters(folder_path="master"):
    """
    Scans the master directory, reading existing multi-formatted layout CSVs
    and newly dropped month-wise standalone reports, unifying into an indexable map.
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        
    csv_files = _glob.glob(os.path.join(folder_path, "*.csv"))
    consolidated = {}
    
    short_months = [calendar.month_abbr[i].lower() for i in range(1, 13)]
    full_months = [calendar.month_name[i].lower() for i in range(1, 13)]
    
    for file_path in csv_files:
        file_name = os.path.basename(file_path)
        try:
            # Check if it's a dynamic monthly dump file (e.g. "April 2026.csv")
            match = re.search(r"([A-Za-z]+)\s+(\d{4})", file_name)
            
            # --- STRATEGY A: STANDALONE MONTH DUMP FILE ---
            if match and not any(k in file_name for k in ["BPC", "BNPL", "PBC", "Master"]):
                df = pd.read_csv(file_path)
                m_name, y_str = match.group(1).lower(), match.group(2)
                m_idx = full_months.index(m_name) + 1 if m_name in full_months else (short_months.index(m_name) + 1 if m_name in short_months else None)
                if not m_idx: continue
                period_key = f"{y_str}-{m_idx:02d}"
                
                cid_col = r_col = t_col = None
                for c in df.columns:
                    cl = str(c).strip().lower()
                    if "customer id" in cl or "cust id" in cl: cid_col = c
                    elif "revenue" in cl or "actual revenue" in cl or "amount" in cl: r_col = c
                    elif "traffic" in cl or "actual traffic" in cl or "article" in cl: t_col = c
                
                if cid_col and r_col and t_col:
                    days_in_month = calendar.monthrange(int(y_str), m_idx)[1]
                    for _, row in df.iterrows():
                        raw_id = str(row[cid_col]).split('.')[0].strip()
                        if raw_id.lower() in ('', 'nan', 'total', 'grand total'): continue
                        if raw_id not in consolidated: consolidated[raw_id] = {"CUSTOMER ID": raw_id}
                        
                        r_val = pd.to_numeric(row[r_col], errors='coerce') or 0
                        t_val = pd.to_numeric(row[t_col], errors='coerce') or 0
                        consolidated[raw_id][f"{period_key} REVENUE"] = r_val
                        consolidated[raw_id][f"{period_key} TRAFFIC"] = t_val
                        consolidated[raw_id][f"{period_key} DAYS"] = days_in_month
                continue

            # --- STRATEGY B: LEGACY STRUCTURE MATRIX RECOGNITION ---
            sample = pd.read_csv(file_path, nrows=5, header=None)
            
            # Sub-case B1 & B2: Multi-row headers with ISO Dates (Jeedimetla, Bollaram, BNPL)
            is_iso_date = False
            for r_idx in [0, 1]:
                if r_idx < len(sample) and any(re.match(r'\d{4}-\d{2}-\d{2}', str(x).strip()) for x in sample.iloc[r_idx]):
                    is_iso_date = True
                    date_row_idx = r_idx
                    break
            
            if is_iso_date:
                # Re-parse headers accurately
                top_headers = pd.read_csv(file_path, nrows=1, header=None).iloc[0].ffill().tolist()
                sub_headers = pd.read_csv(file_path, skiprows=1, nrows=1, header=None).iloc[0].tolist()
                df = pd.read_csv(file_path, skiprows=2, header=None)
                
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

            # Sub-case B3: Autonagar Flat Style ('Apr-25 Traf', 'Apr-25 Rev')
            else:
                df = pd.read_csv(file_path)
                cid_col = None
                for c in df.columns:
                    if str(c).strip().lower() in ["cust id", "customer id"]: cid_col = c
                if not cid_col: continue
                
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
            continue
            
    if not consolidated: return pd.DataFrame()
    return pd.DataFrame(list(consolidated.values()))


# ==========================
# FILE EXCEL EXPORT WORKER
# ==========================
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

# ==========================
# MAIN STREAMLIT APPLICATION
# ==========================
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

# --- App Canvas Header ---
hl, hc, hr = st.columns([1, 8, 1])
if logo: hl.image(logo, width=90)
hc.markdown("<h1 style='margin:0; color:#2f3343;'>Dynamic Customer Cross-Comparison Engine</h1>", unsafe_allow_html=True)
st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

# Compile Repositories
with st.spinner("Indexing Master Data Repositories..."):
    master_db = parse_legacy_and_new_masters("master")

# Sidebar Setup
st.sidebar.header("Configuration Setup")
daily_file = st.sidebar.file_uploader("Upload Target Evaluation File (CSV)", type=["csv"])

deviation_th = st.sidebar.slider("Acceptable Deviation %", 1, 50, 10)
show_mode = st.sidebar.radio("Records View Filter", ["All Records", "Only records matching Master"])

if master_db.empty:
    st.sidebar.error("⚠️ Master Database folder ('master/') has no readable history files.")
    st.info("Please turn the multi-sheet file layers into CSV formats and save them inside the 'master/' directory folder.")
    st.stop()

# Extract unique list of historical timelines available across profiles
available_periods = sorted(list(set([c.split()[0] for c in master_db.columns if "REVENUE" in c])))
month_names_mapping = {f"{i:02d}": calendar.month_name[i] for i in range(1, 13)}

def format_period_label(p):
    try:
        y, m = p.split('-')
        return f"{month_names_mapping.get(m, m)} {y}"
    except:
        return p

baseline_options = ["Cumulative History Average"] + [format_period_label(p) for p in available_periods]
selected_baseline_lbl = st.sidebar.selectbox("Select Comparison Baseline", baseline_options)

# ================================================================
# CROSS MATRIX COMPARISON EXECUTION ENGINE
# ================================================================
if daily_file:
    daily_df = pd.read_csv(daily_file)
    
    # Detect target evaluation file column headers
    cid_col = cname_col = rev_col = traf_col = sd_col = ed_col = None
    for col in daily_df.columns:
        c_l = str(col).strip().lower()
        if "customer id" in c_l or "cust id" in c_l: cid_col = col
        elif "customer name" in c_l or "cutomer name" in c_l: cname_col = col
        elif "revenue" in c_l or "amount" in c_l: rev_col = col
        elif "traffic" in c_l or "article" in c_l: traf_col = col
        elif "start" in c_l: sd_col = col
        elif "end" in c_l: ed_col = col

    missing = [k for k, v in [("ID", cid_col), ("Name", cname_col), ("Revenue", rev_col), ("Traffic", traf_col)] if v is None]
    if missing:
        st.error(f"Target upload missing parameters: {missing}")
        st.stop()

    # Resolve uploaded timeline range properties
    if sd_col and ed_col:
        u_start = pd.to_datetime(daily_df[sd_col].iloc[0], format="%d/%m/%Y", errors="coerce")
        u_end = pd.to_datetime(daily_df[ed_col].iloc[0], format="%d/%m/%Y", errors="coerce")
        if pd.notna(u_start) and pd.notna(u_end):
            u_days = (u_end - u_start).days + 1
            target_range_str = f"{u_start.strftime('%d %b %Y')} to {u_end.strftime('%d %b %Y')} ({u_days} Days)"
        else:
            u_days = 30
            target_range_str = "Custom Month Range (Assumed 30-Day Pro-Rata Base)"
    else:
        u_days = 30
        target_range_str = "Custom Range (Assumed 30-Day Pro-Rata Base)"

    eval_records = []
    master_db["CLEAN_ID"] = master_db["CUSTOMER ID"].astype(str).str.strip()

    for _, row in daily_df.iterrows():
        cust_id = str(row[cid_col]).split('.')[0].strip()
        if cust_id.lower() in ('', 'nan', 'total', 'grand total'): continue
        cust_name = row[cname_col] if cname_col and pd.notna(row[cname_col]) else "Unknown Account"
        
        act_rev = pd.to_numeric(row[rev_col], errors='coerce') or 0
        act_trf = pd.to_numeric(row[traf_col], errors='coerce') or 0
        
        match_hist = master_db[master_db["CLEAN_ID"] == cust_id]
        if match_hist.empty:
            if show_mode == "All Records":
                eval_records.append({
                    "Customer ID": cust_id, "Customer Name": cust_name,
                    "Actual Rev": round(act_rev), "Expected Rev (Pro-Rata)": "", "Rev Variance %": "", "Revenue Status": "No Historical Data",
                    "Actual Trf": round(act_trf), "Expected Trf (Pro-Rata)": "", "Trf Variance %": "", "Traffic Status": "No Historical Data"
                })
            continue
            
        hist_row = match_hist.iloc[0]
        exp_rev = exp_trf = 0
        valid_comparison = False
        
        # Scenario A: Comparing against a specific historical slice chosen in select box
        if selected_baseline_lbl != "Cumulative History Average":
            sel_idx = baseline_options.index(selected_baseline_lbl) - 1
            target_p_key = available_periods[sel_idx]
            
            m_rev = pd.to_numeric(hist_row.get(f"{target_p_key} REVENUE"), errors='coerce') or 0
            m_trf = pd.to_numeric(hist_row.get(f"{target_p_key} TRAFFIC"), errors='coerce') or 0
            m_days = pd.to_numeric(hist_row.get(f"{target_p_key} DAYS"), errors='coerce') or 30
            
            if m_rev > 0 or m_trf > 0:
                exp_rev = (m_rev / m_days) * u_days
                exp_trf = (m_trf / m_days) * u_days
                valid_comparison = True
        
        # Scenario B: Cumulative Historical Global Averages Mode
        else:
            rev_cols = [c for c in master_db.columns if "REVENUE" in c]
            trf_cols = [c for c in master_db.columns if "TRAFFIC" in c]
            
            r_vals = [pd.to_numeric(hist_row[c], errors='coerce') or 0 for c in rev_cols if (pd.to_numeric(hist_row[c], errors='coerce') or 0) > 0]
            t_vals = [pd.to_numeric(hist_row[c], errors='coerce') or 0 for c in trf_cols if (pd.to_numeric(hist_row[c], errors='coerce') or 0) > 0]
            
            if r_vals or t_vals:
                avg_m_rev = np.mean(r_vals) if r_vals else 0
                avg_m_trf = np.mean(t_vals) if t_vals else 0
                exp_rev = (avg_m_rev / 30.44) * u_days
                exp_trf = (avg_m_trf / 30.44) * u_days
                valid_comparison = True

        if valid_comparison:
            r_var = (((act_rev - exp_rev) / exp_rev) * 100) if exp_rev > 0 else np.nan
            t_var = (((act_trf - exp_trf) / exp_trf) * 100) if exp_trf > 0 else np.nan
            
            eval_records.append({
                "Customer ID": cust_id, "Customer Name": cust_name,
                "Actual Rev": round(act_rev), "Expected Rev (Pro-Rata)": round(exp_rev), 
                "Rev Variance %": round(r_var, 2) if not pd.isna(r_var) else "", "Revenue Status": classify(r_var, deviation_th),
                "Actual Trf": round(act_trf), "Expected Trf (Pro-Rata)": round(exp_trf), 
                "Trf Variance %": round(t_var, 2) if not pd.isna(t_var) else "", "Traffic Status": classify(t_var, deviation_th)
            })
        else:
            if show_mode == "All Records":
                eval_records.append({
                    "Customer ID": cust_id, "Customer Name": cust_name,
                    "Actual Rev": round(act_rev), "Expected Rev (Pro-Rata)": "", "Rev Variance %": "", "Revenue Status": "No Historical Data",
                    "Actual Trf": round(act_trf), "Expected Trf (Pro-Rata)": "", "Trf Variance %": "", "Traffic Status": "No Historical Data"
                })

    out_df = pd.DataFrame(eval_records)
    
    if out_df.empty:
        st.warning("No comparative metrics matched to display.")
        st.stop()

    # Informative Context Banner
    st.markdown(f"""
    <div style='background:#f0f7ff; border-left:4px solid #1a73e8; padding:10px 16px; border-radius:6px; margin-bottom:15px; font-size: 15px;'>
    📊 <b>Active Target Selection:</b> {target_range_str} <br>
    🔄 <b>Comparison Baseline Standard:</b> {selected_baseline_lbl} (Pro-rata scaled based on operational days)
    </div>
    """, unsafe_allow_html=True)

    # Main Visual Output Categories
    st.subheader("Comparative Variance Summary")
    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]
    
    for status in status_order:
        sub_grp = out_df[out_df["Revenue Status"] == status]
        if not sub_grp.empty:
            st.markdown(f"#### {status} Status Group — ({len(sub_grp)} Accounts)")
            st.dataframe(
                sub_grp.style.map(color_status, subset=["Revenue Status", "Traffic Status"]),
                use_container_width=True, hide_index=True
            )

    # Multi-Sheet Excel compilation download
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
        write_grouped_sheet(writer, out_df, sheet_name="Cross Analysis Report", workbook=workbook, formats=status_formats)

    st.download_button(
        label="⬇ Download Comparison Report (Excel)",
        data=output.getvalue(),
        file_name=f"Cross_Comparison_Report_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload active target period evaluation file to initialize analytics pipeline.")
