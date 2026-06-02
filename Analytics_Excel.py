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
# PAGE CONFIGURATION & INTERFACE LOOK (YOUR ORIGINAL)
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
                                metric = "REVENUE" if "REV" in col_s.upper() else "TRAFFIC
