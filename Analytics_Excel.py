import streamlit as st
import pandas as pd
import numpy as np
import calendar
import io
import os
import re
import glob as _glob
from PIL import Image

# ==========================
# PAGE CONFIG
# ==========================
st.set_page_config(
    page_title="Analytics (Business and Operations)",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================
# CUSTOM CSS
# ==========================
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

/* ── Sidebar collapse fix ──────────────────────────────────────────
   Keep the re-open arrow always visible and on-screen.
   Without this, after collapsing, the button shifts off-screen
   and there is no way to reopen the sidebar without clearing cache. */
[data-testid="collapsedControl"] {
    display: flex !important;
    visibility: visible !important;
    position: fixed !important;
    top: 50% !important;
    left: 0px !important;
    transform: translateY(-50%) !important;
    z-index: 999999 !important;
    background-color: #2f3343 !important;
    border-radius: 0 8px 8px 0 !important;
    padding: 10px 6px !important;
    box-shadow: 2px 2px 6px rgba(0,0,0,0.3) !important;
}
[data-testid="collapsedControl"] svg {
    fill: white !important;
}
/* ──────────────────────────────────────────────────────────────── */
</style>
""", unsafe_allow_html=True)

# ==========================
# HELPER FUNCTIONS
# ==========================
def format_indian(n):
    """Format a number in Indian numbering system: 1,38,13,220"""
    try:
        n = int(round(n))
    except (ValueError, TypeError):
        return str(n)
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
        if rest:
            parts.append(rest)
        parts.reverse()
        result = ",".join(parts) + "," + last3
    return ("-" if is_negative else "") + result


def parse_dates_robust(series):
    """
    Try multiple explicit date formats to ensure consistent DD/MM/YYYY parsing.
    Falls back to dayfirst=True only as a last resort.
    """
    formats_to_try = [
        "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y",
        "%d %m %Y", "%d %b %Y", "%d-%b-%Y", "%d/%b/%Y",
        "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y",
    ]
    for fmt in formats_to_try:
        try:
            parsed = pd.to_datetime(series, format=fmt, errors='raise')
            return parsed
        except Exception:
            continue
    # Last resort — let pandas infer but strongly hint dayfirst
    return pd.to_datetime(series, dayfirst=True, errors='coerce')


def classify(variance, threshold):
    if pd.isna(variance):
        return "No Historical Data"
    if variance >= threshold:
        return "Excellent"
    elif variance >= 0:
        return "Normal"
    elif variance >= -threshold:
        return "Warning"
    else:
        return "Critical"


def color_status(val):
    colors = {
        "Excellent":          "background-color: #90EE90",
        "Normal":             "background-color: #FFFACD",
        "Warning":            "background-color: #FFD580",
        "Critical":           "background-color: #FF7F7F",
        "No Historical Data": "background-color: #D3D3D3",
    }
    return colors.get(val, "")


def parse_master_excel(file):
    xl = pd.ExcelFile(file)
    master = {}

    for sheet in xl.sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet, header=None)

        id_col_idx   = None
        name_col_idx = None

        has_sub_header = any(
            str(v).strip().upper() == 'TRAFFIC'
            for v in raw.iloc[1].tolist()
        )

        data_start = 2 if has_sub_header else 1

        for scan in range(min(3, len(raw))):
            for cidx, val in enumerate(raw.iloc[scan]):
                s = str(val).strip().lower()
                if ('customer id' in s or 'cust id' in s) and id_col_idx is None:
                    id_col_idx = cidx
                if ('customer name' in s or 'cutomer name' in s) and name_col_idx is None:
                    name_col_idx = cidx

        if id_col_idx is None:
            continue

        month_map = {}

        if has_sub_header:
            date_row = raw.iloc[0]
            sub_row  = raw.iloc[1]
            current_date = None
            for cidx in range(len(date_row)):
                val = date_row.iloc[cidx]
                if hasattr(val, 'year'):
                    current_date = (val.year, val.month)
                    month_map[current_date] = {'t': None, 'r': None}
                if current_date is not None:
                    sub = str(sub_row.iloc[cidx]).strip().upper()
                    if sub == 'TRAFFIC':
                        month_map[current_date]['t'] = cidx
                    elif sub in ('REVENUE', 'REV'):
                        month_map[current_date]['r'] = cidx
        else:
            month_abbr = {
                'apr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9,
                'oct': 10, 'nov': 11, 'dec': 12, 'jan': 1, 'feb': 2, 'mar': 3
            }
            for cidx, val in enumerate(raw.iloc[0].tolist()):
                s = str(val).strip().lower()
                for abbr, mo in month_abbr.items():
                    if abbr in s:
                        yr_match = re.search(r'(\d{2})', s)
                        if yr_match:
                            yr = 2000 + int(yr_match.group(1))
                            key = (yr, mo)
                            if key not in month_map:
                                month_map[key] = {'t': None, 'r': None}
                            if 'traf' in s or 'trf' in s:
                                month_map[key]['t'] = cidx
                            elif 'rev' in s:
                                month_map[key]['r'] = cidx
                        break

        if not month_map:
            continue

        for ridx in range(data_start, len(raw)):
            row = raw.iloc[ridx]
            cid_val = row.iloc[id_col_idx]

            if pd.isna(cid_val):
                continue
            s = str(cid_val).strip().lower()
            if s in ('', 'nan', 'total', 'grand total', 'sl no', 'sno'):
                continue

            clean_id = str(cid_val).replace('.0', '').strip()
            cname = ''
            if name_col_idx is not None and name_col_idx < len(row):
                cname = str(row.iloc[name_col_idx]).strip()
                if cname.lower() in ('nan', ''):
                    cname = ''

            if clean_id not in master:
                master[clean_id] = {'CUSTOMER ID': clean_id, 'CUSTOMER NAME': cname}

            for (yr, mo), cols in month_map.items():
                t_key = f"{yr}-{mo:02d} TRAFFIC"
                r_key = f"{yr}-{mo:02d} REVENUE"
                t_val = 0
                r_val = 0
                if cols['t'] is not None and cols['t'] < len(row):
                    t_val = pd.to_numeric(row.iloc[cols['t']], errors='coerce')
                    t_val = 0 if pd.isna(t_val) else t_val
                if cols['r'] is not None and cols['r'] < len(row):
                    r_val = pd.to_numeric(row.iloc[cols['r']], errors='coerce')
                    r_val = 0 if pd.isna(r_val) else r_val
                master[clean_id][t_key] = master[clean_id].get(t_key, 0) + t_val
                master[clean_id][r_key] = master[clean_id].get(r_key, 0) + r_val

    return pd.DataFrame(list(master.values())) if master else pd.DataFrame()


def detect_daily_columns(df):
    """Detect standard columns in a daily/period format CSV."""
    customer_id_col = customer_name_col = revenue_col = None
    traffic_col = start_date_col = end_date_col = None

    for col in df.columns:
        c = str(col).strip().lower()
        if "customer id"     in c: customer_id_col   = col
        elif "customer name" in c: customer_name_col = col
        elif "amount" in c or "revenue" in c: revenue_col = col
        elif "article" in c or "traffic" in c: traffic_col = col
        elif "start" in c: start_date_col = col
        elif "end"   in c: end_date_col   = col

    return customer_id_col, customer_name_col, revenue_col, traffic_col, start_date_col, end_date_col


def parse_master_from_daily_format(file):
    """
    Parse a daily-format CSV as master data.
    Groups by Customer ID + year + month, summing Revenue and Traffic.
    Returns a DataFrame in the same pivot format as parse_master_excel().
    """
    try:
        df = pd.read_csv(file)
    except Exception:
        return pd.DataFrame()

    cid_col, name_col, rev_col, trf_col, start_col, _ = detect_daily_columns(df)

    if not all([cid_col, rev_col, trf_col, start_col]):
        return pd.DataFrame()

    df[start_col] = parse_dates_robust(df[start_col])
    df["_year"]   = df[start_col].dt.year
    df["_month"]  = df[start_col].dt.month

    # Clean customer ID
    df["_clean_id"] = df[cid_col].astype(str).str.replace(".0", "", regex=False).str.strip()

    grouped = (
        df.groupby(["_clean_id", "_year", "_month"])
        .agg(Revenue=(rev_col, "sum"), Traffic=(trf_col, "sum"))
        .reset_index()
    )

    # Build customer name map
    name_map = {}
    if name_col:
        name_map = (
            df.groupby("_clean_id")[name_col]
            .first()
            .to_dict()
        )

    master = {}
    for _, row in grouped.iterrows():
        cid  = row["_clean_id"]
        key  = f"{int(row['_year'])}-{int(row['_month']):02d}"
        if cid not in master:
            master[cid] = {
                "CUSTOMER ID":   cid,
                "CUSTOMER NAME": str(name_map.get(cid, "")),
            }
        master[cid][f"{key} REVENUE"] = row["Revenue"]
        master[cid][f"{key} TRAFFIC"] = row["Traffic"]

    return pd.DataFrame(list(master.values())) if master else pd.DataFrame()


def compute_average_based_result(customer_id, customer_name, current_revenue,
                                  current_traffic, hist_row, historical_df,
                                  uploaded_days, sd_percent):
    """Compute expected values from average of ALL available FY months."""
    revenue_cols = [c for c in historical_df.columns if "REVENUE" in str(c).upper()]
    traffic_cols = [c for c in historical_df.columns if "TRAFFIC" in str(c).upper()]

    revenue_values, traffic_values = [], []

    for col in revenue_cols:
        val = pd.to_numeric(hist_row[col], errors='coerce')
        if pd.notna(val) and val > 0:
            revenue_values.append(val)

    for col in traffic_cols:
        val = pd.to_numeric(hist_row[col], errors='coerce')
        if pd.notna(val) and val > 0:
            traffic_values.append(val)

    if not revenue_values and not traffic_values:
        return None

    avg_monthly_revenue = np.mean(revenue_values) if revenue_values else 0
    avg_monthly_traffic = np.mean(traffic_values) if traffic_values else 0

    expected_revenue = (avg_monthly_revenue / 30.44) * uploaded_days
    expected_traffic = (avg_monthly_traffic / 30.44) * uploaded_days

    revenue_var = (
        ((current_revenue - expected_revenue) / expected_revenue) * 100
        if expected_revenue > 0 else np.nan
    )
    traffic_var = (
        ((current_traffic - expected_traffic) / expected_traffic) * 100
        if expected_traffic > 0 else np.nan
    )

    return {
        "Customer ID":           customer_id,
        "Customer Name":         customer_name,
        "Actual Revenue":        round(current_revenue),
        "Expected Revenue":      round(expected_revenue),
        "Revenue Variance %":    round(revenue_var, 2) if not pd.isna(revenue_var) else np.nan,
        "Revenue Status":        classify(revenue_var, sd_percent),
        "Actual Traffic":        round(current_traffic),
        "Expected Traffic":      round(expected_traffic),
        "Traffic Variance %":    round(traffic_var, 2) if not pd.isna(traffic_var) else np.nan,
        "Traffic Status":        classify(traffic_var, sd_percent),
        "Avg Revenue Used":      round(avg_monthly_revenue),
        "Avg Traffic Used":      round(avg_monthly_traffic),
        "Months Averaged (Rev)": len(revenue_values),
        "Months Averaged (Trf)": len(traffic_values),
    }


def write_grouped_sheet(writer, df, sheet_name, workbook, formats, status_col="Revenue Status"):
    """
    Write df to an Excel sheet grouped by status with a bold header row per group.
    """
    status_order  = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]
    header_format = workbook.add_format({
        "bold": True, "font_size": 12,
        "bg_color": "#2f3343", "font_color": "#FFFFFF",
        "border": 1,
    })
    col_header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#D9D9D9", "border": 1
    })
    plain_fmt = workbook.add_format({"border": 1})

    ws = writer.book.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    cols = list(df.columns)
    col_width = 22
    current_row = 0

    for status in status_order:
        grp = df[df[status_col] == status]
        if grp.empty:
            continue

        ws.merge_range(current_row, 0, current_row, len(cols) - 1,
                       f"{status}  ({len(grp)})", header_format)
        current_row += 1

        for ci, col in enumerate(cols):
            ws.write(current_row, ci, col, col_header_fmt)
            ws.set_column(ci, ci, col_width)
        current_row += 1

        rev_ci = cols.index("Revenue Status") if "Revenue Status" in cols else None
        trf_ci = cols.index("Traffic Status") if "Traffic Status" in cols else None

        for _, data_row in grp.iterrows():
            for ci, col in enumerate(cols):
                val = data_row[col]
                if isinstance(val, float) and np.isnan(val):
                    val = ""
                cell_fmt = plain_fmt
                if ci == rev_ci or ci == trf_ci:
                    cell_fmt = formats.get(str(val), plain_fmt)
                ws.write(current_row, ci, val, cell_fmt)
            current_row += 1

        current_row += 1  # blank separator between groups


# ==========================
# LOAD LOGO
# ==========================
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None

# ==========================
# SESSION STATE
# ==========================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# ==========================
# LOGIN PAGE
# ==========================
if not st.session_state.authenticated:
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    h_l, h_c, h_r = st.columns([1, 3, 1])
    with h_c:
        c_l, c_r = st.columns([1, 4])
        with c_l:
            if logo:
                st.image(logo, width=100)
        with c_r:
            st.markdown("""
            <div style='padding-top:4px;'>
            <h1 style='font-size:26px;margin-bottom:2px;color:#2f3343;font-weight:700;line-height:1.2;'>
            Analytics (Business and Operations)
            </h1>
            <p style='font-size:15px;color:#555;margin-top:2px;margin-bottom:0;'>
            Headquarters Region - Telangana Postal Circle
            </p>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<hr style='margin:10px 0 8px 0;border-color:#ddd;'>", unsafe_allow_html=True)

    l, c, r = st.columns([2.2, 1.4, 2.2])
    with c:
        st.markdown("""
        <h2 style='text-align:center;font-size:32px;color:#2f3343;
                   margin-top:4px;margin-bottom:12px;'>Login</h2>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            st.markdown("<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Username</p>",
                        unsafe_allow_html=True)
            username = st.text_input("", placeholder="Enter Username",
                                     label_visibility="collapsed", key="usr")
            st.markdown("<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Password</p>",
                        unsafe_allow_html=True)
            password = st.text_input("", type="password", placeholder="Enter Password",
                                     label_visibility="collapsed", key="pwd")
            submitted = st.form_submit_button("Submit", use_container_width=True, type="primary")

        if submitted:
            if username == "admin" and password == "HQR@2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")

    st.stop()

# ==========================
# MAIN APP — POST LOGIN
# ==========================

h_l, h_c, h_r = st.columns([1, 8, 1])
with h_l:
    if logo:
        st.image(logo, width=90)
with h_c:
    st.markdown("""
    <h1 style='font-size:28px;margin-bottom:2px;color:#2f3343;font-weight:700;padding-top:4px;'>
    Bulk Customer Business Analytics
    </h1>
    <p style='font-size:15px;color:#555;margin-top:0;'>
    Headquarters Region - Telangana Postal Circle
    </p>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ---- Sidebar ----
st.sidebar.header("Upload Files")

daily_file  = st.sidebar.file_uploader("Upload Daily / Period File (CSV)", type=["csv"])
master_file = st.sidebar.file_uploader(
    "Upload Master Data File (optional — overrides default)",
    type=["xlsx", "xls", "csv"]
)

_master_candidates = _glob.glob("data/[Mm]aster.xlsx") + _glob.glob("data/[Mm]aster.xls")
DEFAULT_MASTER = _master_candidates[0] if _master_candidates else "data/master.xlsx"

if master_file:
    st.sidebar.success("✅ Using uploaded master data")
elif os.path.exists(DEFAULT_MASTER):
    st.sidebar.info("📂 Using default master data")
else:
    st.sidebar.warning("⚠️ No master data found. Please upload one.")

sd_percent = st.sidebar.slider("Deviation %", min_value=1, max_value=50, value=10)

show_mode = st.sidebar.radio(
    "Filter Records",
    [
        "All records (mark unmatched as 'No Historical Data')",
        "Only records present in Master Data"
    ],
    index=0,
    help=(
        "The daily file may cover a larger region. "
        "Select 'Only records present in Master Data' to limit output "
        "to customers your division has master data for."
    )
)

# ================================================================
# MAIN PROCESS
# ================================================================
if daily_file and (master_file or os.path.exists(DEFAULT_MASTER)):

    # ----------------------------------------------------------------
    # Step 1 — Load daily CSV
    # ----------------------------------------------------------------
    daily_df = pd.read_csv(daily_file)

    # ----------------------------------------------------------------
    # Step 2 — Detect columns FIRST (fixes the NameError crash)
    # ----------------------------------------------------------------
    (customer_id_col, customer_name_col, revenue_col,
     traffic_col, start_date_col, end_date_col) = detect_daily_columns(daily_df)

    missing = [n for n, v in [
        ("Customer ID",   customer_id_col),
        ("Customer Name", customer_name_col),
        ("Revenue",       revenue_col),
        ("Traffic",       traffic_col),
        ("Start Date",    start_date_col),
        ("End Date",      end_date_col),
    ] if v is None]
    if missing:
        st.error(f"Could not detect columns in daily file: {', '.join(missing)}")
        st.stop()

    # ----------------------------------------------------------------
    # Step 3 — Parse dates robustly AFTER columns are known
    # ----------------------------------------------------------------
    daily_df[start_date_col] = parse_dates_robust(daily_df[start_date_col])
    daily_df[end_date_col]   = parse_dates_robust(daily_df[end_date_col])

    # ----------------------------------------------------------------
    # Step 4 — Load master data
    #   Supports: .xlsx/.xls (standard master), daily-format .csv
    # ----------------------------------------------------------------
    with st.spinner("Loading master data..."):
        active_master = master_file if master_file else (
            DEFAULT_MASTER if os.path.exists(DEFAULT_MASTER) else None
        )
        if active_master is None:
            st.error("No master data available. Please upload a master file.")
            st.stop()

        src_name = active_master.name if hasattr(active_master, "name") else str(active_master)

        if src_name.lower().endswith(".csv"):
            # First try interpreting as a daily-format CSV master
            historical_df = parse_master_from_daily_format(active_master)
            if historical_df.empty:
                # Plain CSV fallback
                active_master.seek(0)
                historical_df = pd.read_csv(active_master)
                st.sidebar.info("ℹ️ Master CSV read as plain table.")
            else:
                st.sidebar.success("✅ Daily-format CSV converted to master data.")
        else:
            historical_df = parse_master_excel(active_master)

    if historical_df.empty:
        st.error("Could not read master data. Please check the file.")
        st.stop()

    # ----------------------------------------------------------------
    # Step 5 — Date maths
    # ----------------------------------------------------------------
    upload_start  = pd.to_datetime(daily_df[start_date_col].iloc[0], errors="coerce")
    upload_end    = pd.to_datetime(daily_df[end_date_col].iloc[0],   errors="coerce")
    uploaded_days = (upload_end - upload_start).days + 1
    previous_year = upload_start.year - 1
    upload_month  = upload_start.month
    days_in_month = calendar.monthrange(previous_year, upload_month)[1]

    st.markdown(f"""
    <div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
                border-radius:6px;margin-bottom:12px;font-size:15px;'>
    <b>Period:</b> {upload_start.strftime('%d %b %Y')} → {upload_end.strftime('%d %b %Y')}
    &nbsp;|&nbsp; <b>Days uploaded:</b> {uploaded_days}
    &nbsp;|&nbsp; <b>Comparing against:</b> {calendar.month_name[upload_month]} {previous_year}
    &nbsp;|&nbsp; <b>Days in that month:</b> {days_in_month}
    </div>
    """, unsafe_allow_html=True)

    # ----------------------------------------------------------------
    # Step 6 — Historical column detection
    # ----------------------------------------------------------------
    hist_customer_id_col = revenue_col_hist = traffic_col_hist = None

    for col in historical_df.columns:
        c = str(col).upper()
        if "CUSTOMER ID" in c:
            hist_customer_id_col = col
        if str(previous_year) in c and f"{upload_month:02d}" in c and "REVENUE" in c:
            revenue_col_hist = col
        if str(previous_year) in c and f"{upload_month:02d}" in c and "TRAFFIC" in c:
            traffic_col_hist = col

    if hist_customer_id_col is None:
        st.error("Could not find CUSTOMER ID column in master data.")
        st.stop()

    has_period_col = (revenue_col_hist is not None and traffic_col_hist is not None)

    historical_df["CLEAN_ID"] = (
        historical_df[hist_customer_id_col]
        .astype(str).str.replace(".0", "", regex=False).str.strip()
    )

    # ----------------------------------------------------------------
    # Step 7 — KPI Metrics
    # ----------------------------------------------------------------
    total_revenue   = pd.to_numeric(daily_df[revenue_col], errors="coerce").sum()
    total_traffic   = pd.to_numeric(daily_df[traffic_col], errors="coerce").sum()
    total_customers = daily_df[customer_id_col].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Revenue",   f"₹ {format_indian(total_revenue)}")
    c2.metric("Total Traffic",   format_indian(total_traffic))
    c3.metric("Total Customers", total_customers)

    st.markdown("<hr style='margin:8px 0;border-color:#eee;'>", unsafe_allow_html=True)

    use_average_history = st.checkbox(
        "Analyze 'No Historical Data' customers using average of available historical months"
    )

    # ================================================================
    # CUSTOMER ANALYTICS LOOP
    # ================================================================
    results          = []
    no_hist_raw      = []
    avg_hist_results = []

    def _no_hist_entry(cid, cname, rev, trf):
        return {
            "Customer ID":                cid,
            "Customer Name":              cname,
            "Actual Revenue":             round(rev) if not pd.isna(rev) else 0,
            "Actual Traffic":             round(trf) if not pd.isna(trf) else 0,
            "Historical Monthly Revenue": 0,
            "Revenue Status":             "No Historical Data",
            "Traffic Status":             "No Historical Data",
        }

    for _, row in daily_df.iterrows():

        customer_id     = str(row[customer_id_col]).replace(".0", "").strip()
        customer_name   = row[customer_name_col]
        current_revenue = pd.to_numeric(row[revenue_col], errors="coerce")
        current_traffic = pd.to_numeric(row[traffic_col], errors="coerce")

        historical_match = historical_df[historical_df["CLEAN_ID"] == customer_id]

        if historical_match.empty and show_mode == "Only records present in Master Data":
            continue

        # CASE 1: Not in master at all
        if historical_match.empty:
            no_hist_raw.append(_no_hist_entry(customer_id, customer_name,
                                               current_revenue, current_traffic))
            continue

        hist_row = historical_match.iloc[0]

        # CASE 2: Same-period column entirely missing in master
        if not has_period_col:
            avg_result = compute_average_based_result(
                customer_id, customer_name, current_revenue, current_traffic,
                hist_row, historical_df, uploaded_days, sd_percent
            )
            if avg_result:
                avg_hist_results.append(avg_result)
            else:
                no_hist_raw.append(_no_hist_entry(customer_id, customer_name,
                                                   current_revenue, current_traffic))
            continue

        # CASE 3: Customer in master, same-period column exists
        monthly_revenue = pd.to_numeric(hist_row[revenue_col_hist], errors="coerce")
        monthly_traffic = pd.to_numeric(hist_row[traffic_col_hist], errors="coerce")
        if pd.isna(monthly_revenue): monthly_revenue = 0
        if pd.isna(monthly_traffic): monthly_traffic = 0

        # CASE 3a: Same-period value is zero — try FY average
        if monthly_revenue == 0 and monthly_traffic == 0:
            avg_result = compute_average_based_result(
                customer_id, customer_name, current_revenue, current_traffic,
                hist_row, historical_df, uploaded_days, sd_percent
            )
            if avg_result:
                avg_hist_results.append(avg_result)
            else:
                no_hist_raw.append(_no_hist_entry(customer_id, customer_name,
                                                   current_revenue, current_traffic))
            continue

        # CASE 3b: Normal path — use same-period historical values
        expected_revenue = (monthly_revenue / days_in_month) * uploaded_days
        expected_traffic = (monthly_traffic / days_in_month) * uploaded_days

        revenue_var = (
            ((current_revenue - expected_revenue) / expected_revenue) * 100
            if expected_revenue > 0 else np.nan
        )
        traffic_var = (
            ((current_traffic - expected_traffic) / expected_traffic) * 100
            if expected_traffic > 0 else np.nan
        )

        results.append({
            "Customer ID":                  customer_id,
            "Customer Name":                customer_name,
            "Actual Revenue":               round(current_revenue),
            "Historical Monthly Revenue":   round(monthly_revenue),
            "Historical Avg Revenue":       round(expected_revenue),
            "Revenue Variance %":           round(revenue_var) if not pd.isna(revenue_var) else "",
            "Revenue Status":               classify(revenue_var, sd_percent),
            "Actual Traffic":               round(current_traffic),
            "Historical Monthly Traffic":   round(monthly_traffic),
            "Historical Avg Traffic":       round(expected_traffic),
            "Traffic Variance %":           round(traffic_var) if not pd.isna(traffic_var) else "",
            "Traffic Status":               classify(traffic_var, sd_percent),
        })

    result_df      = pd.DataFrame(results)
    no_hist_df     = pd.DataFrame(no_hist_raw)
    avg_history_df = pd.DataFrame(avg_hist_results)

    count_truly_no_data   = len(no_hist_raw)
    count_avg_processable = len(avg_hist_results)
    count_total_no_hist   = count_truly_no_data + count_avg_processable

    if result_df.empty and no_hist_df.empty and avg_history_df.empty:
        st.warning("No records to display.")
        st.stop()

    # Clean numeric columns
    for col in ["Actual Revenue", "Historical Monthly Revenue", "Historical Avg Revenue",
                "Actual Traffic", "Historical Monthly Traffic", "Historical Avg Traffic"]:
        if col in result_df.columns:
            result_df[col] = (
                pd.to_numeric(result_df[col], errors="coerce").fillna(0).round(0).astype(int)
            )

    # ================================================================
    # SECTION 1 — Customer Analytics (same-period comparison)
    # ================================================================
    st.subheader("Customer Analytics")

    status_order = ["Excellent", "Normal", "Warning", "Critical"]

    for status in status_order:
        if not result_df.empty:
            group_df = result_df[result_df["Revenue Status"] == status]
            if not group_df.empty:
                st.markdown(f"### {status} ({len(group_df)})")
                styled_df = group_df.style.map(color_status, subset=["Revenue Status", "Traffic Status"])
                st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # No Historical Data display logic
    if use_average_history:
        nh_display_df    = no_hist_df
        nh_display_count = count_truly_no_data
    else:
        if not avg_history_df.empty:
            avg_as_no_hist = avg_history_df[
                ["Customer ID", "Customer Name", "Actual Revenue", "Actual Traffic"]
            ].copy()
            avg_as_no_hist["Historical Monthly Revenue"] = 0
            avg_as_no_hist["Revenue Status"] = "No Historical Data"
            avg_as_no_hist["Traffic Status"]  = "No Historical Data"
            nh_display_df = pd.concat([no_hist_df, avg_as_no_hist], ignore_index=True)
        else:
            nh_display_df = no_hist_df
        nh_display_count = count_total_no_hist

    if nh_display_count > 0:
        st.markdown(f"### No Historical Data ({nh_display_count})")
        if not nh_display_df.empty:
            st.dataframe(nh_display_df, use_container_width=True, hide_index=True)

    # ================================================================
    # SECTION 2 — Average-Based Analysis (only when checkbox ticked)
    # ================================================================
    if use_average_history:

        st.markdown("---")
        st.subheader("Average-Based Analysis for No Historical Data Customers")

        st.info(
            f"**Total No Historical Data entries: {count_total_no_hist}**  \n"
            f"Out of these, **{count_avg_processable}** could be processed using "
            f"available FY month averages.  \n"
            f"**{count_truly_no_data}** entries have no historical data in master "
            f"at all and remain unclassified."
        )

        if not avg_history_df.empty:
            display_cols = [
                "Customer ID", "Customer Name",
                "Actual Revenue", "Expected Revenue", "Revenue Variance %", "Revenue Status",
                "Actual Traffic", "Expected Traffic", "Traffic Variance %", "Traffic Status",
            ]
            display_cols = [c for c in display_cols if c in avg_history_df.columns]

            st.markdown("---")
            st.subheader("Average Historical Performance Analysis")

            for status in status_order + ["No Historical Data"]:
                grp = avg_history_df[avg_history_df["Revenue Status"] == status]
                if not grp.empty:
                    st.markdown(f"### {status} ({len(grp)})")
                    styled = grp[display_cols].style.map(
                        color_status, subset=["Revenue Status", "Traffic Status"]
                    )
                    st.dataframe(styled, use_container_width=True, hide_index=True)

    # ================================================================
    # Excel Download
    # ================================================================
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        status_formats = {
            "Excellent":          workbook.add_format({"bg_color": "#90EE90", "border": 1}),
            "Normal":             workbook.add_format({"bg_color": "#FFFACD", "border": 1}),
            "Warning":            workbook.add_format({"bg_color": "#FFD580", "border": 1}),
            "Critical":           workbook.add_format({"bg_color": "#FF7F7F", "border": 1}),
            "No Historical Data": workbook.add_format({"bg_color": "#D3D3D3", "border": 1}),
        }

        sheet1_df = (
            pd.concat([result_df, nh_display_df], ignore_index=True)
            if not nh_display_df.empty else result_df.copy()
        )

        if not sheet1_df.empty:
            write_grouped_sheet(
                writer, sheet1_df,
                sheet_name="Customer Analytics",
                workbook=workbook,
                formats=status_formats,
            )

        if use_average_history and not avg_history_df.empty:
            avg_display_cols = [
                "Customer ID", "Customer Name",
                "Actual Revenue", "Expected Revenue", "Revenue Variance %", "Revenue Status",
                "Actual Traffic", "Expected Traffic", "Traffic Variance %", "Traffic Status",
                "Avg Revenue Used", "Avg Traffic Used",
                "Months Averaged (Rev)", "Months Averaged (Trf)",
            ]
            avg_display_cols = [c for c in avg_display_cols if c in avg_history_df.columns]
            write_grouped_sheet(
                writer, avg_history_df[avg_display_cols],
                sheet_name="Avg-Based Analysis",
                workbook=workbook,
                formats=status_formats,
            )

    st.download_button(
        "⬇ Download Excel Report",
        output.getvalue(),
        file_name="analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif daily_file and not master_file and not os.path.exists(DEFAULT_MASTER):
    st.info("Please also upload the Master Data file in the sidebar.")
elif (master_file or os.path.exists(DEFAULT_MASTER)) and not daily_file:
    st.info("Please also upload the Daily / Period CSV file in the sidebar.")
else:
    st.info("Please upload both files in the sidebar to begin.")
