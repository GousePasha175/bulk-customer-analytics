import streamlit as st
import pandas as pd
import numpy as np
import calendar
import io
import os
from PIL import Image

# ==========================
# PAGE CONFIG
# ==========================
# Sidebar Toggle
if "sidebar_state" not in st.session_state:
    st.session_state.sidebar_state = "expanded"

toggle = st.button("☰ Menu")

if toggle:
    st.session_state.sidebar_state = (
        "collapsed"
        if st.session_state.sidebar_state == "expanded"
        else "expanded"
    )

st.set_page_config(
    page_title="Bulk Customer Analytics",
    layout="wide",
    initial_sidebar_state=st.session_state.sidebar_state
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
</style>
""", unsafe_allow_html=True)

# ==========================
# HELPER FUNCTIONS
# ==========================
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
    """
    Read all sheets from the master Excel.
    Each sheet has 2 header rows:
      row 0 : month dates (datetime) in columns, paired as TRAFFIC | REVENUE
      row 1 : 'TRAFFIC' / 'REVENUE' sub-labels  (or single header row for some sheets)
    Data rows start after the last header row.

    Returns a flat DataFrame with one row per customer:
      CUSTOMER ID  |  CUSTOMER NAME  |  YYYY-MM TRAFFIC  |  YYYY-MM REVENUE  | ...
    """
    xl = pd.ExcelFile(file)
    # Collect per-customer monthly data across all sheets
    # key = clean customer id, value = dict of {col_name: value}
    master = {}

    for sheet in xl.sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet, header=None)

        # ---- locate customer-id column and data-start row ----
        id_col_idx   = None
        name_col_idx = None
        data_start   = None

        # PBC Autonagar has a single header row (row 0) with text month labels
        # All others have row 0 = dates, row 1 = TRAFFIC/REVENUE sub-labels
        # Detect by checking if row 1 contains 'TRAFFIC'
        has_sub_header = any(
            str(v).strip().upper() == 'TRAFFIC'
            for v in raw.iloc[1].tolist()
        )

        if has_sub_header:
            header_row  = 1   # sub-header row
            data_start  = 2
        else:
            header_row  = 0
            data_start  = 1

        # Find CUSTOMER ID and CUSTOMER NAME columns in any of the first 3 rows
        for scan in range(min(3, len(raw))):
            for cidx, val in enumerate(raw.iloc[scan]):
                s = str(val).strip().lower()
                if ('customer id' in s or 'cust id' in s) and id_col_idx is None:
                    id_col_idx = cidx
                if ('customer name' in s or 'cutomer name' in s) and name_col_idx is None:
                    name_col_idx = cidx

        if id_col_idx is None:
            continue  # can't use this sheet

        # ---- build month → (traffic_col, revenue_col) mapping ----
        # For sheets with datetime row 0 + sub-header row 1:
        #   row 0 has datetime at col c, next sub-header says TRAFFIC at c, REVENUE at c+1
        # For PBC Autonagar (single header row 0 with text like 'Apr-25 Traf'):
        #   parse directly from column name text

        month_map = {}  # (year, month) -> {'t': col_idx, 'r': col_idx}

        if has_sub_header:
            # row 0 has dates, row 1 has TRAFFIC/REVENUE
            date_row = raw.iloc[0]
            sub_row  = raw.iloc[1]
            current_date = None
            for cidx in range(len(date_row)):
                val = date_row.iloc[cidx]
                if hasattr(val, 'year'):          # it's a datetime
                    current_date = (val.year, val.month)
                    month_map[current_date] = {'t': None, 'r': None}
                if current_date is not None:
                    sub = str(sub_row.iloc[cidx]).strip().upper()
                    if sub == 'TRAFFIC':
                        month_map[current_date]['t'] = cidx
                    elif sub == 'REVENUE' or sub == 'REV':
                        month_map[current_date]['r'] = cidx
        else:
            # PBC Autonagar: single header row with text labels like 'Apr-25 Traf', 'Apr-25 Rev'
            month_abbr = {
                'apr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9,
                'oct': 10,'nov': 11,'dec': 12,'jan': 1, 'feb': 2, 'mar': 3
            }
            header_vals = raw.iloc[0].tolist()
            for cidx, val in enumerate(header_vals):
                s = str(val).strip().lower()
                # extract month abbreviation
                for abbr, mo in month_abbr.items():
                    if abbr in s:
                        # extract year: look for 2-digit year like '25' -> 2025
                        import re
                        yr_match = re.search(r'(\d{2})', s)
                        if yr_match:
                            yr_2 = int(yr_match.group(1))
                            yr = 2000 + yr_2
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

        # ---- read data rows ----
        for ridx in range(data_start, len(raw)):
            row = raw.iloc[ridx]
            cid_val = row.iloc[id_col_idx]

            # skip empty / summary rows
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

                # Sum across sheets if customer appears in multiple
                master[clean_id][t_key] = master[clean_id].get(t_key, 0) + t_val
                master[clean_id][r_key] = master[clean_id].get(r_key, 0) + r_val

    if not master:
        return pd.DataFrame()

    return pd.DataFrame(list(master.values()))

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
            Bulk Customer Business Analytics
            </h1>
            <p style='font-size:15px;color:#555;margin-top:2px;margin-bottom:0;'>
            Headquarter Region - Telangana Postal Circle
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
        st.markdown(
            "<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Username</p>",
            unsafe_allow_html=True
        )
    
        username = st.text_input(
            "",
            placeholder="Enter Username",
            label_visibility="collapsed",
            key="usr"
        )
    
        st.markdown(
            "<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Password</p>",
            unsafe_allow_html=True
        )
    
        password = st.text_input(
            "",
            type="password",
            placeholder="Enter Password",
            label_visibility="collapsed",
            key="pwd"
        )
    
        submit = st.form_submit_button(
            "Submit",
            use_container_width=True,
            type="primary"
        )
    
    if submit:
        if username == "admin" and password == "HQR@2026":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Invalid Username or Password")
    st.stop()

# ==========================
# MAIN APP — POST LOGIN
# ==========================

# ---- Header ----
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
    Headquarter Region - Telangana Postal Circle
    </p>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin:4px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ---- Sidebar ----
st.sidebar.header("Upload Files")

daily_file = st.sidebar.file_uploader(
    "Upload Daily / Period File (CSV)",
    type=["csv"]
)

master_file = st.sidebar.file_uploader(
    "Upload Master Data File",
    type=["xlsx", "xls", "csv"]
)

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

# ---- Main Process ----
if daily_file and master_file:

    # Read daily CSV
    daily_df = pd.read_csv(daily_file)

    # Read & flatten master Excel into the format the original logic expects:
    # one row per customer, columns: CUSTOMER ID | 2024-05 TRAFFIC | 2024-05 REVENUE ...
    with st.spinner("Reading master data..."):
        if master_file.name.endswith(".csv"):
            historical_df = pd.read_csv(master_file)
        else:
            historical_df = parse_master_excel(master_file)

    if historical_df.empty:
        st.error("Could not read master data. Please check the file.")
        st.stop()

    # ---- Column Detection — Daily File ----
    customer_id_col   = None
    customer_name_col = None
    revenue_col       = None
    traffic_col       = None
    start_date_col    = None
    end_date_col      = None

    for col in daily_df.columns:
        c = str(col).strip().lower()
        if "customer id" in c:
            customer_id_col = col
        elif "customer name" in c:
            customer_name_col = col
        elif "amount" in c or "revenue" in c:
            revenue_col = col
        elif "article" in c or "traffic" in c:
            traffic_col = col
        elif "start" in c:
            start_date_col = col
        elif "end" in c:
            end_date_col = col

    missing = [n for n, v in [
        ("Customer ID", customer_id_col),
        ("Customer Name", customer_name_col),
        ("Revenue", revenue_col),
        ("Traffic", traffic_col),
        ("Start Date", start_date_col),
        ("End Date", end_date_col),
    ] if v is None]

    if missing:
        st.error(f"Could not detect columns in daily file: {', '.join(missing)}")
        st.stop()

    # ---- Date Detection ----
    upload_start  = pd.to_datetime(daily_df[start_date_col].iloc[0])
    upload_end    = pd.to_datetime(daily_df[end_date_col].iloc[0])
    uploaded_days = (upload_end - upload_start).days + 1
    previous_year = upload_start.year - 1
    upload_month  = upload_start.month
    days_in_month = calendar.monthrange(previous_year, upload_month)[1]

    # ---- Date Validation Banner ----
    st.markdown(f"""
    <div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:8px 16px;
                border-radius:6px;margin-bottom:12px;font-size:15px;'>
    <b>Period:</b> {upload_start.strftime('%d %b %Y')} → {upload_end.strftime('%d %b %Y')}
    &nbsp;|&nbsp; <b>Days uploaded:</b> {uploaded_days}
    &nbsp;|&nbsp; <b>Comparing against:</b> {calendar.month_name[upload_month]} {previous_year}
    &nbsp;|&nbsp; <b>Days in that month:</b> {days_in_month}
    </div>
    """, unsafe_allow_html=True)

    # ---- Historical Column Detection (original logic, exact match) ----
    hist_customer_id_col = None
    revenue_col_hist     = None
    traffic_col_hist     = None

    for col in historical_df.columns:
        c = str(col).upper()
        if "CUSTOMER ID" in c:
            hist_customer_id_col = col
        if (str(previous_year) in c
                and f"{upload_month:02d}" in c
                and "REVENUE" in c):
            revenue_col_hist = col
        if (str(previous_year) in c
                and f"{upload_month:02d}" in c
                and "TRAFFIC" in c):
            traffic_col_hist = col

    if hist_customer_id_col is None:
        st.error("Could not find CUSTOMER ID column in master data.")
        st.stop()

    # ---- Clean Customer ID ----
    historical_df["CLEAN_ID"] = (
        historical_df[hist_customer_id_col]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # ---- KPI Metrics ----
    total_revenue   = pd.to_numeric(daily_df[revenue_col],  errors="coerce").sum()
    total_traffic   = pd.to_numeric(daily_df[traffic_col],  errors="coerce").sum()
    total_customers = daily_df[customer_id_col].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Revenue",   f"₹ {round(total_revenue):,}")
    c2.metric("Total Traffic",   f"{round(total_traffic):,}")
    c3.metric("Total Customers", total_customers)

    st.markdown("<hr style='margin:8px 0;border-color:#eee;'>", unsafe_allow_html=True)

    # ---- Customer Analytics Loop (original logic preserved) ----
    results = []

    for _, row in daily_df.iterrows():

        customer_id   = str(row[customer_id_col]).replace(".0", "").strip()
        customer_name = row[customer_name_col]
        current_revenue = pd.to_numeric(row[revenue_col], errors="coerce")
        current_traffic = pd.to_numeric(row[traffic_col], errors="coerce")

        historical_match = historical_df[historical_df["CLEAN_ID"] == customer_id]

        # Apply show_mode filter
        if historical_match.empty and show_mode == "Only records present in Master Data":
            continue

        # No historical data
        if (historical_match.empty
                or revenue_col_hist is None
                or traffic_col_hist is None):
            results.append({
                "Customer ID":               customer_id,
                "Customer Name":             customer_name,
                "Actual Revenue":            round(current_revenue) if not pd.isna(current_revenue) else 0,
                "Actual Traffic":            round(current_traffic) if not pd.isna(current_traffic) else 0,
                "Revenue Variance %":        "",
                "Revenue Status":            "No Historical Data",
                "Traffic Variance %":        "",
                "Traffic Status":            "No Historical Data",
            })
            continue

        # Historical values
        monthly_revenue = pd.to_numeric(
            historical_match[revenue_col_hist].iloc[0], errors="coerce")
        monthly_traffic = pd.to_numeric(
            historical_match[traffic_col_hist].iloc[0], errors="coerce")

        if pd.isna(monthly_revenue): monthly_revenue = 0
        if pd.isna(monthly_traffic): monthly_traffic = 0

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

    result_df = pd.DataFrame(results)

    if result_df.empty:
        st.warning("No records to display.")
        st.stop()

    # Remove decimals
    number_cols = [
        "Actual Revenue", "Historical Monthly Revenue", "Historical Avg Revenue",
        "Actual Traffic", "Historical Monthly Traffic", "Historical Avg Traffic",
        "Revenue Variance %", "Traffic Variance %",
    ]
    for col in number_cols:
        if col in result_df.columns:
            result_df[col] = (
                pd.to_numeric(result_df[col], errors="coerce")
                .fillna(0).round(0).astype(int)
            )

    # ---- Display Grouped by Revenue Status ----
    st.subheader("Customer Analytics")

    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]

    for status in status_order:
        group_df = result_df[result_df["Revenue Status"] == status]
        if not group_df.empty:
            st.markdown(f"### {status} ({len(group_df)})")
            styled_df = group_df.style.map(
                color_status, subset=["Revenue Status", "Traffic Status"])
            st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # ---- Excel Download ----
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Analytics")
        workbook  = writer.book
        worksheet = writer.sheets["Analytics"]

        formats = {
            "Excellent":          workbook.add_format({"bg_color": "#90EE90"}),
            "Normal":             workbook.add_format({"bg_color": "#FFFACD"}),
            "Warning":            workbook.add_format({"bg_color": "#FFD580"}),
            "Critical":           workbook.add_format({"bg_color": "#FF7F7F"}),
            "No Historical Data": workbook.add_format({"bg_color": "#D3D3D3"}),
        }

        rev_col_idx = result_df.columns.get_loc("Revenue Status")
        trf_col_idx = result_df.columns.get_loc("Traffic Status")

        for row_num in range(len(result_df)):
            rs = result_df.iloc[row_num]["Revenue Status"]
            ts = result_df.iloc[row_num]["Traffic Status"]
            worksheet.write(row_num + 1, rev_col_idx, rs, formats[rs])
            worksheet.write(row_num + 1, trf_col_idx, ts, formats[ts])

        for i, col in enumerate(result_df.columns):
            worksheet.set_column(i, i, 22)

    st.download_button(
        "⬇ Download Excel Report",
        output.getvalue(),
        file_name="analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif daily_file and not master_file:
    st.info("Please also upload the Master Data file in the sidebar.")
elif master_file and not daily_file:
    st.info("Please also upload the Daily / Period CSV file in the sidebar.")
else:
    st.info("Please upload both files in the sidebar to begin.")
