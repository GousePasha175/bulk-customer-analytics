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
st.set_page_config(page_title="Bulk Customer Analytics", layout="wide")

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

def load_master_excel(file):
    """
    Load all sheets from the master Excel, detect Customer ID, Traffic and Revenue
    columns per month, and return a unified DataFrame with columns:
    CLEAN_ID, CUSTOMER_NAME, YEAR, MONTH, TRAFFIC, REVENUE
    """
    xl = pd.ExcelFile(file)
    all_records = []

    for sheet in xl.sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet, header=None)

        # Find header row: row that contains 'Customer ID' or 'Cust ID' (case-insensitive)
        header_row = None
        id_col = None
        name_col = None

        for ridx in range(min(5, len(raw))):
            for cidx, val in enumerate(raw.iloc[ridx]):
                s = str(val).strip().lower()
                if "customer id" in s or "cust id" in s:
                    header_row = ridx
                    id_col = cidx
                    break
            if header_row is not None:
                break

        if header_row is None:
            continue  # skip sheet if no customer ID column found

        # Find customer name column in same header row
        for cidx, val in enumerate(raw.iloc[header_row]):
            s = str(val).strip().lower()
            if "customer name" in s or "cutomer name" in s or "name" in s:
                name_col = cidx
                break

        # Find month columns: rows 0..header_row contain datetime objects for months
        # and row header_row+1 (sub-header) has TRAFFIC/REVENUE labels
        # Build mapping: month -> (traffic_col_idx, revenue_col_idx)
        month_cols = {}  # key: (year, month) -> {'traffic': col, 'revenue': col}

        # Scan row 0 (or any row before header) for datetime objects
        for scan_row in range(header_row):
            for cidx, val in enumerate(raw.iloc[scan_row]):
                if isinstance(val, pd.Timestamp) or (
                    hasattr(val, 'year') and hasattr(val, 'month') and not isinstance(val, str)
                ):
                    try:
                        yr = val.year
                        mo = val.month
                        key = (yr, mo)
                        if key not in month_cols:
                            month_cols[key] = {'traffic': None, 'revenue': None}
                        # Now look in sub-header row (header_row or header_row-1 or header_row+1)
                        # for TRAFFIC / REVENUE at cidx and cidx+1
                        for sub_row in range(header_row + 1):
                            t = str(raw.iloc[sub_row].get(cidx, '')).strip().lower()
                            t1 = str(raw.iloc[sub_row].get(cidx + 1, '')).strip().lower()
                            if 'traffic' in t:
                                month_cols[key]['traffic'] = cidx
                            if 'revenue' in t or 'rev' in t:
                                month_cols[key]['revenue'] = cidx
                            if 'traffic' in t1:
                                month_cols[key]['traffic'] = cidx + 1
                            if 'revenue' in t1 or 'rev' in t1:
                                month_cols[key]['revenue'] = cidx + 1
                    except Exception:
                        pass

        # If no month_cols found via datetime, try sub-header approach for sheets
        # like BNPL Hyderabad where row 0 has datetime and row 1 has Traffic/Revenue
        if not month_cols:
            for scan_row in range(min(3, len(raw))):
                for cidx, val in enumerate(raw.iloc[scan_row]):
                    try:
                        if isinstance(val, pd.Timestamp):
                            yr = val.year
                            mo = val.month
                            key = (yr, mo)
                            month_cols[key] = {'traffic': cidx, 'revenue': cidx + 1}
                    except Exception:
                        pass

        if not month_cols:
            continue

        # Data starts from header_row + 2 (skip both header rows)
        data_start = header_row + 2

        for ridx in range(data_start, len(raw)):
            row = raw.iloc[ridx]

            # Skip totally empty rows or summary rows
            cid_val = row.iloc[id_col] if id_col < len(row) else None
            if pd.isna(cid_val) or str(cid_val).strip().lower() in ('', 'nan', 'total', 'grand total'):
                continue

            clean_id = str(cid_val).replace('.0', '').strip()

            cname = ''
            if name_col is not None and name_col < len(row):
                cname = str(row.iloc[name_col]).strip()

            for (yr, mo), cols in month_cols.items():
                t_col = cols.get('traffic')
                r_col = cols.get('revenue')

                traffic = 0
                revenue = 0

                if t_col is not None and t_col < len(row):
                    traffic = pd.to_numeric(row.iloc[t_col], errors='coerce')
                    if pd.isna(traffic):
                        traffic = 0

                if r_col is not None and r_col < len(row):
                    revenue = pd.to_numeric(row.iloc[r_col], errors='coerce')
                    if pd.isna(revenue):
                        revenue = 0

                all_records.append({
                    'CLEAN_ID':      clean_id,
                    'CUSTOMER_NAME': cname,
                    'YEAR':          yr,
                    'MONTH':         mo,
                    'TRAFFIC':       traffic,
                    'REVENUE':       revenue,
                })

    if not all_records:
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    # Deduplicate — same customer can appear in multiple sheets; sum them
    df = df.groupby(['CLEAN_ID', 'CUSTOMER_NAME', 'YEAR', 'MONTH'], as_index=False).sum()
    return df

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

    # Header — centered
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

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.markdown("<hr style='margin:0 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

    # Login form — centered
    l, c, r = st.columns([2.2, 1.4, 2.2])
    with c:
        st.markdown("""
        <h2 style='text-align:center;font-size:32px;color:#2f3343;margin-top:4px;margin-bottom:12px;'>
        Login
        </h2>
        """, unsafe_allow_html=True)

        st.markdown("<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Username</p>",
                    unsafe_allow_html=True)
        username = st.text_input("", placeholder="Enter Username",
                                 label_visibility="collapsed", key="usr")

        st.markdown("<p style='font-size:16px;margin-bottom:2px;font-weight:600;'>Password</p>",
                    unsafe_allow_html=True)
        password = st.text_input("", type="password", placeholder="Enter Password",
                                 label_visibility="collapsed", key="pwd")

        if st.button("Submit", use_container_width=True, type="primary"):
            if username == "admin" and password == "HQR@2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")

    st.stop()

# ==========================
# MAIN APP — POST LOGIN
# ==========================

# ---- Top Header ----
h_l, h_c, h_r = st.columns([1, 6, 1])
with h_l:
    if logo:
        st.image(logo, width=100)
with h_c:
    st.markdown("""
    <div style='padding-top:4px;'>
    <h1 style='font-size:30px;margin-bottom:2px;color:#2f3343;font-weight:700;'>
    Bulk Customer Business Analytics
    </h1>
    <p style='font-size:16px;color:#555;margin-top:2px;margin-bottom:0;'>
    Headquarter Region - Telangana Postal Circle
    </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin:6px 0 10px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# ---- Sidebar ----
st.sidebar.header("Upload Files")

daily_file = st.sidebar.file_uploader(
    "Upload Daily / Period File",
    type=["csv", "xlsx", "xls"]
)

master_file = st.sidebar.file_uploader(
    "Upload Master Data File (Historical)",
    type=["xlsx", "xls", "csv"]
)

sd_percent = st.sidebar.slider("Deviation %", min_value=1, max_value=50, value=10)

show_mode = st.sidebar.radio(
    "Display Records",
    ["All customers (include unmatched)", "Only customers in Master Data"],
    index=0
)

# ---- Main Process ----
if daily_file and master_file:

    # Read daily file
    if daily_file.name.endswith(".csv"):
        daily_df = pd.read_csv(daily_file)
    else:
        daily_df = pd.read_excel(daily_file)

    # Read master file
    with st.spinner("Loading master data..."):
        if master_file.name.endswith(".csv"):
            historical_df = pd.read_csv(master_file)
            # Convert to long format expected by load_master_excel
            # (assume CSV already has CLEAN_ID, YEAR, MONTH, TRAFFIC, REVENUE)
            historical_long = historical_df
        else:
            historical_long = load_master_excel(master_file)

    if historical_long.empty:
        st.error("Could not read master data. Please check the file format.")
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
        ("Customer ID", customer_id_col), ("Customer Name", customer_name_col),
        ("Revenue", revenue_col), ("Traffic", traffic_col),
        ("Start Date", start_date_col), ("End Date", end_date_col)
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
    <div style='background:#f0f7ff;border-left:4px solid #1a73e8;padding:10px 16px;
                border-radius:6px;margin-bottom:14px;'>
    <b>Period Detected:</b> {upload_start.strftime('%d %b %Y')} → {upload_end.strftime('%d %b %Y')}
    &nbsp;|&nbsp; <b>Days:</b> {uploaded_days}
    &nbsp;|&nbsp; <b>Comparing Against:</b> {calendar.month_name[upload_month]} {previous_year}
    &nbsp;|&nbsp; <b>Days in that month:</b> {days_in_month}
    </div>
    """, unsafe_allow_html=True)

    # ---- KPI Metrics ----
    total_revenue   = pd.to_numeric(daily_df[revenue_col],  errors="coerce").sum()
    total_traffic   = pd.to_numeric(daily_df[traffic_col],  errors="coerce").sum()
    total_customers = daily_df[customer_id_col].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Revenue",   f"₹ {round(total_revenue):,}")
    c2.metric("Total Traffic",   f"{round(total_traffic):,}")
    c3.metric("Total Customers", total_customers)

    st.markdown("<hr style='margin:10px 0;border-color:#eee;'>", unsafe_allow_html=True)

    # ---- Filter historical to the relevant month ----
    hist_month = historical_long[
        (historical_long['YEAR']  == previous_year) &
        (historical_long['MONTH'] == upload_month)
    ].copy()

    hist_month['CLEAN_ID'] = (
        hist_month['CLEAN_ID'].astype(str)
        .str.replace('.0', '', regex=False).str.strip()
    )

    # ---- Analytics Loop ----
    results = []

    for _, row in daily_df.iterrows():
        customer_id   = str(row[customer_id_col]).replace('.0', '').strip()
        customer_name = str(row[customer_name_col]).strip()
        current_rev   = pd.to_numeric(row[revenue_col], errors='coerce')
        current_trf   = pd.to_numeric(row[traffic_col], errors='coerce')

        match = hist_month[hist_month['CLEAN_ID'] == customer_id]

        if match.empty:
            if show_mode == "Only customers in Master Data":
                continue  # skip unmatched

            results.append({
                "Customer ID":        customer_id,
                "Customer Name":      customer_name,
                "Actual Revenue":     round(current_rev) if not pd.isna(current_rev) else 0,
                "Actual Traffic":     round(current_trf) if not pd.isna(current_trf) else 0,
                "Hist Monthly Rev":   "",
                "Hist Avg Rev":       "",
                "Revenue Variance %": "",
                "Revenue Status":     "No Historical Data",
                "Hist Monthly Trf":   "",
                "Hist Avg Trf":       "",
                "Traffic Variance %": "",
                "Traffic Status":     "No Historical Data",
            })
            continue

        monthly_rev = float(match['REVENUE'].iloc[0])
        monthly_trf = float(match['TRAFFIC'].iloc[0])

        expected_rev = (monthly_rev / days_in_month) * uploaded_days
        expected_trf = (monthly_trf / days_in_month) * uploaded_days

        rev_var = (
            ((current_rev - expected_rev) / expected_rev) * 100
            if expected_rev > 0 else np.nan
        )
        trf_var = (
            ((current_trf - expected_trf) / expected_trf) * 100
            if expected_trf > 0 else np.nan
        )

        results.append({
            "Customer ID":        customer_id,
            "Customer Name":      customer_name,
            "Actual Revenue":     round(current_rev),
            "Hist Monthly Rev":   round(monthly_rev),
            "Hist Avg Rev":       round(expected_rev),
            "Revenue Variance %": round(rev_var) if not pd.isna(rev_var) else "",
            "Revenue Status":     classify(rev_var, sd_percent),
            "Actual Traffic":     round(current_trf),
            "Hist Monthly Trf":   round(monthly_trf),
            "Hist Avg Trf":       round(expected_trf),
            "Traffic Variance %": round(trf_var) if not pd.isna(trf_var) else "",
            "Traffic Status":     classify(trf_var, sd_percent),
        })

    result_df = pd.DataFrame(results)

    if result_df.empty:
        st.warning("No matching records found.")
        st.stop()

    # Convert numeric cols cleanly
    num_cols = [
        "Actual Revenue", "Hist Monthly Rev", "Hist Avg Rev",
        "Actual Traffic", "Hist Monthly Trf", "Hist Avg Trf",
        "Revenue Variance %", "Traffic Variance %"
    ]
    for col in num_cols:
        if col in result_df.columns:
            result_df[col] = (
                pd.to_numeric(result_df[col], errors='coerce')
                .fillna(0).round(0).astype(int)
            )

    # ---- Display Grouped by Revenue Status ----
    st.subheader("Customer Analytics")

    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]

    for status in status_order:
        grp = result_df[result_df["Revenue Status"] == status]
        if not grp.empty:
            color_map = {
                "Excellent": "#e6f4ea", "Normal": "#fffde7",
                "Warning": "#fff3e0",  "Critical": "#fce8e6",
                "No Historical Data": "#f5f5f5"
            }
            bg = color_map.get(status, "#fff")
            st.markdown(
                f"<div style='background:{bg};padding:6px 12px;border-radius:6px;"
                f"margin-bottom:4px;font-weight:600;font-size:16px;'>"
                f"{status} — {len(grp)} customer(s)</div>",
                unsafe_allow_html=True
            )
            styled = grp.style.map(color_status, subset=["Revenue Status", "Traffic Status"])
            st.dataframe(styled, use_container_width=True, hide_index=True)

    # ---- Excel Download ----
    st.markdown("<br>", unsafe_allow_html=True)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Analytics")
        wb  = writer.book
        ws  = writer.sheets["Analytics"]

        fmt_map = {
            "Excellent":          wb.add_format({"bg_color": "#90EE90"}),
            "Normal":             wb.add_format({"bg_color": "#FFFACD"}),
            "Warning":            wb.add_format({"bg_color": "#FFD580"}),
            "Critical":           wb.add_format({"bg_color": "#FF7F7F"}),
            "No Historical Data": wb.add_format({"bg_color": "#D3D3D3"}),
        }

        rev_col_idx = result_df.columns.get_loc("Revenue Status")
        trf_col_idx = result_df.columns.get_loc("Traffic Status")

        for i in range(len(result_df)):
            rs = result_df.iloc[i]["Revenue Status"]
            ts = result_df.iloc[i]["Traffic Status"]
            ws.write(i + 1, rev_col_idx, rs, fmt_map.get(rs, wb.add_format()))
            ws.write(i + 1, trf_col_idx, ts, fmt_map.get(ts, wb.add_format()))

        for i, col in enumerate(result_df.columns):
            ws.set_column(i, i, 22)

    st.download_button(
        "⬇ Download Excel Report",
        output.getvalue(),
        file_name="analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif daily_file and not master_file:
    st.info("Please upload the Master Data (Historical) file in the sidebar.")
elif master_file and not daily_file:
    st.info("Please upload the Daily / Period file in the sidebar.")
else:
    st.info("Please upload both files in the sidebar to begin.")
