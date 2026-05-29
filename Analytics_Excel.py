import streamlit as st
import pandas as pd
import numpy as np
import calendar
import io
from PIL import Image

# ==========================
# PAGE CONFIG
# ==========================
st.set_page_config(
    page_title="Bulk Customer Analytics",
    layout="wide"
)

# ==========================
# CUSTOM CSS
# ==========================
st.markdown("""
<style>

.block-container {
    padding-top: 3rem !important;
    padding-bottom: 0rem !important;
}

header {
    visibility: hidden;
    height: 0px !important;
}

/* Compact text inputs */
.stTextInput input {
    height: 44px;
    border-radius: 8px;
    font-size: 16px;
}

/* Submit button */
div.stButton > button {
    background-color: #ff4b4b;
    color: white;
    border: none;
    border-radius: 8px;
    width: 100%;
    height: 48px;
    font-size: 20px;
    font-weight: 600;
    margin-top: 8px;
}

div.stButton > button:hover {
    background-color: #e23d3d;
    color: white;
}

</style>
""", unsafe_allow_html=True)

# ==========================
# HELPER FUNCTIONS
# ==========================
def classify(variance, threshold):
    """Classify variance into status categories."""
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
    """Return background color style for a status value."""
    colors = {
        "Excellent":          "background-color: #90EE90",
        "Normal":             "background-color: #FFFACD",
        "Warning":            "background-color: #FFD580",
        "Critical":           "background-color: #FF7F7F",
        "No Historical Data": "background-color: #D3D3D3",
    }
    return colors.get(val, "")


# ==========================
# LOAD LOGO (optional)
# ==========================
import os
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None

# ==========================
# SESSION STATE INIT
# ==========================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# ==========================
# LOGIN SECTION
# ==========================
if not st.session_state.authenticated:

    # Header row — logo + title, centered
    h1, h2, h3 = st.columns([1, 3, 1])
    with h2:
        inner_l, inner_r = st.columns([1, 3])
        with inner_l:
            if logo:
                st.image(logo, width=110)
        with inner_r:
            st.markdown("""
            <div style='padding-top:6px;'>
            <h1 style='font-size:28px;margin-bottom:2px;color:#2f3343;font-weight:700;line-height:1.2;'>
            Bulk Customer Business Analytics
            </h1>
            <p style='font-size:16px;color:#555;margin-top:2px;margin-bottom:0;'>
            Headquarter Region - Telangana Postal Circle
            </p>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    left, center, right = st.columns([2.2, 1.4, 2.2])

    with center:

        st.markdown("""
        <h2 style='
            text-align:center;
            font-size:34px;
            color:#2f3343;
            margin-top:0px;
            margin-bottom:10px;
        '>
        Login
        </h2>
        """, unsafe_allow_html=True)

        st.markdown("<p style='font-size:17px; margin-bottom:2px; font-weight:600;'>Username</p>",
                    unsafe_allow_html=True)

        username = st.text_input(
            "",
            placeholder="Enter Username",
            label_visibility="collapsed",
            key="username_input"
        )

        st.markdown("<p style='font-size:17px; margin-bottom:2px; font-weight:600;'>Password</p>",
                    unsafe_allow_html=True)

        password = st.text_input(
            "",
            type="password",
            placeholder="Enter Password",
            label_visibility="collapsed",
            key="password_input"
        )

        submit = st.button("Submit", use_container_width=True, type="primary")

        if submit:
            if username == "admin" and password == "HQR@2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")

    st.stop()

# ==========================
# MAIN APP (post-login)
# ==========================

# ==========================
# SIDEBAR FILE UPLOADERS
# ==========================
daily_file = st.sidebar.file_uploader(
    "Upload Daily / Period File",
    type=["csv"]
)

historical_file = st.sidebar.file_uploader(
    "Upload Historical Data File",
    type=["csv"]
)

sd_percent = st.sidebar.slider(
    "Deviation %",
    min_value=1,
    max_value=50,
    value=10
)

# ==========================
# MAIN PROCESS
# ==========================
if daily_file and historical_file:

    daily_df      = pd.read_csv(daily_file)
    historical_df = pd.read_csv(historical_file)

    # ------------------------------
    # COLUMN DETECTION — DAILY FILE
    # ------------------------------
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
        elif c == "customer name":
            customer_name_col = col
        elif "amount" in c or "revenue" in c:
            revenue_col = col
        elif "article" in c or "traffic" in c:
            traffic_col = col
        elif "start" in c:
            start_date_col = col
        elif "end" in c:
            end_date_col = col

    # Validate required columns
    missing = [
        name for name, col in [
            ("Customer ID",   customer_id_col),
            ("Customer Name", customer_name_col),
            ("Revenue",       revenue_col),
            ("Traffic",       traffic_col),
            ("Start Date",    start_date_col),
            ("End Date",      end_date_col),
        ]
        if col is None
    ]

    if missing:
        st.error(f"Could not detect these columns in the daily file: {', '.join(missing)}")
        st.stop()

    # ------------------------------
    # DATE DETECTION
    # ------------------------------
    upload_start = pd.to_datetime(daily_df[start_date_col].iloc[0])
    upload_end   = pd.to_datetime(daily_df[end_date_col].iloc[0])
    uploaded_days = (upload_end - upload_start).days + 1

    previous_year = upload_start.year - 1
    upload_month  = upload_start.month
    days_in_month = calendar.monthrange(previous_year, upload_month)[1]

    st.success(
        f"Detected Period: {upload_start.date()} to {upload_end.date()} "
        f"({uploaded_days} days) | Comparing against: "
        f"{previous_year}-{upload_month:02d}"
    )

    # ------------------------------
    # COLUMN DETECTION — HISTORICAL
    # ------------------------------
    hist_customer_id_col = None
    revenue_col_hist     = None
    traffic_col_hist     = None

    for col in historical_df.columns:
        c = str(col).upper()

        if "CUSTOMER ID" in c:
            hist_customer_id_col = col

        if (
            str(previous_year) in c
            and f"{upload_month:02d}" in c
            and "REVENUE" in c
        ):
            revenue_col_hist = col

        if (
            str(previous_year) in c
            and f"{upload_month:02d}" in c
            and "TRAFFIC" in c
        ):
            traffic_col_hist = col

    if hist_customer_id_col is None:
        st.error("Could not detect 'Customer ID' column in historical file.")
        st.stop()

    # ------------------------------
    # CLEAN CUSTOMER ID
    # ------------------------------
    historical_df["CLEAN_ID"] = (
        historical_df[hist_customer_id_col]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # ------------------------------
    # KPI METRICS
    # ------------------------------
    total_revenue   = pd.to_numeric(daily_df[revenue_col],  errors="coerce").sum()
    total_traffic   = pd.to_numeric(daily_df[traffic_col],  errors="coerce").sum()
    total_customers = daily_df[customer_id_col].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Revenue",   f"₹ {round(total_revenue):,}")
    c2.metric("Traffic",   f"{round(total_traffic):,}")
    c3.metric("Customers", total_customers)

    st.markdown("---")

    # ==================================
    # CUSTOMER ANALYTICS LOOP
    # ==================================
    results = []

    for _, row in daily_df.iterrows():

        customer_id   = str(row[customer_id_col]).replace(".0", "").strip()
        customer_name = row[customer_name_col]

        current_revenue = pd.to_numeric(row[revenue_col], errors="coerce")
        current_traffic = pd.to_numeric(row[traffic_col], errors="coerce")

        historical_match = historical_df[historical_df["CLEAN_ID"] == customer_id]

        # No historical data
        if historical_match.empty or revenue_col_hist is None or traffic_col_hist is None:
            results.append({
                "Customer ID":       customer_id,
                "Customer Name":     customer_name,
                "Actual Revenue":    round(current_revenue) if not pd.isna(current_revenue) else 0,
                "Actual Traffic":    round(current_traffic) if not pd.isna(current_traffic) else 0,
                "Revenue Variance %": "",
                "Revenue Status":    "No Historical Data",
                "Traffic Variance %": "",
                "Traffic Status":    "No Historical Data",
            })
            continue

        # Historical values
        monthly_revenue = pd.to_numeric(
            historical_match[revenue_col_hist].iloc[0], errors="coerce"
        )
        monthly_traffic = pd.to_numeric(
            historical_match[traffic_col_hist].iloc[0], errors="coerce"
        )

        if pd.isna(monthly_revenue): monthly_revenue = 0
        if pd.isna(monthly_traffic): monthly_traffic = 0

        expected_revenue = (monthly_revenue / days_in_month) * uploaded_days
        expected_traffic = (monthly_traffic / days_in_month) * uploaded_days

        # Variance
        revenue_var = (
            ((current_revenue - expected_revenue) / expected_revenue) * 100
            if expected_revenue > 0 else np.nan
        )
        traffic_var = (
            ((current_traffic - expected_traffic) / expected_traffic) * 100
            if expected_traffic > 0 else np.nan
        )

        revenue_status = classify(revenue_var, sd_percent)
        traffic_status = classify(traffic_var, sd_percent)

        results.append({
            "Customer ID":                 customer_id,
            "Customer Name":               customer_name,
            "Actual Revenue":              round(current_revenue),
            "Historical Monthly Revenue":  round(monthly_revenue),
            "Historical Avg Revenue":      round(expected_revenue),
            "Revenue Variance %":          round(revenue_var) if not pd.isna(revenue_var) else "",
            "Revenue Status":              revenue_status,
            "Actual Traffic":              round(current_traffic),
            "Historical Monthly Traffic":  round(monthly_traffic),
            "Historical Avg Traffic":      round(expected_traffic),
            "Traffic Variance %":          round(traffic_var) if not pd.isna(traffic_var) else "",
            "Traffic Status":              traffic_status,
        })

    # ==================================
    # BUILD DATAFRAME
    # ==================================
    result_df = pd.DataFrame(results)

    # Remove decimals from numeric columns
    number_cols = [
        "Actual Revenue", "Historical Monthly Revenue", "Historical Avg Revenue",
        "Actual Traffic", "Historical Monthly Traffic", "Historical Avg Traffic",
        "Revenue Variance %", "Traffic Variance %",
    ]

    for col in number_cols:
        if col in result_df.columns:
            result_df[col] = (
                pd.to_numeric(result_df[col], errors="coerce")
                .fillna(0)
                .round(0)
                .astype(int)
            )

    # ==================================
    # DISPLAY — GROUPED BY STATUS
    # ==================================
    st.subheader("Customer Analytics")

    status_order = ["Excellent", "Normal", "Warning", "Critical", "No Historical Data"]

    for status in status_order:
        group_df = result_df[result_df["Revenue Status"] == status]

        if not group_df.empty:
            st.markdown(f"### {status} ({len(group_df)})")

            styled_df = group_df.style.map(
                color_status,
                subset=["Revenue Status", "Traffic Status"]
            )

            st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # ==================================
    # EXCEL DOWNLOAD
    # ==================================
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

        revenue_status_col = result_df.columns.get_loc("Revenue Status")
        traffic_status_col = result_df.columns.get_loc("Traffic Status")

        for row_num in range(len(result_df)):
            rev_status = result_df.iloc[row_num]["Revenue Status"]
            trf_status = result_df.iloc[row_num]["Traffic Status"]

            worksheet.write(row_num + 1, revenue_status_col, rev_status, formats[rev_status])
            worksheet.write(row_num + 1, traffic_status_col, trf_status, formats[trf_status])

        for i, col in enumerate(result_df.columns):
            worksheet.set_column(i, i, 22)

    st.download_button(
        "⬇ Download Excel Report",
        output.getvalue(),
        file_name="analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif daily_file and not historical_file:
    st.info("Please also upload the Historical Data file in the sidebar.")

elif historical_file and not daily_file:
    st.info("Please also upload the Daily / Period file in the sidebar.")

else:
    st.info("Please upload both files in the sidebar to begin.")
