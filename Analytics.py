import streamlit as st
import pandas as pd
import calendar
import os
import numpy as np

# -----------------------------------
# PAGE CONFIG
# -----------------------------------
st.set_page_config(
    page_title="Bulk Customer Analytics",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Bulk Customer Business Analytics")
st.subheader(
    "Headquarter Region - Telangana Postal Circle"
)

st.markdown("---")

# -----------------------------------
# FILE SAVE PATH
# -----------------------------------
SAVE_PATH = (
    r"D:\BulkCustomerAnalytics"
    r"\Saved_Data"
    r"\historical_master.pkl"
)

# -----------------------------------
# STATUS CLASSIFICATION
# -----------------------------------
def classify(var, sd):

    if pd.isna(var):
        return "No Historical Data"

    if var > sd:
        return "Excellent"

    elif -sd <= var <= sd:
        return "Normal"

    elif -30 <= var < -sd:
        return "Warning"

    else:
        return "Critical"


# -----------------------------------
# COLOR CODING
# -----------------------------------
def color_status(val):

    colors = {
        "Excellent": "#90EE90",
        "Normal": "#FFFACD",
        "Warning": "#FFD580",
        "Critical": "#FF7F7F",
        "No Historical Data": "#D3D3D3"
    }

    return (
        f"background-color:"
        f"{colors.get(val, '')}"
    )


# -----------------------------------
# HISTORICAL DATABASE
# -----------------------------------
st.sidebar.header("Historical Database")

historical_upload = st.sidebar.file_uploader(
    "Upload / Replace Historical File",
    type=["xlsx"]
)

if historical_upload is not None:

    historical_df = pd.read_excel(
        historical_upload,
        header=[0, 1]
    )

    # Flatten merged columns
    cols = []

    for c1, c2 in historical_df.columns:

        c1 = str(c1).strip()
        c2 = str(c2).strip()

        cols.append(
            f"{c1}_{c2}"
        )

    historical_df.columns = cols

    historical_df.to_pickle(
        SAVE_PATH
    )

    st.sidebar.success(
        "Historical Database Saved"
    )

# -----------------------------------
# LOAD SAVED DATABASE
# -----------------------------------
if os.path.exists(SAVE_PATH):

    historical_df = pd.read_pickle(
        SAVE_PATH
    )

    st.sidebar.success(
        "Historical Database Loaded"
    )

else:

    st.warning(
        "Upload historical file once."
    )

    st.stop()

# -----------------------------------
# DAILY FILE
# -----------------------------------
daily_file = st.sidebar.file_uploader(
    "Upload Top Performer File",
    type=["csv"]
)

sd_percent = st.sidebar.slider(
    "Deviation %",
    1,
    50,
    10
)

if daily_file:

    daily_df = pd.read_csv(
        daily_file
    )

    # -----------------------------------
    # COLUMN DETECTION
    # -----------------------------------
    customer_id_col = None
    customer_name_col = None
    revenue_col = None
    traffic_col = None
    start_date_col = None
    end_date_col = None

    for col in daily_df.columns:

        c = str(col).strip().lower()

        if "customer id" in c:
            customer_id_col = col

        elif c == "customer name":
            customer_name_col = col

        elif (
            "amount" in c
            or "revenue" in c
        ):
            revenue_col = col

        elif (
            "article" in c
            or "traffic" in c
        ):
            traffic_col = col

        elif "start" in c:
            start_date_col = col

        elif "end" in c:
            end_date_col = col

    # -----------------------------------
    # DATE DETECTION
    # -----------------------------------
    upload_start = pd.to_datetime(
        daily_df[start_date_col].iloc[0]
    )

    upload_end = pd.to_datetime(
        daily_df[end_date_col].iloc[0]
    )

    uploaded_days = (
        upload_end - upload_start
    ).days + 1

    previous_year = (
        upload_start.year - 1
    )

    upload_month = (
        upload_start.month
    )

    target_month = (
        f"{previous_year}-"
        f"{upload_month:02d}"
    )

    days_in_month = (
        calendar.monthrange(
            previous_year,
            upload_month
        )[1]
    )

    st.success(
        f"Detected Period: "
        f"{upload_start.date()} "
        f"to "
        f"{upload_end.date()} "
        f"({uploaded_days} days)"
    )

    # -----------------------------------
    # HISTORICAL COLUMN DETECTION
    # -----------------------------------
    hist_customer_id_col = None
    revenue_col_hist = None
    traffic_col_hist = None

    for col in historical_df.columns:

        c = str(col).upper()

        if "CUSTOMER ID" in c:
            hist_customer_id_col = col

        if (
            target_month in c
            and "REVENUE" in c
        ):
            revenue_col_hist = col

        if (
            target_month in c
            and "TRAFFIC" in c
        ):
            traffic_col_hist = col

    # -----------------------------------
    # CLEAN IDS
    # -----------------------------------
    historical_df["CLEAN_ID"] = (
        historical_df[
            hist_customer_id_col
        ]
        .astype(str)
        .str.replace(
            ".0",
            "",
            regex=False
        )
        .str.strip()
    )

    # -----------------------------------
    # KPI SECTION
    # -----------------------------------
    total_revenue = pd.to_numeric(
        daily_df[revenue_col],
        errors="coerce"
    ).sum()

    total_traffic = pd.to_numeric(
        daily_df[traffic_col],
        errors="coerce"
    ).sum()

    total_customers = (
        daily_df[
            customer_id_col
        ].nunique()
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Revenue",
        f"₹ {round(total_revenue):,}"
    )

    c2.metric(
        "Traffic",
        f"{round(total_traffic):,}"
    )

    c3.metric(
        "Customers",
        total_customers
    )

    st.markdown("---")

    # -----------------------------------
    # ANALYTICS
    # -----------------------------------
    results = []

    for _, row in daily_df.iterrows():

        customer_id = (
            str(
                row[
                    customer_id_col
                ]
            )
            .replace(".0", "")
            .strip()
        )

        customer_name = row[
            customer_name_col
        ]

        current_revenue = pd.to_numeric(
            row[revenue_col],
            errors="coerce"
        )

        current_traffic = pd.to_numeric(
            row[traffic_col],
            errors="coerce"
        )

        historical_match = historical_df[
            historical_df[
                "CLEAN_ID"
            ] == customer_id
        ]

        # -------------------------------
        # NO HISTORICAL DATA
        # -------------------------------
        if (
            historical_match.empty
            or revenue_col_hist is None
            or traffic_col_hist is None
        ):

            results.append({

                "Customer ID":
                    customer_id,

                "Customer Name":
                    customer_name,

                "Actual Revenue":
                    round(
                        current_revenue
                    ),

                "Actual Traffic":
                    round(
                        current_traffic
                    ),

                "Status":
                    "No Historical Data"
            })

            continue

        # -------------------------------
        # HISTORICAL VALUES
        # -------------------------------
        monthly_revenue = pd.to_numeric(
            historical_match[
                revenue_col_hist
            ].iloc[0],
            errors="coerce"
        )

        monthly_traffic = pd.to_numeric(
            historical_match[
                traffic_col_hist
            ].iloc[0],
            errors="coerce"
        )

        # Handle NaN
        if pd.isna(monthly_revenue):
            monthly_revenue = 0

        if pd.isna(monthly_traffic):
            monthly_traffic = 0

        expected_revenue = (
            monthly_revenue
            / days_in_month
        ) * uploaded_days

        expected_traffic = (
            monthly_traffic
            / days_in_month
        ) * uploaded_days

        # Revenue variance
        if expected_revenue > 0:

            revenue_var = (
                (
                    current_revenue
                    - expected_revenue
                )
                /
                expected_revenue
            ) * 100

        else:
            revenue_var = np.nan

        # Traffic variance
        if expected_traffic > 0:

            traffic_var = (
                (
                    current_traffic
                    - expected_traffic
                )
                /
                expected_traffic
            ) * 100

        else:
            traffic_var = np.nan

        status = classify(
            revenue_var,
            sd_percent
        )

        results.append({

            "Customer ID":
                customer_id,

            "Customer Name":
                customer_name,

            "Actual Revenue":
                round(current_revenue),

            "Historical Monthly Revenue":
                round(monthly_revenue),

            "Historical Avg Revenue":
                round(expected_revenue),

            "Revenue Variance %":
                round(
                    revenue_var,
                    1
                ) if not pd.isna(
                    revenue_var
                ) else "",

            "Actual Traffic":
                round(current_traffic),

            "Historical Monthly Traffic":
                round(monthly_traffic),

            "Historical Avg Traffic":
                round(expected_traffic),

            "Traffic Variance %":
                round(
                    traffic_var,
                    1
                ) if not pd.isna(
                    traffic_var
                ) else "",

            "Status":
                status
        })

    result_df = pd.DataFrame(
        results
    )

    # -----------------------------------
    # DISPLAY
    # -----------------------------------
    st.subheader(
        "Customer Analytics"
    )

    styled_df = (
        result_df.style.map(
            color_status,
            subset=["Status"]
        )
    )

    st.dataframe(
        styled_df,
        use_container_width=True,
        hide_index=True
    )

    # -----------------------------------
    # DOWNLOAD
    # -----------------------------------
    csv = result_df.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(
        "⬇ Download Report",
        csv,
        "analytics_report.csv",
        "text/csv"
    )