import streamlit as st
import pandas as pd
import numpy as np
import calendar
import os
import io
import streamlit as st
from PIL import Image

def check_password():

    if st.session_state.get("authenticated"):
        return True

    # ---------- PAGE STYLE ----------
    st.markdown("""
    <style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }

    div[data-testid="stTextInput"]{
        margin-bottom: -10px;
    }

    .login-box {
        background-color: white;
        padding: 30px;
        border-radius: 15px;
        box-shadow: 0px 2px 10px rgba(0,0,0,0.08);
    }
    </style>
    """, unsafe_allow_html=True)

    # ---------- HEADER ----------
    col1, col2, col3 = st.columns([1,2,1])

    with col2:

        st.image(
            "assets/logo.png",
            width=120
        )

        st.markdown("""
        <h1 style="
        text-align:center;
        margin-top:-10px;
        margin-bottom:0px;
        font-size:42px;
        font-weight:700;">
        📊 Bulk Customer Business Analytics
        </h1>
        """, unsafe_allow_html=True)

        st.markdown("""
        <h4 style="
        text-align:center;
        color:#444;
        margin-top:0px;
        margin-bottom:25px;">
        Headquarter Region - Telangana Postal Circle
        </h4>
        """, unsafe_allow_html=True)

    # ---------- LOGIN CENTER ----------
    left, center, right = st.columns([1.5,1.2,1.5])

    with center:

        st.markdown(
            "<h2 style='text-align:center;'>Login</h2>",
            unsafe_allow_html=True
        )

        username = st.text_input(
            "Username",
            placeholder="Enter Username"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter Password"
        )

        st.markdown("<br>", unsafe_allow_html=True)

        login = st.button(
            "Submit",
            use_container_width=True
        )

        if login:

            if (
                username == "admin"
                and password == "HQR@2026"
            ):
                st.session_state.authenticated = True
                st.rerun()

            else:
                st.error(
                    "Invalid Username or Password"
                )

    st.stop()


check_password()
# ---------- PAGE CONFIG ----------
st.set_page_config(
    page_title="Bulk Customer Analytics",
    layout="wide"
)

# ---------- LOAD LOGO ----------
logo = Image.open("assets/logo.png")


# ---------- LOGIN FUNCTION ----------
def check_password():

    # already logged in
    if st.session_state.get("authenticated"):
        return True

    st.markdown("<br>", unsafe_allow_html=True)

    # Logo center
    col1, col2, col3 = st.columns([2, 2, 2])

    with col2:
        st.image(logo, width=180)

    # Title
    st.markdown(
        """
        <h1 style='text-align:center; margin-bottom:0px;'>
        Bulk Customer Business Analytics
        </h1>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <h3 style='text-align:center;
                   color:#444;
                   margin-top:0px;'>
        Headquarter Region - Telangana Postal Circle
        </h3>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # Login Box Center
    left, center, right = st.columns([1, 2, 1])

    with center:

        st.markdown("### Login")

        username = st.text_input("Username")
        password = st.text_input(
            "Password",
            type="password"
        )

        login = st.button(
            "Submit",
            use_container_width=True
        )

        if login:

            if (
                username == "admin"
                and password == "HQR@2026"
            ):
                st.session_state.authenticated = True
                st.rerun()

            else:
                st.error(
                    "Invalid Username or Password"
                )

    st.stop()


check_password()
# ==================================
# PAGE CONFIG
# ==================================
st.set_page_config(
    page_title="Bulk Customer Analytics",
    page_icon="📊",
    layout="wide"
)

logo = Image.open("assets/logo.png")

# ---------- HEADER ----------

left, center, right = st.columns([1,2,1])

with center:

    st.image(
        logo,
        width=180
    )

    st.markdown(
        """
        <h1 style='
        text-align:center;
        margin-top:-15px;
        margin-bottom:0px;
        font-size:48px;
        font-weight:700;'>
        📊 Bulk Customer Business Analytics
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <h3 style='
        text-align:center;
        color:#444;
        margin-top:0px;'>
        Headquarter Region - Telangana Postal Circle
        </h3>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")
# ==================================
# SAVE PATH
# ==================================
SAVE_PATH = (
    r"D:\BulkCustomerAnalytics"
    r"\Saved_Data"
    r"\historical_master.pkl"
)

# ==================================
# STATUS LOGIC
# ==================================
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


# ==================================
# COLOR LOGIC
# ==================================
def color_status(value):

    colors = {
        "Excellent": "#90EE90",
        "Normal": "#FFFACD",
        "Warning": "#FFD580",
        "Critical": "#FF7F7F",
        "No Historical Data": "#D3D3D3"
    }

    return (
        f"background-color: "
        f"{colors.get(value, '')}"
    )


# ==================================
# HISTORICAL DATABASE
# ==================================
st.sidebar.header(
    "Historical Database"
)

historical_upload = (
    st.sidebar.file_uploader(
        "Upload / Replace Historical File",
        type=["xlsx"]
    )
)

# Save historical permanently
if historical_upload is not None:

    historical_df = pd.read_excel(
        historical_upload,
        header=[0, 1]
    )

    cols = []

    for c1, c2 in (
        historical_df.columns
    ):

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

# ==================================
# LOAD SAVED DATABASE
# ==================================
if os.path.exists(
    SAVE_PATH
):

    historical_df = (
        pd.read_pickle(
            SAVE_PATH
        )
    )

    st.sidebar.success(
        "Historical Database Loaded"
    )

else:

    st.warning(
        "Please upload "
        "historical database once."
    )

    st.stop()

# ==================================
# DAILY FILE
# ==================================
daily_file = (
    st.sidebar.file_uploader(
        "Upload Top Performer File",
        type=["csv"]
    )
)

sd_percent = (
    st.sidebar.slider(
        "Deviation %",
        min_value=1,
        max_value=50,
        value=10
    )
)

# ==================================
# MAIN PROCESS
# ==================================
if daily_file:

    daily_df = pd.read_csv(
        daily_file
    )

    # ------------------------------
    # COLUMN DETECTION
    # ------------------------------
    customer_id_col = None
    customer_name_col = None
    revenue_col = None
    traffic_col = None
    start_date_col = None
    end_date_col = None

    for col in (
        daily_df.columns
    ):

        c = str(
            col
        ).strip().lower()

        if (
            "customer id"
            in c
        ):
            customer_id_col = col

        elif (
            c ==
            "customer name"
        ):
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

    # ------------------------------
    # DATE DETECTION
    # ------------------------------
    upload_start = (
        pd.to_datetime(
            daily_df[
                start_date_col
            ].iloc[0]
        )
    )

    upload_end = (
        pd.to_datetime(
            daily_df[
                end_date_col
            ].iloc[0]
        )
    )

    uploaded_days = (
        upload_end
        - upload_start
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

    # ------------------------------
    # HISTORICAL COLUMN DETECTION
    # ------------------------------
    hist_customer_id_col = None
    revenue_col_hist = None
    traffic_col_hist = None

    for col in (
        historical_df.columns
    ):

        c = str(col).upper()

        if (
            "CUSTOMER ID"
            in c
        ):
            hist_customer_id_col = col

        if (
            str(previous_year)
            in c
            and f"{upload_month:02d}"
            in c
            and "REVENUE" in c
        ):
            revenue_col_hist = col

        if (
            str(previous_year)
            in c
            and f"{upload_month:02d}"
            in c
            and "TRAFFIC" in c
        ):
            traffic_col_hist = col

    # ------------------------------
    # CLEAN CUSTOMER ID
    # ------------------------------
    historical_df[
        "CLEAN_ID"
    ] = (
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

    # ------------------------------
    # KPI
    # ------------------------------
    total_revenue = (
        pd.to_numeric(
            daily_df[
                revenue_col
            ],
            errors="coerce"
        ).sum()
    )

    total_traffic = (
        pd.to_numeric(
            daily_df[
                traffic_col
            ],
            errors="coerce"
        ).sum()
    )

    total_customers = (
        daily_df[
            customer_id_col
        ].nunique()
    )

    c1, c2, c3 = (
        st.columns(3)
    )

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

    # ==================================
    # ANALYTICS LIST
    # ==================================
    results = []
        # ==================================
    # CUSTOMER ANALYTICS
    # ==================================
    for _, row in (
        daily_df.iterrows()
    ):

        customer_id = (
            str(
                row[
                    customer_id_col
                ]
            )
            .replace(
                ".0",
                ""
            )
            .strip()
        )

        customer_name = row[
            customer_name_col
        ]

        current_revenue = (
            pd.to_numeric(
                row[
                    revenue_col
                ],
                errors="coerce"
            )
        )

        current_traffic = (
            pd.to_numeric(
                row[
                    traffic_col
                ],
                errors="coerce"
            )
        )

        historical_match = (
            historical_df[
                historical_df[
                    "CLEAN_ID"
                ]
                == customer_id
            ]
        )

        # ------------------------------
        # NO HISTORICAL DATA
        # ------------------------------
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

                "Revenue Variance %":
                    "",

                "Revenue Status":
                    "No Historical Data",

                "Traffic Variance %":
                    "",

                "Traffic Status":
                    "No Historical Data"
            })

            continue

        # ------------------------------
        # HISTORICAL VALUES
        # ------------------------------
        monthly_revenue = (
            pd.to_numeric(
                historical_match[
                    revenue_col_hist
                ].iloc[0],
                errors="coerce"
            )
        )

        monthly_traffic = (
            pd.to_numeric(
                historical_match[
                    traffic_col_hist
                ].iloc[0],
                errors="coerce"
            )
        )

        if pd.isna(
            monthly_revenue
        ):
            monthly_revenue = 0

        if pd.isna(
            monthly_traffic
        ):
            monthly_traffic = 0

        expected_revenue = (
            monthly_revenue
            / days_in_month
        ) * uploaded_days

        expected_traffic = (
            monthly_traffic
            / days_in_month
        ) * uploaded_days

        # ------------------------------
        # VARIANCE
        # ------------------------------
        if (
            expected_revenue > 0
        ):

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

        if (
            expected_traffic > 0
        ):

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

        revenue_status = (
            classify(
                revenue_var,
                sd_percent
            )
        )

        traffic_status = (
            classify(
                traffic_var,
                sd_percent
            )
        )

        results.append({

            "Customer ID":
                customer_id,

            "Customer Name":
                customer_name,

            "Actual Revenue":
                round(
                    current_revenue
                ),

            "Historical Monthly Revenue":
                round(
                    monthly_revenue
                ),

            "Historical Avg Revenue":
                round(
                    expected_revenue
                ),

            "Revenue Variance %":
                round(
                    revenue_var
                )
                if not pd.isna(
                    revenue_var
                )
                else "",

            "Revenue Status":
                revenue_status,

            "Actual Traffic":
                round(
                    current_traffic
                ),

            "Historical Monthly Traffic":
                round(
                    monthly_traffic
                ),

            "Historical Avg Traffic":
                round(
                    expected_traffic
                ),

            "Traffic Variance %":
                round(
                    traffic_var
                )
                if not pd.isna(
                    traffic_var
                )
                else "",

            "Traffic Status":
                traffic_status
        })

    # ==================================
    # DATAFRAME
    # ==================================
    result_df = pd.DataFrame(
        results
    )

    # ==================================
    # REMOVE DECIMALS
    # ==================================
    number_cols = [

        "Actual Revenue",
        "Historical Monthly Revenue",
        "Historical Avg Revenue",

        "Actual Traffic",
        "Historical Monthly Traffic",
        "Historical Avg Traffic",

        "Revenue Variance %",
        "Traffic Variance %"
    ]

    for col in number_cols:

        if col in result_df.columns:

            result_df[col] = (
                pd.to_numeric(
                    result_df[col],
                    errors="coerce"
                )
                .fillna(0)
                .round(0)
                .astype(int)
            )
    st.subheader(
        "Customer Analytics"
    )

    # Group by Revenue Status
    status_order = [
        "Excellent",
        "Normal",
        "Warning",
        "Critical",
        "No Historical Data"
    ]

    for status in (
        status_order
    ):

        group_df = (
            result_df[
                result_df[
                    "Revenue Status"
                ]
                == status
            ]
        )

        if (
            not group_df.empty
        ):

            st.markdown(
                f"### "
                f"{status} "
                f"({len(group_df)})"
            )

            styled_df = (
                group_df.style.map(
                    color_status,
                    subset=[
                        "Revenue Status",
                        "Traffic Status"
                    ]
                )
            )

            st.dataframe(
                styled_df,
                use_container_width=True,
                hide_index=True
            )

    # ==================================
    # EXCEL DOWNLOAD
    # ==================================
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        result_df.to_excel(
            writer,
            index=False,
            sheet_name="Analytics"
        )

        workbook = writer.book
        worksheet = (
            writer.sheets[
                "Analytics"
            ]
        )

        # Status colors
        formats = {
            "Excellent":
                workbook.add_format({
                    "bg_color":
                    "#90EE90"
                }),

            "Normal":
                workbook.add_format({
                    "bg_color":
                    "#FFFACD"
                }),

            "Warning":
                workbook.add_format({
                    "bg_color":
                    "#FFD580"
                }),

            "Critical":
                workbook.add_format({
                    "bg_color":
                    "#FF7F7F"
                }),

            "No Historical Data":
                workbook.add_format({
                    "bg_color":
                    "#D3D3D3"
                })
        }

        revenue_status_col = (
            result_df.columns
            .get_loc(
                "Revenue Status"
            )
        )

        traffic_status_col = (
            result_df.columns
            .get_loc(
                "Traffic Status"
            )
        )

        for row_num in range(
            len(result_df)
        ):

            rev_status = (
                result_df.iloc[
                    row_num
                ][
                    "Revenue Status"
                ]
            )

            trf_status = (
                result_df.iloc[
                    row_num
                ][
                    "Traffic Status"
                ]
            )

            worksheet.write(
                row_num + 1,
                revenue_status_col,
                rev_status,
                formats[
                    rev_status
                ]
            )

            worksheet.write(
                row_num + 1,
                traffic_status_col,
                trf_status,
                formats[
                    trf_status
                ]
            )

        # Auto width
        for i, col in enumerate(
            result_df.columns
        ):

            worksheet.set_column(
                i,
                i,
                22
            )

    excel_data = (
        output.getvalue()
    )

    st.download_button(
        "⬇ Download Excel Report",
        excel_data,
        file_name=
        "analytics_report.xlsx",
        mime=(
            "application/"
            "vnd.openxmlformats-"
            "officedocument."
            "spreadsheetml.sheet"
        )
    )
