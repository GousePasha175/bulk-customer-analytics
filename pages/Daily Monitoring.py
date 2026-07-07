import streamlit as st
import pandas as pd
from PIL import Image
import os
import glob as _glob

# --------------------------------------------------------
# Navigation (Copied from your existing app)
# --------------------------------------------------------

def _render_nav():
    st.sidebar.markdown(
        """
        <div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;
        font-weight:700;
        color:#888;
        text-transform:uppercase;
        letter-spacing:1px;
        margin:0 0 4px 0;'>
        Pages
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.page_link("Analytics_Excel.py", label="🏠 Home")

    for pat, lbl in [
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py","🤚 AEBAS Monitoring"),
        ("pages/Bulk Analytics.py|pages/*[Bb]ulk*.py","📊 Bulk Customer Analytics"),
        ("pages/Delivery Productivity.py|pages/*[Dd]elivery*.py","📦 Delivery Productivity"),
        ("pages/Daily Monitoring.py","📈 Daily Monitoring"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py","📮 POSB Daily Report"),
        ("pages/Sorting_Assistance.py|pages/*[Ss]orting*.py","📮 Sorting Assistance"),
    ]:

        hits = []

        for p in pat.split("|"):
            hits += _glob.glob(p)

        if hits:
            st.sidebar.page_link(
                hits[0].replace("\\", "/"),
                label=lbl
            )

    st.sidebar.markdown(
        "<hr style='margin:8px 0 12px 0;'>",
        unsafe_allow_html=True
    )

# --------------------------------------------------------

st.set_page_config(
    page_title="Division-wise Daily Monitoring",
    page_icon="📈",
    layout="wide"
)
st.markdown("""
<style>

/* Hide Streamlit default navigation */
[data-testid="stSidebarNav"]{
    display:none!important;
}

/* Reduce top spacing */
.block-container{
    padding-top:1.2rem!important;
    padding-bottom:0!important;
}

/* Hide default header */
header{
    visibility:hidden;
    height:0!important;
}

/* Sidebar collapse button */
[data-testid="collapsedControl"]{
    display:flex!important;
    visibility:visible!important;
    opacity:1!important;
    position:fixed!important;
    top:50%!important;
    left:0!important;
    transform:translateY(-50%)!important;
    z-index:999999!important;
    background-color:#2f3343!important;
    border-radius:0 8px 8px 0!important;
    padding:12px 7px!important;
    box-shadow:3px 0 8px rgba(0,0,0,.35)!important;
    cursor:pointer!important;
}

[data-testid="collapsedControl"] button{
    background:transparent!important;
    border:none!important;
    padding:0!important;
}

[data-testid="collapsedControl"] svg{
    fill:white!important;
    color:white!important;
}

/* Buttons */
.stButton>button,
.stDownloadButton>button{
    border-radius:10px;
    font-weight:600;
}

/* File uploader */
[data-testid="stFileUploader"]{
    border-radius:10px;
}

/* Tables */
table{
    border-collapse:collapse;
}

/* Success boxes */
.stAlert{
    border-radius:10px;
}

</style>
""", unsafe_allow_html=True)
_render_nav()
# ==========================
# PAGE HEADER
# ==========================

col_logo, col_title = st.columns([1, 9])

with col_logo:
    st.image("assets/logo.png", width=75)

with col_title:
    st.markdown("""
    <h1 style="margin-bottom:0px;color:#1f2a44;font-size:48px;font-weight:700;">
        Division-wise Daily Monitoring
    </h1>
    <div style="font-size:22px;color:#6c757d;margin-top:6px;">
        Headquarters Region • Telangana Circle
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin-top:20px;margin-bottom:30px;'>", unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.header("📊 Reports")

show_lowest = st.sidebar.checkbox(
    "Division-wise Daily Monitoring",
    value=True
)

show_100 = st.sidebar.checkbox(
    "100% Deposit Offices",
    value=False
)
st.sidebar.markdown("---")
st.sidebar.subheader("📌 Minimum Invoiced Count")

min_bo = st.sidebar.number_input(
    "BPO Minimum Invoiced",
    min_value=1,
    value=30,
    step=1
)

min_other = st.sidebar.number_input(
    "SPO/HPO/GPO/IDC/NDC Minimum Invoiced",
    min_value=1,
    value=100,
    step=1
)
uploaded_file = None

if show_lowest or show_100:
    uploaded_file = st.file_uploader(
        "Upload Daily Monitoring CSV",
        type=["csv"]
    )

if uploaded_file is None:
    st.info("Please upload the Daily Monitoring CSV.")
    st.stop()
MASTER_FILE = "data/Delivery Productivity/Office Master.csv"
master = pd.read_csv(MASTER_FILE)
master.columns = master.columns.str.strip()

VALID_TYPES = [
    "BPO",
    "SPO",
    "HPO",
    "GPO",
    "IDC",
    "NDC"
]

master = master[
    master["office-type-code"].isin(VALID_TYPES)
].copy()

daily = pd.read_csv(uploaded_file)

daily.columns = daily.columns.str.strip()
#st.subheader("Detected Columns")
#st.write(daily.columns.tolist())
daily["office-id"] = daily["office-id"].astype(str)
master["office-id"] = master["office-id"].astype(str)

daily = daily.merge(
    master[
        [
            "office-id",
            "office-name",
            "division-office-name",
            "office-type-code"
        ]
    ],
    on="office-id",
    how="left"
)
# Keep Office Master office name
daily = daily.rename(columns={
    "office-name_y": "office-name"
})

# Remove uploaded office name
if "office-name_x" in daily.columns:
    daily = daily.drop(columns=["office-name_x"])
# --------------------------------------------------------
# Percentage Calculations
# --------------------------------------------------------

daily["Delivery %"] = (
    daily["delivery-count"] / daily["invoice-count"] * 100
).round(2)

daily["Deposit %"] = (
    daily["deposit-count"] / daily["invoice-count"] * 100
).round(2)

daily["Redirect %"] = (
    daily["redirection-count"] / daily["invoice-count"] * 100
).round(2)

daily["Return %"] = (
    daily["return-count"] / daily["invoice-count"] * 100
).round(2)

daily = daily.fillna(0)
# ==========================================================
# Division Selector
# ==========================================================

divisions = sorted(
    daily["division-office-name"]
    .dropna()
    .unique()
)

selected_division = st.selectbox(
    "Select Division",
    divisions
)

division_df = daily[
    daily["division-office-name"] == selected_division
].copy()

st.subheader(selected_division)
# ==========================================================
# Lowest Delivery % Report
# ==========================================================

division_data = division_df.copy()

# Split BO and Others
bo_df = division_data[
    division_data["office-type-code"] == "BPO"
].copy()

other_df = division_data[
    division_data["office-type-code"] != "BPO"
].copy()

# Apply minimum invoice criteria
bo_df = bo_df[
    bo_df["invoice-count"] >= min_bo
]

other_df = other_df[
    other_df["invoice-count"] >= min_other
]

# Lowest Delivery %
bo_delivery = (
    bo_df.sort_values(
        ["Delivery %", "invoice-count"],
        ascending=[True, False]
    )
    .head(25)
)

other_delivery = (
    other_df.sort_values(
        ["Delivery %", "invoice-count"],
        ascending=[True, False]
    )
    .head(15)
)
