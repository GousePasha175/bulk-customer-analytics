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

_render_nav()

# --------------------------------------------------------
# Header
# --------------------------------------------------------

logo_path = "assets/logo.png"

if os.path.exists(logo_path):
    logo = Image.open(logo_path)
else:
    logo = None

c1, c2 = st.columns([1,8])

with c1:
    if logo:
        st.image(logo,width=90)

with c2:

    st.markdown("""
    <h1 style='margin-bottom:0px;color:#2f3343;'>
    Division-wise Daily Monitoring
    </h1>

    <p style='color:gray;'>
    Headquarters Region • Telangana Circle
    </p>
    """,unsafe_allow_html=True)

st.divider()

# --------------------------------------------------------
# Upload
# --------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Daily Monitoring CSV",
    type="csv"
)

if uploaded_file is None:

    st.info("Please upload the Daily Monitoring CSV.")

    st.stop()

# --------------------------------------------------------
# Read Upload
# --------------------------------------------------------

df = pd.read_csv(uploaded_file)

st.success(f"{len(df):,} records loaded successfully.")

st.dataframe(df.head())
