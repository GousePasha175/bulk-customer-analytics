import streamlit as st
import subprocess
import os

st.set_page_config(
    page_title="Regional Bulk Customer Analytics",
    layout="wide"
)

st.title("Regional Bulk Customer Business Tracking")

st.success("App Loaded Successfully")

if st.button("Run Analytics Engine"):
    try:
        result = subprocess.run(
            ["python", "Analytics_Excel.py"],
            capture_output=True,
            text=True
        )

        st.success("Analytics Executed")

        if result.stdout:
            st.text(result.stdout)

        if result.stderr:
            st.error(result.stderr)

    except Exception as e:
        st.error(str(e))