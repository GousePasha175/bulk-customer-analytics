import streamlit as st

st.set_page_config(
    page_title="Delivery Productivity",
    page_icon="📮",
    layout="wide"
)

st.title("📮 Delivery Productivity")

st.sidebar.header("Select Module")

module = st.sidebar.radio(
    "Choose a Report",
    [
        "Delivery Transit Analysis",
        "Division-wise Daily Monitoring",
        "100% Deposit Offices"
    ]
)

st.write("---")

if module == "Delivery Transit Analysis":
    st.info("Transit Analysis module will come here.")

elif module == "Division-wise Daily Monitoring":
    st.info("Daily Monitoring module will come here.")

elif module == "100% Deposit Offices":
    st.info("100% Deposit Offices module will come here.")
