import streamlit as st
import os
import glob as _glob

# ── Shared nav ────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f512 Login")
    _bulk = (_glob.glob("pages/Bulk Analytics.py") +
             _glob.glob("pages/*[Bb]ulk*.py"))
    if _bulk:
        st.sidebar.page_link(_bulk[0].replace("\\", "/"), label="\U0001f4ca Bulk Customer Analytics")
    _posb = (_glob.glob("pages/POSB Daily Report.py") +
             _glob.glob("pages/*[Pp][Oo][Ss][Bb]*.py"))
    if _posb:
        st.sidebar.page_link(_posb[0].replace("\\", "/"), label="\U0001f4ee POSB Daily Report")
    _dig = (_glob.glob("pages/1_Digital_Transactions.py") +
            _glob.glob("pages/*[Dd]igital*.py"))
    if _dig:
        st.sidebar.page_link(_dig[0].replace("\\", "/"), label="\U0001f4bb Digital Transactions")
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)
    _del = (_glob.glob("pages/Delivery Productivity.py") +
            _glob.glob("pages/*[Dd]elivery*.py"))
    if _del:
        st.sidebar.page_link(_del[0].replace("\\", "/"), label="\U0001f4e6 Delivery Productivity")
    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)

from PIL import Image

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Login – Analytics",
    page_icon="\U0001f512",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebarNav"]  { display: none !important; }
.block-container { padding-top: 1.2rem !important; padding-bottom: 0rem !important; }
header { visibility: hidden; height: 0px !important; }
.stTextInput input { height: 44px; border-radius: 8px; font-size: 16px; }
div.stButton > button {
    background-color: #ff4b4b; color: white; border: none;
    border-radius: 8px; width: 100%; height: 48px;
    font-size: 20px; font-weight: 600; margin-top: 8px;
}
div.stButton > button:hover { background-color: #e23d3d; color: white; }
[data-testid="collapsedControl"] {
    display: flex !important; visibility: visible !important; opacity: 1 !important;
    position: fixed !important; top: 50% !important; left: 0px !important;
    transform: translateY(-50%) !important; z-index: 999999 !important;
    background-color: #2f3343 !important; border-radius: 0 8px 8px 0 !important;
    padding: 12px 7px !important; box-shadow: 3px 0 8px rgba(0,0,0,0.35) !important;
    cursor: pointer !important;
}
[data-testid="collapsedControl"] button { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="collapsedControl"] svg    { fill: white !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "just_logged_in" not in st.session_state:
    st.session_state.just_logged_in = False

# ── Load logo ─────────────────────────────────────────────────────────────────
logo_path = "assets/logo.png"
logo = Image.open(logo_path) if os.path.exists(logo_path) else None

# ── If already logged in — show success banner + nav ─────────────────────────
if st.session_state.authenticated:
    _render_nav()
    st.markdown("""
    <div style='background:#e6f4ea;border-left:5px solid #34a853;padding:18px 24px;
                border-radius:8px;margin:40px auto;max-width:520px;text-align:center;'>
        <span style='font-size:36px;'>✅</span><br>
        <h2 style='color:#1e7e34;margin:8px 0 4px 0;font-size:24px;'>
            You are successfully logged in!</h2>
        <p style='color:#444;margin:0;font-size:15px;'>
            Use the sidebar to navigate to a page.</p>
    </div>
    """, unsafe_allow_html=True)

    # Show logout button
    st.markdown("<div style='text-align:center;margin-top:16px;'>", unsafe_allow_html=True)
    if st.button("🔓 Logout", key="logout_btn"):
        st.session_state.authenticated = False
        st.session_state.just_logged_in = False
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# ── Login page (not yet authenticated) ───────────────────────────────────────
# Hide sidebar entirely on the login screen
st.markdown("""
<style>
[data-testid="stSidebar"]        { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
</style>
""", unsafe_allow_html=True)

# Header
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
        Analytics (Business and Operations)</h1>
        <p style='font-size:15px;color:#555;margin-top:2px;margin-bottom:0;'>
        Headquarters Region – Telangana Postal Circle</p>
        </div>""", unsafe_allow_html=True)

st.markdown("<hr style='margin:10px 0 8px 0;border-color:#ddd;'>", unsafe_allow_html=True)

# Login form
l, c, r = st.columns([2.2, 1.4, 2.2])
with c:
    st.markdown(
        "<h2 style='text-align:center;font-size:32px;color:#2f3343;"
        "margin-top:4px;margin-bottom:12px;'>Login</h2>",
        unsafe_allow_html=True)

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
            st.session_state.authenticated  = True
            st.session_state.just_logged_in = True
            st.rerun()
        else:
            st.error("Invalid Username or Password")
