import streamlit as st
import pandas as pd
from datetime import date
from PIL import Image
import os
import glob as _glob
from io import BytesIO
import xlsxwriter

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
st.sidebar.subheader("📅 Report Period")
_dr1, _dr2 = st.sidebar.columns(2)
with _dr1:
    report_from = st.date_input("From", value=date.today())
with _dr2:
    report_to = st.date_input("To", value=date.today())
date_range_str = (
    report_from.strftime('%d.%m.%Y') if report_from == report_to
    else f"{report_from.strftime('%d.%m.%Y')} to {report_to.strftime('%d.%m.%Y')}"
)

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

enable_min_criteria = st.sidebar.checkbox(
    "Enable Minimum Invoiced Criteria",
    value=False,
    help="Off by default — shows all offices regardless of invoice count (minimum = 0). "
         "Turn on to set a minimum invoiced threshold below."
)

if enable_min_criteria:
    min_bo = st.sidebar.number_input(
        "BPO Minimum Invoiced",
        min_value=0,
        value=30,
        step=1
    )

    min_other = st.sidebar.number_input(
        "SPO/HPO/GPO/IDC/NDC Minimum Invoiced",
        min_value=0,
        value=100,
        step=1
    )
else:
    min_bo = 0
    min_other = 0
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
# Collapse to exactly one row per office
# --------------------------------------------------------
# The daily upload may contain more than one row per office (e.g. one row per
# postman/beat rather than one row per office). If left as-is, every count,
# KPI, and "top offices" ranking below would silently treat each such row as
# a separate office — inflating totals and letting the same office dominate
# multiple rows of a ranked list. This aggregation is a no-op when the file
# already has one row per office.
_sum_cols = [c for c in [
    "postman-count", "invoice-count", "delivery-count", "deposit-count",
    "return-count", "redirection-count", "beat-diversion-count"
] if c in daily.columns]
_first_cols = [c for c in daily.columns
               if c not in _sum_cols and c != "office-id"]

if daily["office-id"].duplicated().any():
    _dup_offices = daily["office-id"].duplicated().sum()
    st.info(
        f"ℹ️ The uploaded file had {_dup_offices} extra row(s) sharing an Office ID "
        "with another row (e.g. multiple postmen per office) — these were combined "
        "into a single row per office before computing any figures below."
    )

_agg_dict = {c: "sum" for c in _sum_cols}
_agg_dict.update({c: "first" for c in _first_cols})
daily = daily.groupby("office-id", as_index=False).agg(_agg_dict)

# --------------------------------------------------------
# Percentage Calculations
# --------------------------------------------------------

_inv_safe = daily["invoice-count"].replace(0, pd.NA)
daily["Delivery %"] = (daily["delivery-count"] / _inv_safe * 100).round(2)
daily["Deposit %"] = (daily["deposit-count"] / _inv_safe * 100).round(2)
daily["Redirect %"] = (daily["redirection-count"] / _inv_safe * 100).round(2)
daily["Return %"] = (daily["return-count"] / _inv_safe * 100).round(2)

# Fill missing values precisely rather than a blanket fillna(0), which would
# otherwise turn any unmatched division-office-name / office-type-code into
# the literal number 0 instead of a readable label.
for _c in ["Delivery %", "Deposit %", "Redirect %", "Return %"]:
    daily[_c] = daily[_c].fillna(0)
daily["division-office-name"] = daily["division-office-name"].fillna("Unmapped")
daily["office-type-code"] = daily["office-type-code"].fillna("UNK")
# ==========================================================
# ----------------------------------------------------------
# CSS
# ----------------------------------------------------------

st.markdown("""
<style>

.dm-title{
background:#0B5CAD;
padding:14px;
border-radius:10px;
color:white;
font-size:28px;
font-weight:700;
margin-bottom:18px;
}

.kpi{
background:white;
border-radius:14px;
padding:16px;
text-align:center;
box-shadow:0 2px 10px rgba(0,0,0,.10);
border-top:6px solid #0B5CAD;
margin-bottom:15px;
}

.kpi_head{
font-size:15px;
font-weight:600;
color:#666;
}

.kpi_val{
font-size:30px;
font-weight:700;
color:#0B5CAD;
}

.section{
background:#0B5CAD;
padding:10px 18px;
border-radius:8px;
color:white;
font-size:20px;
font-weight:bold;
margin-top:18px;
margin-bottom:10px;
}

.good{
background:#70AD47;
color:white;
font-weight:bold;
}

.medium{
background:#FFC000;
font-weight:bold;
}

.bad{
background:#FF4D4F;
color:white;
font-weight:bold;
}

</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------
# KPI CARD FUNCTION
# ----------------------------------------------------------

def kpi(title, value):

    st.markdown(f"""
    <div class="kpi">
        <div class="kpi_head">{title}</div>
        <div class="kpi_val">{value}</div>
    </div>
    """, unsafe_allow_html=True)
if show_lowest:
    # DAILY MONITORING REPORTING ENGINE - PART 1A
    # ==========================================================

    # ----------------------------------------------------------
    # Optional Settings
    # ----------------------------------------------------------

    st.sidebar.markdown("---")

    show_full = st.sidebar.checkbox(
        "Show Complete Office List",
        value=False
    )

    show_priority = st.sidebar.checkbox(
        "Include Priority Score",
        value=False,
        help="Priority Score will be added to Excel and tables."
    )

    critical_delivery = st.sidebar.slider(
        "Critical Delivery %",
        40,
        100,
        60
    )

    # ----------------------------------------------------------
    # Additional Calculations
    # ----------------------------------------------------------

    daily["Priority Score"] = (
        daily["invoice-count"] *
        (100 - daily["Delivery %"])
    ).round(0)

    daily["Undelivered"] = (
        daily["invoice-count"] -
        daily["delivery-count"]
    )

    # ----------------------------------------------------------
    # HQ SUMMARY
    # ----------------------------------------------------------

    division_summary = (
        daily.groupby("division-office-name", as_index=False)
        .agg(
            Offices=("office-id", "count"),
            Articles=("invoice-count", "sum"),
            Delivery=("Delivery %", "mean"),
            Deposit=("Deposit %", "mean"),
            Redirect=("Redirect %", "mean"),
            Return=("Return %", "mean"),
        )
    )

    division_summary = division_summary.round(2)

    # ----------------------------------------------------------
    # KPI VALUES
    # ----------------------------------------------------------

    total_divisions = division_summary.shape[0]

    total_offices = len(daily)

    total_articles = int(
        daily["invoice-count"].sum()
    )

    total_bo = (
        daily["office-type-code"] == "BPO"
    ).sum()

    total_other = total_offices - total_bo

    avg_delivery = daily["Delivery %"].mean()

    avg_deposit = daily["Deposit %"].mean()

    avg_redirect = daily["Redirect %"].mean()

    avg_return = daily["Return %"].mean()


    # ----------------------------------------------------------
    # HQ DASHBOARD
    # ----------------------------------------------------------

    st.markdown(
        f"<div class='dm-title'>🏛 Headquarters Region Dashboard — {date_range_str}</div>",
        unsafe_allow_html=True
    )

    r1, r2, r3, r4 = st.columns(4)

    with r1:
        kpi("Divisions", total_divisions)

    with r2:
        kpi("Offices", f"{total_offices:,}")

    with r3:
        kpi("Articles", f"{total_articles:,}")

    with r4:
        kpi("Branch Offices", f"{total_bo:,}")

    r5, r6, r7, r8 = st.columns(4)

    with r5:
        kpi("Average Delivery", f"{avg_delivery:.2f}%")

    with r6:
        kpi("Average Deposit", f"{avg_deposit:.2f}%")

    with r7:
        kpi("Average Redirect", f"{avg_redirect:.2f}%")

    with r8:
        kpi("Average Return", f"{avg_return:.2f}%")

    # ----------------------------------------------------------
    # TOP / BOTTOM DIVISIONS
    # ----------------------------------------------------------

    top_delivery = (
        division_summary
        .sort_values("Delivery", ascending=False)
        .head(3)
    )

    bottom_delivery = (
        division_summary
        .sort_values("Delivery", ascending=True)
        .head(3)
    )

    c1, c2 = st.columns(2)

    with c1:

        st.markdown(
            "<div class='section'>🏆 Top Delivery Divisions</div>",
            unsafe_allow_html=True
        )

    
        st.dataframe(
            top_delivery[
                [
                    "division-office-name",
                    "Delivery",
                    "Articles"
                ]
            ],
            hide_index=True,
            use_container_width=True
        )

    with c2:

        st.markdown(
            "<div class='section'>⚠ Bottom Delivery Divisions</div>",
            unsafe_allow_html=True
        )

        st.dataframe(
            bottom_delivery[
                [
                    "division-office-name",
                    "Delivery",
                    "Articles"
                ]
            ],
            hide_index=True,
            use_container_width=True
        )

    # ----------------------------------------------------------
    # DIVISION SELECTION
    # ----------------------------------------------------------

    divisions = sorted(
        daily["division-office-name"]
        .dropna()
        .unique()
    )

    selected_division = st.selectbox(
        "📍 Select Division",
        divisions
    )

    division_df = daily[
        daily["division-office-name"] == selected_division
    ].copy()
    # ==========================================================
    # DIVISION SUMMARY
    # ==========================================================

    bo_df = division_df[
        division_df["office-type-code"] == "BPO"
    ].copy()

    other_df = division_df[
        division_df["office-type-code"] != "BPO"
    ].copy()

    bo_df = bo_df[
        bo_df["invoice-count"] >= min_bo
    ]

    other_df = other_df[
        other_df["invoice-count"] >= min_other
    ]

    division_articles = int(
        division_df["invoice-count"].sum()
    )

    avg_del = division_df["Delivery %"].mean()

    avg_dep = division_df["Deposit %"].mean()

    avg_red = division_df["Redirect %"].mean()

    avg_ret = division_df["Return %"].mean()

    st.markdown(
        f"<div class='dm-title'>📍 {selected_division} — {date_range_str}</div>",
        unsafe_allow_html=True
    )

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        kpi("Total Offices",len(division_df))

    with c2:
        kpi("Branch Offices",len(bo_df))

    with c3:
        kpi("HO/SO/IDC/NDC",len(other_df))

    with c4:
        kpi("Articles",f"{division_articles:,}")

    c5,c6,c7,c8 = st.columns(4)

    with c5:
        kpi("Delivery",f"{avg_del:.2f}%")

    with c6:
        kpi("Deposit",f"{avg_dep:.2f}%")

    with c7:
        kpi("Redirect",f"{avg_red:.2f}%")

    with c8:
        kpi("Return",f"{avg_ret:.2f}%")

    # ----------------------------------------------------------
    # EXCEL EXPORT — Division Summary + selected Division detail
    # ----------------------------------------------------------
    _dm_xl_buf = BytesIO()
    with pd.ExcelWriter(_dm_xl_buf, engine="xlsxwriter") as _writer:
        _wb = _writer.book
        _title_fmt = _wb.add_format({"bold": True, "font_size": 13, "bg_color": "#0B5CAD",
                                      "font_color": "white", "align": "center",
                                      "valign": "vcenter", "border": 1})
        _hdr_fmt = _wb.add_format({"bold": True, "bg_color": "#2f3343",
                                    "font_color": "white", "border": 1,
                                    "align": "center", "valign": "vcenter", "text_wrap": True})
        _pct_fmt = _wb.add_format({"num_format": "0.00\"%\"", "border": 1, "align": "center"})
        _num_fmt = _wb.add_format({"num_format": "#,##0", "border": 1, "align": "center"})
        _cell_fmt = _wb.add_format({"border": 1})
        _total_fmt = _wb.add_format({"bold": True, "bg_color": "#FFD700", "border": 1, "align": "center"})
        _total_lft = _wb.add_format({"bold": True, "bg_color": "#FFD700", "border": 1})

        def _write_df(sheet_name, df_out, title):
            ncols = len(df_out.columns)
            ws = _wb.add_worksheet(sheet_name)
            _writer.sheets[sheet_name] = ws

            ws.merge_range(0, 0, 0, ncols - 1, title, _title_fmt)
            ws.set_row(0, 24)
            for ci, col in enumerate(df_out.columns):
                ws.write(1, ci, col, _hdr_fmt)
            ws.set_row(1, 30)

            for ri, (_, row) in enumerate(df_out.iterrows(), start=2):
                for ci, col in enumerate(df_out.columns):
                    val = row[col]
                    if "%" in col:
                        ws.write_number(ri, ci, float(val), _pct_fmt)
                    elif isinstance(val, (int, float)) and col != "Office" and col != "Type" and col != "Division":
                        ws.write_number(ri, ci, val, _num_fmt)
                    else:
                        ws.write(ri, ci, val, _cell_fmt)

            for ci, col in enumerate(df_out.columns):
                max_data_len = df_out[col].astype(str).map(len).max() if len(df_out) else 0
                width = max(max_data_len, len(col)) + 4
                ws.set_column(ci, ci, min(width, 45))

            ws.freeze_panes(2, 0)
            ws.autofilter(1, 0, 1 + len(df_out), ncols - 1)

        _write_df("Division Summary", division_summary.rename(columns={
            "division-office-name": "Division",
            "Offices": "Offices", "Articles": "Articles",
            "Delivery": "Delivery %", "Deposit": "Deposit %",
            "Redirect": "Redirect %", "Return": "Return %",
        }), f"Division-wise Daily Monitoring — {date_range_str}")

        _div_detail_cols = {
            "office-name": "Office", "office-type-code": "Type",
            "invoice-count": "Invoiced", "delivery-count": "Delivered",
            "Delivery %": "Delivery %", "deposit-count": "Deposited",
            "Deposit %": "Deposit %", "redirection-count": "Redirected",
            "Redirect %": "Redirect %", "return-count": "Returned",
            "Return %": "Return %",
        }
        _div_detail = division_df[list(_div_detail_cols.keys())].rename(columns=_div_detail_cols)
        _sheet_name = selected_division[:31].replace("/", "-")
        _write_df(_sheet_name, _div_detail, f"{selected_division} — {date_range_str}")

    st.download_button(
        "⬇ Download Excel (Daily Monitoring)",
        data=_dm_xl_buf.getvalue(),
        file_name=f"Daily_Monitoring_{selected_division.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ==========================================================
    # TABS
    # ==========================================================

    delivery_tab,deposit_tab,redirect_tab,return_tab = st.tabs(
    [
    "🚚 Delivery",
    "📥 Deposit",
    "🔀 Redirect",
    "↩ Return"
    ]
    )

    # ==========================================================
    # REPORT FUNCTION
    # ==========================================================

    def render_report(df_other, df_bo, metric, count_col, title, ascending=False):

        other = (
            df_other
            .sort_values([metric, "invoice-count"], ascending=[ascending, False])
            .head(15)
        )

        bo = (
            df_bo
            .sort_values([metric, "invoice-count"], ascending=[ascending, False])
            .head(25)
        )

        left, right = st.columns(2)

        def show_table(df, heading):

            st.markdown(f"### {heading}")

            display = pd.DataFrame({
                "Office": df["office-name"],
                "Invoiced": df["invoice-count"],
                "Delivered": df["delivery-count"]
            })

            if count_col != "delivery-count":
                display[title] = df[count_col]

            display[metric] = df[metric]

            if show_priority:
                display["Priority"] = df["Priority Score"].astype(int)

            st.dataframe(
                display,
                hide_index=True,
                use_container_width=True
            )

        with left:
            show_table(other, f"🏤 Head / Sub Offices ({len(other)})")

        with right:
            show_table(bo, f"🏣 Branch Offices ({len(bo)})")

        if show_full:

            st.markdown("---")

            st.subheader("Complete Office List")

            st.dataframe(
                pd.concat([other, bo]),
                hide_index=True,
                use_container_width=True
            )
    # ==========================================================
    # DELIVERY TAB
    # ==========================================================

    with delivery_tab:

        st.markdown("""
        ### 🚚 Lowest Delivery Percentage

        Offices are arranged in ascending order of **Delivery %**
        after applying the minimum invoiced article criteria.
        """)

        render_report(
            other_df,
            bo_df,
            "Delivery %",
            "delivery-count",
            "Delivered",
            ascending=True
        )

    # ==========================================================
    # DEPOSIT TAB
    # ==========================================================

    with deposit_tab:

        st.markdown("""
        ### 📥 Highest Deposit Percentage

        Offices are arranged in descending order of **Deposit %**
        after applying the minimum invoiced article criteria.
        """)

        render_report(
            other_df,
            bo_df,
            "Deposit %",
            "deposit-count",
            "Deposited",
            ascending=False
        )

    # ==========================================================
    # REDIRECT TAB
    # ==========================================================

    with redirect_tab:

        st.markdown("""
        ### 🔀 Highest Redirect Percentage

        Offices are arranged in descending order of **Redirect %**
        after applying the minimum invoiced article criteria.
        """)

        render_report(
            other_df,
            bo_df,
            "Redirect %",
            "redirection-count",
            "Redirected",
            ascending=False
        )

    # ==========================================================
    # RETURN TAB
    # ==========================================================

    with return_tab:

        st.markdown("""
        ### ↩ Highest Return Percentage

        Offices are arranged in descending order of **Return %**
        after applying the minimum invoiced article criteria.
        """)

        render_report(
            other_df,
            bo_df,
            "Return %",
            "return-count",
            "Returned",
            ascending=False
        )
else:
    st.info("Enable 'Division-wise Daily Monitoring' from the sidebar to view this report.")

# ==========================================================
# 100% DEPOSIT OFFICES (RED-FLAG LIST)
# ==========================================================
# Offices where Deposit % = 100 — every invoiced article was deposited back
# and nothing delivered. Independent of the dashboard above: works whether
# or not "Division-wise Daily Monitoring" is also enabled.

if show_100:
    st.markdown("<hr style='margin-top:30px;margin-bottom:20px;'>", unsafe_allow_html=True)
    st.markdown(
        f"<div class='dm-title'>🚩 100% Deposit Offices (Nothing Delivered) — {date_range_str}</div>",
        unsafe_allow_html=True
    )

    red_flag = daily[daily["Deposit %"] == 100].copy()
    red_flag = red_flag.sort_values(
        ["division-office-name", "invoice-count"], ascending=[True, False]
    )

    if red_flag.empty:
        st.success("✅ No offices found with 100% Deposit — nothing to flag today.")
    else:
        rf1, rf2, rf3 = st.columns(3)
        with rf1:
            kpi("Flagged Offices", len(red_flag))
        with rf2:
            kpi("Divisions Affected", red_flag["division-office-name"].nunique())
        with rf3:
            kpi("Articles Involved", f"{int(red_flag['invoice-count'].sum()):,}")

        red_flag_display = pd.DataFrame({
            "Division": red_flag["division-office-name"],
            "Office": red_flag["office-name"],
            "Type": red_flag["office-type-code"],
            "Invoiced": red_flag["invoice-count"],
            "Deposited": red_flag["deposit-count"],
            "Deposit %": red_flag["Deposit %"],
        })

        st.dataframe(red_flag_display, hide_index=True, use_container_width=True)

        # Excel download
        xl_buf = BytesIO()
        with pd.ExcelWriter(xl_buf, engine="xlsxwriter") as writer:
            wb = writer.book
            ncols = len(red_flag_display.columns)
            ws = wb.add_worksheet("100pct Deposit")
            writer.sheets["100pct Deposit"] = ws

            title_fmt = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#cc0000",
                                        "font_color": "white", "align": "center",
                                        "valign": "vcenter", "border": 1})
            hdr_fmt = wb.add_format({"bold": True, "bg_color": "#2f3343",
                                      "font_color": "white", "border": 1,
                                      "align": "center", "valign": "vcenter"})
            pct_fmt = wb.add_format({"num_format": "0.00\"%\"", "border": 1, "align": "center"})
            num_fmt = wb.add_format({"num_format": "#,##0", "border": 1, "align": "center"})
            cell_fmt = wb.add_format({"border": 1})

            ws.merge_range(0, 0, 0, ncols - 1,
                            f"100% Deposit Offices (Nothing Delivered) — {date_range_str}", title_fmt)
            ws.set_row(0, 24)
            for ci, col in enumerate(red_flag_display.columns):
                ws.write(1, ci, col, hdr_fmt)

            for ri, (_, row) in enumerate(red_flag_display.iterrows(), start=2):
                for ci, col in enumerate(red_flag_display.columns):
                    val = row[col]
                    if "%" in col:
                        ws.write_number(ri, ci, float(val), pct_fmt)
                    elif isinstance(val, (int, float)):
                        ws.write_number(ri, ci, val, num_fmt)
                    else:
                        ws.write(ri, ci, val, cell_fmt)

            for ci, col in enumerate(red_flag_display.columns):
                max_data_len = red_flag_display[col].astype(str).map(len).max() if len(red_flag_display) else 0
                width = max(max_data_len, len(col)) + 4
                ws.set_column(ci, ci, min(width, 45))

            ws.freeze_panes(2, 0)
            ws.autofilter(1, 0, 1 + len(red_flag_display), ncols - 1)

        st.download_button(
            "⬇ Download 100% Deposit Offices (Excel)",
            data=xl_buf.getvalue(),
            file_name=f"100_Percent_Deposit_Offices_{report_from.strftime('%d%m%Y')}_{report_to.strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
