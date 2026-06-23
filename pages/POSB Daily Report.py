import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
from datetime import date, datetime
import calendar
import glob as _glob

import glob as _glob

def _render_nav():
    st.sidebar.markdown(
        """<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)
    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f3e0 Home")
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


# ── Auth guard: redirect to Home (login) if not authenticated ─────────────────
# set_page_config must be the very first Streamlit call
st.set_page_config(
    page_title="POSB Accounts Daily Report",
    page_icon="📊",
    layout="wide",
)

st.markdown("""<style>
[data-testid="stSidebarNav"] { display: none !important; }
[data-testid="collapsedControl"] {
    display: flex !important; visibility: visible !important; opacity: 1 !important;
    position: fixed !important; top: 50% !important; left: 0px !important;
    transform: translateY(-50%) !important; z-index: 999999 !important;
    background-color: #2f3343 !important; border-radius: 0 8px 8px 0 !important;
    padding: 12px 7px !important; box-shadow: 3px 0 8px rgba(0,0,0,0.35) !important;
    cursor: pointer !important;
}
[data-testid="collapsedControl"] button { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="collapsedControl"] svg { fill: white !important; color: white !important; }
</style>""", unsafe_allow_html=True)

if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown(
        "Please go to **🏠 Home** in the sidebar to log in.",
        unsafe_allow_html=False,
    )
    with st.sidebar:
        _render_nav()
    st.stop()

# ─── Constants ────────────────────────────────────────────────────────────────
DIVISIONS = [
    "Hyderabad City",
    "Hyderabad GPO",
    "Hyderabad South East",
    "Secunderabad",
    "Medak",
    "Sangareddy",
]

TARGETS_FY = {
    "Hyderabad City":       115000,
    "Hyderabad GPO":         27000,
    "Hyderabad South East": 177000,
    "Secunderabad":         267000,
    "Medak":                146000,
    "Sangareddy":           138000,
}

ACCOUNT_COLS  = ["MIS", "PPFGP", "SSA", "RD", "SBBAS", "SBSGP", "SCSS", "TD"]
CERT_COLS     = ["KVN", "NSC8", "MSSC"]
ALL_PROD_COLS = ACCOUNT_COLS + CERT_COLS

SCHEME_FULL = {
    "MIS":   "Monthly Income Scheme",
    "PPFGP": "Public Provident Fund",
    "SSA":   "Sukanya Samriddhi Account",
    "RD":    "Recurring Deposit",
    "SBBAS": "Savings Bank Basic",
    "SBSGP": "Savings Bank General",
    "SCSS":  "Senior Citizen Savings Scheme",
    "TD":    "Time Deposit",
    "KVN":   "Kisan Vikas Patra",
    "NSC8":  "National Savings Certificate",
    "MSSC":  "Mahila Samman Savings Certificate",
}

OFFICE_SUFFIXES = {"S.O", "B.O", "H.O", "DC"}


def classify_office(name: str) -> str:
    """Return office class (S.O / B.O / H.O / DC) or '' if unrecognised."""
    if not isinstance(name, str):
        return ""
    parts = name.strip().split()
    if not parts:
        return ""
    last = parts[-1].rstrip(".")
    # normalise to upper for matching
    for suf in OFFICE_SUFFIXES:
        # compare last word (strip trailing dots) case-insensitively
        if last.upper() == suf.replace(".", "").upper() or parts[-1].upper() == suf.upper():
            return suf
    return ""


def is_dc(name: str) -> bool:
    return classify_office(name) == "DC"


# ─── Excel parsing helpers ─────────────────────────────────────────────────────

def parse_product_report(uploaded_file, division_name: str) -> pd.DataFrame | None:
    """Parse a Product Wise A/C Report Excel for a division."""
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        st.error(f"Cannot read {division_name} file: {e}")
        return None

    # Find header row (contains 'Name')
    hdr_row = None
    for i, row in df.iterrows():
        if any(str(c).strip().lower() == "name" for c in row.values):
            hdr_row = i
            break
    if hdr_row is None:
        st.error(f"Could not find header row in {division_name} file.")
        return None

    df.columns = df.iloc[hdr_row]
    df = df.iloc[hdr_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    # Drop total row
    df = df[~df["Name"].astype(str).str.lower().str.startswith("total")]
    # Drop DC offices
    df = df[~df["Name"].apply(is_dc)]
    # Drop empty names
    df = df[df["Name"].astype(str).str.strip() != ""]
    df = df[df["Name"].astype(str).str.strip().str.lower() != "nan"]

    # Coerce product columns to numeric
    for col in ALL_PROD_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        else:
            df[col] = 0

    df["Division"] = division_name
    return df


def parse_summary_excel(uploaded_file, value_cols: list[str]) -> pd.DataFrame | None:
    """Parse a summary-level Excel (Accounts Opened Details or Net Addition)."""
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        st.error(f"Cannot read summary file: {e}")
        return None

    hdr_row = None
    for i, row in df.iterrows():
        if any(str(c).strip().lower() == "name" for c in row.values):
            hdr_row = i
            break
    if hdr_row is None:
        hdr_row = 0

    # Build column names from header row, de-duplicating if needed
    raw_cols = [str(c).strip() for c in df.iloc[hdr_row].tolist()]
    seen = {}
    unique_cols = []
    for c in raw_cols:
        if c in seen:
            seen[c] += 1
            unique_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            unique_cols.append(c)

    df = df.iloc[hdr_row + 1:].reset_index(drop=True)
    df.columns = unique_cols

    df = df[~df["Name"].astype(str).str.lower().str.startswith("total")]
    df = df[df["Name"].astype(str).str.strip() != ""]
    df = df[df["Name"].astype(str).str.strip().str.lower() != "nan"]

    # For Net Addition: the first "Total" column (index 3) = net accounts
    # value_cols that have duplicates will be suffixed; map them to position-based access
    for col in value_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        elif f"{col}_1" in df.columns:
            # use first occurrence via positional slice
            pass
        else:
            df[col] = 0

    # Special handling for Net Addition file: ensure "Total" = first Total col (net a/cs)
    if "Total" not in df.columns and "Total_1" in df.columns:
        # The deduplicated header means first Total is still named "Total"
        # This shouldn't happen, but handle gracefully
        df["Total"] = pd.to_numeric(df["Total_1"], errors="coerce").fillna(0).astype(int)

    return df


# ─── Proportionate / Daily Target logic ────────────────────────────────────────

MONTH_TO_IDX = {
    "April": 1, "May": 2, "June": 3, "July": 4,
    "August": 5, "September": 6, "October": 7, "November": 8,
    "December": 9, "January": 10, "February": 11, "March": 12,
}

def proportionate_target(annual_target: int, month_name: str) -> int:
    m = MONTH_TO_IDX.get(month_name, 1)
    return round(annual_target * m / 12)


def daily_target(annual_target: int, month_name: str,
                 report_date: date, accounts_opened_so_far: int) -> int:
    prop_target = proportionate_target(annual_target, month_name)
    remaining = prop_target - accounts_opened_so_far
    total_days = calendar.monthrange(report_date.year, report_date.month)[1]
    days_remaining = total_days - report_date.day + 1
    if days_remaining <= 0:
        return int(remaining)
    return max(0, int(round(remaining / days_remaining)))


# ─── Report 1 – Office-wise Range Report ───────────────────────────────────────

RANGES = ["0", "1-10", "11-25", "26-50", "51-100", "100+"]

def classify_range(total: int) -> str:
    if total == 0:      return "0"
    elif total <= 10:   return "1-10"
    elif total <= 25:   return "11-25"
    elif total <= 50:   return "26-50"
    elif total <= 100:  return "51-100"
    else:               return "100+"


def build_range_report(division_dfs: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    sl = 1
    for div, df in division_dfs.items():
        df["_total_ac"] = df[ACCOUNT_COLS].sum(axis=1)
        total_ac = int(df["_total_ac"].sum())
        counts = {r: 0 for r in RANGES}
        for _, row in df.iterrows():
            counts[classify_range(int(row["_total_ac"]))] += 1
        rows.append({
            "Sl. No": sl,
            "Division": div,
            "Total Accounts Opened": total_ac,
            **{r: counts[r] for r in RANGES},
        })
        sl += 1
    return pd.DataFrame(rows)


def export_range_report_excel(df, division_dfs=None):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    # ── Shared formats (professional palette) ────────────────────────────────
    def f(**kw):
        base = {"border": 1, "valign": "vcenter"}; base.update(kw)
        return wb.add_format(base)
    title_fmt     = f(bold=True, font_size=13, align="center",   bg_color="#1F3864", font_color="#FFFFFF")
    header_fmt    = f(bold=True, align="center", text_wrap=True, bg_color="#2E75B6", font_color="#FFFFFF")
    sub_hdr_fmt   = f(bold=True, align="center",                 bg_color="#9DC3E6", font_color="#000000")
    div_title_fmt = f(bold=True, font_size=11, align="left",     bg_color="#2E75B6", font_color="#FFFFFF")
    data_fmt      = f(align="center")
    data_left_fmt = f(align="left")
    total_fmt     = f(bold=True, align="center", bg_color="#FFF2CC", font_color="#000000")
    total_left_fmt= f(bold=True, align="left",   bg_color="#FFF2CC", font_color="#000000")
    range_color   = {"0":"#D9D9D9","1-10":"#DDEBF7","11-25":"#E2EFDA","26-50":"#FFF2CC","51-100":"#FCE4D6","100+":"#F4CCCC"}

    # ── Sheet 1: Range Summary ────────────────────────────────────────────────
    ws = wb.add_worksheet("Range Summary")
    nr = len(RANGES); tc = 3 + nr
    ws.merge_range(0,0,0,tc-1, "Office wise Range of Accounts Opened – Division wise Report", title_fmt)
    ws.write(1,0,"Sl. No",header_fmt); ws.write(1,1,"Division",header_fmt)
    ws.write(1,2,"Total Accounts\nOpened",header_fmt)
    ws.merge_range(1,3,1,3+nr-1,"Office wise Range of Accounts Opened",header_fmt)
    ws.write(2,0,"",sub_hdr_fmt); ws.write(2,1,"",sub_hdr_fmt); ws.write(2,2,"",sub_hdr_fmt)
    for i,r in enumerate(RANGES): ws.write(2,3+i,r,sub_hdr_fmt)
    ws.set_column(0,0,6); ws.set_column(1,1,26); ws.set_column(2,2,18)
    ws.set_column(3,3+nr-1,10)
    ri=3
    for _,row in df.iterrows():
        ws.write(ri,0,row["Sl. No"],data_fmt); ws.write(ri,1,row["Division"],data_left_fmt)
        ws.write(ri,2,row["Total Accounts Opened"],data_fmt)
        for i,r in enumerate(RANGES): ws.write(ri,3+i,row[r],data_fmt)
        ri+=1
    ws.write(ri,0,"",total_fmt); ws.write(ri,1,"TOTAL",total_left_fmt)
    ws.write(ri,2,df["Total Accounts Opened"].sum(),total_fmt)
    for i,r in enumerate(RANGES): ws.write(ri,3+i,df[r].sum(),total_fmt)

    # ── Sheet 2: Detailed Office-wise Breakdown ───────────────────────────────
    if division_dfs:
        ws2 = wb.add_worksheet("Office-wise Breakdown")
        det_cols = ["Sl. No","Office Name","Office Type"] + ACCOUNT_COLS + ["Total Accounts","Range"]
        nc2 = len(det_cols)
        ws2.merge_range(0,0,0,nc2-1,"Detailed Office-wise Breakdown – All Divisions",title_fmt)
        ws2.set_row(0,22)
        ws2.set_column(0,0,6); ws2.set_column(1,1,35); ws2.set_column(2,2,10)
        ws2.set_column(3,3+len(ACCOUNT_COLS)-1,8)
        ws2.set_column(3+len(ACCOUNT_COLS),3+len(ACCOUNT_COLS),14)
        ws2.set_column(3+len(ACCOUNT_COLS)+1,3+len(ACCOUNT_COLS)+1,8)
        cur_row = 1; overall_sl = 1
        for div_name, div_df in division_dfs.items():
            ws2.merge_range(cur_row,0,cur_row,nc2-1,f"{div_name} Division",div_title_fmt)
            ws2.set_row(cur_row,18); cur_row+=1
            for ci,col in enumerate(det_cols): ws2.write(cur_row,ci,col,sub_hdr_fmt)
            cur_row+=1
            div_df = div_df.copy()
            div_df["_total_ac"] = div_df[ACCOUNT_COLS].sum(axis=1)
            div_df["_range"]    = div_df["_total_ac"].apply(classify_range)
            div_df["_otype"]    = div_df["Name"].apply(classify_office)
            div_sl=1
            for _,orow in div_df.iterrows():
                rng = orow["_range"]
                rc  = range_color.get(rng,"#FFFFFF")
                rng_fmt = f(align="center",bg_color=rc)
                ws2.write(cur_row,0,div_sl,data_fmt)
                ws2.write(cur_row,1,str(orow["Name"]),data_left_fmt)
                ws2.write(cur_row,2,str(orow["_otype"]),data_fmt)
                for ci,col in enumerate(ACCOUNT_COLS): ws2.write(cur_row,3+ci,int(orow.get(col,0)),data_fmt)
                ws2.write(cur_row,3+len(ACCOUNT_COLS),  int(orow["_total_ac"]),data_fmt)
                ws2.write(cur_row,3+len(ACCOUNT_COLS)+1,rng,rng_fmt)
                cur_row+=1; div_sl+=1; overall_sl+=1
            # subtotal row
            ws2.write(cur_row,0,"",total_fmt); ws2.write(cur_row,1,f"Sub-total – {div_name}",total_left_fmt)
            ws2.write(cur_row,2,"",total_fmt)
            for ci,col in enumerate(ACCOUNT_COLS): ws2.write(cur_row,3+ci,int(div_df[col].sum()),total_fmt)
            ws2.write(cur_row,3+len(ACCOUNT_COLS),int(div_df["_total_ac"].sum()),total_fmt)
            ws2.write(cur_row,3+len(ACCOUNT_COLS)+1,"",total_fmt)
            cur_row+=2

    wb.close(); return output.getvalue()

def _sum_ac_cols(df, div):
    """
    For Accounts Opened files: sum MIS+PPFGP+SSA+RD+SBBAS+SBSGP+SCSS+TD
    (columns at positions 1-8 after Name in the Product Wise A/C Report format).
    Excludes certificates (PRFTS, KVN, NSC8, MSSC) which are in cols 9+.
    """
    if df is None or df.empty: return 0
    row = df[df["Name"].str.contains(div, case=False, na=False)]
    if row.empty: return 0
    dcols = [c for c in row.columns if c != "Name"]
    total = 0
    for c in dcols[:8]:   # first 8 data cols = MIS through TD
        v = pd.to_numeric(row[c].iloc[0], errors="coerce")
        if not pd.isna(v): total += int(v)
    return max(0, total)


def _net_total_col(df, div):
    """
    For Net Addition files: read the first 'Total' column = A/c Opened - A/c Closed.
    Column index 3 in the raw file (after Name, A/c Opened, A/c Closed).
    The de-duplicated header names it 'Total' (first occurrence).
    """
    if df is None or df.empty: return 0
    row = df[df["Name"].str.contains(div, case=False, na=False)]
    if row.empty: return 0
    # 'Total' is the first occurrence after deduplication in parse_summary_excel
    if "Total" in df.columns:
        v = pd.to_numeric(row["Total"].iloc[0], errors="coerce")
        return int(v) if not pd.isna(v) else 0
    # Positional fallback: 3rd data col after Name
    dcols = [c for c in row.columns if c != "Name"]
    if len(dcols) >= 3:
        v = pd.to_numeric(row[dcols[2]].iloc[0], errors="coerce")
        return int(v) if not pd.isna(v) else 0
    return 0


def build_daily_summary(
    ao_date_df,      # accounts opened ON report date
    ao_cumul_df,     # accounts opened UP TO report date (cumulative)
    net_date_df,     # net accounts ON report date
    net_cumul_df,    # net accounts UP TO report date (cumulative FY)
    report_date: date,
    month_name: str,
    working_days_left: int,
) -> pd.DataFrame:
    rows = []
    for div in DIVISIONS:
        annual = TARGETS_FY[div]
        prop   = proportionate_target(annual, month_name)

        opened_today     = _sum_ac_cols(ao_date_df,  div) if ao_date_df  is not None else 0
        opened_till_date = _sum_ac_cols(ao_cumul_df, div) if ao_cumul_df is not None else 0
        net_today        = _net_total_col(net_date_df, div) if net_date_df is not None else 0
        net_till_date    = _net_total_col(net_cumul_df, div) if net_cumul_df is not None else 0

        # Daily target = (Proportionate target − Net addition cumulative) ÷ working days left
        balance   = max(0, prop - net_till_date)
        daily_tgt = max(0, int(round(balance / working_days_left))) if working_days_left > 0 else balance

        shortfall_daily = max(0, daily_tgt - opened_today)
        shortfall_prop  = max(0, prop - net_till_date)
        pct_prop        = round((net_till_date / prop * 100), 0) if prop > 0 else 0

        rows.append({
            "Division": div,
            "Target FY 2026-27": annual,
            f"Proportionate Target upto {month_name}, {report_date.year}": prop,
            f"Daily Target upto {report_date.strftime('%d.%m.%Y')}": daily_tgt,
            f"No. of Accounts Opened on {report_date.strftime('%d.%m.%Y')}": opened_today,
            f"No. of Accounts Opened up to {report_date.strftime('%d.%m.%Y')}": opened_till_date,
            f"Net no. of a/cs opened on {report_date.strftime('%d.%m.%Y')}": net_today,
            f"Net no. of a/cs opened upto {report_date.strftime('%d.%m.%Y')}": net_till_date,
            "Shortfall on daily target": shortfall_daily,
            "Shortfall on proportionate target": shortfall_prop,
            "% achievement of proportionate Target": int(pct_prop),
        })

    total_row = {"Division": "Total HQ Region"}
    for col in rows[0]:
        if col == "Division": continue
        if "%" in col:
            prop_sum = sum(proportionate_target(TARGETS_FY[d], month_name) for d in DIVISIONS)
            net_sum  = sum(r[f"Net no. of a/cs opened upto {report_date.strftime('%d.%m.%Y')}"] for r in rows)
            total_row[col] = int(round(net_sum / prop_sum * 100, 0)) if prop_sum else 0
        elif "Daily Target" in col:
            # Sum of individual daily targets is meaningless; recompute for total
            total_balance   = max(0, sum(proportionate_target(TARGETS_FY[d], month_name) for d in DIVISIONS)
                                   - sum(r[f"Net no. of a/cs opened upto {report_date.strftime('%d.%m.%Y')}"] for r in rows))
            total_row[col]  = max(0, int(round(total_balance / working_days_left))) if working_days_left > 0 else total_balance
        else:
            total_row[col] = sum(r[col] for r in rows)
    rows.append(total_row)
    return pd.DataFrame(rows)


def build_scheme_wise(accounts_opened_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for div in DIVISIONS:
        annual = TARGETS_FY[div]
        ao_row = accounts_opened_df[
            accounts_opened_df["Name"].str.contains(div, case=False, na=False)
        ]
        row_data = {"Division": div, "Target": annual}
        if not ao_row.empty:
            for col in ACCOUNT_COLS:
                row_data[col] = int(ao_row[col].iloc[0]) if col in ao_row.columns else 0
            row_data["Total POSB"] = sum(row_data.get(c, 0) for c in ACCOUNT_COLS)
        else:
            for col in ACCOUNT_COLS:
                row_data[col] = 0
            row_data["Total POSB"] = 0
        rows.append(row_data)

    total = {"Division": "Total HQ Region", "Target": sum(TARGETS_FY[d] for d in DIVISIONS)}
    for col in ACCOUNT_COLS + ["Total POSB"]:
        total[col] = sum(r[col] for r in rows)
    rows.append(total)
    return pd.DataFrame(rows)


# ─── Streamlit UI ─────────────────────────────────────────────────────────────

def main():
    # set_page_config already called at module level above

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        # ── Navigation ───────────────────────────────────────────────────────
        _render_nav()
        report_option = st.radio(
            "Select Report",
            ["Office wise Range Report", "Division wise Summary Reports"]
            )
    
        # ─────────────────────────────────────────────────────────────────────

        # ── Logo (use local assets/logo.png if available, else text only) ──────
        import os as _os
        if _os.path.exists("assets/logo.png"):
            st.image("assets/logo.png", width=80)
        st.title("📮 POSB Daily Report")
        st.markdown("---")
        st.subheader("📅 Report Parameters")

        report_date = st.date_input("Report Date", value=date.today())

        month_options = list(MONTH_TO_IDX.keys())
        # Default to report_date month name
        report_month_name = report_date.strftime("%B")
        # Map calendar month to FY month
        cal_to_fy = {
            4: "April", 5: "May", 6: "June", 7: "July",
            8: "August", 9: "September", 10: "October", 11: "November",
            12: "December", 1: "January", 2: "February", 3: "March",
        }
        default_fy_month = cal_to_fy.get(report_date.month, "April")
        report_month = st.selectbox(
            "Report Month (FY)", month_options,
            index=month_options.index(default_fy_month)
        )

        st.markdown("---")
        st.subheader("⚙️ Daily Target Settings")
        import calendar as _cal
        _days_in_month = _cal.monthrange(report_date.year, report_date.month)[1]
        _days_remaining_cal = _days_in_month - report_date.day + 1
        working_days_left = st.number_input(
            "Working days left in month",
            min_value=1, max_value=31,
            value=max(1, _days_remaining_cal),
            step=1,
            help="Enter the number of working days remaining including today. "
                 "Daily target = (Proportionate Target – Net Addition as on date) ÷ working days left"
        )

        # st.markdown("---")
        # st.subheader("📂 Upload Division Files")
        # st.caption("Product Wise A/C Report – one per Division")

        # div_files = {}
        # for div in DIVISIONS:
        #     div_files[div] = st.file_uploader(
        #         f"{div}", type=["xlsx", "xls"], key=f"div_{div}"
        #     )

        # st.markdown("---")
        # st.subheader("📂 Upload Summary Files")
        # st.caption(
        #     "Upload four files for the Division-wise Summary Report. "
        #     "Each file: Column 1 = Name, Column 2 = Accounts Opened / Net Accounts."
        # )
        # ao_date_file = st.file_uploader(
        #     f"Accounts Opened on {report_date.strftime('%d.%m.%Y')} (daily)",
        #     type=["xlsx", "xls"], key="ao_date_file"
        # )
        # ao_file = st.file_uploader(
        #     f"Accounts Opened up to {report_date.strftime('%d.%m.%Y')} (cumulative)",
        #     type=["xlsx", "xls"], key="ao_file"
        # )
        # net_date_file = st.file_uploader(
        #     f"Net Accounts on {report_date.strftime('%d.%m.%Y')} (daily)",
        #     type=["xlsx", "xls"], key="net_date_file"
        # )
        # net_file = st.file_uploader(
        #     f"Net Accounts up to {report_date.strftime('%d.%m.%Y')} (cumulative FY)",
        #     type=["xlsx", "xls"], key="net_file"
        # )

    # ── Main Area ─────────────────────────────────────────────────────────────
    st.title("📊 POSB Accounts Daily Report")
    st.caption(f"Hyderabad HQ Region  |  Report Date: **{report_date.strftime('%d.%m.%Y')}**  |  Month: **{report_month}**")

    #Removed for making on same page
    #tab1, tab2 = st.tabs(["📁 Office-wise Range Report", "📈 Division-wise Summary Reports"])
    
    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 – Range Report
    # ════════════════════════════════════════════════════════════════════════
    #with tab1:#
    st.header(f"Office-wise Range of Accounts Opened – as on {report_date.strftime('%d.%m.%Y')}")
    st.info(
        "Upload the **Product Wise A/C Report** for each Division in the sidebar. "
        "The report counts only account category groups (MIS, PPFGP, SSA, RD, SBBAS, SBSGP, SCSS, TD). "
        "DC offices are excluded."
    )

    uploaded_divs = {d: f for d, f in div_files.items() if f is not None}
    if not uploaded_divs:
        st.warning("Please upload at least one Division file from the sidebar to generate this report.")
    else:
        division_dfs = {}
        for div, f in uploaded_divs.items():
            df = parse_product_report(f, div)
            if df is not None:
                division_dfs[div] = df

        if division_dfs:
            range_df = build_range_report(division_dfs)

            # ── Display Table ─────────────────────────────────────────
            st.subheader(f"Division-wise Summary — {len(division_dfs)} Division(s) loaded")

            # Styled display
            display_df = range_df.copy()
            total_row = {
                "Sl. No": "",
                "Division": "TOTAL",
                "Total Accounts Opened": display_df["Total Accounts Opened"].sum(),
            }
            for r in RANGES:
                total_row[r] = display_df[r].sum()
            display_df = pd.concat([display_df, pd.DataFrame([total_row])], ignore_index=True)

            st.dataframe(
                display_df.style
                    .set_properties(**{"text-align": "center"})
                    .apply(lambda x: ["background-color: #FFF2CC; font-weight: bold"
                                      if x.name == len(display_df) - 1 else ""
                                      for _ in x], axis=1),
                use_container_width=True,
                hide_index=True,
            )

            # ── Per-division detailed breakdown ───────────────────────
            with st.expander("🔍 Detailed Office-wise Breakdown per Division"):
                for div, df in division_dfs.items():
                    st.markdown(f"**{div} Division**")
                    df["Total Accounts"] = df[ACCOUNT_COLS].sum(axis=1)
                    df["Range"] = df["Total Accounts"].apply(classify_range)
                    show_cols = ["Name"] + ACCOUNT_COLS + ["Total Accounts", "Range"]
                    existing = [c for c in show_cols if c in df.columns]
                    st.dataframe(df[existing], use_container_width=True, hide_index=True)

            # ── Download ──────────────────────────────────────────────
            excel_bytes = export_range_report_excel(range_df, division_dfs)
            st.download_button(
                label="⬇️ Download Range Report as Excel",
                data=excel_bytes,
                file_name=f"Office_Range_Report_{report_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ════════════════════════════════════════════════════════════════════════
    # TAB 2 – Division-wise Summary Reports
    # ════════════════════════════════════════════════════════════════════════
    #with tab2:#
    st.header("Division-wise Summary Reports")

    st.info(
        "📋 **Tab 1** (Office-wise Range Report) uses the Division-wise Product files uploaded in the sidebar.  \n"
        "📋 **Tab 2** (Division-wise Summary Reports) uses the Accounts Opened Details and Net Addition files below."
    )
    # Parse whichever files were uploaded (all optional; show partial results)
    def _parse_or_none(f):
        if f is None: return None
        df = parse_summary_excel(f, ACCOUNT_COLS + CERT_COLS + ["A/c Opened","A/c Closed","Total"])
        return df

    ao_date_df  = _parse_or_none(ao_date_file)
    ao_cumul_df = _parse_or_none(ao_file)
    net_date_df = _parse_or_none(net_date_file)
    net_cumul_df= _parse_or_none(net_file)

    any_summary_file = any(f is not None for f in [ao_date_file, ao_file, net_date_file, net_file])
    if not any_summary_file:
        st.warning("Please upload at least one Summary file from the sidebar to generate this report.")
    else:
            # ── Table 1: Daily Summary ─────────────────────────────────
            st.subheader(
                f"POSB Accounts Daily Report dated {report_date.strftime('%d.%m.%Y')}"
            )

            summary_df = build_daily_summary(
                ao_date_df, ao_cumul_df, net_date_df, net_cumul_df,
                report_date, report_month, working_days_left
            )

            # ── Render daily summary as HTML table (full column width control) ──
            pct_col = "% achievement of proportionate Target"

            # Short header labels
            col_labels = {
                "Division":                                                  "Division",
                "Target FY 2026-27":                                         "Annual<br>Target",
                f"Proportionate Target upto {report_month}, {report_date.year}":
                                                                             f"Prop.<br>Target<br>{report_month[:3]} {report_date.year}",
                f"Daily Target upto {report_date.strftime('%d.%m.%Y')}":     f"Daily<br>Target<br>{report_date.strftime('%d.%m')}",
                f"No. of Accounts Opened on {report_date.strftime('%d.%m.%Y')}":
                                                                             f"A/cs<br>Opened<br>on {report_date.strftime('%d.%m')}",
                f"No. of Accounts Opened up to {report_date.strftime('%d.%m.%Y')}":
                                                                             f"A/cs<br>Opened<br>upto {report_date.strftime('%d.%m')}",
                f"Net no. of a/cs opened on {report_date.strftime('%d.%m.%Y')}":
                                                                             f"Net A/cs<br>on {report_date.strftime('%d.%m')}",
                f"Net no. of a/cs opened upto {report_date.strftime('%d.%m.%Y')}":
                                                                             f"Net A/cs<br>upto {report_date.strftime('%d.%m')}",
                "Shortfall on daily target":                                 "Shortfall<br>Daily",
                "Shortfall on proportionate target":                         "Shortfall<br>Prop.",
                "% achievement of proportionate Target":                     "% Prop.<br>Achiev.",
            }

            def _pct_style(val):
                try:
                    v = float(val)
                    if v < 50:   return "background:#FF0000;color:white;font-weight:700"
                    elif v < 75: return "background:#FFC000;font-weight:700"
                    elif v < 100:return "background:#FFFF00;font-weight:700"
                    else:        return "background:#70AD47;color:white;font-weight:700"
                except: return ""

            cols = list(summary_df.columns)
            pct_idx = cols.index(pct_col)

            # Column pixel widths
            col_widths = {"Division": 150}
            for c in cols:
                if c == "Division": continue
                col_widths[c] = 62 if "%" in c else 75

            # Build HTML
            hdr_style = ("background:#2E75B6;color:white;font-size:11px;font-weight:700;"
                         "text-align:center;vertical-align:bottom;padding:5px 3px;"
                         "white-space:normal;line-height:1.3;border:1px solid #ccc;")
            num_style  = "font-size:12px;text-align:right;padding:4px 6px;border:1px solid #e0e0e0;"
            div_style  = "font-size:12px;text-align:left;padding:4px 6px;border:1px solid #e0e0e0;font-weight:600;"
            tot_style  = "background:#1F3864;color:white;font-weight:700;font-size:12px;text-align:right;padding:4px 6px;border:1px solid #555;"
            tot_div_st = "background:#1F3864;color:white;font-weight:700;font-size:12px;text-align:left;padding:4px 6px;border:1px solid #555;"

            html = ["<div style='overflow-x:auto;'>",
                    "<table style='border-collapse:collapse;width:100%;table-layout:fixed;'>",
                    "<colgroup>"]
            for c in cols:
                html.append(f"<col style='width:{col_widths.get(c,75)}px;'>")
            html.append("</colgroup><thead><tr>")
            for c in cols:
                lbl = col_labels.get(c, c)
                html.append(f"<th style='{hdr_style}'>{lbl}</th>")
            html.append("</tr></thead><tbody>")

            for _, row in summary_df.iterrows():
                is_total = (row["Division"] == "Total HQ Region")
                html.append("<tr>")
                for ci, c in enumerate(cols):
                    val = row[c]
                    disp = f"{int(val):,}" if isinstance(val, (int, float)) and not isinstance(val, bool) else str(val)
                    if c == "Division":
                        html.append(f"<td style='{tot_div_st if is_total else div_style}'>{val}</td>")
                    elif c == pct_col:
                        cell_style = (tot_style if is_total else
                                      num_style + ";" + _pct_style(val))
                        html.append(f"<td style='{cell_style}'>{disp}</td>")
                    else:
                        html.append(f"<td style='{tot_style if is_total else num_style}'>{disp}</td>")
                html.append("</tr>")

            html.append("</tbody></table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)
            st.caption("🟢 ≥100%  🟡 75–99%  🟠 50–74%  🔴 <50% of proportionate target")

            # ── Table 2: Scheme-wise Status ───────────────────────────
            st.subheader(f"Scheme wise status – up to {report_date.strftime('%d.%m.%Y')}")

            scheme_df = build_scheme_wise(ao_cumul_df if ao_cumul_df is not None else pd.DataFrame())

            # Rename columns to full names for display
            scheme_display = scheme_df.rename(columns={k: k for k in ACCOUNT_COLS})

            def style_total_row(df):
                styles = pd.DataFrame("", index=df.index, columns=df.columns)
                styles.iloc[-1] = "background-color:#1F3864; color:white; font-weight:bold"
                return styles

            st.dataframe(
                scheme_display.style.apply(style_total_row, axis=None),
                use_container_width=True,
                hide_index=True,
            )

            # ── Legend / scheme names ──────────────────────────────────
            with st.expander("ℹ️ Scheme Code Reference"):
                legend = pd.DataFrame([
                    {"Code": k, "Full Name": v}
                    for k, v in SCHEME_FULL.items()
                ])
                st.dataframe(legend, hide_index=True, use_container_width=False)

            # ── Download both tables ───────────────────────────────────
            def build_summary_excel(summary_df, scheme_df, report_date, report_month):
                output = BytesIO()
                wb = xlsxwriter.Workbook(output, {"in_memory": True})

                title_fmt = wb.add_format({
                    "bold": True, "font_size": 12, "align": "center",
                    "valign": "vcenter", "bg_color": "#1F3864",
                    "font_color": "white", "border": 1,
                })
                hdr_fmt = wb.add_format({
                    "bold": True, "align": "center", "valign": "vcenter",
                    "bg_color": "#2E75B6", "font_color": "white",
                    "border": 1, "text_wrap": True,
                })
                data_fmt = wb.add_format({"align": "center", "border": 1})
                data_left = wb.add_format({"align": "left", "border": 1})
                total_fmt = wb.add_format({
                    "bold": True, "align": "center", "border": 1,
                    "bg_color": "#1F3864", "font_color": "white",
                })
                total_left = wb.add_format({
                    "bold": True, "align": "left", "border": 1,
                    "bg_color": "#1F3864", "font_color": "white",
                })
                green_fmt = wb.add_format({"align": "center", "border": 1, "bg_color": "#70AD47", "font_color": "white", "bold": True})
                yellow_fmt = wb.add_format({"align": "center", "border": 1, "bg_color": "#FFFF00", "bold": True})
                orange_fmt = wb.add_format({"align": "center", "border": 1, "bg_color": "#FFC000", "bold": True})
                red_fmt = wb.add_format({"align": "center", "border": 1, "bg_color": "#FF0000", "font_color": "white", "bold": True})

                # Sheet 1: Daily Summary
                ws1 = wb.add_worksheet("Daily Summary")
                cols = list(summary_df.columns)
                title = f"POSB Accounts Daily Report dated {report_date.strftime('%d.%m.%Y')}"
                ws1.merge_range(0, 0, 0, len(cols) - 1, title, title_fmt)
                ws1.set_row(0, 30)
                ws1.set_row(1, 45)
                for ci, col in enumerate(cols):
                    ws1.write(1, ci, col, hdr_fmt)
                ws1.set_column(0, 0, 24)
                ws1.set_column(1, len(cols) - 1, 14)
                pct_col_idx = cols.index("% achievement of proportionate Target")
                for ri, row in summary_df.iterrows():
                    is_total = row["Division"] == "Total HQ Region"
                    for ci, col in enumerate(cols):
                        val = row[col]
                        if is_total:
                            fmt = total_left if ci == 0 else total_fmt
                        elif ci == 0:
                            fmt = data_left
                        elif ci == pct_col_idx:
                            try:
                                v = float(val)
                                fmt = green_fmt if v >= 100 else (yellow_fmt if v >= 75 else (orange_fmt if v >= 50 else red_fmt))
                            except Exception:
                                fmt = data_fmt
                        else:
                            fmt = data_fmt
                        ws1.write(ri + 2, ci, val, fmt)

                # Sheet 2: Scheme wise
                ws2 = wb.add_worksheet("Scheme Wise Status")
                scheme_cols = list(scheme_df.columns)
                title2 = f"Scheme wise status – up to {report_date.strftime('%d.%m.%Y')}"
                ws2.merge_range(0, 0, 0, len(scheme_cols) - 1, title2, title_fmt)
                ws2.set_row(0, 28)
                ws2.set_row(1, 30)
                ws2.set_column(0, 0, 28)
                ws2.set_column(1, len(scheme_cols) - 1, 10)
                for ci, col in enumerate(scheme_cols):
                    ws2.write(1, ci, col, hdr_fmt)
                for ri, row in scheme_df.iterrows():
                    is_total = row["Division"] == "Total HQ Region"
                    for ci, col in enumerate(scheme_cols):
                        val = row[col]
                        if is_total:
                            fmt = total_left if ci == 0 else total_fmt
                        else:
                            fmt = data_left if ci == 0 else data_fmt
                        ws2.write(ri + 2, ci, val, fmt)

                wb.close()
                return output.getvalue()

            scheme_df_for_dl = scheme_df if 'scheme_df' in dir() else pd.DataFrame()
            excel_bytes2 = build_summary_excel(summary_df, scheme_df_for_dl, report_date, report_month)
            st.download_button(
                label="⬇️ Download Summary Reports as Excel",
                data=excel_bytes2,
                file_name=f"Division_Summary_Report_{report_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
