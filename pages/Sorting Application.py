import streamlit as st
import pandas as pd
import re
import glob as _glob
from rapidfuzz import fuzz
from streamlit_keyup import st_keyup
from st_keyup import st_keyup

# ─────────────────────────────────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sorting Assistance",
    page_icon="📮",
    layout="wide"
)

# ─────────────────────────────────────────────────────────────────────
# Shared Navigation
# ─────────────────────────────────────────────────────────────────────
def _render_nav():
    st.sidebar.markdown("""<div style='padding:8px 0 4px 0;'>
        <p style='font-size:12px;font-weight:700;color:#888;
           text-transform:uppercase;letter-spacing:1px;margin:0 0 4px 0;'>Pages</p>
        </div>""", unsafe_allow_html=True)

    st.sidebar.page_link("Analytics_Excel.py", label="\U0001f3e0 Home")

    for pat, lbl in [
        ("pages/AEBAS_Monitoring.py|pages/*[Aa][Ee][Bb][Aa][Ss]*.py", "🤚 AEBAS Monitoring"),
        ("pages/Bulk_Analytics.py|pages/*[Bb]ulk*.py", "📊 Bulk Customer Analytics"),
        ("pages/Delivery_Productivity.py|pages/*[Dd]elivery*.py", "📦 Delivery Productivity"),
        ("pages/1_Digital_Transactions.py|pages/*[Dd]igital*.py", "💻 Digital Transactions"),
        ("pages/POSB Daily Report.py|pages/*[Pp][Oo][Ss][Bb]*.py", "📮 POSB Daily Report"),
        ("pages/Sorting Assistance.py|pages/*[Ss]orting*.py", "📮 Sorting Assistance"),
    ]:
        hits = []
        for p in pat.split("|"):
            hits += _glob.glob(p)

        if hits:
            st.sidebar.page_link(hits[0].replace("\\", "/"), label=lbl)

    st.sidebar.markdown("<hr style='margin:8px 0 12px 0;'>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────
# Auth Guard
# ─────────────────────────────────────────────────────────────────────
if not st.session_state.get("authenticated", False):
    st.warning("⚠️ You are not logged in.")
    st.markdown("Please go to **🔐 Login** in the sidebar.")
    with st.sidebar:
        _render_nav()
    st.stop()

with st.sidebar:
    _render_nav()


# ─────────────────────────────────────────────────────────────────────
# Load Master
# ─────────────────────────────────────────────────────────────────────
MASTER_PATH = "data/Sorting Application/Sorting List HQR 2025.xlsx"


@st.cache_data
def load_master():
    df = pd.read_excel(MASTER_PATH, sheet_name="Sorting Lists")
    df.columns = [str(c).strip() for c in df.columns]

    required = [
        "Name of the Post Office",
        "Pincode",
        "Beat",
        "Colony/Area",
        "Door Nos."
    ]

    df = df[required].copy()

    for c in required:
        df[c] = df[c].fillna("").astype(str)

    df["Pincode"] = df["Pincode"].str.replace(".0", "", regex=False)

    return df


df = load_master()


# ─────────────────────────────────────────────────────────────────────
# Normalization
# ─────────────────────────────────────────────────────────────────────
def normalize_text(txt):
    txt = str(txt).upper().strip()

    txt = txt.replace("/", "-")
    txt = txt.replace("\\", "-")
    txt = txt.replace(".", "-")
    txt = txt.replace(",", " ")

    txt = re.sub(r"\s+", " ", txt)
    txt = txt.strip()

    return txt


def normalize_door(txt):
    txt = normalize_text(txt)
    txt = txt.replace(" ", "")
    txt = re.sub(r"-+", "-", txt)
    return txt


# Precompute normalized columns
df["area_norm"] = df["Colony/Area"].apply(normalize_text)
df["door_norm"] = df["Door Nos."].apply(normalize_door)
df["po_norm"] = df["Name of the Post Office"].apply(normalize_text)


# ─────────────────────────────────────────────────────────────────────
# Search Type Detection
# ─────────────────────────────────────────────────────────────────────
def detect_query_type(q):
    q = q.strip()

    if re.fullmatch(r"\d{6}", q):
        return "pincode"

    digits = sum(ch.isdigit() for ch in q)
    nonspace = len(q.replace(" ", ""))

    if nonspace > 0 and digits / nonspace > 0.55:
        return "door"

    return "area"


# ─────────────────────────────────────────────────────────────────────
# Scoring
# ─────────────────────────────────────────────────────────────────────
def weighted_score(query, candidate):
    s1 = fuzz.partial_ratio(query, candidate)
    s2 = fuzz.token_sort_ratio(query, candidate)
    s3 = fuzz.token_set_ratio(query, candidate)

    score = 0.45*s1 + 0.35*s2 + 0.20*s3

    if query in candidate:
        score += 15

    if candidate.startswith(query):
        score += 10

    return score


# ─────────────────────────────────────────────────────────────────────
# Door Range Logic
# ─────────────────────────────────────────────────────────────────────
def extract_numbers(text):
    return [int(x) for x in re.findall(r'\d+', text)]


def door_match(user_input, door_text):
    user = normalize_door(user_input)
    door = normalize_door(door_text)

    if user == door:
        return 100

    fuzzy = fuzz.partial_ratio(user, door)
    if fuzzy >= 90:
        return fuzzy

    if "TO" in door:
        try:
            parts = door.split("TO")
            start = parts[0].strip()
            end = parts[1].strip()

            u_nums = extract_numbers(user)
            s_nums = extract_numbers(start)
            e_nums = extract_numbers(end)

            if u_nums and s_nums and e_nums:
                u = u_nums[-1]
                s = s_nums[-1]
                e = e_nums[-1]

                if s <= u <= e:
                    return 98
        except:
            pass

    return 0


# ─────────────────────────────────────────────────────────────────────
# Smart Search
# ─────────────────────────────────────────────────────────────────────
def smart_search(query):
    qtype = detect_query_type(query)

    temp = df.copy()

    if qtype == "pincode":
        result = temp[temp["Pincode"] == query].copy()
        result["Score"] = 100
        return result

    elif qtype == "door":
        scores = temp["Door Nos."].apply(lambda x: door_match(query, x))
        temp["Score"] = scores
        result = temp[temp["Score"] >= 75]
        return result.sort_values("Score", ascending=False).head(10)

    else:
        query_norm = normalize_text(query)

        scores = []
        for _, row in temp.iterrows():
            area_score = weighted_score(query_norm, row["area_norm"])
            po_score = weighted_score(query_norm, row["po_norm"]) * 0.5

            final = max(area_score, po_score)
            scores.append(final)

        temp["Score"] = scores
        result = temp[temp["Score"] >= 75]

        return result.sort_values("Score", ascending=False).head(10)


# ─────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────
st.title("📮 Sorting Assistance")
st.caption("Type Pincode, Colony, Area or Door Number")

query = st_keyup(
    "Search",
    placeholder="Type Pincode / Area / Door Number...",
    debounce=250
)
st.write("Query =", repr(query))
if query and str(query).strip():
    results = smart_search(query)

    if results.empty:
        st.warning("No matching records found.")
    else:
        st.success(f"Top {len(results)} matches")

        display = results[
            [
                "Score",
                "Name of the Post Office",
                "Pincode",
                "Beat",
                "Colony/Area",
                "Door Nos."
            ]
        ].copy()

        display["Score"] = display["Score"].round(0).astype(int)

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True
        )
