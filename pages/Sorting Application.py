import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Sorting Assistance",
    page_icon="📮",
    layout="wide"
)

MASTER_PATH = "data/Sorting Application/Sorting List HQR 2025.xlsx"


# =========================
# LOAD DATA
# =========================
@st.cache_data
def load_data():
    df = pd.read_excel(MASTER_PATH, sheet_name="Sorting Lists")
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = [
        "Name of the Post Office",
        "Pincode",
        "Beat",
        "Colony/Area",
        "Door Nos."
    ]

    df = df[required_cols].copy()

    for col in required_cols:
        df[col] = df[col].fillna("").astype(str)

    df["Pincode"] = df["Pincode"].str.replace(".0", "", regex=False)

    return df


df = load_data()


# =========================
# NORMALIZATION HELPERS
# =========================
def normalize_text(txt):
    if pd.isna(txt):
        return ""

    txt = str(txt).upper().strip()

    txt = txt.replace("/", "-")
    txt = txt.replace(".", "-")
    txt = txt.replace("\\", "-")
    txt = txt.replace(" TO ", " TO ")

    txt = re.sub(r"\s+", "", txt)
    txt = re.sub(r"-+", "-", txt)

    return txt


def fuzzy_match(a, b):
    return fuzz.partial_ratio(a, b)


# =========================
# DOOR RANGE PARSER
# =========================
def extract_numbers(text):
    return [int(x) for x in re.findall(r'\d+', text)]


def door_range_match(user_input, door_range):
    user_norm = normalize_text(user_input)
    range_norm = normalize_text(door_range)

    # direct match
    if user_norm == range_norm:
        return True

    # fuzzy
    if fuzzy_match(user_norm, range_norm) >= 90:
        return True

    if "TO" not in range_norm:
        return False

    try:
        start, end = [x.strip() for x in range_norm.split("TO")]

        user_nums = extract_numbers(user_norm)
        start_nums = extract_numbers(start)
        end_nums = extract_numbers(end)

        if not user_nums or not start_nums or not end_nums:
            return False

        user_last = user_nums[-1]
        start_last = start_nums[-1]
        end_last = end_nums[-1]

        if start_last <= user_last <= end_last:
            return True

    except:
        pass

    return False


# =========================
# SEARCH FUNCTIONS
# =========================
def search_pincode(pin):
    return df[df["Pincode"] == str(pin)]


def search_area(query):
    q = normalize_text(query)

    scores = []
    for _, row in df.iterrows():
        area = normalize_text(row["Colony/Area"])
        score = fuzzy_match(q, area)

        if q in area:
            score += 20

        scores.append(score)

    temp = df.copy()
    temp["score"] = scores

    return temp[temp["score"] >= 65].sort_values("score", ascending=False)


def search_door(query):
    matches = []

    q_norm = normalize_text(query)

    for _, row in df.iterrows():
        door_text = row["Door Nos."]
        door_norm = normalize_text(door_text)

        matched = False

        # Exact
        if q_norm == door_norm:
            matched = True

        # Fuzzy
        elif fuzzy_match(q_norm, door_norm) >= 85:
            matched = True

        # Range Logic
        elif door_range_match(query, door_text):
            matched = True

        if matched:
            matches.append(row)

    if matches:
        return pd.DataFrame(matches)
    else:
        return pd.DataFrame()


# =========================
# UI
# =========================
st.title("📮 Sorting Assistance")
st.caption("Search office and beat using Pincode, Area or Door Number")

mode = st.radio(
    "Search By",
    ["Pincode", "Area / Colony", "Door Number"],
    horizontal=True
)

query = st.text_input("Enter Search Value")

if st.button("Search"):
    if not query.strip():
        st.warning("Please enter a value.")
        st.stop()

    result = pd.DataFrame()

    if mode == "Pincode":
        result = search_pincode(query)

    elif mode == "Area / Colony":
        result = search_area(query)

    elif mode == "Door Number":
        result = search_door(query)

    if result.empty:
        st.error("No matching records found.")
    else:
        st.success(f"{len(result)} matching entries found")

        display_cols = [
            "Name of the Post Office",
            "Pincode",
            "Beat",
            "Colony/Area",
            "Door Nos."
        ]

        st.dataframe(
            result[display_cols],
            use_container_width=True,
            hide_index=True
        )

        st.subheader("Beat Summary")

        summary = (
            result.groupby(
                ["Name of the Post Office", "Beat"]
            )
            .size()
            .reset_index(name="Matches")
        )

        st.dataframe(summary, use_container_width=True, hide_index=True)
