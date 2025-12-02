# app.py
# Single-upload VIP splitter:
# - Upload one file with headers:
#   email, firstname, lastname, jobtitle, Company Name, country, State, Organisation Type, Segment, Seniority
# - Identify VIP roles (SVP/EVP/Chief) from 'jobtitle'
# - Fuzzy-match 'State' to North Carolina (accepts 'NC', 'North Carolina', 'Charlotte, North Carolina, United States', city names, etc.)
# - Split VIPs into NC vs Other
# - Show counts + previews
# - Download one Excel file with:
#     1) VIP_All_WithBucket
#     2) VIP_NorthCarolina
#     3) VIP_Other

import io
import re
from typing import Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="VIP Splitter (SVP / EVP / Chief)", layout="wide")

REQUIRED_COLS = [
    "email", "firstname", "lastname", "jobtitle", "Company Name",
    "country", "State", "Organisation Type", "Segment", "Seniority"
]

US_STATE_ABBR_TO_NAME = {
    "NC": "North Carolina"
}

# A practical list of NC cities for fuzzy matching when 'State' contains city-level text
NC_CITIES = {
    "charlotte", "raleigh", "durham", "greensboro", "winston-salem", "fayetteville",
    "cary", "wilmington", "asheville", "high point", "concord", "greenville",
    "gastonia", "chapel hill", "jacksonville", "huntersville", "rocky mount",
    "apex", "holly springs", "wake forest", "hickory", "kannapolis", "kernersville",
    "matthews", "monroe", "mooresville", "mint hill", "salisbury", "wilson",
    "thomasville", "shelby", "garner", "goldsboro", "boone", "burlington",
    "morrisville", "new bern"
}

def load_table(file) -> pd.DataFrame:
    name = getattr(file, "name", "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file, dtype=str)
    else:
        df = pd.read_excel(file, dtype=str)
    # Normalise column names for matching, but keep originals
    df.columns = [str(c).strip() for c in df.columns]
    return df

def ensure_required_columns(df: pd.DataFrame) -> Tuple[bool, list]:
    present = set(c.lower() for c in df.columns)
    missing = [c for c in REQUIRED_COLS if c.lower() not in present]
    return (len(missing) == 0, missing)

def col(df: pd.DataFrame, name: str) -> str:
    """Return the actual column name in df matching 'name' case-insensitively."""
    for c in df.columns:
        if c.lower() == name.lower():
            return c
    raise KeyError(name)

def title_is_vip(title) -> bool:
    if not isinstance(title, str):
        return False
    t = title.lower()
    patterns = [
        r"\bsvp\b", r"\bsenior vice president\b",
        r"\bevp\b", r"\bexecutive vice president\b",
        r"\bchief\b",
        r"\bchief [a-z ]+ officer\b",
        r"\bcfo\b", r"\bcio\b", r"\bcto\b", r"\bcoo\b", r"\bciso\b", r"\bcmo\b", r"\bceo\b", r"\bcdo\b", r"\bcro\b", r"\bcco\b"
    ]
    return any(re.search(p, t, flags=re.IGNORECASE) for p in patterns)

def looks_like_nc(state_value: str) -> bool:
    """Heuristic match for North Carolina in free-form 'State' strings."""
    if not isinstance(state_value, str):
        return False
    s = state_value.strip().lower()

    # Quick exits
    if not s:
        return False

    # Direct name checks
    if "north carolina" in s:
        return True

    # Abbreviation checks for NC
    # Examples: "NC", "Charlotte, NC", "Raleigh NC", "(NC)"
    if s == "nc" or s.endswith(", nc") or s.endswith(" nc") or "(nc)" in s:
        return True

    # City-based check for common NC cities present in the string
    for city in NC_CITIES:
        if city in s:
            return True

    # Sometimes the field is comma-separated place strings: try to catch 'United States' noise
    # and a trailing 'nc' token with punctuation
    tokens = [tok.strip(" ,()") for tok in s.split(",") if tok.strip()]
    if any(tok.lower() == "nc" for tok in tokens):
        return True

    return False

st.title("VIP Splitter — SVP / EVP / Chief, North Carolina focus")

with st.sidebar:
    st.header("Upload your file")
    up = st.file_uploader("Excel or CSV with the required headers", type=["xlsx", "xls", "csv"])

st.markdown(
    "Required headers (case-insensitive): "
    "`email`, `firstname`, `lastname`, `jobtitle`, `Company Name`, `country`, `State`, "
    "`Organisation Type`, `Segment`, `Seniority`."
)

if not up:
    st.info("Upload a file to begin.")
    st.stop()

df_raw = load_table(up)
ok, missing = ensure_required_columns(df_raw)
if not ok:
    st.error(f"Missing required columns: {', '.join(missing)}")
    st.stop()

# Canonical views
job_col = col(df_raw, "jobtitle")
state_col = col(df_raw, "State")

df = df_raw.copy()
df["__is_vip__"] = df[job_col].apply(title_is_vip)
df["__is_nc__"] = df[state_col].apply(looks_like_nc)

vip = df[df["__is_vip__"] == True].copy()
vip_nc = vip[vip["__is_nc__"] == True].copy()
vip_other = vip[vip["__is_nc__"] == False].copy()

# For output, add a bucket column and remove helper cols
vip_nc_out = vip_nc.copy()
vip_nc_out["Bucket"] = "North Carolina"
vip_nc_out.drop(columns=["__is_vip__", "__is_nc__"], inplace=True, errors="ignore")

vip_other_out = vip_other.copy()
vip_other_out["Bucket"] = "Other"
vip_other_out.drop(columns=["__is_vip__", "__is_nc__"], inplace=True, errors="ignore")

vip_all = pd.concat([vip_nc_out, vip_other_out], ignore_index=True)

# ---------------- UI: counts + preview ----------------
st.divider()
st.subheader("Overview")
c1, c2, c3 = st.columns(3)
c1.metric("VIP total", len(vip_all))
c2.metric("VIP in North Carolina", len(vip_nc_out))
c3.metric("VIP in Other locations", len(vip_other_out))

with st.expander("Preview — VIP in North Carolina"):
    st.dataframe(vip_nc_out.head(50), use_container_width=True)

with st.expander("Preview — VIP in Other"):
    st.dataframe(vip_other_out.head(50), use_container_width=True)

# ---------------- Export ----------------
def make_excel(vip_all_df, vip_nc_df, vip_other_df) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        # Combined first (single sheet with a Bucket column)
        vip_all_df.to_excel(writer, index=False, sheet_name="VIP_All_WithBucket")
        vip_nc_df.to_excel(writer, index=False, sheet_name="VIP_NorthCarolina")
        vip_other_df.to_excel(writer, index=False, sheet_name="VIP_Other")
    bio.seek(0)
    return bio.getvalue()

excel_bytes = make_excel(vip_all, vip_nc_out, vip_other_out)
st.download_button(
    label="Download Excel (1 file, 3 tabs)",
    data=excel_bytes,
    file_name="VIP_Splits_NorthCarolina.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("VIP roles identified via job titles; North Carolina matched via state text, abbreviations, and common city names.")
