# app.py
# Streamlit app: upload one workbook with multiple sheets.
# Map "title" and "location" only.
# Detect VIP (SVP / EVP / Chief / C-suite).
# Tag North Carolina using fuzzy match.
# Export all original columns (including profileUrl, etc.) + Bucket column in one sheet.

import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="VIP Splitter — Multi-sheet Mapping (All Columns)", layout="wide")

# ------------------------------
# Helpers
# ------------------------------

NC_CITIES = {
    "charlotte", "raleigh", "durham", "greensboro", "winston-salem", "fayetteville",
    "cary", "wilmington", "asheville", "high point", "concord", "greenville",
    "gastonia", "chapel hill", "jacksonville", "huntersville", "rocky mount",
    "apex", "holly springs", "wake forest", "hickory", "kannapolis", "kernersville",
    "matthews", "monroe", "mooresville", "mint hill", "salisbury", "wilson",
    "thomasville", "shelby", "garner", "goldsboro", "boone", "burlington",
    "morrisville", "new bern"
}

def title_is_vip(title) -> bool:
    if not isinstance(title, str):
        return False
    t = title.lower()
    patterns = [
        r"\bsvp\b", r"\bsenior vice president\b",
        r"\bevp\b", r"\bexecutive vice president\b",
        r"\bchief\b", r"\bchief [a-z ]+ officer\b",
        r"\bcfo\b", r"\bcio\b", r"\bcto\b", r"\bcoo\b", r"\bciso\b", r"\bcmo\b",
        r"\bceo\b", r"\bcdo\b", r"\bcro\b", r"\bcco\b"
    ]
    return any(re.search(p, t, flags=re.IGNORECASE) for p in patterns)

def looks_like_nc_from_location(loc_value: str) -> bool:
    """Heuristic match for North Carolina based on free-form location."""
    if not isinstance(loc_value, str):
        return False
    s = loc_value.strip().lower()
    if not s:
        return False

    if "north carolina" in s:
        return True
    if s == "nc" or s.endswith(", nc") or s.endswith(" nc") or "(nc)" in s:
        return True
    for city in NC_CITIES:
        if city in s:
            return True
    tokens = [tok.strip(" ,()").lower() for tok in s.split(",") if tok.strip()]
    if any(tok == "nc" for tok in tokens):
        return True
    return False

def read_all_sheets(file) -> pd.DataFrame:
    """Read every sheet, merge into one DataFrame, preserve all columns."""
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=0, dtype=str)
        except Exception:
            continue
        if df.empty:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        df["_source_sheet"] = sheet
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    # unify columns across sheets
    all_cols = []
    for f in frames:
        all_cols.extend(list(f.columns))
    all_cols = list(dict.fromkeys(all_cols))
    frames = [f.reindex(columns=all_cols) for f in frames]
    merged = pd.concat(frames, ignore_index=True)
    merged["_source_workbook"] = getattr(file, "name", "uploaded.xlsx")
    return merged

def guess_mapping(colnames, target):
    patterns = {
        "title": ["title", "job title", "jobtitle", "position"],
        "location": ["location", "city", "state", "region", "province"]
    }
    for pat in patterns.get(target, [target]):
        for c in colnames:
            if pat in c.lower():
                return c
    return colnames[0] if colnames else None

def build_mapping_ui(df: pd.DataFrame):
    st.subheader("Map columns")
    if df.empty:
        st.warning("No data found in workbook.")
        return None, None

    # View headers
    with st.expander("View detected headers by sheet"):
        for sh in sorted(df["_source_sheet"].dropna().unique()):
            sh_cols = [c for c in df[df["_source_sheet"] == sh].columns if c not in ["_source_sheet", "_source_workbook"]]
            st.markdown(f"**{sh}**")
            st.write(sh_cols)

    candidates = [c for c in df.columns if c not in ["_source_sheet", "_source_workbook"]]
    guess_title = guess_mapping(candidates, "title")
    guess_location = guess_mapping(candidates, "location")

    c1, c2 = st.columns(2)
    with c1:
        title_col = st.selectbox("Select the job title column", candidates,
                                 index=(candidates.index(guess_title) if guess_title in candidates else 0),
                                 key="map_title")
    with c2:
        location_col = st.selectbox("Select the location column", candidates,
                                    index=(candidates.index(guess_location) if guess_location in candidates else 0),
                                    key="map_location")
    return title_col, location_col

def make_excel_one_sheet(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="VIP_All_WithBucket")
    bio.seek(0)
    return bio.getvalue()

# ------------------------------
# Streamlit App
# ------------------------------

st.title("VIP Splitter — Multi-sheet Mapping (All Columns + profileUrl)")
st.markdown("Upload a workbook with multiple sheets, map your job title and location columns, and export **all original columns (including profileUrl)** with a `Bucket` column tagging North Carolina vs Other.")

with st.sidebar:
    st.header("Upload workbook")
    up = st.file_uploader("Excel (.xlsx or .xls)", type=["xlsx", "xls"])

if not up:
    st.info("Upload a workbook to begin.")
    st.stop()

df_raw = read_all_sheets(up)
if df_raw.empty:
    st.error("No data found in workbook.")
    st.stop()

title_col, location_col = build_mapping_ui(df_raw)

go = st.button("Run filter", type="primary")
if not go:
    st.stop()

df = df_raw.copy()
df["__title__"] = df[title_col].astype(str)
df["__location__"] = df[location_col].astype(str)
df["__is_vip__"] = df["__title__"].apply(title_is_vip)
df["__is_nc__"] = df["__location__"].apply(looks_like_nc_from_location)

vip = df[df["__is_vip__"] == True].copy()
vip["Bucket"] = vip["__is_nc__"].map({True: "North Carolina", False: "Other"})

# Remove helper columns but keep all original headers
vip_out = vip.drop(columns=["__title__", "__location__", "__is_vip__", "__is_nc__"], errors="ignore")

# ------------------------------
# Display results
# ------------------------------

st.divider()
st.subheader("Overview")

vip_total = len(vip_out)
vip_nc = int((vip_out["Bucket"] == "North Carolina").sum())
vip_other = vip_total - vip_nc

c1, c2, c3 = st.columns(3)
c1.metric("VIP total", vip_total)
c2.metric("VIP in North Carolina", vip_nc)
c3.metric("VIP in Other locations", vip_other)

with st.expander("Preview — VIP in North Carolina"):
    st.dataframe(vip_out[vip_out["Bucket"] == "North Carolina"].head(50), use_container_width=True)

with st.expander("Preview — VIP in Other"):
    st.dataframe(vip_out[vip_out["Bucket"] == "Other"].head(50), use_container_width=True)

# ------------------------------
# Export
# ------------------------------

excel_bytes = make_excel_one_sheet(vip_out)
st.download_button(
    label="Download Excel (single sheet with all columns)",
    data=excel_bytes,
    file_name="VIP_All_WithBucket.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Includes all original headers (e.g., profileUrl), merged from all sheets, plus a Bucket column for North Carolina vs Other.")
