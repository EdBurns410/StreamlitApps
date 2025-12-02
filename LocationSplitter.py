# LocationSplitter.py
import io
import re
import zipfile
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------------
# Config
# ------------------------------------
SCHEMA: List[str] = [
    "profileUrl","fullName","firstName","lastName","companyName","title","companyId",
    "companyUrl","regularCompanyUrl","summary","titleDescription","industry","companyLocation",
    "location","durationInRole","durationInCompany","pastExperienceCompanyName",
    "pastExperienceCompanyUrl","pastExperienceCompanyTitle","pastExperienceDate",
    "pastExperienceDuration","connectionDegree","profileImageUrl","sharedConnectionsCount",
    "name","vmid","linkedInProfileUrl","isPremium","isOpenLink","query","timestamp",
    "defaultProfileUrl","searchAccountProfileId","searchAccountProfileName"
]

# Main buckets only. Everything else collapses to US.
MAIN_BUCKETS = {
    "Charlotte",
    "North Carolina (not inc Charlotte)",
    "Georgia",
    "Virginia",
    "Tennessee",
    "South Carolina",
    "New York",
    "Texas",
    "Arizona",
    "Illinois",
    "California",
    "Ohio",
    "Washington",
    "US"
}

# Titles that imply influence
TITLE_SENIORITY_PATTERNS = re.compile(
    r"\b(ceo|chief|cso|cto|cdo|cio|coo|cfo|president|founder|cofounder|co-founder|chair|head|vp|vice president|svp|evp|director|partner|principal)\b",
    re.I
)

# Detection helpers
US_GENERIC_REGEX = re.compile(r"^(us|usa|u\.s\.a\.|united states|u\.s\.)$", re.I)
CHARLOTTE_TOKENS = ["charlotte", "charlotte metro", "charlotte-concord-gastonia"]

STATE_NEEDLES = {
    "Georgia": ["georgia"," ga ",", ga","(ga)"],
    "Virginia": ["virginia"," va ",", va","(va)"],
    "Tennessee": ["tennessee"," tn ",", tn","(tn)"],
    "South Carolina": ["south carolina"," sc ",", sc","(sc)"],
    "New York": ["new york"," ny ",", ny","(ny)","nyc"],
    "Texas": ["texas"," tx ",", tx","(tx)"],
    "Arizona": ["arizona"," az ",", az","(az)"],
    "Illinois": ["illinois"," il ",", il","(il)"],
    "California": ["california"," ca ",", ca","(ca)"],
    "Ohio": ["ohio"," oh ",", oh","(oh)"],
    "Washington": ["washington"," wa ",", wa","(wa)"],
}

# ------------------------------------
# File reading
# ------------------------------------
def read_any(file) -> pd.DataFrame:
    name = file.name.lower()
    file.seek(0)

    if name.endswith(".xlsx"):
        try:
            return pd.read_excel(file)
        except Exception:
            file.seek(0)
            return pd.read_excel(file, engine="openpyxl")

    file.seek(0)
    file_bytes = file.getvalue() if hasattr(file, "getvalue") else file.read()

    for enc in ["utf-8", "utf-8-sig", "cp1252", "latin1"]:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc, engine="python", on_bad_lines="warn")
        except UnicodeDecodeError:
            continue
        except Exception:
            continue

    text = file_bytes.decode("cp1252", errors="replace")
    return pd.read_csv(io.StringIO(text), engine="python", on_bad_lines="warn")

def read_excel_sheet(file, sheet_name: str) -> pd.DataFrame:
    file.seek(0)
    try:
        return pd.read_excel(file, sheet_name=sheet_name)
    except Exception:
        file.seek(0)
        return pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")

# ------------------------------------
# Helpers
# ------------------------------------
def normalise_location(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s)
    return re.sub(r"\s+", " ", s.strip())

def ensure_schema(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = pd.DataFrame()
    for col in SCHEMA:
        src = mapping.get(col)
        if src and src in df.columns:
            out[col] = df[src]
        else:
            out[col] = pd.Series([np.nan] * len(df))
    return out

def compute_influencer_flag(row, shared_min: float) -> bool:
    title = str(row.get("title", "") or "")
    shared = row.get("sharedConnectionsCount", np.nan)
    seniority = bool(TITLE_SENIORITY_PATTERNS.search(title))
    try:
        shared = float(shared)
    except Exception:
        shared = np.nan
    high_shared = not np.isnan(shared) and shared >= float(shared_min)
    return seniority or high_shared

def pick_first_nonempty(*vals):
    for v in vals:
        if isinstance(v, str) and v.strip():
            return v
    return None

def dedupe_by_profile(df: pd.DataFrame) -> pd.DataFrame:
    key = []
    for _, r in df.iterrows():
        k = pick_first_nonempty(
            str(r.get("profileUrl") or ""),
            str(r.get("linkedInProfileUrl") or ""),
            str(r.get("defaultProfileUrl") or "")
        )
        key.append(k or "")
    df = df.copy()
    df["__dedupe_key"] = key
    df = df.drop_duplicates(subset="__dedupe_key").drop(columns="__dedupe_key")
    return df

def detect_main_bucket(location: str) -> str:
    l = normalise_location(location)
    low = l.lower()

    if US_GENERIC_REGEX.fullmatch(low.strip()):
        return "US"

    if any(tok in low for tok in CHARLOTTE_TOKENS):
        return "Charlotte"

    if ("north carolina" in low) or re.search(r"(^|[ ,;/\(\)\-])nc($|[ ,;/\(\)\-])", low):
        return "North Carolina (not inc Charlotte)"

    for state, needles in STATE_NEEDLES.items():
        if any(n in low for n in needles):
            return state

    # Any other place collapses to US by policy
    return "US"

# ------------------------------------
# UI
# ------------------------------------
st.set_page_config(page_title="Location Splitter for Dripify", layout="wide")
st.title("Location Splitter for Dripify")

st.markdown("Upload CSV or Excel, map columns, split into main geography buckets, score influencer status, and export a ZIP with CSVs and a summary.")

uploaded = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "csv"])

if uploaded:
    raw_df = read_any(uploaded)

    if uploaded.name.lower().endswith(".xlsx"):
        uploaded.seek(0)
        try:
            xls = pd.ExcelFile(uploaded)
            sheet_name = st.selectbox("Select sheet", xls.sheet_names, index=0)
            raw_df = read_excel_sheet(uploaded, sheet_name)
        except Exception:
            pass

    st.subheader("1. Column mapping")
    st.caption("Map your columns to the target schema. Unmapped fields will be blank.")
    cols = list(raw_df.columns)
    cols_with_skip = ["- Skip -"] + cols

    mapping: Dict[str, Optional[str]] = {}
    priority = [
        "profileUrl","linkedInProfileUrl","defaultProfileUrl","fullName","firstName","lastName",
        "title","companyName","location","sharedConnectionsCount"
    ]
    remaining = [c for c in SCHEMA if c not in priority]

    left, right = st.columns(2)
    with left:
        for col in priority:
            default_idx = cols_with_skip.index(col) if col in cols else 0
            sel = st.selectbox(f"Map → {col}", cols_with_skip, index=default_idx, key=f"map_{col}")
            mapping[col] = None if sel == "- Skip -" else sel
    with right:
        for col in remaining:
            default_idx = cols_with_skip.index(col) if col in cols else 0
            sel = st.selectbox(f"Map → {col}", cols_with_skip, index=default_idx, key=f"map_{col}")
            mapping[col] = None if sel == "- Skip -" else sel

    st.subheader("2. Options")
    infl_shared_min = st.number_input("Influencer shared connections minimum", min_value=0, max_value=1000, value=20, step=5)
    min_contacts_per_state_bucket = st.number_input("Minimum contacts to keep a state bucket", min_value=0, max_value=1000, value=10, step=1)

    st.subheader("3. Preview mapped data")
    df = ensure_schema(raw_df, mapping)
    df["location"] = df["location"].apply(normalise_location)
    df["sharedConnectionsCount"] = pd.to_numeric(df["sharedConnectionsCount"], errors="coerce")
    df["influencerFlag"] = df.apply(lambda r: compute_influencer_flag(r, infl_shared_min), axis=1)
    df = dedupe_by_profile(df)
    st.dataframe(df.head(20), use_container_width=True)

    # ------------------------------------
    # 4. Build groups
    # ------------------------------------
    st.subheader("4. Build groups")

    df["__group"] = df["location"].apply(detect_main_bucket)

    # Enforce main buckets only
    df.loc[~df["__group"].isin(MAIN_BUCKETS), "__group"] = "US"

    # Apply threshold: if a state bucket has less than N contacts, merge to US
    counts = df["__group"].value_counts()
    low_states = {g for g, n in counts.items() if g not in {"US"} and n < int(min_contacts_per_state_bucket)}
    if low_states:
        df.loc[df["__group"].isin(low_states), "__group"] = "US"

    # Summary
    summary = (
        df.assign(influencer=df["influencerFlag"].astype(int))
          .groupby("__group", dropna=False)
          .agg(total=("profileUrl","count"), influencers=("influencer","sum"))
          .sort_values("total", ascending=False)
          .reset_index()
          .rename(columns={"__group":"group"})
    )
    summary["influencer_rate"] = (summary["influencers"] / summary["total"]).round(3)

    st.markdown("Group summary")
    st.dataframe(summary, use_container_width=True)

    # ------------------------------------
    # 5. Download ZIP
    # ------------------------------------
    st.subheader("5. Download ZIP")
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for g, gdf in df.groupby("__group", dropna=False):
            group_name = g if isinstance(g, str) and g.strip() else "US"
            safe_name = re.sub(r"[^A-Za-z0-9\-_ ]+", "", group_name).strip().replace(" ", "_")
            out_df = gdf[SCHEMA + ["influencerFlag"]].copy()
            zf.writestr(f"{safe_name}.csv", out_df.to_csv(index=False))
        zf.writestr("summary.csv", summary.to_csv(index=False))

    st.download_button(
        "Download ZIP of CSVs and summary",
        data=buffer.getvalue(),
        file_name="dripify_location_split.zip",
        mime="application/zip",
        use_container_width=True
    )

else:
    st.caption("Upload a file to begin.")

# Run:
# pip install streamlit pandas openpyxl
# streamlit run LocationSplitter.py
