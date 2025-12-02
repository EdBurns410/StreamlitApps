# app.py
# Streamlit contact selector for MoneyNext
# - Upload Excel/CSV
# - Map columns to fixed schema
# - Filter to "Relevant" + tech/AI/transformation or senior
# - Score, balance seniority, and select best 150-200 per company
# - One-sheet download with full original fields
#
# Requirements:
#   pip install streamlit pandas python-levenshtein openpyxl xlsxwriter

from io import BytesIO
from typing import Dict, List, Optional, Tuple
import re

import pandas as pd
import numpy as np
import streamlit as st

try:
    from Levenshtein import ratio as lev_ratio
except Exception:
    def lev_ratio(a, b):
        a = "" if a is None else str(a).lower()
        b = "" if b is None else str(b).lower()
        return 1.0 if a == b else 0.0

st.set_page_config(page_title="Contact Selector – Event Targeting", layout="wide")

# ------------------------------- Target schema -------------------------------
TARGET_HEADERS: List[str] = [
    "SalesNavUrl","fullName","firstName","lastName","companyName","title",
    "companyId","companyUrl","regularCompanyUrl","summary","titleDescription",
    "industry","companyLocation","location","durationInRole","durationInCompany",
    "connectionDegree","profileImageUrl","sharedConnectionsCount","name","vmid",
    "linkedInProfileUrl","isPremium","isOpenLink","query","timestamp","ProfileUrl",
    "searchAccountProfileId","searchAccountProfileName","Bank Size","Seniority",
    "Focus Area","Relevance"
]

# Seniority weights
SENIORITY_ORDER = [
    "C-Suite","Chair/Board Member","Vice President","Leadership","Director",
    "Head","Senior","Manager","Consultant","Non-Senior Banker","Entry-Level"
]
SENIORITY_WEIGHT = {
    "C-Suite": 10,
    "Chair/Board Member": 10,
    "Vice President": 8,
    "Leadership": 8,
    "Director": 7,
    "Head": 7,
    "Senior": 5,
    "Manager": 5,
    "Consultant": 3,
    "Non-Senior Banker": 3,
    "Entry-Level": 1,
}

# Title keyword scoring
TITLE_KEYWORDS = {
    # signal: weight
    "ai": 6, "artificial intelligence": 6, "genai": 6, "machine learning": 5, "ml": 4, "data": 4,
    "chief": 6, "cdo": 6, "cio": 6, "cto": 6, "ciso": 6, "coo": 5, "cfo": 5,
    "vp": 5, "vice president": 5, "director": 4, "head": 4, "lead": 3, "manager": 2,
    "transformation": 6, "digital": 4, "innovation": 5, "strategy": 4, "architecture": 5,
    "architect": 5, "platform": 3, "engineering": 3, "engineer": 3, "cloud": 4,
    "security": 4, "cyber": 4, "payments": 3, "core banking": 5, "modernisation": 4,
    "modernization": 4, "automation": 4, "risk": 3, "compliance": 3, "fraud": 3,
    "analytics": 4, "business intelligence": 4, "devops": 3, "product": 3,
}
TITLE_NEGATIVE = {
    "recruiter": -4, "talent": -3, "hr": -2, "student": -5, "intern": -6,
    "retired": -10, "consultant at": -2  # generic agency spam titles
}

# Focus area default filters
DEFAULT_FOCUS_KEYWORDS = ["ai","genai","machine learning","ml","data","digital","transformation","innovation","architecture","platform","cloud","security","cyber","payments","automation","core banking","modernisation","modernization","bi","analytics","devops","product"]

# Senior buckets for balancing per company
BUCKET_RULES = {
    "leadership": {"labels": {"C-Suite","Chair/Board Member","Vice President","Leadership","Director","Head"}},
    "manager": {"labels": {"Senior","Manager"}},
    "specialist": {"labels": {"Consultant","Non-Senior Banker","Entry-Level"}},
}
BUCKET_SPLIT = {"leadership": 0.35, "manager": 0.40, "specialist": 0.25}

# ------------------------------- Helpers -------------------------------
def best_guess_mapping(src_cols: List[str], target: List[str]) -> Dict[str, Optional[str]]:
    mapping = {}
    for t in target:
        best = None
        best_score = 0.0
        for c in src_cols:
            score = lev_ratio(str(c).lower(), t.lower())
            # allow loose matches on common aliases
            alias_bonus = 0
            if t in ["linkedInProfileUrl","ProfileUrl"] and re.search("profile", str(c), re.I):
                alias_bonus = 0.05
            if t in ["companyName"] and re.search("company", str(c), re.I):
                alias_bonus = 0.05
            if t in ["firstName","lastName","fullName","title","location","industry"]:
                if re.search(t.replace("Name",""), str(c), re.I):
                    alias_bonus += 0.05
            score += alias_bonus
            if score > best_score:
                best_score = score
                best = c
        mapping[t] = best if best_score >= 0.45 else None
    return mapping

def normalise_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def infer_seniority_from_title(title: str) -> Optional[str]:
    if not isinstance(title, str) or not title:
        return None
    t = title.lower()
    checks = [
        ("c-suite", [" chief ", " cto", " cio", " cdo", " ciso", " cfo", "ceo", "coo"]),
        ("Chair/Board Member", [" chair", " board"]),
        ("Vice President", [" vp", "vice president"]),
        ("Leadership", ["executive", "leadership"]),
        ("Director", [" director "]),
        ("Head", [" head "]),
        ("Senior", [" senior ", " sr ", "principal"]),
        ("Manager", [" manager "]),
        ("Consultant", [" consultant ", " advisory ", " advisor "]),
        ("Non-Senior Banker", [" associate ", " analyst ", "officer"]),
        ("Entry-Level", [" intern ", "graduate", "trainee"]),
    ]
    for label, keys in checks:
        for k in keys:
            if k in f" {t} ":
                return label if label != "c-suite" else "C-Suite"
    return None

def title_score(title: str) -> int:
    if not isinstance(title, str):
        return 0
    t = title.lower()
    score = 0
    for k, w in TITLE_KEYWORDS.items():
        if k in t:
            score += w
    for k, w in TITLE_NEGATIVE.items():
        if k in t:
            score += w
    return score

def focus_hits(text: str, keys: List[str]) -> int:
    if not isinstance(text, str):
        return 0
    t = text.lower()
    return sum(1 for k in keys if k in t)

def seniority_weight(s: Optional[str]) -> int:
    if not s:
        return 0
    return SENIORITY_WEIGHT.get(s, 0)

def derive_bucket(s: Optional[str]) -> str:
    if s in BUCKET_RULES["leadership"]["labels"]:
        return "leadership"
    if s in BUCKET_RULES["manager"]["labels"]:
        return "manager"
    return "specialist"

def safe_lower(x):
    return str(x).lower() if isinstance(x, str) else x

def to_int_safe(x):
    try:
        return int(x)
    except Exception:
        return None

# ------------------------------- UI -------------------------------
st.title("Contact Selector for Event Targeting")

uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx","xls","csv"])
if not uploaded:
    st.stop()

# Read file
input_is_excel = uploaded.name.lower().endswith((".xlsx",".xls"))
if input_is_excel:
    xls = pd.ExcelFile(uploaded)
    sheet = st.selectbox("Select sheet", xls.sheet_names, index=0)
    df_raw = pd.read_excel(xls, sheet_name=sheet, dtype=str)
else:
    df_raw = pd.read_csv(uploaded, dtype=str)

df_raw = normalise_cols(df_raw)
st.caption(f"Rows: {len(df_raw):,} • Columns: {len(df_raw.columns)}")

# Column mapping
st.subheader("Map columns to target schema")
auto_map = best_guess_mapping(df_raw.columns.tolist(), TARGET_HEADERS)

with st.expander("Column mapping", expanded=True):
    mapping = {}
    cols_with_none = ["<None>"] + df_raw.columns.tolist()
    col1, col2, col3 = st.columns(3)
    for i, t in enumerate(TARGET_HEADERS):
        if i % 3 == 0:
            with col1:
                sel = st.selectbox(t, cols_with_none, index=(df_raw.columns.tolist().index(auto_map[t]) + 1) if auto_map[t] in df_raw.columns else 0, key=f"map_{t}")
        elif i % 3 == 1:
            with col2:
                sel = st.selectbox(t, cols_with_none, index=(df_raw.columns.tolist().index(auto_map[t]) + 1) if auto_map[t] in df_raw.columns else 0, key=f"map_{t}")
        else:
            with col3:
                sel = st.selectbox(t, cols_with_none, index=(df_raw.columns.tolist().index(auto_map[t]) + 1) if auto_map[t] in df_raw.columns else 0, key=f"map_{t}")
        mapping[t] = None if sel == "<None>" else sel

required_min = ["companyName","title","Relevance"]  # minimal for selection
missing_required = [t for t in required_min if mapping.get(t) is None]
if missing_required:
    st.error(f"Map required fields: {', '.join(missing_required)}")
    st.stop()

# Build standardised frame (but preserve all original columns for output)
std = pd.DataFrame()
for t in TARGET_HEADERS:
    src = mapping.get(t)
    std[t] = df_raw[src] if src in df_raw.columns else pd.NA

# Enrichment columns
std["_norm_company"] = std["companyName"].fillna("").str.strip()
std["_norm_title"] = std["title"].fillna("").str.strip()

# Filter controls
st.subheader("Filter settings")
c1, c2, c3, c4 = st.columns(4)
with c1:
    max_per_company = st.number_input("Max per company", min_value=50, max_value=500, value=200, step=10)
with c2:
    min_per_company = st.number_input("Min per company if available", min_value=0, max_value=500, value=150, step=10)
with c3:
    require_relevance = st.checkbox('Require Relevance == "Relevant"', value=True)
with c4:
    allow_seniority_only = st.checkbox("Allow seniority pass-through if Focus Area not matched", value=True)

focus_default = ", ".join(DEFAULT_FOCUS_KEYWORDS)
focus_text = st.text_input("Focus Area keywords (comma separated)", value=focus_default)
focus_keys = [k.strip().lower() for k in focus_text.split(",") if k.strip()]

# Compute features
work = std.copy()

# Seniority final (prefer provided, else infer from title)
work["_seniority_final"] = work["Seniority"].where(work["Seniority"].notna(), work["title"].map(infer_seniority_from_title))

# Title and focus scores
work["_title_score"] = work["title"].map(title_score)
work["_focus_hits"] = work["Focus Area"].map(lambda x: focus_hits(x, focus_keys)) + work["title"].map(lambda x: focus_hits(x, focus_keys))
work["_seniority_w"] = work["_seniority_final"].map(seniority_weight)
# Connection degree bonus if present
conn_deg = work.get("connectionDegree")
if conn_deg is not None:
    work["_conn_bonus"] = work["connectionDegree"].map(lambda x: 1 if str(x).strip() in {"1","1st","First","first"} else 0)
else:
    work["_conn_bonus"] = 0

# Base eligibility
elig_relevance = True
if require_relevance:
    elig_relevance = work["Relevance"].fillna("").str.lower().eq("relevant")

elig_focus = (work["_focus_hits"] > 0)
elig_senior = work["_seniority_final"].isin(list(BUCKET_RULES["leadership"]["labels"]) + list(BUCKET_RULES["manager"]["labels"]))

elig = elig_relevance & (elig_focus | (allow_seniority_only & elig_senior))
work = work[elig].copy()

# Composite score
# Weighting: seniority 0.5, title 0.35, focus 0.12, connection 0.03
work["_score"] = (
    work["_seniority_w"].fillna(0)*0.5 +
    work["_title_score"].fillna(0)*0.35 +
    work["_focus_hits"].fillna(0)*0.12 +
    work["_conn_bonus"].fillna(0)*0.03
)

# Bucket for balancing
work["_bucket"] = work["_seniority_final"].map(derive_bucket)

# Selection per company with bucket quotas
def pick_for_company(g: pd.DataFrame, n_max: int, n_min: int) -> pd.DataFrame:
    g = g.sort_values("_score", ascending=False).copy()
    if len(g) <= n_min:
        return g
    # target counts by bucket
    targets = {b: int(round(n_max * p)) for b, p in BUCKET_SPLIT.items()}
    # ensure sum equals n_max
    diff = n_max - sum(targets.values())
    if diff != 0:
        # adjust leadership bucket first
        targets["leadership"] += diff

    picked = []
    remaining = g.copy()
    for b in ["leadership","manager","specialist"]:
        gb = remaining[remaining["_bucket"] == b].sort_values("_score", ascending=False)
        take = min(targets[b], len(gb))
        picked.append(gb.head(take))
        remaining = remaining.drop(gb.head(take).index)

    combined = pd.concat(picked) if picked else pd.DataFrame(columns=g.columns)
    # fill up to n_max with best remaining
    if len(combined) < n_max and len(remaining) > 0:
        need = n_max - len(combined)
        combined = pd.concat([combined, remaining.head(need)])

    # Guarantee at least n_min by backfilling if needed
    if len(combined) < n_min:
        need = min(n_min - len(combined), len(g) - len(combined))
        if need > 0:
            rest = g.drop(combined.index)
            combined = pd.concat([combined, rest.head(need)])

    return combined.head(min(n_max, len(g)))

if work.empty:
    st.warning("No rows matched the filters. Loosen filters or upload a different file.")
    st.stop()

grp = work.groupby(work["companyName"].fillna("").str.strip(), dropna=False)
selected_parts = []
for cname, g in grp:
    if cname is None or str(cname).strip() == "":
        # skip records without company name
        continue
    selected_parts.append(pick_for_company(g, int(max_per_company), int(min_per_company)))

selected = pd.concat(selected_parts) if selected_parts else pd.DataFrame(columns=work.columns)

# Final output: include all original columns, plus a few selection metrics
if selected.empty:
    st.warning("No contacts selected after per-company allocation.")
    st.stop()

# Attach selection metadata
out = selected.copy()
out.insert(0, "_SelectedScore", out["_score"].round(3))
out.insert(1, "_SelectedBucket", out["_bucket"])
out.insert(2, "_SelectedSeniority", out["_seniority_final"])

# Summary
st.subheader("Summary")
sum1, sum2, sum3 = st.columns(3)
with sum1:
    st.metric("Selected contacts", f"{len(out):,}")
with sum2:
    st.metric("Companies covered", f"{out['companyName'].nunique():,}")
with sum3:
    st.metric("Avg per company", f"{(len(out)/max(out['companyName'].nunique(),1)):.1f}")

with st.expander("Per-company counts"):
    counts = out.groupby("companyName").size().reset_index(name="selected")
    st.dataframe(counts.sort_values("selected", ascending=False), use_container_width=True, height=320)

with st.expander("Preview selection"):
    st.dataframe(out.head(50), use_container_width=True, height=360)

# Download
def to_xlsx(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    try:
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Selected")
    except Exception:
        # fallback to openpyxl
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Selected")
    bio.seek(0)
    return bio.read()

xlsx_bytes = to_xlsx(out)
st.download_button(
    label="Download selected contacts (XLSX)",
    data=xlsx_bytes,
    file_name="selected_contacts.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

csv_bytes = out.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Download selected contacts (CSV)",
    data=csv_bytes,
    file_name="selected_contacts.csv",
    mime="text/csv",
)

st.caption("Tip: adjust focus keywords and per-company caps to tune the mix.")
