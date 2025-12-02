"""
Company Standardiser + Top 5 Senior Contacts - Seniority aware
- Upload CSV or XLSX (multi-file, multi-sheet)
- Standardise companyName into companyName_standard with reviewable mapping
- Primary seniority ranking from your 'Seniority' column
- Optional spread across top senior buckets before filling to 5
- Exports a Top-5 CSV in your exact schema
"""

import io
import re
from typing import List, Dict, Tuple, Optional

import pandas as pd
import streamlit as st

# Optional fuzzy library. Falls back to difflib if missing.
try:
    from rapidfuzz import fuzz
    def _ratio(a, b): return fuzz.token_set_ratio(a, b)
except Exception:
    from difflib import SequenceMatcher
    def _ratio(a, b): return int(100 * SequenceMatcher(None, a, b).ratio())

st.set_page_config(page_title="Company Standardiser + Top 5 Senior Contacts", page_icon="🏷️", layout="wide")

# ----------------------------- Helpers -----------------------------

COMPANY_SUFFIXES = {
    "limited", "ltd", "plc", "llc", "inc", "inc.", "corp", "corporation",
    "gmbh", "ag", "sa", "s.a.", "bv", "b.v.", "oy", "ab", "nv", "srl", "s.r.l.",
    "pty", "pty ltd", "kk", "k.k.", "co", "co.", "company", "holdings", "group",
    "lp", "llp", "sasu", "sas", "spa", "s.p.a.", "ulc", "p.l.c.", "pte", "pte. ltd"
}
COUNTRY_TOKENS = {
    "uk", "usa", "us", "u.s.", "u.s.a.", "ireland", "ire", "gb", "uae", "eu",
    "europe", "singapore", "s’pore", "sng", "hk", "hong", "kong", "china",
    "japan", "india", "canada", "ca", "aus", "australia"
}

def normalise_company(raw: str) -> str:
    if pd.isna(raw):
        return ""
    s = str(raw).strip()
    s = re.sub(r"\(.*?\)|\[.*?\]|{.*?}", " ", s)  # remove bracketed content
    s = re.sub(r"[+/,_\.]", " ", s)
    s = s.replace("&", " and ")
    s = s.lower()
    s = re.sub(r"^the\s+", "", s)
    tokens = [t for t in re.split(r"\s+", s) if t]
    cleaned = []
    for t in tokens:
        t_clean = re.sub(r"[^a-z0-9\-]", "", t)
        if t_clean in COMPANY_SUFFIXES or t_clean in COUNTRY_TOKENS:
            continue
        cleaned.append(t_clean)
    s = " ".join(cleaned)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" - ", "-")
    return s

def title_case_company(s: str) -> str:
    if not s:
        return s
    acronyms = {"hsbc", "ubs", "bbva", "bny", "tsb", "abn", "cibc", "rbs", "bbt"}
    out = []
    for t in s.split():
        out.append(t.upper() if t in acronyms or t.upper() in acronyms else t.capitalize())
    return " ".join(out)

def propose_standard_name(company: str) -> str:
    core = normalise_company(company)
    return title_case_company(core) if core else company

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    return None

# Fallback title-based seniority scoring used only if 'Seniority' column not provided
SENIORITY_RULES_TITLE: List[Tuple[str, int, str]] = [
    (r"\b(chief|c[\.\s]*o\b|cdo\b|cfo\b|cio\b|cto\b|cso\b|ciso\b)\b", 100, "C-Suite"),
    (r"\b(managing director|md)(?![a-z])", 96, "Managing Director"),
    (r"\b(president|co-?founder|founder|owner|partner)\b", 94, "Founder or President"),
    (r"\b(executive vice president|evp)\b", 92, "EVP"),
    (r"\b(senior vice president|svp)\b", 90, "SVP"),
    (r"\b(vice president|vp)\b", 80, "Vice President"),
    (r"\b(global head|head of|head)\b", 72, "Head"),
    (r"\b(director|principal)\b", 70, "Director"),
    (r"\b(associate director|ad)\b", 66, "Associate Director"),
    (r"\b(manager|mgr)\b", 60, "Manager"),
    (r"\b(lead)\b", 55, "Lead"),
    (r"\b(senior|sr\.)\b", 50, "Senior"),
    (r"\b(consultant|specialist|architect)\b", 46, "Consultant"),
    (r"\b(associate)\b", 40, "Associate"),
    (r"\b(analyst|researcher)\b", 35, "Analyst"),
    (r"\b(coordinator|administrator)\b", 30, "Coordinator"),
    (r"\b(intern|trainee|apprentice)\b", 15, "Entry-Level"),
]
def score_from_title(title: str) -> Tuple[int, str]:
    if pd.isna(title):
        return 0, "Unknown"
    t = " " + str(title).lower().strip() + " "
    if re.search(r"\bglobal lead\b", t):
        return 68, "Global Lead"
    for pattern, score, bucket in SENIORITY_RULES_TITLE:
        if re.search(pattern, t):
            return score, bucket
    return 0, "Unknown"

# Your Seniority categories and ranking
SENIORITY_CANON = {
    "Chair/Board Member": 102,
    "C-Suite": 100,
    "Vice President": 90,
    "Head": 85,
    "Director": 80,
    "Leadership": 72,
    "Manager": 65,
    "Senior": 55,
    "Consultant": 50,
    "Non-Senior Banker": 40,
    "Entry-Level": 10,
}
PREFERRED_SPREAD_ORDER = ["C-Suite", "Chair/Board Member", "Vice President", "Director", "Head"]

def canonicalise_seniority(value: str) -> str:
    if pd.isna(value):
        return "Unknown"
    s = re.sub(r"[\s/_-]+", " ", str(value)).strip().lower()
    # direct matches
    for k in SENIORITY_CANON.keys():
        if s == k.lower():
            return k
    # light synonym handling
    if s in {"board", "board member", "chair", "chairman", "chairwoman", "chair person", "chairperson"}:
        return "Chair/Board Member"
    if s in {"c suite", "csuite"}:
        return "C-Suite"
    if s in {"vp", "v p"}:
        return "Vice President"
    if s.startswith("head "):
        return "Head"
    if s in {"entry level"}:
        return "Entry-Level"
    if s in {"non senior banker", "non-senior banker"}:
        return "Non-Senior Banker"
    if s == "director":
        return "Director"
    if s == "leadership":
        return "Leadership"
    if s == "manager":
        return "Manager"
    if s == "senior":
        return "Senior"
    if s == "consultant":
        return "Consultant"
    return value.strip()

def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, sheet in sheets.items():
            sheet.to_excel(writer, index=False, sheet_name=name[:31])
    bio.seek(0)
    return bio.getvalue()

# ----------------------------- UI -----------------------------

st.title("🏷️ Company Standardiser + Top 5 Senior Contacts")
st.markdown("Upload CSV or XLSX, standardise company names, then export the five most senior contacts per company. Edit the mapping if needed.")

with st.sidebar:
    st.header("Settings")
    sim_threshold = st.slider("Fuzzy similarity threshold", 80, 100, 94, help="Higher is stricter. 100 disables fuzzy merge.")
    prefer_spread = st.checkbox("Prefer spread across top senior buckets", value=True,
                                help="Try to include one of each: C-Suite, Chair/Board Member, VP, Director, Head before filling remaining slots.")
    auto_apply = st.checkbox("Auto-apply proposed mapping", value=False, help="If ticked, skips manual mapping review.")
    st.markdown("---")
    st.subheader("Optional: load a previous mapping")
    prior_map_file = st.file_uploader("company_mapping.csv", type=["csv"], accept_multiple_files=False)

uploaded = st.file_uploader("Upload CSV or XLSX file(s)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
if not uploaded:
    st.info("Upload at least one CSV or XLSX to begin.")
    st.stop()

# ----------------------------- Read files (multi-sheet aware) -----------------------------

frames = []
for f in uploaded:
    name = f.name
    if name.lower().endswith(".csv"):
        df0 = pd.read_csv(f, dtype=str, keep_default_na=False)
        df0["_source_file"] = name
        df0["_source_sheet_or_type"] = "csv"
        frames.append(df0)
    else:
        raw = f.read()
        xls = pd.ExcelFile(io.BytesIO(raw))
        for sheet in xls.sheet_names:
            df0 = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, dtype=str, keep_default_na=False)
            df0["_source_file"] = name
            df0["_source_sheet_or_type"] = sheet
            frames.append(df0)

df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
df = df.applymap(lambda x: pd.NA if isinstance(x, str) and x.strip() == "" else x)

# ----------------------------- Column mapping -----------------------------

st.subheader("Column mapping")

company_col_guess = find_col(df, ["companyName", "company", "company_name", "Company"])
title_col_guess = find_col(df, ["title", "jobTitle", "job_title", "Title"])
first_col_guess = find_col(df, ["firstName", "first_name", "First Name", "first name"])
last_col_guess = find_col(df, ["lastName", "last_name", "Last Name", "last name"])
full_col_guess = find_col(df, ["fullName", "name", "Name", "Full Name"])
email_col_guess = find_col(df, ["email", "workEmail", "Email", "e-mail", "businessEmail"])
loc_col_guess = find_col(df, ["location", "city", "City", "Region", "Country", "country", "Location"])
li_col_guess = find_col(df, ["linkedInProfileUrl", "linkedin", "LinkedIn", "LinkedIn URL", "profileUrl", "linkedinProfileUrl"])
seniority_col_guess = find_col(df, ["Seniority", "seniority"])

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    company_col = st.selectbox("Company column", options=list(df.columns), index=list(df.columns).index(company_col_guess) if company_col_guess in df.columns else 0)
with col2:
    title_col = st.selectbox("Title column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(title_col_guess) + 1) if title_col_guess in df.columns else 0)
with col3:
    first_col = st.selectbox("First name column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(first_col_guess) + 1) if first_col_guess in df.columns else 0)
with col4:
    last_col = st.selectbox("Last name column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(last_col_guess) + 1) if last_col_guess in df.columns else 0)
with col5:
    full_col = st.selectbox("Full name column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(full_col_guess) + 1) if full_col_guess in df.columns else 0)

col6, col7, col8 = st.columns(3)
with col6:
    email_col = st.selectbox("Email column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(email_col_guess) + 1) if email_col_guess in df.columns else 0)
with col7:
    loc_col = st.selectbox("Location column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(loc_col_guess) + 1) if loc_col_guess in df.columns else 0)
with col8:
    li_col = st.selectbox("LinkedIn URL column", options=["<none>"] + list(df.columns), index=(list(df.columns).index(li_col_guess) + 1) if li_col_guess in df.columns else 0)

# Seniority column selector
seniority_col = st.selectbox("Seniority column (preferred)", options=["<none>"] + list(df.columns),
                             index=(list(df.columns).index(seniority_col_guess) + 1) if seniority_col_guess in df.columns else 0)

if not company_col:
    st.error("Select the company column before continuing.")
    st.stop()

# Build first/last if needed from full name
if (first_col == "<none>" or last_col == "<none>") and full_col != "<none>":
    tmp = df[full_col].fillna("").astype(str).str.strip()
    df["__firstName_tmp"] = tmp.apply(lambda s: s.split()[0] if s else pd.NA)
    df["__lastName_tmp"] = tmp.apply(lambda s: " ".join(s.split()[1:]) if len(s.split()) > 1 else pd.NA)
    if first_col == "<none>":
        first_col = "__firstName_tmp"
    if last_col == "<none>":
        last_col = "__lastName_tmp"

# ----------------------------- Seniority scoring -----------------------------

if seniority_col != "<none>":
    # Use provided Seniority column
    df["seniority_bucket"] = df[seniority_col].apply(canonicalise_seniority)
    df["seniority_score"] = df["seniority_bucket"].apply(lambda x: SENIORITY_CANON.get(x, 0))
else:
    # Fallback to title parsing
    if title_col == "<none>":
        st.warning("No Seniority or Title column found. Seniority will be Unknown.")
        df["seniority_bucket"] = "Unknown"
        df["seniority_score"] = 0
    else:
        scores = df[title_col].apply(score_from_title)
        df["seniority_score"] = scores.apply(lambda t: t[0])
        df["seniority_bucket"] = scores.apply(lambda t: t[1])

# ----------------------------- Company mapping proposal -----------------------------

unique_companies = sorted([c for c in df[company_col].dropna().astype(str).unique()], key=lambda x: x.lower())
# Build proposed mapping via normalised grouping and fuzzy merge
def cluster_names(unique_names: List[str], threshold: int = 94) -> Dict[str, str]:
    normalised_map = {n: normalise_company(n) for n in unique_names}
    groups: Dict[str, List[str]] = {}
    for orig, core in normalised_map.items():
        groups.setdefault(core, []).append(orig)
    mapping: Dict[str, str] = {}
    for core, members in groups.items():
        canonical = title_case_company(core) if core else title_case_company(members[0])
        for m in members:
            mapping[m] = canonical
    cores = list(groups.keys())
    for i, c in enumerate(cores):
        for j in range(i + 1, len(cores)):
            d = cores[j]
            if not c or not d:
                continue
            if _ratio(c, d) >= threshold:
                canon = title_case_company(min(c, d, key=len))
                for orig, core in normalised_map.items():
                    if core in (c, d):
                        mapping[orig] = canon
    return mapping

mapping = cluster_names(unique_companies, threshold=sim_threshold)
mapping_df = pd.DataFrame({"original_companyName": unique_companies,
                           "proposed_standard": [mapping.get(c, propose_standard_name(c)) for c in unique_companies]})

# Optional prior mapping
if prior_map_file is not None:
    try:
        prior_df = pd.read_csv(prior_map_file, dtype=str)
        if {"original_companyName", "proposed_standard"}.issubset(prior_df.columns):
            prior_s = pd.Series(prior_df["proposed_standard"].values, index=prior_df["original_companyName"].values)
            mapping_df["proposed_standard"] = mapping_df.apply(
                lambda r: prior_s.get(r["original_companyName"], r["proposed_standard"]),
                axis=1
            )
            st.success("Loaded prior mapping and applied overrides.")
        else:
            st.warning("Prior mapping CSV must have: original_companyName, proposed_standard.")
    except Exception as e:
        st.warning(f"Could not read prior mapping: {e}")

st.subheader("Step 1 - Review proposed company mapping")
st.caption("Edit the proposed_standard values if needed. The list is A to Z by original name.")
edited_mapping_df = mapping_df.copy() if auto_apply else st.data_editor(mapping_df, num_rows="fixed", use_container_width=True, key="mapping_editor")

# Apply mapping to data
map_series = pd.Series(edited_mapping_df["proposed_standard"].values, index=edited_mapping_df["original_companyName"].values)
df["companyName_standard"] = df[company_col].map(map_series).fillna(df[company_col].apply(propose_standard_name))

st.success(f"Applied mapping to {df[company_col].notna().sum():,} rows from {len(unique_companies):,} unique companies.")

# ----------------------------- Top 5 selection with optional spread -----------------------------

st.subheader("Step 2 - Top 5 most senior contacts per company")

def pick_top5_with_spread(group: pd.DataFrame, spread: bool = True) -> pd.DataFrame:
    g = group.sort_values(["seniority_score"], ascending=[False]).copy()
    if not spread:
        return g.head(5)

    selected_idx = []
    # 1 per preferred bucket if available
    for bucket in PREFERRED_SPREAD_ORDER:
        rows = g[g["seniority_bucket"] == bucket]
        if not rows.empty:
            selected_idx.append(rows.index[0])
        if len(selected_idx) >= 5:
            break

    # Fill the rest by score
    if len(selected_idx) < 5:
        for idx in g.index:
            if idx not in selected_idx:
                selected_idx.append(idx)
                if len(selected_idx) >= 5:
                    break

    return g.loc[selected_idx]

top5_list = []
for cname, grp in df.groupby("companyName_standard", sort=False):
    top5_list.append(pick_top5_with_spread(grp, prefer_spread))
top5 = pd.concat(top5_list) if top5_list else df.head(0)

# ----------------------------- Build Top-5 export in your schema -----------------------------

schema_cols = ["_source_file", "_source_sheet_or_type", "firstName", "lastName",
               "companyName", "title", "location", "Email", "linkedInProfileUrl"]

def safe_get_from(df_src: pd.DataFrame, col_name: str) -> pd.Series:
    if col_name == "firstName":
        return df_src[first_col] if first_col != "<none>" and first_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "lastName":
        return df_src[last_col] if last_col != "<none>" and last_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "title":
        return df_src[title_col] if title_col != "<none>" and title_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "location":
        return df_src[loc_col] if loc_col != "<none>" and loc_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "Email":
        return df_src[email_col] if email_col != "<none>" and email_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "linkedInProfileUrl":
        return df_src[li_col] if li_col != "<none>" and li_col in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "_source_file":
        return df_src["_source_file"] if "_source_file" in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "_source_sheet_or_type":
        return df_src["_source_sheet_or_type"] if "_source_sheet_or_type" in df_src.columns else pd.Series([pd.NA] * len(df_src))
    if col_name == "companyName":
        return df_src["companyName_standard"] if "companyName_standard" in df_src.columns else df_src[company_col]
    return pd.Series([pd.NA] * len(df_src))

top5_schema = pd.DataFrame({c: safe_get_from(top5, c).values for c in schema_cols})

st.markdown("**Preview - Top 5 per company (schema view)**")
st.dataframe(top5_schema.head(50), use_container_width=True)

# ----------------------------- Downloads -----------------------------

st.subheader("Step 3 - Download outputs")

# Mapping CSV
mapping_csv = edited_mapping_df.to_csv(index=False).encode("utf-8")
st.download_button("Download company mapping CSV", mapping_csv, file_name="company_mapping.csv", mime="text/csv")

# Top 5 schema CSV
top5_csv = top5_schema.to_csv(index=False).encode("utf-8")
st.download_button("Download Top 5 per company (schema CSV)", top5_csv, file_name="top5_per_company_schema.csv", mime="text/csv")

# Enriched master CSV
master_cols = ["_source_file", "_source_sheet_or_type", company_col, "companyName_standard",
               "seniority_bucket", "seniority_score"]
for c in [first_col, last_col, full_col, title_col, loc_col, email_col, li_col]:
    if c and c != "<none>" and c in df.columns and c not in master_cols:
        master_cols.append(c)
master_export = df[master_cols].copy()
master_csv = master_export.to_csv(index=False).encode("utf-8")
st.download_button("Download Enriched Master (CSV)", master_csv, file_name="master_enriched.csv", mime="text/csv")

# Excel with all sheets
def to_excel_bytes_local() -> bytes:
    return to_excel_bytes({
        "Top5_schema": top5_schema,
        "Master_enriched": master_export,
        "Company_mapping": edited_mapping_df
    })
wb = to_excel_bytes_local()
st.download_button("Download Excel with all sheets", wb, file_name="company_standardised_outputs.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------- Summary -----------------------------

st.divider()
st.subheader("Summary")
n_companies_before = len(unique_companies)
n_companies_after = top5["companyName"].nunique() if "companyName" in top5.columns else top5["companyName_standard"].nunique()
st.write(f"- Unique companies before standardising: **{n_companies_before:,}**")
st.write(f"- Unique companies after standardising: **{n_companies_after:,}**")
st.write(f"- Rows with ranked seniority: **{int((df['seniority_score'] > 0).sum()):,}**")
st.write(f"- Total rows processed: **{len(df):,}**")
