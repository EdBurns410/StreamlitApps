# merge_excels_app.py
# Upload multiple .xlsx and .csv files, map to a fixed target schema, preview, and download merged output.
# Requirements: streamlit, pandas, openpyxl, python-levenshtein

from io import BytesIO
import json
import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

try:
    # Optional but useful for better auto-matching
    from Levenshtein import ratio as lev_ratio
except Exception:
    def lev_ratio(a, b):  # fallback if python-levenshtein is not installed
        return 1.0 if a.lower() == b.lower() else 0.0

st.set_page_config(page_title="Excel and CSV Merger to Fixed Schema", page_icon="📎", layout="wide")

# ------------------------------- Target schema -------------------------------
TARGET_HEADERS: List[str] = [
    "profileUrl",
    "fullName",
    "firstName",
    "lastName",
    "companyName",
    "title",
    "companyId",
    "companyUrl",
    "regularCompanyUrl",
    "summary",
    "titleDescription",
    "industry",
    "companyLocation",
    "location",
    "durationInRole",
    "durationInCompany",
    "pastExperienceCompanyName",
    "pastExperienceCompanyUrl",
    "pastExperienceCompanyTitle",
    "pastExperienceDate",
    "pastExperienceDuration",
    "connectionDegree",
    "profileImageUrl",
    "sharedConnectionsCount",
    "name",
    "vmid",
    "linkedInProfileUrl",
    "isPremium",
    "isOpenLink",
    "query",
    "timestamp",
    "defaultProfileUrl",
    "searchAccountProfileId",
    "searchAccountProfileName",
]

# Common aliases to improve auto-suggestions
ALIASES: Dict[str, List[str]] = {
    "linkedInProfileUrl": ["linkedin", "linkedin_url", "linkedin profile", "profile url", "li_url", "profileUrl"],
    "profileUrl": ["profile url", "li_url", "linkedin_url"],
    "companyUrl": ["company url", "company_url", "employer_url"],
    "regularCompanyUrl": ["website", "company website", "domain", "root domain", "regular_url"],
    "companyName": ["employer", "organization", "org", "company"],
    "title": ["job title", "position", "role"],
    "fullName": ["name_full", "candidate", "contact", "fullname", "displayName"],
    "firstName": ["fname", "first name", "given_name"],
    "lastName": ["lname", "last name", "surname", "family_name"],
    "profileImageUrl": ["avatar", "photo", "image", "picture"],
    "connectionDegree": ["degree", "conn_degree", "relationship"],
    "isPremium": ["premium", "linkedin_premium"],
    "isOpenLink": ["open_link", "openlink"],
    "companyId": ["company_id", "employer_id"],
    "companyLocation": ["hq_location", "company_city", "employer_location"],
    "durationInRole": ["role_duration", "tenure_role"],
    "durationInCompany": ["company_tenure", "tenure_company"],
    "pastExperienceCompanyName": ["prev_company", "past_company", "experience_company"],
    "pastExperienceCompanyUrl": ["prev_company_url", "past_company_url"],
    "pastExperienceCompanyTitle": ["prev_title", "past_title", "experience_title"],
    "pastExperienceDate": ["prev_date", "experience_date", "dates_employed"],
    "pastExperienceDuration": ["prev_duration", "experience_duration"],
    "vmid": ["entity_urn", "vanity_id"],
    "query": ["search_query", "term", "keywords"],
    "timestamp": ["created_at", "exported_at", "scraped_at", "ts"],
    "defaultProfileUrl": ["fallback_profile_url", "profile_url_default"],
    "searchAccountProfileId": ["account_profile_id", "salesnav_account_id"],
    "searchAccountProfileName": ["account_profile_name", "salesnav_account_name"],
}

# ------------------------------- Utilities -------------------------------
def normalise_header(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().lower()
    s = re.sub(r"[\s\-]+", "_", s)
    s = re.sub(r"[^0-9a-z_]", "", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def tidy_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df

def load_excel_tables(file, read_all_sheets: bool) -> List[Tuple[pd.DataFrame, str, str]]:
    """Read first or all sheets from an Excel file. Row 1 is header."""
    results = []
    if read_all_sheets:
        book = pd.read_excel(file, sheet_name=None, header=0, engine="openpyxl")
        for sheetname, df in book.items():
            results.append((tidy_headers(df), file.name, sheetname))
    else:
        df = pd.read_excel(file, sheet_name=0, header=0, engine="openpyxl")
        results.append((tidy_headers(df), file.name, "Sheet1"))
    return results

def load_csv_table(file) -> List[Tuple[pd.DataFrame, str, str]]:
    """Read a CSV file. Row 1 is header. Use sep=None to auto-detect delimiter."""
    try:
        df = pd.read_csv(file, header=0, sep=None, engine="python")  # auto-detects comma, tab, semicolon
    except Exception:
        file.seek(0)
        df = pd.read_csv(file, header=0)
    return [(tidy_headers(df), file.name, "CSV")]

def load_any_tables(file, read_all_sheets: bool) -> List[Tuple[pd.DataFrame, str, str]]:
    name = file.name.lower()
    if name.endswith(".xlsx"):
        return load_excel_tables(file, read_all_sheets)
    if name.endswith(".csv"):
        return load_csv_table(file)
    raise ValueError(f"Unsupported file type for {file.name}. Only .xlsx and .csv are supported.")

def suggest_source_for_target(target: str, source_headers: List[str]) -> Optional[str]:
    # Try exact norm match
    t_norm = normalise_header(target)
    best = None
    best_score = 0.0
    # Candidate list augmented with aliases
    alias_list = [target] + ALIASES.get(target, [])
    alias_norms = [normalise_header(a) for a in alias_list]

    for s in source_headers:
        s_norm = normalise_header(s)
        score = 0.0
        if s_norm in alias_norms:
            score = 1.0
        else:
            # string similarity as fallback
            score = max(lev_ratio(s_norm, t_norm), *(lev_ratio(s_norm, an) for an in alias_norms))
        if score > best_score:
            best_score = score
            best = s

    # Only accept if similarity is decent
    return best if best_score >= 0.72 else None

def make_excel_download(df: pd.DataFrame, sheet_name: str = "Merged") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return out.getvalue()

def make_csv_download(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

# ------------------------------- UI -------------------------------
st.title("📎 Merge Excel and CSV files to fixed schema")
st.caption("Row 1 is treated as header. Map columns to the required schema, preview, and download.")

st.sidebar.header("Options")
read_all_sheets = st.sidebar.checkbox("Read all sheets from each Excel workbook", value=False,
                                      help="For .xlsx only. CSVs have no sheets.")
dedupe_enabled = st.sidebar.checkbox("Enable de-duplication", value=False)
sample_rows = st.sidebar.number_input("Preview rows", min_value=10, max_value=2000, value=100, step=10)

uploaded = st.file_uploader(
    "Upload .xlsx and .csv files",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="You can mix Excel and CSV files. For Excel, row 1 is used as the header on the first or all sheets."
)

# Mapping persistence
st.subheader("Column mapping")
c_load, c_save = st.columns([1, 1])
with c_load:
    mapping_file = st.file_uploader("Load mapping (.json)", type=["json"], key="map_upload")

if uploaded:
    # Load all tables from all files
    dfs: List[Tuple[pd.DataFrame, str, str]] = []
    for f in uploaded:
        try:
            dfs.extend(load_any_tables(f, read_all_sheets))
        except Exception as e:
            st.error(f"Failed to read {f.name}: {e}")

    if not dfs:
        st.stop()

    st.success(f"Loaded {len(dfs)} table(s) from {len(uploaded)} file(s).")

    with st.expander("View loaded tables"):
        inv = pd.DataFrame(
            [{"file": fn, "sheet_or_type": sh, "rows": df.shape[0], "cols": df.shape[1]} for df, fn, sh in dfs]
        ).sort_values(["file", "sheet_or_type"])
        st.dataframe(inv, use_container_width=True, height=220)

    # Collect source headers across all uploads
    all_source_headers: List[str] = []
    for df, _, _ in dfs:
        all_source_headers.extend(list(df.columns))
    all_source_headers = sorted(pd.unique([str(h).strip() for h in all_source_headers]).tolist())

    # Build initial mapping: target -> suggested source
    initial_map: Dict[str, Optional[str]] = {}
    for t in TARGET_HEADERS:
        initial_map[t] = suggest_source_for_target(t, all_source_headers)

    # If a mapping JSON was uploaded, override suggestions
    if mapping_file is not None:
        try:
            loaded_map: Dict[str, str] = json.load(mapping_file)
            for t in TARGET_HEADERS:
                src = loaded_map.get(t)
                if src in all_source_headers:
                    initial_map[t] = src
            st.info("Loaded mapping from JSON.")
        except Exception as e:
            st.error(f"Could not read mapping JSON: {e}")

    # Interactive mapping table
    st.write("Select, for each target header, the source column to map from. Leave blank to create an empty column.")
    editable_rows = [{"target_header": t, "source_column": initial_map.get(t) or ""} for t in TARGET_HEADERS]
    map_df = pd.DataFrame(editable_rows)

    map_df = st.data_editor(
        map_df,
        use_container_width=True,
        num_rows="fixed",
        height=min(900, 80 + 28 * len(map_df)),
        column_config={
            "target_header": st.column_config.TextColumn("Target header", disabled=True, width="large"),
            "source_column": st.column_config.SelectboxColumn(
                "Source column",
                options=[""] + all_source_headers,
                required=False,
                width="large",
                help="Pick the column from your uploads that should fill this target field."
            ),
        },
        key="map_editor_fixed",
    )

    # Build mapping dict target -> source or None
    target_to_source: Dict[str, Optional[str]] = {
        row["target_header"]: (row["source_column"] if row["source_column"] else None)
        for _, row in map_df.iterrows()
    }

    with c_save:
        if st.button("Download mapping JSON"):
            mapping_json = json.dumps(target_to_source, indent=2).encode("utf-8")
            st.download_button(
                "Save mapping.json",
                data=mapping_json,
                file_name="fixed_schema_mapping.json",
                mime="application/json",
                key="save_map_btn"
            )

    # Apply mapping to each dataframe
    remapped: List[pd.DataFrame] = []
    for df, fn, sh in dfs:
        out = pd.DataFrame()
        for t in TARGET_HEADERS:
            src = target_to_source.get(t)
            if src and src in df.columns:
                out[t] = df[src]
            else:
                out[t] = pd.Series([None] * len(df))
        # Provenance
        out.insert(0, "_source_sheet_or_type", sh)
        out.insert(0, "_source_file", fn)
        remapped.append(out)

    merged = pd.concat(remapped, ignore_index=True)

    # Optional de-duplication by selected keys
    if dedupe_enabled:
        with st.expander("De-duplication settings", expanded=False):
            dedupe_cols = st.multiselect(
                "Choose columns to define duplicates",
                options=["_source_file", "_source_sheet_or_type"] + TARGET_HEADERS,
                default=["linkedInProfileUrl"] if "linkedInProfileUrl" in TARGET_HEADERS else [],
            )
            if dedupe_cols:
                before = len(merged)
                merged = merged.drop_duplicates(subset=dedupe_cols, keep="first")
                after = len(merged)
                st.caption(f"Removed {before - after} duplicate rows using {dedupe_cols}.")

    # Preview
    st.subheader("Preview")
    st.caption(f"Showing first {min(sample_rows, len(merged))} rows. Total rows: {len(merged)}. Columns: {len(merged.columns)}.")
    st.dataframe(merged.head(sample_rows), use_container_width=True, height=420)

    # Downloads
    st.subheader("Download")
    c1, c2 = st.columns(2)
    with c1:
        csv_bytes = make_csv_download(merged)
        st.download_button(
            "Download CSV",
            data=csv_bytes,
            file_name="merged_fixed_schema.csv",
            mime="text/csv"
        )
    with c2:
        xlsx_bytes = make_excel_download(merged, sheet_name="Merged")
        st.download_button(
            "Download Excel",
            data=xlsx_bytes,
            file_name="merged_fixed_schema.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Upload at least one .xlsx or .csv file to begin.")
