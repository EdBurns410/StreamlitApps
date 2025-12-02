# app_alt.py
import io
import re
import unicodedata
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

st.set_page_config(page_title="Company Matcher – Alternate Contacts", layout="wide")

# -----------------------------
# Schemas
# -----------------------------
TARGET_SCHEMA = [
    "Company",
    "Sector",
    "Theme",
    "Notes",
    "Where From",
]

ALT_CONTACT_SCHEMA = [
    "_source_file",
    "_source_sheet",
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
    "Seniority",
    "Result",
    "First Name",
    "Last Name",
    "Title",
    "Person Linkedin Url",
    "City",
    "State",
    "Country",
    "Email",
    "Company",
    "Website",
    "Industry",
    "# Employees",
    "Annual Revenue",
    "Total Funding",
    "Company Phone",
    "Company Linkedin Url",
    "Company Street",
    "Company City",
    "Company Postal Code",
    "Company State",
    "Company Country",
    "Company Founded Year",
]

# Words to strip from company names during normalisation
SUFFIXES = {
    "inc", "inc.", "ltd", "ltd.", "limited", "plc", "llc", "llp", "gmbh", "s.a.", "sa",
    "s.p.a", "spa", "bv", "b.v.", "ag", "corp", "corp.", "corporation", "company", "co",
    "co.", "group", "holdings", "holding", "technologies", "technology", "systems",
    "services", "solutions", "international", "int", "int.", "global", "the", "&", "bank"
}

# -----------------------------
# Helpers
# -----------------------------
def read_csv(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    try:
        return pd.read_csv(uploaded)
    except Exception:
        uploaded.seek(0)
        return pd.read_csv(uploaded, encoding="latin-1")

def auto_map_columns(df_cols: List[str], target_names: List[str]) -> Dict[str, str]:
    df_cols_lower = {c.lower(): c for c in df_cols}
    mapping = {}
    for name in target_names:
        key = name.lower()
        if key in df_cols_lower:
            mapping[name] = df_cols_lower[key]
        else:
            fallback = next((c for c in df_cols if clean_text(c) == clean_text(name)), None)
            mapping[name] = fallback if fallback else None
    return mapping

def clean_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def normalise_company(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = s.strip().lower()
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9& ]+", " ", s)
    tokens = [t for t in s.split() if t and t not in SUFFIXES]
    return " ".join(tokens)

def unique_contacts_count(df: pd.DataFrame, id_col_candidates: List[str]) -> int:
    for c in id_col_candidates:
        if c in df.columns:
            return df[c].nunique(dropna=True)
    key = df.get("Email", "").astype(str) + "|" + df.get("First Name", "").astype(str) + "|" + df.get("Last Name", "").astype(str)
    return key.nunique()

def build_company_index(contacts: pd.DataFrame, company_cols: List[str]) -> pd.DataFrame:
    rows = []
    for ridx, row in contacts.iterrows():
        for cc in company_cols:
            if cc in contacts.columns:
                val = row.get(cc, None)
                if pd.notna(val) and str(val).strip():
                    rows.append({
                        "__row_id__": ridx,
                        "__company_source__": cc,
                        "__company_raw__": str(val),
                        "__company_norm__": normalise_company(str(val))
                    })
    if not rows:
        return pd.DataFrame(columns=["__row_id__", "__company_source__", "__company_raw__", "__company_norm__"])
    return pd.DataFrame(rows).drop_duplicates(subset=["__row_id__", "__company_norm__"])

def best_match(target_name: str, contact_norms: pd.Series, scorer, score_cutoff: int) -> Tuple[int, str, int]:
    choices = contact_norms.fillna("").tolist()
    if not choices:
        return -1, "", 0
    res = process.extractOne(
        normalise_company(target_name),
        choices,
        scorer=scorer,
        score_cutoff=score_cutoff
    )
    if res is None:
        return -1, "", 0
    matched_value, score, idx = res
    true_index = contact_norms.index[idx]
    return true_index, matched_value, score

def run_matching(
    contacts_df: pd.DataFrame,
    targets_df: pd.DataFrame,
    target_company_col: str,
    contact_company_cols: List[str],
    id_col_candidates: List[str],
    scorer_name: str,
    score_cutoff: int,
):
    company_index = build_company_index(contacts_df, contact_company_cols)

    contacts_df = contacts_df.copy()
    contacts_df["__row_id__"] = np.arange(len(contacts_df))

    ci = company_index.merge(
        contacts_df,
        on="__row_id__",
        how="left",
        suffixes=("", "")
    )

    targets_df = targets_df.copy()
    targets_df["__target_norm__"] = targets_df[target_company_col].astype(str).map(normalise_company)

    scorer = {
        "token_set_ratio": fuzz.token_set_ratio,
        "token_sort_ratio": fuzz.token_sort_ratio,
        "ratio": fuzz.ratio,
        "partial_ratio": fuzz.partial_ratio,
    }[scorer_name]

    matched_rows = []
    contact_norm_series = ci["__company_norm__"] if not ci.empty else pd.Series(dtype="object")

    for _, trow in targets_df.iterrows():
        t_name = str(trow[target_company_col])
        t_norm = trow["__target_norm__"]

        # exact first
        matched_ci = ci[ci["__company_norm__"] == t_norm]
        source = "exact"
        if matched_ci.empty and not contact_norm_series.empty:
            idx, _, score = best_match(t_name, contact_norm_series, scorer, score_cutoff)
            if idx == -1:
                continue
            best_norm = ci.loc[idx, "__company_norm__"]
            matched_ci = ci[ci["__company_norm__"] == best_norm]
            source = f"fuzzy:{scorer_name}:{score}"

        if matched_ci.empty:
            continue

        matched_contacts = matched_ci.copy()
        matched_contacts["__target_company__"] = t_name
        matched_contacts["__match_type__"] = source
        matched_rows.append(matched_contacts)

    if matched_rows:
        matched_all = pd.concat(matched_rows, ignore_index=True)

        # Build matched contacts output
        matched_contacts_cols = [c for c in contacts_df.columns if not c.startswith("__")]
        matched_contacts_df = matched_all[
            ["__target_company__", "__company_raw__", "__company_source__", "__match_type__", "__row_id__"]
        ].merge(
            contacts_df[["__row_id__"] + matched_contacts_cols],
            on="__row_id__",
            how="left"
        )

        # Counts per target
        grouped = matched_contacts_df.groupby("__target_company__", as_index=False).apply(
            lambda g: pd.Series({"unique_contacts": unique_contacts_count(g, id_col_candidates)})
        ).reset_index(drop=True)
    else:
        matched_contacts_df = pd.DataFrame()
        grouped = pd.DataFrame(columns=["__target_company__", "unique_contacts"])

    # Summary: keep target CSV exactly as mapped, add Number of Matches
    targets_full = targets_df.drop(columns=[c for c in ["__target_norm__"] if c in targets_df.columns])
    counts = grouped.rename(columns={"__target_company__": "Company"})
    summary_df = targets_full.merge(counts, on="Company", how="left")
    summary_df["unique_contacts"] = summary_df["unique_contacts"].fillna(0).astype(int)
    summary_df = summary_df.rename(columns={"unique_contacts": "Number of Matches"})

    targets_no_match_df = summary_df.loc[summary_df["Number of Matches"] == 0, ["Company"]].copy()

    if not matched_contacts_df.empty:
        sort_keys = ["__target_company__", "__company_raw__"]
        if "Last Name" in matched_contacts_df.columns:
            sort_keys.append("Last Name")
        matched_contacts_df = matched_contacts_df.sort_values(sort_keys)

    return summary_df, matched_contacts_df, targets_no_match_df

def download_xlsx(sheets: Dict[str, pd.DataFrame]) -> bytes:
    engine = "xlsxwriter"
    try:
        import xlsxwriter  # noqa: F401
    except Exception:
        engine = "openpyxl"
    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine=engine) as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        return output.getvalue()

def column_mapper_ui(df: pd.DataFrame, target_schema: List[str], title: str, prefix: str) -> Dict[str, str]:
    st.subheader(title)
    st.caption("Auto detected identical headers. You can override below.")
    cols = df.columns.tolist()
    auto_map = auto_map_columns(cols, target_schema)

    mapping = {}
    ncols = 3
    rows = (len(target_schema) + ncols - 1) // ncols
    for r in range(rows):
        cols_row = st.columns(ncols, gap="small")
        for c in range(ncols):
            idx = r + rows * c
            if idx >= len(target_schema):
                continue
            target = target_schema[idx]
            options = ["Not present"] + cols
            default = auto_map[target] if auto_map[target] in cols else "Not present"
            key = f"{prefix}_{title}_{target}_{r}_{c}"
            sel = cols_row[c].selectbox(
                f"{target}",
                options,
                index=options.index(default),
                key=key
            )
            mapping[target] = None if sel == "Not present" else sel
    return mapping

def remap_df(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    out = pd.DataFrame()
    for target, src in mapping.items():
        if src is None:
            out[target] = pd.Series(dtype="object")
        else:
            out[target] = df[src]
    return out

# -----------------------------
# UI
# -----------------------------
st.title("Alternate Contacts → Target Company Matcher")

with st.expander("Matching options", expanded=True):
    scorer_name = st.selectbox(
        "Fuzzy scoring method",
        ["token_set_ratio", "token_sort_ratio", "ratio", "partial_ratio"],
        help="token_set_ratio is robust to word order and duplicates in company names"
    )
    score_cutoff = st.slider(
        "Minimum match score",
        min_value=60, max_value=100, value=88, step=1,
        help="Higher is stricter"
    )

c1, c2 = st.columns(2)
with c1:
    contacts_file = st.file_uploader("Upload Alternate Contacts CSV", type=["csv"], key="contacts_alt")
    contacts_df_raw = read_csv(contacts_file)
    if not contacts_df_raw.empty:
        st.success(f"Loaded contacts: {contacts_df_raw.shape[0]} rows, {contacts_df_raw.shape[1]} columns")

with c2:
    targets_file = st.file_uploader("Upload Target Companies CSV", type=["csv"], key="targets_alt")
    targets_df_raw = read_csv(targets_file)
    if not targets_df_raw.empty:
        st.success(f"Loaded targets: {targets_df_raw.shape[0]} rows, {targets_df_raw.shape[1]} columns")

if not contacts_df_raw.empty:
    mapping_contacts = column_mapper_ui(contacts_df_raw, ALT_CONTACT_SCHEMA, "Map Alternate Contacts columns", prefix="alt_contacts")
    contacts_df = remap_df(contacts_df_raw, mapping_contacts)
else:
    contacts_df = pd.DataFrame()

if not targets_df_raw.empty:
    mapping_targets = column_mapper_ui(targets_df_raw, TARGET_SCHEMA, "Map Target Companies columns", prefix="alt_targets")
    targets_df = remap_df(targets_df_raw, mapping_targets)
else:
    targets_df = pd.DataFrame()

st.markdown("---")

ready = not contacts_df.empty and not targets_df.empty and "Company" in targets_df.columns
if ready:
    st.subheader("Run matching")
    # Choose which company columns from the alternate schema to use
    suggested_cols = [c for c in ["companyName", "Company", "Company Name"] if c in contacts_df.columns]
    selectable_company_cols = [c for c in contacts_df.columns if "company" in c.lower()]
    company_cols = st.multiselect(
        "Contact company columns to match on",
        options=selectable_company_cols,
        default=suggested_cols or selectable_company_cols[:2],
        help="Include companyName if present"
    )

    # Reasonable ID candidates for unique counts
    id_candidates = [c for c in ["vmid", "companyId", "Email", "linkedInProfileUrl", "Person Linkedin Url"] if c in contacts_df.columns]

    go = st.button("Match now", type="primary")
    if go:
        with st.spinner("Matching..."):
            summary_df, matched_contacts_df, targets_no_match_df = run_matching(
                contacts_df=contacts_df,
                targets_df=targets_df,
                target_company_col="Company",
                contact_company_cols=company_cols,
                id_col_candidates=id_candidates,
                scorer_name=scorer_name,
                score_cutoff=score_cutoff,
            )

        st.success("Done")

        st.subheader("Summary – your target CSV plus Number of Matches")
        st.dataframe(summary_df, use_container_width=True)
        st.download_button(
            "Download summary CSV",
            data=summary_df.to_csv(index=False).encode("utf-8"),
            file_name="alt_summary_target_company_counts.csv",
            mime="text/csv"
        )

        st.subheader("Matched contacts grouped by target company")
        if matched_contacts_df.empty:
            st.info("No matches found at the chosen threshold")
        else:
            display_cols = [c for c in ["__target_company__", "__match_type__", "__company_source__", "__company_raw__"] if c in matched_contacts_df.columns]
            contact_cols = [c for c in matched_contacts_df.columns if not c.startswith("__")]
            show_cols = display_cols + contact_cols
            st.dataframe(matched_contacts_df[show_cols], use_container_width=True)
            st.download_button(
                "Download matched contacts CSV",
                data=matched_contacts_df[show_cols].to_csv(index=False).encode("utf-8"),
                file_name="alt_matched_contacts.csv",
                mime="text/csv"
            )

        st.subheader("Target companies with no matches")
        st.dataframe(targets_no_match_df, use_container_width=True)
        st.download_button(
            "Download no-match targets CSV",
            data=targets_no_match_df.to_csv(index=False).encode("utf-8"),
            file_name="alt_targets_no_match.csv",
            mime="text/csv"
        )

        # XLSX bundle
        xlsx_bytes = download_xlsx({
            "Summary": summary_df,
            "Matched Contacts": matched_contacts_df if not matched_contacts_df.empty else pd.DataFrame(),
            "No Match Targets": targets_no_match_df
        })
        st.download_button(
            "Download all results as XLSX",
            data=xlsx_bytes,
            file_name="alt_company_matching_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
