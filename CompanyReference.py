# streamlit_app.py
# Reference vs Target company matcher with fuzzy mapping and counts
# Requirements: streamlit, pandas, openpyxl, rapidfuzz, xlsxwriter (optional)

from io import BytesIO
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Company Match Counter", page_icon="🏷️", layout="wide")

# ---------------------------- Helpers ----------------------------

def read_table(uploaded_file: BytesIO) -> pd.DataFrame:
    """Read CSV or Excel into DataFrame. Keeps all columns."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file, engine="openpyxl")


def normalise_name(s: Optional[str]) -> str:
    if not isinstance(s, str):
        return ""
    s2 = s.strip()
    # Light normalisation only. Do not over-aggressively strip info.
    s2 = s2.replace("&", "and")
    # collapse whitespace
    s2 = " ".join(s2.split())
    return s2.lower()


def to_excel_bytes(df_dict: dict) -> bytes:
    buf = BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            for sheet, frame in df_dict.items():
                frame.to_excel(writer, index=False, sheet_name=sheet[:31])
    except Exception:
        # Fallback to openpyxl if xlsxwriter missing
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sheet, frame in df_dict.items():
                frame.to_excel(writer, index=False, sheet_name=sheet[:31])
    buf.seek(0)
    return buf.getvalue()

# ---------------------------- UI ----------------------------

st.title("Company Match Counter")

st.markdown(
    "Upload a reference list of all target companies, then a target contacts sheet. "
    "The app will count contacts per reference company with optional fuzzy matching."
)

with st.sidebar:
    st.header("Matching options")
    use_fuzzy = st.checkbox("Enable fuzzy matching", value=True)
    threshold = st.slider("Match threshold", min_value=50, max_value=100, value=92, step=1,
                          help="Minimum similarity score for a target company to map to a reference company")
    algo = st.selectbox("Fuzzy algorithm", ["token_set_ratio", "ratio", "partial_ratio"], index=0,
                        help="token_set_ratio is resilient to extra tokens and order changes")
    normalise = st.checkbox("Normalise names", value=True,
                            help="Lowercase, trim, replace '&' with 'and', collapse spaces")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1) Reference list")
    ref_file = st.file_uploader("Upload reference CSV or Excel", type=["csv", "xlsx", "xls"], key="ref")
    ref_df = None
    if ref_file is not None:
        ref_df = read_table(ref_file)
        st.dataframe(ref_df.head(10))
        # Column B per spec. Warn if not present.
        if ref_df.shape[1] < 2:
            st.error("Reference file must have company names in column B.")
        else:
            st.success("Using column B for company names, per spec.")

with col2:
    st.subheader("2) Target contacts")
    tgt_file = st.file_uploader("Upload target CSV or Excel", type=["csv", "xlsx", "xls"], key="tgt")
    tgt_df = None
    tgt_company_col = None
    if tgt_file is not None:
        tgt_df = read_table(tgt_file)
        if len(tgt_df.columns) == 0:
            st.error("Target file contains no columns.")
        else:
            tgt_company_col = st.selectbox("Select the company name column in target sheet", list(tgt_df.columns))
            st.dataframe(tgt_df.head(10))

run = st.button("Run matching")

# ---------------------------- Core logic ----------------------------

if run:
    if ref_df is None or tgt_df is None or tgt_company_col is None:
        st.error("Please upload both files and select the target company column.")
        st.stop()

    # Extract reference companies from column B
    ref_companies_raw = ref_df.iloc[:, 1].astype(str).fillna("")

    # Build canonical list and optional normalised versions
    ref_companies = ref_companies_raw.tolist()
    ref_index = list(range(len(ref_companies)))  # positions to rejoin later

    if normalise:
        ref_keys = [normalise_name(x) for x in ref_companies]
    else:
        ref_keys = [str(x) if isinstance(x, str) else "" for x in ref_companies]

    # Prepare target companies
    tgt_series_raw = tgt_df[tgt_company_col].astype(str).fillna("")
    tgt_series = tgt_series_raw.apply(normalise_name) if normalise else tgt_series_raw.astype(str)

    # Choose scorer
    scorer = {
        "token_set_ratio": fuzz.token_set_ratio,
        "ratio": fuzz.ratio,
        "partial_ratio": fuzz.partial_ratio,
    }[algo]

    # Build a map from target row to matched reference index
    # Then aggregate counts per reference index
    counts = {i: 0 for i in ref_index}
    matched_pairs = []  # for audit

    # RapidFuzz expects a list of choices. Use list of tuples to carry index
    choices = [(ref_keys[i], i) for i in ref_index]
    choice_strings = [c[0] for c in choices]

    for t_raw, t_key in zip(tgt_series_raw.tolist(), tgt_series.tolist()):
        best = None
        if use_fuzzy:
            res = process.extractOne(t_key, choice_strings, scorer=scorer)
            if res is not None:
                best_choice_str, score, best_pos = res[0], res[1], res[2]
                if score >= threshold:
                    best_index = choices[best_pos][1]
                    counts[best_index] += 1
                    best = (t_raw, ref_companies[best_index], int(score))
        else:
            # Exact match on normalised key
            try:
                pos = ref_keys.index(t_key)
                counts[pos] += 1
                best = (t_raw, ref_companies[pos], 100)
            except ValueError:
                pass
        if best:
            matched_pairs.append(best)

    total_contacts = len(tgt_df)
    count_list = [counts[i] for i in ref_index]
    pct_list = [(c / total_contacts * 100.0) if total_contacts else 0.0 for c in count_list]

    # Build output: original reference with counts and percent
    out_df = ref_df.copy()
    out_df["Matched_Contacts"] = count_list
    out_df["%_of_Total_Contacts"] = [round(p, 2) for p in pct_list]

    st.subheader("Results")
    st.dataframe(out_df.head(50))

    # Audit table
    audit_df = pd.DataFrame(matched_pairs, columns=["Target_Company", "Matched_Reference_Company", "Score"]).sort_values(
        by=["Score"], ascending=False
    )

    with st.expander("View match audit table"):
        st.dataframe(audit_df.head(500))

    # Downloads
    excel_bytes = to_excel_bytes({
        "Results": out_df,
        "Match_Audit": audit_df
    })

    st.download_button(
        label="Download Excel results",
        data=excel_bytes,
        file_name="company_match_counts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="Download CSV results only",
        data=out_df.to_csv(index=False).encode("utf-8"),
        file_name="company_match_counts.csv",
        mime="text/csv",
    )

    st.success("Done. Review counts and download outputs.")
