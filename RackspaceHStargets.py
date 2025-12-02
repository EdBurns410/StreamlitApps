# streamlit_app.py
import io
import unicodedata
import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Contact Filter – Titles & Banks", layout="wide")

# ---------- Targets ----------
RAW_TITLES = [
    "Chief Information Security Officer",
    "CISO",
    "Head of IT Risk and Compliance",
    "Group CIO / CIO / Chief Information Officer",
    "Director, IT Strategy and Transformation",
    "IT Strategy Director",
    "Group CTO / CTO / Chief Technology Officer",
    "Global CTO",
    "Head of IT",
    "IT Director",
    "Director of Technology",
    "Technology and Operations Director",
    "Chief Operating Officer for IT / Information Technology",
    "Director of Operations and Transformation",
    "VP Infrastructure and Operations",
    "Head of Infrastructure",
    "VP Cloud Infrastructure",
    "Director of Cloud Infrastructure",
    "Head of Enterprise Architecture",
    "Chief Data Officer",
    "IT Transformation Director / Head of IT Transformation",
]

BANKS = [
    "Aareal Bank AG","ABANCA","ABC International Bank Plc","ABN AMRO Clearing Bank","AIB Group (UK) Plc",
    "Allica Bank","Alpha Bank London Ltd","Arbuthnot Latham & Co Ltd","Banca Monte dei Paschi di Siena (BMPS)",
    "BANCA SELLA SPA","Bank Of England","Bank of Ireland (UK) Plc","Bank of London and the Middle East Plc",
    "Bper Banca","British Arab Commercial Bank Plc","C Hoare & Co","Caceis","CAF Bank Ltd","CaixaBank",
    "Close Brothers Ltd","Commerzbank AG","Coutts & Co","Danske Bank UK","Deutsche Pfandbriefbank AG","EBRD",
    "EFG International","GB Bank Ltd","Groupe Bpce","GULF INTERNATIONAL BANK (UK) LIMITED","Habib Bank AG Zurich",
    "Halifax","Handlesbanken","ICBC Standard Bank","Intesa Sanpaolo","Investec Bank PLC","Kexim Bank (UK) Ltd",
    "Leeds Building Society","Manchester Building Society","MARSHALL WACE LLP","Monzo Bank","Nationwide Building Society",
    "Nordea Bank Abp","Nottingham Building Society","OneSavings Bank Plc","Oxbury Bank Plc","Paragon Bank Plc","Pictet",
    "Rabobank","Sainsbury's Bank Plc","Secure Trust Bank Plc","Shawbrook Bank Ltd","Skipton Building Society",
    "SMBC Group EMEA","Starling Bank","Tandem Bank Ltd","Tesco Bank","Tescos Bank","The Co-operative Bank plc",
    "Triodos Bank UK Ltd","Ulster Bank","Union Bank UK Plc","Vanquis Bank Ltd","Virgin Money UK PLC",
    "WEST BROMWICH BUILDING SOCIETY","Yorkshire Bank","Yorkshire Building Society","Zopa Bank Ltd",
]

def split_variants(items):
    out = []
    for it in items:
        parts = [p.strip() for p in it.replace("–","/").replace("-", "/").split("/") if p.strip()]
        out.extend(parts)
    # de-dup while preserving order
    seen = set()
    dedup = []
    for p in out:
        if p.lower() not in seen:
            dedup.append(p)
            seen.add(p.lower())
    return dedup

TARGET_TITLES = split_variants(RAW_TITLES)
TARGET_COMPANIES = BANKS

# ---------- Helpers ----------
def norm_txt(x: str) -> str:
    if pd.isna(x):
        return ""
    x = str(x)
    x = unicodedata.normalize("NFKC", x)
    return " ".join(x.strip().lower().split())

def best_fuzzy_match(s: str, choices, scorer=fuzz.token_set_ratio):
    if not s:
        return ("", 0)
    match = process.extractOne(
        s, choices, scorer=scorer, score_cutoff=0
    )
    if match is None:
        return ("", 0)
    name, score, _ = match
    return (name, int(score))

# ---------- UI ----------
st.title("Contact Filter: Titles and Banks")

uploaded = st.file_uploader("Upload CSV with contacts", type=["csv"])
with st.expander("Matching options", expanded=True):
    colA, colB, colC = st.columns([1,1,1])
    with colA:
        title_threshold = st.slider("Job title match threshold", min_value=50, max_value=100, value=86, step=1)
    with colB:
        company_threshold = st.slider("Company match threshold", min_value=50, max_value=100, value=90, step=1)
    with colC:
        show_preview_rows = st.number_input("Preview rows", min_value=5, max_value=50, value=15, step=1)

if uploaded:
    # Read with simple fallback encodings
    read_ok = False
    for enc in ["utf-8-sig", "utf-8", "latin-1"]:
        try:
            df = pd.read_csv(uploaded, encoding=enc)
            read_ok = True
            break
        except Exception:
            continue
    if not read_ok:
        st.error("Could not read the file. Try saving as UTF-8 CSV.")
        st.stop()

    st.success(f"Loaded {len(df):,} rows, {len(df.columns)} columns")

    # Column mapping
    st.subheader("Map your columns")
    c1, c2 = st.columns(2)
    with c1:
        job_col = st.selectbox("Select the Job Title column", options=df.columns.tolist())
    with c2:
        company_col = st.selectbox("Select the Company column", options=df.columns.tolist())

    # Prepare normalized working columns
    work = df.copy()
    work["_job_norm"] = work[job_col].apply(norm_txt)
    work["_co_norm"] = work[company_col].apply(norm_txt)

    # Compute fuzzy scores
    st.info("Scoring titles and companies. This runs locally.")
    work[["_best_title", "_title_score"]] = work["_job_norm"].apply(
        lambda s: pd.Series(best_fuzzy_match(s, [norm_txt(x) for x in TARGET_TITLES], scorer=fuzz.token_set_ratio))
    )
    work[["_best_company", "_company_score"]] = work["_co_norm"].apply(
        lambda s: pd.Series(best_fuzzy_match(s, [norm_txt(x) for x in TARGET_COMPANIES], scorer=fuzz.token_set_ratio))
    )

    work["_match_title"] = work["_title_score"] >= title_threshold
    work["_match_company"] = work["_company_score"] >= company_threshold
    matched_titles = work[work["_match_title"]].copy()
    matched_both = work[work["_match_title"] & work["_match_company"]].copy()

    # Metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Total rows", f"{len(work):,}")
    m2.metric("Title matches", f"{len(matched_titles):,}")
    m3.metric("Title + Company matches", f"{len(matched_both):,}")

    # Optional previews
    st.markdown("### Preview: Title matches")
    st.dataframe(matched_titles.head(int(show_preview_rows)))
    st.markdown("### Preview: Title + Company matches")
    st.dataframe(matched_both.head(int(show_preview_rows)))

    # Downloads
    def df_to_csv_bytes(dfx: pd.DataFrame) -> bytes:
        buf = io.StringIO()
        # Drop helper columns
        out = dfx.drop(columns=[c for c in dfx.columns if c.startswith("_")], errors="ignore")
        out.to_csv(buf, index=False)
        return buf.getvalue().encode("utf-8")

    colD, colE = st.columns(2)
    with colD:
        st.download_button(
            label=f"Download Title Matches CSV ({len(matched_titles):,})",
            data=df_to_csv_bytes(matched_titles),
            file_name="title_matches.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with colE:
        st.download_button(
            label=f"Download Title+Company Matches CSV ({len(matched_both):,})",
            data=df_to_csv_bytes(matched_both),
            file_name="title_and_company_matches.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Diagnostics
    with st.expander("Diagnostics and tuning"):
        st.write("Top unmatched titles by frequency")
        unmatched_titles = work.loc[~work["_match_title"], job_col].fillna("(blank)")
        st.dataframe(unmatched_titles.value_counts().head(20))
        st.write("Top unmatched companies by frequency")
        unmatched_companies = work.loc[~work["_match_company"], company_col].fillna("(blank)")
        st.dataframe(unmatched_companies.value_counts().head(20))
else:
    st.caption("Upload a CSV to begin.")

# Notes:
# - Adjust thresholds if you need stricter or looser matches.
# - token_set_ratio handles partial and re-ordered strings well.
# - You can add more target titles or banks by editing RAW_TITLES / BANKS above.
