# findvpother.py (replace existing file)
import streamlit as st
import pandas as pd
from io import BytesIO
from typing import List, Dict, Optional

# Optional: rapidfuzz for better fuzzy matching
try:
    from rapidfuzz import process, fuzz
    _USE_RAPIDFUZZ = True
except Exception:
    import difflib
    _USE_RAPIDFUZZ = False

st.set_page_config(page_title="State Splitter", layout="wide")

TARGET_HEADERS = [
    "email",
    "firstname",
    "lastname",
    "jobtitle",
    "Company Name",
    "country",
    "State",
    "Organisation Type",
    "Segment",
    "Seniority",
    "Bucket",
]

STATE_TARGETS = ["Virginia", "Tennessee", "Georgia", "South Carolina"]
STATE_ABBREVS = {
    "Virginia": ["va"],
    "Tennessee": ["tn"],
    "Georgia": ["ga"],
    "South Carolina": ["sc"],
}


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file, engine="openpyxl")


def best_match(column: str, candidates: List[str]) -> Optional[str]:
    if not candidates:
        return None
    if _USE_RAPIDFUZZ:
        match = process.extractOne(column, candidates, scorer=fuzz.token_sort_ratio)
        if match:
            candidate, score, _ = match
            return candidate if score >= 50 else None
        return None
    else:
        matches = difflib.get_close_matches(column, candidates, n=1, cutoff=0.4)
        return matches[0] if matches else None


def map_headers(uploaded_cols: List[str]) -> Dict[str, Optional[str]]:
    mapping = {}
    lowered = [c.lower() for c in uploaded_cols]
    for target in TARGET_HEADERS:
        # exact (case-insensitive) match first
        try:
            idx = lowered.index(target.lower())
            mapping[target] = uploaded_cols[idx]
            continue
        except ValueError:
            pass
        candidate = best_match(target, uploaded_cols)
        mapping[target] = candidate
    return mapping


def state_matches(value: str, state: str, threshold: int = 80) -> bool:
    if pd.isna(value):
        return False
    s = str(value).strip().lower()
    target = state.lower()
    if s == target or target in s:
        return True
    for abbr in STATE_ABBREVS.get(state, []):
        if s == abbr or s == abbr + "." or s.upper() == abbr.upper():
            return True
    if _USE_RAPIDFUZZ:
        score = fuzz.token_sort_ratio(s, target)
        return score >= threshold
    else:
        ratio = __import__("difflib").SequenceMatcher(None, s, target).ratio()
        return int(ratio * 100) >= threshold


def to_excel_bytes(sheets: Dict[str, pd.DataFrame], all_columns: List[str]) -> bytes:
    output = BytesIO()
    engine = "xlsxwriter"
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet_name, sheet_df in sheets.items():
            if sheet_df is None or sheet_df.empty:
                pd.DataFrame(columns=all_columns).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        # context manager will save/close
    return output.getvalue()


st.title("State Splitter: Upload and split contacts by state")
st.write(
    "Upload a CSV or Excel file. The app will attempt to map your header row to the expected columns, then extract rows that look like they are from the chosen states and export an Excel with one sheet per state."
)

uploaded_file = st.file_uploader("Upload .csv or .xlsx file", type=["csv", "xlsx"], accept_multiple_files=False)
if uploaded_file is None:
    st.info("Please upload a CSV or XLSX file to begin.")
    st.stop()

# Read file
try:
    df = read_uploaded_file(uploaded_file)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

st.success(f"Loaded file: {uploaded_file.name} with {len(df)} rows and {len(df.columns)} columns")

uploaded_cols = list(df.columns.astype(str))
auto_map = map_headers(uploaded_cols)

# Initialize session state keys used by this app
if "mapping_applied" not in st.session_state:
    st.session_state.mapping_applied = False
if "user_map" not in st.session_state:
    st.session_state.user_map = {}
if "results" not in st.session_state:
    st.session_state.results = {}
if "download_bytes" not in st.session_state:
    st.session_state.download_bytes = None

st.subheader("Header mapping")
st.write("The app has tried to map your file's headers to the expected columns. Adjust any mapping if required.")

# Show mapping form. On submit, store mapping in session_state so it persists across reruns.
with st.form(key="mapping_form"):
    user_map_local = {}
    cols_with_none = ["Not present"] + uploaded_cols
    for target in TARGET_HEADERS:
        default = auto_map.get(target)
        prev = st.session_state.user_map.get(target)
        default_choice = prev if prev in cols_with_none else default
        index = cols_with_none.index(default_choice) if default_choice in cols_with_none else 0
        # explicit key so selectbox keeps its state
        selection = st.selectbox(f"{target}", options=cols_with_none, index=index, key=f"map_{target}")
        user_map_local[target] = None if selection == "Not present" else selection

    st.caption("If a target column is not present in your file, choose Not present. The filtering will use the mapped 'State' column.")
    submit = st.form_submit_button("Apply mapping")
    if submit:
        # persist mapping and mark as applied
        st.session_state.user_map = user_map_local
        st.session_state.mapping_applied = True
        # do NOT call experimental_rerun here; just continue within the same run

# If mapping has not been applied yet, stop here so user can apply it
if not st.session_state.mapping_applied:
    st.info("Review the header mapping and click Apply mapping to continue.")
    st.stop()

# From here onwards, mapping exists in session_state.user_map
mapped_cols = {k: v for k, v in st.session_state.user_map.items() if v}
if "State" not in mapped_cols:
    st.error("You must map a column to 'State' for the state filtering to work.")
    st.stop()

# Rename columns in a stable way
rename_map = {v: k for k, v in mapped_cols.items()}
df_renamed = df.rename(columns=rename_map).copy()

if "State" not in df_renamed.columns:
    st.error("State column not found after renaming. Check your mapping.")
    st.stop()

# Fuzzy threshold
threshold = st.slider("Fuzzy match threshold (higher = stricter)", min_value=50, max_value=100, value=85)

# Compute matches and store in session_state.results so the information remains available across reruns
results = {}
for state in STATE_TARGETS:
    mask = df_renamed["State"].apply(lambda x: state_matches(x, state, threshold=threshold))
    matched_df = df_renamed[mask].copy()
    results[state] = matched_df

st.session_state.results = results

# Display counts
st.subheader("Results")
counts = {state: len(df_state) for state, df_state in results.items()}
cols = st.columns(len(STATE_TARGETS))
for idx, state in enumerate(STATE_TARGETS):
    cols[idx].metric(label=state, value=counts[state])

# Preview
with st.expander("Preview matched rows"):
    for state, df_state in results.items():
        st.write(f"### {state} — {len(df_state)} rows")
        if len(df_state) > 0:
            st.dataframe(df_state.head(100))
        else:
            st.write("No matches found for this state with the current threshold.")

# Download generation and button (always show when bytes exist, and also show immediately on click)
# Download generation and button (always show when bytes exist, and also show immediately on click)
placeholder = st.empty()

def render_download(bytes_data, key_suffix=""):
    """Display the download button for generated bytes."""
    if bytes_data:
        placeholder.download_button(
            label="Download Excel with 4 sheets",
            data=bytes_data,
            file_name="state_matches.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_xlsx_btn_{key_suffix}",
            use_container_width=True,
        )

# If we already have bytes from a prior click, render the button now
render_download(st.session_state.get("download_bytes"), key_suffix="existing")

# Create on demand and show the button immediately
if st.button("Create download file (.xlsx)"):
    try:
        file_bytes = to_excel_bytes(results, all_columns=list(df_renamed.columns))
        st.session_state.download_bytes = file_bytes
        st.success("Download ready. The file contains one sheet per target state.")
        # render the button right away in the same run, with a different key
        render_download(file_bytes, key_suffix="new")
        st.caption(f"File size: {len(file_bytes)} bytes")
    except Exception as e:
        st.error(f"Failed to create Excel file: {e}")
