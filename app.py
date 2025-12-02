import io
import math
import re
import zipfile
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Streamlit - Expandi Splitter
# Purpose: Upload a contacts file, map columns to a standard schema,
# validate required fields, split evenly into N files, force-assign
# an Expandi Team Member per file, and download CSVs named to spec.
# ------------------------------------------------------------

st.set_page_config(page_title="Expandi Splitter", layout="wide")
st.title("Expandi List Splitter")
st.caption("Upload, map, validate, split, and download your Expandi-ready CSVs.")

# Target output schema in the exact order requested
TARGET_COLUMNS: List[str] = [
    "Expandi Team Member",
    "firstName",
    "lastName",
    "companyName",
    "title",
    "location",
    "linkedInProfileUrl",
    "Seniority",
    "State",
    "City",
    "Country Dropdown",
    "Target Event Role",
]

REQUIRED_FIELDS = {"linkedInProfileUrl", "firstName", "companyName"}

# -------------------------------
# Helpers
# -------------------------------

def read_uploaded_file(uploaded) -> pd.DataFrame:
    """Read CSV or Excel into a DataFrame, keeping original headers."""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded)
        except UnicodeDecodeError:
            uploaded.seek(0)
            return pd.read_csv(uploaded, encoding="latin-1")
    if name.endswith(".xlsx") or name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Excel read error: {e}")
            raise
    st.error("Unsupported file type. Please upload a CSV or Excel file.")
    st.stop()


def normalise_header(h: str) -> str:
    return re.sub(r"\s+", " ", str(h).strip()).lower()


def suggest_mapping(src_cols: List[str]) -> Dict[str, Optional[str]]:
    """Suggest a mapping from target column to a source column by fuzzy rules."""
    norm_src = {normalise_header(c): c for c in src_cols}

    aliases = {
        "expandi team member": ["expandi team member", "owner", "assignee", "agent", "team member", "expandi_owner"],
        "firstname": ["first name", "firstname", "given name", "first_name", "f_name"],
        "lastname": ["last name", "lastname", "surname", "last_name", "l_name"],
        "companyname": ["company", "company name", "organisation", "organization", "employer", "companyname"],
        "title": ["title", "job title", "role", "position"],
        "location": ["location", "city - state", "city, state", "geo", "region"],
        "linkedinprofileurl": [
            "linkedin", "linkedin url", "linkedin profile", "linkedinprofileurl", "li url", "profile url",
            "linkedin_profile", "linkedinprofile"
        ],
        "seniority": ["seniority", "level", "seniority level"],
        "state": ["state", "province", "state/region"],
        "city": ["city", "town"],
        "country dropdown": ["country", "country dropdown", "nation"],
        "target event role": ["target event role", "target role", "event role", "role target"],
    }

    mapping: Dict[str, Optional[str]] = {col: None for col in TARGET_COLUMNS}

    for target in TARGET_COLUMNS:
        key = normalise_header(target)
        if key in aliases:
            for alias in aliases[key]:
                n = normalise_header(alias)
                if n in norm_src:
                    mapping[target] = norm_src[n]
                    break
        else:
            # default exact match by normalised header
            if key in norm_src:
                mapping[target] = norm_src[key]

    return mapping


def build_output_frame(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    data = {}
    for col in TARGET_COLUMNS:
        src = mapping.get(col)
        if src is None or src == "-- Not in file --":
            data[col] = [""] * len(df)
        else:
            # Cast to str to avoid NaN and keep uniform output
            series = df[src].astype(str).fillna("").map(lambda x: x.strip())
            data[col] = series
    out = pd.DataFrame(data, columns=TARGET_COLUMNS)
    return out


def validate_required(out_df: pd.DataFrame) -> List[str]:
    problems = []
    for req in REQUIRED_FIELDS:
        if req not in out_df.columns:
            problems.append(f"Missing required column in output: {req}")
            continue
        missing = (out_df[req].astype(str).str.strip() == "").sum()
        if missing > 0:
            problems.append(f"{req} has {missing} empty values after mapping")
    return problems


def contiguous_split(df: pd.DataFrame, n_parts: int) -> List[pd.DataFrame]:
    """Split into n contiguous parts as evenly as possible."""
    if n_parts <= 1:
        return [df]
    parts = np.array_split(df, n_parts)
    return [p.copy() for p in parts]


def safe_filename(name: str) -> str:
    # Remove illegal filename characters and condense spaces
    name = re.sub(r"[\\/:*?\"<>|]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def build_zip(files: List[Dict[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.writestr(f["name"], f["data"])
    return buf.getvalue()


# -------------------------------
# Sidebar - Controls
# -------------------------------
with st.sidebar:
    st.header("Settings")
    shuffle_before_split = st.checkbox("Shuffle rows before splitting", value=False, help="Optional - prevents bias from input ordering.")
    drop_duplicates = st.checkbox(
        "De-duplicate by LinkedIn URL", value=True,
        help="Keeps the first row where linkedInProfileUrl matches, drops the rest."
    )

# -------------------------------
# Main - Upload
# -------------------------------
uploaded = st.file_uploader("Upload a CSV or Excel file", type=["csv", "xlsx", "xls"], accept_multiple_files=False)

if uploaded is None:
    st.info("Upload a file to begin.")
    st.stop()

raw_df = read_uploaded_file(uploaded)
if raw_df.empty:
    st.error("Your file appears to be empty.")
    st.stop()

st.subheader("Step 1 - Map your columns")

src_columns = list(raw_df.columns)
if "_saved_mapping" not in st.session_state:
    st.session_state._saved_mapping = suggest_mapping(src_columns)

# Mapping UI
mapping: Dict[str, Optional[str]] = {}
cols_left, cols_right = st.columns(2)

for i, tgt in enumerate(TARGET_COLUMNS):
    container = cols_left if i % 2 == 0 else cols_right
    with container:
        options = ["-- Not in file --"] + src_columns
        default = 0
        suggested = st.session_state._saved_mapping.get(tgt)
        if suggested and suggested in src_columns:
            default = options.index(suggested) if suggested in options else 0
        selected = st.selectbox(
            f"Map to: {tgt}",
            options=options,
            index=default,
            key=f"map_{tgt}",
        )
        mapping[tgt] = None if selected == "-- Not in file --" else selected

st.caption("Only required fields are linkedInProfileUrl, firstName, companyName.")

# Build output preview
out_df = build_output_frame(raw_df, mapping)

# Validate required
problems = validate_required(out_df)
if problems:
    st.error("Required field issues detected. Fix your mapping or input data.")
    for p in problems:
        st.write("- " + p)
    st.stop()

# Optional de-duplicate and shuffle
if drop_duplicates:
    before = len(out_df)
    out_df = out_df.drop_duplicates(subset=["linkedInProfileUrl"], keep="first").reset_index(drop=True)
    after = len(out_df)
    if after < before:
        st.success(f"De-duplicated by LinkedIn URL - removed {before - after} rows. Kept {after}.")

if shuffle_before_split:
    out_df = out_df.sample(frac=1.0, random_state=42).reset_index(drop=True)

st.subheader("Preview after mapping")
st.dataframe(out_df.head(25), use_container_width=True)
st.write(f"Total rows ready: **{len(out_df)}**")

# -------------------------------
# Step 2 - Split settings
# -------------------------------
st.subheader("Step 2 - Split the file")

num_files = st.number_input("Number of people to split this list across", min_value=1, max_value=200, value=1, step=1)
st.caption("We will split the rows contiguously into this many files.")

# Team member names input
st.markdown("**Team member names**")
default_names = "\n".join([f"Member {i+1}" for i in range(int(num_files))])
name_lines = st.text_area(
    "Enter one name per line. If you enter one name and choose more than one file we will reuse that name. If you provide fewer names than files we will cycle them.",
    value=default_names,
    height=120,
)

team_names = [n.strip() for n in name_lines.splitlines() if n.strip()]
if not team_names:
    team_names = ["Unassigned"]
# Fit to num_files by cycling or truncating
if len(team_names) < num_files:
    k = math.ceil(num_files / len(team_names))
    team_names = (team_names * k)[: num_files]
elif len(team_names) > num_files:
    team_names = team_names[: num_files]

base_name = st.text_input("Base name for files", value="Visprom Contacts")

# -------------------------------
# Step 3 - Build and download
# -------------------------------
st.subheader("Step 3 - Generate your CSVs")

if st.button("Prepare and show downloads", type="primary"):
    parts = contiguous_split(out_df, int(num_files))

    files_payload: List[Dict[str, bytes]] = []

    st.success(f"Prepared {len(parts)} file(s). See downloads below.")

    for i, part in enumerate(parts):
        member = team_names[i] if i < len(team_names) else "Unassigned"
        # Force-assign team member - overwrite any mapped value
        part["Expandi Team Member"] = member

        count = len(part)
        filename_display = f"{base_name} - {member} - {count}.csv"
        filename_safe = safe_filename(filename_display)

        # Ensure output column order and header presence
        part = part[TARGET_COLUMNS]
        data = part.to_csv(index=False).encode("utf-8-sig")  # include BOM for Excel-friendliness
        files_payload.append({"name": filename_safe, "data": data})

        st.download_button(
            label=f"Download {filename_display}",
            data=data,
            file_name=filename_safe,
            mime="text/csv",
            key=f"dl_{i}_{filename_safe}",
        )

    # Also offer ZIP
    zip_name = safe_filename(f"{base_name} - split-{len(parts)}.zip")
    zip_bytes = build_zip(files_payload)
    st.download_button(
        label=f"Download all as ZIP ({zip_name})",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
        key="dl_zip_all",
    )

    # Summary table
    st.markdown("### Summary")
    summary_rows = []
    start = 0
    for i, part in enumerate(parts):
        end = start + len(part)
        summary_rows.append({
            "File #": i + 1,
            "Team Member": team_names[i] if i < len(team_names) else "Unassigned",
            "Rows": len(part),
            "Row range": f"{start + 1} - {end}",
        })
        start = end
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

else:
    st.info("When you are happy with the mapping and settings, click the button above to generate your downloads.")
