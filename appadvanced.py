import io
import math
import re
import zipfile
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Streamlit - Expandi Splitter with Connection-Owner Allocation
# ------------------------------------------------------------

st.set_page_config(page_title="Expandi Splitter", layout="wide")
st.title("Expandi List Splitter")
st.caption("Upload, map, validate, split, and download your Expandi-ready CSVs with connection-owner allocation.")

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

# Required fields. ProfileUrl must be mapped via linkedInProfileUrl.
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


ALIASES = {
    "expandi team member": ["expandi team member", "owner", "assignee", "agent", "team member", "expandi_owner"],
    "firstname": ["first name", "firstname", "given name", "first_name", "f_name"],
    "lastname": ["last name", "lastname", "surname", "last_name", "l_name"],
    "companyname": ["company", "company name", "organisation", "organization", "employer", "companyname"],
    "title": ["title", "job title", "role", "position"],
    "location": ["location", "city - state", "city, state", "geo", "region"],
    "linkedinprofileurl": [
        "linkedin", "linkedin url", "linkedin profile", "linkedinprofileurl", "li url", "profile url",
        "linkedin_profile", "linkedinprofile", "profileurl", "public profile url"
    ],
    "seniority": ["seniority", "level", "seniority level"],
    "state": ["state", "province", "state/region"],
    "city": ["city", "town"],
    "country dropdown": ["country", "country dropdown", "nation"],
    "target event role": ["target event role", "target role", "event role", "role target"],
}

def suggest_mapping(src_cols: List[str]) -> Dict[str, Optional[str]]:
    """Suggest a mapping from target column to a source column by fuzzy rules."""
    norm_src = {normalise_header(c): c for c in src_cols}
    mapping: Dict[str, Optional[str]] = {col: None for col in TARGET_COLUMNS}
    for target in TARGET_COLUMNS:
        key = normalise_header(target)
        # direct hit
        if key in norm_src:
            mapping[target] = norm_src[key]
            continue
        # alias hit
        if key in ALIASES:
            for alias in ALIASES[key]:
                n = normalise_header(alias)
                if n in norm_src:
                    mapping[target] = norm_src[n]
                    break
    return mapping


def build_output_frame(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    data = {}
    for col in TARGET_COLUMNS:
        src = mapping.get(col)
        if src is None or src == "-- Not in file --":
            data[col] = [""] * len(df)
        else:
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
    # LinkedIn URL sanity check
    if "linkedInProfileUrl" in out_df.columns:
        bad_like = out_df["linkedInProfileUrl"].astype(str).str.contains(r"^https?://", case=False, na=False) == False
        if bad_like.any():
            problems.append("Some linkedInProfileUrl values do not look like URLs")
    return problems


def contiguous_split(df: pd.DataFrame, n_parts: int) -> List[pd.DataFrame]:
    if n_parts <= 1:
        return [df]
    parts = np.array_split(df, n_parts)
    return [p.copy() for p in parts]


def safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def build_zip(files: List[Dict[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.writestr(f["name"], f["data"])
    return buf.getvalue()


def detect_url_column(df: pd.DataFrame) -> Optional[str]:
    """Find the LinkedIn URL column in a connections file, prioritising 'profileUrl'."""
    # 1) Exact 'profileUrl'
    if "profileUrl" in df.columns:
        return "profileUrl"
    # 2) Case-insensitive match
    for c in df.columns:
        if normalise_header(c) == "profileurl":
            return c
    # 3) Fallbacks: common variants
    for c in df.columns:
        if normalise_header(c) in {"linkedinprofileurl", "public profile url", "linkedin url", "linkedin_profile"}:
            return c
    # 4) Heuristic: any column that looks mostly like LinkedIn URLs
    for c in df.columns:
        vals = df[c].astype(str)
        if vals.str.contains("linkedin.com/", case=False, na=False).mean() > 0.3:
            return c
    return None



def normalize_url(u: str) -> str:
    u = str(u or "").strip()
    # lower for compare, drop URL params and trailing slash
    u = u.split("?")[0].rstrip("/")
    return u.lower()


def build_member_url_sets(member_files: List[Optional[Tuple[str, pd.DataFrame]]]) -> List[Set[str]]:
    """For each member slot, return a set of normalized LinkedIn URLs from their uploaded file."""
    url_sets: List[Set[str]] = []
    for item in member_files:
        if item is None:
            url_sets.append(set())
            continue
        name, df = item
        col = detect_url_column(df)
        if not col:
            st.warning(f"Could not find a LinkedIn URL column in {name}. Skipping its matches.")
            url_sets.append(set())
            continue
        urls = df[col].astype(str).map(normalize_url)
        url_sets.append(set(u for u in urls if "linkedin.com/" in u))
    return url_sets


def allocate_with_connection_owners(df: pd.DataFrame, member_names: List[str], owner_sets: List[Set[str]]) -> List[pd.DataFrame]:
    """Assign rows to the member who already has the connection. Remaining rows are balanced evenly."""
    n = len(member_names)
    buckets: List[List[int]] = [[] for _ in range(n)]
    url_series = df["linkedInProfileUrl"].astype(str).map(normalize_url)

    # First pass: owner assignment
    for idx, url in url_series.items():
        assigned = False
        if "linkedin.com/" in url:
            for m_idx in range(n):
                if url in owner_sets[m_idx]:
                    buckets[m_idx].append(idx)
                    assigned = True
                    break
        if not assigned:
            # mark as unassigned by using -1 in a temp list
            pass

    assigned_indices = set(i for b in buckets for i in b)
    remaining = [i for i in df.index.tolist() if i not in assigned_indices]

    # Second pass: balance remaining across members to near-equal sizes
    target_sizes = [len(b) for b in buckets]
    base = len(df) // n
    extra = len(df) % n
    # desired final sizes
    desired = [base + (1 if i < extra else 0) for i in range(n)]

    # fill greedily into members with most headroom
    ptr = 0
    for row_idx in remaining:
        # choose member with max (desired - current), tie break by lowest index
        headroom = [desired[i] - target_sizes[i] for i in range(n)]
        m_idx = int(np.argmax(headroom))
        buckets[m_idx].append(row_idx)
        target_sizes[m_idx] += 1
        ptr += 1

    # Build dataframes per member preserving original order within each bucket
    parts: List[pd.DataFrame] = []
    for b in buckets:
        part = df.loc[b].copy()
        parts.append(part.reset_index(drop=True))
    return parts


# -------------------------------
# Sidebar - Controls
# -------------------------------
with st.sidebar:
    st.header("Settings")
    shuffle_before_split = st.checkbox("Shuffle rows before splitting", value=False, help="Optional to prevent input-order bias.")
    drop_duplicates = st.checkbox("De-duplicate by LinkedIn URL", value=True, help="Keeps first row per linkedInProfileUrl.")
    enforce_owner_allocation = st.checkbox("Prioritise connection-owner allocation", value=True, help="Assign rows to members who already have that LinkedIn connection.")

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

st.caption("Required fields: linkedInProfileUrl, firstName, companyName. ProfileUrl must be mapped to linkedInProfileUrl.")

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
st.caption("We will split rows into this many files. Connection-owner allocation runs first if enabled.")

# Team member names input
st.markdown("**Team member names**")
default_names = "\n".join([f"Member {i+1}" for i in range(int(num_files))])
name_lines = st.text_area(
    "Enter one name per line. If fewer than slots, names will be cycled. If more, extras are ignored.",
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

# -------------------------------
# Step 2a - Upload connections per member
# -------------------------------
st.markdown("**Optional: upload each member’s LinkedIn connections or sent requests file**")
st.caption("If provided, any contact whose linkedInProfileUrl matches will be auto-assigned to that member.")

member_files: List[Optional[Tuple[str, pd.DataFrame]]] = []
uploader_cols = st.columns(2)
for i in range(int(num_files)):
    with uploader_cols[i % 2]:
        mf = st.file_uploader(
            f"Connections file for {team_names[i]}",
            type=["csv", "xlsx", "xls"],
            key=f"member_file_{i}",
            accept_multiple_files=False
        )
    if mf is not None:
        try:
            df_mf = read_uploaded_file(mf)
            member_files.append((mf.name, df_mf))
        except Exception:
            member_files.append(None)
    else:
        member_files.append(None)

owner_sets = build_member_url_sets(member_files)

base_name = st.text_input("Base name for files", value="Visprom Contacts")

# -------------------------------
# Step 3 - Build and download
# -------------------------------
st.subheader("Step 3 - Generate your CSVs")

if st.button("Prepare and show downloads", type="primary"):
    # Allocation with owner priority then balancing
    if enforce_owner_allocation and len(team_names) > 1:
        parts = allocate_with_connection_owners(out_df, team_names, owner_sets)
    else:
        parts = contiguous_split(out_df, int(num_files))

    files_payload: List[Dict[str, bytes]] = []

    st.success(f"Prepared {len(parts)} file(s). See downloads below.")

    for i, part in enumerate(parts):
        member = team_names[i] if i < len(team_names) else "Unassigned"
        # Force-assign team member
        part["Expandi Team Member"] = member

        count = len(part)
        filename_display = f"{base_name} - {member} - {count}.csv"
        filename_safe = safe_filename(filename_display)

        # Ensure output column order
        part = part[TARGET_COLUMNS]
        data = part.to_csv(index=False).encode("utf-8-sig")
        files_payload.append({"name": filename_safe, "data": data})

        st.download_button(
            label=f"Download {filename_display}",
            data=data,
            file_name=filename_safe,
            mime="text/csv",
            key=f"dl_{i}_{filename_safe}",
        )

    # ZIP
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
    already_matched_counts = []
    # count matches per part for visibility
    for i, part in enumerate(parts):
        urls = part["linkedInProfileUrl"].astype(str).map(normalize_url)
        matched = len([u for u in urls if u in owner_sets[i]])
        already_matched_counts.append(matched)

    start = 0
    for i, part in enumerate(parts):
        end = start + len(part)
        summary_rows.append({
            "File #": i + 1,
            "Team Member": team_names[i] if i < len(team_names) else "Unassigned",
            "Rows": len(part),
            "Pre-matched connections": already_matched_counts[i],
            "Row range": f"{start + 1} - {end}",
        })
        start = end
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

else:
    st.info("When you are happy with the mapping and settings, click the button above to generate your downloads.")
