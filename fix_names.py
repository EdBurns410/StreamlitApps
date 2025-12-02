import re
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Name Fixer from Email", layout="wide")

st.title("Name Cleaner: Use email to fix firstname / lastname")

st.write(
    "Upload the Excel file with your 4 sheets. "
    "This will try to fix jumbled first/last names using the email pattern and "
    "output a new workbook with corrected columns."
)

uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx"])


# ---------------------- Helper functions ---------------------- #

def normalise_cols(df: pd.DataFrame):
    """
    Return df and a mapping of logical names to actual column names.
    Looks for email / firstname / lastname in a case-insensitive way.
    """
    col_map = {}
    lower_map = {c.lower(): c for c in df.columns}

    for logical, target in {
        "email": "email",
        "firstname": "firstname",
        "lastname": "lastname",
    }.items():
        for col in df.columns:
            if col.strip().lower() == target:
                col_map[logical] = col
                break

    # If something is missing, just skip those rows for fixing
    return df, col_map


def get_email_tokens(email: str):
    """
    Extract name-like tokens from email local part (before '@'),
    split on dots/underscores, strip non-letters.
    """
    if not isinstance(email, str) or "@" not in email:
        return []

    local = email.split("@", 1)[0].lower()
    raw_tokens = re.split(r"[._]", local)
    tokens = []
    for t in raw_tokens:
        t = re.sub(r"[^a-z]", "", t)  # keep only letters
        if t:
            tokens.append(t)
    return tokens


def core_from_firstname(name: str) -> str:
    """
    Get a 'core' token from firstname cell for matching.
    Typically first word, stripped of punctuation.
    """
    if not isinstance(name, str):
        return ""
    # Take first token
    token = name.strip().split()[0]
    token = token.strip(",.").lower()
    token = re.sub(r"[^a-z]", "", token)
    return token


def core_from_lastname(name: str) -> str:
    """
    Get a 'core' token from lastname cell for matching.
    Use text before first comma if present, else first word.
    """
    if not isinstance(name, str):
        return ""
    s = name.strip()
    if "," in s:
        s = s.split(",", 1)[0]
    token = s.split()[0]
    token = token.strip(",.").lower()
    token = re.sub(r"[^a-z]", "", token)
    return token


def nice_first(name: str) -> str:
    """
    Return a cleaned firstname value to write out.
    Soft clean: trim and drop trailing comma.
    """
    if not isinstance(name, str):
        return ""
    s = name.strip()
    if "," in s:
        s = s.split(",", 1)[0]
    return s


def nice_last(name: str) -> str:
    """
    Return a cleaned lastname value to write out.
    Drop everything after first comma (credentials, etc.).
    """
    if not isinstance(name, str):
        return ""
    s = name.strip()
    if "," in s:
        s = s.split(",", 1)[0]
    return s


def fix_name_row(email, fname, lname):
    """
    Decide corrected first/last for a single row using the email.
    Returns (first_fixed, last_fixed, rule_used).
    """
    # Defaults: just cleaned versions of existing
    first_clean = nice_first(fname)
    last_clean = nice_last(lname)
    rule = "original_cleaned"

    tokens = get_email_tokens(email)
    if len(tokens) < 2:
        # Not enough information in email to be confident
        return first_clean, last_clean, rule

    ef = tokens[0]      # email first
    el = tokens[-1]     # email last

    cf = core_from_firstname(fname)
    cl = core_from_lastname(lname)

    # Case 1: firstname/lastname match email orientation
    if cf == ef and cl == el:
        # Already correct, but we still clean off credentials from last
        rule = "matched_email_orientation"
        return first_clean, last_clean, rule

    # Case 2: swapped (firstname matches email last, lastname matches email first)
    if cf == el and cl == ef:
        # Swap them
        new_first = nice_first(lname)
        new_last = nice_last(fname)
        rule = "swapped_based_on_email"
        return new_first, new_last, rule

    # Case 3: firstname matches email first, lastname seems messy or mismatched
    if cf == ef and cl != el:
        # Trust email for surname, but keep existing case if possible
        # If current lastname core contains el somewhere, just clean
        if el in core_from_lastname(lname):
            rule = "lastname_trimmed_to_core"
            return first_clean, last_clean, rule
        else:
            # Overwrite with email-derived last, capitalised naive
            rule = "lastname_overwritten_from_email"
            new_last = el.capitalize()
            return first_clean, new_last, rule

    # Case 4: lastname matches email last, firstname seems messy/mismatched
    if cl == el and cf != ef:
        # Similar logic, but for firstname
        if ef in core_from_firstname(fname):
            rule = "firstname_trimmed_to_core"
            return first_clean, last_clean, rule
        else:
            rule = "firstname_overwritten_from_email"
            new_first = ef.capitalize()
            return new_first, last_clean, rule

    # If we get here, it's ambiguous; keep cleaned original
    return first_clean, last_clean, rule


def process_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = df.copy()
    df, col_map = normalise_cols(df)

    email_col = col_map.get("email")
    first_col = col_map.get("firstname")
    last_col = col_map.get("lastname")

    if not (email_col and first_col and last_col):
        # If required columns missing, just return unchanged
        df["firstname_fixed"] = df.get(first_col, "")
        df["lastname_fixed"] = df.get(last_col, "")
        df["name_fix_rule"] = "missing_required_columns"
        return df

    first_fixed = []
    last_fixed = []
    rules = []

    for _, row in df.iterrows():
        email = row.get(email_col, "")
        fname = row.get(first_col, "")
        lname = row.get(last_col, "")

        f_fix, l_fix, rule = fix_name_row(email, fname, lname)
        first_fixed.append(f_fix)
        last_fixed.append(l_fix)
        rules.append(rule)

    df["firstname_fixed"] = first_fixed
    df["lastname_fixed"] = last_fixed
    df["name_fix_rule"] = rules

    return df


def create_excel_download(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])  # Excel sheet name limit
    output.seek(0)
    return output


# ---------------------- Main app logic ---------------------- #

if uploaded_file is not None:
    # Read all sheets
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    st.success(f"Loaded {len(all_sheets)} sheet(s): {', '.join(all_sheets.keys())}")

    processed_sheets = {}
    for sheet_name, df in all_sheets.items():
        st.subheader(f"Sheet: {sheet_name}")

        processed_df = process_sheet(df, sheet_name)
        processed_sheets[sheet_name] = processed_df

        # Show a preview
        st.write("Preview with fixed names (top 10 rows):")
        st.dataframe(processed_df.head(10))

        # Quick stats on how many rows got touched
        if "name_fix_rule" in processed_df.columns:
            rule_counts = processed_df["name_fix_rule"].value_counts()
            st.write("Name fix rules applied:")
            st.write(rule_counts)

    # Single download for the whole workbook – unique key to avoid clashes
    download_buffer = create_excel_download(processed_sheets)
    st.download_button(
        label="Download corrected Excel workbook",
        data=download_buffer,
        file_name="names_fixed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_corrected_workbook_btn",
    )
else:
    st.info("Upload your .xlsx file to start fixing names.")
