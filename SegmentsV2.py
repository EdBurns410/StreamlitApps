# segmenter.py
# Usage:
#   pip install streamlit pandas openpyxl xlsxwriter rapidfuzz
#   streamlit run segmenter.py

import io
import re
import zipfile
from typing import Dict, List

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

# ----------------------- Config -----------------------

TARGET_HEADERS = [
    "SalesNavUrl","fullName","firstName","lastName","companyName","title","companyId",
    "companyUrl","regularCompanyUrl","summary","titleDescription","industry",
    "companyLocation","location","durationInRole","durationInCompany",
    "pastExperienceCompanyName","pastExperienceCompanyUrl","pastExperienceCompanyTitle",
    "pastExperienceDate","pastExperienceDuration","connectionDegree","profileImageUrl",
    "sharedConnectionsCount","name","vmid","linkedInProfileUrl","isPremium","isOpenLink",
    "query","timestamp","ProfileUrl","searchAccountProfileId","searchAccountProfileName",
    "Bank Size","Seniority","Focus Area","Relevance",
]

# Target states plus a catch-all to guarantee coverage of every row
STATE_LABELS = {
    "North Carolina": {"names": ["North Carolina", "N Carolina"], "abbr": ["NC"]},
    "Virginia": {"names": ["Virginia"], "abbr": ["VA"]},
    "South Carolina": {"names": ["South Carolina", "S Carolina"], "abbr": ["SC"]},
    "Tennessee": {"names": ["Tennessee"], "abbr": ["TN"]},
    "Georgia": {"names": ["Georgia"], "abbr": ["GA"]},
    "Other States": {"names": [], "abbr": []},  # fallback
}

CREDIT_UNION_HINTS = [
    r"\bcredit union\b",
    r"\bfcu\b",
    r"\bfederal credit union\b",
    r"\bcu\b",
]
BANK_HINTS = [
    r"\bbank\b",
    r"\bbanco\b",
    r"\bbanca\b",
    r"\bbancorp\b",
    r"\btrust\b",
    r"\bnational bank\b",
]

FOCUS_BUCKETS = {
    "Data_AI_Automation": [
        "data","analytics","machine learning","ml","ai","artificial intelligence",
        "automation","rpa","robotic process","data science","scientist","data engineer",
        "ml engineer","mlops","genai","gen ai","llm","prompt","knowledge graph",
        "nlp","computer vision","model risk"
    ],
    "Transformation_Digital_Innovation": [
        "transformation","change","digital","innovation","innovate",
        "enterprise architect","architecture","architect","product","design",
        "experience","cx","ux","strategy","pm","programme","program manager",
        "agile","scrum","delivery","cto","cdo","cpo","head of digital"
    ],
    "Risk_Compliance": [
        "risk","compliance","fraud","conduct","audit","aml","sanctions",
        "financial crime","kyc","model validation","governance",
        "regulatory","prudential","sox","basel","grc","bsa"
    ],
}

STATE_FUZZ_THRESHOLD = 85

# ----------------------- Helpers -----------------------

def norm(s: str) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().lower()

def contains_any(text: str, patterns: List[str]) -> bool:
    t = norm(text)
    return any(p in t for p in patterns)

def regex_any(text: str, regex_list: List[str]) -> bool:
    t = str(text or "")
    return any(re.search(rx, t, flags=re.IGNORECASE) for rx in regex_list)

def classify_company_type(company_name: str) -> str:
    if not company_name or pd.isna(company_name):
        return "Bank"
    name = str(company_name)
    if regex_any(name, CREDIT_UNION_HINTS):
        return "Credit Union"
    if regex_any(name, BANK_HINTS):
        return "Bank"
    return "Bank"  # default to Bank for outreach bias

def classify_focus(title: str) -> str:
    t = norm(title)
    if not t:
        return "All Others"
    if contains_any(t, FOCUS_BUCKETS["Data_AI_Automation"]):
        return "Data / AI / Automation"
    if contains_any(t, FOCUS_BUCKETS["Risk_Compliance"]):
        return "Risk and Compliance"
    if contains_any(t, FOCUS_BUCKETS["Transformation_Digital_Innovation"]):
        return "Transformation, digital, innovation"
    return "All Others"

def best_state_match(location: str) -> str:
    if not location or pd.isna(location):
        return "Other States"
    text = str(location)
    candidates = []
    for state, meta in STATE_LABELS.items():
        if state == "Other States":
            continue
        for token in set(meta["names"] + meta["abbr"]):
            score = fuzz.partial_ratio(token.lower(), text.lower())
            candidates.append((state, token, score))
    if not candidates:
        return "Other States"
    state, _, score = max(candidates, key=lambda x: x[2])
    return state if score >= STATE_FUZZ_THRESHOLD else "Other States"

def auto_guess_mapping(in_cols: List[str], target: str) -> str:
    for c in in_cols:
        if norm(c) == norm(target):
            return c
    aliases = {
        "linkedInProfileUrl": ["linkedinprofileurl","linkedin url","profile url","profileurl","linkedin"],
        "ProfileUrl": ["profileurl","profile url","linkedin url","linkedinprofileurl"],
        "companyName": ["company","company_name","employer"],
        "title": ["job title","role","position","headline"],
        "fullName": ["fullname","name","display name"],
        "firstName": ["firstname","first name","given name"],
        "lastName": ["lastname","last name","surname","family name"],
        "location": ["location","city, state","city","state","region"],
        "companyLocation": ["company location","hq location","company city","company state"],
        "SalesNavUrl": ["salesnavurl","sales navigator url","sales navigator link"],
        "query": ["search query","query string"],
        "timestamp": ["time","created at","added at","timestamp"],
        "Relevance": ["relevance","is relevant","relevant"],
        "Focus Area": ["focus","focus area","focus_area"],
        "Seniority": ["seniority","level","seniority level"],
        "Bank Size": ["bank size","size","segment"],
    }
    if target in aliases:
        scored = process.extractOne(
            norm(target),
            [norm(c) for c in in_cols] + aliases[target],
            scorer=fuzz.ratio
        )
        if scored:
            for c in in_cols:
                if norm(c) == scored[0]:
                    return c
    scored = process.extractOne(norm(target), [norm(c) for c in in_cols], scorer=fuzz.token_sort_ratio)
    if scored and scored[1] >= 85:
        for c in in_cols:
            if norm(c) == scored[0]:
                return c
    return ""

def apply_mapping(df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    out = pd.DataFrame()
    for tgt in TARGET_HEADERS:
        src = mapping.get(tgt, "")
        if src and src in df.columns:
            out[tgt] = df[src].astype(str)
        else:
            out[tgt] = pd.Series([None] * len(df))
    return out

def segment(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    # Derive guaranteed classifications
    df["_CompanyType"] = df.get("companyName", pd.Series([None]*len(df))).apply(classify_company_type)
    df["_Focus"] = df.get("title", pd.Series([None]*len(df))).apply(classify_focus)
    df["_State"] = df.get("location", pd.Series([None]*len(df))).apply(best_state_match)

    states = list(STATE_LABELS.keys())  # includes "Other States"
    company_types = ["Bank", "Credit Union"]
    focuses = [
        "Data / AI / Automation",
        "Transformation, digital, innovation",
        "Risk and Compliance",
        "All Others",
    ]

    buckets = {}
    for st_name in states:
        for ct in company_types:
            for fc in focuses:
                mask = (df["_State"] == st_name) & (df["_CompanyType"] == ct) & (df["_Focus"] == fc)
                key = f"{st_name} - {ct} - {fc}"
                buckets[key] = df[mask].copy()
    return buckets

def _safe_sheet_name(name: str, used: set) -> str:
    n = re.sub(r'[\[\]\:\*\?\/\\]', '_', name)[:31] or "Sheet"
    base, i = n, 1
    while n in used:
        suffix = f"_{i}"
        n = (base[:31 - len(suffix)] + suffix)
        i += 1
    used.add(n)
    return n

def _excel_bytes_from_df(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = io.BytesIO()
    try:
        import xlsxwriter  # noqa
        engine = "xlsxwriter"
    except ModuleNotFoundError:
        engine = "openpyxl"
    with pd.ExcelWriter(buf, engine=engine) as writer:
        df.to_excel(writer, index=False, sheet_name=_safe_sheet_name(sheet_name, set()))
    return buf.getvalue()

def _sanitize_filename(name: str) -> str:
    n = re.sub(r'[^\w\-\.\(\) ]+', '_', name)
    return n.strip().replace("  ", " ")

# ----------------------- Messaging -----------------------

def domain_labels(company_type: str) -> Dict[str, str]:
    if company_type == "Credit Union":
        return {"BANKING": "credit union", "BANKS": "credit unions"}
    return {"BANKING": "banking", "BANKS": "banks"}

def state_across(state: str) -> str:
    return "the US" if state == "Other States" else state

def state_in(state: str) -> str:
    return "the US" if state == "Other States" else state

def focus_phrase_short(focus: str) -> str:
    if focus == "Data / AI / Automation":
        return "data and AI"
    if focus == "Transformation, digital, innovation":
        return "digital transformation and innovation"
    if focus == "Risk and Compliance":
        return "risk and compliance"
    return "your work"

def message_pack(state: str, company_type: str, focus: str) -> Dict[str, str]:
    dom = domain_labels(company_type)
    fp = focus_phrase_short(focus)

    conn = (
        f"Hi {{FirstName}},\n\n"
        f"I’m connecting with {dom['BANKING']} leaders across {state_across(state)} who are driving projects in AI, data and automation. "
        f"It would be great to connect and share insights ahead of a summit in Charlotte this November.\n\n"
        f"Best regards, Mark"
    )
    f1 = (
        f"Hi {{FirstName}},\n\n"
        f"I’m reaching out to senior {dom['BANKING']} leaders in {state_in(state)} as we prepare for the Banking Transformation Summit in Charlotte this Nov "
        f"- focused on data, AI and digital change. Thought this could be relevant for you.\n\n"
        f"Best regards, Mark"
    )
    f2 = (
        f"Hi {{FirstName}},\n\n"
        f"We organise Banking Transformation Summit – the leading AI and transformation event for {dom['BANKS']}. "
        f"Given your role with {fp}, I thought this would be highly relevant and wanted to share a VIP guest pass to attend.\n\n"
        f"Best regards, Mark"
    )
    return {"Connection": conn, "Follow Up 1": f1, "Follow Up 2": f2}

def build_summary_df(buckets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for seg_name, df in buckets.items():
        try:
            state, company_type, focus = [x.strip() for x in seg_name.split(" - ", 3)]
        except Exception:
            state, company_type, focus = "Other States", "Bank", "All Others"
        msgs = message_pack(state, company_type, focus)
        rows.append({
            "Segment": seg_name,
            "Count": len(df),
            "Connection": msgs["Connection"],
            "Follow Up 1": msgs["Follow Up 1"],
            "Follow Up 2": msgs["Follow Up 2"],
        })
    return pd.DataFrame(rows).sort_values("Segment").reset_index(drop=True)

# ----------------------- UI -----------------------

st.set_page_config(page_title="Segmented Outreach Builder", layout="wide")
st.title("Segmented Outreach Builder")

uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
if not uploaded:
    st.stop()

if uploaded.name.lower().endswith(".csv"):
    raw = pd.read_csv(uploaded, dtype=str, keep_default_na=False)
else:
    raw = pd.read_excel(uploaded, dtype=str, keep_default_na=False)

st.write(f"Rows: {len(raw):,}  Columns: {len(raw.columns)}")
st.dataframe(raw.head(10), use_container_width=True)

st.subheader("Map columns to target schema")
in_cols = list(raw.columns)

mapping: Dict[str, str] = {}
cols_left, cols_right = st.columns(2)
half = len(TARGET_HEADERS) // 2
with cols_left:
    for tgt in TARGET_HEADERS[:half]:
        guess = auto_guess_mapping(in_cols, tgt)
        opts = [""] + in_cols
        idx = opts.index(guess) if guess in in_cols else 0
        mapping[tgt] = st.selectbox(f"{tgt}", options=opts, index=idx)
with cols_right:
    for tgt in TARGET_HEADERS[half:]:
        guess = auto_guess_mapping(in_cols, tgt)
        opts = [""] + in_cols
        idx = opts.index(guess) if guess in in_cols else 0
        mapping[tgt] = st.selectbox(f"{tgt}", options=opts, index=idx)

st.caption("Every contact is assigned. Unmapped columns are left blank in the export.")

if st.button("Build segments and download ZIP"):
    df = apply_mapping(raw, mapping)

    # Safeguards for key fields
    for must in ["companyName", "title", "location"]:
        if must not in df.columns:
            df[must] = None

    buckets = segment(df)

    # Summary workbook
    summary_df = build_summary_df(buckets)
    summary_bytes = _excel_bytes_from_df(summary_df, sheet_name="Summary")

    # ZIP with per-segment CSVs + summary.xlsx
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("summary.xlsx", summary_bytes)
        for name, d in buckets.items():
            cols = [c for c in TARGET_HEADERS if c in d.columns] + [c for c in d.columns if c not in TARGET_HEADERS]
            d = d[cols]
            csv_bytes = d.to_csv(index=False).encode("utf-8-sig")
            zf.writestr(_sanitize_filename(f"{name}.csv"), csv_bytes)

    st.success("ZIP created.")
    st.download_button(
        "Download ZIP of segments + summary.xlsx",
        data=zip_buf.getvalue(),
        file_name="segmented_outreach.zip",
        mime="application/zip",
    )

    totals = {k: len(v) for k, v in buckets.items()}
    stats = pd.DataFrame(sorted(totals.items()), columns=["Segment", "Count"])
    st.dataframe(stats, use_container_width=True)

    st.info("Adjust STATE_FUZZ_THRESHOLD, keyword lists, and Bank/Credit Union hints to tune allocations.")
