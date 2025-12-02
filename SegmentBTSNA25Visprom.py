# app.py
# Streamlit segmentation app for Banks vs Credit Unions with 14 segments and ZIP export
# ------------------------------------------------------------
# pip install streamlit pandas openpyxl chardet

import io
import re
import zipfile
from typing import Dict, Optional

import chardet
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bank & Credit Union Segmentation", layout="wide")

# ------------------------------------------------------------
# Canonical headers (MUST be present in row 1)
# ------------------------------------------------------------
REQUIRED_HEADERS = [
    "email", "firstname", "lastname", "jobtitle",
    "Company Name", "country", "State", "Seniority"
]

# Replace the old SENIOR_FOR_OTHERS block with this:

# Values are compared case-insensitively
SENIOR_FOR_OTHERS = {
    "director",
    "c-suite",
    "vice president",
    "leadership",
    "head",
    "chair/board member",
}

SENIOR_TITLE_REGEX = re.compile(
    r"\b("
    r"chief|c[-\s]?suite|cfo|cio|cto|cdo|ciso|cro|ceo|coo|"
    r"chair|board\s+member|"
    r"executive\s+vice\s+president|evp|senior\s+vice\s+president|svp|vice\s+president|vp|"
    r"managing\s+director|md|director|head|"
    r"senior\s+manager|manager|managing\b"
    r")\b", re.IGNORECASE
)

def is_most_senior(seniority: str, jobtitle: str) -> bool:
    """
    True if seniority column clearly indicates leadership per spec,
    or (fallback) if the job title itself carries senior leadership tokens.
    """
    s = norm(seniority)  # lowercased + trimmed
    if s in SENIOR_FOR_OTHERS:
        return True
    # Some files might contain variants like "Senior Vice President" or "EVP" in title:
    if not s or s not in SENIOR_FOR_OTHERS:
        if SENIOR_TITLE_REGEX.search(jobtitle or ""):
            return True
    return False


# ------------------------------------------------------------
# Normalisation helpers
# ------------------------------------------------------------
def norm(s: str) -> str:
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def expand_jobtitle(t: str) -> str:
    """Lowercase + add expansions (AI, ML, IAM, PCI, etc.) to boost keyword recall."""
    t = norm(t)
    t = re.sub(r"[|/&(),+]", " ", t)
    t = re.sub(r"[-]+", " ", t)
    expansions = [
        (r"\bai\b", " artificial intelligence"),
        (r"\bml\b", " machine learning"),
        (r"\bnlp\b", " natural language processing"),
        (r"\brpa\b", " robotic process automation automation"),
        (r"\bbi\b", " business intelligence"),
        (r"\bcx\b", " customer experience"),
        (r"\bux\b", " user experience"),
        (r"\bui\b", " user interface"),
        (r"\bgrc\b", " governance risk compliance"),
        (r"\biam\b", " identity and access management identity access"),
        (r"\bkyc\b", " know your customer"),
        (r"\bkyb\b", " know your business"),
        (r"\bcdd\b", " customer due diligence"),
        (r"\bedd\b", " enhanced due diligence"),
        (r"\bofac\b", " sanctions ofac"),
        (r"\bpci\b", " pci dss"),
        (r"\bncino\b", " core banking ncino"),
        (r"\bfed\s?now\b|\bfednow\b", " instant payments fednow"),
        (r"\bertp\b", " real time payments"),
        (r"\bitgc\b", " it general controls"),
        (r"\bpsm\b", " payment systems management"),
    ]
    for rx, add in expansions:
        if re.search(rx, t):
            t += " " + add
    if re.search(r"\bmachine learnin\b", t):
        t += " machine learning"
    return t

def token_regex(phrase: str, flags=re.IGNORECASE):
    safe = re.escape(phrase.strip().lower())
    safe = re.sub(r"\\\s+", r"[\\s\\-/]+", safe)  # allow space/slash/hyphen
    return re.compile(rf"\b{safe}\b", flags)

# ------------------------------------------------------------
# Robust file reading (always uses first row as headers)
# ------------------------------------------------------------
def read_uploaded_file(file) -> pd.DataFrame:
    name = file.name.lower()

    if name.endswith(".xlsx"):
        df = pd.read_excel(file, engine="openpyxl", header=0)
    else:
        raw = file.getvalue()
        enc = chardet.detect(raw).get("encoding") or "utf-8"
        try:
            df = pd.read_csv(io.BytesIO(raw), engine="python", sep=None, header=0)
        except Exception:
            try:
                df = pd.read_csv(io.BytesIO(raw), engine="python", sep=None, header=0, encoding=enc)
            except Exception:
                df = pd.read_csv(io.BytesIO(raw), engine="python", sep=None, header=0,
                                 encoding="latin-1", on_bad_lines="skip")

    df.columns = [str(c).strip() for c in df.columns]
    return df

# ------------------------------------------------------------
# Organisation (Bank vs Credit Union)
# ------------------------------------------------------------
CU_PATTERNS = [
    r"\bfederal credit union\b",
    r"\bcredit union\b",
    r"\bteachers? cu\b",
    r"\bcu\b",
    r"\bfcu\b",
]
CU_COMPILED = [re.compile(pat, re.IGNORECASE) for pat in CU_PATTERNS]

def classify_org(company_name: str) -> str:
    s = " ".join(re.findall(r"[A-Za-z0-9&.'\-]+", str(company_name) or "", flags=re.UNICODE)).strip().lower()
    if any(rx.search(s) for rx in CU_COMPILED):
        return "Credit Union"
    if re.search(r"\bcu\b|\bfcu\b", s):
        return "Credit Union"
    return "Bank"

# ------------------------------------------------------------
# Strong Technology anchors (tight eligibility)
# ------------------------------------------------------------
TECH_STRONG_ANCHORS = [
    r"\b(it|information technology)\b",
    r"\bsystem(s)?\b", r"\bsystems?\s+engineer(ing)?\b",
    r"\bapplication\s+analyst\b", r"\bapplication(s)?\s+support\b",
    r"\binfrastructure\b", r"\bnetwork(ing)?\b", r"\bnetwork\s+security\b",
    r"\b(devops|sre)\b", r"\bcloud\b", r"\bplatform\b", r"\bendpoint\b",
    r"\bservice\s+desk\b", r"\bhelp\s*desk\b", r"\bsupport(\s+technician|\s+specialist|\b)\b",
    r"\barchitect\b", r"\bsolutions?\s+architect\b", r"\benterprise\s+architect\b",
    r"\badministrator\b", r"\bit\s+operations\b", r"\bit\s+service\s+delivery\b",
    r"\bsoftware\s+engineer\b", r"\bprincipal\s+software\s+engineer\b",
    r"\bengineering\s+director\b", r"\btechnical\b", r"\btechnical\s+engineering\b",
    r"\bworkday\s+application\s+support\b", r"\btelecommunications?\b",
    r"\bcommunications\s+systems?\s+specialist\b"
]
TECH_ANCHOR_RX = [re.compile(p, re.IGNORECASE) for p in TECH_STRONG_ANCHORS]
def tech_eligible(job: str) -> bool:
    return any(rx.search(job) for rx in TECH_ANCHOR_RX)

# ------------------------------------------------------------
# Segment rules (STRICT); order is priority
#   Credit Union: Risk > Compliance > Data > AI > Digital Banking > Technology > Others > Unassigned
#   Bank:        Risk > Compliance > Data > AI > Innovation > Technology > Others > Unassigned
# ------------------------------------------------------------
def build_segment_rules():
    # RISK — must contain "risk" (core) or explicit risk phrases
    RISK_CORE = ["risk"]
    RISK_GEN = [
        "credit risk", "liquidity risk", "model risk", "operational risk", "enterprise risk",
        "risk officer", "risk management", "portfolio risk", "erm"
    ]

    # COMPLIANCE — explicit regulatory/financial crimes signals
    COMP_CORE = ["compliance", "audit", "assurance"]
    COMP_GEN = [
        "regulatory", "internal audit", "sox", "ofac", "bsa", "aml", "fraud", "sanctions",
        "financial crime", "grc", "controls", "controls testing", "quality control",
        "kyc", "kyb", "cdd", "edd", "pci dss", "transaction monitoring",
        "information security risk and governance"
    ]

    # DATA — VERY tight: data/analytics/analyst/analysis/BI only
    DATA_CORE = ["data", "analytics", "analyst", "analysis", "business intelligence", "bi"]
    DATA_GEN = [
        # include only phrases that contain those stems to avoid false positives
        "data science", "data engineering", "data engineer", "data governance",
        "data platform", "data warehouse", "data operations", "data architecture",
        "advanced analytics", "pricing analytics", "strategic analytics", "finance analytics",
        "bi developer", "bi analyst", "enterprise analytics", "market intelligence", "member intelligence"
    ]

    # AI — explicit AI/ML only
    AI_CORE = ["ai", "artificial intelligence", "machine learning", "ml"]
    AI_GEN = ["generative ai", "gen ai", "deep learning", "nlp", "llm"]

    # DIGITAL BANKING — explicit banking-channel words
    DIGIBANK_CORE = ["digital banking", "online banking", "mobile banking"]
    DIGIBANK_GEN = [
        "digital channels", "digital onboarding", "client experience", "omnichannel",
        "digital operations", "contact center digital", "contact centre digital"
    ]

    # INNOVATION (BANKS) — tight innovation signals only
    INNOV_CORE = ["innovation"]
    INNOV_GEN = ["incubation", "r&d", "lab", "ventures", "innovation & integration"]

    # TECHNOLOGY — strong IT/engineering anchors only (see tech_eligible)
    TECH_CORE = [
        "it", "infrastructure", "network", "systems", "platform",
        "engineer", "architect", "administrator", "support", "operations",
        "devops", "sre", "cloud", "service desk", "helpdesk",
        "application analyst", "software engineer"
    ]
    TECH_GEN = [
        "cloud operations", "cloud administrator", "endpoint", "configuration",
        "system administrator", "systems engineer", "environment", "server", "hardware",
        "telecommunications", "communications engineer", "network operations", "iam",
        "enterprise architect", "solutions architect", "solution architect", "architecture",
        "it governance", "it operations", "it service delivery", "it manager", "it director",
        "workday application support", "eiam", "it administration", "it asset",
        "endpoint management engineer", "principal it engineer",
        "it operations and service management", "communications systems specialist",
        "it delivery manager", "service delivery manager", "it support specialist",
        "it support manager", "network and telecommunications",
        "technical facilitator", "it technician", "it hardware technician",
        "it support", "it business relationship manager",
        "enterprise cloud platform manager"
    ]

    # Priority lists
    cu_rules = [
        ("Risk", RISK_CORE, RISK_GEN),
        ("Compliance", COMP_CORE, COMP_GEN),
        ("Data", DATA_CORE, DATA_GEN),
        ("AI", AI_CORE, AI_GEN),
        ("Digital Banking", DIGIBANK_CORE, DIGIBANK_GEN),
        ("Technology", TECH_CORE, TECH_GEN),
        ("Others (most senior)", [], []),
    ]
    bank_rules = [
        ("Risk", RISK_CORE, RISK_GEN),
        ("Compliance", COMP_CORE, COMP_GEN),
        ("Data", DATA_CORE, DATA_GEN),
        ("AI", AI_CORE, AI_GEN),
        ("Innovation", INNOV_CORE, INNOV_GEN),
        ("Technology", TECH_CORE, TECH_GEN),
        ("Others (most senior)", [], []),
    ]

    def compile_rules(rules):
        compiled = []
        for seg, core, gens in rules:
            core_rx = [token_regex(c) for c in core]
            gen_rx = [token_regex(g) for g in gens]
            compiled.append((seg, core_rx, gen_rx))
        return compiled

    return compile_rules(cu_rules), compile_rules(bank_rules)

CU_RULES, BANK_RULES = build_segment_rules()

# ------------------------------------------------------------
# Overrides & Exceptions (intent-first, but conservative)
# ------------------------------------------------------------
def dual_overrides(job: str, org_type: str) -> Optional[str]:
    # Risk + Analytics → Risk (keeps 'Head of Advanced Analytics and AI' in Data unless 'risk' present)
    if re.search(r"\brisk\b", job) and re.search(r"\banalytics?\b", job):
        return "Risk"
    # IAM/Identity Access → Compliance
    if re.search(r"\b(identity\s*(and\s*)?access|iam|identity access management)\b", job):
        return "Compliance"
    # Digital + Product (banks): Innovation
    if org_type == "Bank" and re.search(r"\bdigital\b", job) and re.search(r"\bproduct\b", job):
        return "Innovation"
    # Digital Banking anchors for CUs
    if org_type == "Credit Union" and re.search(r"\b(digital|online|mobile)\s+banking\b", job):
        return "Digital Banking"
    return None

def special_exceptions(job: str, org_type: str) -> Optional[str]:
    # IT Business Analyst → Technology
    if re.search(r"^(it|information technology)\b", job) and re.search(r"\bbusiness analyst\b", job):
        return "Technology"
    return None

# ------------------------------------------------------------
# Matching with STRICT Data/Tech & priority ordering
# ------------------------------------------------------------
def match_segment(jobtitle: str, seniority: str, org_type: str) -> Optional[str]:
    jt = expand_jobtitle(jobtitle)
    sr = norm(seniority)
    rules = CU_RULES if org_type == "Credit Union" else BANK_RULES

    # 0) Overrides
    ovr = dual_overrides(jt, org_type)
    if ovr:
        return ovr

    ex = special_exceptions(jt, org_type)
    if ex:
        return ex

    # 1) Core hits in priority order
    for seg, core_rx, _ in rules:
        if seg == "Others (most senior)":
            continue
        if any(rx.search(jt) for rx in core_rx):
            if seg == "Technology" and not tech_eligible(jt):
                continue
            if seg == "Data":
                # enforce tight data rule: one of the strict stems must be present
                if not re.search(r"\b(data|analytics?|analyst|analysis|business intelligence|bi)\b", jt):
                    continue
            if seg == "Innovation":
                # must contain 'innovation' or close kin (not vague 'transformation')
                if not re.search(r"\binnovation|incubation|r&d|lab|ventures\b", jt):
                    continue
            return seg

    # 2) General hits (score), still respecting strictness
    best_seg = None
    best_score = 0
    for seg, _, gen_rx in rules:
        if seg == "Others (most senior)":
            continue
        score = 0
        for rx in gen_rx:
            m = rx.findall(jt)
            if m:
                score += len(m)
        if score > 0:
            if seg == "Technology" and not tech_eligible(jt):
                continue
            if seg == "Data" and not re.search(r"\b(data|analytics?|analyst|analysis|business intelligence|bi)\b", jt):
                continue
            if seg == "Innovation" and not re.search(r"\binnovation|incubation|r&d|lab|ventures\b", jt):
                continue
            if score > best_score:
                best_score = score
                best_seg = seg
    if best_seg:
        return best_seg

    # 3) Others (most senior) via Seniority
    if sr in SENIOR_FOR_OTHERS:
        return "Others (most senior)"

    # 4) Unassigned if nothing matches (no forced catch-all)
    return None

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Bank & Credit Union Segmentation")
st.caption("Headers are taken **from row 1**: `email, firstname, lastname, jobtitle, Company Name, country, State, Seniority`.")

with st.sidebar:
    st.header("1) Upload file")
    uploaded = st.file_uploader("Upload .xlsx or .csv (headers in row 1)", type=["xlsx", "csv"])

if not uploaded:
    st.info("Upload a file to begin.")
    st.stop()

df_raw = read_uploaded_file(uploaded)
st.success(f"Loaded {len(df_raw):,} rows, {len(df_raw.columns)} columns.")

# Validate headers
missing = [h for h in REQUIRED_HEADERS if h not in df_raw.columns]
if missing:
    st.error(f"Missing required header(s) (must be in row 1): {', '.join(missing)}")
    st.write("Detected headers:", list(df_raw.columns))
    st.stop()

# Build working dataframe with canonical names (exact columns)
df = pd.DataFrame({
    "email": df_raw["email"],
    "firstname": df_raw["firstname"],
    "lastname": df_raw["lastname"],
    "jobtitle": df_raw["jobtitle"],
    "Company Name": df_raw["Company Name"],
    "country": df_raw["country"],
    "State": df_raw["State"],
    "Seniority": df_raw["Seniority"],
})

# Classify and segment
st.header("2) Classify organisations & segment")
with st.spinner("Classifying organisations and segmenting contacts..."):
    df["Organisation Type"] = df["Company Name"].apply(classify_org)
    df["Segment"] = [
        match_segment(jt, sr, org)
        for jt, sr, org in zip(df["jobtitle"], df["Seniority"], df["Organisation Type"])
    ]
    df["Segment"] = df["Segment"].fillna("Unassigned")

# ------------------------------------------------------------
# Summary + Samples (fixed ordering; includes Unassigned)
# ------------------------------------------------------------
st.subheader("Summary")

CU_SEG_ORDER = ["Risk", "Compliance", "Data", "AI", "Digital Banking", "Technology", "Others (most senior)", "Unassigned"]
BANK_SEG_ORDER = ["Risk", "Compliance", "Data", "AI", "Innovation", "Technology", "Others (most senior)", "Unassigned"]

counts = df.groupby(["Organisation Type", "Segment"]).size().reset_index(name="count")
cu_order = {seg: i for i, seg in enumerate(CU_SEG_ORDER)}
bank_order = {seg: i for i, seg in enumerate(BANK_SEG_ORDER)}

def seg_rank(org: str, seg: str) -> int:
    if org == "Credit Union":
        return cu_order.get(seg, 999)
    return bank_order.get(seg, 999)

counts["seg_rank"] = counts.apply(lambda r: seg_rank(r["Organisation Type"], r["Segment"]), axis=1)
counts = counts.sort_values(by=["Organisation Type", "seg_rank"]).drop(columns=["seg_rank"])
st.dataframe(counts, use_container_width=True)

st.subheader("Samples")
org_pick = st.selectbox("Organisation Type", ["Credit Union", "Bank"])
seg_pick = st.selectbox("Segment", CU_SEG_ORDER if org_pick == "Credit Union" else BANK_SEG_ORDER)
sample_n = st.slider("Rows to preview", min_value=5, max_value=100, value=20, step=5)
sample_df = df[(df["Organisation Type"] == org_pick) & (df["Segment"] == seg_pick)].head(sample_n)
st.dataframe(sample_df, use_container_width=True)

# ------------------------------------------------------------
# Export ZIP (now also includes Unassigned)
# ------------------------------------------------------------
st.header("3) Download per-segment CSVs (ZIP)")
buffer = io.BytesIO()
with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
    # Credit Union segments
    for seg in CU_SEG_ORDER:
        subset = df[(df["Organisation Type"] == "Credit Union") & (df["Segment"] == seg)]
        zf.writestr(f"Credit Union - {seg}.csv", subset.to_csv(index=False))
    # Bank segments
    for seg in BANK_SEG_ORDER:
        subset = df[(df["Organisation Type"] == "Bank") & (df["Segment"] == seg)]
        zf.writestr(f"Bank - {seg}.csv", subset.to_csv(index=False))
st.download_button("Download all segments as ZIP", data=buffer.getvalue(),
                   file_name="segments_export.zip", mime="application/zip")

with st.expander("Advanced: Credit Union detection patterns"):
    st.write("Patterns:", ", ".join([p for p in [r"\bfederal credit union\b", r"\bcredit union\b", r"\bteachers? cu\b", r"\bcu\b", r"\bfcu\b"]]))
    st.caption("If you see a misclassification, share the company name and we’ll harden the regex safely.")


# ------------------------------------------------------------
# Master CSV (merged) download — FIXED sort
# ------------------------------------------------------------
st.header("4) Download master segment list (single CSV)")

include_banks = st.checkbox("Include Banks", value=True)
include_cus = st.checkbox("Include Credit Unions", value=True)

master_df = df.copy()
if not include_banks and include_cus:
    master_df = master_df[master_df["Organisation Type"] == "Credit Union"]
elif include_banks and not include_cus:
    master_df = master_df[master_df["Organisation Type"] == "Bank"]

export_cols = [
    "email", "firstname", "lastname", "jobtitle",
    "Company Name", "country", "State",
    "Organisation Type", "Segment", "Seniority"
]
export_cols = [c for c in export_cols if c in master_df.columns]
master_out = master_df[export_cols].copy()

# Build a per-row rank for Segment that depends on Organisation Type
CU_SEG_ORDER = ["Risk", "Compliance", "Data", "AI", "Digital Banking", "Technology", "Others (most senior)", "Unassigned"]
BANK_SEG_ORDER = ["Risk", "Compliance", "Data", "AI", "Innovation", "Technology", "Others (most senior)", "Unassigned"]

order_map = {
    "Credit Union": {s: i for i, s in enumerate(CU_SEG_ORDER)},
    "Bank": {s: i for i, s in enumerate(BANK_SEG_ORDER)},
}

def seg_rank_row(org_type: str, seg: str) -> int:
    return order_map.get(org_type, {}).get(seg, 999)

if not master_out.empty:
    master_out["seg_rank"] = master_out.apply(
        lambda r: seg_rank_row(r["Organisation Type"], r["Segment"]),
        axis=1
    )
    master_out = master_out.sort_values(
        by=["Organisation Type", "seg_rank", "Company Name", "lastname", "firstname"]
    ).drop(columns=["seg_rank"])

master_csv_bytes = master_out.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    label="Download master CSV",
    data=master_csv_bytes,
    file_name="master_segment_list.csv",
    mime="text/csv",
    disabled=master_out.empty
)
