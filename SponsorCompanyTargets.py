# sponsor_target_matcher.py
# Streamlit app to compare sponsor target lists against a scraped master list with robust canonicalisation + fuzzy matching.
# Exports multi-sheet Excel. Falls back to openpyxl if xlsxwriter is missing.

import io
import re
from collections import defaultdict
from typing import List, Tuple, Dict, Optional

import pandas as pd
import streamlit as st

# Fuzzy library
try:
    from rapidfuzz import process, fuzz
    HAVE_RAPIDFUZZ = True
except Exception:
    import difflib
    HAVE_RAPIDFUZZ = False

# Excel writer engine fallback
def get_excel_writer(bytes_buf):
    try:
        import xlsxwriter  # noqa: F401
        return pd.ExcelWriter(bytes_buf, engine="xlsxwriter")
    except Exception:
        return pd.ExcelWriter(bytes_buf, engine="openpyxl")

st.set_page_config(page_title="Sponsor Target Matcher", page_icon="📊", layout="wide")
st.title("📊 Sponsor Target Matcher")
st.caption("Paste scraped companies, add sponsors and targets, then export clean Excel. Now with strong canonicalisation for brand variants.")

# ---------------- Canonicalisation helpers ----------------

PUNCT_RE = re.compile(r"[\.\,&\-/\(\)'\*]")

COMMON_SUFFIXES = [
    r",?\s+incorporated$", r",?\s+inc\.$", r",?\s+inc$", r",?\s+co\.$", r",?\s+company$",
    r",?\s+corp\.$", r",?\s+corporation$", r",?\s+llc$", r",?\s+l\.l\.c\.$",
    r",?\s+lp$", r",?\s+l\.p\.$", r",?\s+llp$", r",?\s+l\.l\.p\.$", r",?\s+plc$",
    r",?\s+limited$", r",?\s+ltd\.$", r",?\s+ltd$", r",?\s+sa$", r",?\s+ag$", r",?\s+spa$",
    r",?\s+gmbh$", r",?\s+pte\.?\s+ltd\.?$",
    r",?\s+n\.a\.$", r",?\s+na$", r",?\s+national\s+association$",
    r",?\s+credit\s+union$", r",?\s+financial\s+group$", r",?\s+financials?$",
    r",?\s+group$", r",?\s+holdings?$", r",?\s+bank$", r",?\s+bank\s+na$",
    r",?\s+securities$", r",?\s+markets?$", r",?\s+capital\s+markets?$",
]

SUFFIX_RE = re.compile("(" + "|".join(COMMON_SUFFIXES) + r")\s*$", re.IGNORECASE)

# Unit and geography labels to drop to get brand-core
STOP_TOKENS = {
    "bank", "banks", "credit", "union", "financial", "finance", "group", "holding", "holdings",
    "corp", "corporation", "company", "limited", "ltd", "inc", "llc", "plc", "na", "n.a",
    "usa", "u.s.", "us", "uk", "canada", "texas", "ny", "nyc", "california", "america",
    "banamex", "cib", "securities", "markets", "capital", "national", "association",
    "nb", "n.a.", "nv", "ag", "sa", "spa", "gmbh"
}

def strip_parentheticals(s: str) -> str:
    # Remove any (...) blocks
    return re.sub(r"\s*\([^)]*\)", "", s)

def normalise(name: str) -> str:
    if not isinstance(name, str):
        return ""
    s = name.strip()
    s = strip_parentheticals(s)
    s = PUNCT_RE.sub(" ", s)
    s = SUFFIX_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).lower().strip()
    return s

def brand_core_key(name: str) -> str:
    n = normalise(name)
    tokens = [t for t in n.split() if t and t not in STOP_TOKENS]
    return " ".join(tokens) or n

def dedupe_keep_first(seq: List[str]) -> List[str]:
    seen = set()
    out = []
    for x in seq:
        k = x.strip()
        if k and k.lower() not in seen:
            seen.add(k.lower())
            out.append(k)
    return out

def parse_multiline(text: str) -> List[str]:
    return dedupe_keep_first([ln.strip() for ln in text.splitlines() if ln.strip()])

def best_match_q(query: str, universe_display: List[str], universe_key: List[str]) -> Tuple[Optional[str], float]:
    """Return (best_match_original_display, score 0-100) using brand core keys."""
    qkey = brand_core_key(query)
    if not qkey or not universe_key:
        return None, 0.0

    if HAVE_RAPIDFUZZ:
        match = process.extractOne(qkey, universe_key, scorer=fuzz.token_set_ratio)
        if match is None:
            return None, 0.0
        idx = match[2]
        return universe_display[idx], float(match[1])
    else:
        scores = [difflib.SequenceMatcher(None, qkey, cand).ratio() for cand in universe_key]
        if not scores:
            return None, 0.0
        idx = int(max(range(len(scores)), key=lambda i: scores[i]))
        return universe_display[idx], float(scores[idx] * 100.0)

# ---------------- UI ----------------

with st.sidebar:
    st.subheader("Settings")
    match_threshold = st.slider("Match threshold", 50, 100, 88, 1,
                                help="Minimum score to accept a fuzzy match as included.")
    show_previews = st.checkbox("Show preview tables", value=True)
    st.markdown("---")
    st.markdown("**Sponsors**")
    default_sponsors = ["LifeRay", "StackAI", "Synechron", "Thoughtspot"]
    sponsor_text = st.text_area("Sponsor names (one per line)", value="\n".join(default_sponsors), height=120)
    st.markdown("---")
    st.markdown("**Manual Aliases (optional)**")
    st.caption("One per line. Use 'Variant => Canonical'. Example: 'Citibank => Citigroup'")
    alias_text = st.text_area("Alias map", value="", height=120)

st.subheader("1) Paste scraped company names")
scraped_block = st.text_area(
    "Scraped companies - paste the unique list here (one per line). You can also upload a CSV/XLSX below.",
    value="", height=200,
    placeholder="BNP Paribas CIB\nCitibank (Banamex USA)\nCitigroup (Texas)\nThe Toronto-Dominion Bank"
)

uploaded = st.file_uploader("Or upload CSV/XLSX with a single column of company names", type=["csv", "xlsx"])
uploaded_names: List[str] = []

if uploaded is not None:
    try:
        if uploaded.name.lower().endswith(".csv"):
            df_u = pd.read_csv(uploaded)
        else:
            df_u = pd.read_excel(uploaded)
        # Guess name column
        col_guess = None
        for c in df_u.columns:
            if str(c).strip().lower() in {"company", "companyname", "name", "organisation", "organization"}:
                col_guess = c
                break
        if col_guess is None:
            col_guess = df_u.columns[0]
        uploaded_names = df_u[col_guess].dropna().astype(str).tolist()
        st.success(f"Loaded {len(uploaded_names)} rows from {uploaded.name} using column '{col_guess}'.")
    except Exception as e:
        st.error(f"Could not read file. {e}")

# Combine and dedupe raw scraped names
scraped_raw = dedupe_keep_first(parse_multiline(scraped_block) + uploaded_names)

# Build alias map
alias_map: Dict[str, str] = {}
for line in parse_multiline(alias_text):
    if "=>" in line:
        left, right = [p.strip() for p in line.split("=>", 1)]
        if left and right:
            alias_map[left.lower()] = right

def apply_alias(name: str) -> str:
    return alias_map.get(name.lower(), name)

# Canonicalise scraped list into brand cores
canonical_to_originals: Dict[str, List[str]] = defaultdict(list)
for nm in scraped_raw:
    aliased = apply_alias(nm)
    key = brand_core_key(aliased)
    canonical_to_originals[key].append(nm)

# Pick a display representative per canonical entity (first occurrence)
canonical_display = []
canonical_keys = []
for key, originals in canonical_to_originals.items():
    representative = originals[0]
    canonical_display.append(representative)
    canonical_keys.append(key)

st.write(f"Scraped universe: raw **{len(scraped_raw)}** → canonical entities **{len(canonical_display)}**")

if show_previews and scraped_raw:
    prev_df = pd.DataFrame(
        [(k, ", ".join(v[:6]) + (" ..." if len(v) > 6 else "")) for k, v in sorted(canonical_to_originals.items(), key=lambda x: x[0])]
        , columns=["Canonical Key", "Example Variants"])
    st.expander("Preview canonicalisation - click to expand").dataframe(prev_df, use_container_width=True)

st.subheader("2) Add sponsors and paste their target companies")
sponsors = parse_multiline(sponsor_text)
if not sponsors:
    st.info("Add at least one sponsor in the sidebar.")
    st.stop()

sponsor_targets: Dict[str, List[str]] = {}
cols = st.columns(2)
for i, sp in enumerate(sponsors):
    with cols[i % 2]:
        with st.expander(f"{sp} - paste targets", expanded=False):
            txt = st.text_area(f"Targets for {sp} (one per line)", key=f"ta_{sp}", height=180)
            sponsor_targets[sp] = parse_multiline(txt)

# ---------------- Matching ----------------

if st.button("Generate report"):
    if not canonical_display:
        st.error("Please provide scraped companies first.")
        st.stop()

    sponsor_results: Dict[str, Dict[str, pd.DataFrame]] = {}
    missing_map = defaultdict(set)

    for sp, targets in sponsor_targets.items():
        rows = []
        for t in targets:
            # Apply alias on target to align with manual mappings
            t_eff = apply_alias(t)
            best, score = best_match_q(t_eff, canonical_display, canonical_keys)
            in_scrape = bool(best) and score >= match_threshold

            # Also include the canonical key so you can sanity check what we compared
            rows.append({
                "Target": t,
                "TargetCore": brand_core_key(t_eff),
                "BestMatchInScrape": best if best else "",
                "MatchScore": round(score, 1),
                "IncludedInScrape": "Yes" if in_scrape else "No"
            })
            if not in_scrape:
                missing_map[t].add(sp)

        df = pd.DataFrame(rows).sort_values(["IncludedInScrape", "MatchScore"], ascending=[True, False])
        sponsor_results[sp] = {
            "all": df,
            "in": df[df["IncludedInScrape"] == "Yes"].copy(),
            "out": df[df["IncludedInScrape"] == "No"].copy(),
        }

    # Build summary of missing across all sponsors
    summary_rows = []
    for company, sps in sorted(missing_map.items(), key=lambda x: x[0].lower()):
        summary_rows.append({
            "MissingCompany": company,
            "RequestedBySponsors": ", ".join(sorted(sps))
        })
    df_summary = pd.DataFrame(summary_rows)

    if show_previews:
        st.subheader("Preview")
        for sp in sponsors:
            res = sponsor_results.get(sp)
            if not res:
                continue
            st.markdown(f"**{sp}**")
            st.write("All targets with matches and scores")
            st.dataframe(res["all"], use_container_width=True)
            st.write("Included in scrape")
            st.dataframe(res["in"], use_container_width=True)
            st.write("Not included in scrape")
            st.dataframe(res["out"], use_container_width=True)
            st.markdown("---")
        st.markdown("**Summary - Missing across all sponsors**")
        st.dataframe(df_summary, use_container_width=True)

    # ---------------- Excel export ----------------
    output = io.BytesIO()
    with get_excel_writer(output) as writer:
        # One sheet per sponsor
        for sp in sponsors:
            res = sponsor_results.get(sp)
            ws_name = re.sub(r"[\\/\*\?\[\]]", "_", sp)[:31] or "Sponsor"
            res["all"].to_excel(writer, index=False, sheet_name=ws_name, startrow=1)
            ws = writer.sheets[ws_name]
            try:
                # xlsxwriter formatting if available
                wb = writer.book
                titlefmt = wb.add_format({"bold": True, "font_size": 12})
                ws.write(0, 0, "Section A - All targets with match and score", titlefmt)
                row_b = len(res["all"]) + 4
                ws.write(row_b - 1, 0, "Section B - Targets included in the scrape", titlefmt)
                res["in"].to_excel(writer, index=False, sheet_name=ws_name, startrow=row_b)
                row_c = row_b + len(res["in"]) + 3
                ws.write(row_c - 1, 0, "Section C - Targets not included in the scrape", titlefmt)
                res["out"].to_excel(writer, index=False, sheet_name=ws_name, startrow=row_c)
                ws.set_column(0, 1, 36)
                ws.set_column(2, 2, 42)
                ws.set_column(3, 4, 14)
            except Exception:
                # openpyxl minimal formatting path
                pass

        # Summary
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        try:
            ws_s = writer.sheets["Summary"]
            wb = writer.book
            titlefmt = wb.add_format({"bold": True, "font_size": 12})
            ws_s.write(0, 0, "Missing across all sponsors", titlefmt)
            ws_s.set_column(0, 0, 45)
            ws_s.set_column(1, 1, 40)
        except Exception:
            pass

    st.success("Report generated.")
    st.download_button(
        label="⬇️ Download Excel report",
        data=output.getvalue(),
        file_name="sponsor_target_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.markdown("**Tips**")
st.markdown("""
- Use the Alias map to collapse brand families, e.g.:
  - `Citibank => Citigroup`
  - `RBC Capital Markets => Royal Bank of Canada`
  - `BNP Paribas CIB => BNP Paribas`
- The brand-core match ignores generic tokens like bank, group, na, usa, cib, and regions, and strips anything in parentheses.
- Start with threshold 88-92. If you see false negatives, slide down a bit. If you see false positives, slide up.
""")
