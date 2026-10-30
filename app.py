"""OTT Title Matcher - Streamlit web app.

Workflow: 1 Upload -> 2 Inspect -> 3 Confirm -> 4 Preview -> 5 Process -> 6 Results -> 7 Download
Run locally:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from core.inspector import QuestionMapping
from core import (
    DEFAULT_CHECKS,
    DEFAULT_THRESHOLD,
    MappingError,
    OTTError,
    TitleMatcher,
    build_audit_report,
    build_layout,
    columns_to_spec,
    default_alias_table,
    detect_unsupported_features,
    inspect_snapshots,
    load_all_snapshots,
    load_custom_rules,
    preview_matches,
    process_workbook,
    scorer_name,
    validate_layout,
    load_snapshot,
    pairs_to_text,
    parse_pairs,
)
from core.inspector import best_sheet
from core.pairs import destinations_of, questions_of, sources_of

st.set_page_config(page_title="OTT Title Matcher", page_icon="📺", layout="wide")

MAX_MB = 50
YELLOW, GREEN = "#FFFF00", "#00B050"


def legend():
    st.markdown(
        f"""
<div style="display:flex;gap:24px;flex-wrap:wrap;margin:6px 0 12px 0">
  <span><span style="background:{YELLOW};padding:2px 14px;border:1px solid #999;border-radius:4px">&nbsp;</span>
  &nbsp;<b>Yellow</b> = exact match, changed from 1 to 0</span>
  <span><span style="background:{GREEN};padding:2px 14px;border:1px solid #999;border-radius:4px">&nbsp;</span>
  &nbsp;<b>Green</b> = normalized or fuzzy title match, changed from 1 to 0</span>
</div>""",
        unsafe_allow_html=True,
    )


def reset_downstream():
    for k in ("result", "preview", "audit_bytes"):
        st.session_state.pop(k, None)


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("Settings")
    st.markdown("**Question pairs to check**")
    st.caption(
        "One pair per line: *source -> destination*. Source titles are looked up in the destination "
        "question and that title's 1 becomes 0. Change the numbers to match each survey. "
        "Several sources are allowed: `Q4, Q5 -> Q12`."
    )
    pairs_text = st.text_area(
        "Question pairs", value=pairs_to_text(DEFAULT_CHECKS), height=170,
        label_visibility="collapsed", key="pairs_text", on_change=reset_downstream,
    )
    st.button(
        "Reset to default pairs (Q2→Q3, Q2→Q7 …)",
        on_click=lambda: st.session_state.update(pairs_text=pairs_to_text(DEFAULT_CHECKS)),
    )
    threshold = st.slider(
        "Fuzzy-match threshold (%)", 80, 100, int(DEFAULT_THRESHOLD),
        help="Minimum similarity for a minor-spelling match. Short titles need a stricter score automatically.",
        on_change=reset_downstream,
    )
    with st.expander("Custom aliases / exceptions (optional)"):
        st.caption(
            'JSON, e.g. {"aliases": {"Bigg Boss": ["BB"]}, "exceptions": [["Bigg Boss", "Bigg Boss OTT"]]}'
        )
        rules_text = st.text_area("Rules", value="", height=140, label_visibility="collapsed",
                                  on_change=reset_downstream)
    st.caption(f"Similarity engine: {scorer_name()}")
    st.markdown("**Colour legend**")
    legend()

try:
    checks = parse_pairs(pairs_text)
except OTTError as exc:
    st.title("📺 OTT Title Matcher & Excel Auto-Corrector")
    for e in getattr(exc, "errors", [str(exc)]):
        st.error(f"Question pairs: {e}")
    st.stop()
SRC_QS = sources_of(checks)
USED_QS = questions_of(checks)
pairs_key = hashlib.sha1(pairs_text.strip().encode()).hexdigest()[:6]

st.title("📺 OTT Title Matcher & Excel Auto-Corrector")
st.write(
    "Finds OTT titles you listed in one question and sets the same title's `1` to `0` in later questions "
    "(`2` is always kept). Your original file is never changed. "
    "**Question numbers differ per survey - set the pairs in the sidebar before you start.**"
)

# ------------------------------------------------------------------ step 1
st.subheader("Step 1 - Upload")
up = st.file_uploader("Excel workbook (.xlsx)", type=["xlsx"])
if up is None:
    st.info("Upload a survey workbook to begin.")
    st.stop()

data = up.getvalue()
if len(data) > MAX_MB * 1024 * 1024:
    st.error(f"The file is larger than {MAX_MB} MB. Please split it or remove unneeded sheets.")
    st.stop()

file_id = hashlib.sha256(data).hexdigest()
def clear_file_state():
    """Forget everything about the previous file but keep the user's question pairs."""
    keep = st.session_state.get("pairs_text")
    st.session_state.clear()
    if keep is not None:
        st.session_state["pairs_text"] = keep


if st.session_state.get("file_id") != file_id:
    clear_file_state()
    st.session_state["file_id"] = file_id
    try:
        with st.spinner("Reading workbook..."):
            snaps = load_all_snapshots(data)
        st.session_state["snaps"] = snaps
        st.session_state["sheets"] = list(snaps)
        st.session_state["features"] = detect_unsupported_features(data)
    except OTTError as exc:
        clear_file_state()
        st.error(str(exc))
        st.stop()

# detection depends on the question pairs, so it is recomputed (cheaply) whenever they change
if st.session_state.get("layouts_for") != pairs_key:
    st.session_state["layouts"] = inspect_snapshots(st.session_state["snaps"], checks)
    st.session_state["default_sheet"] = best_sheet(st.session_state["layouts"])
    st.session_state["layouts_for"] = pairs_key
layouts = st.session_state["layouts"]
for w in st.session_state["features"]:
    st.warning(w)

# ------------------------------------------------------------------ step 2
st.subheader("Step 2 - Inspect")
sheets = st.session_state["sheets"]
sheet = st.selectbox(
    "Worksheet to process", sheets, index=sheets.index(st.session_state["default_sheet"]),
    on_change=reset_downstream,
)
st.caption("Worksheets found: " + ", ".join(f"**{s}**" for s in sheets))
det = layouts[sheet]

if det.confident:
    st.success(f"All required questions ({', '.join(USED_QS)}) were detected. Please still check the mapping below.")
else:
    st.warning("The layout could not be detected with confidence. Check and complete the mapping below.")
    for issue in det.issues:
        st.write(f"- {issue}")

rows = []
for q in sorted(set(USED_QS) | set(det.questions), key=lambda x: int(x[1:])):
    m = det.questions.get(q) or QuestionMapping(q)
    rows.append(
        {
            "Question": q,
            "Role": "Source" if q in SRC_QS else ("Destination" if q in USED_QS else "not used"),
            "Detected columns": m.spec or "-",
            "Label cell": m.label_cell or "-",
            "Merged label": "yes" if m.merged else "",
        }
    )
st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Question-label row", det.label_row)
c2.metric("Title-header row", det.header_row)
c3.metric("First respondent row", det.first_row)
c4.metric("Last respondent row", det.last_row)

# ------------------------------------------------------------------ step 3
st.subheader("Step 3 - Confirm mapping")
st.caption("Use Excel notation: `BK:BQ` or `BK,BL,BM` (or a mix). Edit anything that looks wrong.")
key = f"{file_id[:8]}-{sheet}-{pairs_key}"
r1, r2, r3, r4 = st.columns(4)
label_row = r1.number_input("Question-label row", 1, 100000, int(det.label_row), key=f"lr-{key}")
header_row = r2.number_input("Title-header row", 1, 100000, int(det.header_row), key=f"hr-{key}")
first_row = r3.number_input("First respondent row", 1, 1000000, int(det.first_row), key=f"fr-{key}")
last_row = r4.number_input("Last respondent row", 1, 1000000, int(det.last_row), key=f"ls-{key}")

specs = {}
cols = st.columns(3)
for i, q in enumerate(USED_QS):
    role = "source" if q in SRC_QS else "destination"
    guess = det.questions.get(q)
    specs[q] = cols[i % 3].text_input(
        f"{q} columns ({role})", value=guess.spec if guess else "", key=f"sp-{q}-{key}"
    )

confirmed = st.checkbox("I have checked the mapping above and it is correct", key=f"ok-{key}")

try:
    layout = build_layout(sheet, label_row, header_row, first_row, last_row, specs)
    snap_key = f"snap-{key}"
    if snap_key not in st.session_state:
        st.session_state[snap_key] = load_snapshot(data, sheet)
    snap = st.session_state[snap_key]
    errors, warnings = validate_layout(snap, layout, checks)
except OTTError as exc:
    errors = list(getattr(exc, "errors", [str(exc)]))
    warnings, layout = [], None

for e in errors:
    st.error(e)
for w in warnings:
    st.warning(w)

try:
    custom_alias, custom_exc = load_custom_rules(rules_text)
except ValueError as exc:
    st.error(str(exc))
    st.stop()


def make_matcher() -> TitleMatcher:
    table = default_alias_table()
    for k, v in custom_alias.items():
        table.add(k, v)
    return TitleMatcher(threshold=threshold, aliases=table, exceptions=custom_exc)


ready = layout is not None and not errors and confirmed
if not ready:
    st.info("Fix any errors and tick the confirmation box to enable Preview and Process.")
    st.stop()

with st.expander(f"Question pairs being checked ({len(checks)})", expanded=True):
    st.table(pd.DataFrame(
        [{"Check": n, "Source (never modified)": s, "Destination (1 -> 0)": d} for n, s, d in checks]
    ))

# ------------------------------------------------------------------ step 4
st.subheader("Step 4 - Preview (nothing is changed yet)")
if st.button("Preview matches"):
    try:
        with st.spinner("Matching titles..."):
            st.session_state["preview"] = preview_matches(data, layout, make_matcher(), checks)
    except OTTError as exc:
        st.error(str(exc))
if "preview" in st.session_state:
    pv = st.session_state["preview"]
    a, b, c = st.columns(3)
    a.metric("Cells that would change 1 to 0", pv["would_change"])
    b.metric("Title matches", pv["totals"]["title_matches"])
    c.metric("Need manual review", pv["totals"]["review_total"])
    st.markdown(
        "**Classification:** *exact* -> yellow; *normalized*, *alias* (recognised variant) and *fuzzy* -> green; "
        "*ambiguous / near match* -> listed for review and never applied."
    )
    if pv["matches"]:
        st.dataframe(pd.DataFrame(pv["matches"]).head(500), hide_index=True, use_container_width=True)
        if len(pv["matches"]) > 500:
            st.caption("Showing the first 500 distinct matches.")
    else:
        st.info("No title matches found with the current mapping.")
    if pv["review"]:
        st.markdown("**Manual-review list**")
        st.dataframe(pd.DataFrame(pv["review"]).head(500), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ step 5
st.subheader("Step 5 - Process")
if st.button("Run all six checks and create corrected workbook", type="primary"):
    try:
        with st.spinner("Applying corrections and verifying the saved file..."):
            res = process_workbook(data, up.name, layout, make_matcher(), checks=checks)
            st.session_state["result"] = res
            st.session_state["audit_bytes"] = build_audit_report(res)
    except OTTError as exc:
        st.error(str(exc))
    except Exception as exc:  # unexpected - show something actionable
        st.error(f"The workbook could not be processed: {exc}")

# ------------------------------------------------------------------ step 6 + 7
res = st.session_state.get("result")
if res is not None:
    st.subheader("Step 6 - Results")
    if res.verified:
        st.success("Verified: the saved corrected workbook was re-opened and only the intended cells differ.")
    else:
        st.error("Verification found problems - do not use the corrected file:")
        for i in res.verification_issues:
            st.write(f"- {i}")
    for w in res.feature_warnings + res.layout_warnings:
        st.warning(w)
    legend()
    t = res.totals
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cells changed 1 to 0", t["changed"])
    m2.metric("Preserved (value 2)", t["preserved_2"])
    m3.metric("Already 0", t["already_0"])
    m4.metric("Blank skipped", t["blank"])
    m5, m6, m7, m8 = st.columns(4)
    m5.metric("Other values skipped", t["other"])
    m6.metric("Source titles checked", t["titles_checked"])
    m7.metric("Ambiguous (review)", t["ambiguous"])
    m8.metric("Near matches (review)", t["near_match"])
    n1, n2, n3, n4 = st.columns(4)
    n1.metric("Exact matches", t["exact"])
    n2.metric("Normalized matches", t["normalized"])
    n3.metric("Variant (alias) matches", t["alias"])
    n4.metric("Fuzzy matches", t["fuzzy"])
    st.markdown("**Per question-pair check**")
    df = pd.DataFrame(res.per_check).rename(columns={
        "check": "Check", "source": "Source", "destination": "Destination",
        "titles_checked": "Titles checked", "exact": "Exact", "normalized": "Normalized",
        "alias": "Alias", "fuzzy": "Fuzzy", "changed": "Changed 1→0", "preserved_2": "Kept 2",
        "already_0": "Already 0", "blank": "Blank", "other": "Other",
        "duplicates_suppressed": "Duplicates", "ambiguous": "Ambiguous", "near_match": "Near match",
    })
    keep = ["Check", "Source", "Destination", "Titles checked", "Exact", "Normalized", "Alias", "Fuzzy",
            "Changed 1→0", "Kept 2", "Already 0", "Blank", "Other", "Duplicates", "Ambiguous", "Near match"]
    st.dataframe(df[keep], hide_index=True, use_container_width=True)
    st.caption(
        "If several source titles or checks hit the same cell it is changed and counted once; "
        "exact beats normalized, normalized beats alias, alias beats fuzzy."
    )

    st.subheader("Step 7 - Download")
    stem = re.sub(r"\.xlsx$", "", up.name, flags=re.I)
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Corrected workbook", res.corrected_bytes, f"{stem}_corrected.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
        disabled=not res.verified,
    )
    d2.download_button(
        "⬇️ Audit report", st.session_state["audit_bytes"], f"{stem}_audit_report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
