# OTT Title Matcher & Excel Auto-Corrector

Browser-based tool (Streamlit) that reads survey workbooks, finds OTT show/movie titles you listed in one
question, and in the later questions changes that title's `1` to `0` (keeps `2`), highlights each changed cell,
and writes an audit report. Your original file is never modified.

| Check | Source (never modified) | Destination (1 → 0) |
|---|---|---|
| 1 | Q2 | Q3 |
| 2 | Q2 | Q7 |
| 3 | Q6 | Q7 |
| 4 | Q2 | Q15 |
| 5 | Q6 | Q15 |
| 6 | Q14 | Q15 |

### Question numbers differ per survey? Set the pairs yourself
The six checks above are only the **default**. In the app sidebar, **Question pairs to check**, type one
`source -> destination` pair per line using *this survey's* numbers, for example:

```
Q4 -> Q9
Q4, Q5 -> Q12        # several sources for one destination
Q20 -> Q21
```
The tool then looks for exactly those questions in the sheet, maps their columns, and the audit report records the
pairs used. A question must be only a source or only a destination (sources are never modified). Try
`sample_data/sample_survey_renumbered.xlsx` (pairs listed in `sample_data/README_renumbered.txt`).

**Rules.** Destination cell `1` → `0`; `2` kept; `0`, blank, formulas and any other value untouched.
**Yellow** (`FFFF00`) = exact match (ignoring case/outer spaces). **Green** (`00B050`) = normalized, recognised
variant (alias) or fuzzy match. Only cells actually changed are coloured.

## Use it on your Mac without installing anything locally (Streamlit Community Cloud)

1. Create a free GitHub account (github.com) and a new **private or public** repository, e.g. `ott-title-matcher`.
2. In the repo click **Add file → Upload files**, drag in the *contents* of this folder (`app.py`, `core/`,
   `requirements.txt`, `.streamlit/`, `tests/`, `sample_data/`, …) and commit.
3. Go to **share.streamlit.io**, sign in with GitHub, click **Create app → Deploy a public app from GitHub**.
4. Choose your repo, branch `main`, main file `app.py`, then **Deploy**. After a couple of minutes you get a
   URL like `https://your-name-ott-title-matcher.streamlit.app` that opens in Safari/Chrome.
5. Upload `sample_data/sample_survey.xlsx` first to see how it works.

Privacy: uploaded workbooks are processed in memory on the host; use a private repo/app if the survey is sensitive.
(Streamlit Community Cloud apps can be restricted to invited viewers in the app's Share settings.)

## Using the app
1. **Upload** the `.xlsx`.  2. **Inspect** (after setting the question pairs in the sidebar) – detected sheets, question columns, header/respondent rows.
3. **Confirm** – edit anything (Excel notation `BK:BQ` or `BK,BL,BM`), set the fuzzy threshold (default 92 %),
tick the confirmation box. The tool never silently guesses: unresolved questions block processing.
4. **Preview** – dry run showing exact / normalized / alias / fuzzy matches and the manual-review list.
5. **Process**, 6. **Results** (stats + colour legend), 7. **Download** corrected workbook and audit report.

## How matching works
Order: exact → normalized (case, punctuation, spacing, joined words, "The", season labels `Season 20`/`S20`/`20th Season`)
→ recognised variants (`core/aliases.py`) → conservative fuzzy (RapidFuzz ratio ≥ threshold).
Safeguards: different season numbers never match; a number on one side only (`Bigg Boss` vs `Bigg Boss 20`)
goes to review; titles under 7 letters need ≥ 96 %, under 5 letters are never fuzzy-matched; equally good
candidates are *ambiguous* and are only listed in **Review Matches**, never applied.
Add aliases/exceptions permanently in `core/aliases.py` or per session in the sidebar JSON box.

If several source titles/checks reach the same cell it is changed and counted once; winner = exact > normalized >
alias > fuzzy, then higher score, then lower check number (also stated in the audit report).

## Audit report (3 sheets)
**Summary** (file, time, sheet, mappings, threshold, totals + per-check stats, legend), **Audit Log** (every match,
with title match vs. actual modification distinguished), **Review Matches**.

## Known limits
* OpenPyXL cannot keep charts, images, slicers, macros, sparklines; the app warns when it sees them. Formulas,
  styles, widths, merges and other sheets are preserved. Verify the output in Excel on a copy first.
* Only `.xlsx` (save `.xls` as `.xlsx` first). Several titles in one source cell may be separated by new line, `;` or `|`
  (not commas, since titles may contain them).
* Recommended size: up to ~5,000 respondents × 200 columns (≈ 25 s, ~0.5 GB RAM). Streamlit Cloud's free tier has
  limited memory, so split very large files.

## Run / test locally (optional)
```
pip install -r requirements.txt
streamlit run app.py
python -m pytest
python make_samples.py     # regenerates sample_data/
```
