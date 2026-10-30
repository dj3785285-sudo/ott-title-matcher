"""Cell-correction behaviour, highlighting, preservation, audit counts (mandatory tests 1-4, 9-11, 13-20)."""
import hashlib
import io

import pytest
from openpyxl import load_workbook

from core import CHECKS, MappingError, TitleMatcher, build_audit_report, process_workbook
from core.matching import HeaderSet
from tests.helpers import HEADERS, SurveyBuilder, fill_hex, layout_for, run

YELLOW, GREEN = "FFFF00", "00B050"


def out_ws(res, sheet="Survey"):
    return load_workbook(io.BytesIO(res.corrected_bytes))[sheet]


def val(res, builder, i, q, header):
    return out_ws(res, builder.sheet).cell(row=builder.row_of(i), column=builder.col_of(q, header)).value


def cell(res, builder, i, q, header):
    return out_ws(res, builder.sheet).cell(row=builder.row_of(i), column=builder.col_of(q, header))


# ---------------------------------------------------------------------------------------
# Tests 1-4 + 14/16/17: the four destination-cell rules
# ---------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "orig,expected,color,action_key",
    [
        (1, 0, YELLOW, "changed"),
        (2, 2, "", "preserved_2"),
        (0, 0, "", "already_0"),
        (None, None, "", "blank"),
        (7, 7, "", "other"),
        ("x", "x", "", "other"),
    ],
    ids=["one", "two", "zero", "blank", "other-number", "other-text"],
)
def test_exact_match_cell_rules(orig, expected, color, action_key):
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[orig, None, None, None, None])
    _, _, res = run(b)
    c = cell(res, b, 0, "Q3", "Bigg Boss")
    assert c.value == expected
    assert fill_hex(c) == color
    assert res.per_check[0][action_key] == 1
    assert res.totals["changed"] == (1 if orig == 1 else 0)


def test_text_one_is_changed_to_text_zero():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=["1", None, None, None, None])
    _, _, res = run(b)
    c = cell(res, b, 0, "Q3", "Bigg Boss")
    assert c.value == "0" and fill_hex(c) == YELLOW


def test_float_one_is_changed():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[1.0, None, None, None, None])
    _, _, res = run(b)
    assert val(res, b, 0, "Q3", "Bigg Boss") == 0


def test_formula_cell_is_never_overwritten():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[None] * 5)
    b.formulas[(b.row_of(0), b.col_of("Q3", "Bigg Boss"))] = "=1"
    _, _, res = run(b)
    assert val(res, b, 0, "Q3", "Bigg Boss") == "=1"
    assert res.totals["changed"] == 0 and res.totals["formula_skipped"] == 1


# ---------------------------------------------------------------------------------------
# Tests 5-7, 15: normalized / fuzzy -> green
# ---------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "source,header,key",
    [
        ("The Kapil Sharma Show", "Kapil Sharma Show", "normalized"),
        ("  KAPIL  SHARMA SHOW ", "Kapil Sharma Show", "normalized"),
        ("Mirzapurr", "Mirzapur", "fuzzy"),
        ("Big Boss", "Bigg Boss", "alias"),
        ("BiggBoss", "Bigg Boss", "normalized"),
    ],
)
def test_non_exact_matches_are_green(source, header, key):
    b = SurveyBuilder()
    b.add(Q2=[source], Q3=[1] * 5)
    _, _, res = run(b)
    c = cell(res, b, 0, "Q3", header)
    assert c.value == 0 and fill_hex(c) == GREEN
    assert res.per_check[0][key] == 1 and res.totals["changed"] == 1
    # the other four destination cells of that respondent stay untouched and uncoloured
    for h in HEADERS:
        if h != header:
            other = cell(res, b, 0, "Q3", h)
            assert other.value == 1 and fill_hex(other) == ""


def test_season_variation_is_green():
    b = SurveyBuilder(headers=["Bigg Boss Season 20", "Panchayat"])
    b.add(Q2=["Bigg Boss 20"], Q3=[1, 1])
    _, _, res = run(b)
    c = cell(res, b, 0, "Q3", "Bigg Boss Season 20")
    assert c.value == 0 and fill_hex(c) == GREEN


def test_different_season_not_changed():
    b = SurveyBuilder(headers=["Bigg Boss 20", "Panchayat"])
    b.add(Q2=["Bigg Boss 19"], Q3=[1, 1])
    _, _, res = run(b)
    assert res.totals["changed"] == 0 and val(res, b, 0, "Q3", "Bigg Boss 20") == 1


# ---------------------------------------------------------------------------------------
# Test 8: unrelated titles
# ---------------------------------------------------------------------------------------
def test_unrelated_titles_change_nothing():
    b = SurveyBuilder()
    b.add(Q2=["Paatal Lok", "Sacred Games", "Money Heist"], Q3=[1] * 5, Q7=[1] * 5, Q15=[1] * 5)
    _, _, res = run(b)
    assert res.totals["changed"] == 0
    assert res.audit_rows == []
    assert not res.changed_cells


# ---------------------------------------------------------------------------------------
# Test 9: several source titles -> different headers
# ---------------------------------------------------------------------------------------
def test_multiple_titles_hit_different_headers():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss", "mirzapurr", "Squid Game"], Q3=[1, 1, 1, 1, 1])
    _, _, res = run(b)
    changed = {h: val(res, b, 0, "Q3", h) for h in HEADERS}
    assert changed == {
        "Bigg Boss": 0, "Kapil Sharma Show": 1, "Mirzapur": 0, "Panchayat": 1, "Squid Game": 0,
    }
    assert fill_hex(cell(res, b, 0, "Q3", "Bigg Boss")) == YELLOW
    assert fill_hex(cell(res, b, 0, "Q3", "Mirzapur")) == GREEN
    assert fill_hex(cell(res, b, 0, "Q3", "Squid Game")) == YELLOW
    assert res.totals["changed"] == 3


# ---------------------------------------------------------------------------------------
# Test 10: several source titles -> same destination cell
# ---------------------------------------------------------------------------------------
def test_same_cell_changed_and_counted_once_exact_wins():
    b = SurveyBuilder()
    # alias title in the FIRST column, exact title in the second: exact must win -> yellow
    b.add(Q2=["Big Boss", "Bigg Boss", "BiggBoss"], Q3=[1, None, None, None, None])
    _, _, res = run(b)
    c = cell(res, b, 0, "Q3", "Bigg Boss")
    assert c.value == 0 and fill_hex(c) == YELLOW
    assert res.totals["changed"] == 1
    assert res.per_check[0]["duplicates_suppressed"] == 2
    rows = [r for r in res.audit_rows if r["Destination Cell"] == b.addr_of(0, "Q3", "Bigg Boss")]
    assert len(rows) == 3
    assert sum(r["Cell Modified"] == "Yes" for r in rows) == 1
    assert [r for r in rows if r["Cell Modified"] == "Yes"][0]["Match Type"] == "exact"
    assert all(r["Action Taken"].startswith("Duplicate") for r in rows if r["Cell Modified"] == "No")


def test_same_cell_across_checks_counted_once_lowest_check_wins_tie():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q6=["Bigg Boss"], Q14=["Bigg Boss"], Q7=[1] * 5, Q15=[1] * 5)
    _, _, res = run(b)
    # Q7 is reached by checks 2 and 3; Q15 by checks 4, 5, 6 -> two unique cells changed overall
    assert res.totals["changed"] == 2
    by_check = {r["check"]: r["changed"] for r in res.per_check}
    assert by_check == {1: 0, 2: 1, 3: 0, 4: 1, 5: 0, 6: 0}
    assert res.totals["duplicates_suppressed"] == 3  # check 3 on Q7; checks 5, 6 on Q15


def test_exact_in_later_check_beats_fuzzy_in_earlier_check():
    b = SurveyBuilder()
    b.add(Q2=["Mirzapurr"], Q6=["Mirzapur"], Q7=[None, None, 1, None, None])
    _, _, res = run(b)
    c = cell(res, b, 0, "Q7", "Mirzapur")
    assert c.value == 0 and fill_hex(c) == YELLOW  # exact from Q6 (check 3) wins over fuzzy from Q2 (check 2)
    assert {r["check"]: r["changed"] for r in res.per_check}[3] == 1


# ---------------------------------------------------------------------------------------
# Test 11: all six question-pair relationships
# ---------------------------------------------------------------------------------------
@pytest.mark.parametrize("check_no,sq,dq", CHECKS)
def test_each_of_the_six_checks_in_isolation(check_no, sq, dq):
    b = SurveyBuilder()
    b.add(**{sq: ["Panchayat"], "Q3": [1] * 5, "Q7": [1] * 5, "Q15": [1] * 5})
    _, _, res = run(b)
    # every destination reachable from this source gets exactly its Panchayat cell changed
    targets = {d for _, s, d in CHECKS if s == sq}
    assert targets, "source must feed at least one destination"
    for q in ("Q3", "Q7", "Q15"):
        for h in HEADERS:
            c = cell(res, b, 0, q, h)
            should = q in targets and h == "Panchayat"
            assert c.value == (0 if should else 1), (q, h)
            assert fill_hex(c) == (YELLOW if should else "")
    assert {r["check"]: r["changed"] for r in res.per_check}[check_no] >= 0
    # source cells are untouched
    for q in ("Q2", "Q6", "Q14"):
        for j in range(3):
            assert cell(res, b, 0, q, j).value == (["Panchayat"][0] if (q == sq and j == 0) else None)


def test_only_listed_pairs_are_checked():
    # Q3 is not a source for anything and Q2 is never compared with Q14 etc.
    assert [(s, d) for _, s, d in CHECKS] == [
        ("Q2", "Q3"), ("Q2", "Q7"), ("Q6", "Q7"), ("Q2", "Q15"), ("Q6", "Q15"), ("Q14", "Q15"),
    ]
    b = SurveyBuilder()
    b.add(Q14=["Panchayat"], Q3=[1] * 5, Q7=[1] * 5)  # Q14 only feeds Q15
    _, _, res = run(b)
    assert res.totals["changed"] == 0


# ---------------------------------------------------------------------------------------
# Test 12: different layouts all work end to end
# ---------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "kw",
    [dict(), dict(top_pad=4), dict(left_pad=0), dict(left_pad=6, top_pad=2, merged_labels=True), dict(gap=3)],
    ids=lambda kw: str(kw) or "default",
)
def test_layout_variants_end_to_end(kw):
    b = SurveyBuilder(**kw)
    b.add(Q2=["Mirzapur"], Q3=[1] * 5, Q7=[1] * 5)
    _, layout, res = run(b)
    assert layout.confident
    assert val(res, b, 0, "Q3", "Mirzapur") == 0
    assert val(res, b, 0, "Q7", "Mirzapur") == 0
    assert res.totals["changed"] == 2


def test_manual_mapping_overrides_detection():
    from core import build_layout, columns_to_spec
    b = SurveyBuilder(top_pad=2)
    b.add(Q2=["Panchayat"], Q3=[1] * 5)
    data = b.build()
    specs = {q: columns_to_spec(b.cols[q]) for q in b.cols}
    layout = build_layout("Survey", b.label_row, b.header_row, b.first_row, b.row_of(0), specs)
    res = process_workbook(data, "x.xlsx", layout, TitleMatcher())
    assert res.totals["changed"] == 1


# ---------------------------------------------------------------------------------------
# Test 13: respondent-row alignment
# ---------------------------------------------------------------------------------------
def test_only_the_matching_respondents_row_changes():
    b = SurveyBuilder()
    for i in range(6):
        b.add(Q2=["Bigg Boss"] if i == 3 else [None], Q3=[1] * 5)
    _, _, res = run(b)
    ws = out_ws(res)
    for i in range(6):
        for h in HEADERS:
            c = ws.cell(row=b.row_of(i), column=b.col_of("Q3", h))
            if i == 3 and h == "Bigg Boss":
                assert c.value == 0 and fill_hex(c) == YELLOW
            else:
                assert c.value == 1 and fill_hex(c) == ""
    assert res.audit_rows[0]["Respondent Row"] == b.row_of(3)
    assert res.audit_rows[0]["Destination Cell"] == b.addr_of(3, "Q3", "Bigg Boss")
    assert res.audit_rows[0]["Source Cell"].endswith(str(b.row_of(3)))


def test_each_respondent_uses_own_title():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    b.add(Q2=["Panchayat"], Q3=[1] * 5)
    b.add(Q2=["Squid Game", "Mirzapur"], Q3=[2, 1, 1, 1, 1])
    _, _, res = run(b)
    assert [val(res, b, 0, "Q3", h) for h in HEADERS] == [0, 1, 1, 1, 1]
    assert [val(res, b, 1, "Q3", h) for h in HEADERS] == [1, 1, 1, 0, 1]
    # headers: Bigg Boss, Kapil Sharma Show, Mirzapur, Panchayat, Squid Game
    assert [val(res, b, 2, "Q3", h) for h in HEADERS] == [2, 1, 0, 1, 0]


# ---------------------------------------------------------------------------------------
# Tests 16-18: no stray highlighting, 2 preserved, other cells/sheets untouched
# ---------------------------------------------------------------------------------------
def test_original_bytes_are_never_modified_and_everything_else_preserved():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss", "Panchayat"], Q3=[1, 5, 2, 1, "keep"], Q7=[1, 1, 1, 1, 1], Q15=[None, 2, 0, 1, 1])
    b.add(Q2=[None], Q3=[1] * 5)
    data = b.build()
    digest = hashlib.sha256(data).hexdigest()
    layout = layout_for(data)
    res = process_workbook(data, "s.xlsx", layout, TitleMatcher())
    assert hashlib.sha256(data).hexdigest() == digest  # original untouched
    assert res.verified, res.verification_issues

    wo = load_workbook(io.BytesIO(data))
    wn = load_workbook(io.BytesIO(res.corrected_bytes))
    assert wo.sheetnames == wn.sheetnames == ["Survey", "Notes"]
    changed = {c["cell"] for c in res.changed_cells}
    assert changed == {
        b.addr_of(0, "Q3", "Bigg Boss"), b.addr_of(0, "Q3", "Panchayat"),
        b.addr_of(0, "Q7", "Bigg Boss"), b.addr_of(0, "Q7", "Panchayat"),
        b.addr_of(0, "Q15", "Panchayat"),
    }
    for name in wo.sheetnames:
        for row in wo[name].iter_rows():
            for co in row:
                cn = wn[name][co.coordinate]
                if name == "Survey" and co.coordinate in changed:
                    continue
                assert cn.value == co.value, (name, co.coordinate)
                assert fill_hex(cn) == fill_hex(co), (name, co.coordinate)
    # formatting survives
    assert wn["Survey"].column_dimensions["A"].width == 17
    assert wn["Survey"].cell(row=b.header_row, column=1).font.bold
    assert fill_hex(wn["Notes"]["C3"]) == "FFC000"
    # Q15 value 2 for Panchayat... was 1 (col 4) -> changed; the cell holding 2 (Kapil) stays
    assert wn["Survey"][b.addr_of(0, "Q15", "Kapil Sharma Show")].value == 2


def test_no_highlight_when_title_matches_but_cell_not_one():
    b = SurveyBuilder(src_cols=4)
    # headers: Bigg Boss, Kapil Sharma Show, Mirzapur, Panchayat, Squid Game
    b.add(Q2=["Bigg Boss", "Kapil Sharma Show", "Mirzapur", "Panchayat"], Q3=[2, 0, None, "x", 3])
    _, _, res = run(b)
    ws = out_ws(res)
    for c in ws[b.row_of(0)]:
        assert fill_hex(c) == ""
    c1 = res.per_check[0]
    assert res.totals["changed"] == 0 and res.changed_cells == []
    assert (c1["preserved_2"], c1["already_0"], c1["blank"], c1["other"]) == (1, 1, 1, 1)
    assert [ws.cell(row=b.row_of(0), column=b.col_of("Q3", h)).value for h in HEADERS] == [2, 0, None, "x", 3]


def test_source_columns_never_modified_even_if_they_contain_ones():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss", 1, 1], Q6=[1, 1, 1], Q14=[1, 1, 1], Q3=[1] * 5)
    data, layout, res = run(b)
    for q in ("Q2", "Q6", "Q14"):
        for j in range(3):
            assert cell(res, b, 0, q, j).value == load_workbook(io.BytesIO(data))["Survey"].cell(
                row=b.row_of(0), column=b.cols[q][j]
            ).value


# ---------------------------------------------------------------------------------------
# Ambiguity: never applied, listed for review
# ---------------------------------------------------------------------------------------
def test_ambiguous_match_is_not_applied_and_listed_for_review():
    b = SurveyBuilder(headers=["Bigg Boss", "bigg boss", "Panchayat"])
    b.add(Q2=["Bigg Boss"], Q3=[1, 1, 1])
    _, _, res = run(b)
    assert res.totals["changed"] == 0
    assert res.per_check[0]["ambiguous"] == 1
    rows = [r for r in res.review_rows if r["Check"] == 1]
    assert len(rows) == 2
    assert all(r["Status"] == "Ambiguous" for r in rows)
    from openpyxl.utils import get_column_letter as L
    assert {r["Destination Header Cell"] for r in rows} == {
        f"{L(b.cols['Q3'][0])}{b.header_row}", f"{L(b.cols['Q3'][1])}{b.header_row}",
    }


def test_near_matches_listed_for_review_not_applied():
    b = SurveyBuilder(headers=["Bigg Boss 20", "Panchayat"])
    b.add(Q2=["Bigg Boss"], Q3=[1, 1])
    b.add(Q2=["Panchayatt"], Q3=[1, 1])  # fuzzy hit (>=92) -> applied
    _, _, res = run(b)
    assert val(res, b, 0, "Q3", "Bigg Boss 20") == 1
    review = [r for r in res.review_rows if r["Source Title"] == "Bigg Boss"]
    assert review and "one side only" in review[0]["Reason for Manual Review"]
    assert val(res, b, 1, "Q3", "Panchayat") == 0


def test_review_rows_aggregate_occurrences():
    b = SurveyBuilder(headers=["Bigg Boss 20", "Panchayat"])
    for _ in range(3):
        b.add(Q2=["Bigg Boss"], Q3=[1, 1])
    _, _, res = run(b)
    rows = [r for r in res.review_rows if r["Check"] == 1]
    assert len(rows) == 1 and rows[0]["Occurrences"] == 3


# ---------------------------------------------------------------------------------------
# Validation errors
# ---------------------------------------------------------------------------------------
def test_process_refuses_invalid_mapping():
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    data = b.build()
    layout = layout_for(data)
    layout.questions["Q15"].columns = []
    with pytest.raises(MappingError) as e:
        process_workbook(data, "x.xlsx", layout, TitleMatcher())
    assert "Q15" in str(e.value)


# ---------------------------------------------------------------------------------------
# Test 19: audit-report counts   |   Test 20: persistence after save/reopen
# ---------------------------------------------------------------------------------------
def scenario():
    b = SurveyBuilder()
    # r0: exact 1 (Q3), normalized-ish alias 1 (Q3), 2 preserved
    b.add(Q2=["Bigg Boss", "Mirzapurr", "Panchayat"], Q3=[1, None, 1, 2, None])
    # r1: title matches but cell is 0 / blank
    b.add(Q2=["Kapil Sharma Show", "Squid Game"], Q3=[None, 0, None, None, None])
    # r2: unrelated title; destination from Q6
    b.add(Q2=["Unknown Show"], Q6=["The Kapil Sharma Show"], Q7=[None, 1, None, None, 5])
    return b


def test_audit_counts_are_correct():
    b = scenario()
    _, _, res = run(b)
    t = res.totals
    # titles checked: each Q2 title is checked three times (checks 1, 2, 4), each Q6 title twice (3, 5)
    assert t["titles_checked"] == 3 * 3 + 2 * 3 + 1 * 3 + 1 * 2 == 20
    by = {r["check"]: r for r in res.per_check}
    # check 1: Q2 -> Q3 : Bigg Boss (exact, 1->0), Mirzapurr (fuzzy, 1->0), Panchayat (exact, cell 2),
    #                      Kapil Sharma Show (exact, blank cell), Squid Game (exact, cell 0)
    assert by[1]["titles_checked"] == 6
    assert (by[1]["exact"], by[1]["fuzzy"], by[1]["normalized"]) == (4, 1, 0)
    assert (by[1]["changed"], by[1]["preserved_2"], by[1]["already_0"], by[1]["blank"]) == (2, 1, 1, 1)
    # check 2 / 4: Q2 -> Q7 / Q15: respondent 0 and 1 have blank destination cells everywhere
    assert by[2]["changed"] == 0 and by[2]["blank"] == 5 and by[4]["blank"] == 5
    # check 3: Q6 -> Q7 : "The Kapil Sharma Show" normalized -> cell 1 -> 0 (green)
    assert (by[3]["titles_checked"], by[3]["normalized"], by[3]["changed"]) == (1, 1, 1)
    assert by[5]["titles_checked"] == 1 and by[5]["blank"] == 1 and by[5]["duplicates_suppressed"] == 0
    assert by[6]["titles_checked"] == 0
    assert t["changed"] == 3
    assert t["changed"] == sum(r["changed"] for r in res.per_check)
    assert len(res.changed_cells) == 3
    # every audit row marked as modified corresponds to a changed cell, one-to-one
    modified = [r for r in res.audit_rows if r["Cell Modified"] == "Yes"]
    assert len(modified) == 3
    assert {r["Destination Cell"] for r in modified} == {c["cell"] for c in res.changed_cells}
    assert sorted(r["Highlight Color"] for r in modified) == ["Green", "Green", "Yellow"]
    # every changed row has original 1 and final 0
    assert all(r["Original Value"] == 1 and r["Final Value"] == 0 for r in modified)


def test_audit_report_workbook_structure_and_values(tmp_path):
    b = scenario()
    _, layout, res = run(b)
    path = tmp_path / "audit.xlsx"
    path.write_bytes(build_audit_report(res))
    wb = load_workbook(path)
    assert wb.sheetnames == ["Summary", "Audit Log", "Review Matches"]
    ws = wb["Summary"]
    flat = {str(c.value) for row in ws.iter_rows() for c in row if c.value is not None}
    for needle in (
        "survey.xlsx", "92%", "Yellow = Exact match, changed from 1 to 0.",
        "Cells changed from 1 to 0", "Cells preserved because value was 2",
        "Statistics per question-pair check", "Survey",
    ):
        assert needle in flat, needle
    log = wb["Audit Log"]
    headers = [c.value for c in log[1]]
    for h in (
        "Worksheet", "Source Question", "Destination Question", "Respondent Row", "Source Title",
        "Source Cell", "Destination Title Header", "Destination Header Cell", "Destination Cell",
        "Original Value", "Final Value", "Match Type", "Match Confidence", "Action Taken", "Highlight Color",
    ):
        assert h in headers, h
    assert log.max_row - 1 == len(res.audit_rows)
    colors = {log.cell(row=r, column=headers.index("Highlight Color") + 1).value for r in range(2, log.max_row + 1)}
    assert {"Yellow", "Green"} <= colors
    # colour cells in the audit log carry the fill
    for r in range(2, log.max_row + 1):
        c = log.cell(row=r, column=headers.index("Highlight Color") + 1)
        if c.value == "Yellow":
            assert fill_hex(c) == YELLOW
        if c.value == "Green":
            assert fill_hex(c) == GREEN
    assert wb["Review Matches"].max_row >= 2


def test_summary_totals_match_per_check_sums():
    _, _, res = run(scenario())
    for key, total in res.totals.items():
        if key in ("title_matches", "review_total"):
            continue
        assert total == sum(r[key] for r in res.per_check), key


def test_saved_corrected_workbook_reopens_with_changes_and_fills(tmp_path):
    b = scenario()
    _, _, res = run(b)
    path = tmp_path / "corrected.xlsx"
    path.write_bytes(res.corrected_bytes)
    ws = load_workbook(path)["Survey"]  # fresh load from disk
    exact = ws[b.addr_of(0, "Q3", "Bigg Boss")]
    fuzzy = ws[b.addr_of(0, "Q3", "Mirzapur")]
    norm = ws[b.addr_of(2, "Q7", "Kapil Sharma Show")]
    assert (exact.value, fuzzy.value, norm.value) == (0, 0, 0)
    assert (fill_hex(exact), fill_hex(fuzzy), fill_hex(norm)) == (YELLOW, GREEN, GREEN)
    assert ws[b.addr_of(0, "Q3", "Panchayat")].value == 2 and fill_hex(ws[b.addr_of(0, "Q3", "Panchayat")]) == ""
    assert ws[b.addr_of(2, "Q7", "Squid Game")].value == 5
    assert res.verified and res.verification_issues == []


def test_verifier_detects_tampering():
    from core.processor import verify_output
    b = SurveyBuilder()
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    data, _, res = run(b)
    wb = load_workbook(io.BytesIO(res.corrected_bytes))
    wb["Survey"].cell(row=b.row_of(0), column=b.col_of("Q3", "Panchayat")).value = 0  # unintended change
    buf = io.BytesIO()
    wb.save(buf)
    issues = verify_output(data, buf.getvalue(), {("Survey", b.row_of(0), b.col_of("Q3", "Bigg Boss")): "Yellow"})
    assert any("unexpectedly" in i or "highlight" in i for i in issues)


def test_invalid_file_gives_clear_error():
    from core import WorkbookError, load_snapshot
    with pytest.raises(WorkbookError):
        load_snapshot(b"this is not excel", "Sheet1")
    with pytest.raises(WorkbookError):
        load_snapshot(b"PK\x03\x04garbage", "Sheet1")


def test_preview_is_a_dry_run_with_classifications():
    from core import preview_matches
    b = scenario()
    data = b.build()
    layout = layout_for(data)
    prev = preview_matches(data, layout, TitleMatcher())
    kinds = {m["Classification"] for m in prev["matches"]}
    assert {"exact", "fuzzy", "normalized"} <= kinds
    assert prev["would_change"] == 3
    colours = {m["Classification"]: m["Colour if changed"] for m in prev["matches"]}
    assert colours["exact"] == "Yellow" and colours["fuzzy"] == "Green"
