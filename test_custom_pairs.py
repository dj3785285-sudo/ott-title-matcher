"""Question numbers differ per survey: pairs and detection must be configurable."""
import io

import pytest
from openpyxl import load_workbook

from core import (
    DEFAULT_CHECKS, MappingError, TitleMatcher, build_audit_report, build_layout, parse_pairs,
    process_workbook, validate_layout, load_snapshot, columns_to_spec, preview_matches,
)
from core.pairs import destinations_of, sources_of
from tests.helpers import HEADERS, SurveyBuilder, fill_hex, layout_for, run

# a survey that uses completely different question numbers
RENAME = {"Q2": "Q4", "Q3": "Q9", "Q6": "Q5", "Q7": "Q12", "Q14": "Q20", "Q15": "Q21"}
PAIRS = "Q4 -> Q9\nQ4 -> Q12\nQ5 -> Q12\nQ4 -> Q21\nQ5 -> Q21\nQ20 -> Q21"


def cell(res, b, i, internal_q, header):
    ws = load_workbook(io.BytesIO(res.corrected_bytes))[b.sheet]
    return ws.cell(row=b.row_of(i), column=b.col_of(internal_q, header))


# ------------------------------------------------------------------ parsing
def test_default_checks_unchanged():
    assert [(s, d) for _, s, d in DEFAULT_CHECKS] == [
        ("Q2", "Q3"), ("Q2", "Q7"), ("Q6", "Q7"), ("Q2", "Q15"), ("Q6", "Q15"), ("Q14", "Q15"),
    ]


@pytest.mark.parametrize(
    "text,expected",
    [
        ("Q4 -> Q9", [("Q4", "Q9")]),
        ("q4 → q9", [("Q4", "Q9")]),
        ("4 > 9", [("Q4", "Q9")]),
        ("Question 04 => Question 09", [("Q4", "Q9")]),
        ("Q4 -> Q9\n\n# comment\nQ5 -> Q9  # trailing", [("Q4", "Q9"), ("Q5", "Q9")]),
        ("Q4, Q5 -> Q12", [("Q4", "Q12"), ("Q5", "Q12")]),
        ("Q4 and Q5 -> Q12", [("Q4", "Q12"), ("Q5", "Q12")]),
        ("Q4 -> Q9\nQ4 -> Q9", [("Q4", "Q9")]),  # duplicates dropped
    ],
)
def test_parse_pairs_formats(text, expected):
    got = parse_pairs(text)
    assert [(s, d) for _, s, d in got] == expected
    assert [n for n, _, _ in got] == list(range(1, len(expected) + 1))


@pytest.mark.parametrize(
    "text,needle",
    [
        ("", "at least one"),
        ("Q4", "write it as"),
        ("Q4 -> ", "write it as"),
        ("Q4 -> Q4", "itself"),
        ("Qx -> Q9", "not a question number"),
        ("Q4 -> Q9\nQ9 -> Q12", "both as a source and as a destination"),
    ],
)
def test_parse_pairs_errors(text, needle):
    with pytest.raises(MappingError) as e:
        parse_pairs(text)
    assert needle in str(e.value)


def test_roles():
    checks = parse_pairs(PAIRS)
    assert sources_of(checks) == ["Q4", "Q5", "Q20"]
    assert destinations_of(checks) == ["Q9", "Q12", "Q21"]


# ------------------------------------------------------------------ detection
def test_detection_follows_the_configured_pairs():
    checks = parse_pairs(PAIRS)
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    layout = layout_for(b.build(), checks=checks)
    assert layout.confident, layout.issues
    for internal, label in RENAME.items():
        assert layout.cols(label) == b.cols[internal], label


def test_default_pairs_on_a_renumbered_survey_report_missing_questions():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    layout = layout_for(b.build())  # default Q2/Q3/Q6/Q7/Q14/Q15
    assert not layout.confident
    assert any("Q2" in i and "question-pair list" in i for i in layout.issues)
    errors, _ = validate_layout(load_snapshot(b.build(), "Survey"), layout)
    assert errors  # blocked, never silently guessed


def test_extra_unrelated_questions_in_the_sheet_are_detected_but_not_required():
    # Q4, Q5, Q9 exist as well as the required ones; only the pair's questions matter
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5, Q7=[1] * 5)
    checks = parse_pairs("Q4 -> Q9")
    layout = layout_for(b.build(), checks=checks)
    assert layout.confident
    assert layout.cols("Q4") and layout.cols("Q9") and layout.cols("Q12")  # also detected


# ------------------------------------------------------------------ processing
def test_renumbered_survey_gives_same_result_as_the_default_numbering():
    def fill(bld):
        bld.add(Q2=["Bigg Boss", "Mirzapurr"], Q6=["The Kapil Sharma Show"], Q14=["Panchayat"],
                Q3=[1, 1, 1, 2, 0], Q7=[1, 1, 1, 1, 1], Q15=[1, 1, 1, 1, 1])
        bld.add(Q2=["Squid Game"], Q3=[1] * 5, Q7=[None, 2, 1, 1, 1], Q15=[1] * 5)

    base = SurveyBuilder(); fill(base)
    _, _, r_base = run(base)
    new = SurveyBuilder(rename=RENAME); fill(new)
    _, _, r_new = run(new, checks=parse_pairs(PAIRS))
    assert r_new.totals == r_base.totals
    assert [c["cell"] for c in r_new.changed_cells] == [c["cell"] for c in r_base.changed_cells]
    assert [c["color"] for c in r_new.changed_cells] == [c["color"] for c in r_base.changed_cells]
    assert r_new.verified


def test_single_custom_pair_touches_only_its_destination():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Panchayat"], Q3=[1] * 5, Q7=[1] * 5, Q15=[1] * 5)
    _, _, res = run(b, checks=parse_pairs("Q4 -> Q12"))  # source Q4 (=Q2), destination Q12 (=Q7)
    assert cell(res, b, 0, "Q7", "Panchayat").value == 0
    assert fill_hex(cell(res, b, 0, "Q7", "Panchayat")) == "FFFF00"
    for q in ("Q3", "Q15"):
        assert cell(res, b, 0, q, "Panchayat").value == 1
    assert res.totals["changed"] == 1 and [r["check"] for r in res.per_check] == [1]


def test_several_sources_for_one_destination():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss"], Q6=["Mirzapur"], Q3=[1] * 5, Q7=[1] * 5)
    _, _, res = run(b, checks=parse_pairs("Q4, Q5 -> Q12"))
    assert cell(res, b, 0, "Q7", "Bigg Boss").value == 0
    assert cell(res, b, 0, "Q7", "Mirzapur").value == 0
    assert cell(res, b, 0, "Q3", "Bigg Boss").value == 1  # Q9 is not a destination here
    assert res.totals["changed"] == 2


def test_source_columns_still_protected_with_custom_numbers():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss", 1, 1], Q3=[1] * 5)
    data, layout, res = run(b, checks=parse_pairs("Q4 -> Q9"))
    ws = load_workbook(io.BytesIO(res.corrected_bytes))["Survey"]
    assert [ws.cell(row=b.row_of(0), column=c).value for c in b.cols["Q2"]] == ["Bigg Boss", 1, 1]


def test_overlap_and_role_conflicts_blocked_for_custom_pairs():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Bigg Boss"], Q3=[1] * 5)
    data = b.build()
    specs = {RENAME[i]: columns_to_spec(b.cols[i]) for i in RENAME}
    specs["Q9"] = columns_to_spec(b.cols["Q2"] + b.cols["Q3"])  # destination swallows source
    layout = build_layout("Survey", b.label_row, b.header_row, b.first_row, b.row_of(0), specs)
    with pytest.raises(MappingError) as e:
        process_workbook(data, "x.xlsx", layout, TitleMatcher(), checks=parse_pairs("Q4 -> Q9"))
    assert "mapped to both" in str(e.value)


def test_manual_mapping_with_custom_question_numbers():
    b = SurveyBuilder(top_pad=2, rename=RENAME)
    b.add(Q2=["Panchayat"], Q3=[1] * 5)
    data = b.build()
    specs = {"Q4": columns_to_spec(b.cols["Q2"]), "Q9": columns_to_spec(b.cols["Q3"])}
    layout = build_layout("Survey", b.label_row, b.header_row, b.first_row, b.row_of(0), specs)
    res = process_workbook(data, "x.xlsx", layout, TitleMatcher(), checks=parse_pairs("Q4 -> Q9"))
    assert res.totals["changed"] == 1


def test_preview_uses_custom_pairs():
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Panchayat"], Q3=[1] * 5, Q7=[1] * 5)
    data = b.build()
    checks = parse_pairs("Q4 -> Q9")
    layout = layout_for(data, checks=checks)
    pv = preview_matches(data, layout, TitleMatcher(), checks=checks)
    assert pv["would_change"] == 1 and pv["matches"][0]["Check"] == "1: Q4 -> Q9"


def test_audit_report_lists_the_pairs_used(tmp_path):
    b = SurveyBuilder(rename=RENAME)
    b.add(Q2=["Panchayat"], Q3=[1] * 5)
    _, _, res = run(b, checks=parse_pairs("Q4 -> Q9\nQ4 -> Q12"))
    p = tmp_path / "a.xlsx"
    p.write_bytes(build_audit_report(res))
    ws = load_workbook(p)["Summary"]
    flat = {str(c.value) for row in ws.iter_rows() for c in row if c.value}
    assert "Q4 (source) -> Q9 (destination)" in flat and "Q4 (source) -> Q12 (destination)" in flat
    log = load_workbook(p)["Audit Log"]
    assert {log.cell(row=r, column=3).value for r in range(2, log.max_row + 1)} == {"Q4"}
