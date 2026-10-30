"""Title-matching engine (requirements 1, 5-8 of the mandatory test list)."""
import pytest

from core.matching import (
    ALIAS, EXACT, FUZZY, NONE, NORMALIZED, REVIEW, HeaderSet, TitleMatcher, similarity,
)
from core.normalize import normalize_title


@pytest.fixture
def m():
    return TitleMatcher(threshold=92)


def kind(m, a, b):
    return m.compare(a, b).type


# ---- exact -----------------------------------------------------------------
@pytest.mark.parametrize("a,b", [("Bigg Boss", "Bigg Boss"), ("bigg boss", "BIGG BOSS"), ("  Bigg Boss ", "Bigg Boss")])
def test_exact_ignores_case_and_outer_spaces(m, a, b):
    assert kind(m, a, b) == EXACT


def test_repeated_inner_spaces_are_normalized_not_exact(m):
    assert kind(m, "Bigg   Boss", "Bigg Boss") == NORMALIZED


# ---- normalized -------------------------------------------------------------
@pytest.mark.parametrize(
    "a,b",
    [
        ("The Kapil Sharma Show", "Kapil Sharma Show"),
        ("Bigg Boss 20", "Bigg Boss Season 20"),
        ("BiggBoss", "Bigg Boss"),
        ("Bigg-Boss!", "Bigg Boss"),
        ("Grey's Anatomy", "Greys Anatomy"),
        ("Tom & Jerry", "Tom and Jerry"),
        ("Pokémon", "Pokemon"),
    ],
)
def test_normalized_matches(m, a, b):
    assert kind(m, a, b) == NORMALIZED


# ---- season variations -------------------------------------------------------
@pytest.mark.parametrize(
    "a,b",
    [
        ("Season 20", "20"),
        ("Bigg Boss S20", "Bigg Boss 20"),
        ("Bigg Boss Season 02", "Bigg Boss 2"),
        ("Bigg Boss 20th Season", "Bigg Boss 20"),
        ("Bigg Boss Seasons 20", "Bigg Boss 20"),
    ],
)
def test_season_label_variants(m, a, b):
    assert kind(m, a, b) == NORMALIZED


def test_different_season_numbers_never_match(m):
    assert kind(m, "Bigg Boss 19", "Bigg Boss 20") == NONE
    assert kind(m, "Bigg Boss Season 19", "Bigg Boss 20") == NONE


def test_season_on_one_side_goes_to_review(m):
    assert kind(m, "Bigg Boss", "Bigg Boss 20") == REVIEW


def test_normalized_form_examples():
    assert normalize_title("  The  Kapil Sharma Show! ") == "kapil sharma show"
    assert normalize_title("Bigg Boss Season 20") == "bigg boss 20"
    assert normalize_title(None) == ""


# ---- aliases ---------------------------------------------------------------------
def test_alias_big_boss(m):
    assert kind(m, "Big Boss", "Bigg Boss") == ALIAS


def test_alias_with_season(m):
    assert kind(m, "BB 20", "Bigg Boss Season 20") == ALIAS


def test_alias_abbreviation(m):
    assert kind(m, "KBC", "Kaun Banega Crorepati") == ALIAS


def test_custom_alias_and_exception(m):
    m.add_alias("Panchayat", ["Phulera Gaon Show"])
    assert kind(m, "Phulera Gaon Show", "Panchayat") == ALIAS
    m.add_exception("Bigg Boss", "Big Boss")
    assert kind(m, "Big Boss", "Bigg Boss") == NONE


# ---- fuzzy ---------------------------------------------------------------------------
def test_minor_spelling_error_is_fuzzy(m):
    r = m.compare("Mirzapurr", "Mirzapur")
    assert r.type == FUZZY and r.score >= 92


def test_fuzzy_threshold_is_adjustable():
    strict = TitleMatcher(threshold=97)
    assert strict.compare("Mirzapurr", "Mirzapur").type == REVIEW  # near match, not applied


def test_similarity_formula():
    assert round(similarity("bigboss", "biggboss"), 2) == 93.33
    assert similarity("abc", "abc") == 100.0
    assert similarity("abc", "xyz") == 0.0


def test_unrelated_titles_do_not_match(m):
    for a, b in [("Panchayat", "Paatal Lok"), ("Mirzapur", "Sacred Games"), ("Squid Game", "Money Heist")]:
        assert kind(m, a, b) == NONE


def test_similar_but_different_titles_not_auto_matched(m):
    # same franchise, different show: must never be applied automatically
    assert kind(m, "Bigg Boss", "Bigg Boss OTT") in (REVIEW, NONE)
    assert kind(m, "Bigg Boss Marathi", "Bigg Boss Kannada") == NONE


def test_short_titles_are_strict(m):
    assert kind(m, "Kohra", "Kohrra") == REVIEW  # 1 typo in a 5/6-letter title: review only
    assert kind(m, "Aha", "Ahaa") == REVIEW  # <5 letters: never auto-matched
    assert kind(m, "Zee", "Zeee") != FUZZY


def test_blank_values_never_match(m):
    assert kind(m, "", "Bigg Boss") == NONE
    assert kind(m, None, "Bigg Boss") == NONE
    assert kind(m, "Bigg Boss", "   ") == NONE


def test_invalid_threshold_rejected():
    with pytest.raises(ValueError):
        TitleMatcher(threshold=10)


# ---- resolve (title vs all headers) -------------------------------------------------
def test_resolve_picks_exact_over_fuzzy(m):
    hs = HeaderSet([(5, "Mirzapur"), (6, "Mirzapurr")])
    res = m.resolve("Mirzapur", hs)
    assert res.status == "matched" and res.best.col == 5 and res.best.result.type == EXACT


def test_resolve_duplicate_headers_is_ambiguous(m):
    hs = HeaderSet([(5, "Bigg Boss"), (9, "bigg boss")])
    res = m.resolve("Bigg Boss", hs)
    assert res.status == "ambiguous" and len(res.candidates) == 2


def test_resolve_two_equally_close_fuzzy_headers_is_ambiguous(m):
    hs = HeaderSet([(5, "Mirzapurr"), (6, "Mirzapuur")])  # both 94.1% from "Mirzapur"
    res = m.resolve("Mirzapur", hs)
    assert res.status == "ambiguous" and {c.col for c in res.candidates} == {5, 6}


def test_resolve_clearly_better_fuzzy_header_wins(m):
    hs = HeaderSet([(5, "Mirzapurr"), (6, "Mirzapurrr")])  # 94.1% vs 88.9% (below threshold)
    res = m.resolve("Mirzapur", hs)
    assert res.status == "matched" and res.best.col == 5 and res.best.result.type == FUZZY


def test_resolve_none_and_review(m):
    hs = HeaderSet([(5, "Panchayat")])
    assert m.resolve("Squid Game", hs).status == "none"
    hs2 = HeaderSet([(5, "Bigg Boss 20")])
    assert m.resolve("Bigg Boss", hs2).status == "review"
