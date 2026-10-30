"""Title-matching engine.

Matching order (first rule that fires wins):

1. ``exact``       - equal after trimming and case-folding.
2. ``normalized``  - equal after punctuation/space/season-label/"The" normalisation,
                     including joined vs. separated words.
3. ``alias``       - both titles belong to the same recognised variant group.
4. ``fuzzy``       - conservative similarity >= threshold (default 92).

Guards against false positives
------------------------------
* Different numbers never fuzzy-match ("Bigg Boss 19" vs "Bigg Boss 20").
* A number on one side only ("Bigg Boss" vs "Bigg Boss 20") is sent to review.
* Short titles (< ``min_fuzzy_len`` letters) need a much higher score, and very
  short titles (< ``short_fuzzy_len``) are never fuzzy-matched at all.
* If several destination headers qualify equally well the title is *ambiguous*
  and goes to the manual-review list instead of being applied.
"""
from __future__ import annotations

import itertools
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from .aliases import AliasTable, default_alias_table
from .normalize import (
    compact,
    digit_tokens,
    normalize_title,
    strip_trailing_number,
)

try:  # RapidFuzz is the preferred scorer (used on Streamlit Cloud).
    from rapidfuzz import fuzz as _rf_fuzz
except Exception:  # pragma: no cover - fallback keeps the engine usable anywhere
    _rf_fuzz = None

# Match types ---------------------------------------------------------------
EXACT = "exact"
NORMALIZED = "normalized"
ALIAS = "alias"
FUZZY = "fuzzy"
REVIEW = "review"
NONE = "none"

AUTO_TYPES = (EXACT, NORMALIZED, ALIAS, FUZZY)
# Higher number = stronger evidence.  Used to resolve several titles that
# point at the same destination cell (exact > normalized > alias > fuzzy).
PRIORITY: Dict[str, int] = {EXACT: 4, NORMALIZED: 3, ALIAS: 2, FUZZY: 1}
# Yellow for exact; green for every other automatic match type.
COLOR_FOR: Dict[str, str] = {EXACT: "Yellow", NORMALIZED: "Green", ALIAS: "Green", FUZZY: "Green"}

DEFAULT_THRESHOLD = 92.0


# Similarity ----------------------------------------------------------------
def _lcs_len(a: str, b: str) -> int:
    if len(a) < len(b):
        a, b = b, a
    prev = [0] * (len(b) + 1)
    for ca in a:
        cur = [0]
        for j, cb in enumerate(b, 1):
            cur.append(prev[j - 1] + 1 if ca == cb else max(prev[j], cur[j - 1]))
        prev = cur
    return prev[-1]


def similarity(a: str, b: str) -> float:
    """Normalised Indel similarity 0-100 (identical to ``rapidfuzz.fuzz.ratio``)."""
    if not a and not b:
        return 100.0
    if not a or not b:
        return 0.0
    if _rf_fuzz is not None:
        return float(_rf_fuzz.ratio(a, b))
    return 200.0 * _lcs_len(a, b) / (len(a) + len(b))


def scorer_name() -> str:
    return "rapidfuzz" if _rf_fuzz is not None else "built-in fallback (same formula)"


# Result types ----------------------------------------------------------------
@dataclass(frozen=True)
class MatchResult:
    type: str
    score: float = 0.0
    reason: str = ""


NO_MATCH = MatchResult(NONE, 0.0, "")


@dataclass(frozen=True)
class Candidate:
    col: int
    header: str
    result: MatchResult


@dataclass
class Resolution:
    """Outcome of comparing one source title with every header of a question."""

    status: str  # 'matched' | 'ambiguous' | 'review' | 'none'
    candidates: List[Candidate] = field(default_factory=list)
    reason: str = ""

    @property
    def best(self) -> Optional[Candidate]:
        return self.candidates[0] if self.status == "matched" and self.candidates else None


_header_ids = itertools.count(1)


class HeaderSet:
    """Destination-question headers: a list of (column index, header text)."""

    def __init__(self, items: Iterable[Tuple[int, str]]):
        self.items: List[Tuple[int, str]] = list(items)
        self.uid = next(_header_ids)

    def __iter__(self):
        return iter(self.items)

    def __len__(self):
        return len(self.items)


class TitleMatcher:
    def __init__(
        self,
        threshold: float = DEFAULT_THRESHOLD,
        aliases: AliasTable | None = None,
        exceptions: Iterable[Tuple[str, str]] = (),
        min_fuzzy_len: int = 7,
        short_fuzzy_len: int = 5,
        short_threshold: float = 96.0,
        tie_margin: float = 2.0,
        review_margin: float = 12.0,
    ):
        if not 50.0 <= float(threshold) <= 100.0:
            raise ValueError("Fuzzy threshold must be between 50 and 100.")
        self.threshold = float(threshold)
        self.aliases = aliases if aliases is not None else default_alias_table()
        self.min_fuzzy_len = min_fuzzy_len
        self.short_fuzzy_len = short_fuzzy_len
        self.short_threshold = max(short_threshold, self.threshold)
        self.tie_margin = tie_margin
        self.review_floor = max(60.0, self.threshold - review_margin)
        self._exceptions: set = set()
        for a, b in exceptions:
            self.add_exception(a, b)
        self._cache: Dict[Tuple[str, int], Resolution] = {}
        self._norm_cache: Dict[str, Tuple[str, str, Tuple[str, ...], str]] = {}

    # -- extensibility ------------------------------------------------------
    def add_alias(self, canonical: str, variants: Iterable[str] = ()) -> None:
        self.aliases.add(canonical, variants)
        self._cache.clear()

    def add_exception(self, a: str, b: str) -> None:
        ka, kb = compact(normalize_title(a)), compact(normalize_title(b))
        if ka and kb:
            self._exceptions.add(frozenset((ka, kb)))
        self._cache.clear()

    # -- helpers --------------------------------------------------------------
    def _forms(self, text: str) -> Tuple[str, str, Tuple[str, ...], str]:
        """(normalised, compact, digits, alias-group-of-base)."""
        hit = self._norm_cache.get(text)
        if hit is None:
            norm = normalize_title(text)
            base = compact(strip_trailing_number(norm))
            hit = (norm, compact(norm), digit_tokens(norm), self.aliases.group_of(base) if base else "")
            self._norm_cache[text] = hit
        return hit

    # -- pairwise comparison --------------------------------------------------
    def compare(self, source: str, header: str) -> MatchResult:
        if source is None or header is None:
            return NO_MATCH
        s_raw, h_raw = str(source).strip(), str(header).strip()
        if not s_raw or not h_raw:
            return NO_MATCH

        s_norm, s_key, s_dig, s_grp = self._forms(s_raw)
        h_norm, h_key, h_dig, h_grp = self._forms(h_raw)
        if not s_key or not h_key:
            return NO_MATCH

        if frozenset((s_key, h_key)) in self._exceptions:
            return NO_MATCH

        if s_raw.casefold() == h_raw.casefold():
            return MatchResult(EXACT, 100.0, "identical ignoring case and outer spaces")

        if s_key == h_key:
            return MatchResult(NORMALIZED, 100.0, "identical after normalisation")

        if s_grp and s_grp == h_grp and s_dig == h_dig:
            return MatchResult(ALIAS, 100.0, "recognised title variant")

        # ---- numbers: never fuzzy-match different seasons ------------------
        if s_dig != h_dig:
            if s_dig and h_dig:
                return NO_MATCH  # e.g. Bigg Boss 19 vs Bigg Boss 20: different titles
            s_base = compact(strip_trailing_number(s_norm))
            h_base = compact(strip_trailing_number(h_norm))
            same_base = bool(s_base) and (
                s_base == h_base
                or (s_grp and s_grp == h_grp)
                or similarity(s_base, h_base) >= self.threshold
            )
            if same_base:
                return MatchResult(
                    REVIEW,
                    100.0 if s_base == h_base else round(similarity(s_base, h_base), 1),
                    "season/number appears on one side only - confirm it is the same title",
                )
            return NO_MATCH

        # ---- fuzzy -----------------------------------------------------------
        score = round(similarity(s_key, h_key), 2)
        shortest = min(len(s_key), len(h_key))
        if shortest < self.short_fuzzy_len:
            if score >= self.review_floor:
                return MatchResult(
                    REVIEW, score, f"very short title (<{self.short_fuzzy_len} letters): never auto-matched"
                )
            return NO_MATCH
        required = self.threshold if shortest >= self.min_fuzzy_len else self.short_threshold
        if score >= required:
            return MatchResult(FUZZY, score, f"similarity {score:.1f}% >= {required:.0f}%")
        if score >= self.review_floor:
            return MatchResult(
                REVIEW, score, f"similar ({score:.1f}%) but below the required {required:.0f}%"
            )
        return NO_MATCH

    # -- one source title vs. all destination headers -------------------------
    def resolve(self, title: str, headers: HeaderSet) -> Resolution:
        key = (str(title).strip(), headers.uid)
        hit = self._cache.get(key)
        if hit is not None:
            return hit
        out = self._resolve(key[0], headers)
        self._cache[key] = out
        return out

    def _resolve(self, title: str, headers: HeaderSet) -> Resolution:
        cands: List[Candidate] = []
        for col, text in headers:
            res = self.compare(title, text)
            if res.type != NONE:
                cands.append(Candidate(col, text, res))
        if not cands:
            return Resolution("none")

        qualified = [c for c in cands if c.result.type in AUTO_TYPES]
        if qualified:
            top_priority = max(PRIORITY[c.result.type] for c in qualified)
            top = [c for c in qualified if PRIORITY[c.result.type] == top_priority]
            if top_priority == PRIORITY[FUZZY]:
                top.sort(key=lambda c: (-c.result.score, c.col))
                top = [c for c in top if top[0].result.score - c.result.score <= self.tie_margin]
            else:
                top.sort(key=lambda c: c.col)
            if len(top) == 1:
                return Resolution("matched", top)
            return Resolution(
                "ambiguous",
                top,
                f"{len(top)} destination headers match equally well - choose manually",
            )

        near = sorted(
            (c for c in cands if c.result.type == REVIEW),
            key=lambda c: (-c.result.score, c.col),
        )[:5]
        if near:
            return Resolution("review", near, near[0].result.reason)
        return Resolution("none")
