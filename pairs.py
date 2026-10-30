"""Configurable question-pair checks.

Question numbers differ from survey to survey, so the checks are data, not code.
The default list is the original six checks; the web app lets the user type any pairs, e.g.

    Q4 -> Q9
    Q4, Q12 > Q9        (several sources for one destination are fine)
    Q5 → Q12

A question can be a source or a destination, never both (sources are never modified).
"""
from __future__ import annotations

import re
from typing import List, Sequence, Tuple

from .errors import MappingError

Check = Tuple[int, str, str]  # (check number, source question, destination question)

DEFAULT_CHECKS: List[Check] = [
    (1, "Q2", "Q3"),
    (2, "Q2", "Q7"),
    (3, "Q6", "Q7"),
    (4, "Q2", "Q15"),
    (5, "Q6", "Q15"),
    (6, "Q14", "Q15"),
]

_QID = re.compile(r"^\s*(?:q|question)?\s*\.?\s*0*(\d{1,3})\s*$", re.I)
_ARROW = re.compile(r"\s*(?:->|=>|-->|→|➜|>|to)\s*", re.I)


def normalize_qid(text: str) -> str:
    """'q2', 'Question 02', '2' -> 'Q2'.  Raises ValueError for anything else."""
    m = _QID.match(str(text))
    if not m:
        raise ValueError(f"'{text}' is not a question number (use e.g. Q2).")
    return f"Q{int(m.group(1))}"


def pairs_to_text(checks: Sequence[Check]) -> str:
    return "\n".join(f"{s} -> {d}" for _, s, d in checks)


def parse_pairs(text: str) -> List[Check]:
    """Parse one 'source -> destination' pair per line.

    The part before the arrow may list several sources separated by commas; each becomes its own
    check.  Duplicates are dropped, numbering follows the order typed.
    """
    errors: List[str] = []
    out: List[Check] = []
    seen = set()
    for lineno, raw in enumerate((text or "").splitlines(), start=1):
        line = raw.split("#")[0].strip()
        if not line:
            continue
        parts = _ARROW.split(line)
        if len(parts) != 2 or not parts[0].strip() or not parts[1].strip():
            errors.append(f"Line {lineno} ('{raw.strip()}'): write it as  Q2 -> Q3")
            continue
        try:
            dst = normalize_qid(parts[1])
            sources = [normalize_qid(s) for s in re.split(r"[,&+/]|\band\b", parts[0]) if s.strip()]
        except ValueError as exc:
            errors.append(f"Line {lineno}: {exc}")
            continue
        for src in sources:
            if src == dst:
                errors.append(f"Line {lineno}: {src} cannot be compared with itself.")
            elif (src, dst) not in seen:
                seen.add((src, dst))
                out.append((len(out) + 1, src, dst))
    if not out and not errors:
        errors.append("Enter at least one question pair, e.g.  Q2 -> Q3")
    check_pairs(out, errors)
    if errors:
        raise MappingError(errors)
    return out


def check_pairs(checks: Sequence[Check], errors: List[str]) -> None:
    both = sorted({s for _, s, _ in checks} & {d for _, _, d in checks}, key=lambda q: int(q[1:]))
    for q in both:
        errors.append(
            f"{q} is used both as a source and as a destination. Source answers are never modified, "
            "so a question must be only one of the two."
        )


def sources_of(checks: Sequence[Check]) -> List[str]:
    return _ordered(s for _, s, _ in checks)


def destinations_of(checks: Sequence[Check]) -> List[str]:
    return _ordered(d for _, _, d in checks)


def questions_of(checks: Sequence[Check]) -> List[str]:
    """Every question used, sorted by question number."""
    return sorted({q for _, s, d in checks for q in (s, d)}, key=lambda q: int(q[1:]))


def _ordered(items) -> List[str]:
    seen, out = set(), []
    for q in items:
        if q not in seen:
            seen.add(q)
            out.append(q)
    return sorted(out, key=lambda q: int(q[1:]))
