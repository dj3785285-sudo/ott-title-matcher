"""Automatic detection of question layout + validation of (possibly edited) mappings.

Nothing here hard-codes column letters, header rows or sheet names.  Detection is
heuristic, so every result carries ``confident`` and ``issues``; the web app
always requires the user to confirm the mapping before anything is changed.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .errors import MappingError
from .excel_utils import addr, columns_to_spec, col_letter, parse_columns
from .pairs import (
    DEFAULT_CHECKS,
    check_pairs,
    destinations_of,
    normalize_qid,
    questions_of,
    sources_of,
)
from .workbook_io import Snapshot, is_blank

# Defaults (the original six checks).  Real runs pass their own ``checks`` list.
SOURCE_QUESTIONS = tuple(sources_of(DEFAULT_CHECKS))
DEST_QUESTIONS = tuple(destinations_of(DEFAULT_CHECKS))
ALL_QUESTIONS = tuple(questions_of(DEFAULT_CHECKS))

SCAN_ROWS = 15  # how many top rows are searched for question labels
HEADER_SEARCH_ROWS = 5
DATA_SAMPLE_ROWS = 200

QUESTION_RE = re.compile(r"^\s*(?:q|question)\s*\.?\s*0*(\d{1,3})(?!\d)", re.I)


def question_number(value) -> Optional[int]:
    """'Q2', 'Q2_1', 'Q 2.', 'Question 2 - Which ...' -> 2."""
    if not isinstance(value, str):
        return None
    m = QUESTION_RE.match(value)
    return int(m.group(1)) if m else None


@dataclass
class QuestionMapping:
    qid: str
    columns: List[int] = field(default_factory=list)
    label_cell: str = ""
    merged: bool = False
    note: str = ""

    @property
    def spec(self) -> str:
        return columns_to_spec(self.columns)


@dataclass
class SheetLayout:
    sheet: str
    label_row: int = 1
    header_row: int = 2
    first_row: int = 3
    last_row: int = 3
    questions: Dict[str, QuestionMapping] = field(default_factory=dict)
    confident: bool = False
    issues: List[str] = field(default_factory=list)

    def cols(self, qid: str) -> List[int]:
        q = self.questions.get(qid)
        return list(q.columns) if q else []

    @property
    def found(self) -> int:
        return sum(1 for q in self.questions if self.cols(q))


def build_layout(
    sheet: str,
    label_row: int,
    header_row: int,
    first_row: int,
    last_row: int,
    specs: Dict[str, str],
) -> SheetLayout:
    """Create a layout from user input; column specs use Excel notation ('BK:BQ')."""
    layout = SheetLayout(sheet, int(label_row), int(header_row), int(first_row), int(last_row))
    errors = []
    for raw_qid, spec in specs.items():
        qid = normalize_qid(raw_qid)
        try:
            cols = parse_columns(spec)
        except ValueError as exc:
            errors.append(f"{qid}: {exc}")
            cols = []
        layout.questions[qid] = QuestionMapping(qid, cols, note="user-confirmed")
    if errors:
        raise MappingError(errors)
    layout.confident = True
    return layout


# --------------------------------------------------------------------------
# Detection
# --------------------------------------------------------------------------
def _nonempty_in_column(snap: Snapshot, col: int, r0: int, r1: int) -> bool:
    return any(not is_blank(snap.value(r, col)) for r in range(r0, r1 + 1))


def _looks_like_header(snap: Snapshot, row: int, cols: Sequence[int]) -> bool:
    """True when the row holds text titles instead of 0/1/2 codes."""
    vals = [snap.value(row, c) for c in cols]
    vals = [v for v in vals if not is_blank(v)]
    if len(vals) < 2:  # a single odd value is data, not a sub-header row
        return False
    texty = 0
    for v in vals:
        if isinstance(v, str) and v.strip() not in {"0", "1", "2"}:
            try:
                float(v)
            except ValueError:
                texty += 1
    return texty / len(vals) >= 0.5


def inspect_snapshot(snap: Snapshot, checks=None) -> SheetLayout:
    """Detect the layout.  Every question label found (Q1, Q2, ... any number) is mapped;
    the questions named in ``checks`` (default: the six standard ones) are the *required* ones."""
    checks = list(checks if checks is not None else DEFAULT_CHECKS)
    required = questions_of(checks)
    dest_qids = destinations_of(checks)
    layout = SheetLayout(snap.sheet)
    for q in required:
        layout.questions[q] = QuestionMapping(q)

    # 1. find the label row ---------------------------------------------------
    best_row, best_count, row_labels = None, 0, {}
    for r in range(1, min(snap.max_row, SCAN_ROWS) + 1):
        labels = []
        for c in range(1, snap.max_col + 1):
            qn = question_number(snap.value(r, c))
            if qn is not None:
                labels.append((c, qn))
        row_labels[r] = labels
        count = len({qn for _, qn in labels})
        if count > best_count:
            best_row, best_count = r, count

    if best_row is None:
        layout.issues.append(
            "No question labels (Q2, Q3, Q6, Q7, Q14, Q15) were found in the top "
            f"{SCAN_ROWS} rows. Please enter the column ranges manually."
        )
        layout.last_row = snap.max_row
        return layout

    layout.label_row = best_row
    labels = sorted(row_labels[best_row])

    # 2. group consecutive labels of the same question into column segments ----
    groups: List[Tuple[int, int, int]] = []  # (qnum, start_col, label_count)
    for col, qn in labels:
        if groups and groups[-1][0] == qn:
            continue
        groups.append((qn, col, 0))

    found_before: Dict[str, int] = {}
    for i, (qn, start, _) in enumerate(groups):
        end = groups[i + 1][1] - 1 if i + 1 < len(groups) else snap.max_col
        merged = snap.merged_at(best_row, start)
        is_merged = bool(merged and merged[3] > merged[1])
        if is_merged:
            end = max(min(end, snap.max_col), merged[3])
        # trim trailing columns with no content at all
        r_hi = min(snap.max_row, best_row + DATA_SAMPLE_ROWS)
        while end > start and not _nonempty_in_column(snap, end, best_row, r_hi):
            end -= 1
        qid = f"Q{qn}"
        if qid in found_before:
            layout.issues.append(
                f"{qid} appears more than once in row {best_row}; the first occurrence "
                f"({col_letter(found_before[qid])}) was used - please confirm."
            )
            continue
        found_before[qid] = start
        layout.questions[qid] = QuestionMapping(
            qid,
            list(range(start, end + 1)),
            label_cell=addr(best_row, start),
            merged=is_merged,
            note="merged label cell" if is_merged else "",
        )

    missing = [q for q in required if not layout.cols(q)]
    if missing:
        layout.issues.append(
            f"Could not find label(s) for: {', '.join(missing)}. Enter those column ranges manually "
            "(or correct the question numbers in the question-pair list)."
        )

    # 3. header row (titles of the destination questions) ----------------------
    dest_cols = [c for q in dest_qids for c in layout.cols(q)]
    header_row = best_row + 1
    if dest_cols:
        chosen, best_guess, best_share = None, None, 0.0
        for r in range(best_row + 1, best_row + 1 + HEADER_SEARCH_ROWS):
            if r > snap.max_row:
                break
            filled = sum(
                1 for c in dest_cols if isinstance(snap.value(r, c), str) and snap.value(r, c).strip()
            )
            share = filled / len(dest_cols)
            if share >= 0.6:
                chosen = r
                break
            if share > best_share:
                best_share, best_guess = share, r
        if chosen is not None:
            header_row = chosen
        elif best_guess is not None:
            header_row = best_guess
            layout.issues.append(
                f"Title headers in row {best_guess} are only partly filled; please check the header row."
            )
        else:
            layout.issues.append("No row below the question labels contains title headers.")
    else:
        layout.issues.append("Header row could not be determined because no destination question was found.")
    layout.header_row = header_row

    # 4. first respondent row ----------------------------------------------------
    first = header_row + 1
    skipped = 0
    while skipped < 3 and first <= snap.max_row and dest_cols and _looks_like_header(snap, first, dest_cols):
        first += 1
        skipped += 1
    if skipped:
        layout.issues.append(
            f"Skipped {skipped} text-only row(s) below the header row; respondent data is assumed to "
            f"start at row {first} - please confirm."
        )
    layout.first_row = first

    # 5. last respondent row -------------------------------------------------------
    mapped = [c for q in layout.questions for c in layout.cols(q)]
    last = snap.max_row
    while last > first and not any(not is_blank(snap.value(last, c)) for c in mapped):
        last -= 1
    layout.last_row = max(last, first)

    layout.confident = not layout.issues and all(layout.cols(q) for q in required)
    return layout


def inspect_snapshots(snaps: Dict[str, Snapshot], checks=None) -> Dict[str, SheetLayout]:
    return {name: inspect_snapshot(s, checks) for name, s in snaps.items()}


def best_sheet(layouts: Dict[str, SheetLayout]) -> str:
    """Sheet with the most detected questions (confident sheets win ties)."""
    return max(layouts, key=lambda n: (layouts[n].found, layouts[n].confident, -list(layouts).index(n)))


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------
def dest_headers(snap: Snapshot, layout: SheetLayout, qid: str) -> List[Tuple[int, str]]:
    """(column, title) for every non-empty header cell of a destination question."""
    out = []
    for c in layout.cols(qid):
        v = snap.value(layout.header_row, c)
        if v is None or isinstance(v, bool):
            continue
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        text = str(v).strip()
        if text:
            out.append((c, text))
    return out


def validate_layout(snap: Snapshot, layout: SheetLayout, checks=None) -> Tuple[List[str], List[str]]:
    """Return (errors, warnings).  Errors block processing.  ``checks`` = the question pairs in use."""
    checks = list(checks if checks is not None else DEFAULT_CHECKS)
    used = questions_of(checks)
    dest_qids = destinations_of(checks)
    errors: List[str] = []
    warnings: List[str] = []
    check_pairs(checks, errors)

    for qid in used:
        if not layout.cols(qid):
            errors.append(f"{qid} has no columns mapped. Enter its column range (e.g. BK:BQ).")
    if layout.header_row < 1:
        errors.append("Header row must be 1 or greater.")
    if layout.first_row <= layout.header_row:
        errors.append("The first respondent row must be below the title-header row.")
    if layout.last_row < layout.first_row:
        errors.append("The last respondent row must not be above the first respondent row.")
    if layout.first_row > snap.max_row:
        errors.append(
            f"The first respondent row ({layout.first_row}) is beyond the end of the sheet "
            f"(last used row is {snap.max_row})."
        )

    # a column may belong to one question only -> source answers can never be edited
    owner: Dict[int, str] = {}
    for qid in used:
        for c in layout.cols(qid):
            if c in owner and owner[c] != qid:
                errors.append(
                    f"Column {col_letter(c)} is mapped to both {owner[c]} and {qid}. "
                    "Question ranges must not overlap (source answers must never be modified)."
                )
            owner.setdefault(c, qid)

    for qid in dest_qids:
        if layout.cols(qid) and not dest_headers(snap, layout, qid):
            errors.append(
                f"{qid}: no title headers were found in row {layout.header_row}. "
                "Check the header row number or the column range."
            )
        else:
            titles = [t.casefold() for _, t in dest_headers(snap, layout, qid)]
            dupes = sorted({t for t in titles if titles.count(t) > 1})
            if dupes:
                warnings.append(
                    f"{qid} has duplicate title headers ({', '.join(dupes[:5])}); "
                    "matches to them will be treated as ambiguous and sent to review."
                )
            empties = len(layout.cols(qid)) - len(titles)
            if empties:
                warnings.append(f"{qid}: {empties} column(s) have an empty header and will be ignored.")

    for qid in used:
        beyond = [c for c in layout.cols(qid) if c > snap.max_col]
        if beyond:
            warnings.append(f"{qid}: column(s) beyond the last used column are mapped (they are empty).")
    return errors, warnings
