"""Builds small synthetic survey workbooks with configurable layouts for the tests."""
from __future__ import annotations

import io
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from core import (
    TitleMatcher,
    build_audit_report,
    inspect_snapshot,
    load_snapshot,
    process_workbook,
)

HEADERS = ["Bigg Boss", "Kapil Sharma Show", "Mirzapur", "Panchayat", "Squid Game"]
ALL_Q = ["Q2", "Q3", "Q6", "Q7", "Q14", "Q15"]
SRC = {"Q2", "Q6", "Q14"}


class SurveyBuilder:
    """Lay out a survey sheet.

    Row ``top_pad+1``: question labels, row ``top_pad+2``: headers, data from ``top_pad+3``.
    Column ``left_pad`` (if >0) holds respondent ids; questions follow, one after the other,
    separated by ``gap`` empty columns.
    """

    def __init__(
        self,
        sheet: str = "Survey",
        top_pad: int = 0,
        left_pad: int = 1,
        merged_labels: bool = False,
        src_cols: int = 3,
        headers: List[str] | None = None,
        gap: int = 0,
        extra_sheet: bool = True,
        rename: Dict[str, str] | None = None,
    ):
        self.sheet = sheet
        self.top_pad = top_pad
        self.left_pad = left_pad
        self.merged_labels = merged_labels
        self.src_cols = src_cols
        self.headers = list(headers or HEADERS)
        self.gap = gap
        self.extra_sheet = extra_sheet
        self.rename = dict(rename or {})  # internal role name (Q2..Q15) -> label printed in the sheet
        self.cols: Dict[str, List[int]] = {}
        col = left_pad + 1
        for q in ALL_Q:
            n = src_cols if q in SRC else len(self.headers)
            self.cols[q] = list(range(col, col + n))
            col += n + gap
        self.respondents: List[Dict[str, list]] = []
        self.dest_headers: Dict[str, List[str]] = {
            q: list(self.headers) for q in ALL_Q if q not in SRC
        }
        self.formulas: Dict[tuple, str] = {}

    # -- geometry ----------------------------------------------------------
    @property
    def label_row(self):
        return self.top_pad + 1

    @property
    def header_row(self):
        return self.top_pad + 2

    @property
    def first_row(self):
        return self.top_pad + 3

    def row_of(self, i: int) -> int:
        return self.first_row + i

    def col_of(self, q: str, header: str | int) -> int:
        if isinstance(header, int):
            return self.cols[q][header]
        return self.cols[q][self.dest_headers[q].index(header)]

    def addr_of(self, i: int, q: str, header) -> str:
        return f"{get_column_letter(self.col_of(q, header))}{self.row_of(i)}"

    # -- data -----------------------------------------------------------------
    def add(self, **answers) -> int:
        self.respondents.append(answers)
        return len(self.respondents) - 1

    def dest_values(self, **by_header) -> dict:
        return by_header

    def build(self) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = self.sheet
        if self.top_pad:
            ws.cell(row=1, column=1, value="Survey export - generated for testing")
        lr, hr = self.label_row, self.header_row
        if self.left_pad:
            ws.cell(row=hr, column=1, value="Resp ID")
        for q in ALL_Q:
            cols = self.cols[q]
            label = self.rename.get(q, q)
            ws.cell(row=lr, column=cols[0], value=f"{label}. Which shows have you watched?")
            if self.merged_labels and len(cols) > 1:
                ws.merge_cells(start_row=lr, start_column=cols[0], end_row=lr, end_column=cols[-1])
            for j, c in enumerate(cols):
                if q in SRC:
                    ws.cell(row=hr, column=c, value=f"{label}_{j + 1}")
                else:
                    ws.cell(row=hr, column=c, value=self.dest_headers[q][j])
        for i, resp in enumerate(self.respondents):
            r = self.row_of(i)
            if self.left_pad:
                ws.cell(row=r, column=1, value=1000 + i)
            for q, values in resp.items():
                for j, v in enumerate(values):
                    if v is not None:
                        ws.cell(row=r, column=self.cols[q][j], value=v)
        for (r, c), f in self.formulas.items():
            ws.cell(row=r, column=c, value=f)
        # some formatting that must survive
        ws.cell(row=hr, column=max(1, self.left_pad)).font = Font(bold=True)
        ws.column_dimensions["A"].width = 17
        if self.extra_sheet:
            ws2 = wb.create_sheet("Notes")
            ws2["A1"] = "Q3 should stay untouched"
            ws2["B2"] = 1
            ws2["C3"] = 1
            ws2["C3"].fill = PatternFill("solid", start_color="FFC000", end_color="FFC000")
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()


def layout_for(data: bytes, sheet: str = "Survey", checks=None):
    return inspect_snapshot(load_snapshot(data, sheet), checks)


def run(builder: SurveyBuilder, threshold: float = 92.0, matcher: TitleMatcher | None = None, checks=None):
    data = builder.build()
    layout = layout_for(data, builder.sheet, checks)
    m = matcher or TitleMatcher(threshold=threshold)
    res = process_workbook(data, "survey.xlsx", layout, m, checks=checks)
    return data, layout, res


def fill_hex(cell) -> str:
    rgb = cell.fill.fgColor.rgb
    return rgb[-6:].upper() if isinstance(rgb, str) and cell.fill.fill_type == "solid" else ""
