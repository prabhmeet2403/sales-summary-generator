r"""
summary_writer.py
==================
Builds the final Sales & Forecast Summary workbook from the aggregated
GroupSummary objects produced by aggregator.py.

Column layout (fully dynamic based on config.NUM_PRIOR_YEARS_SHOWN /
config.YEARS_WITH_MARGIN_SHOWN, but for the default configuration this
reproduces the sample workbook's layout exactly):

    Name | POC | <year-2> Total | <year-1> Total | <year-1> Margin |
    Q1 | Q2 | Q3 | Q4 | Total | Margin | Comments
      \_______________________/  \____________________________/
            prior years                    current year

Formulas (not hardcoded values) are used wherever the value is derived
from cells written on the SAME sheet -- current-year Total (=SUM of that
row's Q1:Q4) and every subtotal row -- exactly matching the convention
used in the manually built workbook. Figures that come from a *different*
workbook/sheet (quarters, margin, prior-year totals) are necessarily
written as computed values, the same way the original human-built
Summary does it.
"""
from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import config
from aggregator import GroupSummary


class SummaryWriter:
    def __init__(self, target_year: int, prior_years: List[int], years_with_margin: List[int]):
        self.target_year = target_year
        self.prior_years = sorted(prior_years)
        self.years_with_margin = set(years_with_margin)
        self._plan_columns()

    # ------------------------------------------------------------------
    def _plan_columns(self) -> None:
        self.col_name = 1
        self.col_poc = 2
        idx = 3
        self.prior_year_cols: Dict[int, Dict[str, int]] = {}
        for year in self.prior_years:
            entry = {"total": idx}
            idx += 1
            if year in self.years_with_margin:
                entry["margin"] = idx
                idx += 1
            self.prior_year_cols[year] = entry

        self.quarter_cols: Dict[str, int] = {}
        for q in config.QUARTER_ORDER:
            self.quarter_cols[q] = idx
            idx += 1
        self.col_current_total = idx
        idx += 1
        self.col_current_margin = idx
        idx += 1
        self.col_comments = idx
        self.last_col = idx

        # every numeric column, used for subtotal formulas & number format
        self.numeric_cols: List[int] = (
            [c["total"] for c in self.prior_year_cols.values()]
            + [c["margin"] for c in self.prior_year_cols.values() if "margin" in c]
            + list(self.quarter_cols.values())
            + [self.col_current_total, self.col_current_margin]
        )

    # ------------------------------------------------------------------
    def build(self, sections: List[Tuple[config.OutputSection, List[GroupSummary]]]) -> Workbook:
        wb = Workbook()
        ws = wb.active
        ws.title = str(self.target_year)

        self._write_headers(ws)
        current_row = 3

        for section, groups in sections:
            if section.heading:
                self._write_banner_row(ws, current_row, section.heading, fill=config.SECTION_FILL)
                current_row += 1
                current_row += section.blank_rows_after_heading

            self._write_banner_row(ws, current_row, section.title, fill=config.SUBHEADING_FILL)
            current_row += 1
            current_row += section.blank_rows_after_title

            data_start = current_row
            for group in groups:
                self._write_group_row(ws, current_row, group, section.show_poc)
                current_row += 1
            data_end = current_row - 1

            current_row += section.blank_rows_after_data
            if data_end >= data_start:
                self._write_subtotal_row(ws, current_row, section.subtotal_label, data_start, data_end)
            else:
                self._write_banner_row(ws, current_row, section.subtotal_label, fill=config.SUBTOTAL_FILL)
            current_row += 1
            current_row += section.blank_rows_after_subtotal

        self._apply_sheet_formatting(ws, current_row)
        return wb

    # ------------------------------------------------------------------
    def _write_headers(self, ws: Worksheet) -> None:
        bold_center = Font(name=config.FONT_NAME, size=config.FONT_SIZE, bold=True)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Row 1: year banners
        for year, cols in self.prior_year_cols.items():
            start = cols["total"]
            end = cols.get("margin", cols["total"])
            if end > start:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
            cell = ws.cell(row=1, column=start, value=year)
            cell.font = bold_center
            cell.alignment = center

        cur_start = self.quarter_cols["Q1"]
        cur_end = self.col_current_margin
        ws.merge_cells(start_row=1, start_column=cur_start, end_row=1, end_column=cur_end)
        cell = ws.cell(row=1, column=cur_start, value=self.target_year)
        cell.font = bold_center
        cell.alignment = center

        # Row 2: column labels
        labels: Dict[int, str] = {
            self.col_name: "Name",
            self.col_poc: "POC",
            self.col_current_total: "Total",
            self.col_current_margin: "Margin",
            self.col_comments: "Comments",
        }
        for q, col in self.quarter_cols.items():
            labels[col] = q
        for year, cols in self.prior_year_cols.items():
            labels[cols["total"]] = "Total"
            if "margin" in cols:
                labels[cols["margin"]] = "Margin"

        for col in range(1, self.last_col + 1):
            cell = ws.cell(row=2, column=col, value=labels.get(col, ""))
            cell.font = bold_center
            cell.alignment = center
            cell.fill = PatternFill("solid", fgColor=config.HEADER_FILL)

        for col in range(1, self.last_col + 1):
            ws.cell(row=1, column=col).fill = PatternFill("solid", fgColor=config.HEADER_FILL)

    def _write_banner_row(self, ws: Worksheet, row: int, text: str, fill: Optional[str] = None) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=self.last_col)
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = Font(name=config.FONT_NAME, size=config.FONT_SIZE, bold=True)
        if fill:
            for col in range(1, self.last_col + 1):
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=fill)

    def _write_group_row(self, ws: Worksheet, row: int, group: GroupSummary, show_poc: bool) -> None:
        ws.cell(row=row, column=self.col_name, value=group.group_name)
        if show_poc and group.poc:
            ws.cell(row=row, column=self.col_poc, value=group.poc)

        for year, cols in self.prior_year_cols.items():
            total, margin = group.historical.get(year, (None, None))
            if total is not None:
                ws.cell(row=row, column=cols["total"], value=round(total, 2))
            if "margin" in cols and margin is not None:
                ws.cell(row=row, column=cols["margin"], value=round(margin, 2))

        for q, col in self.quarter_cols.items():
            ws.cell(row=row, column=col, value=group.quarters.get(q, 0.0))

        q1_letter = get_column_letter(self.quarter_cols["Q1"])
        q4_letter = get_column_letter(self.quarter_cols["Q4"])
        ws.cell(
            row=row,
            column=self.col_current_total,
            value=f"=SUM({q1_letter}{row}:{q4_letter}{row})",
        )
        ws.cell(row=row, column=self.col_current_margin, value=group.total_margin)

        if group.comment:  # Rule 6: comment blanks stay blank
            c = ws.cell(row=row, column=self.col_comments, value=group.comment)
            c.alignment = Alignment(wrap_text=True, vertical="top")

    def _write_subtotal_row(self, ws: Worksheet, row: int, label: str, data_start: int, data_end: int) -> None:
        cell = ws.cell(row=row, column=self.col_name, value=label)
        cell.font = Font(name=config.FONT_NAME, size=config.FONT_SIZE, bold=True)
        for col in self.numeric_cols:
            letter = get_column_letter(col)
            formula_cell = ws.cell(row=row, column=col, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
            formula_cell.font = Font(name=config.FONT_NAME, size=config.FONT_SIZE, bold=True)
        for col in range(1, self.last_col + 1):
            ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=config.SUBTOTAL_FILL)

    # ------------------------------------------------------------------
    def _apply_sheet_formatting(self, ws: Worksheet, last_row: int) -> None:
        base_font = Font(name=config.FONT_NAME, size=config.FONT_SIZE)
        for row in ws.iter_rows(min_row=1, max_row=max(last_row, 2), min_col=1, max_col=self.last_col):
            for cell in row:
                if cell.font is None or cell.font.name is None:
                    cell.font = base_font

        for col in self.numeric_cols:
            for row in range(3, last_row + 1):
                ws.cell(row=row, column=col).number_format = config.CURRENCY_FORMAT

        ws.column_dimensions[get_column_letter(self.col_name)].width = config.NAME_COLUMN_WIDTH
        ws.column_dimensions[get_column_letter(self.col_poc)].width = config.POC_COLUMN_WIDTH
        for col in self.numeric_cols:
            ws.column_dimensions[get_column_letter(col)].width = config.NUMBER_COLUMN_WIDTH
        ws.column_dimensions[get_column_letter(self.col_comments)].width = config.COMMENTS_COLUMN_WIDTH

        ws.freeze_panes = "A3"
        ws.sheet_view.showGridLines = False
