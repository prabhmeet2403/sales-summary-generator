"""
monthly_view.py
================
Prepares the per-month Actual/Forecast revenue and margin breakdown that
powers the "<year> Monthly Performance" worksheet (Worksheet 2).

This module adds no new business rule and duplicates no existing
calculation. It:

  - reuses the exact same `ProjectRow` objects `excel_reader.py` already
    read for Worksheet 1 (`monthly_revenue` / `monthly_margin`, already
    parsed per project row, per calendar month -- see
    `excel_reader.read_project_rows`),
  - reuses the exact same grouping key `aggregator.py` already uses
    (`normalize_name(row.group)`) and the exact same per-section
    Sub-Group/DS-code filter (`section.ds_codes`),
  - reuses the exact same GROUP LIST `aggregator.py` already computed
    and validated (the `GroupSummary` objects passed in) -- Worksheet 2
    shows precisely the same groups, in the same order, as Worksheet 1.
    This module never independently decides whether a group belongs in
    the Summary (Rule 6's blank-group drop, sorting, etc. have already
    happened by the time this module runs),
  - reuses each group's already-matched Comment (`GroupSummary.comment`)
    instead of re-querying the ClientComments sheet,
  - reuses each group's already-computed `total_revenue`/`total_margin`
    instead of re-deriving them from the monthly figures a second time.

The only genuinely new computation here is re-bucketing the SAME monthly
figures by calendar month instead of by quarter (`aggregator.py` already
throws the month-level detail away once it has summed into Q1-Q4), and
reading each month's own Actual/Forecast role label directly from the
sheet's own type-header row -- using the exact column position
`build_column_map` already resolved -- instead of hardcoding which
calendar months are "Actual" and which are "Forecast". If the source
sheet's own labels ever change (e.g. "Actual" becomes "Closed", or the
Actual/Forecast split moves from May/June to a different month), this
keeps working with no code change, because it reads the sheet's own
words rather than assuming a fixed month range.
"""
from __future__ import annotations

from collections import Counter, OrderedDict
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import config
from aggregator import GroupSummary
from excel_reader import ColumnMap, ProjectRow, normalize_name
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class MonthlyGroupSummary:
    """One group's month-by-month view, mirroring `GroupSummary` but
    bucketed by calendar month (1-12) instead of by quarter. Every field
    other than `monthly_revenue`/`monthly_margin` is copied verbatim
    from the already-validated `GroupSummary` for the same group."""

    group_name: str
    poc: Optional[str]
    monthly_revenue: Dict[int, float]
    monthly_margin: Dict[int, float]
    total_revenue: float
    total_margin: float
    confidence: Optional[str]
    comment: Optional[str]


def resolve_month_roles(ws: Worksheet, cmap: ColumnMap) -> Dict[int, str]:
    """Return {month: role_text} by reading each month's own type-header
    cell directly -- the exact same cell `build_column_map` already
    inspected to decide "this is a revenue column" (see
    `excel_reader.build_column_map`), just read again here for its
    literal text instead of being collapsed into the generic "revenue"
    role bucket. Whatever the sheet's own header actually says (e.g.
    "Actual", "Forecast") is used verbatim (title-cased for a clean
    display) -- nothing about which calendar months are which is
    hardcoded anywhere in this function.
    """
    roles: Dict[int, str] = {}
    if not cmap.type_header_row:
        return roles
    for month, cols in cmap.months.items():
        revenue_col = cols.get("revenue")
        if revenue_col is None:
            continue
        raw = ws.cell(cmap.type_header_row, revenue_col).value
        text = str(raw).strip() if raw is not None else ""
        roles[month] = text.title() if text else "Actual"
    return roles


def build_monthly_sections(
    rows: List[ProjectRow],
    cmap: ColumnMap,
    ws_main: Worksheet,
    section_results: List[Tuple["config.OutputSection", List[GroupSummary]]],
) -> List[Tuple["config.OutputSection", List[MonthlyGroupSummary]]]:
    """For every section/group in `section_results`, compute that same
    group's month-by-month revenue and margin.

    The set of groups, their order, their POC, and their comment are all
    taken as-is from `section_results` -- this function only adds the
    monthly breakdown on top of groups that already exist; it never
    adds, removes, or reorders a group relative to whatever it's given.

    Callers pass this Worksheet 1's own `section_results` plus (per
    `config.WORKSHEET2_ADDITIONAL_SECTIONS`) any sections that belong
    only on Worksheet 2 -- see `main.py`/`gui/runner.py`'s
    `worksheet2_extra_section_results`. This function itself has no
    opinion on which sections came from where; it just computes a
    monthly breakdown for whatever section/group pairs it's handed.
    """
    # Group every row by normalised name ONCE, so every group below is a
    # simple dict lookup instead of a fresh linear scan each time.
    rows_by_name: "OrderedDict[str, List[ProjectRow]]" = OrderedDict()
    for r in rows:
        rows_by_name.setdefault(normalize_name(r.group), []).append(r)

    monthly_results: List[Tuple[config.OutputSection, List[MonthlyGroupSummary]]] = []

    for section, groups in section_results:
        monthly_groups: List[MonthlyGroupSummary] = []
        for g in groups:
            group_rows = [
                r for r in rows_by_name.get(normalize_name(g.group_name), [])
                if r.ds_code in section.ds_codes
            ]

            monthly_revenue: Dict[int, float] = {m: 0.0 for m in cmap.months}
            monthly_margin: Dict[int, float] = {m: 0.0 for m in cmap.months}
            confidence_counter: Counter = Counter()

            for r in group_rows:
                for month, revenue in r.monthly_revenue.items():
                    monthly_revenue[month] = monthly_revenue.get(month, 0.0) + revenue
                for month, margin in r.monthly_margin.items():
                    monthly_margin[month] = monthly_margin.get(month, 0.0) + margin
                if cmap.renewal_confidence:
                    raw_conf = ws_main.cell(r.row_index, cmap.renewal_confidence).value
                    if raw_conf is not None and str(raw_conf).strip() != "":
                        confidence_counter[str(raw_conf).strip()] += 1

            monthly_revenue = {m: round(v, 2) for m, v in monthly_revenue.items()}
            monthly_margin = {m: round(v, 2) for m, v in monthly_margin.items()}
            confidence = confidence_counter.most_common(1)[0][0] if confidence_counter else None

            monthly_groups.append(
                MonthlyGroupSummary(
                    group_name=g.group_name,
                    poc=g.poc,
                    monthly_revenue=monthly_revenue,
                    monthly_margin=monthly_margin,
                    total_revenue=g.total_revenue,
                    total_margin=g.total_margin,
                    confidence=confidence,
                    comment=g.comment,
                )
            )
        monthly_results.append((section, monthly_groups))

    return monthly_results
