"""
historical_lookup.py
=====================
Provides an independent, from-first-principles RECOMPUTATION of a
group's prior-year Total Revenue / Total Margin, used by aggregator.py:

- as the value for Margin in the most recent prior year (there is no
  same-sheet embedded reference column for margin anywhere in the
  workbook, so this is the only source for it);
- as a fallback for Total when that year's embedded same-sheet reference
  column (see aggregator.attach_historical) is entirely blank for a
  group;
- as a cross-check against the embedded reference column, to flag
  drift in the validation report.

Two strategies are used, chosen automatically per sheet:

1. "group_match" -- the prior-year sheet has its own Group / Sub-Group
   columns (true for every year that follows the modern schema, e.g.
   2025). We aggregate that sheet exactly the same way the current year
   is aggregated: same Group name, same Sub-Group/DS-code family, so a
   Track-1 group is only ever compared to that group's Track-1 rows in
   the prior year (not, say, a same-named Staffing row). The monthly
   Actual/Margin columns are always summed directly rather than trusting
   that sheet's own cached "Total Revenue" cell, because that cached
   cell is occasionally left blank even when monthly data exists (and,
   more rarely, is simply stale) -- summing the twelve monthly cells is
   the one thing that is always self-consistent.

2. "fuzzy_name_match" -- older sheets (e.g. 2024) have no Group column at
   all, just a free-text project Name. We fall back to a case-insensitive
   substring match of the group name against the Name column, then sum
   that row's monthly columns (assumed to be the 12 columns immediately
   before the sheet's "Total" column) for the same reason as above. No
   margin is available in this mode (that mirrors the sample workbook,
   whose oldest historical column has no margin).

Every result records which strategy was used (or "no_sheet" /
"not_found") so the validation report can surface exactly how much of
the historical data is a hard match vs. a best-effort guess.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import config
from excel_reader import (
    MasterWorkbook,
    ColumnMap,
    ProjectRow,
    SheetNotFoundError,
    build_column_map,
    read_project_rows,
    normalize_name,
    normalize_header,
    as_number,
    is_blank,
)

logger = logging.getLogger("sfae.historical_lookup")


@dataclass
class HistoricalResult:
    total_revenue: Optional[float]
    total_margin: Optional[float]
    method: str  # "group_match" | "fuzzy_name_match" | "not_found" | "no_sheet" | "duplicate_fuzzy_claim"
    sheet_name: Optional[str]


class HistoricalLookup:
    def __init__(self, master: MasterWorkbook):
        self.master = master
        self._sheet_cache: Dict[int, Optional[Tuple[str, ColumnMap, List[ProjectRow]]]] = {}
        self._fuzzy_cache: Dict[int, Optional[Tuple[str, int, int, List[int]]]] = {}
        # Sheets with no Group/Sub-Group column (e.g. a legacy year) can't
        # tell two of today's sections apart for the same customer name
        # (e.g. "Databricks" under Projects vs. under Staffing). We track
        # which (year, normalised name) pairs have already been credited
        # to a section so the same legacy total is never handed out twice.
        self._fuzzy_claimed: Dict[int, set] = {}

    # ------------------------------------------------------------------
    def _load_group_sheet(self, year: int):
        if year in self._sheet_cache:
            return self._sheet_cache[year]
        try:
            sheet_name = self.master.main_sheet_name(year)
        except SheetNotFoundError:
            self._sheet_cache[year] = None
            return None

        ws = self.master.sheet(sheet_name)
        try:
            cmap = build_column_map(ws)
        except Exception:
            # Legacy sheet with no proper Name+Group header row at all;
            # the fuzzy name-matching fallback will handle it instead.
            self._sheet_cache[year] = None
            return None
        if cmap.group is None:
            self._sheet_cache[year] = None
            return None
        rows = read_project_rows(ws, cmap)
        result = (sheet_name, cmap, rows)
        self._sheet_cache[year] = result
        return result

    def _load_fuzzy_sheet(self, year: int):
        """Locate (sheet, name_col, total_col, monthly_cols) for a
        legacy-layout sheet that has no Group column at all.

        Every column is still identified by its own header text where
        one exists (Total, and each monthly column's "Actual"/"Forecast"
        role label one row above the field header -- the same
        role-keyword check `build_column_map` uses for the modern
        sheets). The one column this specific legacy layout gives no
        header at all -- the project Name column -- cannot be found by
        header text by definition, so it is identified by its CONTENT
        instead: among every blank-header column, the Name column is
        the one with the most non-blank data cells beneath it (a
        legacy sheet may have more than one blank-header column, e.g. a
        genuinely empty spacer column, but only the real Name column is
        densely populated with text). This is immune to column
        reordering -- unlike an earlier version of this method, which
        assumed Name always sat immediately after a "Status" column and
        broke the one time a *different* labelled column happened to
        land in that position instead.
        """
        if year in self._fuzzy_cache:
            return self._fuzzy_cache[year]
        try:
            sheet_name = self.master.main_sheet_name(year)
        except SheetNotFoundError:
            self._fuzzy_cache[year] = None
            return None

        ws = self.master.sheet(sheet_name)
        for row in range(2, 15):
            values = [normalize_header(ws.cell(row, c).value) for c in range(1, ws.max_column + 1)]
            if "total" not in values:
                continue

            total_col = values.index("total") + 1
            name_col = None
            for c, v in enumerate(values, start=1):
                if v == "name":
                    name_col = c
                    break

            if name_col is None:
                # Content-based fallback: score every blank-header
                # column by how many non-blank data cells it has below
                # this header row, and take the best-populated one.
                best_col, best_count = None, 0
                for c in range(1, len(values) + 1):
                    if values[c - 1]:  # has a real header -> not a candidate
                        continue
                    count = sum(
                        1
                        for r in range(row + 1, ws.max_row + 1)
                        if not is_blank(ws.cell(r, c).value)
                    )
                    if count > best_count:
                        best_col, best_count = c, count
                name_col = best_col if best_col is not None else 1

            # Monthly columns: identified by their own role label
            # ("Actual"/"Forecast") one row above the field-header row,
            # exactly like the modern sheets -- never by position
            # relative to the Total column.
            type_row_values = [
                normalize_header(ws.cell(row - 1, c).value) for c in range(1, ws.max_column + 1)
            ]
            monthly_cols = [
                c
                for c, type_val in enumerate(type_row_values, start=1)
                if any(k in type_val for k in config.REVENUE_ROLE_KEYWORDS)
            ]
            if not monthly_cols:
                # Extremely defensive fallback for a layout with no
                # role-label row at all; keeps the tool working rather
                # than raising, at the cost of falling back to a
                # positional guess only in that unseen case.
                monthly_cols = [c for c in range(max(1, total_col - 12), total_col)]

            self._fuzzy_cache[year] = (sheet_name, name_col, total_col, monthly_cols)
            return self._fuzzy_cache[year]

        self._fuzzy_cache[year] = None
        return None

    def recompute_q4(self, year: int, group_name: str, ds_codes: List[int]) -> Optional[float]:
        """Sum just the Oct+Nov+Dec revenue for `group_name` from
        `year`'s own sheet -- used purely as an independent drift
        detector (see aggregator.attach_historical). Returns None if the
        group can't be found in that year's sheet at all (a legacy sheet
        with no Group column is not supported here, since the drift
        check is only meaningful where we can precisely re-scope by
        Group + Sub-Group)."""
        loaded = self._load_group_sheet(year)
        if loaded is None:
            return None
        _, _, rows = loaded
        target = normalize_name(group_name)
        matched = [r for r in rows if r.ds_code in ds_codes and normalize_name(r.group) == target]
        if not matched:
            return None
        total = 0.0
        for r in matched:
            for month in (10, 11, 12):
                total += r.monthly_revenue.get(month, 0.0)
        return round(total, 2)

    # ------------------------------------------------------------------
    def recompute(self, year: int, group_name: str, ds_codes: List[int]) -> HistoricalResult:
        """Independently recompute Total Revenue / Total Margin for
        `group_name` (scoped to `ds_codes`) from `year`'s own sheet."""
        loaded = self._load_group_sheet(year)
        if loaded is not None:
            sheet_name, cmap, rows = loaded
            target = normalize_name(group_name)
            matched = [
                r for r in rows
                if r.ds_code in ds_codes and normalize_name(r.group) == target
            ]
            if matched:
                total_rev = sum(sum(r.monthly_revenue.values()) for r in matched)
                total_marg = sum(sum(r.monthly_margin.values()) for r in matched)
                return HistoricalResult(round(total_rev, 2), round(total_marg, 2), "group_match", sheet_name)
            return HistoricalResult(None, None, "not_found", sheet_name)

        fuzzy = self._load_fuzzy_sheet(year)
        if fuzzy is None:
            return HistoricalResult(None, None, "no_sheet", None)

        sheet_name, name_col, total_col, monthly_cols = fuzzy
        ws = self.master.sheet(sheet_name)
        target = normalize_name(group_name)
        if not target:
            return HistoricalResult(None, None, "not_found", sheet_name)

        claimed = self._fuzzy_claimed.setdefault(year, set())
        if target in claimed:
            # Another section already claimed this legacy, un-coded total
            # for the same customer name this run; Rule 6 treats the
            # remaining occurrence(s) as a numeric blank (0) rather than
            # double-counting the same historical dollars twice.
            return HistoricalResult(0.0, None, "duplicate_fuzzy_claim", sheet_name)

        total = 0.0
        found = False
        for r in range(1, ws.max_row + 1):
            name_val = ws.cell(r, name_col).value
            if is_blank(name_val):
                continue
            norm_name = normalize_name(name_val)
            if target in norm_name or norm_name in target:
                monthly_values = [ws.cell(r, c).value for c in monthly_cols]
                if all(is_blank(v) for v in monthly_values):
                    # A stray section-header/label row (e.g. a bare
                    # "Staffing" row) can accidentally satisfy the
                    # substring test against a longer target name; such
                    # a row never carries real project data, so it must
                    # not count as a match.
                    continue
                monthly_sum = sum(as_number(v) for v in monthly_values)
                # Safety net: if the monthly-column heuristic finds
                # nothing but the sheet's own Total cell is populated,
                # trust the Total cell rather than reporting a false 0.
                total += monthly_sum if monthly_sum else as_number(ws.cell(r, total_col).value)
                found = True

        if found:
            claimed.add(target)
            return HistoricalResult(round(total, 2), None, "fuzzy_name_match", sheet_name)
        return HistoricalResult(None, None, "not_found", sheet_name)
