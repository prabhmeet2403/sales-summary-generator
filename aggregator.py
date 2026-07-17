"""
aggregator.py
=============
Implements the core business rules from the project spec:

  Rule 1  - group all rows by Group, one Summary row per unique Group.
  Rule 2  - Q1..Q4 = sum of the relevant calendar months' revenue across
            every project row under that Group.
  Rule 3  - Total = Q1 + Q2 + Q3 + Q4 (written as an Excel formula in the
            output, exactly like the human-built workbook).
  Rule 4  - Margin = sum of every project's monthly Margin (Actual -
            Salary, or the sheet's own Margin column when present) across
            every project row under that Group.
  Rule 6  - numeric blanks are treated as 0; comment blanks stay blank.

Plus the additional rules inferred from comparing the master workbook to
the manually produced Summary workbook (see README.md for the full
write-up):
  - Rows are scoped to a Summary "section" by their Sub-Group / DS-code
    (see config.OUTPUT_SECTIONS).
  - A group is dropped entirely ("Skipped Blank Groups") when both its
    current-year revenue AND margin are exactly zero.
  - Groups with non-zero revenue are sorted alphabetically; groups with
    zero revenue (overhead/non-billable cost centres) are kept in the
    order they first appear in the source sheet and listed after the
    billable groups.
"""
from __future__ import annotations

import logging
from collections import Counter, OrderedDict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import config
from excel_reader import ProjectRow, normalize_name, as_number, is_blank
from comment_mapper import CommentMapper
from historical_lookup import HistoricalLookup
from validator import SectionStats

logger = logging.getLogger("sfae.aggregator")


def month_to_quarter(month: int) -> str:
    for q, months in config.QUARTER_MONTHS.items():
        if month in months:
            return q
    raise ValueError(f"Month {month} is not a valid calendar month (1-12).")


@dataclass
class GroupSummary:
    group_name: str
    ds_code: Optional[int]
    raw_sub_group: Optional[str]
    poc: Optional[str]
    quarters: Dict[str, float]
    # Per-quarter margin breakdown, computed from the exact same monthly
    # margin figures (and the same month_to_quarter() mapping) already
    # used to compute `total_margin` below -- added to support the
    # Summary workbook's Q1-Q4 "Margin" sub-columns without changing how
    # margin itself is calculated (see aggregate_section()).
    quarter_margins: Dict[str, float]
    total_revenue: float
    total_margin: float
    first_seen_row: int
    row_count: int
    comment: Optional[str] = None
    comment_method: str = "not_found"
    # year -> (embedded_reference_sum, was_any_row_populated)
    embedded_ref_totals: Dict[int, Tuple[float, bool]] = field(default_factory=dict)
    # year -> (embedded_Q4-only_reference_sum, was_any_row_populated) --
    # used purely as an independent drift detector, see attach_historical
    embedded_q4_totals: Dict[int, Tuple[float, bool]] = field(default_factory=dict)
    historical: Dict[int, Tuple[Optional[float], Optional[float]]] = field(default_factory=dict)
    historical_source: Dict[int, str] = field(default_factory=dict)
    has_renewal_confidence: bool = False


def _group_rows_in_order(rows: List[ProjectRow]) -> "OrderedDict[str, List[ProjectRow]]":
    """Group rows by normalised Group name while preserving first-seen
    order (needed for the non-billable/overhead ordering rule)."""
    grouped: "OrderedDict[str, List[ProjectRow]]" = OrderedDict()
    for r in rows:
        key = normalize_name(r.group)
        grouped.setdefault(key, []).append(r)
    return grouped


def aggregate_section(
    all_rows: List[ProjectRow],
    section: config.OutputSection,
    stats: SectionStats,
) -> List[GroupSummary]:
    """Apply Rules 1-4 & 6 to every row belonging to `section`."""
    section_rows = [r for r in all_rows if r.ds_code in section.ds_codes]
    if section.row_range is not None:
        # Only set for sections that share a DS-code with another
        # section (see config.OutputSection.row_range) -- an
        # additional, purely narrowing tiebreaker on top of the
        # DS-code match above, never applied to (and therefore never
        # changing behavior for) any section that doesn't need it.
        start, end = section.row_range
        section_rows = [r for r in section_rows if start <= r.row_index <= end]
    grouped = _group_rows_in_order(section_rows)

    summaries: List[GroupSummary] = []
    for _, group_rows in grouped.items():
        display_name = group_rows[0].group
        quarters = {q: 0.0 for q in config.QUARTER_ORDER}
        quarter_margins = {q: 0.0 for q in config.QUARTER_ORDER}
        total_margin = 0.0
        sheet_margin_sum = 0.0
        has_sheet_margin = False
        poc_counter: Counter = Counter()
        sub_group_counter: Counter = Counter()
        ds_code_counter: Counter = Counter()
        first_seen = min(r.row_index for r in group_rows)

        for r in group_rows:
            for month, revenue in r.monthly_revenue.items():
                quarters[month_to_quarter(month)] += revenue
            for month, margin in r.monthly_margin.items():
                quarter_margins[month_to_quarter(month)] += margin
                total_margin += margin
            if r.sheet_total_margin is not None:
                sheet_margin_sum += r.sheet_total_margin
                has_sheet_margin = True
            if r.poc:
                poc_counter[r.poc] += 1
            if r.sub_group_raw:
                sub_group_counter[r.sub_group_raw] += 1
            if r.ds_code is not None:
                ds_code_counter[r.ds_code] += 1

        total_margin = round(total_margin, 2)
        total_revenue = round(sum(quarters.values()), 2)
        quarters = {q: round(v, 2) for q, v in quarters.items()}
        quarter_margins = {q: round(v, 2) for q, v in quarter_margins.items()}

        # Rule 6 / "Skipped Blank Groups": drop groups with zero revenue
        # AND zero margin - i.e. genuinely no activity this year --
        # unless this section has opted out (see
        # config.OutputSection.skip_blank_groups), because for it a $0
        # group is expected, not a sign of no real activity.
        if (
            section.skip_blank_groups
            and abs(total_revenue) < config.ZERO_TOLERANCE
            and abs(total_margin) < config.ZERO_TOLERANCE
        ):
            stats.skipped_blank_groups.append(display_name)
            continue

        if (
            has_sheet_margin
            and abs(round(sheet_margin_sum, 2) - total_margin) > config.CROSS_CHECK_TOLERANCE
        ):
            stats.margin_cross_check_mismatches.append(
                f"{display_name} (computed {total_margin:,.2f} vs sheet {sheet_margin_sum:,.2f})"
            )

        poc = poc_counter.most_common(1)[0][0] if poc_counter else None
        raw_sub_group = sub_group_counter.most_common(1)[0][0] if sub_group_counter else None
        ds_code = ds_code_counter.most_common(1)[0][0] if ds_code_counter else None

        # Same-sheet embedded historical reference columns (e.g. a bare
        # "2024" header or a "2025_Total" header on THIS year's sheet).
        # Summed across the group's own rows, with `present` tracking so
        # a group whose reference is genuinely blank everywhere can fall
        # back to recomputing from that year's own sheet instead of
        # silently reporting 0 (see historical_lookup.py).
        embedded_ref_totals: Dict[int, Tuple[float, bool]] = {}
        years_seen = {y for r in group_rows for y in r.historical_refs_raw}
        for year in years_seen:
            total = 0.0
            present = False
            for r in group_rows:
                raw = r.historical_refs_raw.get(year)
                if raw is not None:
                    present = True
                    total += as_number(raw)
            embedded_ref_totals[year] = (round(total, 2), present)

        embedded_q4_totals: Dict[int, Tuple[float, bool]] = {}
        q4_years_seen = {y for r in group_rows for y in r.historical_q4_raw}
        for year in q4_years_seen:
            total = 0.0
            present = False
            for r in group_rows:
                raw = r.historical_q4_raw.get(year)
                if raw is not None:
                    present = True
                    total += as_number(raw)
            embedded_q4_totals[year] = (round(total, 2), present)

        has_renewal_confidence = any(r.has_renewal_confidence for r in group_rows)

        summaries.append(
            GroupSummary(
                group_name=display_name,
                ds_code=ds_code,
                raw_sub_group=raw_sub_group,
                poc=poc,
                quarters=quarters,
                quarter_margins=quarter_margins,
                total_revenue=total_revenue,
                total_margin=total_margin,
                first_seen_row=first_seen,
                row_count=len(group_rows),
                embedded_ref_totals=embedded_ref_totals,
                embedded_q4_totals=embedded_q4_totals,
                has_renewal_confidence=has_renewal_confidence,
            )
        )
        stats.groups_processed += 1

    return summaries


def sort_groups(groups: List[GroupSummary], sort_alphabetically: bool = True) -> List[GroupSummary]:
    """Billable groups (non-zero current-year revenue) are sorted
    alphabetically (when `sort_alphabetically` is True for this section);
    non-billable / overhead groups (zero revenue, but kept because they
    carry a non-zero margin) are always appended afterwards in their
    original source order. This mirrors the manually built Summary
    workbook, where the small Staffing block is left in source order
    while the larger Projects block is alphabetised."""
    billable = [g for g in groups if abs(g.total_revenue) > config.ZERO_TOLERANCE]
    non_billable = [g for g in groups if abs(g.total_revenue) <= config.ZERO_TOLERANCE]
    if sort_alphabetically:
        billable.sort(key=lambda g: normalize_name(g.group_name))
    else:
        billable.sort(key=lambda g: g.first_seen_row)
    non_billable.sort(key=lambda g: g.first_seen_row)
    return billable + non_billable


def attach_comments(groups: List[GroupSummary], mapper: CommentMapper, stats: SectionStats) -> None:
    """Rule 5: pull each group's comment from the ClientComments sheet.

    Only an EXACT (Sub-Group/DS-code, Client List) match is used. A
    same-named client that only appears under a *different* section's
    code in the comments sheet (e.g. IDBS logged only as a
    DS30_Projection comment while this group is a DS10_Secured Summary
    row) is deliberately left blank rather than borrowing that other
    section's narrative -- verified against the sample workbook, whose
    IDBS row has no comment even though the ClientComments sheet does
    have an IDBS entry filed under a different code.
    """
    for g in groups:
        comment, method = mapper.lookup(g.group_name, g.ds_code)
        if method == "exact":
            g.comment = comment
            g.comment_method = method
            stats.comments_matched += 1
        else:
            g.comment = None
            g.comment_method = method
            stats.missing_comments += 1


def attach_historical(
    groups: List[GroupSummary],
    lookup: HistoricalLookup,
    years: List[int],
    years_with_margin: List[int],
    ds_codes: List[int],
    stats: SectionStats,
) -> None:
    """Populate each group's prior-year Total/Margin figures.

    Strategy, per year, in priority order:

    1. TOTAL -- if the current sheet carries an embedded same-sheet
       reference column for that year (a bare "2024" or a "2025_Total"
       style header) and at least one of the group's rows has a
       non-blank value in it, trust the SUM of that reference column
       across the group's rows. This was reverse-engineered by
       cross-checking dozens of groups against the sample Summary and is
       the single best predictor of the sample's historical figures --
       far better than recomputing from that year's own sheet, because
       the reference column is what was actually captured at the time
       the business tracked that number, and the sheets keep getting
       lightly revised after the fact.
    2. TOTAL fallback -- if the reference column is entirely blank for
       every one of the group's rows (nothing to sum), recompute fresh
       from that year's own sheet instead of reporting a false zero.
    3. MARGIN -- always recomputed fresh from that year's own sheet
       (Group + Sub-Group/DS-code scoped), since no equivalent
       same-sheet margin reference column exists anywhere in the
       workbook.
    4. Manual overrides (`config.HISTORICAL_OVERRIDES`) always win, for
       the rare case where a human has confirmed the source workbook
       itself contains bad data (see config.py for the documented
       mechanism -- empty, and therefore inert, by default).

    A same-sheet "2025_Q4"-style reference also exists and was tested as
    an independent drift detector: when the full-year embedded reference
    disagrees with a fresh recomputation, does a Q4-only recomputation
    agree with the embedded Q4 reference? If so, the reasoning goes, the
    underlying sheet hasn't moved since the references were captured, so
    the full-year *reference* (not the recomputation) must be the one
    with the error. This held for "Databricks" (Q4 matches; the
    recomputed full-year figure turned out to be the correct one) -- but
    it does NOT generalise: "Icertis" has the identical signature (Q4
    matches, full-year reference and recomputation disagree by $300),
    yet there the embedded *reference* is the one the sample workbook
    agrees with, not the recomputation. With only a single Q4 checkpoint
    available there is no way to tell, in general, which of the two
    conflicting numbers is right -- both are equally consistent with "an
    earlier quarter changed after the snapshot was taken". Rather than
    ship a heuristic that silently trades a wrong cell for a different
    wrong cell, drift is always surfaced in the validation report as
    "Historical Reference vs Recompute" for a human to check, and the
    embedded reference (the better default in the large majority of
    cases) is what gets written.
    """
    for g in groups:
        for year in years:
            embedded_total, embedded_present = g.embedded_ref_totals.get(year, (0.0, False))
            recompute = lookup.recompute(year, g.group_name, ds_codes)

            if embedded_present:
                total = embedded_total
                source = "embedded_reference"
                if (
                    recompute.total_revenue is not None
                    and abs(recompute.total_revenue - embedded_total) > config.HISTORICAL_DRIFT_TOLERANCE
                ):
                    detail = (
                        f"{g.group_name} ({year} Total): embedded reference={embedded_total:,.2f} "
                        f"vs recomputed from {recompute.sheet_name or 'source'}={recompute.total_revenue:,.2f}"
                    )
                    # Purely diagnostic corroboration: does a fresh
                    # Q4-only recomputation agree with this sheet's own
                    # embedded Q4 reference? If yes, the underlying data
                    # is confirmed unchanged for at least that quarter,
                    # which narrows the disagreement down to an earlier
                    # quarter -- worth flagging, but NOT a reliable
                    # signal for which of the two full-year numbers is
                    # correct (verified: it points the right way for
                    # some groups and the wrong way for others), so it
                    # never changes which value gets written.
                    embedded_q4, embedded_q4_present = g.embedded_q4_totals.get(year, (0.0, False))
                    if embedded_q4_present:
                        recomputed_q4 = lookup.recompute_q4(year, g.group_name, ds_codes)
                        if recomputed_q4 is not None:
                            if abs(recomputed_q4 - embedded_q4) <= config.HISTORICAL_DRIFT_TOLERANCE:
                                detail += " [Q4 corroborates recompute - drift likely in an earlier quarter]"
                            else:
                                detail += " [Q4 also drifts - underlying sheet revised since reference was captured]"
                    stats.historical_drift.append(detail)
            elif recompute.total_revenue is not None:
                total = recompute.total_revenue
                source = recompute.method
            else:
                # No embedded reference AND no match at all when
                # recomputing from that year's own sheet. The sample
                # workbook distinguishes two cases here: a real tracked
                # client account with genuinely no revenue that year is
                # shown as an explicit 0 (e.g. a brand-new client such
                # as "Aldevron"), while an internal/overhead line that
                # was never itself a sales account is left visually
                # blank (e.g. "AI Forward", "Bench"). The one reliable,
                # data-driven signal that tracks this distinction is
                # whether the group carries a Renewal Confidence value
                # on any of its rows at all (real accounts always have
                # one -- even "0%" -- overhead lines never do).
                total = 0.0 if g.has_renewal_confidence else None
                if total is None:
                    stats.historical_missing.append(f"{g.group_name} ({year})")
                source = "not_found"

            margin: Optional[float] = None
            if year in years_with_margin:
                if recompute.total_margin is not None:
                    margin = recompute.total_margin
                elif total is not None:
                    # We already know this group had a tracked (even if
                    # zero) total for this year; default margin to 0
                    # too rather than leaving it blank.
                    margin = 0.0
                if margin is None:
                    stats.historical_missing.append(f"{g.group_name} ({year} Margin)")

            override_total = config.HISTORICAL_OVERRIDES.get((g.group_name, year, "total"))
            if override_total is not None:
                total = override_total
                source = "manual_override"
            override_margin = config.HISTORICAL_OVERRIDES.get((g.group_name, year, "margin"))
            if override_margin is not None:
                margin = override_margin

            g.historical[year] = (total, margin)
            g.historical_source[year] = source
