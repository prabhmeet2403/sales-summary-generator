"""
ai.data.frames
===============
Builds pandas DataFrames as a query-optimized view over Phase 1's
already-computed ``GroupSummary`` and ``MonthlyGroupSummary`` objects.

Every column produced here is either a direct attribute readout or a
direct dictionary-value readout of an existing Phase 1 field -- there is
no arithmetic, aggregation, or business rule anywhere in this module.
The column-to-source mapping is documented on each builder function
specifically so it can be audited at a glance, and is mechanically
checked by ``tests/test_ai_dataframe_parity.py`` against the real
fixture workbook on every change to this file or to
``aggregator.py``/``monthly_view.py``.

A note on missing values: pandas represents a Python ``None`` in an
object-dtype column (``comment``, ``poc``) as ``NaN`` (a float) once the
column is read back from the DataFrame -- this is standard pandas
behavior, not a data-loss bug. Code reading these columns (including
``ai.data.filters``) should check for missingness with ``pandas.isna()``,
which correctly treats both ``None`` and ``NaN`` as "missing," rather
than an ``is None`` comparison, which will not match a value pandas has
silently converted to ``NaN``.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 3 for the
approved design.
"""

from __future__ import annotations

import logging
from typing import Dict, List, Optional, Tuple

import pandas as pd

from aggregator import GroupSummary, month_to_quarter
from config import OutputSection
from monthly_view import MonthlyGroupSummary

logger = logging.getLogger(__name__)

#: Column names shared by both DataFrames for the same underlying
#: concept (revenue, margin), so filtering code (``ai.data.filters``)
#: can apply the same clause to either DataFrame without needing to
#: know which one it received.
COLUMN_REVENUE = "revenue"
COLUMN_MARGIN = "margin"


def _parse_confidence_pct(raw_confidence: Optional[str]) -> Optional[float]:
    """Parse a display-formatted confidence string (e.g. ``"75%"``) into
    a plain float (``75.0``).

    This is a pure string-format conversion -- the same number, a
    different representation -- not a recalculation of the underlying
    confidence value, which is why it lives here rather than being
    treated as new business logic.

    Args:
        raw_confidence: The raw string as already produced by Phase 1
            (``MonthlyGroupSummary.confidence``), or ``None``.

    Returns:
        The parsed percentage as a float, or ``None`` if ``raw_confidence``
        is ``None`` or not parseable as a percentage.
    """
    if raw_confidence is None:
        return None
    text = raw_confidence.strip().rstrip("%").strip()
    try:
        return float(text)
    except ValueError:
        logger.debug("Could not parse confidence value %r as a percentage; treating as missing.", raw_confidence)
        return None


def build_groups_dataframe(
    section_results: List[Tuple[OutputSection, List[GroupSummary]]],
    target_year: int,
) -> pd.DataFrame:
    """Build one row per ``GroupSummary``, across every section.

    Column-to-source mapping (every column is a direct readout, no
    computation):

    ==================  =============================================
    Column              Source
    ==================  =============================================
    section             ``OutputSection.key``
    section_title       ``OutputSection.title``
    group               ``GroupSummary.group_name``
    poc                 ``GroupSummary.poc``
    ds_code             ``GroupSummary.ds_code``
    year                the ``target_year`` argument (constant per call)
    q1_revenue..q4_revenue  ``GroupSummary.quarters["Q1"]``..``["Q4"]``
    q1_margin..q4_margin    ``GroupSummary.quarter_margins["Q1"]``..``["Q4"]``
    revenue             ``GroupSummary.total_revenue``
    margin              ``GroupSummary.total_margin``
    comment             ``GroupSummary.comment``
    has_renewal_confidence  ``GroupSummary.has_renewal_confidence``
    ==================  =============================================

    Args:
        section_results: The same structure Worksheet 1 is built from
            (see ``summary_writer.SummaryWriter.build``'s ``sections``
            parameter) -- attached to ``GenerationResult.section_results``.
        target_year: The Summary's target year, as already resolved by
            Phase 1 and attached to ``GenerationResult.target_year``.

    Returns:
        A DataFrame with one row per group, per section. Empty (with
        the correct columns, zero rows) if ``section_results`` is empty.
    """
    records: List[dict] = []
    for section, groups in section_results:
        for group in groups:
            records.append(
                {
                    "section": section.key,
                    "section_title": section.title,
                    "group": group.group_name,
                    "poc": group.poc,
                    "ds_code": group.ds_code,
                    "year": target_year,
                    "q1_revenue": group.quarters.get("Q1", 0.0),
                    "q2_revenue": group.quarters.get("Q2", 0.0),
                    "q3_revenue": group.quarters.get("Q3", 0.0),
                    "q4_revenue": group.quarters.get("Q4", 0.0),
                    "q1_margin": group.quarter_margins.get("Q1", 0.0),
                    "q2_margin": group.quarter_margins.get("Q2", 0.0),
                    "q3_margin": group.quarter_margins.get("Q3", 0.0),
                    "q4_margin": group.quarter_margins.get("Q4", 0.0),
                    COLUMN_REVENUE: group.total_revenue,
                    COLUMN_MARGIN: group.total_margin,
                    "comment": group.comment,
                    "has_renewal_confidence": group.has_renewal_confidence,
                }
            )

    columns = [
        "section", "section_title", "group", "poc", "ds_code", "year",
        "q1_revenue", "q2_revenue", "q3_revenue", "q4_revenue",
        "q1_margin", "q2_margin", "q3_margin", "q4_margin",
        COLUMN_REVENUE, COLUMN_MARGIN, "comment", "has_renewal_confidence",
    ]
    frame = pd.DataFrame.from_records(records, columns=columns)
    logger.debug("Built groups DataFrame: %d rows across %d section(s).", len(frame), len(section_results))
    return frame


def build_monthly_dataframe(
    monthly_section_results: List[Tuple[OutputSection, List[MonthlyGroupSummary]]],
    month_roles: Dict[int, str],
    target_year: int,
) -> pd.DataFrame:
    """Build one row per (group, month), across every section.

    Column-to-source mapping (every column is a direct readout; the one
    exception -- ``confidence_pct`` -- is a pure string-to-float parse
    of ``confidence``, documented on :func:`_parse_confidence_pct`, not
    a recalculation):

    ==================  =============================================
    Column              Source
    ==================  =============================================
    section             ``OutputSection.key``
    section_title       ``OutputSection.title``
    group               ``MonthlyGroupSummary.group_name``
    poc                 ``MonthlyGroupSummary.poc``
    year                the ``target_year`` argument (constant per call)
    month               the dictionary key of ``monthly_revenue``/``monthly_margin`` (1-12)
    quarter             ``aggregator.month_to_quarter(month)`` -- reused, not reimplemented
    role                ``month_roles[month]`` (already-resolved Actual/Forecast label)
    revenue             ``MonthlyGroupSummary.monthly_revenue[month]``
    margin              ``MonthlyGroupSummary.monthly_margin[month]``
    confidence          ``MonthlyGroupSummary.confidence`` (raw display string)
    confidence_pct      parsed from ``confidence`` (see :func:`_parse_confidence_pct`)
    comment             ``MonthlyGroupSummary.comment``
    ==================  =============================================

    Args:
        monthly_section_results: The same structure Worksheet 2 is
            built from -- attached to
            ``GenerationResult.monthly_section_results``.
        month_roles: Already-resolved ``{month: "Actual"|"Forecast"}``
            mapping -- attached to ``GenerationResult.month_roles``.
        target_year: The Summary's target year.

    Returns:
        A DataFrame with one row per group per month, per section.
        Empty (with the correct columns, zero rows) if
        ``monthly_section_results`` is empty.
    """
    records: List[dict] = []
    for section, groups in monthly_section_results:
        for group in groups:
            for month, revenue in sorted(group.monthly_revenue.items()):
                margin = group.monthly_margin.get(month, 0.0)
                records.append(
                    {
                        "section": section.key,
                        "section_title": section.title,
                        "group": group.group_name,
                        "poc": group.poc,
                        "year": target_year,
                        "month": month,
                        "quarter": month_to_quarter(month),
                        "role": month_roles.get(month, "Actual"),
                        COLUMN_REVENUE: revenue,
                        COLUMN_MARGIN: margin,
                        "confidence": group.confidence,
                        "confidence_pct": _parse_confidence_pct(group.confidence),
                        "comment": group.comment,
                    }
                )

    columns = [
        "section", "section_title", "group", "poc", "year", "month", "quarter", "role",
        COLUMN_REVENUE, COLUMN_MARGIN, "confidence", "confidence_pct", "comment",
    ]
    frame = pd.DataFrame.from_records(records, columns=columns)
    logger.debug(
        "Built monthly DataFrame: %d rows across %d section(s).",
        len(frame), len(monthly_section_results),
    )
    return frame
