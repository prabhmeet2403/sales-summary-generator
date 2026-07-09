"""
ai.data.filters
================
The single, shared filtering mechanism every AI tool uses instead of
implementing its own filtering logic.

:func:`apply_filter` operates purely on the DataFrame layer
(``ai.data.frames``) via boolean-mask composition -- it never reads a
Phase 1 object directly and never performs a business calculation. A
:class:`Filter` field that has no corresponding column in the given
DataFrame is silently skipped (documented per-field below) rather than
raising, since the same :class:`Filter` instance is designed to be
reused across DataFrames of different shapes (e.g. the group-level
DataFrame has no ``month`` column; the monthly DataFrame does).

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 4 (and
Revision 2 Section 4, which first introduced this design) for the
approved design.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import List, Literal, Optional

import pandas as pd

logger = logging.getLogger(__name__)

#: The two Actual/Forecast role values Phase 1 produces (see
#: ``monthly_view.resolve_month_roles``). Not a closed enum in Phase 1
#: itself (it echoes whatever text the source workbook's own header
#: says), but these are the two values every fixture and every
#: production workbook seen so far actually uses.
Role = Literal["Actual", "Forecast"]


@dataclass
class Filter:
    """A structured, shareable description of "which subset of the
    business data" a tool call or dashboard view applies to.

    Every field is optional; an all-``None`` ``Filter`` matches every
    row. This same object is also the backbone of conversational state
    (Phase 2b's ``ConversationState``) -- "what the user means" and
    "what data gets queried" are represented by exactly one object so
    they cannot drift apart from each other.

    Attributes:
        client: Matches the ``group`` column, case-insensitively,
            exactly (not a substring match).
        poc: Matches the ``poc`` column, case-insensitively, exactly.
        section: Matches the ``section`` column exactly (this is the
            section's stable ``key``, e.g. ``"projects_track1"``, not
            its display title).
        quarters: Matches the ``quarter`` column (only present on the
            monthly-grain DataFrame; silently skipped on the group-grain
            DataFrame, which has no per-row quarter).
        months: Matches the ``month`` column (1-12; monthly-grain only).
        years: Matches the ``year`` column (present on both DataFrames).
        role: Matches the ``role`` column (``"Actual"`` or
            ``"Forecast"``; monthly-grain only).
        min_revenue: Inclusive lower bound on the ``revenue`` column.
        max_revenue: Inclusive upper bound on the ``revenue`` column.
        min_margin: Inclusive lower bound on the ``margin`` column.
        max_margin: Inclusive upper bound on the ``margin`` column.
        min_confidence_pct: Inclusive lower bound on the
            ``confidence_pct`` column (monthly-grain only).
    """

    client: Optional[str] = None
    poc: Optional[str] = None
    section: Optional[str] = None
    quarters: Optional[List[str]] = None
    months: Optional[List[int]] = None
    years: Optional[List[int]] = None
    role: Optional[Role] = None
    min_revenue: Optional[float] = None
    max_revenue: Optional[float] = None
    min_margin: Optional[float] = None
    max_margin: Optional[float] = None
    min_confidence_pct: Optional[float] = None

    def is_empty(self) -> bool:
        """Return ``True`` if no field is set (this filter matches every row)."""
        return all(value is None for value in vars(self).values())


def apply_filter(dataframe: pd.DataFrame, data_filter: Filter) -> pd.DataFrame:
    """Apply a :class:`Filter` to a DataFrame produced by ``ai.data.frames``.

    This is the *only* place filtering logic is implemented; every tool
    (from Phase 2b onward) is expected to call this rather than writing
    its own boolean conditions, so a new filter dimension or a bug fix
    here benefits every tool automatically.

    Args:
        dataframe: A DataFrame produced by
            :func:`ai.data.frames.build_groups_dataframe` or
            :func:`ai.data.frames.build_monthly_dataframe`.
        data_filter: The filter to apply.

    Returns:
        A new DataFrame containing only the matching rows. Returns
        ``dataframe`` unchanged (not a copy) when ``data_filter.is_empty()``,
        to avoid an unnecessary copy on the common "no filter" case.
    """
    if data_filter.is_empty():
        return dataframe

    mask = pd.Series(True, index=dataframe.index)

    if data_filter.client is not None:
        mask &= _case_insensitive_equals(dataframe, "group", data_filter.client)
    if data_filter.poc is not None:
        mask &= _case_insensitive_equals(dataframe, "poc", data_filter.poc)
    if data_filter.section is not None:
        mask &= _column_equals(dataframe, "section", data_filter.section)
    if data_filter.quarters is not None:
        mask &= _column_isin(dataframe, "quarter", data_filter.quarters)
    if data_filter.months is not None:
        mask &= _column_isin(dataframe, "month", data_filter.months)
    if data_filter.years is not None:
        mask &= _column_isin(dataframe, "year", data_filter.years)
    if data_filter.role is not None:
        mask &= _column_equals(dataframe, "role", data_filter.role)
    if data_filter.min_revenue is not None:
        mask &= _column_ge(dataframe, "revenue", data_filter.min_revenue)
    if data_filter.max_revenue is not None:
        mask &= _column_le(dataframe, "revenue", data_filter.max_revenue)
    if data_filter.min_margin is not None:
        mask &= _column_ge(dataframe, "margin", data_filter.min_margin)
    if data_filter.max_margin is not None:
        mask &= _column_le(dataframe, "margin", data_filter.max_margin)
    if data_filter.min_confidence_pct is not None:
        mask &= _column_ge(dataframe, "confidence_pct", data_filter.min_confidence_pct)

    return dataframe[mask]


def _warn_missing_column(column: str) -> None:
    logger.debug(
        "Filter field targeting column '%s' was set, but this DataFrame has no such "
        "column; that filter clause is being skipped for this DataFrame shape.",
        column,
    )


def _case_insensitive_equals(dataframe: pd.DataFrame, column: str, value: str) -> pd.Series:
    if column not in dataframe.columns:
        _warn_missing_column(column)
        return pd.Series(True, index=dataframe.index)
    return dataframe[column].astype(str).str.lower() == value.lower()


def _column_equals(dataframe: pd.DataFrame, column: str, value: object) -> pd.Series:
    if column not in dataframe.columns:
        _warn_missing_column(column)
        return pd.Series(True, index=dataframe.index)
    return dataframe[column] == value


def _column_isin(dataframe: pd.DataFrame, column: str, values: list) -> pd.Series:
    if column not in dataframe.columns:
        _warn_missing_column(column)
        return pd.Series(True, index=dataframe.index)
    return dataframe[column].isin(values)


def _column_ge(dataframe: pd.DataFrame, column: str, value: float) -> pd.Series:
    if column not in dataframe.columns:
        _warn_missing_column(column)
        return pd.Series(True, index=dataframe.index)
    return dataframe[column] >= value


def _column_le(dataframe: pd.DataFrame, column: str, value: float) -> pd.Series:
    if column not in dataframe.columns:
        _warn_missing_column(column)
        return pd.Series(True, index=dataframe.index)
    return dataframe[column] <= value
