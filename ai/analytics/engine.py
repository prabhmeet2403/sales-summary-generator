"""
ai.analytics.engine
=====================
The single implementation of generic, reusable analytical operations
(sum, rank, compare, trend, group-and-aggregate) every tool uses instead
of writing its own version of the same shaping logic.

``AnalyticsEngine`` never redefines what a number means -- revenue,
margin, and quarter are exactly what ``aggregator.py``/``monthly_view.py``
already computed. It only provides one canonical implementation of
generic operations (the same category of operation a spreadsheet
``SUM``/``RANK``/``GROUPBY`` formula performs) applied to the DataFrame
layer (``ai.data.frames``) through the Universal Filter Engine
(``ai.data.filters``).

Every method takes an explicit ``dataframe`` parameter rather than
assuming one fixed grain of data -- some analyses (a client's yearly
total) are naturally computed over ``BusinessContext.groups_df``, while
others (a quarter-over-quarter comparison) require row-level filtering
by quarter, which only ``BusinessContext.monthly_df`` supports (see
``ai.data.filters``'s own documented "quarters/months/role filter
fields are monthly-grain-only" behavior). Making this explicit, rather
than picking one DataFrame implicitly, is what lets this one engine
correctly serve both kinds of analysis.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 3.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Optional

import pandas as pd

from ai.context import BusinessContext
from ai.data.filters import Filter, apply_filter

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ComparisonResult:
    """The result of comparing one metric between two filtered slices.

    Attributes:
        label_a: A human-readable label for the first slice (e.g. ``"Q2"``).
        label_b: A human-readable label for the second slice (e.g. ``"Q3"``).
        value_a: The metric's total for the first slice.
        value_b: The metric's total for the second slice.
        delta: ``value_b - value_a``.
        percent_change: ``delta / value_a * 100``, or ``None`` if
            ``value_a`` is zero (a percentage change from zero is
            undefined, not "infinite" or "0%" -- reporting ``None`` and
            letting the caller state this explicitly is more honest
            than picking an arbitrary convention).
    """

    label_a: str
    label_b: str
    value_a: float
    value_b: float
    delta: float
    percent_change: Optional[float]


class AnalyticsEngine:
    """Generic analytical operations over the DataFrame layer.

    Attributes:
        context: retained only for tools that need direct access to
            context beyond what a single DataFrame operation requires
            (e.g. resolving a display name); the engine's own methods
            operate on whichever ``dataframe`` is passed to them.
    """

    def __init__(self, context: BusinessContext) -> None:
        self.context = context

    def kpi(self, dataframe: pd.DataFrame, metric: str, data_filter: Filter) -> float:
        """Sum ``metric`` across every row matching ``data_filter``.

        Args:
            dataframe: The DataFrame to read from (``groups_df`` or
                ``monthly_df``).
            metric: The column to sum (typically ``"revenue"`` or
                ``"margin"``).
            data_filter: Which rows to include.

        Returns:
            The sum, or ``0.0`` if no rows match or ``metric`` is not a
            column of ``dataframe``.
        """
        filtered = apply_filter(dataframe, data_filter)
        if filtered.empty or metric not in filtered.columns:
            return 0.0
        return float(filtered[metric].sum())

    def rank(
        self,
        dataframe: pd.DataFrame,
        metric: str,
        data_filter: Filter,
        *,
        top_n: int = 10,
        ascending: bool = False,
    ) -> pd.DataFrame:
        """Return the top (or bottom) ``top_n`` rows by ``metric``.

        Args:
            dataframe: The DataFrame to read from.
            metric: The column to rank by.
            data_filter: Which rows to consider before ranking.
            top_n: How many rows to return.
            ascending: If ``True``, return the *lowest* ``metric``
                values instead of the highest.

        Returns:
            A DataFrame of at most ``top_n`` rows, sorted by ``metric``.
            Empty if no rows match or ``metric`` is not a column.
        """
        filtered = apply_filter(dataframe, data_filter)
        if filtered.empty or metric not in filtered.columns:
            return filtered
        return filtered.nsmallest(top_n, metric) if ascending else filtered.nlargest(top_n, metric)

    def compare(
        self,
        dataframe: pd.DataFrame,
        metric: str,
        filter_a: Filter,
        filter_b: Filter,
        *,
        label_a: str = "A",
        label_b: str = "B",
    ) -> ComparisonResult:
        """Compare ``metric``'s total between two filtered slices of the
        same DataFrame.

        Args:
            dataframe: The DataFrame to read from.
            metric: The column to compare.
            filter_a: The first slice.
            filter_b: The second slice.
            label_a: Display label for the first slice.
            label_b: Display label for the second slice.

        Returns:
            A :class:`ComparisonResult`.
        """
        value_a = self.kpi(dataframe, metric, filter_a)
        value_b = self.kpi(dataframe, metric, filter_b)
        delta = value_b - value_a
        percent_change = (delta / value_a * 100.0) if value_a != 0 else None
        return ComparisonResult(
            label_a=label_a, label_b=label_b, value_a=value_a, value_b=value_b,
            delta=delta, percent_change=percent_change,
        )

    def trend(self, dataframe: pd.DataFrame, metric: str, data_filter: Filter, *, group_by: str) -> pd.DataFrame:
        """Sum ``metric`` grouped by ``group_by`` (e.g. ``"month"``),
        after applying ``data_filter``.

        Args:
            dataframe: The DataFrame to read from (typically
                ``monthly_df`` when grouping by month or quarter).
            metric: The column to sum within each group.
            data_filter: Which rows to include before grouping.
            group_by: The column to group by (e.g. ``"month"``,
                ``"quarter"``).

        Returns:
            A DataFrame with one row per distinct ``group_by`` value and
            a ``metric`` column holding that group's sum. Empty if no
            rows match or either column is absent.
        """
        filtered = apply_filter(dataframe, data_filter)
        if filtered.empty or metric not in filtered.columns or group_by not in filtered.columns:
            return pd.DataFrame(columns=[group_by, metric])
        return filtered.groupby(group_by, as_index=False)[metric].sum()

    def aggregate(self, dataframe: pd.DataFrame, metric: str, group_by: str, data_filter: Filter) -> pd.DataFrame:
        """Sum ``metric`` grouped by ``group_by`` (e.g. ``"poc"``,
        ``"section"``), after applying ``data_filter``.

        This is functionally identical to :meth:`trend` (both are a
        filtered group-and-sum); kept as a separate, identically-named
        method per the approved architecture so callers reading tool
        code can tell "this is a breakdown by category" (``aggregate``)
        from "this is a breakdown over time" (``trend``) at a glance,
        even though the underlying operation is the same.

        Args:
            dataframe: The DataFrame to read from.
            metric: The column to sum within each group.
            group_by: The column to group by.
            data_filter: Which rows to include before grouping.

        Returns:
            A DataFrame with one row per distinct ``group_by`` value.
        """
        return self.trend(dataframe, metric, data_filter, group_by=group_by)
