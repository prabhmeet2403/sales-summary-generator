"""
ai.tools.comparison
=====================
Compares a metric (revenue or margin) between two quarters, optionally
scoped to a specific client, POC, or section.

Operates on ``BusinessContext.monthly_df``, since quarter-level
row-filtering is only meaningful at that grain (see
``ai.data.filters``'s documented behavior: the ``quarters`` filter field
is a no-op on ``groups_df``, which has no per-row quarter).
"""

from __future__ import annotations

from dataclasses import replace
from typing import ClassVar

from ai.analytics.engine import AnalyticsEngine
from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolError, ToolResult
from ai.tools.schemas import FILTER_PROPERTIES, filter_from_arguments

_VALID_QUARTERS = {"Q1", "Q2", "Q3", "Q4"}
_VALID_METRICS = {"revenue", "margin"}


class QuarterComparisonTool(BaseTool):
    """Compares revenue or margin between two quarters."""

    name: ClassVar[str] = "quarter_comparison"
    display_name: ClassVar[str] = "Quarter Comparison Tool"
    description: ClassVar[str] = (
        "Compares revenue or margin between two quarters (e.g. Q2 vs Q3), optionally "
        "scoped to a single client, POC, or section. Use this for questions like "
        "'compare Q2 and Q3 for HPE' or 'how did margin change from Q1 to Q2'."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            **{k: v for k, v in FILTER_PROPERTIES.items() if k != "quarters"},
            "quarter_a": {"type": "string", "enum": ["Q1", "Q2", "Q3", "Q4"], "description": "The first (baseline) quarter."},
            "quarter_b": {"type": "string", "enum": ["Q1", "Q2", "Q3", "Q4"], "description": "The second (comparison) quarter."},
            "metric": {"type": "string", "enum": ["revenue", "margin"], "description": "Which metric to compare. Defaults to revenue."},
        },
        "required": ["quarter_a", "quarter_b"],
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        quarter_a = arguments.get("quarter_a")
        quarter_b = arguments.get("quarter_b")
        metric = arguments.get("metric", "revenue")

        if quarter_a not in _VALID_QUARTERS or quarter_b not in _VALID_QUARTERS:
            raise ToolError(f"quarter_a and quarter_b must each be one of {sorted(_VALID_QUARTERS)}.")
        if metric not in _VALID_METRICS:
            raise ToolError(f"metric must be one of {sorted(_VALID_METRICS)}.")

        base_filter = filter_from_arguments(arguments)
        filter_a = replace(base_filter, quarters=[quarter_a])
        filter_b = replace(base_filter, quarters=[quarter_b])

        engine = AnalyticsEngine(context)
        comparison = engine.compare(
            context.monthly_df, metric, filter_a, filter_b, label_a=quarter_a, label_b=quarter_b
        )

        scope = base_filter.client or base_filter.poc or base_filter.section or "the filtered scope"
        if comparison.percent_change is None:
            change_text = f"a change of ${comparison.delta:,.0f} (baseline was $0, percent change is undefined)"
        else:
            direction = "increase" if comparison.delta >= 0 else "decrease"
            change_text = f"a {abs(comparison.percent_change):.1f}% {direction}"

        summary = (
            f"{metric.capitalize()} for {scope}: "
            f"{quarter_a} = ${comparison.value_a:,.0f}, {quarter_b} = ${comparison.value_b:,.0f} "
            f"({change_text})."
        )
        return ToolResult(
            summary=summary,
            raw={
                "quarter_a": quarter_a, "value_a": comparison.value_a,
                "quarter_b": quarter_b, "value_b": comparison.value_b,
                "delta": comparison.delta, "percent_change": comparison.percent_change,
            },
        )
