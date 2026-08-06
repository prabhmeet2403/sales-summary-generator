"""
ai.tools.revenue
=================
Revenue analysis: a single total, or a ranking of groups by revenue,
over any combination of the shared filter fields.

Reads ``BusinessContext.groups_df``/``monthly_df`` through
``AnalyticsEngine`` -- no revenue figure is computed here; every number
already exists on ``GroupSummary``/``MonthlyGroupSummary`` (see
``ai.data.frames``'s column-to-source mapping).
"""

from __future__ import annotations

from typing import ClassVar

from ai.analytics.engine import AnalyticsEngine
from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolResult
from ai.tools.schemas import FILTER_PROPERTIES, filter_from_arguments


class RevenueAnalysisTool(BaseTool):
    """Reports total revenue, or a top-N ranking by revenue, for a
    filtered slice of the business data."""

    name: ClassVar[str] = "revenue_analysis"
    display_name: ClassVar[str] = "Revenue Analysis Tool"
    description: ClassVar[str] = (
        "Reports total revenue for a filtered set of clients, or ranks clients by revenue. "
        "Use this for questions like 'what is total revenue', 'what is HPE's revenue', "
        "'top 5 clients by revenue', or 'revenue for the Staffing section in Q2'."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            **FILTER_PROPERTIES,
            "top_n": {
                "type": "integer",
                "description": "If set, return the top N clients ranked by revenue instead of a single total.",
            },
        },
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        data_filter = filter_from_arguments(arguments)
        dataframe = context.monthly_df if data_filter.quarters else context.groups_df
        engine = AnalyticsEngine(context)
        top_n = arguments.get("top_n")

        if top_n:
            ranked = engine.rank(dataframe, "revenue", data_filter, top_n=int(top_n))
            if ranked.empty:
                return ToolResult(summary="No matching clients were found for this request.")
            lines = [f"{row['group']}: ${row['revenue']:,.0f}" for _, row in ranked.iterrows()]
            summary = f"Top {len(ranked)} clients by revenue:\n" + "\n".join(lines)
            return ToolResult(summary=summary, raw={"ranking": ranked[["group", "revenue"]].to_dict("records")})

        total = engine.kpi(dataframe, "revenue", data_filter)
        scope = data_filter.client or data_filter.section or "the filtered scope"
        summary = f"Total revenue for {scope}: ${total:,.0f}."
        return ToolResult(summary=summary, raw={"revenue": total})
