"""
ai.tools.margin
=================
Margin analysis: a single total, or a ranking of groups by margin, over
any combination of the shared filter fields. Mirrors
``ai.tools.revenue.RevenueAnalysisTool`` exactly, for the ``margin``
metric instead of ``revenue``.
"""

from __future__ import annotations

from typing import ClassVar

from ai.analytics.engine import AnalyticsEngine
from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolResult
from ai.tools.schemas import FILTER_PROPERTIES, filter_from_arguments


class MarginAnalysisTool(BaseTool):
    """Reports total margin, or a top-N ranking by margin, for a
    filtered slice of the business data."""

    name: ClassVar[str] = "margin_analysis"
    display_name: ClassVar[str] = "Margin Analysis Tool"
    description: ClassVar[str] = (
        "Reports total margin for a filtered set of clients, or ranks clients by margin. "
        "Use this for questions like 'what is total margin', 'what is HPE's margin', "
        "'top 5 clients by margin', or 'lowest margin clients'."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            **FILTER_PROPERTIES,
            "top_n": {
                "type": "integer",
                "description": "If set, return the top N clients ranked by margin instead of a single total.",
            },
            "ascending": {
                "type": "boolean",
                "description": "If true with top_n set, return the LOWEST margin clients instead of the highest.",
            },
        },
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        data_filter = filter_from_arguments(arguments)
        dataframe = context.monthly_df if data_filter.quarters else context.groups_df
        engine = AnalyticsEngine(context)
        top_n = arguments.get("top_n")
        ascending = bool(arguments.get("ascending", False))

        if top_n:
            ranked = engine.rank(dataframe, "margin", data_filter, top_n=int(top_n), ascending=ascending)
            if ranked.empty:
                return ToolResult(summary="No matching clients were found for this request.")
            direction = "lowest" if ascending else "highest"
            lines = [f"{row['group']}: ${row['margin']:,.0f}" for _, row in ranked.iterrows()]
            summary = f"{len(ranked)} clients with the {direction} margin:\n" + "\n".join(lines)
            return ToolResult(summary=summary, raw={"ranking": ranked[["group", "margin"]].to_dict("records")})

        total = engine.kpi(dataframe, "margin", data_filter)
        scope = data_filter.client or data_filter.section or "the filtered scope"
        summary = f"Total margin for {scope}: ${total:,.0f}."
        return ToolResult(summary=summary, raw={"margin": total})
