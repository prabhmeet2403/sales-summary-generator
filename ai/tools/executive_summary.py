"""
ai.tools.executive_summary
=============================
A concise, board-level summary of the generated Sales & Forecast
Summary, assembled entirely from figures other tools/modules already
compute:

- Revenue figures: ``AnalyticsEngine.kpi``/``rank`` over
  ``BusinessContext.groups_df`` -- the exact same engine calls
  ``ai.tools.revenue.RevenueAnalysisTool`` makes, not a second
  implementation of "total revenue" or "top client".
- Validation figures: ``BusinessContext.report`` (the Phase 1
  ``ValidationReport``) -- the exact same object
  ``ai.tools.validation.ValidationSummaryTool`` reads, using only its
  existing aggregate properties and per-section lists (the same
  ``len()``/``sum()`` readout pattern ``ValidationReport`` itself
  already uses for ``total_skipped_blank_groups`` etc.).

This tool performs no business calculation of its own: every number in
its output already existed on ``BusinessContext``/``ValidationReport``
before this tool ran. Its only job is arranging already-correct facts
into the five sections a management/board audience expects, and
deriving a short, fact-conditioned list of recommended follow-ups
(e.g. "review unmatched comments" only when unmatched comments > 0) --
never an invented observation.
"""

from __future__ import annotations

from typing import ClassVar, List

from ai.analytics.engine import AnalyticsEngine
from ai.context import BusinessContext
from ai.data.filters import Filter
from ai.tools.base import BaseTool, ToolCategory, ToolResult


class ExecutiveSummaryTool(BaseTool):
    """Assembles a concise executive/board-level summary from figures
    already computed elsewhere (revenue via ``AnalyticsEngine``,
    validation via the existing ``ValidationReport``)."""

    name: ClassVar[str] = "executive_summary"
    display_name: ClassVar[str] = "Executive Summary Tool"
    description: ClassVar[str] = (
        "Produces a concise, board-level summary of the entire generated Summary, combining "
        "revenue highlights, validation status, and key business observations into one report. "
        "Use this for requests like 'generate an executive summary', 'summarize this report', "
        "'give me a management summary', 'prepare a board summary', 'what should management "
        "know', or 'what are the key highlights' -- not for a single-metric question like "
        "'what is total revenue', which the Revenue Analysis Tool already answers directly."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {"type": "object", "properties": {}}

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        engine = AnalyticsEngine(context)
        no_filter = Filter()

        total_revenue = engine.kpi(context.groups_df, "revenue", no_filter)
        highest_ranked = engine.rank(context.groups_df, "revenue", no_filter, top_n=1)
        lowest_ranked = engine.rank(context.groups_df, "revenue", no_filter, top_n=1, ascending=True)

        highest_client = highest_ranked.iloc[0] if not highest_ranked.empty else None
        lowest_client = lowest_ranked.iloc[0] if not lowest_ranked.empty else None
        # Only show a distinct "lowest" line when it's actually a different
        # client than "highest" -- otherwise it's the same fact restated.
        show_lowest = (
            lowest_client is not None
            and highest_client is not None
            and lowest_client["group"] != highest_client["group"]
        )

        report = context.report

        lines: List[str] = ["## Executive Summary"]
        group_count = len(context.group_names())
        if report is not None:
            lines.append(
                f"This Summary covers {group_count} client(s) with total revenue of "
                f"${total_revenue:,.0f}. Generation {'completed successfully' if report.success else 'did not complete successfully'}, "
                f"processing {report.total_groups_processed} group(s) with "
                f"{len(report.warnings)} warning(s) and {len(report.errors)} error(s)."
            )
        else:
            lines.append(
                f"This Summary covers {group_count} client(s) with total revenue of ${total_revenue:,.0f}."
            )

        lines.append("\n## Revenue Highlights")
        lines.append(f"- Total Revenue: ${total_revenue:,.0f}")
        if highest_client is not None:
            lines.append(f"- Highest Revenue Client: {highest_client['group']} (${highest_client['revenue']:,.0f})")
        if show_lowest:
            lines.append(f"- Lowest Revenue Client: {lowest_client['group']} (${lowest_client['revenue']:,.0f})")

        recommendations: List[str] = []

        if report is not None:
            lines.append("\n## Validation Status")
            lines.append(f"- Groups Processed: {report.total_groups_processed}")
            lines.append(f"- Comments Matched: {report.total_comments_matched}")
            lines.append(f"- Warnings: {len(report.warnings)}")
            lines.append(f"- Errors: {len(report.errors)}")

            historical_missing_count = sum(len(s.historical_missing) for s in report.sections)
            historical_drift_count = sum(len(s.historical_drift) for s in report.sections)

            lines.append("\n## Key Business Observations")
            if highest_client is not None:
                lines.append(f"- Largest client by revenue: {highest_client['group']}")
            lines.append(f"- Missing comments: {report.total_missing_comments}")
            if historical_missing_count:
                lines.append(f"- Historical records missing: {historical_missing_count}")
            if historical_drift_count:
                lines.append(f"- Historical reference vs. recompute drift flagged: {historical_drift_count}")
            if report.unmapped_sub_groups:
                total_unmapped_rows = sum(report.unmapped_sub_groups.values())
                lines.append(f"- Unmapped sub-groups excluded from this Summary: {total_unmapped_rows} row(s)")
            if report.total_skipped_blank_groups:
                lines.append(f"- Skipped blank groups: {report.total_skipped_blank_groups}")

            if report.total_missing_comments:
                recommendations.append("Review unmatched client comments before finalizing the report.")
            if historical_missing_count:
                recommendations.append("Investigate missing historical records, which may affect year-over-year comparisons.")
            if historical_drift_count:
                recommendations.append("Verify the flagged historical reference vs. recompute variances before distributing this report.")
            if report.unmapped_sub_groups:
                recommendations.append("Review unmapped sub-groups excluded from this Summary to confirm nothing material was left out.")
            if report.warnings:
                recommendations.append("Review the generation warnings before distributing this report.")
            if report.total_skipped_blank_groups:
                recommendations.append("Review skipped blank groups for completeness.")

        if len(recommendations) < 3:
            recommendations.append("Confirm revenue and margin figures align with financial reporting expectations before distribution.")
        if len(recommendations) < 3:
            recommendations.append("No material issues were flagged in this generation; a final review is still recommended before distribution.")
        recommendations = recommendations[:5]

        lines.append("\n## Recommended Management Attention")
        lines.extend(f"- {item}" for item in recommendations)

        summary = "\n".join(lines)
        raw = {
            "total_revenue": total_revenue,
            "highest_revenue_client": highest_client["group"] if highest_client is not None else None,
            "lowest_revenue_client": lowest_client["group"] if show_lowest else None,
            "groups_processed": report.total_groups_processed if report else None,
            "comments_matched": report.total_comments_matched if report else None,
            "warning_count": len(report.warnings) if report else None,
            "error_count": len(report.errors) if report else None,
            "recommendations": recommendations,
        }
        return ToolResult(summary=summary, raw=raw)
