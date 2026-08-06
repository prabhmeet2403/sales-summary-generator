"""
ai.tools.validation
=====================
Reports on this generation's own process/quality metrics -- the same
values the "Upload & Generate" dashboard already shows in its
Validation Summary KPI cards and "Full Validation Report" expander
(Groups Processed, Comments Matched, Missing Comments, Skipped Blank
Groups, Warnings, Errors, Time Taken) -- by reading the Phase 1
``ValidationReport`` already attached to ``BusinessContext.report``.

This tool never re-validates or recomputes anything: it is a pure
readout of ``ValidationReport``'s own existing fields and
``render()`` output, exactly like every other tool in this package is a
pure readout of ``groups_df``/``monthly_df``.
"""

from __future__ import annotations

from typing import ClassVar

from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolError, ToolResult


class ValidationSummaryTool(BaseTool):
    """Reports the generation's own validation/process metrics (as
    opposed to business data like revenue or margin)."""

    name: ClassVar[str] = "validation_summary"
    display_name: ClassVar[str] = "Validation Summary Tool"
    description: ClassVar[str] = (
        "Reports this Summary's generation/validation metrics: Groups Processed, "
        "Comments Matched, Missing Comments, Skipped Blank Groups, Warnings, Errors, "
        "and Time Taken -- the same figures shown on the Upload & Generate dashboard's "
        "Validation Summary cards and Full Validation Report. Use this for questions like "
        "'explain the validation summary', 'how many comments matched', 'how many warnings "
        "are there', or 'explain the validation report'. Set include_full_report=true only "
        "when the user explicitly wants the full/detailed report text rather than the summary "
        "figures."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            "include_full_report": {
                "type": "boolean",
                "description": (
                    "If true, include the full rendered validation report text (all "
                    "per-section detail, warnings, and errors) in addition to the "
                    "summary figures. Defaults to false."
                ),
            },
        },
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        report = context.report
        if report is None:
            raise ToolError("No validation report is available for this Summary.")

        include_full_report = bool(arguments.get("include_full_report", False))

        summary_lines = [
            f"Generation status: {'Successful' if report.success else 'Failed'}",
            f"Groups Processed: {report.total_groups_processed}",
            f"Comments Matched: {report.total_comments_matched}",
            f"Missing Comments: {report.total_missing_comments}",
            f"Skipped Blank Groups: {report.total_skipped_blank_groups}",
            f"Warnings: {len(report.warnings)}",
            f"Errors: {len(report.errors)}",
        ]
        if context.elapsed_seconds is not None:
            summary_lines.append(f"Time Taken: {context.elapsed_seconds:.1f} seconds")
        if report.warnings:
            summary_lines.append("Warning details: " + "; ".join(report.warnings))
        if report.errors:
            summary_lines.append("Error details: " + "; ".join(report.errors))
        if report.unmapped_sub_groups:
            unmapped = ", ".join(f"{code} ({count} row(s))" for code, count in sorted(report.unmapped_sub_groups.items()))
            summary_lines.append(f"Unmapped sub-groups: {unmapped}")

        raw = {
            "success": report.success,
            "groups_processed": report.total_groups_processed,
            "comments_matched": report.total_comments_matched,
            "missing_comments": report.total_missing_comments,
            "skipped_blank_groups": report.total_skipped_blank_groups,
            "warning_count": len(report.warnings),
            "error_count": len(report.errors),
            "warnings": list(report.warnings),
            "errors": list(report.errors),
            "elapsed_seconds": context.elapsed_seconds,
        }

        if include_full_report:
            summary_lines.append("")
            summary_lines.append("Full Validation Report:")
            summary_lines.append(report.render())
            raw["full_report"] = report.render()

        return ToolResult(summary="\n".join(summary_lines), raw=raw)
