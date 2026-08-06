"""
tests/test_ai_tools.py
========================
Tests all five Phase 2b tools (Revenue, Margin, Quarter Comparison,
Client Lookup, POC Lookup) directly against real fixture data, plus the
plugin registry's auto-discovery mechanism.

Usage:
    python tests/test_ai_tools.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.tools.base import ToolCategory, ToolError  # noqa: E402
from ai.tools.registry import UnknownToolError, discover_tools  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"

_EXPECTED_TOOLS = {
    "revenue_analysis", "margin_analysis", "quarter_comparison", "client_lookup", "poc_lookup",
    "validation_summary", "executive_summary",
}


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test tools.")
            return 1
        ctx = BusinessContext.from_generation_result(result, elapsed_seconds=12.5)
        registry = discover_tools()

        # --- registry auto-discovery ---
        discovered_names = {t.name for t in registry.all_tools()}
        if discovered_names != _EXPECTED_TOOLS:
            problems.append(f"Expected tools {_EXPECTED_TOOLS}, discovered {discovered_names}")
        for tool in registry.all_tools():
            if tool.category != ToolCategory.ANALYSIS:
                problems.append(f"Tool '{tool.name}' should be ANALYSIS category, got {tool.category}")
            if not tool.description:
                problems.append(f"Tool '{tool.name}' has an empty description")
            if not isinstance(tool.schema, dict) or "properties" not in tool.schema:
                problems.append(f"Tool '{tool.name}' has a malformed schema")

        # --- unknown tool dispatch raises UnknownToolError ---
        try:
            registry.dispatch("not_a_real_tool", {}, ctx)
            problems.append("Dispatching an unknown tool name should raise UnknownToolError")
        except UnknownToolError:
            pass

        # --- revenue_analysis: total for a known client ---
        expected_hpe_revenue = float(
            ctx.groups_df[(ctx.groups_df["group"] == "HPE") & (ctx.groups_df["section"] == "staffing_secured")]["revenue"].iloc[0]
        )
        rev_result = registry.dispatch("revenue_analysis", {"client": "HPE", "section": "staffing_secured"}, ctx)
        if f"{expected_hpe_revenue:,.0f}" not in rev_result.summary:
            problems.append(f"revenue_analysis summary did not contain the expected figure: {rev_result.summary}")
        if rev_result.raw["revenue"] != expected_hpe_revenue:
            problems.append("revenue_analysis raw payload did not match the real DataFrame value")

        # --- revenue_analysis: ranking ---
        rank_result = registry.dispatch("revenue_analysis", {"top_n": 3}, ctx)
        if rank_result.raw is None or len(rank_result.raw["ranking"]) != 3:
            problems.append("revenue_analysis with top_n=3 should return exactly 3 ranked entries")
        else:
            revenues = [row["revenue"] for row in rank_result.raw["ranking"]]
            if revenues != sorted(revenues, reverse=True):
                problems.append("revenue_analysis ranking is not sorted descending")

        # --- margin_analysis: ascending ranking (lowest margin) ---
        low_margin_result = registry.dispatch("margin_analysis", {"top_n": 3, "ascending": True}, ctx)
        margins = [row["margin"] for row in low_margin_result.raw["ranking"]]
        if margins != sorted(margins):
            problems.append("margin_analysis with ascending=True is not sorted ascending")

        # --- quarter_comparison: matches an independent Analytics Engine computation ---
        from ai.analytics.engine import AnalyticsEngine
        from ai.data.filters import Filter
        engine = AnalyticsEngine(ctx)
        expected_comparison = engine.compare(
            ctx.monthly_df, "revenue",
            Filter(client="HPE", section="staffing_secured", quarters=["Q2"]),
            Filter(client="HPE", section="staffing_secured", quarters=["Q3"]),
        )
        comparison_result = registry.dispatch(
            "quarter_comparison",
            {"quarter_a": "Q2", "quarter_b": "Q3", "client": "HPE", "section": "staffing_secured"},
            ctx,
        )
        if comparison_result.raw["value_a"] != expected_comparison.value_a:
            problems.append("quarter_comparison tool result did not match an independent AnalyticsEngine computation")

        # --- quarter_comparison: invalid quarter raises ToolError from run(),
        #     but dispatch() catches it into a graceful ToolResult ---
        try:
            registry.get("quarter_comparison").run({"quarter_a": "Q9", "quarter_b": "Q2"}, ctx)
            problems.append("quarter_comparison.run() with an invalid quarter should raise ToolError")
        except ToolError:
            pass
        graceful_quarter_result = registry.dispatch("quarter_comparison", {"quarter_a": "Q9", "quarter_b": "Q2"}, ctx)
        if "could not complete" not in graceful_quarter_result.summary:
            problems.append("dispatch() should catch ToolError into a graceful failure summary, not propagate it")

        # --- client_lookup: known client ---
        lookup_result = registry.dispatch("client_lookup", {"client": "aldevron"}, ctx)  # case-insensitive
        if lookup_result.raw["client"] != "Aldevron":
            problems.append(f"client_lookup did not resolve case-insensitively: {lookup_result.raw}")

        # --- client_lookup: unknown client is a graceful ToolResult via dispatch, not a raised exception ---
        graceful_result = registry.dispatch("client_lookup", {"client": "NotARealClient"}, ctx)
        if "could not complete" not in graceful_result.summary:
            problems.append("Unknown client lookup via dispatch() should produce a graceful failure summary, not raise")

        # --- poc_lookup: known POC ---
        poc_result = registry.dispatch("poc_lookup", {"poc": "Vijay"}, ctx)
        expected_poc_revenue = float(ctx.groups_df[ctx.groups_df["poc"] == "Vijay"]["revenue"].sum())
        if poc_result.raw["total_revenue"] != expected_poc_revenue:
            problems.append("poc_lookup total_revenue did not match an independent DataFrame sum")

        # --- required-argument validation: client_lookup/poc_lookup need
        #     their key argument -- run() raises ToolError directly ---
        try:
            registry.get("client_lookup").run({}, ctx)
            problems.append("client_lookup.run() with no client argument should raise ToolError")
        except ToolError:
            pass

        # --- validation_summary: figures match the real ValidationReport ---
        validation_result = registry.dispatch("validation_summary", {}, ctx)
        report = result.report
        if str(report.total_groups_processed) not in validation_result.summary:
            problems.append(f"validation_summary did not report the correct Groups Processed: {validation_result.summary}")
        if validation_result.raw["comments_matched"] != report.total_comments_matched:
            problems.append("validation_summary raw payload did not match the real ValidationReport comments_matched")
        if validation_result.raw["warning_count"] != len(report.warnings):
            problems.append("validation_summary raw payload did not match the real ValidationReport warning count")
        if validation_result.raw["elapsed_seconds"] != 12.5:
            problems.append("validation_summary did not surface the elapsed_seconds passed into BusinessContext")
        if "full_report" in validation_result.raw:
            problems.append("validation_summary should not include the full report unless requested")

        full_report_result = registry.dispatch("validation_summary", {"include_full_report": True}, ctx)
        if full_report_result.raw.get("full_report") != report.render():
            problems.append("validation_summary with include_full_report=True did not return the real rendered report")

        # --- validation_summary: missing report is a graceful ToolError via dispatch ---
        ctx_no_report = BusinessContext(
            target_year=ctx.target_year,
            prior_years=ctx.prior_years,
            groups_df=ctx.groups_df,
            monthly_df=ctx.monthly_df,
            fingerprint=ctx.fingerprint,
            report=None,
        )
        no_report_result = registry.dispatch("validation_summary", {}, ctx_no_report)
        if "could not complete" not in no_report_result.summary:
            problems.append("validation_summary with no report should produce a graceful failure summary via dispatch")

        # --- executive_summary: assembles from the same figures the other
        #     tools/AnalyticsEngine already produce -- no new calculation ---
        from ai.analytics.engine import AnalyticsEngine as _Engine
        from ai.data.filters import Filter as _Filter
        engine_check = _Engine(ctx)
        expected_total_revenue = engine_check.kpi(ctx.groups_df, "revenue", _Filter())
        expected_top = engine_check.rank(ctx.groups_df, "revenue", _Filter(), top_n=1)

        exec_result = registry.dispatch("executive_summary", {}, ctx)
        for expected_heading in (
            "## Executive Summary", "## Revenue Highlights", "## Validation Status",
            "## Key Business Observations", "## Recommended Management Attention",
        ):
            if expected_heading not in exec_result.summary:
                problems.append(f"executive_summary is missing the '{expected_heading}' section")

        if exec_result.raw["total_revenue"] != expected_total_revenue:
            problems.append("executive_summary total_revenue did not match an independent AnalyticsEngine.kpi computation")
        if not expected_top.empty and exec_result.raw["highest_revenue_client"] != expected_top.iloc[0]["group"]:
            problems.append("executive_summary highest_revenue_client did not match an independent AnalyticsEngine.rank computation")
        if exec_result.raw["groups_processed"] != report.total_groups_processed:
            problems.append("executive_summary groups_processed did not match the real ValidationReport")
        if exec_result.raw["comments_matched"] != report.total_comments_matched:
            problems.append("executive_summary comments_matched did not match the real ValidationReport")
        if not (3 <= len(exec_result.raw["recommendations"]) <= 5):
            problems.append(f"executive_summary should return 3-5 recommendations, got {len(exec_result.raw['recommendations'])}")

        # --- executive_summary: still produces a coherent result with no report ---
        no_report_exec_result = registry.dispatch("executive_summary", {}, ctx_no_report)
        if "## Executive Summary" not in no_report_exec_result.summary:
            problems.append("executive_summary should still produce a summary when no ValidationReport is available")
        if "## Validation Status" in no_report_exec_result.summary:
            problems.append("executive_summary should omit Validation Status entirely when no ValidationReport is available")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL TOOL CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
