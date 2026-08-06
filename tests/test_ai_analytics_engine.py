"""
tests/test_ai_analytics_engine.py
====================================
Tests ``ai.analytics.engine.AnalyticsEngine``'s kpi/rank/compare/trend/
aggregate operations against real fixture data, cross-checked against
independent pandas computations (not the engine's own logic) to prove
correctness rather than just self-consistency.

Usage:
    python tests/test_ai_analytics_engine.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.analytics.engine import AnalyticsEngine  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.data.filters import Filter  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test the Analytics Engine.")
            return 1
        ctx = BusinessContext.from_generation_result(result)
        engine = AnalyticsEngine(ctx)

        # --- kpi: matches an independent pandas sum ---
        independent_total = float(ctx.groups_df["revenue"].sum())
        engine_total = engine.kpi(ctx.groups_df, "revenue", Filter())
        if engine_total != independent_total:
            problems.append(f"kpi() total ({engine_total}) != independent sum ({independent_total})")

        independent_hpe = float(ctx.groups_df[ctx.groups_df["group"] == "HPE"]["revenue"].sum())
        engine_hpe = engine.kpi(ctx.groups_df, "revenue", Filter(client="HPE"))
        if engine_hpe != independent_hpe:
            problems.append("kpi() with a client filter did not match an independent filtered sum")

        # --- kpi: unknown metric column returns 0.0, not a KeyError ---
        if engine.kpi(ctx.groups_df, "not_a_real_column", Filter()) != 0.0:
            problems.append("kpi() with an unknown metric column should return 0.0")

        # --- rank: descending by default, correct length, matches independent nlargest ---
        independent_top5 = ctx.groups_df.nlargest(5, "revenue")["group"].tolist()
        engine_top5 = engine.rank(ctx.groups_df, "revenue", Filter(), top_n=5)["group"].tolist()
        if engine_top5 != independent_top5:
            problems.append(f"rank() top 5 ({engine_top5}) != independent nlargest ({independent_top5})")

        # --- rank: ascending ---
        independent_bottom3 = ctx.groups_df.nsmallest(3, "margin")["group"].tolist()
        engine_bottom3 = engine.rank(ctx.groups_df, "margin", Filter(), top_n=3, ascending=True)["group"].tolist()
        if engine_bottom3 != independent_bottom3:
            problems.append("rank() with ascending=True did not match independent nsmallest")

        # --- compare: correct delta and percent_change ---
        comparison = engine.compare(
            ctx.monthly_df, "revenue",
            Filter(section="staffing_secured", quarters=["Q2"]),
            Filter(section="staffing_secured", quarters=["Q3"]),
            label_a="Q2", label_b="Q3",
        )
        independent_q2 = float(ctx.monthly_df[(ctx.monthly_df["section"] == "staffing_secured") & (ctx.monthly_df["quarter"] == "Q2")]["revenue"].sum())
        independent_q3 = float(ctx.monthly_df[(ctx.monthly_df["section"] == "staffing_secured") & (ctx.monthly_df["quarter"] == "Q3")]["revenue"].sum())
        if comparison.value_a != independent_q2 or comparison.value_b != independent_q3:
            problems.append("compare() values did not match independent filtered sums")
        expected_delta = independent_q3 - independent_q2
        if abs(comparison.delta - expected_delta) > 1e-9:
            problems.append(f"compare() delta ({comparison.delta}) != expected ({expected_delta})")
        if independent_q2 != 0:
            expected_pct = expected_delta / independent_q2 * 100.0
            if comparison.percent_change is None or abs(comparison.percent_change - expected_pct) > 1e-6:
                problems.append("compare() percent_change did not match the expected calculation")

        # --- compare: percent_change is None when the baseline is zero, not a ZeroDivisionError ---
        zero_baseline_comparison = engine.compare(
            ctx.groups_df, "revenue", Filter(client="NoSuchClientAtAll"), Filter(client="HPE"),
        )
        if zero_baseline_comparison.percent_change is not None:
            problems.append("compare() with a zero baseline should report percent_change=None, not a computed value")

        # --- trend: grouped sums match an independent groupby ---
        independent_trend = ctx.monthly_df.groupby("month")["revenue"].sum().reset_index()
        engine_trend = engine.trend(ctx.monthly_df, "revenue", Filter(), group_by="month")
        if not independent_trend.equals(engine_trend.sort_values("month").reset_index(drop=True)):
            problems.append("trend() did not match an independent groupby('month').sum()")

        # --- aggregate: grouped sums match an independent groupby on a different column ---
        independent_by_poc = ctx.groups_df.groupby("poc")["revenue"].sum().reset_index()
        engine_by_poc = engine.aggregate(ctx.groups_df, "revenue", "poc", Filter())
        if not independent_by_poc.equals(engine_by_poc.sort_values("poc").reset_index(drop=True).sort_values("poc").reset_index(drop=True)):
            # Allow for row-order differences; compare as sets of tuples instead.
            independent_set = set(independent_by_poc.itertuples(index=False, name=None))
            engine_set = set(engine_by_poc.itertuples(index=False, name=None))
            if independent_set != engine_set:
                problems.append("aggregate() did not match an independent groupby('poc').sum()")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL ANALYTICS ENGINE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
