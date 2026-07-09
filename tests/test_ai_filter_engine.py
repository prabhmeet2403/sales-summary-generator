"""
tests/test_ai_filter_engine.py
================================
Tests ``ai.data.filters.apply_filter`` against real fixture data --
every field individually, several in combination, and the documented
"skip clause when column absent" behavior for filter fields that only
apply to one of the two DataFrame shapes.

Usage:
    python tests/test_ai_filter_engine.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.data.frames import build_groups_dataframe, build_monthly_dataframe  # noqa: E402
from ai.data.filters import Filter, apply_filter  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test filters.")
            return 1

        groups_df = build_groups_dataframe(result.section_results, result.target_year)
        monthly_df = build_monthly_dataframe(result.monthly_section_results, result.month_roles, result.target_year)

        # --- Empty filter matches everything, returns the same object ---
        empty = Filter()
        if not empty.is_empty():
            problems.append("Filter() should report is_empty() == True")
        filtered = apply_filter(groups_df, empty)
        if filtered is not groups_df:
            problems.append("apply_filter with an empty Filter should return the same DataFrame object, not a copy")
        if len(filtered) != len(groups_df):
            problems.append("apply_filter with an empty Filter should match every row")

        # --- client filter (case-insensitive, exact match) ---
        by_client = apply_filter(groups_df, Filter(client="aldevron"))
        if not (len(by_client) >= 1 and (by_client["group"] == "Aldevron").all()):
            problems.append(f"Filter(client='aldevron') did not correctly match 'Aldevron' rows: {by_client['group'].tolist()}")

        no_client_match = apply_filter(groups_df, Filter(client="Definitely Not A Real Client"))
        if len(no_client_match) != 0:
            problems.append("Filter with an unknown client should match zero rows")

        # --- poc filter ---
        by_poc = apply_filter(groups_df, Filter(poc="Vijay"))
        if len(by_poc) == 0 or not (by_poc["poc"] == "Vijay").all():
            problems.append("Filter(poc='Vijay') did not correctly filter by POC")

        # --- section filter ---
        by_section = apply_filter(groups_df, Filter(section="staffing_secured"))
        if len(by_section) != 4:  # known fixture fact: Staffing-Secured has 4 groups
            problems.append(f"Filter(section='staffing_secured') matched {len(by_section)} rows, expected 4")

        # --- revenue range filter ---
        by_revenue = apply_filter(groups_df, Filter(min_revenue=100_000))
        if not (by_revenue["revenue"] >= 100_000).all():
            problems.append("Filter(min_revenue=100000) let a lower-revenue row through")
        by_revenue_range = apply_filter(groups_df, Filter(min_revenue=10_000, max_revenue=100_000))
        if not ((by_revenue_range["revenue"] >= 10_000) & (by_revenue_range["revenue"] <= 100_000)).all():
            problems.append("Filter(min_revenue, max_revenue) did not correctly bound the range")

        # --- combined filter (AND semantics across fields) ---
        combined = apply_filter(groups_df, Filter(poc="Neeraj", section="projects_track1", min_revenue=1))
        if not (
            (combined["poc"] == "Neeraj").all()
            and (combined["section"] == "projects_track1").all()
            and (combined["revenue"] >= 1).all()
        ):
            problems.append("Combined Filter did not apply all clauses with AND semantics")

        # --- quarter/month/role filters: no-op (with a debug log, not an error)
        #     on groups_df, which has no per-row quarter/month/role ---
        groups_with_quarter_filter = apply_filter(groups_df, Filter(quarters=["Q2"]))
        if len(groups_with_quarter_filter) != len(groups_df):
            problems.append(
                "Filter(quarters=[...]) should be a no-op on groups_df (no 'quarter' column there), "
                f"but row count changed from {len(groups_df)} to {len(groups_with_quarter_filter)}"
            )

        # --- quarter/month/role filters: DO apply on monthly_df ---
        by_quarter = apply_filter(monthly_df, Filter(quarters=["Q2"]))
        if not (by_quarter["quarter"] == "Q2").all() or len(by_quarter) == 0:
            problems.append("Filter(quarters=['Q2']) did not correctly filter monthly_df by quarter")

        by_month = apply_filter(monthly_df, Filter(months=[1, 2]))
        if not by_month["month"].isin([1, 2]).all() or len(by_month) == 0:
            problems.append("Filter(months=[1, 2]) did not correctly filter monthly_df by month")

        by_role = apply_filter(monthly_df, Filter(role="Forecast"))
        if not (by_role["role"] == "Forecast").all() or len(by_role) == 0:
            problems.append("Filter(role='Forecast') did not correctly filter monthly_df by role")
        by_role_actual = apply_filter(monthly_df, Filter(role="Actual"))
        if len(by_role) + len(by_role_actual) != len(monthly_df):
            problems.append("Actual + Forecast row counts should sum to the total monthly_df row count")

        # --- confidence filter (monthly_df only; no-op on groups_df) ---
        by_confidence = apply_filter(monthly_df, Filter(min_confidence_pct=50))
        if len(by_confidence) > 0 and not (by_confidence["confidence_pct"] >= 50).all():
            problems.append("Filter(min_confidence_pct=50) let a lower-confidence row through")

        # --- years filter applies to both DataFrame shapes ---
        by_year_groups = apply_filter(groups_df, Filter(years=[2026]))
        if len(by_year_groups) != len(groups_df):
            problems.append("Filter(years=[2026]) should match every groups_df row (single-year fixture)")
        by_year_monthly = apply_filter(monthly_df, Filter(years=[2026]))
        if len(by_year_monthly) != len(monthly_df):
            problems.append("Filter(years=[2026]) should match every monthly_df row (single-year fixture)")
        by_wrong_year = apply_filter(groups_df, Filter(years=[1999]))
        if len(by_wrong_year) != 0:
            problems.append("Filter(years=[1999]) should match zero rows")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL FILTER ENGINE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
