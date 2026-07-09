"""
tests/test_ai_dataframe_parity.py
===================================
Mechanically proves ``ai.data.frames``'s "no recalculation" claim: every
column in ``groups_df``/``monthly_df`` is checked against the exact
Phase 1 field it is documented to come from, for every row, against the
real fixture workbook. This is the test the Phase 2 architecture
document names specifically as the proof that the DataFrame layer never
drifts from Phase 1's own numbers (Architecture Plan Revision 2 Section
16, Revision 3 Section 13).

Usage:
    python tests/test_ai_dataframe_parity.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.data.frames import build_groups_dataframe, build_monthly_dataframe  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


def _check_groups_dataframe_parity(section_results, groups_df, problems: list) -> None:
    for section, groups in section_results:
        for group in groups:
            matches = groups_df[(groups_df["group"] == group.group_name) & (groups_df["section"] == section.key)]
            if len(matches) != 1:
                problems.append(
                    f"Expected exactly 1 groups_df row for ({section.key!r}, {group.group_name!r}), found {len(matches)}"
                )
                continue
            row = matches.iloc[0]

            expectations = {
                "poc": group.poc,
                "ds_code": group.ds_code,
                "q1_revenue": group.quarters.get("Q1", 0.0),
                "q2_revenue": group.quarters.get("Q2", 0.0),
                "q3_revenue": group.quarters.get("Q3", 0.0),
                "q4_revenue": group.quarters.get("Q4", 0.0),
                "q1_margin": group.quarter_margins.get("Q1", 0.0),
                "q2_margin": group.quarter_margins.get("Q2", 0.0),
                "q3_margin": group.quarter_margins.get("Q3", 0.0),
                "q4_margin": group.quarter_margins.get("Q4", 0.0),
                "revenue": group.total_revenue,
                "margin": group.total_margin,
                "comment": group.comment,
                "has_renewal_confidence": group.has_renewal_confidence,
            }
            for column, expected in expectations.items():
                actual = row[column]
                # pandas represents a Python None as NaN for object/float
                # columns in some circumstances; treat NaN and None as equal.
                if actual != expected and not (_is_nan(actual) and expected is None):
                    problems.append(
                        f"groups_df[{section.key}/{group.group_name}].{column} = {actual!r}, "
                        f"expected {expected!r} (from GroupSummary)"
                    )


def _check_monthly_dataframe_parity(monthly_section_results, month_roles, monthly_df, problems: list) -> None:
    for section, groups in monthly_section_results:
        for group in groups:
            for month, expected_revenue in group.monthly_revenue.items():
                matches = monthly_df[
                    (monthly_df["group"] == group.group_name)
                    & (monthly_df["section"] == section.key)
                    & (monthly_df["month"] == month)
                ]
                if len(matches) != 1:
                    problems.append(
                        f"Expected exactly 1 monthly_df row for "
                        f"({section.key!r}, {group.group_name!r}, month={month}), found {len(matches)}"
                    )
                    continue
                row = matches.iloc[0]
                expected_margin = group.monthly_margin.get(month, 0.0)
                expected_role = month_roles.get(month, "Actual")

                if row["revenue"] != expected_revenue:
                    problems.append(
                        f"monthly_df[{group.group_name}/month={month}].revenue = {row['revenue']!r}, "
                        f"expected {expected_revenue!r} (from MonthlyGroupSummary.monthly_revenue)"
                    )
                if row["margin"] != expected_margin:
                    problems.append(
                        f"monthly_df[{group.group_name}/month={month}].margin = {row['margin']!r}, "
                        f"expected {expected_margin!r} (from MonthlyGroupSummary.monthly_margin)"
                    )
                if row["role"] != expected_role:
                    problems.append(
                        f"monthly_df[{group.group_name}/month={month}].role = {row['role']!r}, "
                        f"expected {expected_role!r} (from month_roles)"
                    )
                if row["poc"] != group.poc and not (_is_nan(row["poc"]) and group.poc is None):
                    problems.append(
                        f"monthly_df[{group.group_name}/month={month}].poc = {row['poc']!r}, "
                        f"expected {group.poc!r}"
                    )
                if row["comment"] != group.comment and not (_is_nan(row["comment"]) and group.comment is None):
                    problems.append(
                        f"monthly_df[{group.group_name}/month={month}].comment mismatch: "
                        f"got {row['comment']!r}, expected {group.comment!r}"
                    )


def _is_nan(value: object) -> bool:
    try:
        return value != value  # NaN is the only value not equal to itself
    except Exception:  # noqa: BLE001 - defensive, value may not support !=
        return False


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print(f"Generation FAILED ({result.error_title}: {result.error_message}) - cannot test parity.")
            return 1

        groups_df = build_groups_dataframe(result.section_results, result.target_year)
        monthly_df = build_monthly_dataframe(result.monthly_section_results, result.month_roles, result.target_year)

        expected_group_rows = sum(len(groups) for _, groups in result.section_results)
        expected_monthly_rows = sum(
            len(group.monthly_revenue) for _, groups in result.monthly_section_results for group in groups
        )
        if len(groups_df) != expected_group_rows:
            problems.append(f"groups_df has {len(groups_df)} rows, expected {expected_group_rows}")
        if len(monthly_df) != expected_monthly_rows:
            problems.append(f"monthly_df has {len(monthly_df)} rows, expected {expected_monthly_rows}")

        _check_groups_dataframe_parity(result.section_results, groups_df, problems)
        _check_monthly_dataframe_parity(result.monthly_section_results, result.month_roles, monthly_df, problems)

        if not problems:
            print(f"groups_df: {len(groups_df)} rows, all columns match GroupSummary exactly. PASS")
            print(f"monthly_df: {len(monthly_df)} rows, all columns match MonthlyGroupSummary exactly. PASS")

    if problems:
        print("\nFAILURES:")
        for p in problems[:30]:
            print(f"  - {p}")
        if len(problems) > 30:
            print(f"  ... and {len(problems) - 30} more")
        print(f"\nFAIL - {len(problems)} parity violation(s).")
        return 1

    print("\nALL DATAFRAME PARITY CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
