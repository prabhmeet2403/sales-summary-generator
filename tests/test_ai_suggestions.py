"""
tests/test_ai_suggestions.py
===============================
Tests ``ai.ui.suggestions``: the quick-action groups, the follow-up
mapping keyed by tool display name, and the lightweight
BusinessContext-driven reordering -- against real fixture-generated
data, not a mock.

Usage:
    python tests/test_ai_suggestions.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.ui.suggestions import (  # noqa: E402
    QUICK_ACTION_GROUPS,
    contextual_quick_action_groups,
    follow_up_suggestions,
)

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"

_EXPECTED_GROUPS = {"Executive Summary", "Validation", "Business Analysis", "Search"}
_EXPECTED_GROUP_SIZES = {"Executive Summary": 5, "Validation": 4, "Business Analysis": 4, "Search": 4}
_EXPECTED_FOLLOW_UP_TOOLS = {
    "Revenue Analysis Tool", "Margin Analysis Tool", "Validation Summary Tool",
    "Quarter Comparison Tool", "Client Lookup Tool", "POC Lookup Tool", "Executive Summary Tool",
}


def main() -> int:
    problems: list = []

    # --- static content shape ---
    if set(QUICK_ACTION_GROUPS.keys()) != _EXPECTED_GROUPS:
        problems.append(f"Unexpected quick-action group set: {set(QUICK_ACTION_GROUPS.keys())}")
    for group_name, chips in QUICK_ACTION_GROUPS.items():
        expected_size = _EXPECTED_GROUP_SIZES.get(group_name)
        if len(chips) != expected_size:
            problems.append(f"Group '{group_name}' should have exactly {expected_size} suggestions, has {len(chips)}")
        for label, prompt in chips:
            if not label or not prompt:
                problems.append(f"Empty label/prompt in group '{group_name}': {(label, prompt)}")

    # --- follow-up mapping: every tool in the registry has an entry, every
    #     entry has at least one chip, and an unknown/absent source falls
    #     back to the documented default rather than raising ---
    for tool_name in _EXPECTED_FOLLOW_UP_TOOLS:
        chips = follow_up_suggestions([tool_name])
        if not chips:
            problems.append(f"No follow-up suggestions registered for '{tool_name}'")

    if not follow_up_suggestions(None):
        problems.append("follow_up_suggestions(None) should return the default set, not empty")
    if not follow_up_suggestions(["Some Unregistered Tool"]):
        problems.append("An unrecognized source should fall back to the default set, not empty")

    # --- context-aware reordering, against a real generation ---
    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test context-aware suggestions.")
            return 1

        ctx = BusinessContext.from_generation_result(result, elapsed_seconds=3.0)

        # This fixture has 0 report-level warnings and 1 missing comment
        # (see tests/compare_with_manual.py's own printed validation
        # report) -- assert against the real report rather than a
        # hardcoded assumption, so this test tracks the fixture itself.
        report = result.report
        groups = contextual_quick_action_groups(ctx)

        if set(groups.keys()) != _EXPECTED_GROUPS:
            problems.append(f"contextual_quick_action_groups changed the group set: {set(groups.keys())}")
        for group_name, chips in groups.items():
            expected_size = _EXPECTED_GROUP_SIZES.get(group_name)
            if len(chips) != expected_size or {c[0] for c in chips} != {c[0] for c in QUICK_ACTION_GROUPS[group_name]}:
                problems.append(f"contextual_quick_action_groups altered the item set of '{group_name}'")

        validation_labels = [label for label, _ in groups["Validation"]]
        if report.total_missing_comments > 0 and validation_labels[0] != "Review Comments":
            problems.append(
                f"With {report.total_missing_comments} missing comment(s), 'Review Comments' "
                f"should be prioritized first in Validation: {validation_labels}"
            )
        if len(report.warnings) == 0 and validation_labels.index("Review Warnings") == 0:
            problems.append("With 0 warnings, 'Review Warnings' should not be the top Validation suggestion")

        # Multiple quarters exist in monthly_df -> Compare Quarters should lead
        # Business Analysis; multiple clients exist -> Show Top Clients should
        # lead Executive Summary.
        business_labels = [label for label, _ in groups["Business Analysis"]]
        executive_labels = [label for label, _ in groups["Executive Summary"]]

        multiple_quarters = not ctx.monthly_df.empty and ctx.monthly_df["quarter"].nunique() > 1
        multiple_clients = len(ctx.group_names()) > 1

        if multiple_quarters and business_labels[0] != "Compare Quarters":
            problems.append(f"With multiple quarters, 'Compare Quarters' should lead Business Analysis: {business_labels}")
        if multiple_clients and executive_labels[0] != "Show Top Clients":
            problems.append(f"With multiple clients, 'Show Top Clients' should lead Executive Summary: {executive_labels}")

        # Search quick actions should be filled in with a REAL client/POC name
        # from this generation's own data, not a hardcoded example -- so the
        # same quick action reliably invokes its tool for any dataset, not
        # just this fixture.
        search_prompts = {label: prompt for label, prompt in groups["Search"]}
        real_client = ctx.group_names()[0]
        real_poc = ctx.poc_names()[0]
        if real_client not in search_prompts["Find Client"]:
            problems.append(f"'Find Client' prompt should reference a real client from this data: {search_prompts['Find Client']!r}")
        if real_poc not in search_prompts["Find POC"]:
            problems.append(f"'Find POC' prompt should reference a real POC from this data: {search_prompts['Find POC']!r}")
        if real_client not in search_prompts["Search Revenue"]:
            problems.append(f"'Search Revenue' prompt should reference a real client from this data: {search_prompts['Search Revenue']!r}")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL SUGGESTION CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
