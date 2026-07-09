"""
tests/test_ai_context.py
==========================
Tests ``ai.context.BusinessContext`` construction, its lookup helpers,
its grounding summary, its fingerprint stability, and its error handling
for an incomplete/failed ``GenerationResult``.

Usage:
    python tests/test_ai_context.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import GenerationResult, generate_summary  # noqa: E402
from ai.context import BusinessContext, IncompleteGenerationResultError  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test BusinessContext.")
            return 1

        ctx = BusinessContext.from_generation_result(result)

        # --- basic shape ---
        if ctx.target_year != 2026:
            problems.append(f"target_year = {ctx.target_year}, expected 2026")
        if ctx.prior_years != [2024, 2025]:
            problems.append(f"prior_years = {ctx.prior_years}, expected [2024, 2025]")
        if len(ctx.groups_df) != 32:
            problems.append(f"groups_df has {len(ctx.groups_df)} rows, expected 32")
        if len(ctx.monthly_df) != 32 * 12:
            problems.append(f"monthly_df has {len(ctx.monthly_df)} rows, expected {32 * 12}")

        # --- lookup helpers ---
        group_names = ctx.group_names()
        if "Aldevron" not in group_names:
            problems.append("group_names() did not include 'Aldevron'")
        if len(group_names) != len(set(group_names)):
            problems.append("group_names() returned duplicates")

        poc_names = ctx.poc_names()
        if "Vijay" not in poc_names or "Neeraj" not in poc_names:
            problems.append(f"poc_names() missing an expected POC: {poc_names}")
        if any(name != name for name in poc_names):  # NaN check
            problems.append("poc_names() should never include a NaN/missing value")

        section_keys = ctx.section_keys()
        if set(section_keys) != {"projects_track1", "staffing_secured"}:
            problems.append(f"section_keys() = {section_keys}, expected the two known fixture sections")

        # --- grounding summary: a compact, factual string, not a data dump ---
        summary = ctx.grounding_summary()
        if "2026" not in summary:
            problems.append("grounding_summary() did not mention the target year")
        if str(len(ctx.groups_df)) not in summary:
            problems.append("grounding_summary() did not mention the group count")
        if len(summary) > 2000:
            problems.append(
                f"grounding_summary() is {len(summary)} characters -- expected a compact "
                "aggregate summary, not a per-row data dump"
            )

        # --- fingerprint: stable and deterministic for the same input ---
        ctx_again = BusinessContext.from_generation_result(result)
        if ctx.fingerprint != ctx_again.fingerprint:
            problems.append("Building a BusinessContext twice from the same GenerationResult produced different fingerprints")
        if not ctx.fingerprint:
            problems.append("fingerprint should never be empty for a successful generation")

        # --- immutability: BusinessContext is a frozen dataclass ---
        try:
            ctx.target_year = 9999  # type: ignore[misc]
            problems.append("BusinessContext should be immutable (frozen dataclass) but allowed attribute assignment")
        except AttributeError:
            pass  # expected

    # --- error handling: a failed generation must raise, not silently build an empty context ---
    failed_result = GenerationResult(success=False)
    try:
        BusinessContext.from_generation_result(failed_result)
        problems.append("BusinessContext.from_generation_result should raise for a failed GenerationResult")
    except IncompleteGenerationResultError:
        pass  # expected

    # --- error handling: a successful-looking result missing Phase 2 fields must also raise ---
    incomplete_result = GenerationResult(success=True)  # success=True but no section_results etc.
    try:
        BusinessContext.from_generation_result(incomplete_result)
        problems.append("BusinessContext.from_generation_result should raise when Phase 2 fields are missing")
    except IncompleteGenerationResultError as exc:
        if "section_results" not in str(exc):
            problems.append(f"Error message should name the missing field(s); got: {exc}")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL BUSINESSCONTEXT CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
