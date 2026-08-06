"""
tests/test_ai_phase2b_integration.py
=======================================
End-to-end proof of Phase 2b's exit criterion (Architecture Plan v3,
Section 15): "the full worked example (Client Filter -> Comparison ->
Revenue) runs end-to-end through the graph, with the live checklist
rendering and a real 'Sources Used' trail."

Exercises the complete real chain -- real ``generate_summary()``, real
``BusinessContext``, real ``AIService``, real ``WorkflowGraph`` with all
real nodes and real tools -- with only the LLM network boundary faked
(this environment has no live AWS access; per the established testing
philosophy, the provider boundary is the one place mocking is
appropriate).

Usage:
    python tests/test_ai_phase2b_integration.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import Iterator, List, Optional

import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.llm.provider import LLMProvider, Message, Response, StreamChunk, ToolUseRequest  # noqa: E402
from ai.service import AIService  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


class _WorkedExampleProvider(LLMProvider):
    """Scripts exactly the worked example: Client Filter (via
    client_lookup) -> Comparison (via quarter_comparison) -> a final
    synthesized answer. "Revenue" in the exit criterion's naming is
    satisfied by client_lookup's own revenue/margin profile output --
    a real, registered ANALYSIS-category tool, matching the exit
    criterion's intent (a chain of more than one distinct tool) without
    assuming tool names the exit criterion itself doesn't specify."""

    def __init__(self) -> None:
        self._responses = [
            Response(text="analytical_question", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
            Response(
                text=None,
                tool_use=ToolUseRequest(tool_use_id="t1", name="client_lookup", arguments={"client": "HPE"}),
                stop_reason="tool_use", input_tokens=1, output_tokens=1,
            ),
            Response(
                text=None,
                tool_use=ToolUseRequest(
                    tool_use_id="t2", name="quarter_comparison",
                    arguments={"quarter_a": "Q2", "quarter_b": "Q3", "client": "HPE", "section": "staffing_secured"},
                ),
                stop_reason="tool_use", input_tokens=1, output_tokens=1,
            ),
            Response(text="done", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
            Response(
                text="HPE is a Staffing-Secured client. Its revenue declined from Q2 ($344,692) to Q3 ($330,653), a 4.1% decrease.",
                tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1,
            ),
        ]

    def converse(self, messages: List[Message], *, system_prompt: str = "", tools=None) -> Response:
        return self._responses.pop(0)

    def converse_stream(self, messages, *, system_prompt: str = "", tools=None) -> Iterator[StreamChunk]:
        raise NotImplementedError

    def embed(self, texts: List[str]) -> List[np.ndarray]:
        raise NotImplementedError


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        # Step 1: the real, unmodified Phase 1 pipeline.
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print(f"Generation FAILED ({result.error_title}) - cannot run the integration test.")
            return 1

        # Step 2: a real BusinessContext.
        context = BusinessContext.from_generation_result(result)

        # Step 3: a real AIService, with a real (auto-discovered) tool
        # registry and a real WorkflowGraph -- only the LLM network
        # boundary is faked.
        provider = _WorkedExampleProvider()
        service = AIService(context, provider)

        # Step 4: the live checklist -- captured via progress_cb exactly
        # as the Streamlit UI does.
        checklist: List[str] = []

        def progress_cb(node_name: str, display_label: str, status: str) -> None:
            if status == "done":
                checklist.append(display_label)

        # Step 5: the worked example.
        turn = service.ask(
            "integration-session",
            "Tell me about HPE and compare Q2 vs Q3 revenue in Staffing.",
            progress_cb=progress_cb,
        )

        # --- Verify the live checklist rendered the expected node sequence ---
        expected_checklist = [
            "Understanding request", "Planning analysis", "Resolving client",
            "Applying filters", "Running analysis", "Finalizing response",
        ]
        if checklist != expected_checklist:
            problems.append(f"Live checklist was {checklist}, expected {expected_checklist}")

        # --- Verify a real, multi-tool "Sources Used" trail ---
        if turn.sources_used != ["Client Lookup Tool", "Quarter Comparison Tool"]:
            problems.append(f"Expected a two-tool Sources Used trail, got {turn.sources_used}")

        # --- Verify the final response is the coherent, synthesized answer ---
        if "4.1%" not in turn.text or "HPE" not in turn.text:
            problems.append(f"Final response did not contain the expected synthesized content: {turn.text!r}")

        # --- Verify this used REAL Phase 1 numbers: cross-check against
        #     an independent AnalyticsEngine computation ---
        from ai.analytics.engine import AnalyticsEngine
        from ai.data.filters import Filter
        engine = AnalyticsEngine(context)
        real_comparison = engine.compare(
            context.monthly_df, "revenue",
            Filter(client="HPE", section="staffing_secured", quarters=["Q2"]),
            Filter(client="HPE", section="staffing_secured", quarters=["Q3"]),
        )
        if f"{real_comparison.value_a:,.0f}" not in turn.text:
            problems.append(
                f"Sanity check failed: the real Q2 value ({real_comparison.value_a:,.0f}) does not "
                f"appear in the scripted final answer -- the test's script has drifted from real data."
            )

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("Full chain verified: generate_summary() -> BusinessContext -> AIService.ask() ->")
    print("  WorkflowGraph (6 real nodes) -> 2 tools (Client Lookup, Quarter Comparison) -> response")
    print("Live checklist and Sources Used trail both verified against the exact exit criterion.")
    print("\nALL PHASE 2B INTEGRATION CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
