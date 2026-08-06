"""
tests/test_ai_tool_routing.py
================================
Verifies that EVERY registered tool is discoverable, actually gets
dispatched when the Analysis node's tool-use loop calls it by name,
and that its result correctly reaches the final answer and the
"Sources Used" trail -- for all eight tools this project ships:

    Revenue Analysis Tool, Margin Analysis Tool, Quarter Comparison
    Tool, Client Lookup Tool, POC Lookup Tool, Validation Summary
    Tool, Executive Summary Tool.

This does not (and cannot, without a live LLM) test whether a real
model's classification/tool-selection judgment picks the "right" tool
for a given natural-language message -- that judgment happens inside
Bedrock itself. What it does verify, mechanically and for real, is
everything downstream of that judgment: that once a tool is selected
by name, the registry actually has it, dispatch actually runs it
against real fixture data, and the result actually flows through to
``AssistantTurn.sources_used`` and the synthesized answer -- the
precise chain "verify every tool is discoverable, routable, and
actually invoked" depends on, for any tool the model does select.

Usage:
    python tests/test_ai_tool_routing.py
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
from ai.tools.registry import discover_tools  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


class _ScriptedProvider(LLMProvider):
    """Returns responses from a fixed script, in order -- no network call."""

    def __init__(self, responses: List[Response]) -> None:
        self._responses = list(responses)

    def converse(self, messages, *, system_prompt: str = "", tools=None) -> Response:
        return self._responses.pop(0)

    def converse_stream(self, messages, *, system_prompt: str = "", tools=None) -> Iterator[StreamChunk]:
        raise NotImplementedError

    def embed(self, texts: List[str]) -> List[np.ndarray]:
        raise NotImplementedError


def _single_tool_script(tool_name: str, arguments: dict, final_text: str) -> List[Response]:
    """The four-call script one analytical question answered by exactly
    one tool call produces: intent detection, the tool-use request, the
    loop's stopping call, and final synthesis -- mirrors
    ``test_ai_service_facade.py``'s ``_revenue_lookup_script`` pattern,
    generalized to any registered tool."""
    return [
        Response(text="analytical_question", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        Response(
            text=None,
            tool_use=ToolUseRequest(tool_use_id="t1", name=tool_name, arguments=arguments),
            stop_reason="tool_use", input_tokens=1, output_tokens=1,
        ),
        Response(text="done", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        Response(text=final_text, tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
    ]


# (tool_name, arguments, expected display_name) for every registered tool.
_TOOL_CASES = [
    ("revenue_analysis", {"client": "HPE", "section": "staffing_secured"}, "Revenue Analysis Tool"),
    ("margin_analysis", {"top_n": 3}, "Margin Analysis Tool"),
    ("quarter_comparison", {"quarter_a": "Q2", "quarter_b": "Q3", "client": "HPE", "section": "staffing_secured"}, "Quarter Comparison Tool"),
    ("client_lookup", {"client": "Aldevron"}, "Client Lookup Tool"),
    ("poc_lookup", {"poc": "Vijay"}, "POC Lookup Tool"),
    ("validation_summary", {}, "Validation Summary Tool"),
    ("executive_summary", {}, "Executive Summary Tool"),
]


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test tool routing.")
            return 1

        ctx = BusinessContext.from_generation_result(result, elapsed_seconds=4.2)
        registry = discover_tools()

        discovered_names = {t.name for t in registry.all_tools()}
        expected_names = {name for name, _, _ in _TOOL_CASES}
        if discovered_names != expected_names:
            problems.append(
                f"Registered tool set does not match the 8 expected tools.\n"
                f"  Discovered: {sorted(discovered_names)}\n"
                f"  Expected:   {sorted(expected_names)}"
            )

        for tool_name, arguments, expected_display_name in _TOOL_CASES:
            # --- discoverable: the registry actually has this tool ---
            try:
                tool = registry.get(tool_name)
            except Exception as exc:  # noqa: BLE001
                problems.append(f"[{tool_name}] not discoverable via the registry: {exc}")
                continue
            if tool.display_name != expected_display_name:
                problems.append(f"[{tool_name}] display_name is {tool.display_name!r}, expected {expected_display_name!r}")

            # --- routable + actually invoked: dispatch runs against real data ---
            try:
                direct_result = registry.dispatch(tool_name, arguments, ctx)
            except Exception as exc:  # noqa: BLE001
                problems.append(f"[{tool_name}] dispatch raised an unexpected exception: {exc}")
                continue
            if not direct_result.summary:
                problems.append(f"[{tool_name}] dispatch produced an empty summary")

            # --- full pipeline: AIService.ask() actually calls this tool
            #     when the model's tool-use response names it, and the
            #     final answer + Sources Used correctly reflect it ---
            final_text = f"Synthesized answer for {tool_name}."
            provider = _ScriptedProvider(_single_tool_script(tool_name, arguments, final_text))
            service = AIService(ctx, provider, registry=registry)
            turn = service.ask(f"session-{tool_name}", f"A question routed to {tool_name}")

            if turn.text != final_text:
                problems.append(
                    f"[{tool_name}] final answer was not generated from this tool's synthesis step: {turn.text!r}"
                )
            if turn.sources_used != [expected_display_name]:
                problems.append(
                    f"[{tool_name}] Sources Used incorrect: got {turn.sources_used}, expected [{expected_display_name!r}]"
                )

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print(f"ALL {len(_TOOL_CASES)} TOOLS VERIFIED: discoverable, routable, invoked, and correctly sourced.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
