"""
tests/test_ai_service_facade.py
=================================
Tests ``ai.service.AIService`` (Phase 2b: workflow-graph-backed) against
a scripted fake ``LLMProvider`` -- no real Bedrock call. Verifies the
full node sequence executes correctly through the public ``ask()``
method, conversation state persists and resolves a follow-up's
ellipsis across turns, error handling doesn't corrupt the displayed
transcript, and the module boundary rule holds.

Usage:
    python tests/test_ai_service_facade.py
"""
from __future__ import annotations

import ast
import sys
import tempfile
from pathlib import Path
from typing import Iterator, List, Optional

import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.llm.provider import LLMProvider, LLMProviderError, Message, Response, StreamChunk, ToolUseRequest  # noqa: E402
from ai.service import AIService, AIServiceError  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


class _ScriptedProvider(LLMProvider):
    """Returns responses from a fixed script, in order, one per call."""

    def __init__(self, responses: List[Response]) -> None:
        self._responses = list(responses)
        self.call_count = 0
        self.fail_on_call: Optional[int] = None

    def converse(self, messages: List[Message], *, system_prompt: str = "", tools=None) -> Response:
        self.call_count += 1
        if self.fail_on_call == self.call_count:
            raise LLMProviderError("synthetic failure", retryable=True)
        return self._responses.pop(0)

    def converse_stream(self, messages, *, system_prompt: str = "", tools=None) -> Iterator[StreamChunk]:
        raise NotImplementedError

    def embed(self, texts: List[str]) -> List[np.ndarray]:
        raise NotImplementedError


def _revenue_lookup_script(final_text: str, tool_name: str = "revenue_analysis", arguments: Optional[dict] = None) -> List[Response]:
    """The four-call script one analytical question with one tool call
    produces: intent detection, one tool-use request, the loop's
    stopping call, and final synthesis."""
    return [
        Response(text="analytical_question", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        Response(
            text=None,
            tool_use=ToolUseRequest(tool_use_id="t1", name=tool_name, arguments=arguments or {}),
            stop_reason="tool_use", input_tokens=1, output_tokens=1,
        ),
        Response(text="done", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        Response(text=final_text, tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
    ]


def _check_no_streamlit_import(module_path: Path) -> Optional[str]:
    tree = ast.parse(module_path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "streamlit" or alias.name.startswith("streamlit."):
                    return f"{module_path}: found 'import {alias.name}'"
        if isinstance(node, ast.ImportFrom) and node.module and (
            node.module == "streamlit" or node.module.startswith("streamlit.")
        ):
            return f"{module_path}: found 'from {node.module} import ...'"
    return None


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot test AIService.")
            return 1
        ctx = BusinessContext.from_generation_result(result)

        # --- basic ask(): full graph runs, final text and sources_used are correct ---
        provider = _ScriptedProvider(_revenue_lookup_script("HPE's revenue is $1,314,659.", arguments={"client": "HPE"}))
        progress_events = []
        service = AIService(ctx, provider)
        turn = service.ask("s1", "What is HPE's revenue?", progress_cb=lambda n, l, s: progress_events.append((n, s)))
        if turn.text != "HPE's revenue is $1,314,659.":
            problems.append(f"ask() did not return the expected final response: {turn.text!r}")
        if turn.sources_used != ["Revenue Analysis Tool"]:
            problems.append(f"ask() sources_used incorrect: {turn.sources_used}")
        node_names_that_ran = {n for n, status in progress_events if status == "done"}
        if node_names_that_ran != {"intent_detection", "planning", "entity_resolution", "filtering", "analysis", "response"}:
            problems.append(f"Unexpected set of nodes ran: {node_names_that_ran}")

        # --- displayed transcript contains exactly the user message + final assistant text ---
        history = service.history_for("s1")
        if len(history) != 2 or history[0].text != "What is HPE's revenue?" or history[1].text != turn.text:
            problems.append(f"history_for() did not contain the expected 2-message transcript: {history}")

        # --- follow-up: ellipsis resolves via ConversationState, end to end through ask() ---
        provider._responses = _revenue_lookup_script("HPE's margin is $1,314,659 - wait, $228,367.", tool_name="margin_analysis")
        turn2 = service.ask("s1", "Show margin")
        if turn2.sources_used != ["Margin Analysis Tool"]:
            problems.append(f"Follow-up turn sources_used incorrect: {turn2.sources_used}")
        if service._sessions["s1"].state.active_filter.client != "HPE":
            problems.append("Conversation state did not carry the client forward across turns")

        # --- session isolation ---
        provider._responses = [
            Response(text="general_conversation", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
            Response(text="You're welcome!", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        ]
        service.ask("s2", "Thanks!")
        if service._sessions["s2"].state.active_filter.client is not None:
            problems.append("A new session should not inherit another session's conversation state")

        # --- error handling: provider failure rolls back the transcript, raises AIServiceError ---
        provider2 = _ScriptedProvider(_revenue_lookup_script("won't get here"))
        provider2.fail_on_call = 1  # fail on the very first call (intent detection)
        service2 = AIService(ctx, provider2)
        history_before = service2.history_for("s3")
        try:
            service2.ask("s3", "This will fail")
            problems.append("ask() should raise AIServiceError when the provider fails")
        except AIServiceError as exc:
            if not exc.retryable:
                problems.append("AIServiceError should propagate the provider's retryable flag")
        if service2.history_for("s3") != history_before:
            problems.append("A failed ask() call should roll back the just-appended user message")

        # --- reset_session ---
        service.reset_session("s1")
        if service.history_for("s1") != []:
            problems.append("reset_session() did not clear history")
        service.reset_session("unknown-session")  # must not raise

        # --- general_conversation and unsupported_request skip Analysis entirely (no tool calls, cheaper) ---
        provider3 = _ScriptedProvider([
            Response(text="general_conversation", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
            Response(text="Hi there!", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        ])
        service3 = AIService(ctx, provider3)
        turn3 = service3.ask("s4", "Hello!")
        if provider3.call_count != 2:  # intent detection + response synthesis only, no analysis loop
            problems.append(f"A general_conversation turn should make exactly 2 provider calls, made {provider3.call_count}")
        if turn3.sources_used != []:
            problems.append("A general_conversation turn should have no sources_used")

        provider4 = _ScriptedProvider([
            Response(text="unsupported_request", tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1),
        ])
        service4 = AIService(ctx, provider4)
        turn4 = service4.ask("s5", "Make me a chart")
        if provider4.call_count != 1:  # intent detection only -- Response uses a fixed message, no LLM call
            problems.append(f"An unsupported_request turn should make exactly 1 provider call, made {provider4.call_count}")
        if "chart" not in turn4.text.lower() and "not yet available" not in turn4.text.lower() and "can't create" not in turn4.text.lower():
            problems.append(f"unsupported_request response should explain the limitation: {turn4.text!r}")

    # --- module boundary rule ---
    ai_root = PROJECT_ROOT / "ai"
    non_ui_modules = [
        p for p in ai_root.rglob("*.py")
        if "ui" not in p.relative_to(ai_root).parts and p.name != "__init__.py"
    ]
    for module_path in non_ui_modules:
        violation = _check_no_streamlit_import(module_path)
        if violation:
            problems.append(f"Module boundary violation: {violation}")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL AISERVICE FACADE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
