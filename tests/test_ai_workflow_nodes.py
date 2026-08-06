"""
tests/test_ai_workflow_nodes.py
==================================
Tests individual workflow node behaviors not already covered by the
graph-mechanics test or the end-to-end integration test: intent
parsing/fallback, entity resolution's longest-match rule, Planning's
defensive category check, and the future-phase stub nodes' documented
"should never run, raises loudly if it does" contract.

Usage:
    python tests/test_ai_workflow_nodes.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import Iterator, List

import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from ai.context import BusinessContext  # noqa: E402
from ai.llm.provider import LLMProvider, Response, StreamChunk  # noqa: E402
from ai.session import ConversationState  # noqa: E402
from ai.tools.registry import ToolRegistry, discover_tools  # noqa: E402
from ai.workflow.graph import Intent, ResolvedEntities, WorkflowState  # noqa: E402
from ai.workflow.nodes.entity_resolution import EntityResolutionNode  # noqa: E402
from ai.workflow.nodes.filtering import FilteringNode  # noqa: E402
from ai.workflow.nodes.future_stubs import ExportNode, ReportingNode, VisualizationNode  # noqa: E402
from ai.workflow.nodes.intent_detection import IntentDetectionNode  # noqa: E402
from ai.workflow.nodes.planning import PlanningNode  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"


class _CannedProvider(LLMProvider):
    def __init__(self, text: str) -> None:
        self._text = text

    def converse(self, messages, *, system_prompt: str = "", tools=None) -> Response:
        return Response(text=self._text, tool_use=None, stop_reason="end_turn", input_tokens=1, output_tokens=1)

    def converse_stream(self, messages, *, system_prompt: str = "", tools=None) -> Iterator[StreamChunk]:
        raise NotImplementedError

    def embed(self, texts: List[str]) -> List[np.ndarray]:
        raise NotImplementedError


def _state(message: str = "hi") -> WorkflowState:
    return WorkflowState(user_message=message, conversation_state=ConversationState())


def main() -> int:
    problems: list = []

    # --- IntentDetectionNode: correctly parses each valid response ---
    for expected_intent, canned_text in [
        (Intent.ANALYTICAL_QUESTION, "analytical_question"),
        (Intent.GENERAL_CONVERSATION, "general_conversation"),
        (Intent.UNSUPPORTED_REQUEST, "unsupported_request"),
        (Intent.ANALYTICAL_QUESTION, "  Analytical_Question  "),  # tolerant of whitespace/case
    ]:
        node = IntentDetectionNode(_CannedProvider(canned_text))
        result_state = node.run(_state())
        if result_state.intent != expected_intent:
            problems.append(f"IntentDetectionNode with response {canned_text!r} produced {result_state.intent}, expected {expected_intent}")

    # --- IntentDetectionNode: unparseable response defaults to ANALYTICAL_QUESTION, doesn't crash ---
    node = IntentDetectionNode(_CannedProvider("completely unrelated gibberish"))
    result_state = node.run(_state())
    if result_state.intent != Intent.ANALYTICAL_QUESTION:
        problems.append("IntentDetectionNode should default to ANALYTICAL_QUESTION on an unparseable response")

    # --- IntentDetectionNode.should_run() is always True ---
    if not IntentDetectionNode(_CannedProvider("x")).should_run(_state()):
        problems.append("IntentDetectionNode.should_run() should always return True")

    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        ctx = BusinessContext.from_generation_result(result)

        # --- EntityResolutionNode: finds a known client mentioned in the message ---
        entity_node = EntityResolutionNode(ctx)
        entity_state = _state("What is HPE's revenue this year?")
        entity_state.intent = Intent.ANALYTICAL_QUESTION
        resolved_state = entity_node.run(entity_state)
        if resolved_state.resolved_entities.client != "HPE":
            problems.append(f"EntityResolutionNode failed to find 'HPE' in the message: {resolved_state.resolved_entities}")

        # --- EntityResolutionNode: no known entity mentioned -> None, not a crash ---
        no_entity_state = _state("What is the weather like today?")
        no_entity_state.intent = Intent.ANALYTICAL_QUESTION
        no_entity_result = entity_node.run(no_entity_state)
        if no_entity_result.resolved_entities.client is not None:
            problems.append("EntityResolutionNode should find no client in an unrelated message")

        # --- EntityResolutionNode: longest match preferred (avoid a short false-positive match) ---
        hitachi_state = _state("How is Hitachi Asia doing?")
        hitachi_state.intent = Intent.ANALYTICAL_QUESTION
        hitachi_result = entity_node.run(hitachi_state)
        if hitachi_result.resolved_entities.client != "Hitachi Asia":
            problems.append(f"EntityResolutionNode should resolve the specific client mentioned; got {hitachi_result.resolved_entities.client!r}")

        # --- EntityResolutionNode.should_run() gates on intent ---
        general_state = _state("thanks!")
        general_state.intent = Intent.GENERAL_CONVERSATION
        if entity_node.should_run(general_state):
            problems.append("EntityResolutionNode.should_run() should be False for GENERAL_CONVERSATION intent")

        # --- FilteringNode: merges resolved entities with carried-forward state ---
        filtering_node = FilteringNode()
        filter_state = _state("Show margin")
        filter_state.intent = Intent.ANALYTICAL_QUESTION
        filter_state.conversation_state.active_filter.client = "Aldevron"  # simulate a prior turn
        filter_state.resolved_entities = ResolvedEntities(client=None, poc=None)  # nothing new resolved
        filtered_result = filtering_node.run(filter_state)
        if filtered_result.active_filter.client != "Aldevron":
            problems.append("FilteringNode should fall back to the conversation's carried-forward client when none is newly resolved")

        # --- PlanningNode: downgrades intent when no ANALYSIS tools are registered ---
        empty_registry = ToolRegistry()  # deliberately empty
        planning_node = PlanningNode(empty_registry)
        planning_state = _state("What is total revenue?")
        planning_state.intent = Intent.ANALYTICAL_QUESTION
        planning_result = planning_node.run(planning_state)
        if planning_result.intent != Intent.UNSUPPORTED_REQUEST:
            problems.append("PlanningNode should downgrade to UNSUPPORTED_REQUEST when no ANALYSIS tools are registered")

        # --- PlanningNode: leaves intent alone when tools ARE registered ---
        real_registry = discover_tools()
        planning_node_real = PlanningNode(real_registry)
        planning_state2 = _state("What is total revenue?")
        planning_state2.intent = Intent.ANALYTICAL_QUESTION
        planning_result2 = planning_node_real.run(planning_state2)
        if planning_result2.intent != Intent.ANALYTICAL_QUESTION:
            problems.append("PlanningNode should leave intent as ANALYTICAL_QUESTION when ANALYSIS tools exist")

    # --- Future-phase stub nodes: should_run() always False, run() raises loudly if ever called ---
    for stub_node in (VisualizationNode(), ReportingNode(), ExportNode()):
        if stub_node.should_run(_state()):
            problems.append(f"{type(stub_node).__name__}.should_run() should always return False in Phase 2b")
        try:
            stub_node.run(_state())
            problems.append(f"{type(stub_node).__name__}.run() should raise NotImplementedError if ever called")
        except NotImplementedError:
            pass

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL WORKFLOW NODE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
