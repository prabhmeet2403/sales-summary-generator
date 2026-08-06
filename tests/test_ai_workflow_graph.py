"""
tests/test_ai_workflow_graph.py
==================================
Tests ``ai.workflow.graph``'s mechanics directly: node ordering,
``should_run()`` skip behavior, progress callback firing, and
``ExecutionTrace`` recording -- independent of any specific node's
business behavior (covered by ``test_ai_tools.py`` and the end-to-end
integration test).

Usage:
    python tests/test_ai_workflow_graph.py
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import ClassVar, List, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from ai.session import ConversationState  # noqa: E402
from ai.tools.base import ToolResult  # noqa: E402
from ai.workflow.graph import ExecutionTrace, WorkflowGraph, WorkflowNode, WorkflowState  # noqa: E402


class _AlwaysRunNode(WorkflowNode):
    name: ClassVar[str] = "always_run"
    display_label: ClassVar[str] = "Always Runs"

    def should_run(self, state: WorkflowState) -> bool:
        return True

    def run(self, state: WorkflowState) -> WorkflowState:
        state.final_response = "ran"
        return state


class _NeverRunNode(WorkflowNode):
    name: ClassVar[str] = "never_run"
    display_label: ClassVar[str] = "Never Runs"

    def should_run(self, state: WorkflowState) -> bool:
        return False

    def run(self, state: WorkflowState) -> WorkflowState:
        raise AssertionError("This node's run() should never be called")


class _OrderRecordingNode(WorkflowNode):
    def __init__(self, name: str, order_log: List[str]) -> None:
        self.name = name
        self.display_label = name
        self._order_log = order_log

    def should_run(self, state: WorkflowState) -> bool:
        return True

    def run(self, state: WorkflowState) -> WorkflowState:
        self._order_log.append(self.name)
        return state


def main() -> int:
    problems: list = []

    # --- should_run() gating: a False node's run() is never invoked ---
    graph = WorkflowGraph([_AlwaysRunNode(), _NeverRunNode()])
    state = WorkflowState(user_message="hi", conversation_state=ConversationState())
    final_state = graph.execute(state)
    if final_state.final_response != "ran":
        problems.append("The always-run node's effect was not applied")

    # --- node execution order matches the list order given to WorkflowGraph ---
    order_log: List[str] = []
    ordered_graph = WorkflowGraph([
        _OrderRecordingNode("first", order_log),
        _OrderRecordingNode("second", order_log),
        _OrderRecordingNode("third", order_log),
    ])
    ordered_graph.execute(WorkflowState(user_message="hi", conversation_state=ConversationState()))
    if order_log != ["first", "second", "third"]:
        problems.append(f"Node execution order was {order_log}, expected ['first', 'second', 'third']")

    # --- progress_cb fires running/done for each node that actually runs, and not for skipped nodes ---
    progress_events: List[Tuple[str, str, str]] = []
    graph_with_skip = WorkflowGraph([_AlwaysRunNode(), _NeverRunNode()])
    graph_with_skip.execute(
        WorkflowState(user_message="hi", conversation_state=ConversationState()),
        progress_cb=lambda name, label, status: progress_events.append((name, label, status)),
    )
    event_node_names = {name for name, _, _ in progress_events}
    if "never_run" in event_node_names:
        problems.append("progress_cb should never fire for a node whose should_run() is False")
    if event_node_names != {"always_run"}:
        problems.append(f"Expected progress events only for 'always_run', got node names: {event_node_names}")
    statuses_for_always_run = [status for name, _, status in progress_events if name == "always_run"]
    if statuses_for_always_run != ["running", "done"]:
        problems.append(f"Expected ['running', 'done'] for always_run, got {statuses_for_always_run}")

    # --- ExecutionTrace: node_steps recorded with timing, tool_steps recorded with display names ---
    trace_state = WorkflowState(user_message="hi", conversation_state=ConversationState())
    trace_graph = WorkflowGraph([_AlwaysRunNode()])
    final_trace_state = trace_graph.execute(trace_state)
    if len(final_trace_state.trace.node_steps) != 1 or final_trace_state.trace.node_steps[0].node_name != "always_run":
        problems.append("ExecutionTrace did not record the executed node")
    if final_trace_state.trace.node_steps[0].elapsed_seconds < 0:
        problems.append("Recorded node elapsed_seconds should be non-negative")

    trace = ExecutionTrace()
    trace.record_tool("revenue_analysis", "Revenue Analysis Tool", {"client": "HPE"}, ToolResult(summary="ok"))
    trace.record_tool("revenue_analysis", "Revenue Analysis Tool", {"client": "HPE"}, ToolResult(summary="ok again"))
    trace.record_tool("margin_analysis", "Margin Analysis Tool", {}, ToolResult(summary="ok"))
    if trace.sources_used() != ["Revenue Analysis Tool", "Margin Analysis Tool"]:
        problems.append(
            f"sources_used() should de-duplicate while preserving first-seen order; got {trace.sources_used()}"
        )

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL WORKFLOW GRAPH MECHANICS CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
