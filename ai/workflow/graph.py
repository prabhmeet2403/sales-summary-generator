"""
ai.workflow.graph
===================
The workflow graph: an ordered sequence of independently extensible
``WorkflowNode``s that together turn one user message into one
coherent response.

Each node does exactly one thing (Single Responsibility) and depends
only on ``WorkflowState`` in and out -- adding a future node means
writing one new class and inserting it into
``ai.workflow.build_default_workflow_graph``'s node list; no existing
node changes. ``WorkflowNode.should_run()`` is what keeps a simple
question fast: a factual question skips Visualization/Reporting/Export
entirely rather than running through all nine nodes unconditionally.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 1.
"""

from __future__ import annotations

import logging
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Callable, ClassVar, Dict, List, Optional

from ai.data.filters import Filter
from ai.session import ConversationState
from ai.tools.base import ToolResult

logger = logging.getLogger(__name__)

#: Signature for the live-progress callback ``WorkflowGraph.execute``
#: invokes as each node starts and finishes. Arguments are
#: ``(node_name, display_label, status)`` where ``status`` is
#: ``"running"`` or ``"done"``.
ProgressCallback = Callable[[str, str, str], None]


class Intent(str, Enum):
    """The request classifications the Intent Detection node produces.

    Deliberately scoped to what Phase 2b's tool registry can actually
    act on: only :data:`ANALYTICAL_QUESTION` triggers the
    Analysis node. As chart/report/export/search tools are added in
    later phases, this enum is expected to grow (an additive change,
    not a redesign of the classification mechanism itself).
    """

    ANALYTICAL_QUESTION = "analytical_question"
    GENERAL_CONVERSATION = "general_conversation"
    UNSUPPORTED_REQUEST = "unsupported_request"


@dataclass
class ResolvedEntities:
    """Named entities the Entity Resolution node found in the user's message.

    Attributes:
        client: A client/group name recognized in the message, if any.
        poc: A POC name recognized in the message, if any.
    """

    client: Optional[str] = None
    poc: Optional[str] = None


@dataclass
class ToolStep:
    """One tool invocation recorded in an :class:`ExecutionTrace`."""

    tool_name: str
    display_name: str
    arguments: dict
    result: ToolResult


@dataclass
class NodeStep:
    """One workflow node's execution, recorded in an :class:`ExecutionTrace`."""

    node_name: str
    elapsed_seconds: float


@dataclass
class ExecutionTrace:
    """Everything that happened while answering one message: which
    nodes ran, how long each took, and which tools were invoked with
    what arguments and results.

    This is what makes Explainability possible without any separate
    bookkeeping: "Sources Used" is a direct readout of
    :meth:`sources_used`, not a reconstruction after the fact.
    """

    node_steps: List[NodeStep] = field(default_factory=list)
    tool_steps: List[ToolStep] = field(default_factory=list)

    def record_node(self, node_name: str, elapsed_seconds: float) -> None:
        self.node_steps.append(NodeStep(node_name=node_name, elapsed_seconds=elapsed_seconds))

    def record_tool(self, tool_name: str, display_name: str, arguments: dict, result: ToolResult) -> None:
        self.tool_steps.append(
            ToolStep(tool_name=tool_name, display_name=display_name, arguments=arguments, result=result)
        )

    def sources_used(self) -> List[str]:
        """Return the display name of every tool invoked, in call order,
        with duplicates removed but original order preserved."""
        seen: Dict[str, None] = {}
        for step in self.tool_steps:
            seen.setdefault(step.display_name, None)
        return list(seen.keys())


@dataclass
class WorkflowState:
    """The single object every node reads from and writes to.

    Attributes:
        user_message: The current turn's raw user message.
        conversation_state: The session's carried-forward state (see
            ``ai.session.ConversationState``). Nodes read the
            previously active filter/metric from here and the
            Filtering node writes the merged result back.
        intent: Set by the Intent Detection node.
        resolved_entities: Set by the Entity Resolution node.
        active_filter: Set by the Filtering node.
        analysis_results: Appended to by the Analysis node, one entry
            per tool call.
        final_response: Set by the Response node -- the text shown to
            the user.
        trace: Accumulates node- and tool-level execution detail for
            explainability and for updating ``conversation_state``
            after the turn completes.
    """

    user_message: str
    conversation_state: ConversationState
    intent: Optional[Intent] = None
    resolved_entities: ResolvedEntities = field(default_factory=ResolvedEntities)
    active_filter: Filter = field(default_factory=Filter)
    analysis_results: List[ToolResult] = field(default_factory=list)
    final_response: Optional[str] = None
    trace: ExecutionTrace = field(default_factory=ExecutionTrace)


class WorkflowNode(ABC):
    """One stage of the workflow graph.

    Attributes:
        name: A short, stable, code-friendly identifier (e.g.
            ``"intent_detection"``).
        display_label: A human-readable label for the live progress
            checklist (e.g. ``"Understanding request"``).
    """

    name: ClassVar[str]
    display_label: ClassVar[str]

    @abstractmethod
    def should_run(self, state: WorkflowState) -> bool:
        """Return whether this node applies to the current request.

        Args:
            state: The workflow state so far.

        Returns:
            ``True`` if :meth:`run` should be called.
        """
        raise NotImplementedError

    @abstractmethod
    def run(self, state: WorkflowState) -> WorkflowState:
        """Execute this node.

        Args:
            state: The workflow state so far.

        Returns:
            The (possibly same, mutated) state, for the next node.
        """
        raise NotImplementedError


class WorkflowGraph:
    """Executes an ordered list of :class:`WorkflowNode`s against one
    :class:`WorkflowState`."""

    def __init__(self, nodes: List[WorkflowNode]) -> None:
        self._nodes = nodes

    def execute(self, initial_state: WorkflowState, progress_cb: Optional[ProgressCallback] = None) -> WorkflowState:
        """Run every applicable node in order.

        Args:
            initial_state: The state to start from.
            progress_cb: Optional callback invoked as
                ``(node_name, display_label, "running"|"done")`` around
                each node that actually runs (nodes whose
                ``should_run()`` returns ``False`` are never reported,
                so the checklist only ever shows what genuinely
                happened).

        Returns:
            The final :class:`WorkflowState`, after every applicable
            node has run.
        """
        state = initial_state
        for node in self._nodes:
            if not node.should_run(state):
                logger.debug("Skipping node '%s' (should_run() is False)", node.name)
                continue

            if progress_cb:
                progress_cb(node.name, node.display_label, "running")

            started_at = time.monotonic()
            state = node.run(state)
            elapsed = time.monotonic() - started_at
            state.trace.record_node(node.name, elapsed)
            logger.debug("Node '%s' completed in %.3fs", node.name, elapsed)

            if progress_cb:
                progress_cb(node.name, node.display_label, "done")

        return state
