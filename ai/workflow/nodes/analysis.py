"""
ai.workflow.nodes.analysis
=============================
Runs the iterative tool-use loop: repeatedly calls the configured
``LLMProvider`` with the full analysis tool catalog until the model
either stops requesting tools or a bounded number of tool calls has
been made, recording every call into ``WorkflowState.trace`` and
``WorkflowState.analysis_results``.

This is the node the approved architecture describes as containing "the
iterative Bedrock tool-use loop" (Architecture Plan v3 Section 1.3) --
deciding which *specific* tool(s) to call and in what order is this
node's responsibility; ``PlanningNode`` upstream only confirms a
suitable tool category exists at all.

Tool execution itself never talks to an LLM and never fails silently:
an unknown tool name or a tool-level failure (``ToolError``, already
handled gracefully by ``ToolRegistry.dispatch``) is reported back to the
model as a normal tool result describing the problem, giving the model
a chance to explain the limitation in its final answer rather than the
turn failing outright.
"""

from __future__ import annotations

import logging
from typing import ClassVar, List

from ai.context import BusinessContext
from ai.data.filters import Filter
from ai.llm.provider import LLMProvider, Message, ToolResultBlock, ToolSchema
from ai.tools.registry import ToolRegistry, UnknownToolError
from ai.workflow.graph import Intent, WorkflowNode, WorkflowState

logger = logging.getLogger(__name__)


class AnalysisNode(WorkflowNode):
    """Runs the tool-use loop against the registered analysis tools."""

    name: ClassVar[str] = "analysis"
    display_label: ClassVar[str] = "Running analysis"

    #: Bounded so a model that keeps requesting tools indefinitely
    #: cannot turn one user message into an unbounded number of calls
    #: (Architecture Plan v3 Section 17's "planner loop cost/latency"
    #: risk mitigation).
    MAX_TOOL_CALLS: ClassVar[int] = 5

    def __init__(self, provider: LLMProvider, registry: ToolRegistry, context: BusinessContext) -> None:
        self._provider = provider
        self._registry = registry
        self._context = context

    def should_run(self, state: WorkflowState) -> bool:
        return state.intent == Intent.ANALYTICAL_QUESTION

    def run(self, state: WorkflowState) -> WorkflowState:
        tool_schemas = [
            ToolSchema(name=s["name"], description=s["description"], input_schema=s["input_schema"])
            for s in self._registry.schemas_for_bedrock()
        ]
        messages: List[Message] = [Message(role="user", text=self._build_prompt(state))]

        for call_number in range(1, self.MAX_TOOL_CALLS + 1):
            response = self._provider.converse(messages, tools=tool_schemas)
            if response.tool_use is None:
                logger.debug("Analysis loop stopped after %d tool call(s): model returned a direct response.", call_number - 1)
                break

            tool_use = response.tool_use
            messages.append(Message(role="assistant", tool_use=tool_use))

            try:
                tool = self._registry.get(tool_use.name)
                result = self._registry.dispatch(tool_use.name, tool_use.arguments, self._context)
                state.analysis_results.append(result)
                state.trace.record_tool(tool_use.name, tool.display_name, tool_use.arguments, result)
                result_text = result.summary
            except UnknownToolError as exc:
                logger.warning("Model requested an unregistered tool: %s", exc)
                result_text = f"Error: {exc}"

            messages.append(
                Message(role="user", tool_result=ToolResultBlock(tool_use_id=tool_use.tool_use_id, text=result_text))
            )
        else:
            logger.warning(
                "Analysis loop reached the maximum of %d tool call(s) without the model "
                "signaling it was done; proceeding to Response with whatever results were gathered.",
                self.MAX_TOOL_CALLS,
            )

        return state

    @staticmethod
    def _build_prompt(state: WorkflowState) -> str:
        """Build the prompt for the tool-use loop, including the active
        filter as explicit context so the model doesn't have to
        re-derive it from the raw message alone."""
        filter_description = AnalysisNode._describe_filter(state.active_filter)
        return f"{state.user_message}\n\n(Active filter context: {filter_description})"

    @staticmethod
    def _describe_filter(data_filter: Filter) -> str:
        parts = []
        if data_filter.client:
            parts.append(f"client={data_filter.client}")
        if data_filter.poc:
            parts.append(f"poc={data_filter.poc}")
        if data_filter.section:
            parts.append(f"section={data_filter.section}")
        if data_filter.quarters:
            parts.append(f"quarters={data_filter.quarters}")
        return ", ".join(parts) if parts else "none"
