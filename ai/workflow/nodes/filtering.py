"""
ai.workflow.nodes.filtering
==============================
Builds this turn's active ``Filter`` from resolved entities, merged with
whatever the conversation's carried-forward state already had active.

This is the concrete mechanism behind resolving a follow-up like "show
margin" without the user repeating "for HPE, Q2 vs Q3": this turn
resolves no new client, so ``ConversationState.merge_filter`` falls back
to the previous turn's client, exactly as documented on
``ai.session.ConversationState``.
"""

from __future__ import annotations

from typing import ClassVar

from ai.data.filters import Filter
from ai.workflow.graph import Intent, WorkflowNode, WorkflowState


class FilteringNode(WorkflowNode):
    """Merges this turn's resolved entities into the conversation's
    active filter."""

    name: ClassVar[str] = "filtering"
    display_label: ClassVar[str] = "Applying filters"

    def should_run(self, state: WorkflowState) -> bool:
        return state.intent == Intent.ANALYTICAL_QUESTION

    def run(self, state: WorkflowState) -> WorkflowState:
        new_filter = Filter(client=state.resolved_entities.client, poc=state.resolved_entities.poc)
        state.active_filter = state.conversation_state.merge_filter(new_filter)
        return state
