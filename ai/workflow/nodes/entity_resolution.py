"""
ai.workflow.nodes.entity_resolution
======================================
Recognizes client and POC names mentioned in the user's message by
matching against the exact, finite vocabulary of names already present
in this generation's data (``BusinessContext.group_names()`` /
``poc_names()``).

A substring match against a known, finite vocabulary is deliberately
used here instead of an LLM call: since every valid client/POC name is
already enumerable from the data itself, matching against that list is
strictly more reliable than asking a model to recall or guess a name,
and costs nothing. The longest matching name is preferred over a
shorter one to avoid a shorter name matching inside a longer one
(e.g. preferring "Hitachi Asia" over an unrelated shorter match).
"""

from __future__ import annotations

from typing import ClassVar, List, Optional

from ai.context import BusinessContext
from ai.workflow.graph import Intent, ResolvedEntities, WorkflowNode, WorkflowState


class EntityResolutionNode(WorkflowNode):
    """Finds known client/POC names mentioned in the user's message."""

    name: ClassVar[str] = "entity_resolution"
    display_label: ClassVar[str] = "Resolving client"

    def __init__(self, context: BusinessContext) -> None:
        self._context = context

    def should_run(self, state: WorkflowState) -> bool:
        return state.intent == Intent.ANALYTICAL_QUESTION

    def run(self, state: WorkflowState) -> WorkflowState:
        message_lower = state.user_message.lower()
        client = self._longest_match(message_lower, self._context.group_names())
        poc = self._longest_match(message_lower, self._context.poc_names())
        state.resolved_entities = ResolvedEntities(client=client, poc=poc)
        return state

    @staticmethod
    def _longest_match(message_lower: str, candidates: List[str]) -> Optional[str]:
        matches = [candidate for candidate in candidates if candidate.lower() in message_lower]
        return max(matches, key=len) if matches else None
