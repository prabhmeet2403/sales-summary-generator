"""
ai.session
===========
Conversational state that persists across turns within one chat
session, plus the session container itself.

``ConversationState`` is built directly on ``ai.data.filters.Filter``
rather than a second, parallel "remembered entities" structure --
"what the user means" and "what data gets queried" are represented by
exactly one object so they cannot drift apart from each other. See
``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 1.2 (Filtering
node) and Revision 2 Section 5 for the approved design.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import List, Optional

from ai.data.filters import Filter
from ai.llm.provider import Message

logger = logging.getLogger(__name__)


@dataclass
class ConversationState:
    """Carries forward across turns so a follow-up like "show margin"
    can be resolved without the user repeating the client/quarter they
    already specified.

    Attributes:
        active_filter: The most specific :class:`Filter` established so
            far in this conversation. Updated after each turn with
            whatever entities/scoping that turn's tool calls used.
        active_metric: Which metric ("revenue" or "margin") the
            conversation is currently focused on, if any.
        last_tool_called: The name of the most recently invoked tool,
            for simple heuristics (e.g. suggested follow-ups in a later
            phase).
    """

    active_filter: Filter = field(default_factory=Filter)
    active_metric: Optional[str] = None
    last_tool_called: Optional[str] = None

    def merge_filter(self, new_filter: Filter) -> Filter:
        """Merge a newly resolved filter with what's already active,
        preferring the new value for any field it actually sets, and
        falling back to the existing active value otherwise.

        This is the concrete mechanism behind ellipsis resolution: if
        this turn didn't mention a client, the previous turn's client
        carries forward; if it did, the new client takes over.

        Args:
            new_filter: The filter resolved from this turn's message.

        Returns:
            The merged filter. Does not mutate ``self.active_filter``;
            callers are expected to assign the result back explicitly
            once a turn completes successfully.
        """
        return Filter(
            client=new_filter.client if new_filter.client is not None else self.active_filter.client,
            poc=new_filter.poc if new_filter.poc is not None else self.active_filter.poc,
            section=new_filter.section if new_filter.section is not None else self.active_filter.section,
            quarters=new_filter.quarters if new_filter.quarters is not None else self.active_filter.quarters,
            months=new_filter.months if new_filter.months is not None else self.active_filter.months,
            years=new_filter.years if new_filter.years is not None else self.active_filter.years,
            role=new_filter.role if new_filter.role is not None else self.active_filter.role,
            min_revenue=new_filter.min_revenue if new_filter.min_revenue is not None else self.active_filter.min_revenue,
            max_revenue=new_filter.max_revenue if new_filter.max_revenue is not None else self.active_filter.max_revenue,
            min_margin=new_filter.min_margin if new_filter.min_margin is not None else self.active_filter.min_margin,
            max_margin=new_filter.max_margin if new_filter.max_margin is not None else self.active_filter.max_margin,
            min_confidence_pct=(
                new_filter.min_confidence_pct if new_filter.min_confidence_pct is not None
                else self.active_filter.min_confidence_pct
            ),
        )


@dataclass
class ChatSession:
    """One conversation: its message history and its conversational state.

    Attributes:
        session_id: Identifies this conversation.
        state: Persists across turns (see :class:`ConversationState`).
        messages: The literal transcript, oldest first.
    """

    session_id: str
    state: ConversationState = field(default_factory=ConversationState)
    messages: List[Message] = field(default_factory=list)
