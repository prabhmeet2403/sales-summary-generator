"""
ai.workflow.nodes.intent_detection
=====================================
Classifies the user's message into one of the categories
``ai.workflow.graph.Intent`` defines, using a single small, targeted
LLM call. This determines whether the rest of the analytical pipeline
(Planning through Analysis) should run at all -- a greeting or an
explicitly unsupported request (a chart, a dashboard, an exported file
-- none of which exist yet in Phase 2b) skips straight to the Response
node rather than spending an LLM call trying to select an irrelevant
tool.

A request for a text-based executive/board/management summary of the
data is explicitly called out as ``analytical_question`` (it's answered
directly in this chat by ``ai.tools.executive_summary``, not exported
anywhere), since without that clarification the word "report"/"summary"
in a message like "summarize this report" or "prepare a board summary"
reads as pattern-matching the unsupported-request examples (a PDF/Word
report) rather than a normal data question.
"""

from __future__ import annotations

import logging
from typing import ClassVar

from ai.llm.provider import LLMProvider, Message
from ai.workflow.graph import Intent, WorkflowNode, WorkflowState

logger = logging.getLogger(__name__)

_CLASSIFICATION_PROMPT = """\
Classify the following user message into exactly one category:

- analytical_question: a question about business data (revenue, margin, \
clients, POCs, comparisons, rankings, etc.), INCLUDING a request for a \
text summary of the data itself -- an executive summary, board summary, \
management summary, report summary, or overall summary, or a request \
like "summarize this report" / "what should management know" / "what \
are the key highlights", all count as analytical_question, since these \
are answered directly in this chat from existing data, not exported \
anywhere.
- general_conversation: a greeting, thanks, or small talk with no \
business data question
- unsupported_request: explicitly asks for a capability that is not yet \
available -- specifically, a visual chart or dashboard, or a file the \
user could download/export (a PDF, Word, or Excel file distinct from \
this chat). A request to summarize or explain the data in the chat \
itself is analytical_question, not this category, even if the user \
uses the word "report" or "summary" -- only classify as \
unsupported_request when a chart/dashboard/downloadable file is \
explicitly what's being asked for.

Message: "{message}"

Respond with exactly one word: analytical_question, general_conversation, \
or unsupported_request. Do not add any other text."""


class IntentDetectionNode(WorkflowNode):
    """Classifies the user's message into an :class:`Intent`."""

    name: ClassVar[str] = "intent_detection"
    display_label: ClassVar[str] = "Understanding request"

    def __init__(self, provider: LLMProvider) -> None:
        self._provider = provider

    def should_run(self, state: WorkflowState) -> bool:
        return True  # always the first node

    def run(self, state: WorkflowState) -> WorkflowState:
        prompt = _CLASSIFICATION_PROMPT.format(message=state.user_message)
        response = self._provider.converse([Message(role="user", text=prompt)])
        state.intent = self._parse_intent((response.text or "").strip().lower())
        return state

    @staticmethod
    def _parse_intent(raw_text: str) -> Intent:
        for intent in Intent:
            if intent.value in raw_text:
                return intent
        logger.warning(
            "Could not parse an Intent from the model's classification response %r; "
            "defaulting to ANALYTICAL_QUESTION so the request is still attempted.",
            raw_text,
        )
        return Intent.ANALYTICAL_QUESTION
