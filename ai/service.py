"""
ai.service
===========
``AIService`` is the single entry point anything outside the ``ai``
package should call. The Streamlit "AI Assistant" page calls this; a
future REST API would call the exact same methods with zero changes
below this facade -- see Architecture Plan Section 11.1.

Phase 2b: :meth:`AIService.ask` runs the full nine-node workflow graph
(``ai.workflow``) for every message, dispatching through the plugin
tool registry (``ai.tools``) as the graph's Analysis node determines is
relevant. This replaces Phase 2a's single direct provider call -- the
documented, planned evolution from "Phase 2a: one grounded call, no
graph yet" to "Phase 2b: the workflow graph replaces this method's
internals" (see ``ai/README.md``'s Phase 2a notes). The public
``ask(session_id, message)`` signature is unchanged from Phase 2a.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from ai.context import BusinessContext
from ai.llm.provider import LLMProvider, LLMProviderError, Message
from ai.session import ChatSession
from ai.tools.registry import ToolRegistry, discover_tools
from ai.workflow.factory import build_default_workflow_graph
from ai.workflow.graph import ProgressCallback, WorkflowState

logger = logging.getLogger(__name__)


class AIServiceError(Exception):
    """Raised when :class:`AIService` cannot fulfill a request.

    Attributes:
        message: Plain-language description of the failure.
        retryable: Whether the caller may reasonably retry the same
            request.
    """

    def __init__(self, message: str, *, retryable: bool = False) -> None:
        super().__init__(message)
        self.message = message
        self.retryable = retryable


@dataclass(frozen=True)
class AssistantTurn:
    """One complete assistant response.

    Attributes:
        text: The assistant's natural-language response.
        sources_used: Display names of every tool invoked while
            answering this turn, in call order (the "Sources Used"
            explainability trail). Empty for a general-conversation or
            unsupported-request turn that used no tools.
    """

    text: str
    sources_used: List[str] = field(default_factory=list)


class AIService:
    """Facade over the AI platform's capabilities for one generated Summary.

    Attributes:
        context: The :class:`~ai.context.BusinessContext` this service
            answers questions about.
    """

    def __init__(
        self,
        context: BusinessContext,
        provider: LLMProvider,
        registry: Optional[ToolRegistry] = None,
    ) -> None:
        """Construct a service bound to one generation, one provider,
        and one tool registry.

        Args:
            context: The business data this service will answer
                questions about.
            provider: The LLM provider to use for every call this
                service makes.
            registry: The tool registry to dispatch through. Defaults
                to auto-discovering every registered tool under
                ``ai.tools`` (:func:`ai.tools.registry.discover_tools`).
                Exposed as a parameter so tests can supply a smaller,
                controlled registry.
        """
        self.context = context
        self._provider = provider
        self._registry = registry if registry is not None else discover_tools()
        self._graph = build_default_workflow_graph(provider, self._registry, context)
        self._sessions: Dict[str, ChatSession] = {}

    def ask(
        self,
        session_id: str,
        message: str,
        progress_cb: Optional[ProgressCallback] = None,
    ) -> AssistantTurn:
        """Ask a question within a given conversation session.

        Runs the full workflow graph (intent detection through response
        synthesis). Conversation state (the active filter/metric) is
        retained per ``session_id`` for the lifetime of this
        :class:`AIService` instance, so a follow-up like "show margin"
        resolves using the client/quarter established in an earlier
        turn without the user repeating it.

        Args:
            session_id: Identifies which conversation this message
                belongs to. A new, unseen ``session_id`` starts a fresh
                conversation with empty state.
            message: The user's message.
            progress_cb: Optional callback for live per-node progress
                (see ``ai.workflow.graph.WorkflowGraph.execute``).

        Returns:
            The assistant's response, including the sources-used trail.

        Raises:
            AIServiceError: If the underlying provider call fails.
        """
        session = self._sessions.setdefault(session_id, ChatSession(session_id=session_id))
        session.messages.append(Message(role="user", text=message))

        initial_state = WorkflowState(user_message=message, conversation_state=session.state)

        try:
            final_state = self._graph.execute(initial_state, progress_cb=progress_cb)
        except LLMProviderError as exc:
            # Roll back the just-appended user message so a failed
            # request does not leave a dangling, unanswered turn in
            # the displayed transcript.
            session.messages.pop()
            logger.error("AIService.ask failed for session %s: %s", session_id, exc.message)
            raise AIServiceError(
                f"The assistant could not answer that question right now: {exc.message}",
                retryable=exc.retryable,
            ) from exc

        response_text = final_state.final_response or "I'm not able to provide an answer right now."
        session.messages.append(Message(role="assistant", text=response_text))

        # Persist this turn's filter back onto the session's carried-
        # forward state, so the next turn's Filtering node can fall
        # back to it (see ai.session.ConversationState.merge_filter).
        session.state.active_filter = final_state.active_filter
        if final_state.trace.tool_steps:
            session.state.last_tool_called = final_state.trace.tool_steps[-1].tool_name

        sources_used = final_state.trace.sources_used()
        logger.info(
            "AIService.ask completed for session %s (intent=%s, tools_used=%s)",
            session_id, final_state.intent, sources_used,
        )
        return AssistantTurn(text=response_text, sources_used=sources_used)

    def reset_session(self, session_id: str) -> None:
        """Discard conversation history and state for the given session.

        Args:
            session_id: The session to clear. Clearing an unknown
                session id is a no-op, not an error.
        """
        self._sessions.pop(session_id, None)
        logger.debug("Session cleared: %s", session_id)

    def history_for(self, session_id: str) -> List[Message]:
        """Return the current displayed conversation transcript for a session.

        Args:
            session_id: The session to inspect.

        Returns:
            A copy of the message list (mutating the returned list does
            not affect the service's internal state). This is the
            top-level user/assistant transcript shown in the chat UI --
            it does not include the internal tool-use back-and-forth a
            given turn's Analysis node conducted, which is discarded
            once that turn's final response is synthesized.
        """
        session = self._sessions.get(session_id)
        return list(session.messages) if session else []
