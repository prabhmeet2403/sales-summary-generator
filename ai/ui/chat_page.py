"""
ai.ui.chat_page
==================
The "AI Assistant" Streamlit page: renders an empty state until a
Summary has been generated, then a chat interface backed by
``ai.service.AIService``.

This module is the only place in ``ai`` (besides the rest of ``ai.ui``)
permitted to import ``streamlit`` -- it calls ``AIService`` exactly the
way a future REST API's request handler would call it, through the same
public methods, per Architecture Plan Section 11.1's service-facade rule.
"""

from __future__ import annotations

import logging

import streamlit as st

from ai.context import BusinessContext, IncompleteGenerationResultError
from ai.llm.factory import UnknownProviderError, get_provider
from ai.service import AIService, AIServiceError
from ai.settings import AISettingsError, load_ai_settings
from ai.ui.message_render import render_message
from ai.ui.progress_display import ProgressChecklist
from ai.ui.suggestions import LAST_SOURCES_KEY, render_follow_up_suggestions, render_welcome_panel
from components.ui import empty_state, info_box

logger = logging.getLogger(__name__)

CHAT_SESSION_ID = "main"


def render(gen_result: object) -> None:
    """Render the AI Assistant page.

    Args:
        gen_result: The current ``gui.runner.GenerationResult`` (as
            stored in ``st.session_state["gen_result"]``), or ``None``
            if no generation has run yet this session.
    """
    if gen_result is None or not getattr(gen_result, "success", False):
        empty_state(
            title="No Summary Generated Yet",
            description="Upload a Master workbook and click \"Generate Summary\" on the "
            "Upload & Generate page first -- Chatbot Prabh answers questions about "
            "the most recently generated Summary.",
            icon="🤖",
        )
        return

    try:
        context = get_or_build_context(gen_result)
    except IncompleteGenerationResultError as exc:
        empty_state(title="Summary Data Unavailable", description=str(exc), icon="⚠️")
        return

    try:
        service = get_or_build_service(context)
    except (AISettingsError, UnknownProviderError) as exc:
        info_box(
            f"Chatbot Prabh is not configured yet: {exc}. "
            "Set the required environment variables (see ai/README.md) and reload the page.",
            variant="red",
        )
        return

    st.markdown("#### Chatbot Prabh")
    st.caption(f"Ask questions about the {context.target_year} Summary ({len(context.groups_df)} groups).")

    history = service.history_for(CHAT_SESSION_ID)
    for message in history:
        render_message(message.role, message.text or "")

    suggestion_slot = st.empty()
    with suggestion_slot.container():
        if not history:
            suggested_prompt = render_welcome_panel(context, key_prefix="chatpage_welcome")
        else:
            suggested_prompt = render_follow_up_suggestions(key_prefix=f"chatpage_{len(history)}")

    user_message = st.chat_input("Ask a question about this Summary…")
    message_to_send = user_message or suggested_prompt
    if not message_to_send:
        return

    # A new message is about to be answered -- clear the now-stale
    # welcome/follow-up row and render a fresh one for *this* turn once
    # it completes, rather than waiting for an unrelated future rerun.
    # Keyed by the post-turn message count so the same key persists
    # correctly across reruns (Streamlit ties a button's click state to
    # its key, not to which code path rendered it).
    suggestion_slot.empty()
    _handle_message(service, message_to_send)
    render_follow_up_suggestions(key_prefix=f"chatpage_{len(service.history_for(CHAT_SESSION_ID))}")


def _handle_message(service: AIService, message: str) -> None:
    """Send one message through the shared ``AIService`` and render the
    exchange -- the single code path both the chat input and every
    suggestion chip use, so there is exactly one place that calls
    ``service.ask``."""
    render_message("user", message)

    with st.chat_message("assistant"):
        progress_placeholder = st.empty()
        checklist = ProgressChecklist(progress_placeholder)
        try:
            turn = service.ask(CHAT_SESSION_ID, message, progress_cb=checklist.update)
        except AIServiceError as exc:
            checklist.clear()
            st.error(exc.message)
            logger.error("AI Assistant request failed: %s", exc.message)
            return

        checklist.clear()
        st.markdown(turn.text)
        if turn.sources_used:
            with st.expander(f"Sources Used: {', '.join(turn.sources_used)}", expanded=False):
                for source in turn.sources_used:
                    st.markdown(f"✓ {source}")

    st.session_state[LAST_SOURCES_KEY] = turn.sources_used


def get_or_build_context(gen_result: object) -> BusinessContext:
    """Return the cached :class:`BusinessContext` for this generation,
    building (and caching) a new one if this is a new generation.

    A new generation's fingerprint differs from any previously cached
    one, so this naturally invalidates and rebuilds -- including
    discarding the previous :class:`AIService` (and its conversation
    history) via :func:`get_or_build_service`'s matching cache key.
    """
    elapsed_seconds = st.session_state.get("elapsed_seconds")
    candidate = BusinessContext.from_generation_result(gen_result, elapsed_seconds=elapsed_seconds)  # type: ignore[arg-type]
    cached = st.session_state.get("ai_context")
    if cached is None or cached.fingerprint != candidate.fingerprint:
        st.session_state["ai_context"] = candidate
        st.session_state.pop("ai_service", None)  # force AIService rebuild too
        logger.info("Built a new BusinessContext (fingerprint=%s)", candidate.fingerprint)
        return candidate
    return cached


def get_or_build_service(context: BusinessContext) -> AIService:
    """Return the cached :class:`AIService` for this context, building
    one (with a freshly loaded provider) if none is cached yet."""
    cached = st.session_state.get("ai_service")
    if cached is not None:
        return cached

    settings = load_ai_settings()
    provider = get_provider(settings)
    service = AIService(context, provider)
    st.session_state["ai_service"] = service
    logger.info("Built a new AIService for context fingerprint=%s", context.fingerprint)
    return service
