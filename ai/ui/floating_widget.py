"""
ai.ui.floating_widget
========================
Copilot-style floating assistant: a circular button that opens an
enterprise-style side panel. Reuses the same AIService/session as the
full "AI Assistant" page (ai.ui.chat_page) -- same session id, same
conversation, same chat rendering, same workflow graph -- so a user
isn't juggling two separate threads or two separate code paths.

UI/UX only in this module: no business logic, no calls into AIService
beyond ``ask``/``history_for``, which the full chat page already relies
on identically.
"""

from __future__ import annotations

from typing import Optional

import streamlit as st

from ai.context import BusinessContext, IncompleteGenerationResultError
from ai.llm.factory import UnknownProviderError
from ai.service import AIService, AIServiceError
from ai.settings import AISettingsError
from ai.ui.chat_page import CHAT_SESSION_ID, get_or_build_context, get_or_build_service
from ai.ui.message_render import render_message
from ai.ui.progress_display import ProgressChecklist
from ai.ui.suggestions import LAST_SOURCES_KEY, render_follow_up_suggestions, render_welcome_panel

_OPEN_KEY = "floating_assistant_open"


def render_floating_assistant(gen_result: object) -> None:
    """Render the floating toggle button, or the open panel -- never both.

    Args:
        gen_result: The current ``gui.runner.GenerationResult`` (or
            ``None``). The panel shows an empty/not-configured state as
            appropriate, matching ``ai.ui.chat_page.render``'s own
            handling.
    """
    is_open = st.session_state.get(_OPEN_KEY, False)
    toggle_slot = st.empty()
    panel_slot = st.empty()

    if not is_open:
        with toggle_slot.container():
            if st.button("✨", key="floating_assistant_toggle_btn", help="Chatbot Prabh"):
                st.session_state[_OPEN_KEY] = True
                is_open = True
        if not is_open:
            return
        # Just opened: show the panel in this same run rather than
        # waiting for a second, unrelated interaction.
        toggle_slot.empty()
        with panel_slot.container(key="floating_assistant_panel"):
            _render_panel(gen_result)
        return

    # Already open at the start of this run.
    with panel_slot.container(key="floating_assistant_panel"):
        still_open = _render_panel(gen_result)
    if still_open:
        return
    # Just closed: show the button in this same run rather than
    # waiting for a second, unrelated interaction. Does not need
    # st.rerun() -- that call previously caused unrelated session-state
    # timing issues elsewhere in this app.
    panel_slot.empty()
    with toggle_slot.container():
        st.button("✨", key="floating_assistant_toggle_btn", help="Chatbot Prabh")


def _render_panel(gen_result: object) -> bool:
    """Render the panel's contents.

    Returns:
        Whether the panel should remain open after this run. When the
        close button is clicked, the header above still renders once
        (unavoidable, since the button lives in it), but the caller
        immediately clears the whole panel slot afterward -- avoiding
        ``st.rerun()``, which has caused unrelated session-state timing
        issues elsewhere in this app (see the "Respect Phase 1" fix in
        an earlier phase).
    """
    header_col, close_col = st.columns([6, 1])
    with header_col:
        st.markdown(
            '<div class="copilot-title">✨ Chatbot Prabh</div>'
            '<div class="copilot-subtitle">Ask questions about the generated '
            "Sales &amp; Forecast Summary.</div>",
            unsafe_allow_html=True,
        )
    with close_col:
        if st.button("✕", key="floating_assistant_close_btn"):
            st.session_state[_OPEN_KEY] = False
            return False

    if gen_result is None or not getattr(gen_result, "success", False):
        _render_empty_state("👋", "Generate a Summary first, then come back to ask questions about it.")
        return True

    try:
        context = get_or_build_context(gen_result)
    except IncompleteGenerationResultError as exc:
        _render_empty_state("⚠️", str(exc))
        return True

    try:
        service = get_or_build_service(context)
    except (AISettingsError, UnknownProviderError):
        _render_not_configured()
        return True

    history = service.history_for(CHAT_SESSION_ID)
    suggestion_slot = st.empty()
    if not history:
        with suggestion_slot.container():
            suggested_prompt = _render_welcome(context)
    else:
        for message in history:
            render_message(message.role, message.text or "")
        with suggestion_slot.container():
            suggested_prompt = render_follow_up_suggestions(key_prefix=f"floating_assistant_{len(history)}")

    user_message = st.chat_input("Ask a question…", key="floating_assistant_input")
    message_to_send = user_message or suggested_prompt
    if message_to_send:
        # A new message is about to be answered -- clear the now-stale
        # welcome/follow-up row and render a fresh one for *this* turn
        # once it completes, rather than waiting for an unrelated
        # future rerun. Historical messages above were already rendered
        # directly (not inside this placeholder), so they are untouched.
        # Keyed by the post-turn message count so the same key persists
        # correctly across reruns (a button's click state is tied to
        # its key, not to which code path rendered it).
        suggestion_slot.empty()
        _handle_message(service, message_to_send)
        render_follow_up_suggestions(key_prefix=f"floating_assistant_{len(service.history_for(CHAT_SESSION_ID))}")

    _autoscroll()
    return True


def _handle_message(service: AIService, message: str) -> None:
    """Send one message through the shared AIService and render the
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
            st.markdown(f'<div class="copilot-error">⚠ {exc.message}</div>', unsafe_allow_html=True)
            return
        checklist.clear()
        st.markdown(turn.text)
        if turn.sources_used:
            with st.expander(f"Sources Used: {', '.join(turn.sources_used)}", expanded=False):
                for source in turn.sources_used:
                    st.markdown(f"✓ {source}")

    st.session_state[LAST_SOURCES_KEY] = turn.sources_used


def _render_welcome(context: BusinessContext) -> Optional[str]:
    """Render the shared welcome header/subtext and context-aware
    quick-action suggestions -- identical content to the full "AI
    Assistant" page's welcome panel (``ai.ui.chat_page.render``), just
    inside this panel's own container.
    """
    return render_welcome_panel(context, key_prefix="floating_assistant")


def _render_empty_state(icon: str, message: str) -> None:
    st.markdown(
        f'<div class="copilot-empty"><div class="copilot-empty-emoji">{icon}</div>'
        f"<p>{message}</p></div>",
        unsafe_allow_html=True,
    )
    st.chat_input("Configure AI to start chatting…", key="floating_assistant_input_disabled", disabled=True)


def _render_not_configured() -> None:
    st.markdown(
        '<div class="copilot-config-card">'
        '<div class="copilot-config-icon">⚠</div>'
        '<div class="copilot-config-title">AI isn\'t configured yet</div>'
        "<p>Chatbot Prabh has been installed successfully.</p>"
        "<p><strong>To enable AI:</strong></p>"
        "<ol><li>Configure Bedrock credentials</li><li>Restart the application</li></ol>"
        '<div class="copilot-config-doclink">Setup Guide — ai/README.md</div>'
        "</div>",
        unsafe_allow_html=True,
    )
    st.chat_input("Configure AI to start chatting…", key="floating_assistant_input_disabled", disabled=True)


def _autoscroll() -> None:
    """Scroll the panel to its newest content. A small, self-contained
    script scoped to the panel's own stable CSS class -- no component
    framework, no dependency beyond what unsafe_allow_html already
    permits elsewhere in this app."""
    st.markdown(
        """
        <script>
        const panel = window.parent.document.querySelector('.st-key-floating_assistant_panel');
        if (panel) { panel.scrollTop = panel.scrollHeight; }
        </script>
        """,
        unsafe_allow_html=True,
    )
