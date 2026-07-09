"""
ai.ui.message_render
=======================
Renders one chat turn using Streamlit's native chat elements
(``st.chat_message``) rather than custom HTML, so the AI Assistant's
chat bubbles inherit Streamlit's own styling and accessibility behavior
rather than a second, hand-rolled visual language.
"""

from __future__ import annotations

from typing import List, Optional

import streamlit as st


def render_message(role: str, text: str, sources_used: Optional[List[str]] = None) -> None:
    """Render one chat turn.

    Args:
        role: ``"user"`` or ``"assistant"`` -- passed straight to
            ``st.chat_message`` for its built-in avatar/alignment.
        text: The message text (rendered as Markdown).
        sources_used: If provided and non-empty, rendered as a
            collapsed "Sources Used" expander beneath the message --
            the explainability trail for an assistant turn that used
            one or more tools.
    """
    with st.chat_message(role):
        st.markdown(text)
        if sources_used:
            with st.expander("Sources Used", expanded=False):
                for source in sources_used:
                    st.markdown(f"✓ {source}")
