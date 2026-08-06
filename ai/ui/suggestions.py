"""
ai.ui.suggestions
====================
Business-focused quick-action and follow-up suggestion chips for
Chatbot Prabh, shared identically by the full "AI Assistant" page
(``ai.ui.chat_page``) and the floating widget (``ai.ui.floating_widget``)
so the two surfaces behave identically rather than maintaining two
copies of the same suggestion content.

A suggestion is nothing more than a ``(label, prompt)`` pair rendered
as a button. Clicking one submits ``prompt`` through the exact same
``AIService.ask()`` call a typed question would use -- this module
never calls ``AIService`` itself. It only renders buttons and returns
the clicked prompt text (or ``None``); the caller passes that straight
into the same message-handling function it already uses for typed
input, so there is exactly one code path from "user intent" (typed or
clicked) to ``AIService.ask()``.

The "context-aware" reordering below is deliberately simple: a handful
of ``if``/``else`` checks over facts already sitting on
``BusinessContext``/``ValidationReport`` (warnings, missing comments,
quarter/client counts). It reorders which suggestions appear first
within their existing group; it never adds, removes, or invents a
suggestion, and it is not a recommendation engine.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

import streamlit as st

from ai.context import BusinessContext

Suggestion = Tuple[str, str]  # (label, prompt)

#: Shared session-state key both surfaces write to after every
#: ``AIService.ask()`` call and read from when deciding which
#: follow-up chips to show under the most recent answer. Shared
#: (not per-surface) because the floating widget and the full page
#: read/write the exact same underlying conversation.
LAST_SOURCES_KEY = "prabh_last_sources_used"

WELCOME_HEADER = "Your Sales & Forecast Summary has been generated successfully."
WELCOME_SUBTEXT = "Select one of the suggested questions below or type your own."

# ---- Quick actions (welcome panel) ----------------------------------------

QUICK_ACTION_GROUPS: Dict[str, List[Suggestion]] = {
    "Executive Summary": [
        ("Generate Executive Summary", "Generate an Executive Summary."),
        ("Summarize Revenue", "Summarize total revenue."),
        ("Analyze Margins", "Analyze margins across all clients."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
        ("Forecast Analysis", "What does the forecast look like for the remaining quarters?"),
    ],
    "Validation": [
        ("Explain Validation Summary", "Explain the Validation Summary."),
        ("Review Warnings", "Review the warnings from this generation."),
        ("Review Comments", "Review the missing or unmatched comments."),
        ("Show Processing Statistics", "Show the processing statistics: groups processed, comments matched, and time taken."),
    ],
    "Business Analysis": [
        ("Compare Quarters", "Compare Q2 vs Q3 revenue."),
        ("Compare Clients", "Compare revenue across clients."),
        ("Highest Revenue Clients", "Which clients have the highest revenue?"),
        ("Lowest Margin Clients", "Which clients have the lowest margin?"),
    ],
    "Search": [
        ("Find Client", "Look up a specific client."),
        ("Find POC", "Look up a specific POC."),
        ("Review Comments", "Review the missing or unmatched comments."),
        ("Search Revenue", "Search for a specific client's revenue."),
    ],
}

# ---- Follow-up suggestions, keyed by which tool produced the answer ------

_FOLLOW_UPS_BY_TOOL: Dict[str, List[Suggestion]] = {
    "Revenue Analysis Tool": [
        ("Analyze Margins", "Analyze margins across all clients."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
        ("Compare Quarters", "Compare Q2 vs Q3 revenue."),
        ("Explain Validation Summary", "Explain the Validation Summary."),
        ("Generate Executive Summary", "Generate an Executive Summary."),
    ],
    "Margin Analysis Tool": [
        ("Summarize Revenue", "Summarize total revenue."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
        ("Lowest Margin Clients", "Which clients have the lowest margin?"),
    ],
    "Validation Summary Tool": [
        ("Review Missing Comments", "Review the missing or unmatched comments."),
        ("Summarize Revenue", "Summarize total revenue."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
        ("Generate Executive Summary", "Generate an Executive Summary."),
    ],
    "Quarter Comparison Tool": [
        ("Analyze Margins", "Analyze margins across all clients."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
        ("Forecast Analysis", "What does the forecast look like for the remaining quarters?"),
    ],
    "Client Lookup Tool": [
        ("Client Revenue", "What is this client's revenue?"),
        ("Client Margins", "What is this client's margin?"),
        ("Client Forecast", "What is the forecast for this client?"),
        ("Compare Quarterly Performance", "Compare this client's quarterly performance."),
    ],
    "POC Lookup Tool": [
        ("Client Revenue", "What is this client's revenue?"),
        ("Compare Clients", "Compare revenue across clients."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
    ],
    "Executive Summary Tool": [
        ("Explain Validation Summary", "Explain the Validation Summary."),
        ("Summarize Revenue", "Summarize total revenue."),
        ("Show Top Clients", "Who are the top clients by revenue?"),
    ],
}

_DEFAULT_FOLLOW_UPS: List[Suggestion] = [
    ("Summarize Revenue", "Summarize total revenue."),
    ("Show Top Clients", "Who are the top clients by revenue?"),
    ("Explain Validation Summary", "Explain the Validation Summary."),
]


def follow_up_suggestions(sources_used: Optional[List[str]]) -> List[Suggestion]:
    """Return the follow-up chips to show under one assistant turn,
    keyed off ``AssistantTurn.sources_used`` (the tool display name(s)
    the workflow graph's Analysis node actually called). Falls back to
    a small, generally-useful default set for a turn that used no
    tools (general conversation) or an unrecognized tool name.
    """
    if sources_used:
        for source in sources_used:
            if source in _FOLLOW_UPS_BY_TOOL:
                return _FOLLOW_UPS_BY_TOOL[source]
    return _DEFAULT_FOLLOW_UPS


def contextual_quick_action_groups(context: BusinessContext) -> Dict[str, List[Suggestion]]:
    """Return :data:`QUICK_ACTION_GROUPS` with items reordered within
    their existing group using simple, already-available
    ``BusinessContext``/``ValidationReport`` facts -- never adding,
    removing, or inventing a suggestion.
    """
    groups = {name: list(items) for name, items in QUICK_ACTION_GROUPS.items()}

    report = context.report
    if report is not None:
        has_warnings = len(report.warnings) > 0
        has_missing_comments = report.total_missing_comments > 0

        def _validation_priority(item: Suggestion) -> int:
            label = item[0]
            if label == "Review Comments" and has_missing_comments:
                return 0
            if label == "Review Warnings" and has_warnings:
                return 1
            return 2

        groups["Validation"].sort(key=_validation_priority)

    multiple_quarters = (
        not context.monthly_df.empty
        and "quarter" in context.monthly_df.columns
        and context.monthly_df["quarter"].nunique() > 1
    )
    multiple_clients = len(context.group_names()) > 1

    def _business_priority(item: Suggestion) -> int:
        label = item[0]
        if label == "Compare Quarters" and multiple_quarters:
            return 0
        if label == "Compare Clients" and multiple_clients:
            return 1
        return 2

    groups["Business Analysis"].sort(key=_business_priority)

    if multiple_clients:
        groups["Executive Summary"].sort(key=lambda item: 0 if item[0] == "Show Top Clients" else 1)

    # "Find Client"/"Find POC"/"Search Revenue" are otherwise too vague to
    # reliably invoke their tool (the model has no name to look up) -- fill
    # in a real client/POC name already present in this generation's own
    # data, so the same quick action works for whatever workbook is
    # currently loaded rather than a hardcoded example client.
    group_names = context.group_names()
    poc_names = context.poc_names()
    search_group = []
    for label, prompt in groups["Search"]:
        if label == "Find Client" and group_names:
            prompt = f"Look up {group_names[0]}."
        elif label == "Find POC" and poc_names:
            prompt = f"Look up POC {poc_names[0]}."
        elif label == "Search Revenue" and group_names:
            prompt = f"What is {group_names[0]}'s revenue?"
        search_group.append((label, prompt))
    groups["Search"] = search_group

    return groups


# ---- Shared rendering: every chip funnels into the same pipeline ---------

def render_welcome_panel(context: BusinessContext, key_prefix: str) -> Optional[str]:
    """Render the welcome header, subtext, and the grouped quick-action
    suggestions. Returns the clicked chip's prompt text, or ``None`` if
    nothing was clicked this run.
    """
    st.markdown(f'<div class="card-title">{WELCOME_HEADER}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="card-subtitle">{WELCOME_SUBTEXT}</div>', unsafe_allow_html=True)

    with st.container(key=f"{key_prefix}_quick_actions"):
        clicked = _render_suggestion_groups(contextual_quick_action_groups(context), key_prefix)
    return clicked


def render_follow_up_suggestions(key_prefix: str) -> Optional[str]:
    """Render one row of follow-up chips beneath the most recent
    assistant turn, based on the tool(s) that answered it (tracked via
    :data:`LAST_SOURCES_KEY`). Returns the clicked chip's prompt text,
    or ``None`` if nothing was clicked this run.
    """
    sources_used = st.session_state.get(LAST_SOURCES_KEY)
    chips = follow_up_suggestions(sources_used)
    if not chips:
        return None

    st.markdown('<div class="section-label">Related Questions</div>', unsafe_allow_html=True)
    with st.container(key=f"{key_prefix}_followups"):
        cols = st.columns(len(chips))
        for column, (label, prompt) in zip(cols, chips):
            with column:
                if st.button(label, key=f"{key_prefix}_followup_{label}", use_container_width=True):
                    return prompt
    return None


def _render_suggestion_groups(groups: Dict[str, List[Suggestion]], key_prefix: str) -> Optional[str]:
    """Render each quick-action group under a plain section label.
    Returns the clicked chip's prompt text, or ``None``."""
    clicked: Optional[str] = None
    for group_name, chips in groups.items():
        st.markdown(f'<div class="section-label">{group_name}</div>', unsafe_allow_html=True)
        cols = st.columns(len(chips))
        for column, (label, prompt) in zip(cols, chips):
            with column:
                if st.button(label, key=f"{key_prefix}_qa_{group_name}_{label}", use_container_width=True):
                    clicked = prompt
    return clicked
