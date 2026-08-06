"""
ai.tools.schemas
==================
Shared JSON Schema fragments reused across tool definitions.

Centralizing the filter fields here means every tool that accepts
filtering exposes the *same* argument shape to the model -- this is
what "the LLM only has to learn one filter shape across every tool"
(Universal Filter Engine, ``ai.data.filters``) means concretely at the
schema level, not just in the Python implementation.
"""

from __future__ import annotations

from ai.data.filters import Filter

#: The filter properties every analytical tool's schema should include
#: (merged into that tool's own ``schema["properties"]``). Field names
#: match ``ai.data.filters.Filter``'s attribute names exactly, so a
#: tool's ``run()`` can construct a ``Filter`` directly from
#: ``arguments`` with no name translation.
FILTER_PROPERTIES: dict = {
    "client": {
        "type": "string",
        "description": "Filter to a single client/group name (exact match, case-insensitive).",
    },
    "poc": {
        "type": "string",
        "description": "Filter to a single POC (account owner) name (exact match, case-insensitive).",
    },
    "section": {
        "type": "string",
        "enum": ["projects_track1", "staffing_secured"],
        "description": "Filter to a single Summary section.",
    },
    "quarters": {
        "type": "array",
        "items": {"type": "string", "enum": ["Q1", "Q2", "Q3", "Q4"]},
        "description": "Filter to one or more quarters. Only meaningful for month-level analysis (quarter comparisons, trends).",
    },
}


def filter_from_arguments(arguments: dict) -> Filter:
    """Build a :class:`~ai.data.filters.Filter` from a tool call's
    arguments, reading exactly the fields declared in
    :data:`FILTER_PROPERTIES`.

    Args:
        arguments: The tool call's arguments dict. Fields not present
            in :data:`FILTER_PROPERTIES` are ignored (a tool's own
            arguments, like ``top_n``, coexist in the same dict).

    Returns:
        A :class:`Filter` with whichever fields were present set, and
        the rest left at their ``None`` defaults.
    """
    return Filter(
        client=arguments.get("client"),
        poc=arguments.get("poc"),
        section=arguments.get("section"),
        quarters=arguments.get("quarters"),
    )
