"""
ai.tools.base
==============
The plugin interface every AI tool implements.

A tool is a small, self-contained unit that reads already-validated
business data (through ``ai.analytics.engine.AnalyticsEngine`` and the
DataFrame layer) and returns a factual, structured :class:`ToolResult`.
A tool never performs a business calculation Phase 1 didn't already
perform, and never talks to an LLM itself -- deciding *which* tool to
call, and narrating the result, are the workflow graph's job (see
``ai.workflow``), not the tool's.

New tools are added by creating one new module under ``ai.tools``
containing one ``BaseTool`` subclass -- ``ai.tools.registry`` discovers
it automatically. No existing tool file, and no code in
``ai.workflow``, needs to change.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 8.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from typing import ClassVar, Optional

from ai.context import BusinessContext


class ToolCategory(str, Enum):
    """The closed set of capability categories a tool belongs to.

    A closed enum (rather than a free string) lets the workflow graph's
    Planning node, the follow-up suggestion engine (a later phase), and
    the tool registry all reason about "what kinds of capability exist"
    structurally, without string-matching individual tool names.
    """

    ANALYSIS = "analysis"
    VISUALIZATION = "visualization"
    SEARCH = "search"
    EXPORT = "export"
    REPORTING = "reporting"
    DASHBOARD = "dashboard"
    UTILITY = "utility"


@dataclass(frozen=True)
class ToolResult:
    """The structured result a tool returns.

    Attributes:
        summary: A short, factual, numeric description of what the tool
            found (e.g. ``"Q2 revenue for HPE: $115,195. Q3: $104,529
            (-9.2%)."``). This is what gets narrated by the workflow
            graph's Response node -- it is deliberately terse and
            numbers-first, not conversational prose, since the tool's
            job is to be correct, not eloquent.
        raw: Optional machine-readable payload (e.g. a list of row
            dicts for a ranking) for a caller that needs the structured
            data directly, such as a later chart/export step in a
            future phase. ``None`` when a tool has nothing beyond its
            summary to offer.

    Note:
        There is no ``table``/``chart`` field yet. Those arrive as
        additive fields once ``TableSpec``/``ChartSpec`` exist
        (Phase 2c) and a tool actually needs to populate them -- adding
        them now, with nothing to put in them, would be exactly the
        kind of speculative field the project's implementation
        standards ask to avoid.
    """

    summary: str
    raw: Optional[dict] = None


class ToolError(Exception):
    """Raised by a tool's :meth:`BaseTool.run` when it cannot fulfill a
    request (e.g. an unknown client name).

    Caught by ``ai.workflow.nodes.analysis.AnalysisNode`` and converted
    into a result the model can read and explain in plain language --
    never silently swallowed into a fabricated or zero-filled answer.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


class BaseTool(ABC):
    """The interface every AI tool implements.

    Attributes (all class-level, set by each subclass):
        name: A unique, stable identifier (e.g. ``"revenue_analysis"``),
            matching the name Bedrock will use in a tool-use request.
        display_name: A human-readable name for the "Sources Used"
            explainability trail (e.g. ``"Revenue Analysis Tool"``).
        description: A natural-language description shown to the model
            so it can decide when this tool is relevant.
        category: Which :class:`ToolCategory` this tool belongs to.
        schema: A JSON Schema object describing this tool's expected
            arguments, shown to the model alongside ``description``.
    """

    name: ClassVar[str]
    display_name: ClassVar[str]
    description: ClassVar[str]
    category: ClassVar[ToolCategory]
    schema: ClassVar[dict]

    @abstractmethod
    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        """Execute this tool.

        Args:
            arguments: The tool call's arguments, already validated
                against ``schema`` by the caller.
            context: The business data to read from. A tool never reads
                anything outside ``context`` (no file I/O, no other
                global state).

        Returns:
            A :class:`ToolResult`.

        Raises:
            ToolError: If the request cannot be fulfilled (e.g. an
                unrecognized client name) -- never returns a fabricated
                or guessed result instead.
        """
        raise NotImplementedError
