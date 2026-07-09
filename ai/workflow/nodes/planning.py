"""
ai.workflow.nodes.planning
============================
Confirms the classified intent can actually be served by the current
tool registry before the rest of the pipeline spends further effort on
it.

Phase 2b's tool registry has exactly one populated category
(``ToolCategory.ANALYSIS``) -- deciding *between* several non-trivial
categories, which is this node's fuller responsibility per the approved
architecture, has nothing to select between yet. What this node does
today that is genuinely worth doing: verifying at least one ANALYSIS
tool is actually registered before ``AnalysisNode`` spends an LLM call
attempting to use one. As chart/search/reporting tools are registered
in later phases, this node is the natural place to add real
category-selection logic -- an additive change to this file, not to the
graph or to any tool.
"""

from __future__ import annotations

import logging
from typing import ClassVar

from ai.tools.base import ToolCategory
from ai.tools.registry import ToolRegistry
from ai.workflow.graph import Intent, WorkflowNode, WorkflowState

logger = logging.getLogger(__name__)


class PlanningNode(WorkflowNode):
    """Confirms a suitable tool category is available for the
    classified intent."""

    name: ClassVar[str] = "planning"
    display_label: ClassVar[str] = "Planning analysis"

    def __init__(self, registry: ToolRegistry) -> None:
        self._registry = registry

    def should_run(self, state: WorkflowState) -> bool:
        return state.intent == Intent.ANALYTICAL_QUESTION

    def run(self, state: WorkflowState) -> WorkflowState:
        if not self._registry.tools_in_category(ToolCategory.ANALYSIS):
            logger.warning(
                "Intent was ANALYTICAL_QUESTION but no ANALYSIS-category tools are "
                "registered; downgrading to UNSUPPORTED_REQUEST rather than running "
                "Analysis with nothing to call."
            )
            state.intent = Intent.UNSUPPORTED_REQUEST
        return state
