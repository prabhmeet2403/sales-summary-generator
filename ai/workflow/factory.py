"""
ai.workflow.factory
======================
Assembles the approved nine-node workflow graph (Architecture Plan v3
Section 1.2) in order. This is the one place the node sequence is
defined -- adding a future node means inserting one line here, not
changing any node's own code.
"""

from __future__ import annotations

from ai.context import BusinessContext
from ai.llm.provider import LLMProvider
from ai.tools.registry import ToolRegistry
from ai.workflow.graph import WorkflowGraph
from ai.workflow.nodes.analysis import AnalysisNode
from ai.workflow.nodes.entity_resolution import EntityResolutionNode
from ai.workflow.nodes.filtering import FilteringNode
from ai.workflow.nodes.future_stubs import ExportNode, ReportingNode, VisualizationNode
from ai.workflow.nodes.intent_detection import IntentDetectionNode
from ai.workflow.nodes.planning import PlanningNode
from ai.workflow.nodes.response import ResponseNode


def build_default_workflow_graph(
    provider: LLMProvider,
    registry: ToolRegistry,
    context: BusinessContext,
) -> WorkflowGraph:
    """Construct the standard nine-node workflow graph.

    Args:
        provider: The LLM provider every LLM-calling node uses.
        registry: The tool registry ``AnalysisNode`` dispatches through.
        context: The business data every data-reading node uses.

    Returns:
        A :class:`~ai.workflow.graph.WorkflowGraph` with all nine nodes
        in the approved order. Visualization, Reporting, and Export are
        included per the approved architecture but never actually
        execute in Phase 2b (see ``ai.workflow.nodes.future_stubs``).
    """
    return WorkflowGraph(
        [
            IntentDetectionNode(provider),
            PlanningNode(registry),
            EntityResolutionNode(context),
            FilteringNode(),
            AnalysisNode(provider, registry, context),
            VisualizationNode(),
            ReportingNode(),
            ExportNode(),
            ResponseNode(provider),
        ]
    )
