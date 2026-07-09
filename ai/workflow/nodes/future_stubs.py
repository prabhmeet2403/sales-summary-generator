"""
ai.workflow.nodes.future_stubs
================================
Visualization, Reporting, and Export are part of the approved nine-node
workflow graph (Architecture Plan v3 Section 1), but no chart, report,
or export tool exists yet -- those arrive in Phase 2c/2d. Per the
approved phased design ("Reporting/Export/Visualization can initially
be stubs that should_run() -> False always"), all three nodes are
implemented here as always-declining stubs.

``should_run()`` always returning ``False`` means ``run()`` is never
actually invoked by ``WorkflowGraph.execute`` in Phase 2b. Each ``run()``
raises rather than silently doing nothing or returning unchanged state,
specifically so a future bug in ``should_run()`` (e.g. an edit that
accidentally makes it return ``True``) fails loudly during development
instead of quietly producing an empty result -- this is a deliberate,
narrow use of ``NotImplementedError`` as a defensive assertion, not an
unfinished implementation of something these nodes are supposed to do
in Phase 2b.

Each will be replaced by a real implementation in the phase that
introduces its first tool (Visualization in Phase 2c alongside the
Chart Generation tool; Reporting and Export in Phase 2c/2d alongside
the report/export tools) -- an isolated change to this file, not to
``ai.workflow.graph`` or to any other node.
"""

from __future__ import annotations

from typing import ClassVar

from ai.workflow.graph import WorkflowNode, WorkflowState


class VisualizationNode(WorkflowNode):
    """Stub for Phase 2c: will consult the Chart Recommendation Engine
    and produce chart specs once chart-producing tools exist."""

    name: ClassVar[str] = "visualization"
    display_label: ClassVar[str] = "Creating visualization"

    def should_run(self, state: WorkflowState) -> bool:
        return False  # no chart tool is registered until Phase 2c

    def run(self, state: WorkflowState) -> WorkflowState:
        raise NotImplementedError(
            "VisualizationNode.run() should never be called in Phase 2b -- "
            "should_run() always returns False. If this was reached, "
            "should_run()'s condition was changed without implementing this node."
        )


class ReportingNode(WorkflowNode):
    """Stub for Phase 2c/2d: will invoke the Dashboard Designer or a
    Report Template once those exist."""

    name: ClassVar[str] = "reporting"
    display_label: ClassVar[str] = "Preparing report"

    def should_run(self, state: WorkflowState) -> bool:
        return False  # no dashboard/report tool is registered until Phase 2c/2d

    def run(self, state: WorkflowState) -> WorkflowState:
        raise NotImplementedError(
            "ReportingNode.run() should never be called in Phase 2b -- "
            "should_run() always returns False. If this was reached, "
            "should_run()'s condition was changed without implementing this node."
        )


class ExportNode(WorkflowNode):
    """Stub for Phase 2d: will produce downloadable artifacts once the
    export pipeline exists."""

    name: ClassVar[str] = "export"
    display_label: ClassVar[str] = "Finalizing export"

    def should_run(self, state: WorkflowState) -> bool:
        return False  # no export tool is registered until Phase 2d

    def run(self, state: WorkflowState) -> WorkflowState:
        raise NotImplementedError(
            "ExportNode.run() should never be called in Phase 2b -- "
            "should_run() always returns False. If this was reached, "
            "should_run()'s condition was changed without implementing this node."
        )
