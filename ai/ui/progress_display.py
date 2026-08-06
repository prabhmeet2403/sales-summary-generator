"""
ai.ui.progress_display
=========================
Renders the live, per-node progress checklist while
``ai.workflow.graph.WorkflowGraph.execute`` runs, using the exact same
"a placeholder updated in place as work progresses" pattern ``app.py``'s
existing generation flow already established (``_StepDriver``) -- this
module is the AI Assistant's analogue of that proven pattern, not a
new approach to incremental Streamlit UI updates.

Only nodes that actually run are ever shown (a node whose
``should_run()`` returned ``False`` never appears), so the checklist
always reflects exactly what happened for this specific request.
"""

from __future__ import annotations

from typing import Dict, List, Tuple

import streamlit as st

_RUNNING_ICON = "⏳"
_DONE_ICON = "✓"


class ProgressChecklist:
    """Tracks and renders per-node progress into a Streamlit placeholder.

    Usage:
        placeholder = st.empty()
        checklist = ProgressChecklist(placeholder)
        service.ask(session_id, message, progress_cb=checklist.update)
    """

    def __init__(self, placeholder: "st.delta_generator.DeltaGenerator") -> None:
        self._placeholder = placeholder
        self._steps: Dict[str, Tuple[str, str]] = {}
        self._order: List[str] = []

    def update(self, node_name: str, display_label: str, status: str) -> None:
        """Record one node's status change and re-render the checklist.

        Args:
            node_name: The node's stable identifier.
            display_label: The human-readable label to show.
            status: ``"running"`` or ``"done"``.
        """
        if node_name not in self._steps:
            self._order.append(node_name)
        self._steps[node_name] = (display_label, status)
        self._render()

    def clear(self) -> None:
        """Remove the checklist from view once the turn is complete."""
        self._placeholder.empty()

    def _render(self) -> None:
        lines = []
        for node_name in self._order:
            display_label, status = self._steps[node_name]
            icon = _DONE_ICON if status == "done" else _RUNNING_ICON
            lines.append(f"{icon} {display_label}")
        self._placeholder.markdown("  \n".join(lines))
