"""
ai
==
Phase 2: the AI-powered Business Intelligence platform built on top of
the Phase 1 Sales Forecast Automation Engine.

This package is entirely additive. It reads Phase 1's already-validated
output (via the optional fields attached to ``gui.runner.GenerationResult``
-- see ``ai.context.BusinessContext``) and never re-implements, re-parses,
or recalculates anything Phase 1 already produced.

Architectural boundary (enforced by convention, checked in code review
and by ``tests/test_ai_architecture_boundaries.py``):

- No module outside ``ai.ui`` may import ``streamlit``.
- No module outside ``ai.llm.providers`` may import ``boto3`` or reference
  a specific LLM provider by name.
- No module in this package may import from ``ai.ui``.
- No Phase 1 module (``excel_reader``, ``aggregator``, ``comment_mapper``,
  ``historical_lookup``, ``monthly_view``, ``summary_writer``,
  ``validator``, ``main``) may import anything from ``ai``.

See ``ai/README.md`` for the full package guide and
``Phase2_AI_Assistant_Architecture_Plan_v3.md`` for the approved
architecture this package implements.

Current implementation status: Phase 2a (foundation layer) -- settings,
the LLM provider abstraction, the business data context, the DataFrame
query layer, the universal filter engine, and the service facade. The
workflow graph (planner), tools, charts, dashboards, reports, search,
and the Streamlit UI are implemented in later phases per the approved
phased rollout.
"""

from __future__ import annotations

__all__: list[str] = []
