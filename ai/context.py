"""
ai.context
===========
``BusinessContext`` is the single object every AI tool, the workflow
graph, and the service facade use to access Phase 1's already-validated
output. It never opens a workbook and never calls into
``excel_reader``/``aggregator``/``monthly_view`` itself -- it is
constructed from the optional fields Phase 1's own
``gui.runner.generate_summary`` already attaches to its
``GenerationResult`` (see Section 5.1 of the approved architecture),
and does nothing more than index and reshape data those fields already
contain.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 0 and
Revision 2 Section 3 for the approved design.
"""

from __future__ import annotations

import hashlib
import logging
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, List, Optional

import pandas as pd

from ai.data.frames import build_groups_dataframe, build_monthly_dataframe

if TYPE_CHECKING:
    # Imported only for type checking to avoid a hard runtime dependency
    # of this module on gui.runner beyond the one attribute-existence
    # check performed in from_generation_result -- keeps ai.context
    # importable in contexts that construct a BusinessContext some other
    # way in the future (e.g. a future multi-workbook context, see
    # Architecture Plan Section 12) without requiring gui.runner.
    from gui.runner import GenerationResult
    from validator import ValidationReport

logger = logging.getLogger(__name__)


class IncompleteGenerationResultError(Exception):
    """Raised when a :class:`~gui.runner.GenerationResult` lacks the
    Phase 2 fields :class:`BusinessContext` requires.

    This happens for a failed generation (``success=False``), or
    defensively for a ``GenerationResult`` produced by a version of the
    pipeline that predates these fields. Callers (the "AI Assistant"
    Streamlit page, in later phases) are expected to catch this and
    render an empty state rather than let it propagate as an unhandled
    exception -- see Architecture Plan Section 16's error-handling table.
    """

    def __init__(self, missing_fields: List[str]) -> None:
        super().__init__(
            "Cannot build a BusinessContext: GenerationResult is missing "
            f"required field(s): {', '.join(missing_fields)}. This usually means "
            "the generation failed, or produced no output to analyze."
        )
        self.missing_fields = missing_fields


@dataclass(frozen=True)
class BusinessContext:
    """Query-ready view over one successful Summary generation.

    Attributes:
        target_year: The Summary's target year.
        prior_years: The prior years shown alongside the target year
            (e.g. ``[2024, 2025]`` for a 2026 Summary).
        groups_df: One row per group per section (see
            ``ai.data.frames.build_groups_dataframe``).
        monthly_df: One row per group per month per section (see
            ``ai.data.frames.build_monthly_dataframe``).
        fingerprint: A short, stable identifier for this specific
            generation, used to key session/cache invalidation (a new
            generation produces a new fingerprint; the same generation
            always produces the same fingerprint).
        report: The Phase 1 ``ValidationReport`` produced by this
            generation (groups processed, comments matched, warnings,
            errors, the full rendered validation report, etc.) -- the
            exact same object the dashboard's Validation Summary cards
            and "Full Validation Report" expander already read from.
            ``None`` only if this context was built without a report,
            which does not happen for a successful generation.
        elapsed_seconds: Wall-clock generation time, matching the
            dashboard's "Time Taken" card. This is computed by the
            Streamlit page around the ``generate_summary`` call (it is
            not part of Phase 1's own output), so it is passed in
            separately rather than read off ``GenerationResult``.
            ``None`` if unavailable.
    """

    target_year: int
    prior_years: List[int] = field(default_factory=list)
    groups_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    monthly_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    fingerprint: str = ""
    report: Optional["ValidationReport"] = None
    elapsed_seconds: Optional[float] = None

    @classmethod
    def from_generation_result(
        cls,
        generation_result: "GenerationResult",
        elapsed_seconds: Optional[float] = None,
    ) -> "BusinessContext":
        """Build a :class:`BusinessContext` from a successful generation.

        Args:
            generation_result: The result of
                ``gui.runner.generate_summary``. Must be successful and
                carry the Phase 2 fields (``section_results``,
                ``monthly_section_results``, ``month_roles``,
                ``target_year``, ``prior_years``).
            elapsed_seconds: Optional wall-clock generation time (the
                dashboard's "Time Taken" card), passed through as-is
                since it is computed by the caller, not by Phase 1.

        Returns:
            A new, immutable :class:`BusinessContext`.

        Raises:
            IncompleteGenerationResultError: If ``generation_result`` is
                unsuccessful or is missing any of the required Phase 2
                fields.
        """
        required_fields = {
            "section_results": generation_result.section_results,
            "monthly_section_results": generation_result.monthly_section_results,
            "month_roles": generation_result.month_roles,
            "target_year": generation_result.target_year,
        }
        missing = [name for name, value in required_fields.items() if value is None]
        if not generation_result.success or missing:
            if not generation_result.success and not missing:
                missing = ["success"]
            raise IncompleteGenerationResultError(missing)

        groups_df = build_groups_dataframe(
            generation_result.section_results,  # type: ignore[arg-type]
            generation_result.target_year,  # type: ignore[arg-type]
        )
        monthly_df = build_monthly_dataframe(
            generation_result.monthly_section_results,  # type: ignore[arg-type]
            generation_result.month_roles,  # type: ignore[arg-type]
            generation_result.target_year,  # type: ignore[arg-type]
        )
        fingerprint = _compute_fingerprint(generation_result)

        context = cls(
            target_year=generation_result.target_year,  # type: ignore[arg-type]
            prior_years=list(generation_result.prior_years or []),
            groups_df=groups_df,
            monthly_df=monthly_df,
            fingerprint=fingerprint,
            report=generation_result.report,
            elapsed_seconds=elapsed_seconds,
        )
        logger.info(
            "Built BusinessContext for target_year=%d: %d group rows, %d monthly rows, fingerprint=%s",
            context.target_year, len(context.groups_df), len(context.monthly_df), context.fingerprint,
        )
        return context

    def group_names(self) -> List[str]:
        """Return every distinct group (client) name in scope, in the
        order they first appear in ``groups_df``."""
        if self.groups_df.empty:
            return []
        return list(dict.fromkeys(self.groups_df["group"]))

    def poc_names(self) -> List[str]:
        """Return every distinct, non-null POC name in scope."""
        if self.groups_df.empty or "poc" not in self.groups_df.columns:
            return []
        return list(dict.fromkeys(self.groups_df["poc"].dropna()))

    def section_keys(self) -> List[str]:
        """Return every distinct section key in scope (e.g.
        ``["projects_track1", "staffing_secured"]``)."""
        if self.groups_df.empty:
            return []
        return list(dict.fromkeys(self.groups_df["section"]))

    def grounding_summary(self) -> str:
        """Render a compact, factual summary of this context's scope,
        suitable as system-prompt grounding for an LLM call.

        This is deliberately a small, aggregate description (year,
        section names, group/row counts, total revenue/margin) rather
        than a dump of every row -- see Architecture Plan Section 9's
        "grounding context, not raw data" design.
        """
        if self.groups_df.empty:
            return f"No business data is available for target year {self.target_year}."

        lines = [
            f"Target year: {self.target_year}"
            + (f" (prior years shown: {', '.join(str(y) for y in self.prior_years)})" if self.prior_years else ""),
            f"Sections: {', '.join(self.section_keys())}",
            f"Total groups: {len(self.groups_df)}",
            f"Total revenue in scope: ${self.groups_df['revenue'].sum():,.0f}",
            f"Total margin in scope: ${self.groups_df['margin'].sum():,.0f}",
        ]
        return "\n".join(lines)


def _compute_fingerprint(generation_result: "GenerationResult") -> str:
    """Compute a short, stable fingerprint identifying this specific
    generation, from fields the ``ValidationReport`` already carries
    (source file, output file, target year, generation timestamp) --
    no new hashing of raw Excel bytes is performed.
    """
    report = generation_result.report
    basis = "|".join(
        [
            str(report.source_file if report else ""),
            str(report.output_file if report else ""),
            str(generation_result.target_year),
            str(report.generated_at if report else ""),
        ]
    )
    return hashlib.sha256(basis.encode("utf-8")).hexdigest()[:16]
