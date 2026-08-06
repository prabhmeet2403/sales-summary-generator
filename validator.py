"""
validator.py
============
Collects statistics while the workbook is being processed and renders a
human-readable validation report, e.g.:

    Workbook Loaded
    Groups Processed : 84
    Comments Matched : 82
    Missing Comments : 2
    Skipped Blank Groups : 1
    Generation Successful
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List


@dataclass
class SectionStats:
    section_title: str
    groups_processed: int = 0
    comments_matched: int = 0
    missing_comments: int = 0
    skipped_blank_groups: List[str] = field(default_factory=list)
    historical_missing: List[str] = field(default_factory=list)
    historical_drift: List[str] = field(default_factory=list)
    margin_cross_check_mismatches: List[str] = field(default_factory=list)


@dataclass
class ValidationReport:
    source_file: str = ""
    output_file: str = ""
    target_year: int = 0
    workbook_loaded: bool = False
    main_sheet: str = ""
    comments_sheet: str = ""
    prior_year_sheets: Dict[int, str] = field(default_factory=dict)
    sections: List[SectionStats] = field(default_factory=list)
    unmapped_sub_groups: Dict[str, int] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    success: bool = False
    generated_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    # ---- convenience aggregate properties --------------------------------
    @property
    def total_groups_processed(self) -> int:
        return sum(s.groups_processed for s in self.sections)

    @property
    def total_comments_matched(self) -> int:
        return sum(s.comments_matched for s in self.sections)

    @property
    def total_missing_comments(self) -> int:
        return sum(s.missing_comments for s in self.sections)

    @property
    def total_skipped_blank_groups(self) -> int:
        return sum(len(s.skipped_blank_groups) for s in self.sections)

    def new_section(self, title: str) -> SectionStats:
        stats = SectionStats(section_title=title)
        self.sections.append(stats)
        return stats

    # ---- rendering ---------------------------------------------------
    def render(self) -> str:
        lines: List[str] = []
        lines.append("=" * 70)
        lines.append("SALES FORECAST AUTOMATION ENGINE - VALIDATION REPORT")
        lines.append("=" * 70)
        lines.append(f"Generated At        : {self.generated_at}")
        lines.append(f"Source Workbook     : {self.source_file}")
        lines.append(f"Target Year         : {self.target_year}")
        lines.append(f"Main Sheet          : {self.main_sheet}")
        lines.append(f"Comments Sheet      : {self.comments_sheet or '(not found - comments left blank)'}")
        for yr, sh in sorted(self.prior_year_sheets.items()):
            lines.append(f"Prior Year Sheet    : {yr} -> {sh}")
        lines.append("-" * 70)
        lines.append("Workbook Loaded" if self.workbook_loaded else "Workbook Load FAILED")
        lines.append("")

        for s in self.sections:
            lines.append(f"[Section] {s.section_title}")
            lines.append(f"  Groups Processed        : {s.groups_processed}")
            lines.append(f"  Comments Matched        : {s.comments_matched}")
            lines.append(f"  Missing Comments        : {s.missing_comments}")
            lines.append(f"  Skipped Blank Groups    : {len(s.skipped_blank_groups)}")
            if s.skipped_blank_groups:
                lines.append(f"    -> {', '.join(s.skipped_blank_groups)}")
            if s.historical_missing:
                lines.append(f"  Historical Data Missing : {len(s.historical_missing)}")
                lines.append(f"    -> {', '.join(s.historical_missing)}")
            if s.historical_drift:
                lines.append(f"  Historical Reference vs Recompute Drift: {len(s.historical_drift)}")
                for d in s.historical_drift:
                    lines.append(f"    -> {d}")
            if s.margin_cross_check_mismatches:
                lines.append(f"  Margin Cross-Check Flags: {len(s.margin_cross_check_mismatches)}")
                lines.append(f"    -> {', '.join(s.margin_cross_check_mismatches)}")
            lines.append("")

        lines.append("-" * 70)
        lines.append("TOTALS")
        lines.append(f"  Groups Processed        : {self.total_groups_processed}")
        lines.append(f"  Comments Matched        : {self.total_comments_matched}")
        lines.append(f"  Missing Comments        : {self.total_missing_comments}")
        lines.append(f"  Skipped Blank Groups    : {self.total_skipped_blank_groups}")

        if self.unmapped_sub_groups:
            lines.append("")
            lines.append("Unmapped Sub-Groups (excluded from Summary; see config.OUTPUT_SECTIONS):")
            for code, count in sorted(self.unmapped_sub_groups.items()):
                lines.append(f"    {code}: {count} row(s)")

        if self.warnings:
            lines.append("")
            lines.append("Warnings:")
            for w in self.warnings:
                lines.append(f"  - {w}")

        if self.errors:
            lines.append("")
            lines.append("Errors:")
            for e in self.errors:
                lines.append(f"  - {e}")

        lines.append("-" * 70)
        lines.append(f"Output Workbook     : {self.output_file}")
        lines.append("Generation Successful" if self.success else "Generation FAILED")
        lines.append("=" * 70)
        return "\n".join(lines)

    def save(self, path: Path) -> None:
        path.write_text(self.render(), encoding="utf-8")
