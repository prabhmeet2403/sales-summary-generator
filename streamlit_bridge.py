"""
streamlit_bridge.py
====================
The ONLY new code that connects the Streamlit front end to the Sales
Forecast Automation Engine's calculation backend. Every number that
ends up in the generated Summary workbook is produced exclusively by
the existing, unmodified backend modules:

    config, excel_reader, aggregator, comment_mapper,
    historical_lookup, summary_writer, validator

...and the existing, unmodified GUI orchestration wrapper:

    gui.runner  (generate_summary, discover_years, GenerationResult, GenerationError)

which was already built as a UI-framework-agnostic bridge for the
desktop (Tkinter) app and is reused here as-is -- not duplicated.

This file adds exactly ONE piece of new behaviour that didn't already
exist anywhere: a lightweight, pre-generation PREVIEW (detected year,
detected sheets, number of groups) for the "once uploaded" panel the
Streamlit UI shows before the user clicks Generate. Even that preview
does not reimplement any calculation -- it calls the same
`aggregate_section` function the real run uses, so "Number of Groups"
in the preview is guaranteed to match the real run's
"Groups Processed" count.

Nothing in this file should ever be described as "business logic": it
only opens files, calls existing functions, and repackages their
results into a shape convenient for the UI layer.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config  # noqa: E402
from excel_reader import (  # noqa: E402
    MasterWorkbook,
    SheetNotFoundError,
    ColumnNotFoundError,
    build_column_map,
    read_project_rows,
)
from aggregator import aggregate_section  # noqa: E402
from validator import SectionStats  # noqa: E402

# Re-exported as-is for the Streamlit app to use directly -- these are
# the exact same functions/classes the Tkinter desktop GUI calls.
from gui.runner import (  # noqa: E402,F401
    generate_summary,
    GenerationResult,
    GenerationError,
)


@dataclass
class PreviewResult:
    """Everything the "once uploaded" panel needs to show, computed
    without running the full generation pipeline."""

    target_year: int
    available_years: List[int]
    main_sheet: str
    comments_sheet: Optional[str]
    prior_year_sheets: Dict[int, str] = field(default_factory=dict)
    total_sheets_in_workbook: int = 0
    num_groups: int = 0
    num_rows: int = 0
    section_group_counts: Dict[str, int] = field(default_factory=dict)


def preview_workbook(path: str, year: Optional[int] = None) -> PreviewResult:
    """Open the Master workbook far enough to report the same headline
    facts the real run will use -- target year, the sheets that will be
    read, and how many Summary groups will result -- without writing
    anything or touching comments/historical data (neither is needed to
    answer "how many groups").

    Raises GenerationError (imported from gui.runner) with a
    user-friendly message on any problem, exactly like the desktop
    GUI's own preview step.
    """
    try:
        master = MasterWorkbook(path)
    except Exception as exc:  # noqa: BLE001
        raise GenerationError(
            "Could Not Open Workbook",
            f"'{Path(path).name}' could not be opened as an Excel workbook.\n\n{exc}",
        ) from exc

    available_years = master.available_years()
    if not available_years:
        raise GenerationError(
            "No Forecast Sheets Found",
            "No sheet matching \"Sales by Customer- <year>\" was found in this "
            f"workbook.\n\nAvailable sheets:\n{', '.join(master.wb.sheetnames)}",
        )
    target_year = year or max(available_years)

    try:
        main_sheet_name = master.main_sheet_name(target_year)
    except SheetNotFoundError as exc:
        raise GenerationError("Sheet Not Found", str(exc)) from exc
    ws_main = master.sheet(main_sheet_name)

    try:
        cmap = build_column_map(ws_main)
    except ColumnNotFoundError as exc:
        raise GenerationError("Required Column Missing", str(exc)) from exc
    if cmap.group is None:
        raise GenerationError(
            "Required Column Missing",
            f"Required column 'Group' was not found on sheet '{main_sheet_name}'.",
        )
    if not cmap.months:
        raise GenerationError(
            "Required Columns Missing",
            f"No monthly Actual/Forecast columns were found on sheet '{main_sheet_name}'.",
        )

    rows = read_project_rows(ws_main, cmap)

    if cmap.sub_group is None:
        all_codes = {r.ds_code for r in rows}
        sections_config = [
            config.OutputSection(
                key="all",
                heading=None,
                title="Summary",
                subtotal_label="Subtotal",
                ds_codes=list(all_codes),
                show_poc=True,
            )
        ]
    else:
        sections_config = config.OUTPUT_SECTIONS

    section_group_counts: Dict[str, int] = {}
    total_groups = 0
    for section in sections_config:
        stats = SectionStats(section.title)
        groups = aggregate_section(rows, section, stats)
        section_group_counts[section.title] = len(groups)
        total_groups += len(groups)

    comments_sheet_name = master.comments_sheet_name(target_year)

    prior_year_sheets: Dict[int, str] = {}
    for pyear in range(target_year - config.NUM_PRIOR_YEARS_SHOWN, target_year):
        try:
            prior_year_sheets[pyear] = master.main_sheet_name(pyear)
        except SheetNotFoundError:
            pass

    return PreviewResult(
        target_year=target_year,
        available_years=available_years,
        main_sheet=main_sheet_name,
        comments_sheet=comments_sheet_name,
        prior_year_sheets=prior_year_sheets,
        total_sheets_in_workbook=len(master.wb.sheetnames),
        num_groups=total_groups,
        num_rows=len(rows),
        section_group_counts=section_group_counts,
    )
