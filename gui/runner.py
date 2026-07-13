"""
gui/runner.py
=============
Adapts the exact same orchestration steps as main.py's CLI `main()` into
a form suitable for a GUI: progress callbacks instead of console/log
output, a structured result object instead of a process exit code, and
every exception caught and turned into a friendly message instead of a
traceback on stderr.

IMPORTANT: this module contains no calculation or business logic of its
own. Every number in the generated workbook is produced by the exact
same, unmodified backend modules the CLI uses:
    config, excel_reader, aggregator, comment_mapper,
    historical_lookup, summary_writer, validator
This file only sequences calls into them and reports progress -- it is
the GUI's counterpart to main.py, not a replacement for it. main.py
itself is untouched and remains fully usable from the command line.
"""
from __future__ import annotations

import sys
import traceback
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config  # noqa: E402
from excel_reader import (  # noqa: E402
    MasterWorkbook,
    ProjectRow,
    SheetNotFoundError,
    ColumnNotFoundError,
    build_column_map,
    read_project_rows,
)
from comment_mapper import CommentMapper  # noqa: E402
from historical_lookup import HistoricalLookup  # noqa: E402
from aggregator import (  # noqa: E402
    GroupSummary,
    aggregate_section,
    sort_groups,
    attach_comments,
    attach_historical,
)
from monthly_view import MonthlyGroupSummary, build_monthly_sections, resolve_month_roles  # noqa: E402
from sheet_copy import copy_source_sheet_as_new_worksheet  # noqa: E402
from summary_writer import SummaryWriter  # noqa: E402
from validator import ValidationReport  # noqa: E402

ProgressCallback = Callable[[str], None]


def _noop(_message: str) -> None:
    return None


@dataclass
class GenerationResult:
    """Everything the GUI needs to render a result screen, success or
    failure, without re-parsing text output.

    Phase 2 fields (``section_results`` through ``prior_years`` below)
    are additive: they carry the exact same already-computed objects
    ``generate_summary`` builds for the CLI/GUI/Streamlit success path,
    so the Phase 2 AI layer (``ai.context.BusinessContext``) can consume
    them without re-parsing the source workbook or recomputing any
    business value. They are ``None`` on any unsuccessful generation.
    """

    success: bool
    report: Optional[ValidationReport] = None
    output_path: Optional[Path] = None
    report_path: Optional[Path] = None
    error_title: str = ""
    error_message: str = ""
    error_details: str = ""
    available_years: List[int] = field(default_factory=list)
    # --- Phase 2 additions -------------------------------------------
    # See ai/context.py's BusinessContext.from_generation_result, the
    # sole consumer of these fields.
    section_results: Optional[List[Tuple[config.OutputSection, List[GroupSummary]]]] = None
    monthly_section_results: Optional[List[Tuple[config.OutputSection, List[MonthlyGroupSummary]]]] = None
    rows: Optional[List[ProjectRow]] = None
    month_roles: Optional[Dict[int, str]] = None
    target_year: Optional[int] = None
    prior_years: Optional[List[int]] = None


class GenerationError(Exception):
    """Raised for known, user-actionable failure conditions (bad file,
    missing sheet/column, etc.) so the GUI can show a concise message
    instead of a raw traceback. Unexpected exceptions are still caught
    separately and shown with full details for troubleshooting."""

    def __init__(self, title: str, message: str):
        super().__init__(message)
        self.title = title
        self.message = message


def discover_years(input_path: str) -> List[int]:
    """Open the workbook just far enough to list available forecast
    years, for populating the GUI's year dropdown. Raises
    GenerationError with a friendly message on failure."""
    try:
        master = MasterWorkbook(input_path)
    except Exception as exc:  # noqa: BLE001 - surfaced to the user as-is
        raise GenerationError(
            "Could Not Open Workbook",
            f"'{Path(input_path).name}' could not be opened as an Excel workbook.\n\n{exc}",
        ) from exc

    years = master.available_years()
    if not years:
        raise GenerationError(
            "No Forecast Sheets Found",
            "No sheet matching \"Sales by Customer- <year>\" was found in this "
            f"workbook.\n\nAvailable sheets:\n{', '.join(master.wb.sheetnames)}",
        )
    return years


def generate_summary(
    input_path: str,
    output_dir: str,
    year: Optional[int] = None,
    progress_cb: Optional[ProgressCallback] = None,
) -> GenerationResult:
    """Run the full generation pipeline, mirroring main.py's `main()`
    step-for-step, and return a GenerationResult instead of printing to
    the console / exiting the process.
    """
    progress = progress_cb or _noop
    report = ValidationReport()
    input_path_obj = Path(input_path)
    output_dir_obj = Path(output_dir)

    try:
        if not input_path_obj.exists():
            raise GenerationError(
                "Master Workbook Not Found",
                f"The selected file does not exist:\n{input_path_obj}",
            )
        output_dir_obj.mkdir(parents=True, exist_ok=True)
        report.source_file = str(input_path_obj)

        progress("Loading master workbook…")
        try:
            master = MasterWorkbook(str(input_path_obj))
        except Exception as exc:  # noqa: BLE001
            raise GenerationError(
                "Could Not Open Workbook",
                f"'{input_path_obj.name}' could not be opened as an Excel workbook.\n\n{exc}",
            ) from exc
        report.workbook_loaded = True

        progress("Detecting available forecast years…")
        available_years = master.available_years()
        if not available_years:
            raise GenerationError(
                "No Forecast Sheets Found",
                "No sheet matching \"Sales by Customer- <year>\" was found in this "
                f"workbook.\n\nAvailable sheets:\n{', '.join(master.wb.sheetnames)}",
            )
        target_year = year or max(available_years)
        report.target_year = target_year

        progress(f"Target year: {target_year} — locating main sheet…")
        try:
            main_sheet_name = master.main_sheet_name(target_year)
        except SheetNotFoundError as exc:
            raise GenerationError("Sheet Not Found", str(exc)) from exc
        report.main_sheet = main_sheet_name
        ws_main = master.sheet(main_sheet_name)

        progress(f"Reading columns on '{main_sheet_name}'…")
        try:
            cmap = build_column_map(ws_main)
        except ColumnNotFoundError as exc:
            raise GenerationError("Required Column Missing", str(exc)) from exc

        if cmap.group is None:
            raise GenerationError(
                "Required Column Missing",
                f"Required column 'Group' was not found on sheet '{main_sheet_name}'. "
                "Projects cannot be grouped into Summary rows without it.",
            )
        if not cmap.months:
            raise GenerationError(
                "Required Columns Missing",
                f"No monthly Actual/Forecast columns were found on sheet '{main_sheet_name}'. "
                "Expected 12 monthly columns with an 'Actual' or 'Forecast' row label above them.",
            )

        progress("Reading project rows…")
        rows = read_project_rows(ws_main, cmap)

        if cmap.sub_group is None:
            progress("No Sub-Group column found — using a single, un-sectioned Summary…")
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
            # No Sub-Group column at all in this fallback -- every code
            # already lands in the single synthetic "all" section above,
            # so there is nothing left for the Worksheet-2-only
            # Projection sections to add.
            worksheet2_extra_sections_config: list = []
        else:
            sections_config = config.OUTPUT_SECTIONS
            worksheet2_extra_sections_config = config.WORKSHEET2_ADDITIONAL_SECTIONS
            configured_codes = {
                c for s in (sections_config + worksheet2_extra_sections_config) for c in s.ds_codes
            }
            unmapped = Counter(
                r.sub_group_raw for r in rows
                if r.ds_code is not None and r.ds_code not in configured_codes
            )
            for code, count in unmapped.items():
                report.unmapped_sub_groups[code] = count

        progress("Locating client comments sheet…")
        comments_sheet_name = master.comments_sheet_name(target_year)
        report.comments_sheet = comments_sheet_name or ""
        comment_mapper = CommentMapper(master.sheet(comments_sheet_name) if comments_sheet_name else None)
        if not comments_sheet_name:
            report.warnings.append(
                f"No '{target_year}_ClientComments'-style sheet found; all comments were left blank."
            )

        progress("Locating prior-year sheets…")
        prior_years = [target_year - i for i in range(config.NUM_PRIOR_YEARS_SHOWN, 0, -1)]
        years_with_margin = (
            prior_years[-config.YEARS_WITH_MARGIN_SHOWN:] if config.YEARS_WITH_MARGIN_SHOWN else []
        )
        historical = HistoricalLookup(master)
        for pyear in prior_years:
            try:
                report.prior_year_sheets[pyear] = master.main_sheet_name(pyear)
            except SheetNotFoundError:
                report.warnings.append(
                    f"No source sheet found for prior year {pyear}; its historical figures default to 0."
                )

        section_results = []
        for section in sections_config:
            progress(f"Aggregating section: {section.title}…")
            stats = report.new_section(section.title)
            groups = aggregate_section(rows, section, stats)
            groups = sort_groups(groups, sort_alphabetically=section.sort_alphabetically)
            attach_comments(groups, comment_mapper, stats)
            attach_historical(groups, historical, prior_years, years_with_margin, section.ds_codes, stats)
            section_results.append((section, groups))

        # Same aggregation mechanism as above, for the two Projection
        # sections that belong ONLY on Worksheet 2 -- see
        # config.WORKSHEET2_ADDITIONAL_SECTIONS. Kept out of
        # `section_results` (and therefore out of Worksheet 1) entirely;
        # combined with it only when building Worksheet 2's monthly view
        # below.
        worksheet2_extra_section_results = []
        for section in worksheet2_extra_sections_config:
            progress(f"Aggregating section: {section.title}…")
            stats = report.new_section(section.title)
            groups = aggregate_section(rows, section, stats)
            groups = sort_groups(groups, sort_alphabetically=section.sort_alphabetically)
            attach_comments(groups, comment_mapper, stats)
            attach_historical(groups, historical, prior_years, years_with_margin, section.ds_codes, stats)
            worksheet2_extra_section_results.append((section, groups))

        progress("Building the Summary workbook…")
        writer = SummaryWriter(target_year, prior_years, years_with_margin)
        month_roles = resolve_month_roles(ws_main, cmap)
        # Worksheet 2 gets the combined (Worksheet 1 + Projection) monthly
        # view. `monthly_section_results` -- what gets attached to
        # GenerationResult below, and from there is the only thing the
        # AI layer (ai/context.py) ever reads -- stays sliced to exactly
        # Worksheet 1's own sections, so the AI's view of the data is
        # completely unaffected by Worksheet 2's Projection sections
        # (same rows, same order, same objects; slicing rather than a
        # second `build_monthly_sections` call avoids computing anything
        # twice).
        worksheet2_monthly_section_results = build_monthly_sections(
            rows, cmap, ws_main, section_results + worksheet2_extra_section_results,
        )
        monthly_section_results = worksheet2_monthly_section_results[:len(section_results)]
        wb = writer.build(section_results, worksheet2_monthly_section_results, month_roles)

        progress(f"Copying '{main_sheet_name}' as a third worksheet…")
        copy_source_sheet_as_new_worksheet(wb, str(input_path_obj), main_sheet_name, cmap.comments)

        output_filename = f"Sales_and_Forecast_Summary_{target_year}.xlsx"
        output_path = output_dir_obj / output_filename
        progress(f"Saving workbook to {output_path.name}…")
        try:
            wb.save(output_path)
            writer.patch_cached_formula_values(output_path)
        except PermissionError as exc:
            raise GenerationError(
                "Cannot Save Workbook",
                f"'{output_path.name}' could not be saved -- it may be open in Excel.\n"
                "Close the file and try again.",
            ) from exc
        report.output_file = str(output_path)
        report.success = True

        progress("Saving validation report…")
        report_path = output_dir_obj / f"{output_path.stem}_validation_report.txt"
        report.save(report_path)

        progress("Done.")
        return GenerationResult(
            success=True,
            report=report,
            output_path=output_path,
            report_path=report_path,
            available_years=available_years,
            section_results=section_results,
            monthly_section_results=monthly_section_results,
            rows=rows,
            month_roles=month_roles,
            target_year=target_year,
            prior_years=prior_years,
        )

    except GenerationError as exc:
        report.errors.append(exc.message)
        report.success = False
        report_path = _save_failure_report(output_dir_obj, report)
        return GenerationResult(
            success=False,
            report=report,
            report_path=report_path,
            error_title=exc.title,
            error_message=exc.message,
        )
    except Exception as exc:  # noqa: BLE001 - top-level safety net
        report.errors.append(f"Unexpected error: {exc}")
        report.success = False
        details = traceback.format_exc()
        report_path = _save_failure_report(output_dir_obj, report)
        return GenerationResult(
            success=False,
            report=report,
            report_path=report_path,
            error_title="Unexpected Error",
            error_message=str(exc),
            error_details=details,
        )


def _save_failure_report(output_dir_obj: Path, report: ValidationReport) -> Optional[Path]:
    """Best-effort save of the (failed) validation report so the user
    can still inspect what happened; never raises."""
    try:
        output_dir_obj.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = output_dir_obj / f"validation_report_FAILED_{timestamp}.txt"
        report.save(report_path)
        return report_path
    except Exception:  # noqa: BLE001
        return None
