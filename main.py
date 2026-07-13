"""
main.py
=======
Command-line entry point for the Sales Forecast Automation Engine.

Usage
-----
    python main.py
        Reads the first .xlsx file found in input/, writes the Summary
        workbook to output/, and prints a validation report.

    python main.py --input "/path/to/Master.xlsx" --output-dir "/path/to/out" --year 2026
        Explicit overrides for the input file, output directory, and
        target year (defaults to the most recent year found in the
        workbook).
"""
from __future__ import annotations

import argparse
import logging
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional

import config
from excel_reader import (
    MasterWorkbook,
    SheetNotFoundError,
    ColumnNotFoundError,
    build_column_map,
    read_project_rows,
)
from comment_mapper import CommentMapper
from historical_lookup import HistoricalLookup
from aggregator import aggregate_section, sort_groups, attach_comments, attach_historical
from monthly_view import build_monthly_sections, resolve_month_roles
from sheet_copy import copy_source_sheet_as_new_worksheet
from summary_writer import SummaryWriter
from validator import ValidationReport


def setup_logging(log_path: Path) -> logging.Logger:
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(name)s | %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(fmt)

    root.handlers = [file_handler, console_handler]
    return logging.getLogger("sfae.main")


def find_default_input() -> Optional[Path]:
    candidates = sorted(config.INPUT_DIR.glob("*.xlsx"))
    candidates = [c for c in candidates if not c.name.startswith("~$")]
    return candidates[0] if candidates else None


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate the Sales & Forecast Summary workbook from the Master workbook."
    )
    parser.add_argument(
        "--input", "-i", type=str, default=None,
        help="Path to the Master workbook. Defaults to the first .xlsx file found in input/.",
    )
    parser.add_argument(
        "--output-dir", "-o", type=str, default=None,
        help="Directory to write the Summary workbook to. Defaults to output/.",
    )
    parser.add_argument(
        "--year", "-y", type=int, default=None,
        help="Target year to generate the Summary for. Defaults to the most recent year found in the workbook.",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = config.LOG_DIR / f"run_{timestamp}.log"
    logger = setup_logging(log_path)

    report = ValidationReport()

    input_path = Path(args.input) if args.input else find_default_input()
    if input_path is None or not input_path.exists():
        msg = (
            f"No input workbook found. Place the Master workbook in "
            f"'{config.INPUT_DIR}' or pass --input <path-to-workbook>."
        )
        logger.error(msg)
        print(msg)
        return 1

    output_dir = Path(args.output_dir) if args.output_dir else config.OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    report.source_file = str(input_path)

    try:
        master = MasterWorkbook(str(input_path))
        report.workbook_loaded = True

        available_years = master.available_years()
        if not available_years:
            raise SheetNotFoundError(
                "No sheet matching 'Sales by Customer- <year>' was found in the "
                f"workbook. Available sheets: {', '.join(master.wb.sheetnames)}"
            )
        target_year = args.year or max(available_years)
        report.target_year = target_year
        logger.info("Target year resolved to %s (years available: %s)", target_year, available_years)

        main_sheet_name = master.main_sheet_name(target_year)
        report.main_sheet = main_sheet_name
        ws_main = master.sheet(main_sheet_name)
        cmap = build_column_map(ws_main)

        if cmap.group is None:
            raise ColumnNotFoundError(
                f"Required column 'Group' was not found on sheet '{main_sheet_name}'. "
                "Projects cannot be grouped into Summary rows without it."
            )
        if not cmap.months:
            raise ColumnNotFoundError(
                f"No monthly Actual/Forecast columns were found on sheet '{main_sheet_name}'. "
                "Expected 12 monthly columns with an 'Actual' or 'Forecast' row label above them."
            )

        rows = read_project_rows(ws_main, cmap)
        logger.info("Read %d project rows from sheet '%s'.", len(rows), main_sheet_name)

        if cmap.sub_group is None:
            logger.warning(
                "No 'Sub-Group' column found on '%s'; falling back to a single, "
                "un-sectioned Summary. Section-based scoping (which block of the "
                "Summary a group belongs to) and historical-year matching will be "
                "based on Group name only.",
                main_sheet_name,
            )
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

        comments_sheet_name = master.comments_sheet_name(target_year)
        report.comments_sheet = comments_sheet_name or ""
        comment_mapper = CommentMapper(master.sheet(comments_sheet_name) if comments_sheet_name else None)
        if not comments_sheet_name:
            report.warnings.append(
                f"No '{target_year}_ClientComments'-style sheet found; all comments were left blank."
            )

        prior_years = [target_year - i for i in range(config.NUM_PRIOR_YEARS_SHOWN, 0, -1)]
        years_with_margin = (
            prior_years[-config.YEARS_WITH_MARGIN_SHOWN:] if config.YEARS_WITH_MARGIN_SHOWN else []
        )
        historical = HistoricalLookup(master)
        for year in prior_years:
            try:
                report.prior_year_sheets[year] = master.main_sheet_name(year)
            except SheetNotFoundError:
                report.warnings.append(
                    f"No source sheet found for prior year {year}; its historical figures default to 0."
                )

        section_results = []
        for section in sections_config:
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
            stats = report.new_section(section.title)
            groups = aggregate_section(rows, section, stats)
            groups = sort_groups(groups, sort_alphabetically=section.sort_alphabetically)
            attach_comments(groups, comment_mapper, stats)
            attach_historical(groups, historical, prior_years, years_with_margin, section.ds_codes, stats)
            worksheet2_extra_section_results.append((section, groups))

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

        copy_source_sheet_as_new_worksheet(wb, str(input_path), main_sheet_name, cmap.comments, writer._formula_cache)

        output_filename = f"Sales_and_Forecast_Summary_{target_year}.xlsx"
        output_path = output_dir / output_filename
        wb.save(output_path)
        writer.patch_cached_formula_values(output_path)
        report.output_file = str(output_path)
        report.success = True
        logger.info("Summary workbook written to %s", output_path)

    except (SheetNotFoundError, ColumnNotFoundError) as exc:
        report.errors.append(str(exc))
        report.success = False
        logger.error(str(exc))
    except Exception as exc:  # pragma: no cover - top-level safety net
        report.errors.append(f"Unexpected error: {exc}")
        report.success = False
        logger.exception("Unexpected error while generating the Summary workbook.")

    report_text = report.render()
    print()
    print(report_text)
    report_path = config.LOG_DIR / f"validation_report_{timestamp}.txt"
    report.save(report_path)
    logger.info("Validation report written to %s", report_path)

    return 0 if report.success else 1


if __name__ == "__main__":
    raise SystemExit(main())
