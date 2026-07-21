"""
config.py
=========
Central configuration for the Sales Forecast Automation Engine.

Every piece of knowledge that is specific to *this* business (as opposed
to generic Excel/aggregation logic) lives in this file. The engine itself
(excel_reader / aggregator / comment_mapper / historical_lookup /
summary_writer) never hardcodes a row number, a customer name, a column
letter, or a sheet index -- it always asks this file "what does column X
look like?" or "which codes belong in section Y?".

If next month a brand-new Sub-Group code shows up, or a column header is
renamed, this is the only file that should need a one-line edit.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
LOG_DIR = PROJECT_ROOT / "logs"

for _d in (INPUT_DIR, OUTPUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------
# Sheet-name detection (regular expressions, matched case-insensitively)
# --------------------------------------------------------------------------
# The master workbook's main project-level sheet is named something like
# "Sales by Customer- 2026" (spacing/dashes drift slightly month to
# month). We locate it by pattern and by the 4-digit year embedded in the
# name -- never by a fixed sheet index or a fixed literal string.
MAIN_SHEET_PATTERN = r"sales\s*by\s*customer.*?(\d{4})"

# The per-client narrative-comments sheet is named "<year>_ClientComments".
COMMENTS_SHEET_PATTERN = r"(\d{4}).*client\s*comments"

# Sheets that look like the comments sheet but are backups/duplicates and
# must be skipped even though they match the pattern above (e.g.
# "2026_ClientComments (2)").
COMMENTS_SHEET_EXCLUDE_KEYWORDS = ["copy", "backup", "(2)", "(3)", "old", "bak", "draft"]

# --------------------------------------------------------------------------
# Quarter definition (Rule 2) -- purely calendar-month based.
# --------------------------------------------------------------------------
QUARTER_MONTHS: Dict[str, tuple] = {
    "Q1": (1, 2, 3),
    "Q2": (4, 5, 6),
    "Q3": (7, 8, 9),
    "Q4": (10, 11, 12),
}
QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]

# --------------------------------------------------------------------------
# Row "role" keywords (found in the header row directly ABOVE the
# Name/POC/Service field-header row) that tell us what a given monthly
# column represents. Matching is done by substring on normalised text, so
# "Actual", "Actuals" and "Forecast" are all treated as the revenue role.
# --------------------------------------------------------------------------
REVENUE_ROLE_KEYWORDS = ["actual", "forecast"]
SALARY_ROLE_KEYWORDS = ["salary"]
MARGIN_ROLE_KEYWORDS = ["margin"]

# --------------------------------------------------------------------------
# Sub-Group ("DS-code") classification
# --------------------------------------------------------------------------
# Every genuine project row in the master workbook is tagged with a
# Sub-Group value of the form "DS<NN>_<Status>", e.g. "DS10_Secured",
# "DS70_Secured", "DS30_Projection". This is the most reliable dynamic
# signal available for (a) telling a real data row apart from a section
# title / subtotal / blank row, and (b) knowing which block of the
# printed Sales Summary a row belongs to. It is far more robust than
# trying to infer section boundaries from blank rows or bold text, and it
# keeps working even if new rows are inserted anywhere in the sheet.
#
# The mapping below was reverse engineered by comparing the raw master
# workbook to the manually produced Summary workbook. It is kept in this
# config file -- NOT in the code -- specifically so that a business user
# can extend it in one place if a brand-new numeric code appears in a
# future month (e.g. a new staffing client tagged "DS86_Secured").
#
# Any Sub-Group code found in the data that is NOT listed in any section
# below is deliberately EXCLUDED from the Summary and reported in the
# validation log under "Unmapped Sub-Groups" -- so a new business line
# never silently disappears without a trace.


@dataclass
class OutputSection:
    key: str                      # internal identifier
    heading: Optional[str]        # optional top-level heading printed above the block
    title: str                    # the block's own title row
    subtotal_label: str           # text used on the block's subtotal row
    ds_codes: List[int]           # numeric DS-codes that belong to this block
    show_poc: bool = False        # whether the POC column is populated for this block
    sort_alphabetically: bool = True  # billable groups sorted A-Z vs. kept in source order
    # Exact blank-row spacing, reproduced verbatim from the manually
    # built Summary workbook (it is NOT uniform across sections there,
    # so this is deliberately explicit per-section rather than a single
    # global "one blank line between things" rule).
    blank_rows_after_heading: int = 0   # heading row -> title row (only used if `heading` is set)
    blank_rows_after_title: int = 0     # title row -> first data row
    blank_rows_after_data: int = 1      # last data row -> subtotal row
    blank_rows_after_subtotal: int = 1  # subtotal row -> next section (or end of sheet)
    # Almost always None -- a section is normally identified purely by
    # its `ds_codes`, matched anywhere on the source sheet. Only set
    # (dynamically, at runtime -- never a hardcoded pair of numbers
    # here) when two DIFFERENT sections legitimately share the same
    # DS-code (e.g. "Investments" reuses the DS10_Secured code that
    # "projects_track1" already uses) and so need a source row-range
    # as a tiebreaker, in addition to the DS-code, to tell their rows
    # apart. See main.py/gui/runner.py's `_disambiguate_shared_ds_code_sections`.
    row_range: Optional[Tuple[int, int]] = None
    # Rule 6 ("skip blank groups": a group with zero revenue AND zero
    # margin this year is dropped, on the assumption that means no
    # real activity happened) is the right default for every existing
    # section, where a $0 group really does mean a dormant/inactive
    # engagement. "Investments" is structurally different -- its rows
    # are deliberately non-billable internal projects that may
    # genuinely have $0 revenue and margin every month, and still need
    # to be shown (the user explicitly asked to "copy every Investment
    # project"). Only that section sets this False; every other
    # section keeps Rule 6's existing behavior unchanged.
    skip_blank_groups: bool = True


OUTPUT_SECTIONS: List[OutputSection] = [
    OutputSection(
        key="projects_track1",
        heading="Solutions and Staff Augmentation (Projects)",
        title="Solutions and Staff Augmentation (Projects) - Track 1",
        subtotal_label="Subtotal : Track 1",
        ds_codes=[10],
        show_poc=False,
        sort_alphabetically=True,
        blank_rows_after_heading=1,   # heading, BLANK, title (rows 3,4,5 in the sample)
        blank_rows_after_title=0,     # title immediately followed by first data row
        blank_rows_after_data=1,      # one blank row before the subtotal
        blank_rows_after_subtotal=1,  # one blank row before the next section
    ),
    OutputSection(
        key="staffing_secured",
        heading=None,
        title="Staffing- Secured",
        subtotal_label="Subtotal : Staffing- Secured",
        ds_codes=[70, 80, 85],
        show_poc=True,
        # Observed to be listed in original source order in the sample
        # workbook rather than re-sorted alphabetically (a small, stable
        # list of staffing accounts) -- configurable here if that ever
        # changes.
        sort_alphabetically=False,
        blank_rows_after_title=1,     # one blank row before the first data row
        blank_rows_after_data=1,      # one blank row before the subtotal (matches every other section)
        blank_rows_after_subtotal=1,  # one blank row before "Investments" (no longer the last section)
    ),
    OutputSection(
        key="investments",
        heading=None,
        title="Investments",
        subtotal_label="Subtotal : Investments",
        # Reuses DS10_Secured -- the SAME code "projects_track1" above
        # already matches -- because that's genuinely what the source
        # sheet uses for these rows (confirmed directly against the
        # Master workbook). Matching by DS-code alone would therefore
        # pull "Investments"' own rows into "projects_track1" too (and
        # vice versa); `row_range` is filled in dynamically at runtime
        # (see main.py/gui/runner.py's
        # `_disambiguate_shared_ds_code_sections`) as the tiebreaker,
        # using the source sheet's own "Investments"/next-section
        # heading rows -- never a hardcoded pair of row numbers here.
        ds_codes=[10],
        show_poc=False,        # matches "projects_track1" above, the closest structural analog
        sort_alphabetically=True,
        blank_rows_after_title=0,     # title immediately followed by first data row (matches source)
        blank_rows_after_data=0,      # subtotal immediately follows the last data row (matches source)
        blank_rows_after_subtotal=0,  # last section on Worksheet 1
        # These are deliberately non-billable internal projects that
        # may genuinely show $0 revenue/margin every month -- still
        # need to appear (Rule 6's "skip a $0 group" default assumption
        # -- no real activity -- doesn't hold for this section).
        skip_blank_groups=False,
    ),
]

# The master workbook's "Sales by Customer- <year>" sheet also carries two
# Projection (forecast, not-yet-secured) blocks -- "Track 1 (Projection)"
# (Sub-Group DS30_Projection) and "Track 2 (Projection)" (Sub-Group
# DS50_Projection) -- sitting after "TOTAL Secured" and before "TOTAL
# Projection" in the source sheet.
#
# These are deliberately kept OUT of `OUTPUT_SECTIONS` above -- per an
# explicit, confirmed business rule, the Projection blocks belong ONLY on
# the "<year> Monthly Performance" worksheet (Worksheet 2), not on the
# "Multi-Year Revenue & Margin" summary (Worksheet 1). `main.py`/`gui/runner.py`
# aggregate this list with the exact same `aggregate_section` mechanism
# `OUTPUT_SECTIONS` uses, then combine the two sets of results only when
# building Worksheet 2's monthly view -- Worksheet 1 is built from
# `OUTPUT_SECTIONS` alone and never sees these two sections. Without this
# list at all, DS30_Projection/DS50_Projection match no section anywhere
# and every one of those rows is silently dropped (surfaced only in the
# validation report's "Unmapped Sub-Groups" list).
#
# The source sheet gives neither block its own individual subtotal row
# (only a combined "TOTAL Projection" spanning both, which the existing
# "TOTAL Secured"/"Solutions and Staff Augmentation Total" grand-total
# rows above are likewise never reproduced for either Track 1 or Track 2
# on Worksheet 1), so `subtotal_label` below extends the same
# "Subtotal : <title>" convention `OUTPUT_SECTIONS` already uses rather
# than copying source text verbatim. `show_poc=False` mirrors
# "projects_track1" above, the closest structural analog (individual
# named engagements under a "Track N" label).
WORKSHEET2_ADDITIONAL_SECTIONS: List[OutputSection] = [
    OutputSection(
        key="projects_track1_projection",
        heading=None,
        title="Track 1 (Projection)",
        subtotal_label="Subtotal : Track 1 (Projection)",
        ds_codes=[30],
        show_poc=False,
        sort_alphabetically=True,
        blank_rows_after_title=0,
        blank_rows_after_data=1,
        blank_rows_after_subtotal=1,
    ),
    OutputSection(
        key="projects_track2_projection",
        heading=None,
        title="Track 2 (Projection)",
        subtotal_label="Subtotal : Track 2 (Projection)",
        ds_codes=[50],
        show_poc=False,
        sort_alphabetically=True,
        blank_rows_after_title=0,
        blank_rows_after_data=1,
        blank_rows_after_subtotal=0,  # last section on Worksheet 2
    ),
]

# Regex used to pull the numeric part out of a Sub-Group value, e.g.
# "DS10_Secured" -> 10.
DS_CODE_PATTERN = r"DS\s*0*?(\d+)"

# --------------------------------------------------------------------------
# Blank-group suppression
# --------------------------------------------------------------------------
# A group is dropped from the Summary entirely when BOTH its aggregated
# current-year revenue AND its aggregated current-year margin are exactly
# zero across every one of its project rows -- i.e. the group had no
# activity at all this year. (Confirmed against the "Doppel" account in
# the sample files: it is the one group the manual summary drops, and it
# is the only DS10 group whose revenue AND margin are both 0.)
ZERO_TOLERANCE = 1e-6

# Tolerance used when cross-checking our own monthly-margin sum against
# the sheet's own "Total Margin" column (a data-quality sanity check, not
# a hard requirement).
CROSS_CHECK_TOLERANCE = 1.0

# --------------------------------------------------------------------------
# Historical year columns
# --------------------------------------------------------------------------
# The Summary shows N prior years of context before the current year's
# quarters. The most recent prior year also shows a Margin column; older
# years show Total only. (2026 Summary => 2024 Total-only, 2025 Total +
# Margin -- exactly mirrored here.)
NUM_PRIOR_YEARS_SHOWN = 2
YEARS_WITH_MARGIN_SHOWN = 1  # counted from the most recent prior year backwards

# How much a same-sheet historical reference figure is allowed to differ
# from a fresh recomputation off that year's own sheet before it is
# flagged in the validation report as "Historical Reference vs
# Recompute" drift. This never changes which number is written to the
# workbook -- it only decides whether to raise a flag for a human to
# spot-check the two source numbers against each other.
HISTORICAL_DRIFT_TOLERANCE = 1.0

# Optional, explicit corrections for a specific (Group, year, metric)
# historical figure. Empty by default -- this tool never guesses or
# silently "fixes" a number on its own. It exists so that if a human
# confirms the *source* workbook contains a data-entry error (e.g. a
# stray extra zero) or a since-corrected figure, the correction can be
# recorded once, here, with a comment explaining why -- rather than
# either shipping a wrong number forever or hand-editing the generated
# file every month.
#
# Example:
#   HISTORICAL_OVERRIDES = {
#       ("Churchs-Chicken", 2025, "total"): 7500.0,  # source AV19 = 75000, confirmed typo (should be 7500) by <name>, <date>
#   }
HISTORICAL_OVERRIDES: Dict[tuple, float] = {}

# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------
FONT_NAME = "Calibri"
FONT_SIZE = 10
# Standard Excel "Currency" format (not "Accounting") -- the $ sits
# directly against the number instead of being left-padded with the
# Accounting style's `_(`/`*` alignment tokens. This is not a guess: it
# is the exact number format read from the Aug 6 NK workbook's own
# "Total Revenue"/"Total Margin" columns (Sales by Customer- 2026,
# columns AN/AO) -- confirmed to render negatives as "-$1,234" (a plain
# minus sign), matching this single-section format's default behaviour.
CURRENCY_FORMAT = '"$"#,##0'
HEADER_FILL = "FFD9E1F2"
SECTION_FILL = "FFBFBFBF"
SUBHEADING_FILL = "FFE7E6E6"
SUBTOTAL_FILL = "FFCCFFFF"
# Fill colors read directly from the Aug 6 NK workbook rather than
# approximated. TOTAL_MARGIN_HEADER_FILL (yellow) is the literal RGB of
# that workbook's "Total Revenue"/"Total Margin" summary-column headers
# (Sales by Customer- 2026, columns AN/AO) -- kept scoped to just the
# final yearly Total/Margin header cells in both of this project's
# worksheets, matching that source column's own distinct header style.
#
# TOTAL_DATA_FILL (green) and MARGIN_DATA_FILL (orange) are the fills
# that workbook uses on EVERY recurring revenue/margin column, not just
# the final summary ones -- its monthly "Actual"/"Forecast" columns use
# the same green as "Total Revenue" (theme 9/accent6), and its monthly
# "Margin" columns use a separate orange (theme 5/accent2), both at the
# same ~0.6 tint. Both are theme colors in the source file, resolved
# here to their exact rendered RGB via the standard OOXML HSL-tint
# formula and cross-checked by pixel-sampling a rendered copy of that
# workbook -- both methods agree exactly for both colors -- so the
# same visual colors apply in this project's own generated workbook
# regardless of which theme it uses.
TOTAL_MARGIN_HEADER_FILL = "FFFFFF00"
TOTAL_DATA_FILL = "FFB4E5A2"
MARGIN_DATA_FILL = "FFF6C6AD"
# FINAL_MARGIN_DATA_FILL (green) applies only to the one Margin column
# that sits immediately before Comments (Worksheet 1's `col_current_
# margin`) or Confidence (Worksheet 2's `col_margin`) -- i.e. the
# final/current-period Margin figure, never the prior-year or
# quarterly/monthly per-period Margin columns, which keep
# MARGIN_DATA_FILL's original orange. Deliberately the exact same
# shade as TOTAL_DATA_FILL (referenced directly, not just copied, so
# the two can never drift apart if TOTAL_DATA_FILL is ever changed).
FINAL_MARGIN_DATA_FILL = TOTAL_DATA_FILL
BORDER_COLOR = "FF000000"  # thin black border applied to every populated cell
# Column widths are no longer configured as fixed constants here --
# Worksheets 1/2/3 are all sized to their actual content, after the
# whole workbook is built, by column_autofit.py (see
# SummaryWriter.autofit_worksheets in main.py/gui/runner.py).
