"""
sheet_copy.py
================
Appends a visual copy of one worksheet from the uploaded Master
workbook onto the generated Summary workbook, as a third sheet -- used
for the "Sales by Customer- <year>" raw-data sheet requested alongside
the two generated (aggregated) sheets `summary_writer.py` already
produces.

This is a distinct operation from everything in `summary_writer.py`:
that module turns already-aggregated `GroupSummary` objects into new,
formatted rows it designs itself. This module does the opposite -- it
copies a worksheet that already exists, cell for cell, style for
style, exactly as the source workbook has it, with one column (the
Comments column, located dynamically by its header text -- never a
hardcoded letter) physically removed and every column after it shifted
left by one, mirroring exactly what Excel's own "Delete Column" does.
Kept separate so neither module's job is duplicated inside the other.

VALUES ONLY, never formulas -- this is a deliberate correctness fix,
not a simplification for its own sake. The source sheet's formulas
routinely reference OTHER sheets in the Master workbook (e.g.
`=VLOOKUP($A13,'Salary Projections 2026'!$A:$D,3,FALSE)`), which this
module intentionally does not also copy (only the one sheet named is
brought over). Carrying a formula like that into the output verbatim
leaves it pointing at a sheet name that does not exist in THIS
workbook; Excel cannot resolve that internally, treats it as needing
an external source, and shows both the "this workbook contains links
to external sources" warning and a #REF!/#N/A in the cell -- neither
of which has anything to do with the column-removal below. Writing
each cell's already-computed value (read from a `data_only=True` load
of the source -- i.e. exactly what Excel itself last displayed there,
never recomputed by this code) produces a sheet that looks identical
but carries no formula for Excel to fail to resolve.
"""

from __future__ import annotations

from copy import copy, deepcopy
from typing import Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.worksheet import Worksheet


def copy_source_sheet_as_new_worksheet(
    wb: Workbook,
    source_path: str,
    sheet_name: str,
    comments_col: Optional[int],
) -> None:
    """Append `sheet_name` from `source_path` onto `wb` as a new sheet
    with the exact same name, copying everything needed to make it
    LOOK identical -- computed cell values (never formulas -- see
    module docstring), fonts/fills/borders/number formats/alignment,
    column widths, row heights, merged cells, freeze panes/split panes/
    scroll position, autofilter, conditional formatting, print/page
    setup, and hidden rows/columns -- with one exception: if
    `comments_col` is given (1-based column index, e.g.
    `ColumnMap.comments` -- resolved dynamically by header text
    elsewhere, never hardcoded), that ENTIRE column is removed (not
    just blanked), and every column after it shifts left by one,
    exactly like using Excel's own "Delete Column" on the source sheet.

    Args:
        wb: The in-progress output workbook (already has its other
            sheet(s) from `SummaryWriter.build()`) to append to.
        source_path: Path to the uploaded Master workbook.
        sheet_name: Exact sheet name to copy (becomes the new sheet's
            name too).
        comments_col: 1-based column index of the Comments column on
            the SOURCE sheet, or None if it has none (nothing is
            removed in that case).
    """
    source_wb_values = load_workbook(source_path, data_only=True)
    source_wb_styles = load_workbook(source_path, data_only=False)
    src_values = source_wb_values[sheet_name]
    src_styles = source_wb_styles[sheet_name]

    new_ws = wb.create_sheet(title=sheet_name)
    _copy_cells(src_values, src_styles, new_ws, comments_col)
    _copy_dimensions(src_styles, new_ws, comments_col)
    _copy_merged_cells(src_styles, new_ws, comments_col)
    _copy_sheet_view(src_styles, new_ws, comments_col)
    _copy_page_setup(src_styles, new_ws, comments_col)
    _copy_conditional_formatting(src_styles, new_ws, comments_col)


def _remap_column(col_idx: int, removed_col: Optional[int]) -> Optional[int]:
    """1-based column index after removing `removed_col` -- unchanged
    if before it, shifted left by one if after it, or None if it IS
    the removed column. Returns `col_idx` unchanged if `removed_col`
    is None (nothing being removed)."""
    if removed_col is None:
        return col_idx
    if col_idx == removed_col:
        return None
    if col_idx < removed_col:
        return col_idx
    return col_idx - 1


def _remap_range_string(range_str: Optional[str], removed_col: Optional[int]) -> Optional[str]:
    """Shift a range string to account for `removed_col` being deleted.
    Handles a single cell (`"D70"`), a single range (`"A2:AO56"`), a
    full-column range (`"AO1:AO1048576"`), and a MULTI-range string --
    conditional formatting's `sqref` in particular can be several
    space-separated ranges/cells covering disjoint areas (e.g.
    `"E16 E18 F51:F67 ..."`), each remapped independently here. Returns
    None if every piece was entirely inside the removed column (nothing
    left of the whole thing to keep).
    """
    if removed_col is None or not range_str:
        return range_str

    if " " in range_str:
        remapped_pieces = [
            piece for piece in (
                _remap_single_range(token, removed_col) for token in range_str.split()
            ) if piece is not None
        ]
        return " ".join(remapped_pieces) if remapped_pieces else None

    return _remap_single_range(range_str, removed_col)


def _remap_single_range(range_str: str, removed_col: int) -> Optional[str]:
    """Remap exactly one cell or range token (no spaces) -- see
    `_remap_range_string`, which this implements the single-piece case
    for."""
    if ":" not in range_str:
        col_str, row = coordinate_from_string(range_str)
        new_col = _remap_column(column_index_from_string(col_str), removed_col)
        if new_col is None:
            return None
        return f"{get_column_letter(new_col)}{row}"

    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    if min_col == max_col == removed_col:
        return None
    new_min = min_col if min_col < removed_col else max(min_col - 1, 1)
    new_max = max_col if max_col < removed_col else max(max_col - 1, 1)
    return f"{get_column_letter(new_min)}{min_row}:{get_column_letter(new_max)}{max_row}"


def _copy_cells(src_values: Worksheet, src_styles: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    for row in src_styles.iter_rows():
        for cell in row:
            new_col = _remap_column(cell.column, comments_col)
            if new_col is None:
                continue  # this cell is in the removed Comments column

            new_cell = dest.cell(row=cell.row, column=new_col)
            # The already-computed value, not the formula -- see module
            # docstring. Reading from the `data_only=True` sibling sheet
            # rather than anything derived from `cell` itself, since
            # `cell` (from the `data_only=False` load) holds formula
            # text for a formula cell, not its result.
            new_cell.value = src_values.cell(row=cell.row, column=cell.column).value

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)


def _copy_dimensions(src: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    for key, dim in src.column_dimensions.items():
        try:
            old_idx = column_index_from_string(key)
        except ValueError:
            continue  # not a plain column letter (e.g. a column-group marker) -- skip
        new_idx = _remap_column(old_idx, comments_col)
        if new_idx is None:
            continue  # the removed Comments column's own dimension
        new_dim = dest.column_dimensions[get_column_letter(new_idx)]
        new_dim.width = dim.width
        new_dim.hidden = dim.hidden
        new_dim.outline_level = dim.outline_level
        new_dim.bestFit = dim.bestFit

    for key, dim in src.row_dimensions.items():
        new_dim = dest.row_dimensions[key]
        new_dim.height = dim.height
        new_dim.hidden = dim.hidden
        new_dim.outline_level = dim.outline_level

    dest.sheet_format.defaultColWidth = src.sheet_format.defaultColWidth
    dest.sheet_format.defaultRowHeight = src.sheet_format.defaultRowHeight


def _copy_merged_cells(src: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    for merged_range in src.merged_cells.ranges:
        remapped = _remap_range_string(str(merged_range), comments_col)
        if remapped is not None and ":" in remapped:  # a single-cell result isn't a merge
            dest.merge_cells(remapped)


def _copy_sheet_view(src: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    """Copy freeze panes, split panes, scroll position, zoom, and
    selection -- everything governing how the sheet scrolls and what's
    visible on open -- by deep-copying the source's own `SheetView`
    object wholesale and only then fixing up the handful of fields
    that are actual cell/column references, rather than going through
    `Worksheet.freeze_panes`'s convenience setter.

    That setter is LOSSY for a pane that's both frozen and
    independently scrolled: its only input is `pane.topLeftCell`, which
    it (incorrectly, for this purpose) assumes IS the freeze split
    point. In the real Master workbook this sheet is copied from, the
    freeze split is at column D / row 3 (`xSplit=3`, `ySplit=2`), but
    the sheet had additionally been scrolled down before saving, so
    `topLeftCell` is `"D70"` -- merely where the scrollable pane's
    view happens to start, not the split boundary. Round-tripping
    `freeze_panes = "D70"` through the setter reinterprets row 70 AS
    the split point, freezing the top 69 rows instead of 2 -- an
    enormous, wrong frozen region, which is exactly what was breaking
    scrolling on the copied sheet.
    """
    dest.views.sheetView[0] = deepcopy(src.sheet_view)

    pane = dest.sheet_view.pane
    if pane is not None:
        if pane.topLeftCell:
            pane.topLeftCell = _remap_range_string(pane.topLeftCell, comments_col)
        if comments_col is not None and pane.xSplit:
            # xSplit is a COUNT of frozen columns from the left, not a
            # column reference -- if the removed column was itself
            # among the frozen columns, there is now one fewer.
            if comments_col <= pane.xSplit:
                pane.xSplit = pane.xSplit - 1

    for selection in dest.sheet_view.selection:
        if selection.activeCell:
            selection.activeCell = _remap_range_string(selection.activeCell, comments_col)
        if selection.sqref:
            selection.sqref = _remap_range_string(selection.sqref, comments_col)


def _copy_page_setup(src: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    if src.auto_filter and src.auto_filter.ref:
        remapped_ref = _remap_range_string(src.auto_filter.ref, comments_col)
        if remapped_ref:
            dest.auto_filter.ref = remapped_ref

    dest.page_setup = copy(src.page_setup)
    dest.page_margins = copy(src.page_margins)
    dest.print_options = copy(src.print_options)
    if src.print_area:
        dest.print_area = _remap_range_string(src.print_area, comments_col)
    if src.print_title_cols:
        dest.print_title_cols = src.print_title_cols
    if src.print_title_rows:
        dest.print_title_rows = src.print_title_rows
    dest.oddHeader = copy(src.oddHeader)
    dest.oddFooter = copy(src.oddFooter)
    dest.evenHeader = copy(src.evenHeader)
    dest.evenFooter = copy(src.evenFooter)
    dest.firstHeader = copy(src.firstHeader)
    dest.firstFooter = copy(src.firstFooter)
    dest.sheet_properties.pageSetUpPr = copy(src.sheet_properties.pageSetUpPr)


def _copy_conditional_formatting(src: Worksheet, dest: Worksheet, comments_col: Optional[int]) -> None:
    for cf_range in src.conditional_formatting:
        remapped_sqref = _remap_range_string(str(cf_range.sqref), comments_col)
        if not remapped_sqref:
            continue  # this rule applied only to the now-removed column
        for rule in cf_range.rules:
            dest.conditional_formatting.add(remapped_sqref, copy(rule))
