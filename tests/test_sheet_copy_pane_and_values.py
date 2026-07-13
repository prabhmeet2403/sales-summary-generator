"""
tests/test_sheet_copy_pane_and_values.py
===========================================
Two focused regression tests for bugs found (and fixed) in
sheet_copy.py after initial delivery:

1. A frozen pane that has ALSO been independently scrolled (its
   `topLeftCell` differs from its actual split point) must round-trip
   exactly -- not through `Worksheet.freeze_panes`'s convenience
   setter, which is lossy for exactly this case (see
   `sheet_copy._copy_sheet_view`'s docstring). Constructed directly
   here (xSplit=3, ySplit=2, topLeftCell="D70") rather than relying on
   any fixture happening to have this same shape, since the bug is
   specifically about frozen-AND-scrolled panes, not simple ones.

2. Every cell in the copied sheet must carry its computed VALUE, never
   a formula -- a formula referencing a sheet this module doesn't also
   copy (e.g. a cross-sheet VLOOKUP) produces both an "external links"
   warning and a #REF!/#N/A when Excel opens the output, which has
   nothing to do with the column-removal feature itself.

Usage:
    python tests/test_sheet_copy_pane_and_values.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from sheet_copy import copy_source_sheet_as_new_worksheet  # noqa: E402


def _build_source_workbook(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Source"
    ws["A1"] = "Name"
    ws["B1"] = "Comments"
    ws["C1"] = "Value"
    ws["A2"] = "Row2"
    ws["B2"] = "a comment"
    ws["C2"] = 100

    # A formula referencing a DIFFERENT sheet that this test deliberately
    # does NOT ask sheet_copy to also bring over -- reproducing the
    # exact "VLOOKUP into another tab" real-world scenario.
    other = wb.create_sheet("Other")
    other["A1"] = 42
    ws["D2"] = "=Other!A1*2"

    # Frozen pane (columns A-C, rows 1-2 -- xSplit=3, ySplit=2) that has
    # ALSO been scrolled down independently, so topLeftCell is nowhere
    # near the actual split point -- the exact shape that broke the old
    # `freeze_panes = topLeftCell` round-trip.
    ws.freeze_panes = "D3"  # establishes the xSplit=3/ySplit=2 split
    ws.sheet_view.pane.topLeftCell = "D70"  # then scrolled down independently

    wb.save(path)

    # openpyxl has no formula engine, so D2's `=Other!A1*2` has no cached
    # <v> at all after a plain save -- unlike a file real Excel has
    # actually opened and calculated at least once. Inject the value a
    # real Excel session would have cached (42*2=84), reusing the exact
    # same injection helper `SummaryWriter.patch_cached_formula_values`
    # already relies on, rather than a second implementation of it.
    from summary_writer import _inject_cached_formula_values
    _inject_cached_formula_values(path, [("Source", "D2", 84.0)])


def main() -> int:
    problems: list = []

    with tempfile.TemporaryDirectory() as tmp:
        source_path = f"{tmp}/source.xlsx"
        _build_source_workbook(source_path)

        out_wb = Workbook()
        out_wb.remove(out_wb.active)
        copy_source_sheet_as_new_worksheet(out_wb, source_path, "Source", comments_col=2)
        out_path = f"{tmp}/out.xlsx"
        out_wb.save(out_path)

        new_ws = load_workbook(out_path)["Source"]

        # --- 1. Pane split/scroll fidelity ---
        pane = new_ws.sheet_view.pane
        if pane is None:
            problems.append("Pane is missing entirely on the copied sheet")
        else:
            # Comments (column B, index 2) IS one of the 3 originally-
            # frozen columns (A-C), so removing it correctly leaves only
            # 2 columns frozen (A, C-now-shifted-to-B) -- xSplit must
            # shift too, not stay at 3.
            if pane.xSplit != 2:
                problems.append(f"xSplit should be 2 (one of the 3 frozen columns was removed), got {pane.xSplit}")
            if pane.ySplit != 2:
                problems.append(f"ySplit should be 2 (the TRUE split point, row-based, unaffected by a column removal), got {pane.ySplit} -- this is exactly the old bug (row-1 from topLeftCell instead)")
            # D (column 4) shifts left by one to C once column B is removed.
            if pane.topLeftCell != "C70":
                problems.append(f"topLeftCell should shift from 'D70' to 'C70', got {pane.topLeftCell!r}")
            if pane.state != "frozen":
                problems.append(f"pane.state should be 'frozen', got {pane.state!r}")

        # --- 2. Values only, never formulas ---
        # D2 (source formula "=Other!A1*2") shifts to C2 (Comments/B removed).
        new_cell = new_ws["C2"]
        if isinstance(new_cell.value, str) and new_cell.value.startswith("="):
            problems.append(f"Cell C2 is still a formula ({new_cell.value!r}), should be the computed value 84")
        elif new_cell.value != 84:
            problems.append(f"Cell C2 should be the computed value 84, got {new_cell.value!r}")

        # --- 3. No formula anywhere on the copied sheet ---
        formula_cells = [
            cell.coordinate for row in new_ws.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
        if formula_cells:
            problems.append(f"Formula cell(s) found on the copied sheet (should be none): {formula_cells}")

        # --- 4. Comments column actually removed ---
        if new_ws.max_column != 3:  # Name, Value, (was D formula) = 3 columns after removing Comments
            problems.append(f"Expected 3 columns after removing Comments, got {new_ws.max_column}")
        headers = [new_ws.cell(row=1, column=c).value for c in range(1, new_ws.max_column + 1)]
        if "Comments" in headers:
            problems.append("'Comments' header still present")

        # --- 5. No external references in the output package ---
        import zipfile
        with zipfile.ZipFile(out_path) as zf:
            if any("external" in n.lower() for n in zf.namelist()):
                problems.append("Output package contains an externalLink part")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL SHEET_COPY PANE/VALUE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
