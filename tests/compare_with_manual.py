"""
tests/compare_with_manual.py
=============================
Automated regression test: regenerates the Summary workbook from the
fixture Master workbook and compares EVERY populated cell against the
golden, manually-produced Summary workbook byte-for-value.

Usage:
    python tests/compare_with_manual.py

Exit code 0 and "ALL CELLS MATCH" means the generated workbook is a
perfect match. Any other outcome prints every mismatched cell (sheet,
coordinate, row label, expected value, actual value) and exits non-zero,
so this can be wired into CI.

This test is intentionally dependency-free (no LibreOffice / no pytest
required): since the only formulas this project ever writes are simple
`=SUM(<cell>:<cell>)` range sums (on the same sheet), a tiny built-in
evaluator resolves them instead of shelling out to a spreadsheet engine.
"""
from __future__ import annotations

import re
import sys
import tempfile
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import range_boundaries

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import main as sfae_main  # noqa: E402

FIXTURE_MASTER_GLOB = "*.xlsx"
GOLDEN_SUMMARY = Path(__file__).resolve().parent / "fixtures" / "manual_summary_2026.xlsx"
GOLDEN_SHEET = "2026"
NUMERIC_TOLERANCE = 0.015  # ignore sub-₹0.01 floating point noise (float64 can land at ~0.0100000002)


def _find_fixture_master() -> Path:
    fixtures_dir = Path(__file__).resolve().parent / "fixtures"
    candidates = sorted(fixtures_dir.glob(FIXTURE_MASTER_GLOB))
    candidates = [c for c in candidates if c.name != GOLDEN_SUMMARY.name]
    if candidates:
        return candidates[0]
    # Fall back to whatever the user has dropped in input/
    input_candidates = sorted((PROJECT_ROOT / "input").glob("*.xlsx"))
    input_candidates = [c for c in input_candidates if not c.name.startswith("~$")]
    if input_candidates:
        return input_candidates[0]
    raise FileNotFoundError(
        "No Master workbook found for testing. Place one at "
        f"{fixtures_dir}/ or {PROJECT_ROOT}/input/."
    )


_SUM_FORMULA_RE = re.compile(r"^=SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$", re.IGNORECASE)


def resolve_cell(ws: Worksheet, row: int, col: int, _depth: int = 0):
    """Return a cell's effective value, evaluating a same-sheet
    `=SUM(A1:A2)` formula if present (the only formula type this project
    ever writes) -- recursively, since a subtotal row's formula sums a
    range of cells that are themselves `=SUM(...)` formulas."""
    if _depth > 5:
        return ws.cell(row, col).value  # safety net against any cycle
    value = ws.cell(row, col).value
    if isinstance(value, str) and value.startswith("="):
        m = _SUM_FORMULA_RE.match(value.strip())
        if m:
            min_col, min_row, max_col, max_row = range_boundaries(f"{m.group(1)}:{m.group(2)}")
            total = 0.0
            saw_number = False
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    v = resolve_cell(ws, r, c, _depth + 1)
                    if isinstance(v, (int, float)):
                        total += v
                        saw_number = True
            return round(total, 2) if saw_number else 0
        return value  # unrecognised formula shape - compare literally
    return value


def _values_match(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= NUMERIC_TOLERANCE
    if a is None:
        a = ""
    if b is None:
        b = ""
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    return a == b


def compare_workbooks(generated_path: Path, golden_path: Path, golden_sheet: str) -> list:
    generated_wb = load_workbook(generated_path, data_only=False)
    golden_wb = load_workbook(golden_path, data_only=True)

    if golden_sheet not in golden_wb.sheetnames:
        raise ValueError(f"Golden workbook has no sheet '{golden_sheet}'.")
    # The generated sheet is named after the target year, same as golden.
    generated_ws = generated_wb[golden_sheet] if golden_sheet in generated_wb.sheetnames else generated_wb.active
    golden_ws = golden_wb[golden_sheet]

    max_row = max(generated_ws.max_row, golden_ws.max_row)
    max_col = max(generated_ws.max_column, golden_ws.max_column)

    diffs = []
    for r in range(1, max_row + 1):
        row_label = golden_ws.cell(r, 1).value or generated_ws.cell(r, 1).value
        for c in range(1, max_col + 1):
            golden_val = golden_ws.cell(r, c).value
            generated_val = resolve_cell(generated_ws, r, c)
            if golden_val is None and generated_val is None:
                continue
            if not _values_match(golden_val, generated_val):
                diffs.append(
                    {
                        "row": r,
                        "col": c,
                        "row_label": row_label,
                        "expected": golden_val,
                        "actual": generated_val,
                    }
                )
    return diffs


def main() -> int:
    fixture_master = _find_fixture_master()

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        exit_code = sfae_main.main(
            ["--input", str(fixture_master), "--output-dir", str(tmp_dir), "--year", "2026"]
        )
        if exit_code != 0:
            print("Generation run FAILED - cannot compare. Check the validation report above.")
            return 1

        generated_files = sorted(tmp_dir.glob("*.xlsx"))
        if not generated_files:
            print("No output workbook was generated.")
            return 1
        generated_path = generated_files[0]

        diffs = compare_workbooks(generated_path, GOLDEN_SUMMARY, GOLDEN_SHEET)

    if not diffs:
        print(f"ALL CELLS MATCH ({GOLDEN_SUMMARY.name}, sheet '{GOLDEN_SHEET}') - 0 differences. PASS")
        return 0

    from openpyxl.utils import get_column_letter

    print(f"{len(diffs)} CELL DIFFERENCE(S) FOUND vs {GOLDEN_SUMMARY.name} (sheet '{GOLDEN_SHEET}'):\n")
    for d in diffs:
        coord = f"{get_column_letter(d['col'])}{d['row']}"
        print(
            f"  Row {d['row']:>3} [{d['row_label']!s:45.45}] {coord:>5}: "
            f"expected={d['expected']!r:>20}  actual={d['actual']!r}"
        )
    print(f"\nFAIL - {len(diffs)} difference(s).")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
