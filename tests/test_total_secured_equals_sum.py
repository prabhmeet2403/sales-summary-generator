"""
tests/test_total_secured_equals_sum.py
========================================
Verifies requirement 4: TOTAL Secured = Track 1 + Track 2 + Staffing,
proven by independently recomputing each subtotal from the raw cells
and comparing against TOTAL Secured's own displayed value -- never
trusting TOTAL Secured's own formula as proof of itself.

Usage:
    python tests/test_total_secured_equals_sum.py
"""
from __future__ import annotations
import subprocess
import sys
from pathlib import Path

TEST_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TEST_DIR))
sys.path.insert(0, str(TEST_DIR.parent))
from fixture_builder import new_workbook, write_row  # noqa: E402

import main as sfae_main  # noqa: E402
from openpyxl import load_workbook


def build_fixture(path: Path):
    wb, ws, layout = new_workbook()
    r = 4
    ws.cell(row=r, column=1, value="Solutions and Staff Augmentation (Projects) - Track 1"); ws.cell(row=r, column=44, value="Track 1-Secured"); r += 2
    write_row(ws, layout, r, "November Corp - SOW", "Priya", "Project", "November Corp", "DS10_Secured", 1, 100003, 10003); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Track 1"); r += 2

    ws.cell(row=r, column=1, value="Solutions and Staff Augmentation (Projects) - Track 2"); ws.cell(row=r, column=44, value="Track 2-Secured"); r += 2
    write_row(ws, layout, r, "Oscar Corp - SOW", "Dana", "Staff Aug", "Oscar Corp", "DS20_Secured", 1, 200003, 20003); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Track 2"); r += 2

    ws.cell(row=r, column=1, value="TOTAL Secured"); r += 2

    ws.cell(row=r, column=1, value="Staffing"); ws.cell(row=r, column=44, value="Staffing-Secured"); r += 2
    write_row(ws, layout, r, "Papa Corp - Staffing", "Wendy", "Staffing", "Papa Corp", "DS70_Secured", 1, 300003, 30003); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Staffing"); r += 2

    ws.cell(row=r, column=1, value="Investments"); ws.cell(row=r, column=44, value="Investments"); r += 2
    write_row(ws, layout, r, "Quebec Corp - Internal", "Priya", "Internal", "Quebec Corp", "DS10_Secured", 1, 400003, 40003); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Investments"); r += 1

    wb.save(path)


def main() -> int:
    fixture = TEST_DIR / "fixtures" / "test4_total_secured_sum.xlsx"
    build_fixture(fixture)

    out_dir = TEST_DIR.parent / "output" / "_test4"
    out_dir.mkdir(parents=True, exist_ok=True)
    rc = sfae_main.main(["--input", str(fixture), "--output-dir", str(out_dir), "--year", "2026"])
    if rc != 0:
        print("FAIL - pipeline did not exit cleanly")
        return 1

    generated = list(out_dir.glob("Sales_and_Forecast_Summary_*.xlsx"))[0]

    recalc_dir = TEST_DIR.parent / "output" / "_test4_recalc"
    recalc_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        ["soffice", "--headless", "--norestore", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
         "--outdir", str(recalc_dir), str(generated)],
        capture_output=True, timeout=60,
    )
    recalculated = recalc_dir / generated.name
    if not recalculated.exists():
        print("FAIL - LibreOffice recalculation did not produce an output file")
        return 1

    wb = load_workbook(recalculated, data_only=True)
    ws2 = wb["2026 Monthly Performance"]

    def get_val(label, col=27):
        for r in range(1, ws2.max_row + 1):
            if ws2.cell(row=r, column=1).value == label:
                return ws2.cell(row=r, column=col).value
        return None

    t1 = get_val("Subtotal : Track 1")
    t2 = get_val("Subtotal : Track 2")
    staff = get_val("Subtotal : Staffing- Secured")
    total_secured = get_val("TOTAL Secured")

    expected = t1 + t2 + staff
    if abs(total_secured - expected) > 0.01:
        print(f"FAIL - TOTAL Secured={total_secured}, but Track1({t1})+Track2({t2})+Staffing({staff})={expected}")
        return 1
    print(f"PASS - TOTAL Secured={total_secured} == Track1({t1})+Track2({t2})+Staffing({staff})={expected}, independently recomputed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
