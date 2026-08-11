"""
tests/test_preview_workbook_validation_parity.py
==================================================
Regression suite for a confirmed inconsistency: streamlit_bridge.py's
preview_workbook() (feeds app.py's pre-generation "once uploaded"
panel) had its own independent processing path that never called
row_validator.validate_rows() -- unlike main.py and gui/runner.py,
both of which already do. For a workbook with an unrecognized Group
marker, this let the preview silently report a misleading
"Track 1: 0 groups" with no error at all, while the real generation
(through either other entry point) correctly rejected the same file
with a clear "Unknown Group marker" error -- an inconsistent,
confusing user experience, though never a source of data loss in the
delivered workbook itself (generation was always protected).

Fix: preview_workbook() now calls the exact same row_validator.
validate_rows() function main.py and gui/runner.py already use (not a
reimplementation), and raises GenerationError -- preview_workbook's
own, pre-existing error convention -- when any row is ERROR-status,
using the identical Missing Name / Missing Group / raw-reason message
construction already established in gui/runner.py.

Usage:
    python tests/test_preview_workbook_validation_parity.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEST_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TEST_DIR))
sys.path.insert(0, str(PROJECT_ROOT))

from fixture_builder import new_workbook, write_row  # noqa: E402
from streamlit_bridge import preview_workbook, GenerationError  # noqa: E402


def build_full_fixture(track1_heading: str = None, track2_heading: str = None,
                        staffing_heading: str = None) -> Path:
    """Track 1 + Track 2 + Staffing + Investments, each with 2
    customers, DS10 shared between Track 1 and Investments (the real
    production shape). Heading TEXT for any of the three named
    sections can be overridden independently; Group markers always
    stay at their real, correct values."""
    wb, ws, layout = new_workbook()
    r = 4
    ws.cell(row=r, column=1, value=track1_heading or "Solutions and Staff Augmentation (Projects) - Track 1")
    ws.cell(row=r, column=44, value="Track 1-Secured")
    r += 2
    write_row(ws, layout, r, "Track1CustA - SOW", "Priya", "Project", "Track1CustA", "DS10_Secured", 1, 100000, 10000); r += 1
    write_row(ws, layout, r, "Track1CustB - SOW", "Priya", "Project", "Track1CustB", "DS10_Secured", 1, 50000, 5000); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Track 1"); r += 2

    ws.cell(row=r, column=1, value="Investments")
    ws.cell(row=r, column=44, value="Investments")
    r += 2
    write_row(ws, layout, r, "InvestCust - SOW", "Vijay", "Internal", "InvestCust", "DS10_Secured", 1, 0, 0); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Investments"); r += 2

    ws.cell(row=r, column=1, value=track2_heading or "Solutions and Staff Augmentation (Projects) - Track 2")
    ws.cell(row=r, column=44, value="Track 2-Secured")
    r += 2
    write_row(ws, layout, r, "Track2Cust - SOW", "Dana", "Staff Aug", "Track2Cust", "DS20_Secured", 1, 70000, 7000); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Track 2"); r += 2

    ws.cell(row=r, column=1, value=staffing_heading or "Staffing")
    ws.cell(row=r, column=44, value="Staffing-Secured")
    r += 2
    write_row(ws, layout, r, "StaffCust - SOW", "Wendy", "Staffing", "StaffCust", "DS70_Secured", 1, 40000, 4000); r += 1
    ws.cell(row=r, column=1, value="Sub-total : Staffing")

    path = Path(tempfile.mkdtemp()) / "input.xlsx"
    wb.save(path)
    return path


def test_A_valid_workbook_preview_succeeds() -> bool:
    path = build_full_fixture()
    try:
        p = preview_workbook(str(path), year=2026)
    except GenerationError as exc:
        print(f"  FAIL: valid workbook was rejected -- {exc.title}: {exc.message}")
        return False
    expected = {
        "Solutions and Staff Augmentation (Projects) - Track 1": 2,
        "Investments": 1,
        "Solutions and Staff Augmentation (Projects) - Track 2": 1,
        "Staffing- Secured": 1,
    }
    ok = all(p.section_group_counts.get(title) == count for title, count in expected.items())
    if not ok:
        print(f"  FAIL: group counts don't match expected -- got {p.section_group_counts}")
    return ok


def test_B_unknown_group_marker_raises_clearly() -> bool:
    wb, ws, layout = new_workbook()
    r = 4
    ws.cell(row=r, column=1, value="Solutions and Staff Augmentation (Projects) - Track 1")
    ws.cell(row=r, column=44, value="Track 99-Secured")  # unconfigured
    r += 2
    write_row(ws, layout, r, "NewCorp - SOW", "Priya", "Project", "NewCorp", "DS10_Secured", 1, 10000, 1000)
    path = Path(tempfile.mkdtemp()) / "input.xlsx"
    wb.save(path)

    try:
        p = preview_workbook(str(path), year=2026)
        print(f"  FAIL: expected GenerationError, got a silent success -- num_groups={p.num_groups}, "
              f"section_group_counts={p.section_group_counts}")
        return False
    except GenerationError as exc:
        ok = "Unknown Group marker" in exc.message
        if not ok:
            print(f"  FAIL: raised, but message doesn't mention 'Unknown Group marker' -- got: {exc.message!r}")
        return ok


def test_C_blank_group_on_customer_row_says_missing_group() -> bool:
    wb, ws, layout = new_workbook()
    r = 4
    ws.cell(row=r, column=1, value="Solutions and Staff Augmentation (Projects) - Track 1")
    ws.cell(row=r, column=44, value="Track 1-Secured")
    r += 2
    write_row(ws, layout, r, "GoodCorp - SOW", "Priya", "Project", "GoodCorp", "DS10_Secured", 1, 100000, 10000)
    r += 1
    ws.cell(row=r, column=1, value="BlankGroupCorp - SOW")
    ws.cell(row=r, column=44, value="")  # blank Group on a genuine customer row (Sub-Group IS set)
    ws.cell(row=r, column=45, value="DS10_Secured")
    path = Path(tempfile.mkdtemp()) / "input.xlsx"
    wb.save(path)

    try:
        preview_workbook(str(path), year=2026)
        print("  FAIL: expected GenerationError, got a silent success")
        return False
    except GenerationError as exc:
        has_missing_group = "Missing Group" in exc.message
        has_wrong_message = "Unknown Group marker: BlankGroupCorp" in exc.message
        if not has_missing_group or has_wrong_message:
            print(f"  FAIL: expected 'Missing Group', not 'Unknown Group marker: <name>' -- got: {exc.message!r}")
            return False
        return True


def test_D_track1_heading_variation_preview_remains_correct() -> bool:
    path = build_full_fixture(track1_heading="Solutions and Staff Augmentation (Projects) - Track 1 (Secured)")
    try:
        p = preview_workbook(str(path), year=2026)
    except GenerationError as exc:
        print(f"  FAIL: renamed Track 1 heading incorrectly rejected -- {exc.title}: {exc.message}")
        return False
    # section_group_counts is keyed by config.py's canonical title (the
    # stable internal identity), not by whatever heading text happens
    # to be in this particular workbook -- that stability is exactly
    # the property being verified here.
    count = p.section_group_counts.get("Solutions and Staff Augmentation (Projects) - Track 1")
    if count != 2:
        print(f"  FAIL: expected 2 Track 1 groups despite the renamed heading, got {p.section_group_counts}")
        return False
    return True


def test_E_track2_and_staffing_heading_variations() -> bool:
    path = build_full_fixture(
        track2_heading="Solutions and Staff Augmentation (Projects) - Track 2 - Secured",
        staffing_heading="Staffing (Secured)",
    )
    try:
        p = preview_workbook(str(path), year=2026)
    except GenerationError as exc:
        print(f"  FAIL: renamed Track 2/Staffing headings incorrectly rejected -- {exc.title}: {exc.message}")
        return False
    ok = (
        p.section_group_counts.get("Solutions and Staff Augmentation (Projects) - Track 2") == 1
        and p.section_group_counts.get("Staffing- Secured") == 1
    )
    if not ok:
        print(f"  FAIL: group counts don't match expected -- got {p.section_group_counts}")
    return ok


def main() -> int:
    tests = [
        ("A: valid production-shaped workbook, preview succeeds with correct counts", test_A_valid_workbook_preview_succeeds),
        ("B: unknown Group marker no longer silently shows '0 groups' -- raises clearly", test_B_unknown_group_marker_raises_clearly),
        ("C: blank Group on a customer row says 'Missing Group', not 'Unknown Group marker: <name>'", test_C_blank_group_on_customer_row_says_missing_group),
        ("D: Track 1 heading variation, preview stays correct", test_D_track1_heading_variation_preview_remains_correct),
        ("E: Track 2 and Staffing heading variations, preview stays correct", test_E_track2_and_staffing_heading_variations),
    ]
    all_ok = True
    for label, fn in tests:
        try:
            ok = fn()
        except Exception as exc:  # noqa: BLE001
            print(f"  FAIL ({label}): unexpected exception -- {type(exc).__name__}: {exc}")
            ok = False
        print(f"{'PASS' if ok else 'FAIL'}: {label}")
        all_ok &= ok
    print(f"\n{'ALL TESTS PASS' if all_ok else 'SOME TESTS FAILED'}")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
