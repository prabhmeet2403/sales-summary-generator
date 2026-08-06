"""
tests/test_cached_value_xml_validity.py
==========================================
Guards against a regression that shipped once already: the
cached-formula-value patch (`summary_writer._inject_cached_formula_values`,
called via `SummaryWriter.patch_cached_formula_values` right after
`wb.save(...)` -- see Issue 1 of the Phase 1 fixes) must always produce
strictly well-formed worksheet XML, or Excel opens the file with a
repair-dialog ("Replaced Part: /xl/worksheets/sheetN.xml, Load error.
Line 1, column 0.").

The original implementation spliced raw XML text by hand (find a cell
by its `r="..."` attribute, locate its `<v>` via string search, splice
in a new one). That broke as soon as a `<v>` element wasn't already
empty: the regex used to replace it only matched an EMPTY `<v></v>` (or
self-closing `<v/>`), so a non-empty `<v>123</v>` was only partially
replaced, leaving a dangling, unmatched `</v>` behind -- e.g.
`<v>999</v>` becoming `<v>NEW</v>999</v>`. This test locks in the fix
(parsing each sheet into a real element tree and editing it via the
tree API, which cannot produce a mismatched tag by construction) by
directly reproducing that exact scenario and confirming the result
still parses.

Usage:
    python tests/test_cached_value_xml_validity.py
"""
from __future__ import annotations

import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from gui.runner import generate_summary  # noqa: E402
from summary_writer import _patch_sheet_cached_values  # noqa: E402

FIXTURE_MASTER = Path(__file__).resolve().parent / "fixtures" / "master_2026.xlsx"

_SML_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def main() -> int:
    problems: list = []

    # --- 1. Direct reproduction of the exact scenario that broke the
    #        old regex-based patch: a cell whose <v> is NOT empty. ---
    sheet_xml_with_existing_value = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<sheetData><row r="36"><c r="C36" s="18"><f>SUM(C7:C34)</f><v>999</v></c></row></sheetData>'
        b'</worksheet>'
    )
    patched = _patch_sheet_cached_values(sheet_xml_with_existing_value, [("C36", 12345.67)])
    if patched is None:
        problems.append("_patch_sheet_cached_values returned None for valid input")
    else:
        try:
            root = ET.fromstring(patched)
        except ET.ParseError as exc:
            problems.append(f"Patching a cell with an already non-empty <v> produced invalid XML: {exc}\n  Output: {patched!r}")
        else:
            v_elems = root.findall(f".//{_SML_NS}v")
            if len(v_elems) != 1:
                problems.append(f"Expected exactly one <v> element after patching, found {len(v_elems)}")
            elif v_elems[0].text != "12345.67":
                problems.append(f"<v> text is {v_elems[0].text!r}, expected '12345.67'")

    # --- 2. Patching the SAME sheet TWICE in a row (simulating any
    #        double-processing) must still yield valid XML with the
    #        LATEST value, not corruption from stacking edits. ---
    once = _patch_sheet_cached_values(sheet_xml_with_existing_value, [("C36", 111.0)])
    twice = _patch_sheet_cached_values(once, [("C36", 222.0)])
    try:
        root = ET.fromstring(twice)
        v_text = root.find(f".//{_SML_NS}v").text
        if v_text != "222":
            problems.append(f"Double-patching should leave the LATEST value, got {v_text!r}")
    except ET.ParseError as exc:
        problems.append(f"Double-patching the same sheet produced invalid XML: {exc}")

    # --- 3. A real, full generation must produce strictly well-formed
    #        XML for every worksheet, checked with the same parser
    #        Excel's own OOXML consumer is built on the same standard
    #        as. ---
    with tempfile.TemporaryDirectory() as tmp:
        result = generate_summary(str(FIXTURE_MASTER), tmp, 2026, progress_cb=lambda m: None)
        if not result.success:
            print("Generation FAILED - cannot verify XML validity.")
            return 1

        with zipfile.ZipFile(result.output_path) as z:
            sheet_parts = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
            if not sheet_parts:
                problems.append("No worksheet XML parts found in the generated file")
            for part in sheet_parts:
                data = z.read(part)
                try:
                    root = ET.fromstring(data)
                except ET.ParseError as exc:
                    problems.append(f"{part} is not well-formed XML: {exc}")
                    continue
                # Every formula cell must have EXACTLY one <v> child with
                # non-empty text -- the actual "totals visible on open"
                # requirement, verified structurally.
                for cell in root.iter(f"{_SML_NS}c"):
                    if cell.find(f"{_SML_NS}f") is None:
                        continue
                    v_elems = cell.findall(f"{_SML_NS}v")
                    if len(v_elems) != 1 or not v_elems[0].text:
                        problems.append(f"{part} cell {cell.get('r')} does not have exactly one populated cached value")

            bad_member = z.testzip()
            if bad_member is not None:
                problems.append(f"Zip CRC check failed for member: {bad_member}")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL CACHED-VALUE XML VALIDITY CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
