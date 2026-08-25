"""
Regression suite for the INPUT-heading-as-display-label fix.

Business requirement:
The Group marker (e.g. "Track 1-Secured") is the sole, authoritative
source of INTERNAL section identity -- this was already true and remains
completely unchanged.

The human-readable heading text in column A of the source workbook is now
the authoritative DISPLAY LABEL for that run: when a business user renames
a section's heading (while leaving its Group marker untouched), that exact
renamed text should appear everywhere the section's heading is printed in
the output (Worksheet 1, Worksheet 2), instead of silently reverting to
config.py's fixed section.title.

Sheet 3 already got this right incidentally, since it's a verbatim copy
of the source.

Root cause:
excel_reader.py's read_project_rows() read the heading text only to test
"is this a heading row", then discarded it -- it was never stored anywhere.

summary_writer.py's build() always printed config.py's section.title.

Fix:
read_project_rows() gained an optional out-parameter,
section_headings: Dict[str, str], populated as:

    {section_key: exact heading text}

using setdefault so a later row that reuses the same marker for a different,
decorative purpose (e.g. "Sub-Total : Secured" reusing "Staffing-Secured")
never overwrites the genuine heading.

summary_writer.py's build() now accepts this mapping and prefers it over
section.title wherever a section's own heading is printed, falling back to
section.title when a section had no rows this run (and therefore no heading
to capture).

Usage:
    python tests/test_input_heading_propagation.py
"""

from __future__ import annotations

import sys
import subprocess
import tempfile
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEST_DIR = Path(__file__).resolve().parent

sys.path.insert(0, str(TEST_DIR))
sys.path.insert(0, str(PROJECT_ROOT))

import openpyxl  # noqa: E402
import config  # noqa: E402
from fixture_builder import new_workbook, write_row  # noqa: E402


SECTIONS = {
    "track1": {
        "heading": "Solutions and Staff Augmentation (Projects) - Track 1",
        "marker": "Track 1-Secured",
        "sub_group": "DS10_Secured",
        "key": "projects_track1",
    },
    "track2": {
        "heading": "Solutions and Staff Augmentation (Projects) - Track 2",
        "marker": "Track 2-Secured",
        "sub_group": "DS20_Secured",
        "key": "projects_track2",
    },
    "staffing": {
        "heading": "Staffing",
        "marker": "Staffing-Secured",
        "sub_group": "DS70_Secured",
        "key": "staffing_secured",
    },
    "investments": {
        "heading": "Investments",
        "marker": "Investments",
        "sub_group": "DS10_Secured",
        "key": "investments",
    },
    "track1proj": {
        "heading": "Track 1 (Projection)",
        "marker": "Track 1-Projections",
        "sub_group": "DS30_Projection",
        "key": "projects_track1_projection",
    },
    "track2proj": {
        "heading": "Track 2 (Projection)",
        "marker": "Track 2-Projections",
        "sub_group": "DS50_Projection",
        "key": "projects_track2_projection",
    },
    "staffing_projections": {
        "heading": "Staffing (Projections)",
        "marker": "Staffing-Projections",
        "sub_group": "DS95_Secured",
        "key": "staffing_projections",
    },
}


BASE_VALUES = {
    "track1": 111111,
    "track2": 222222,
    "staffing": 333333,
    "investments": 0,
    "track1proj": 444444,
    "track2proj": 555555,
    "staffing_projections": 666666,
}


def build_full_fixture(heading_overrides: dict | None = None) -> Path:
    """
    Build a fixture containing all 6 sections.

    heading_overrides maps a section key, e.g. "track1", to replacement
    heading text. Any section not present in the dictionary keeps its
    original heading.
    """
    heading_overrides = heading_overrides or {}

    wb, ws, layout = new_workbook()

    r = 4

    for key in [
        "track1",
        "track2",
        "investments",
        "track1proj",
        "track2proj",
        "staffing_projections",
        "staffing",
    ]:
        section = SECTIONS[key]

        heading = heading_overrides.get(
            key,
            section["heading"],
        )

        ws.cell(
            row=r,
            column=1,
            value=heading,
        )

        ws.cell(
            row=r,
            column=44,
            value=section["marker"],
        )

        r += 2

        value = BASE_VALUES[key]

        write_row(
            ws,
            layout,
            r,
            f"{key}CustA - SOW",
            "Priya",
            "Project",
            f"{key}CustA",
            section["sub_group"],
            1,
            value,
            value // 10 if value else 0,
        )

        r += 1

        ws.cell(
            row=r,
            column=1,
            value=f"Sub-total : {key}",
        )

        r += 2

    ws.cell(
        row=r,
        column=1,
        value="TOTAL Secured",
    )

    ws.cell(
        row=r + 1,
        column=1,
        value="TOTAL Prospecting",
    )

    path = Path(tempfile.mkdtemp()) / "input.xlsx"
    wb.save(path)

    return path


def run_entry_point(
    entry: str,
    input_path: Path,
    out_dir: Path,
):
    """
    Run either the main.py or GUI entry point.

    entry:
        "main" -> main.main(...)
        "gui"  -> gui.runner.generate_summary(...)
    """

    if entry == "main":
        code = f"""
import sys
sys.path.insert(0, {str(PROJECT_ROOT)!r})

import main as sfae_main

rc = sfae_main.main([
    '--input',
    {str(input_path)!r},
    '--output-dir',
    {str(out_dir)!r},
    '--year',
    '2026',
])

print('RC=' + str(rc))
"""

    else:
        code = f"""
import sys
sys.path.insert(0, {str(PROJECT_ROOT)!r})

from gui.runner import generate_summary

result = generate_summary(
    {str(input_path)!r},
    {str(out_dir)!r},
    year=2026,
)

print('RC=' + ('0' if result.success else '1'))

if not result.success:
    print('TITLE=' + str(result.error_title))
    print(
        'MESSAGE='
        + str(result.error_message).replace(chr(10), ' | ')
    )
"""

    # Use the currently running Python interpreter instead of requiring
    # a separate "python3" executable to exist on Windows.
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=60,
        cwd=str(PROJECT_ROOT),
    )

    rc = None

    for line in proc.stdout.splitlines():
        if line.startswith("RC="):
            rc = int(line[3:])

    return rc, proc.stdout, proc.stderr


def get_headings_by_row(
    output_file: Path,
    sheet: str,
) -> dict:
    """
    Return {row_number: text} for every non-empty column-A cell.
    """
    wb = openpyxl.load_workbook(
        output_file,
        data_only=True,
    )

    ws = wb[sheet]

    return {
        row: ws.cell(row=row, column=1).value
        for row in range(1, ws.max_row + 1)
        if ws.cell(row=row, column=1).value
    }


def find_heading_text(
    output_file: Path,
    sheet: str,
    section_key_hint_words: list,
) -> str | None:
    """
    Find the heading banner row whose text contains all supplied
    hint words, case-insensitively.

    Returns the exact text found.
    """
    headings = get_headings_by_row(
        output_file,
        sheet,
    )

    for text in headings.values():
        text_l = str(text).lower()

        if all(
            word.lower() in text_l
            for word in section_key_hint_words
        ):
            return text

    return None


def get_annual_total(
    output_file: Path,
    label: str,
) -> float:
    """
    Read the cached formula values directly from the workbook.

    IMPORTANT:
    This intentionally does NOT use LibreOffice / soffice.

    The production summary_writer.py is expected to populate cached
    formula values. Therefore openpyxl(data_only=True) can read the
    values directly without an external spreadsheet application.
    """
    wb = openpyxl.load_workbook(
        output_file,
        data_only=True,
    )

    ws2 = wb["2026 Monthly Performance"]

    for row in range(1, ws2.max_row + 1):
        if ws2.cell(row=row, column=1).value == label:
            monthly = [
                ws2.cell(row=row, column=column).value
                for column in range(3, 27, 2)
            ]

            numeric_values = [
                value
                for value in monthly
                if isinstance(value, (int, float))
            ]

            assert numeric_values, (
                f"No cached numeric values found for {label!r}. "
                "The workbook may not contain the expected cached "
                "formula values."
            )

            return sum(numeric_values)

    raise AssertionError(
        f"Could not find {label!r} in "
        "'2026 Monthly Performance'."
    )


# ---------------------------------------------------------------------
# Test A-F:
# Each section's heading is changed independently.
# ---------------------------------------------------------------------


@pytest.mark.parametrize(
    "key",
    list(SECTIONS.keys()),
)
def test_single_section_heading_change(
    key: str,
) -> None:
    section = SECTIONS[key]

    changed = section["heading"] + " -- CHANGED"

    path = build_full_fixture(
        {key: changed}
    )

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"main.py failed for {key}: "
        f"rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    # Track 1/2 Projection are Worksheet-2-only sections.
    #
    # Worksheet 1 excludes Projection sections. This is an established,
    # unrelated architectural fact and is intentionally not changed here.
    if key in (
        "track1proj",
        "track2proj",
        "staffing_projections",
    ):
        sheets_to_check = [
            "2026 Monthly Performance",
        ]
    else:
        sheets_to_check = [
            "Multi-Year Revenue & Margin",
            "2026 Monthly Performance",
        ]

    for sheet in sheets_to_check:
        found = find_heading_text(
            output_file,
            sheet,
            [changed],
        )

        assert found == changed, (
            f"{key} / {sheet}: expected renamed heading "
            f"{changed!r}, found {found!r}"
        )

    # Sheet 3 is a verbatim source copy and must also contain
    # the changed heading.
    found_sheet3 = find_heading_text(
        output_file,
        "2026 SOW Performance",
        [changed],
    )

    assert found_sheet3 == changed, (
        f"{key} / 2026 SOW Performance: expected "
        f"{changed!r}, found {found_sheet3!r}"
    )


# ---------------------------------------------------------------------
# Test G:
# All six sections changed simultaneously.
# ---------------------------------------------------------------------


def test_all_six_simultaneously() -> None:
    overrides = {
        key: SECTIONS[key]["heading"]
        + " :: RENAMED :: "
        + key.upper()
        for key in SECTIONS
    }

    path = build_full_fixture(overrides)

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"main.py failed for all-six test: "
        f"rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    for key, changed in overrides.items():

        if key in (
            "track1proj",
            "track2proj",
            "staffing_projections",
        ):
            sheets_to_check = [
                "2026 Monthly Performance",
                "2026 SOW Performance",
            ]
        else:
            sheets_to_check = [
                "Multi-Year Revenue & Margin",
                "2026 Monthly Performance",
                "2026 SOW Performance",
            ]

        for sheet in sheets_to_check:
            found = find_heading_text(
                output_file,
                sheet,
                [changed],
            )

            assert found == changed, (
                f"All-six / {key} / {sheet}: "
                f"expected {changed!r}, found {found!r}"
            )

    # Data integrity:
    # Renaming headings must not change business calculations.
    expected_secured = (
        BASE_VALUES["track1"]
        + BASE_VALUES["track2"]
        + BASE_VALUES["staffing"]
    )

    expected_prospecting = (
        BASE_VALUES["track1proj"]
        + BASE_VALUES["track2proj"]
        + BASE_VALUES["staffing_projections"]
    )

    total_secured = get_annual_total(
        output_file,
        "TOTAL Secured",
    )

    total_prospecting = get_annual_total(
        output_file,
        "TOTAL Prospecting",
    )

    assert total_secured == expected_secured, (
        f"TOTAL Secured = {total_secured}, "
        f"expected {expected_secured}"
    )

    assert total_prospecting == expected_prospecting, (
        f"TOTAL Prospecting = {total_prospecting}, "
        f"expected {expected_prospecting}"
    )

    label_map = {
        "track1": "Subtotal : Track 1",
        "track2": "Subtotal : Track 2",
        "staffing": "Subtotal : Staffing- Secured",
        "investments": "Subtotal : Investments",
        "track1proj": "Subtotal : Track 1 (Projection)",
        "track2proj": "Subtotal : Track 2 (Projection)",
        "staffing_projections": "Subtotal : Staffing- Projections",
    }

    for key in SECTIONS:
        actual = get_annual_total(
            output_file,
            label_map[key],
        )

        expected = BASE_VALUES[key]

        assert actual == expected, (
            f"{key} subtotal = {actual}, "
            f"expected {expected}"
        )


# ---------------------------------------------------------------------
# Test H:
# Original headings remain unchanged.
# ---------------------------------------------------------------------


def test_original_headings_regression() -> None:
    path = build_full_fixture({})

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"Original-heading regression failed: "
        f"rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    for key, section in SECTIONS.items():

        if key in (
            "track1proj",
            "track2proj",
            "staffing_projections",
        ):
            sheets_to_check = [
                "2026 Monthly Performance",
            ]
        else:
            sheets_to_check = [
                "Multi-Year Revenue & Margin",
                "2026 Monthly Performance",
            ]

        for sheet in sheets_to_check:
            headings = get_headings_by_row(
                output_file,
                sheet,
            )

            assert section["heading"] in headings.values(), (
                f"Original heading for {key} was not found "
                f"in {sheet}: {section['heading']!r}"
            )


# ---------------------------------------------------------------------
# Test I:
# A configured section with zero rows falls back to config.py title.
# ---------------------------------------------------------------------


def test_zero_row_section_fallback() -> None:
    # Build a fixture with Track 1 only.
    #
    # Track 2, Staffing, Investments and both Projection sections
    # never appear, so their display headings are not captured from
    # the source workbook.
    #
    # The writer must therefore fall back to config.py's section.title.

    wb, ws, layout = new_workbook()

    r = 4

    ws.cell(
        row=r,
        column=1,
        value="Solutions and Staff Augmentation (Projects) - Track 1",
    )

    ws.cell(
        row=r,
        column=44,
        value="Track 1-Secured",
    )

    r += 2

    write_row(
        ws,
        layout,
        r,
        "OnlyCorp - SOW",
        "Priya",
        "Project",
        "OnlyCorp",
        "DS10_Secured",
        1,
        50000,
        5000,
    )

    r += 1

    ws.cell(
        row=r,
        column=1,
        value="Sub-total : Track 1",
    )

    path = Path(tempfile.mkdtemp()) / "input.xlsx"

    wb.save(path)

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"Zero-row fallback failed: rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    headings = get_headings_by_row(
        output_file,
        "2026 Monthly Performance",
    )

    values = list(headings.values())

    # Track 2 must fall back to config.py's canonical title.
    assert SECTIONS["track2"]["heading"] in values, (
        "Track 2's config.py title was not found in the "
        f"zero-row fallback output. Headings: {values}"
    )

    # Staffing-Projections (also absent from this fixture) must
    # likewise fall back to config.py's canonical title, not a blank
    # heading and not another section's text. Looked up directly from
    # config.py (rather than SECTIONS["staffing_projections"]["heading"],
    # which deliberately holds the real workbook's own input-heading
    # text, not config.py's fallback title -- those two are expected
    # to differ).
    staffing_projections_fallback_title = next(
        s.title
        for s in config.WORKSHEET2_ADDITIONAL_SECTIONS
        if s.key == "staffing_projections"
    )

    assert staffing_projections_fallback_title in values, (
        "Staffing-Projections' config.py title was not found in the "
        f"zero-row fallback output. Headings: {values}"
    )

    assert None not in values, (
        "A None heading was printed in the zero-row fallback output."
    )

    assert "" not in values, (
        "A blank heading was printed in the zero-row fallback output."
    )


# ---------------------------------------------------------------------
# Test J:
# Arbitrary heading text is preserved verbatim.
# ---------------------------------------------------------------------


def test_arbitrary_heading_text() -> None:
    arbitrary = "2026 Secured Solutions Projects"

    path = build_full_fixture(
        {"track1": arbitrary}
    )

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"Arbitrary-heading test failed: rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    for sheet in [
        "Multi-Year Revenue & Margin",
        "2026 Monthly Performance",
    ]:
        headings = get_headings_by_row(
            output_file,
            sheet,
        )

        assert arbitrary in headings.values(), (
            f"Arbitrary heading {arbitrary!r} was not found "
            f"in {sheet}. Headings: {list(headings.values())}"
        )


# ---------------------------------------------------------------------
# Test K:
# Track 1 / Investments share the same DS-code but remain isolated.
# ---------------------------------------------------------------------


def test_track1_investments_isolation() -> None:
    t1_changed = (
        "Track 1 -- Secured Projects Renamed"
    )

    inv_changed = (
        "Investments -- Internal Renamed"
    )

    path = build_full_fixture(
        {
            "track1": t1_changed,
            "investments": inv_changed,
        }
    )

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"Track 1 / Investments isolation failed: "
        f"rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    headings = get_headings_by_row(
        output_file,
        "2026 Monthly Performance",
    )

    values = list(headings.values())

    assert t1_changed in values, (
        "Track 1's own renamed heading is missing. "
        f"Headings: {values}"
    )

    assert inv_changed in values, (
        "Investments' own renamed heading is missing. "
        f"Headings: {values}"
    )

    assert t1_changed != inv_changed

    # Verify the shared DS-code did not cause the calculations
    # to cross-contaminate.
    t1_total = get_annual_total(
        output_file,
        "Subtotal : Track 1",
    )

    inv_total = get_annual_total(
        output_file,
        "Subtotal : Investments",
    )

    assert t1_total == BASE_VALUES["track1"], (
        f"Track 1 total = {t1_total}, "
        f"expected {BASE_VALUES['track1']} "
        "(possible cross-contamination)"
    )

    assert inv_total == BASE_VALUES["investments"], (
        f"Investments total = {inv_total}, "
        f"expected {BASE_VALUES['investments']} "
        "(possible cross-contamination)"
    )


# ---------------------------------------------------------------------
# Test K2:
# Staffing-Secured / Staffing-Projections are two distinct sections
# (different Group markers, different DS-codes, different categories --
# "secured" vs "prospecting") and must remain fully isolated. Renaming
# one's heading must not affect the other's heading or totals.
# ---------------------------------------------------------------------


def test_staffing_secured_and_projections_isolation() -> None:
    staffing_changed = (
        "Staffing Secured -- Renamed"
    )

    staffing_proj_changed = (
        "Staffing Projections -- Renamed"
    )

    path = build_full_fixture(
        {
            "staffing": staffing_changed,
            "staffing_projections": staffing_proj_changed,
        }
    )

    out = path.parent / "out"

    rc, stdout, stderr = run_entry_point(
        "main",
        path,
        out,
    )

    assert rc == 0, (
        f"Staffing-Secured / Staffing-Projections isolation failed: "
        f"rc={rc}\n"
        f"STDOUT:\n{stdout[-1000:]}\n"
        f"STDERR:\n{stderr[-1000:]}"
    )

    output_file = (
        out / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    headings = get_headings_by_row(
        output_file,
        "2026 Monthly Performance",
    )

    values = list(headings.values())

    assert staffing_changed in values, (
        "Staffing-Secured's own renamed heading is missing. "
        f"Headings: {values}"
    )

    assert staffing_proj_changed in values, (
        "Staffing-Projections' own renamed heading is missing. "
        f"Headings: {values}"
    )

    assert staffing_changed != staffing_proj_changed

    # Worksheet 1 must show Staffing-Secured (a "secured" section) but
    # must NOT show Staffing-Projections (a "prospecting" section) --
    # same established rule already verified for Track 1/2 Projection.
    ws1_headings = get_headings_by_row(
        output_file,
        "Multi-Year Revenue & Margin",
    )
    ws1_values = list(ws1_headings.values())

    assert staffing_changed in ws1_values, (
        "Staffing-Secured's renamed heading is missing from Worksheet 1. "
        f"Headings: {ws1_values}"
    )

    assert staffing_proj_changed not in ws1_values, (
        "Staffing-Projections (a Worksheet-2-only, prospecting-category "
        "section) incorrectly appeared on Worksheet 1. "
        f"Headings: {ws1_values}"
    )

    # Verify the two sections' totals did not cross-contaminate.
    staffing_total = get_annual_total(
        output_file,
        "Subtotal : Staffing- Secured",
    )

    staffing_proj_total = get_annual_total(
        output_file,
        "Subtotal : Staffing- Projections",
    )

    assert staffing_total == BASE_VALUES["staffing"], (
        f"Staffing-Secured total = {staffing_total}, "
        f"expected {BASE_VALUES['staffing']} "
        "(possible cross-contamination with Staffing-Projections)"
    )

    assert staffing_proj_total == BASE_VALUES["staffing_projections"], (
        f"Staffing-Projections total = {staffing_proj_total}, "
        f"expected {BASE_VALUES['staffing_projections']} "
        "(possible cross-contamination with Staffing-Secured)"
    )

    # Staffing-Secured must contribute to TOTAL Secured;
    # Staffing-Projections must contribute to TOTAL Prospecting -- not
    # the other way around, and not both.
    total_secured = get_annual_total(
        output_file,
        "TOTAL Secured",
    )

    total_prospecting = get_annual_total(
        output_file,
        "TOTAL Prospecting",
    )

    assert BASE_VALUES["staffing"] <= total_secured, (
        "Staffing-Secured's own value does not appear to be included "
        "in TOTAL Secured."
    )

    assert BASE_VALUES["staffing_projections"] <= total_prospecting, (
        "Staffing-Projections' own value does not appear to be "
        "included in TOTAL Prospecting."
    )


# ---------------------------------------------------------------------
# Test L:
# main.py vs gui/runner.py parity.
# ---------------------------------------------------------------------


def test_main_gui_parity() -> None:
    overrides = {
        key: SECTIONS[key]["heading"]
        + " (PARITY-TEST)"
        for key in SECTIONS
    }

    path = build_full_fixture(overrides)

    out_main = path.parent / "out_main"
    out_gui = path.parent / "out_gui"

    rc_main, stdout_main, stderr_main = run_entry_point(
        "main",
        path,
        out_main,
    )

    rc_gui, stdout_gui, stderr_gui = run_entry_point(
        "gui",
        path,
        out_gui,
    )

    assert rc_main == 0, (
        f"main.py failed during parity test: rc={rc_main}\n"
        f"STDOUT:\n{stdout_main[-1000:]}\n"
        f"STDERR:\n{stderr_main[-1000:]}"
    )

    assert rc_gui == 0, (
        f"gui/runner.py failed during parity test: rc={rc_gui}\n"
        f"STDOUT:\n{stdout_gui[-1000:]}\n"
        f"STDERR:\n{stderr_gui[-1000:]}"
    )

    from openpyxl.worksheet.formula import ArrayFormula

    main_file = (
        out_main / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    gui_file = (
        out_gui / "Sales_and_Forecast_Summary_2026.xlsx"
    )

    wb_main = openpyxl.load_workbook(
        main_file,
        data_only=False,
    )

    wb_gui = openpyxl.load_workbook(
        gui_file,
        data_only=False,
    )

    def normalize(value):
        if isinstance(value, ArrayFormula):
            return (
                "AF",
                value.ref,
                value.text,
            )

        return value

    assert wb_main.sheetnames == wb_gui.sheetnames, (
        "main.py and gui/runner.py produced different "
        "worksheet names."
    )

    differences = []

    for sheet in wb_main.sheetnames:
        ws_main = wb_main[sheet]
        ws_gui = wb_gui[sheet]

        max_row = max(
            ws_main.max_row,
            ws_gui.max_row,
        )

        max_column = max(
            ws_main.max_column,
            ws_gui.max_column,
        )

        for row in range(
            1,
            max_row + 1,
        ):
            for column in range(
                1,
                max_column + 1,
            ):
                main_value = normalize(
                    ws_main.cell(
                        row=row,
                        column=column,
                    ).value
                )

                gui_value = normalize(
                    ws_gui.cell(
                        row=row,
                        column=column,
                    ).value
                )

                if main_value != gui_value:
                    differences.append(
                        (
                            sheet,
                            row,
                            column,
                            main_value,
                            gui_value,
                        )
                    )

    assert not differences, (
        f"main.py and gui/runner.py differ in "
        f"{len(differences)} cells. "
        f"First differences: {differences[:10]}"
    )


# ---------------------------------------------------------------------
# Standalone execution support.
#
# Running:
#
#     python tests/test_input_heading_propagation.py
#
# should still execute the same pytest suite.
# ---------------------------------------------------------------------


if __name__ == "__main__":
    raise SystemExit(
        pytest.main(
            [
                __file__,
                "-q",
            ]
        )
    )