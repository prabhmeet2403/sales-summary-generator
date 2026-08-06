"""
row_validator.py
================
A new, purely additive validation layer. It runs a parallel scan of the
same worksheet `read_project_rows` reads, using the same heading
detection and helper functions (imported, not duplicated, so the two
scans can never drift out of sync), and classifies every physical row
as VALID, WARNING, or ERROR -- including rows that `read_project_rows`
silently excludes.

This module changes no existing behavior. It does not filter, alter,
or reorder anything `aggregator.py`/`monthly_view.py` see. It exists
purely to make the previously-silent exclusion of blank/malformed rows
visible, and to surface a section_key/ds_code mismatch as a WARNING --
something no existing code path reports today.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

from openpyxl.worksheet.worksheet import Worksheet

import config
from excel_reader import ColumnMap, is_blank, extract_ds_code, GROUP_MARKER_TO_SECTION_KEY

SECTION_TITLES_BY_KEY = {}
for _s in list(config.OUTPUT_SECTIONS) + list(config.WORKSHEET2_ADDITIONAL_SECTIONS):
    SECTION_TITLES_BY_KEY[_s.key] = _s.title

DS_CODES_BY_SECTION = {
    s.key: s.ds_codes
    for s in list(config.OUTPUT_SECTIONS) + list(config.WORKSHEET2_ADDITIONAL_SECTIONS)
}
ALL_CONFIGURED_CODES = {c for codes in DS_CODES_BY_SECTION.values() for c in codes}


@dataclass
class RowValidation:
    row_index: int
    name: Optional[str]
    group: Optional[str]
    section_key: Optional[str]
    ds_code: Optional[int]
    status: str  # "VALID" | "WARNING" | "ERROR"
    reason: str


def validate_rows(ws: Worksheet, cmap: ColumnMap) -> List[RowValidation]:
    """Classify every physical data-shaped row on the sheet. Mirrors
    read_project_rows's own heading-scan exactly (same helper
    functions, same is_heading_candidate guard) so its section_key
    determination is always consistent with what the real pipeline
    sees -- this is a read-only, additional pass, not a replacement."""
    results: List[RowValidation] = []
    start = cmap.field_header_row + 1
    current_section = None

    for r in range(start, ws.max_row + 1):
        first_cell = ws.cell(r, cmap.name).value
        sub_group_val = ws.cell(r, cmap.sub_group).value if cmap.sub_group else None
        group_val_for_marker = ws.cell(r, cmap.group).value if cmap.group else None
        is_heading_candidate = cmap.sub_group is None or is_blank(sub_group_val)

        if is_heading_candidate:
            text = first_cell.strip().lower() if isinstance(first_cell, str) else ""
            matched_any = False

            # The Group column is the ONLY classification signal (same
            # rule as excel_reader.py's read_project_rows) -- heading
            # text is never inspected for classification.
            if group_val_for_marker in GROUP_MARKER_TO_SECTION_KEY:
                current_section = GROUP_MARKER_TO_SECTION_KEY[group_val_for_marker]
                matched_any = True

            # A heading-shaped row (blank Sub-Group) whose Group column
            # is non-blank but matches NO configured section's
            # group_marker is a hard ERROR, not a warning: today's
            # classifier (read_project_rows) silently leaves
            # current_section unchanged in this exact case, meaning
            # every row under this new/renamed/typo'd section gets
            # silently absorbed into whatever section came before it --
            # confirmed directly (a genuine "Track 5-Secured" section
            # added without updating config.py inflated the preceding
            # section's total by the new section's own revenue, with no
            # warning at all). Since the Group column is the sole
            # source of truth, an unrecognized non-blank value there is
            # unambiguous -- unlike heading TEXT (presentation-only,
            # the reason this used to be a soft, text-pattern-based
            # WARNING). Blank-Group decorative rows (most subtotal/
            # grand-total banners) are entirely unaffected, since this
            # only fires when Group is non-blank.
            if not matched_any and not is_blank(group_val_for_marker):
                results.append(RowValidation(
                    row_index=r, name=first_cell, group=group_val_for_marker,
                    section_key=current_section, ds_code=None,
                    status="ERROR",
                    reason=f"Unknown Group marker: {group_val_for_marker}",
                ))

            # A heading-shaped row with BLANK Group and non-decorative
            # heading text is a softer, informational case -- kept as
            # a WARNING since a blank Group value doesn't misclassify
            # anything (current_section simply doesn't change), it's
            # just a heading with no explicit section identity at all.
            decorative_keywords = ("subtotal", "sub-total", "total secured",
                                    "total prospecting", "total projection")
            decorative = bool(text) and (
                any(kw in text for kw in decorative_keywords)
                or text.strip() in ("secured", "projection", "prospecting")
            )
            if not matched_any and text and not decorative and is_blank(group_val_for_marker):
                results.append(RowValidation(
                    row_index=r, name=first_cell, group=None,
                    section_key=current_section, ds_code=None,
                    status="WARNING",
                    reason=(
                        f"Heading-shaped row (blank Sub-Group) with a blank Group "
                        f"column value -- rows below it will continue to be "
                        f"classified as {current_section!r} (whatever section was "
                        f"last recognized), not as a new section."
                    ),
                ))

            continue

        name_val = ws.cell(r, cmap.name).value
        group_val = ws.cell(r, cmap.group).value if cmap.group else None

        # Rows with a blank Sub-Group entirely (not heading-shaped,
        # just genuinely empty) never reach read_project_rows either;
        # skip them here too, silently -- truly blank spacer rows are
        # not worth reporting.
        if cmap.sub_group and is_blank(sub_group_val) is False:
            pass  # has a sub-group; fall through to field checks below
        elif cmap.sub_group and is_blank(sub_group_val):
            continue
        elif not cmap.sub_group and is_blank(group_val):
            continue

        ds_code = extract_ds_code(sub_group_val) if cmap.sub_group else None

        if is_blank(name_val) and is_blank(group_val):
            continue  # a genuine blank spacer row with only a stray Sub-Group tag -- not worth reporting as an error on its own

        if is_blank(name_val) or is_blank(group_val):
            missing = "Name" if is_blank(name_val) else "Group"
            results.append(RowValidation(
                row_index=r, name=name_val, group=group_val,
                section_key=current_section, ds_code=ds_code,
                status="ERROR",
                reason=f"{missing} is missing -- this row will never become a ProjectRow and will not appear anywhere in the output.",
            ))
            continue

        if current_section is None:
            results.append(RowValidation(
                row_index=r, name=name_val, group=group_val,
                section_key=None, ds_code=ds_code,
                status="ERROR",
                reason="Row has Name and Group but appears before any recognized section heading -- invalid section.",
            ))
            continue

        # From here on, the row WILL become a ProjectRow. Check revenue/margin.
        month_rev_cols = [4 + 3 * (m - 1) for m in range(1, 13)]
        month_marg_cols = [6 + 3 * (m - 1) for m in range(1, 13)]
        revenue = sum(v for c in month_rev_cols if isinstance((v := ws.cell(r, c).value), (int, float)))
        margin = sum(v for c in month_marg_cols if isinstance((v := ws.cell(r, c).value), (int, float)))

        warnings = []
        if abs(revenue) < 0.01 and abs(margin) < 0.01:
            warnings.append("Revenue and margin are both zero")
        if ds_code is not None and ds_code not in ALL_CONFIGURED_CODES:
            warnings.append(f"DS code {ds_code} does not match any configured section")
        elif ds_code is not None and current_section in DS_CODES_BY_SECTION:
            if ds_code not in DS_CODES_BY_SECTION[current_section]:
                warnings.append(
                    f"DS code {ds_code} does not match this row's own physical section "
                    f"({current_section}, expects {DS_CODES_BY_SECTION[current_section]}) -- "
                    f"the row is still counted here (section_key is authoritative for "
                    f"placement), but its Sub-Group tag looks like it may not have been "
                    f"updated after a copy-paste."
                )

        if warnings:
            results.append(RowValidation(
                row_index=r, name=name_val, group=group_val,
                section_key=current_section, ds_code=ds_code,
                status="WARNING", reason="; ".join(warnings),
            ))
        else:
            results.append(RowValidation(
                row_index=r, name=name_val, group=group_val,
                section_key=current_section, ds_code=ds_code,
                status="VALID", reason="Name, Group, section, and DS code all consistent; non-zero activity.",
            ))

    return results


def print_report(results: List[RowValidation]) -> None:
    print(f"{'Row':>5} | {'Customer':<30} | {'Status':<7} | Reason")
    print("-" * 100)
    for r in results:
        name = (r.name or "")[:30]
        print(f"{r.row_index:>5} | {name:<30} | {r.status:<7} | {r.reason}")
