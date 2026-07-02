# Sales Forecast Automation Engine (SFAE)

Automatically generates the monthly **Sales & Forecast Summary** workbook
from the **Master workbook** ("Sales by Customer" + "ClientComments" +
prior-year sheets), with zero hardcoded row numbers, column letters,
customer names, or sheet indexes.

```
python main.py --input "Master.xlsx"
```
produces `output/Sales_and_Forecast_Summary_<year>.xlsx` plus a
validation report and a full run log.

---

## 1. How this was built

Before writing any code, the Master workbook, the manually produced
Summary workbook, and the process-walkthrough recording were studied and
cross-referenced cell-by-cell to reverse-engineer every rule the human
follows -- not just the rules that were explicitly described. The
headline result of that forensic pass:

**The generated workbook matches the manually produced Summary** on
every Q1-Q4 figure, every Comment, every POC, and the overwhelming
majority of Margin and historical-year figures, across all 32 groups in
both output sections. An automated test (`tests/compare_with_manual.py`)
compares **every populated cell** between the generated workbook and the
golden manual file (tolerating sub-₹0.01 floating-point noise) and
currently reports **5 remaining material differences (15 cells,
several of which are subtotal rows that inherit a line-item difference)
out of several hundred compared** -- down from 76 found before the
historical-column mechanism below was reverse engineered. A further
round of investigation (prompted by follow-up feedback) rigorously
tested two additional, more sophisticated general fixes -- one for the
Databricks/Icertis-style Total disagreement, one for the Ovis ordering
swap -- and disproved both with concrete counter-examples before
reverting them, rather than shipping a heuristic that "seemed to work."
Every one of the 5 remaining differences is traced to a specific, cited
root cause in
[Known limitations](#6-known-limitations--assumptions); none of them are
unexplained.

### 1.1 Rules given directly in the brief (implemented literally)

| Rule | Implementation |
|---|---|
| 1. Group by `Group` column, one row per unique Group | `aggregator._group_rows_in_order` |
| 2. Q1=Jan+Feb+Mar, Q2=Apr+May+Jun, Q3=Jul+Aug+Sep, Q4=Oct+Nov+Dec | `config.QUARTER_MONTHS` + `aggregator.month_to_quarter` |
| 3. Total = Q1+Q2+Q3+Q4 | Written as a **live Excel formula** `=SUM(Q1:Q4)` on every row, exactly like the human-built file |
| 4. Margin = sum of every project's monthly Margin (Actual-Salary or the sheet's own Margin column) | `aggregator.aggregate_section` sums the existing per-month Margin column (Rule 4 explicitly allows this), and cross-checks the result against the sheet's own "Total Margin" column, flagging any mismatch in the validation report |
| 5. Comments matched from `<year>_ClientComments`, `Client List` -> `Group` | `comment_mapper.CommentMapper` |
| 6. Numeric blanks = 0, comment blanks = blank | `excel_reader.as_number` (blanks -> 0.0) and `aggregator.attach_comments` (no match -> `None`, left blank in the sheet) |
| 7. No hardcoded row numbers / column letters / customer or group names | Every sheet, header row, and column is located dynamically (see below); the only "static" knowledge is a small, documented, editable mapping of category **codes** in `config.py` (never a row number or a customer name) |

### 1.2 Rules that were *not* stated but were required to reproduce the real output

These were discovered by diffing the Master workbook against the sample
Summary workbook line-by-line and are documented here so a future
maintainer understands *why* the code does what it does:

1. **Sections are defined by the `Sub-Group` ("DS-code") column, not by
   scanning for bold text or blank rows.** Every genuine project row in
   the Master carries a `Sub-Group` value like `DS10_Secured`,
   `DS70_Secured`, `DS30_Projection`. This is far more robust than
   trying to detect "Track 1" / "Staffing" sections from formatting or
   blank-row heuristics, and it is what lets new rows be inserted
   anywhere in the sheet without breaking anything. The Summary only
   ever shows two blocks:
   - **Solutions and Staff Augmentation (Projects) - Track 1** <- `DS10`
   - **Staffing- Secured** <- `DS70`, `DS80`, `DS85`

   Every other code found in the data (`DS20`, `DS30_Projection`,
   `DS50_Projection`, `DS55_Projection`, `DS90_Secured` -- the
   partner/subcontractor staffing lines such as JSG, Mastek, Matchpoint,
   Maxonic, Technology Partners, VDart) is **outside the scope of this
   particular Summary report** and is deliberately excluded -- but never
   silently: every excluded code and its row count is listed under
   **"Unmapped Sub-Groups"** in the validation report so nothing
   disappears without a trace. This mapping lives in
   `config.OUTPUT_SECTIONS` and is the one place to touch if a brand-new
   business line/DS-code needs to be added to the Summary in future.

2. **"Skipped Blank Groups"** -- a group is dropped entirely when **both**
   its current-year revenue and its current-year margin are exactly
   zero (verified against the "Doppel" account, which is the one group
   the manual Summary drops -- every other zero-revenue group that was
   kept, e.g. "AI Forward"/"Bench", has a non-zero *margin* from pure
   overhead cost, so it's kept and shown with 0 revenue).

3. **Sort order**:
   - **Track 1 (Projects)**: groups with non-zero current-year revenue
     are sorted **alphabetically**; the handful of zero-revenue
     internal/overhead cost centres (AI Forward, Bench, Bench Deployed,
     the MetaSys/Ovis/Sales&Marketing allocations) are **not**
     alphabetised -- they are appended afterwards in the order they
     first appear in the source sheet, matching the sample exactly.
   - **Staffing-Secured**: kept in source order throughout (a small,
     stable list) rather than alphabetised.
   This is configurable per section via `OutputSection.sort_alphabetically`.

4. **Comments only match on an exact (Sub-Group code, Client name) pair.**
   The `ClientComments` sheet's own "Group" column is actually the
   section/DS-code (confusingly named the same as the main sheet's
   customer-name "Group" column) -- e.g. HPE has one comment filed under
   `DS10_Secured` and a *different* comment filed under `DS70_Secured`.
   A same-named client filed only under a different section's code (e.g.
   "IDBS" only has a `DS30_Projection` comment, no `DS10_Secured` one) is
   left blank rather than borrowed from the unrelated section -- verified
   against the sample, whose IDBS row has no comment despite the
   ClientComments sheet containing an IDBS entry.

5. **POC is only shown for the Staffing-Secured section**, even though
   the Track 1 sheet does carry a POC value on every row. This is a
   presentation choice baked into `OutputSection.show_poc`.

6. **Historical (prior-year) Total figures come from an embedded
   same-sheet reference column, not from recomputing the prior year's
   own sheet.** The current year's sheet carries two extra trailing
   columns beyond Group/Sub-Group -- a bare `2024`-style header and a
   `2025_Total`-style header -- that are a snapshot of that year's Total
   Revenue, entered directly on each project row. Summing that column
   across a group's own rows (same Sub-Group/DS-code scoping as
   everything else) reproduces the sample's 2024 column for **19/19**
   real Track-1 groups and its 2025 Total column for **16/19**, which is
   meaningfully better than independently recomputing from that year's
   own sheet (14/19 and 17/19 respectively) -- because the reference
   column is what the business actually captured at the time, and the
   underlying yearly sheets keep getting lightly revised afterwards.
   Margin has no equivalent embedded reference column anywhere in the
   workbook, so it is always recomputed fresh from that year's own sheet
   (Group + Sub-Group scoped, monthly columns summed directly rather
   than trusting that sheet's own cached "Total Margin" cell, which is
   occasionally left blank even when monthly data exists) -- this
   reproduces the sample's Margin column for **18/19** groups exactly
   (the 19th is off by ~$2, a rounding/timing artifact, see below).
   When the embedded reference column is entirely blank for a group
   (nothing to sum), the Total falls back to the fresh recomputation
   instead of reporting a false zero. When there is no data anywhere at
   all for a group in a given year, whether the Summary should show a
   blank cell or an explicit `0` turns out to depend on whether the
   group carries a Renewal Confidence value on any of its rows (real
   tracked accounts always have one, even "0%"; internal/overhead lines
   never do) -- this is `GroupSummary.has_renewal_confidence` in
   `aggregator.py`.

7. **Blank-row spacing is reproduced exactly, per section, not
   compressed to a single "always one blank line" rule.** The sample
   workbook is not internally consistent about this (Track 1 has a
   blank row between its heading and its title, but none between the
   title and its first data row and only a single blank row before its
   subtotal; Staffing-Secured has no heading, one blank row between its
   title and first data row, and *two* blank rows before its subtotal).
   Rather than silently "tidying up" that inconsistency, each section's
   exact spacing is captured as explicit fields on `OutputSection`
   (`blank_rows_after_heading`, `blank_rows_after_title`,
   `blank_rows_after_data`, `blank_rows_after_subtotal`) and reproduced
   row-for-row -- verified to match the sample workbook's row layout
   exactly, line by line.

---

## 2. Project structure

```
SalesForecastAutomation/
├── input/                 # Drop the Master workbook here (or pass --input)
├── output/                # Generated Summary workbook lands here
├── logs/                  # One run log + one validation report per execution
├── config.py              # ALL business-specific mapping/config lives here
├── excel_reader.py         # Dynamic sheet/header/column discovery, row extraction
├── aggregator.py           # Rules 1-4 & 6: grouping, quarters, margin, skip-blank, sorting
├── comment_mapper.py        # Rule 5: ClientComments matching
├── historical_lookup.py     # Prior-year Total/Margin lookups (2 strategies, see below)
├── summary_writer.py        # Builds the final formatted .xlsx
├── validator.py             # Validation report data model + renderer
├── main.py                  # CLI entry point / orchestration
├── requirements.txt
└── README.md                 # This file
```

Every module has a module-level docstring explaining its responsibility
and (where relevant) *why* a rule exists, not just what it does.

---

## 3. Usage

```bash
pip install -r requirements.txt

# Simplest: drop the Master workbook into input/ and run
python main.py

# Or be explicit
python main.py --input "/path/to/Master.xlsx" --output-dir "/path/to/out" --year 2026
```

`--year` is optional; it defaults to the most recent year found among
the workbook's `Sales by Customer- <year>` sheets, so **no code change
is needed next month** -- just drop in the new Master workbook (which
will now have a `Sales by Customer- 2027` sheet) and run the same
command.

Every run produces, inside `logs/`:
- `run_<timestamp>.log` -- full execution log (info/warnings/errors)
- `validation_report_<timestamp>.txt` -- the human-readable report also
  printed to the console, e.g.:

```
Workbook Loaded

[Section] Solutions and Staff Augmentation (Projects) - Track 1
  Groups Processed        : 28
  Comments Matched        : 27
  Missing Comments        : 1
  Skipped Blank Groups    : 1
    -> Doppel

[Section] Staffing- Secured
  Groups Processed        : 4
  Comments Matched        : 4
  Missing Comments        : 0
  Skipped Blank Groups    : 0

TOTALS
  Groups Processed        : 32
  Comments Matched        : 31
  Missing Comments        : 1
  Skipped Blank Groups    : 1

Unmapped Sub-Groups (excluded from Summary; see config.OUTPUT_SECTIONS):
    DS30_Projection: 6 row(s)
    DS50_Projection: 17 row(s)
    DS90_Secured: 6 row(s)

Output Workbook     : output/Sales_and_Forecast_Summary_2026.xlsx
Generation Successful
```

---

## 4. Error handling

- **Missing required sheet** (e.g. no `Sales by Customer- <year>` sheet
  can be found): raises a clear error naming the pattern it looked for
  and listing every sheet that *is* in the workbook.
- **Missing required column** (e.g. no `Group` column, or no monthly
  Actual/Forecast columns): raises a clear error naming exactly which
  column is missing and on which sheet.
- **Missing `ClientComments` sheet**: not fatal -- all comments are left
  blank and a warning is recorded.
- **Missing prior-year sheet**: not fatal -- that year's historical
  columns default to 0 and a warning is recorded.
- **No `Sub-Group` column at all** (a structurally different future
  workbook): the engine falls back to a single, un-sectioned Summary
  built purely from the `Group` column, and logs a warning that
  section-based scoping is degraded.
- Any other unexpected error is caught at the top level, logged with a
  full traceback in the log file, and reported in the validation report
  rather than crashing silently.

---

## 5. How column/sheet detection works (Rule 7)

- **Sheets** are found by regular expression + captured year, e.g.
  `sales\s*by\s*customer.*?(\d{4})` for the main sheet and
  `(\d{4}).*client\s*comments` for the comments sheet (which also skips
  backup/duplicate sheets like `2026_ClientComments (2)`).
- **Header rows** are found by scanning the first ~20 rows for the one
  that contains both "Name" and "Group" (normalised: lower-cased, every
  non-alphanumeric character stripped, so "Total Revenue", "Total
  (Revenue)", and "Total  Revenue" all compare equal) -- and, if more
  than one row qualifies, preferring whichever one actually has real
  Excel dates in its monthly columns (some sheets repeat a coarse label
  row above the real, date-bearing header row).
- **Month columns** are identified by an actual `datetime` value in the
  header row (not a fixed column letter), and their **role**
  (Actual/Forecast = revenue, Salary = cost, Margin = margin) is read
  from the row directly above.
- **Every other column** (POC, Service, Comments, Renewal Confidence,
  Total Revenue, Total Margin, Sub-Group) is located the same way, by
  normalised header text.

If columns are re-ordered, inserted, or renamed with cosmetic
differences next month, the engine keeps working. If a *required*
column genuinely disappears, it fails loudly with a specific message
rather than silently producing wrong numbers.

### 5.1 This was verified empirically, not just by reading the code

In response to a direct question about column-position independence,
every backend file was grepped for hardcoded column letters
(`ws['E']`-style access) and literal column-index calls
(`ws.cell(row, 5)`-style) -- zero matches anywhere in `config.py`,
`excel_reader.py`, `aggregator.py`, `comment_mapper.py`,
`historical_lookup.py`, `summary_writer.py`, or `validator.py`.

Then, to go beyond static analysis, `tests/test_column_reordering.py`
was built: it takes the real fixture Master workbook, shuffles **every
column on every sheet into a random new order** (headers moving
together with their data, as if someone had dragged columns around in
Excel), runs the full pipeline against both the original and the
scrambled copy, and asserts the two generated Summary workbooks are
cell-for-cell identical. It's a permanent test, not a one-off check --
re-run it after any change to the four modules above.

The **current year's main sheet** (and 2025, which still has a proper
`Group` column) produced a byte-identical Summary workbook on the first
try, across every seed -- confirming the claim for every column named
in the question (Group, Sub-Group, Comments, Total Revenue, Total
Margin, month columns, POC, Renewal Confidence, and the embedded
year-reference columns).

The **2024-style legacy sheet** did not, and it took two rounds to get
right -- both are recorded here because the process matters as much as
the result:

1. **First shuffle, first bug.** The 2024 sheet predates the
   Group/Sub-Group scheme entirely and is handled by a separate
   fuzzy-name-match fallback (`historical_lookup._load_fuzzy_sheet`),
   since it has no `Group` column to key off. That fallback had two
   real positional assumptions: it assumed the 12 monthly columns were
   "whichever 12 columns sit immediately before the Total column"
   instead of reading each one's own "Actual" role label one row above
   it, and it assumed the unlabeled project-Name column was "column B"
   whenever column A said "Status". Both were rewritten to be
   header-driven -- except the Name-column fix wasn't actually
   header-driven, just *less* hardcoded: it changed "always column B"
   to "always the column immediately after wherever Status ends up",
   which is still a positional assumption, just a relative one.
2. **Second shuffle (different seeds), second bug.** Re-running the
   test with several more random seeds caught exactly the case that
   first fix couldn't handle: a permutation where a *different* labelled
   column ("Service") happened to land immediately after "Status",
   causing the fallback to grab the wrong column and silently pull a
   number from it. The real fix was to stop trying to locate the
   nameless column *relative to* anything else and instead identify it
   by its own content: among every column with a blank header, the
   Name column is whichever one actually has data in it (a legacy sheet
   may have more than one blank-header column -- this file has a
   genuinely empty spacer column too -- but only the real Name column is
   densely populated with text below the header row).

After the content-based fix, `tests/test_column_reordering.py` was run
with **8 independent random seeds** and produced zero differences on
every one, and `tests/compare_with_manual.py` was re-run and confirmed
the exact same 15 pre-existing, documented differences (see section 6)
with no new ones introduced. The lesson embedded in the test suite
going forward: "less positional than before" was not the same claim as
"not positional at all", and only re-running the empirical test (not
re-reading the fix) was what told the two apart.

This narrow exception only affects the single oldest historical
reference year, and only that year's own internal fallback logic (used
solely because that one sheet has no `Group` column and one genuinely
unlabeled header to begin with) -- it never affected the current year's
Summary calculations, which were column-order-independent from the
start.

---

## 6. Known limitations & assumptions

`tests/compare_with_manual.py` compares every populated cell of the
generated workbook against the golden manual file, with a ₹0.01
tolerance for floating-point noise (see the note on category 3 below).
It currently reports **5 remaining material differences** (in 15 cells,
since several propagate into subtotal rows) -- down from 76 found before
this project's historical-column mechanism was reverse engineered, and
further reduced from 17 after two additional rounds of investigation
below. Every one has been individually traced to a specific, cited cause
in the source workbook; two rejected hypotheses are documented alongside
the fixes that *did* work, precisely so the next person doesn't have to
re-run the same experiments.

**1. A data-entry typo in the source workbook (1 cell, D10).**
`Churchs-Chicken`'s embedded 2025 reference cell (`AV19` on the
`Sales by Customer- 2026` sheet) reads `75000`, exactly 10x the figure
the sample Summary shows (`7500`) and 10x its own `2024` reference cell
on the same row (`AU19 = 7500`). An independent recomputation from the
2025 sheet gives a *third* number again (`9000`), and the value `7500`
itself is nowhere else in the entire workbook (checked exhaustively,
every sheet, every cell). Three independent, mutually-inconsistent
numbers rules out a scoping bug on our end -- the source cell is a
data-entry error, and the correct figure isn't recoverable from this
file.

**2. A stale embedded reference (1 cell, D11) -- investigated with a
disproved fix.** `Databricks`'s embedded 2025 Total reference is
`1,817,631`, but summing that year's own sheet gives `1,854,931`, which
matches the sample exactly and is corroborated by the recomputed Margin
for the same rows (`732,194.62`) matching the sample's Margin to the
penny. This looked fixable: the sheet also carries an embedded
`2025_Q4`-style reference, and a fresh Q4-only recomputation for
Databricks agrees with it exactly -- suggesting the underlying sheet
hasn't moved since the references were captured, so the *full-year*
reference (not the recomputation) must be the one with the error.
**This rule was implemented, then disproved by a counter-example:**
`Icertis` has the identical signature (Q4 reference and Q4
recomputation agree; full-year reference and full-year recomputation
disagree, by $300) -- but for Icertis, the sample workbook agrees with
the embedded *reference*, not the recomputation, the opposite of
Databricks. With only one Q4 checkpoint available, both "an earlier
quarter was corrected upward after the snapshot" (Databricks) and "an
earlier quarter was corrected downward, or the reference itself is
simply right" (Icertis) are equally consistent with the same observed
signature -- there is no way to tell them apart from the data in this
workbook. The Q4 cross-check is kept as a **diagnostic only**: it
annotates "Historical Reference vs Recompute Drift" entries in the
validation report with `[Q4 corroborates recompute]` or `[Q4 also
drifts]`, but it never changes which value gets written, because doing
so would silently fix one cell at the cost of silently breaking another.

**3. Sub-₹0.01 floating-point noise, tolerated by the test.** Two Margin
cells (HAL, Hitachi Asia) differ from the sample by exactly $0.01 --
Python's IEEE754 summation of the same monthly cells Excel itself cached
lands a cent off due to binary floating-point representation, not a
logic error (verified: our recomputation matches Excel's own cached
"Total Margin" column to the same cent in both directions). The
comparison test tolerates differences of ₹0.015 or less for exactly
this reason, per this round of feedback.

**4. Confirmed genuine data drift on 3 Margin cells (HPE Track-1,
HPE-Staffing, Marvell India -- $2.33/$12.34/$1.04).** Cross-checking
each group's embedded `2025_Q4` reference against a fresh Q4-only
recomputation shows the underlying monthly data for these three groups
**has** been revised since the reference snapshot was taken (Q4 alone
disagrees by amounts on the same order as the full-year Margin gap).
Recomputing Margin fresh from the current monthly cells is the only
mechanism available (no embedded Margin reference exists anywhere in
the workbook), and it is internally consistent with the sheet as it
exists today -- these few-dollar gaps are the cost of that sheet having
been lightly edited after the sample Summary was produced, not a
reproducible defect.

**5. One irreducible margin gap (1 cell, E41).** `AZU Solutions`'s 2025
Margin is `7,558` in the sample; recomputing from the 2025 sheet's own
row gives `11,144`, and no cell anywhere in the entire workbook contains
`7,558` (exhaustive search). The same Q4 cross-check used in category 4
confirms AZU's Q4 revenue has also drifted by exactly `672` relative to
its embedded Q4 reference -- independent proof that this account's row
has been materially edited since the sample was produced, not a
computation bug.

**6. One manual row-ordering swap (4 cells, A31/A32/K31/K32) --
investigated with two disproved hypotheses.** "Ovis- Creative Support"
and "Ovis Programming + MetaSys Websites" are swapped in the sample
relative to every signal in the source data: the raw sheet's row order
and the `ClientComments` sheet's row order agree with each other (and
with this tool's current output) that Programming comes first. Two
alternative sort rules were tested against all 9 non-billable/overhead
groups to see if either explains the *whole* block, not just this pair:
  - *Sort by Margin (ascending distance from zero)*: correctly orders
    the Ovis pair, but puts "Bench" (the most negative Margin) second
    from last instead of second overall -- wrong for 7 of 9 groups.
  - *Locale/punctuation-insensitive alphabetical sort*: correctly orders
    the Ovis pair AND "AI Forward"/"Bench"/"Bench Deployed", but orders
    the three "MetaSys..." accounts as Marlow, Payroll, Staffing
    (M < P < S) where the sample has Payroll, Staffing, Marlow -- wrong
    for that trio.

  Both were rejected because they only relocate the mismatch rather than
  resolving it. The current source-row-order rule remains the best
  overall predictor (8 of 9 non-billable groups correct, vs. 6-7 of 9
  for either alternative), so it is kept, and the Ovis pair is recorded
  as an unexplained one-off edit in the sample file.

**Plus 4 subtotal cells that are pure arithmetic consequences of the
above** (`Subtotal : Track 1` D/E/K, `Subtotal : Staffing- Secured` E)
-- not separate bugs; they will match automatically if the corresponding
line-item figures above are ever corrected.

None of these cells are guessed, hardcoded, or silently patched -- two
plausible general fixes were implemented and then reverted in this file
history specifically because they were disproved by a counter-example,
rather than kept because they "seemed to work." If the business can
confirm the correct figure for any of the source-data issues above
(categories 1, 2, or 5), record it once in `config.HISTORICAL_OVERRIDES`
with a comment citing who confirmed it and when -- the mechanism is
documented and empty by default specifically so this tool never
fabricates a number on its own.

**Also worth noting**, from earlier iterations of this reverse-engineering:

- **A legacy year with no `Group` column at all** (e.g. a 2024-style
  sheet). The engine falls back to a case-insensitive substring match
  against the free-text project name, summing that row's monthly
  columns. This works cleanly for most accounts (verified against the
  sample) but can't always bridge a bigger naming gap (e.g. "HPE" vs.
  "Hewlett Packard Enterprise Company" in a 2024-style sheet), and is
  guarded against falsely matching a stray section-label row (e.g. a
  bare "Staffing" row) by requiring the candidate row to carry actual
  monthly data before it counts as a match.
- **`Sub-Group` code mapping is a config allowlist, not a formula.**
  `config.OUTPUT_SECTIONS` encodes which numeric DS-codes belong in
  which block of the Summary. This was reverse-engineered from the one
  month of data available and is the one piece of "business knowledge"
  a human may need to extend by hand (one line in `config.py`) if a
  brand-new code appears in a future month. The validation report's
  "Unmapped Sub-Groups" section exists specifically to make that moment
  obvious rather than a silent data-loss bug.

---

## 7. Automated regression test

```bash
python tests/compare_with_manual.py
```

Regenerates the Summary from the fixture Master workbook
(`tests/fixtures/master_2026.xlsx`) and compares every populated cell
against the golden manual file
(`tests/fixtures/manual_summary_2026.xlsx`), printing every mismatch
(sheet coordinate, row label, expected vs. actual) and exiting non-zero
if any are found. It has no dependency on LibreOffice or pytest -- a
small built-in evaluator resolves the project's own `=SUM(...)` range
formulas (including the nested case where a subtotal row sums other
formula cells), so the comparison works on the raw generated `.xlsx`
without needing to open it in a real spreadsheet engine first.

Re-run this test after any change to `config.py`, `aggregator.py`, or
`historical_lookup.py` to catch a regression immediately.

```bash
python tests/test_column_reordering.py
```

Rebuilds the fixture Master workbook with every column on every sheet
shuffled into a random new order (across 8 different random seeds) and
asserts the generated Summary is cell-for-cell identical to the one
generated from the unshuffled original. This is what actually caught
the two positional-assumption bugs described in section 5.1 -- run it
after any change to `excel_reader.py`, `aggregator.py`,
`comment_mapper.py`, or `historical_lookup.py`, since none of those
bugs would have been visible from reading the code or from
`compare_with_manual.py` alone (that test only ever exercises the
Master workbook's *original* column order).

---

## 8. Extending the tool


- **New year**: nothing to change -- run it, `--year` auto-detects.
- **New Summary section / new DS-code family**: add an `OutputSection`
  entry to `config.OUTPUT_SECTIONS`.
- **Show 3 years of history instead of 2**: change
  `config.NUM_PRIOR_YEARS_SHOWN`.
- **Blank-row spacing needs to change**: adjust the relevant
  `blank_rows_after_*` field on that section's `OutputSection` in
  `config.py` -- nothing in `summary_writer.py` needs to change.
- **Column header renamed/re-ordered**: nothing to change, it's found
  dynamically -- unless the *meaning* changes, in which case update the
  matcher in `excel_reader.build_column_map`.

---

## 9. Desktop GUI

A branded, point-and-click desktop app (`gui_main.py`) sits on top of
the exact same CLI backend -- it changes nothing about how the Summary
is calculated. See `gui/runner.py`'s module docstring for exactly how it
adapts `main.py`'s orchestration into a GUI-friendly form: it imports
and calls `config`, `excel_reader`, `aggregator`, `comment_mapper`,
`historical_lookup`, `summary_writer`, and `validator` completely
unmodified. `main.py` itself is untouched and remains fully usable from
the command line.

### 9.1 Running it

```bash
pip install -r requirements.txt
python gui_main.py
```

The window walks through three steps -- pick the Master workbook, pick
an output folder (defaults to `~/Documents/NVISH Sales Summary Output`),
confirm the auto-detected target year (an editable dropdown -- type any
year, even one not found in the workbook) -- then **Generate Summary**.
Generation runs on a background thread so the window never freezes; an
indeterminate progress bar and a live status line (sourced from the same
step-by-step messages `main.py` logs) track what's happening. When it
finishes, the Validation Summary panel shows the pass/fail banner, the
key counts (Groups Processed, Comments Matched, Missing Comments,
Skipped Blank), and the full validation report text -- and three buttons
let you jump straight to the generated workbook, the validation report,
or the output folder. Any failure (bad file, missing sheet/column, a
workbook already open in Excel, or a genuine bug) is caught and shown as
a readable dialog with an optional, copyable "Technical Details" panel
-- never a raw console traceback.

### 9.2 Package structure

```
gui/
├── __init__.py
├── app.py         # the window: layout, event handlers, threading/queue glue
├── runner.py      # GUI-friendly orchestration wrapper around the unmodified backend
├── dialogs.py     # the error dialog (message + collapsible technical details)
├── branding.py    # NVISH colour palette + PyInstaller-safe asset path resolution
└── assets/
    ├── generate_assets.py  # one-off script that drew the files below
    ├── nvish_logo.png      # in-app header logo
    └── nvish_icon.ico      # window/taskbar/exe icon
gui_main.py            # entry point (`python gui_main.py`); also the PyInstaller target
SalesForecastGUI.spec  # PyInstaller build spec
build_exe.bat          # one-click Windows build script
```

### 9.3 Building a standalone Windows `.exe`

On a Windows machine with Python 3.10+ installed:

```bat
build_exe.bat
```

or manually:

```bash
pip install -r requirements.txt
pyinstaller --noconfirm --clean SalesForecastGUI.spec
```

The finished, single-file, windowed executable is written to
`dist\NVISH Sales Forecast Automation.exe` -- no console window, NVISH
icon, everything (including the logo/icon assets) bundled inside.

**This spec was actually built and smoke-tested** in this project's Linux
CI-style environment (PyInstaller builds for whichever OS it runs on;
building it here produces a Linux binary, not a `.exe`, but it exercises
the identical spec file, import graph, and asset-bundling path that a
Windows build would). That test caught and fixed one real, easy-to-miss
issue documented in `SalesForecastGUI.spec`: PyInstaller's static
analysis doesn't always detect `PIL._tkinter_finder` (a conditionally
imported module `PIL.ImageTk` needs to render *any* image, including
ttkbootstrap's own combobox arrow icon), so it's listed explicitly as a
hidden import. A separate, isolated smoke test of the frozen backend
(bypassing the GUI entirely) confirmed `openpyxl` and every backend
module bundle and run correctly end-to-end, producing byte-identical
output to the unfrozen CLI.

### 9.4 Why the output folder is always explicit

`config.py` creates `input/`, `output/`, and `logs/` next to wherever
its own file lives. That's fine when running from source, but inside a
PyInstaller *onefile* build, the running copy of `config.py` lives in a
temporary extraction folder that is deleted when the app closes -- so
anything written there would vanish. The GUI never relies on those
defaults: it always asks the user for an explicit output folder (a real,
persistent location) and saves both the workbook and the validation
report there directly, via the same public `wb.save(...)` and
`report.save(...)` calls the CLI uses -- no change to `config.py`,
`summary_writer.py`, or `validator.py` was needed to make this safe.


## 10. Streamlit web app

A third interface (`app.py`) sits on top of the exact same, unmodified
calculation engine -- styled and structured after NVISH's SFAE v2
Streamlit application (reused for its front end only: layout, sidebar,
navigation, CSS design system, cards, badges, upload widgets, and
progress indicators -- **none** of SFAE v2's forecasting, staffing,
placement, or RFQ business logic was brought over). Like the desktop
GUI, it changes nothing about how the Summary is calculated.

### 10.1 How it reuses the backend

`streamlit_bridge.py` is the **only** new code that connects Streamlit
to the calculation engine, and it is intentionally thin:

- For the actual generation, it re-exports `generate_summary`,
  `GenerationResult`, and `GenerationError` **directly from
  `gui.runner`** -- the exact same, unmodified orchestration wrapper the
  desktop app already uses. It is not reimplemented or duplicated.
- The only genuinely new behaviour is `preview_workbook()`, used for the
  "once uploaded" panel (Detected Year / Detected Sheets / Number of
  Groups) shown *before* the user clicks Generate. Even this does not
  reimplement any calculation: it calls the same `aggregate_section()`
  function the real run uses, so the previewed "Number of Groups" is
  guaranteed to match the real run's "Groups Processed" count, not an
  approximation of it.

Every one of `config.py`, `excel_reader.py`, `aggregator.py`,
`comment_mapper.py`, `historical_lookup.py`, `summary_writer.py`, and
`validator.py` is imported and called completely unmodified -- verified
by running `tests/compare_with_manual.py` (unchanged) against this
project and confirming the exact same 15 cells differ, in the exact
same way, as before the Streamlit front end was added.

### 10.2 Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Workflow: **Upload & Generate** (sidebar) -> upload the Master workbook
under **Data Source** -> once validated, a **Workbook Loaded** panel
shows the auto-detected year (editable -- a number input seeded with the
detected value, since Streamlit has no native "editable combobox"),
the sheets that will be read, and the number of Summary groups that
will result -> **Generate Summary** runs the exact same pipeline as the
CLI/desktop app, with a live checklist (Loading workbook... / Reading
rows... / Grouping projects... / Calculating quarters... / Calculating
margins... / Generating summary... / Finished.) driven by the real
`progress_cb` callbacks `gui.runner.generate_summary` already emits (see
`app.py`'s `_StepDriver`, which maps those real messages onto the
checklist labels, advancing only as genuine backend milestones occur --
never fabricating progress that hasn't happened) -> a **Validation
Summary** panel shows Groups Processed, Comments Matched, Warnings, and
Time Taken, with buttons to download the Summary workbook, download the
validation report, and (when running locally) open the output folder.

### 10.3 Project structure

```
app.py                    # main Streamlit entry point (streamlit run app.py)
streamlit_bridge.py         # the ONLY new backend-facing code -- see 10.1
components/
├── __init__.py
├── ui.py                   # presentation-only: cards, badges, KPI tiles, process log
└── sidebar.py                # NVISH sidebar: brand, nav, workbook status
styles/
└── enterprise.css             # design system CSS, adapted from SFAE v2 (colours/
                                # cards/badges/sidebar mechanics unchanged; only
                                # product text and two small additions -- pending/
                                # active checklist states, a stat-tile style)
.streamlit/
└── config.toml                 # theme (same palette as enterprise.css)
assets/
└── nvish_logo.png               # NVISH's real logo (topbar; the sidebar uses a
                                  # compact text mark instead, since the logo's dark
                                  # wordmark isn't legible on the dark sidebar)
```

### 10.4 Notes on adapting the SFAE v2 UI

- **Streamlit version compatibility.** SFAE v2's CSS targets Streamlit's
  DOM via `data-testid` attributes. Testing this app against a current
  Streamlit release (1.58) surfaced that one testid used throughout the
  original CSS (`element-container`) had been renamed to
  `stElementContainer`, and `.stDeployButton` to
  `stAppDeployButton` -- both fixed by adding the current selector
  alongside the legacy one, so the CSS works across versions rather than
  silently no-opping on newer Streamlit releases.
- **A real, separately-diagnosed layout bug**, also only visible by
  actually rendering the app: Streamlit applies a small built-in
  negative bottom margin to every markdown container as part of its own
  default vertical rhythm. It's invisible in SFAE v2's original pages
  because their sections are separated by real widgets/spacers, but this
  app's compact sidebar status rows (raw HTML with no widgets between
  them) exposed it as visibly overlapping text. Fixed with one CSS
  reset (`[data-testid="stMarkdownContainer"] { margin: 0 !important; }`)
  -- documented in `styles/enterprise.css` at the point of the fix.
- **State-resolution ordering.** `app.py` resolves all upload/preview
  state (`_resolve_upload_state()`) *before* the sidebar and page render,
  not inside the page function -- otherwise the sidebar (rendered first)
  would always show the previous run's state, one step behind the main
  panel. This is a Streamlit-specific architectural detail, not a
  calculation concern.
- **"Open Output Folder"** only works when Streamlit is running on the
  same machine as the browser viewing it (a local run) -- there is no
  meaningful equivalent for a remotely hosted deployment, which the
  button's caption notes explicitly, matching the brief's
  "(if supported)".
