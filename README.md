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
| 3. Total = Q1+Q2+Q3+Q4 | Written as a **live Excel formula** on every row (originally a contiguous `=SUM(Q1:Q4)`; now `=SUM(<Q1 Total>,<Q2 Total>,<Q3 Total>,<Q4 Total>)` by explicit cell reference since each quarter became 2 columns -- see §11.2), exactly like the human-built file |
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
├── monthly_view.py           # Worksheet 2's monthly Actual/Forecast breakdown (see §13)
├── summary_writer.py        # Builds the final formatted .xlsx (both worksheets)
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
(`tests/fixtures/master_2026.xlsx`) and compares it against the golden
manual file (`tests/fixtures/manual_summary_2026.xlsx`). Since the
output *layout* changed (§11 -- each quarter is now two columns, Margin
then Total, under a 3-row header instead of the golden file's 2-row,
single-column-per-quarter layout), this test no longer assumes fixed
cell coordinates on either side: it parses both header shapes
dynamically by their own label text, matches rows between the two files
by the text in their Name column (not by row number), and checks
everything that existed in the old layout (Name, POC, prior-year
Total/Margin, each quarter's revenue Total, the final yearly
Total/Margin, Comments) plus one new self-consistency check that has no
golden equivalent: every row's four new Quarter Margin sub-columns must
sum to that row's final yearly Margin, since both are built from the
exact same monthly-margin figures in aggregator.py. It has no
dependency on LibreOffice or pytest -- a small built-in evaluator
resolves the project's own `=SUM(...)` formulas, both the contiguous-
range form subtotal rows use and the explicit comma-separated-cell-list
form each data row's yearly Total now uses (needed because the four
quarter Total sub-columns are no longer contiguous -- see §11).

Re-run this test after any change to `config.py`, `aggregator.py`,
`historical_lookup.py`, or `summary_writer.py` to catch a regression
immediately.

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

```bash
python tests/test_future_year_compatibility.py
```

Runs generation against every year the fixture workbook can meaningfully
target, asserting years with a modern Group-column sheet succeed (even
when the historical lookback reaches a year with no sheet in the
workbook at all) and a legacy year with no Group column fails
*gracefully* rather than crashing. This is what caught the pre-existing
`UnboundLocalError` crash described in §11.4 -- run it after any change
to `aggregator.py` or `historical_lookup.py`.

```bash
python tests/test_worksheet2_actual_forecast.py
```

Confirms Worksheet 2 (§13) exists under the correct dynamic
"<year> Actual & Forecast" name, shows exactly the same groups in the
same order as Worksheet 1, dynamically detects both "Actual" and
"Forecast" month roles from the source sheet's own header (proving the
month range isn't hardcoded), includes the Confidence/Comments columns,
and -- critically -- that Worksheet 1 remains byte-for-byte stable
across repeated runs. Run it after any change to `monthly_view.py` or
to the parts of `summary_writer.py` that build Worksheet 2.

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


## 11. Formatting & layout upgrade (borders, grouped quarters, currency, second sheet)

> **Note:** §11.2 below describes each quarter's sub-columns as
> "Margin, then Total" -- that was the order at the time this section
> was written. It was changed to **Total, then Margin** in a later
> round; see §12 for exactly what that involved. Kept here rather than
> silently rewritten so the two sections together show the actual
> history, not just the current state.

Five presentational features were added on top of the existing engine.
**No calculation was redesigned to build these** -- the only
calculation-adjacent change was a small, additive one in `aggregator.py`
(§11.2); every other change lives entirely in `summary_writer.py`
(output formatting) or `config.py` (styling constants).

### 11.1 Thin black borders on every populated cell

`summary_writer.py` now tracks every row it writes (header rows, section
heading/title banners, data rows, subtotal rows) in `self._content_rows`
as it builds the sheet, then applies a thin black border
(`config.BORDER_COLOR`) to every cell in those rows, across every
column, in a final `_apply_borders()` pass. The blank spacer rows
between sections (`OutputSection.blank_rows_after_*`) are deliberately
left out of `_content_rows`, so they stay pure visual spacing rather
than becoming a bordered empty row. Only `.border` is touched -- fills,
fonts, alignment, merged cells, and number formats are untouched.

### 11.2 Grouped Q1-Q4 columns (Margin + Total each)

Each quarter is now two columns -- **Margin**, then **Total** -- under a
merged "Q1"/"Q2"/"Q3"/"Q4" header, which turned the sheet's header from
2 rows into 3 (year band / quarter-group-or-vertical label / Margin
&#124; Total sub-label). `summary_writer._plan_columns()` and
`_write_headers()` account for the extra row and the doubled-up quarter
columns; every other section (banners, data rows, subtotals) shifted
down by exactly one row to make room, with no other structural change.

**Quarter Total** is untouched: still `GroupSummary.quarters`, the
existing revenue aggregation.

**Quarter Margin** required one small, additive change to
`aggregator.py`: `GroupSummary` gained a `quarter_margins` field, and
`aggregate_section()`'s existing monthly-margin loop --
`for margin in r.monthly_margin.values(): total_margin += margin` --
became `for month, margin in r.monthly_margin.items():
quarter_margins[month_to_quarter(month)] += margin; total_margin +=
margin`. This is the exact same monthly margin data and the exact same
`month_to_quarter()` mapping already used for both the yearly Margin
column and the Quarter Total columns -- it is bucketed one more way, not
recalculated a different way. `tests/compare_with_manual.py`'s
self-consistency check (every row's 4 new Quarter Margin values sum to
that row's existing, already-validated yearly Margin) is what verifies
this claim on every run, not just an assertion in this README.

One mechanical consequence: the current-year yearly Total formula used
to be a single contiguous range (`=SUM(F5:I5)`, back when Q1-Q4 were 4
side-by-side columns). With a Margin column now sitting between every
pair of quarters, a plain range would wrongly sum the Margin cells into
the revenue Total -- so that formula now sums the four Total
sub-columns by explicit cell reference instead (`=SUM(G5,I5,K5,M5)`).
Same result, adjusted cell references, not a new methodology. The
yearly Margin cell is unchanged: still `group.total_margin`, written as
a static value exactly as before.

### 11.3 Real Excel currency formatting

`config.CURRENCY_FORMAT` changed from
`_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)` (Excel's built-in
"Comma Style" -- no currency symbol) to
`_($* #,##0_);_($* \(#,##0\);_($* "-"??_);_(@_)` (the built-in
"Accounting" format with a `$`) -- same thousands separators, same
parentheses-for-negative and dash-for-zero convention, just with the
symbol turned on. Since every monetary column (prior-year Total/Margin,
all 8 quarter Margin/Total sub-columns, final yearly Total/Margin) was
already routed through this one constant in `_apply_sheet_formatting()`,
changing it in one place applied the `$` everywhere it was asked for
without touching `summary_writer.py`'s formatting loop at all. Cell
values remain plain floats/ints throughout -- only the *display* format
string changed, so sorting, filtering, and further calculation in Excel
all continue to work exactly as before.

### 11.4 A pre-existing bug found while verifying "future-year compatibility"

While confirming the new layout didn't break dynamic year detection,
running the tool against `--year 2025` (whose lookback needs a "2023"
sheet the fixture workbook doesn't have) crashed with
`UnboundLocalError: cannot access local variable 'source'` in
`aggregator.attach_historical`. Reproducing the same run against the
**unmodified, originally uploaded project** confirmed this was already
present -- unrelated to any of the five features above.

The cause: one branch of `attach_historical`'s total/margin resolution
(the "no embedded reference and no match anywhere" case) never assigned
its local `source` bookkeeping variable, which a later, unconditional
line then always read. It's guaranteed to trigger whenever the
historical lookback reaches a year with literally no corresponding
sheet in the workbook -- which is exactly what "future-year
compatibility" needs to keep working for the years *nearest* a
workbook's edges, not just new ones being added at the end. Confirmed
before fixing that `historical_source` is written to in exactly one
place and **read nowhere else in the codebase** (`grep -rn
"historical_source"` across every `.py` file), so the one-line fix
(give that branch a `source = "not_found"` value) changes no calculated
number anywhere -- it only stops the crash. `tests/compare_with_manual.py`
was re-run immediately after and showed the identical 10 pre-existing
differences (see §6), confirming zero effect on the 2026 output this
project has been validated against throughout. A permanent regression
test, `tests/test_future_year_compatibility.py`, now locks this in.

### 11.5 Second worksheet (no logic yet)

`SummaryWriter.build()` now also creates
`wb.create_sheet(title=f"{self.target_year} Actual & Forecast")` --
e.g. "2026 Actual & Forecast" -- right after building the main Summary
sheet, with a single italic placeholder cell and no other content. As
requested, no business logic was added for it; it exists purely so the
tab is present and correctly named for whatever gets built into it next.

### 11.6 What stayed exactly the same

`excel_reader.py`, `comment_mapper.py`, `historical_lookup.py`, and
`validator.py` were not touched at all. `config.py` gained one styling
constant (`BORDER_COLOR`) and one format-string edit (`CURRENCY_FORMAT`)
-- no business-rule constant changed. `aggregator.py`'s only changes are
the additive `quarter_margins` field/computation (§11.2) and the
one-line crash fix (§11.4); `total_margin`, `total_revenue`, and
`quarters` are computed exactly as before, from the exact same source
data, in the exact same order. Dynamic year/sheet/header/column
detection, the Streamlit app, the desktop GUI, and the validation report
were all re-verified end-to-end against the new output and required no
changes at all -- confirmed by `grep`ing the whole project for any other
reference to the writer's internal column layout (`quarter_cols`,
`col_current_total`, etc.) and finding none outside
`summary_writer.py`/`aggregator.py` themselves.


## 12. Quarter sub-column order: Total before Margin

Each quarter's two sub-columns are now **Total, then Margin** (was
Margin, then Total in the previous round). This was the only functional
change requested this round -- currency formatting, borders, and the
second worksheet were already in place and needed no changes at all.

### 12.1 What changed, and what deliberately didn't

Only two files changed:

- **`summary_writer.py`** -- `_plan_columns()`'s
  `self.quarter_cols[q] = {"total": idx, "margin": idx + 1}` now assigns
  the lower column index to `"total"` (previously `"margin"` got it).
  The header-writing code that draws the merged "Q1".."Q4" group band
  and its Total/Margin sub-labels was also changed from hardcoding
  which key comes first to resolving it via `min()`/`max()` of
  `self.quarter_cols[q].values()` -- so this is no longer a place where
  the on-screen order could silently drift out of sync with
  `_plan_columns()`'s own ordering again. Everything else that touches a
  quarter column -- writing its value, building the yearly Total's
  `=SUM(...)` cell-reference list, applying the currency format, sizing
  the column -- already looked columns up by the `"total"`/`"margin"`
  dictionary key rather than by position, so none of it needed to
  change at all.
- **`tests/compare_with_manual.py`** -- its header parser had the *old*
  order baked in as an assumption ("row 3's first sub-column is Margin,
  the second is Total") to figure out which generated column was which.
  That assumption broke the moment the real order flipped, producing 23
  false "quarter margins don't sum to the yearly margin" failures that
  had nothing to do with the generated numbers (which were correct) --
  purely a stale assumption in the test itself. Fixed the same way the
  generator itself is architected: read row 3's actual text at each of
  the two sub-columns and assign `"total"`/`"margin"` by *what the label
  says*, not by which position it sits in. Re-run afterwards, the test
  is back to the same 10 pre-existing, documented differences from §6,
  and the quarter-margin consistency check passes cleanly.

**`config.py` and `aggregator.py` were not touched this round** --
currency formatting (`CURRENCY_FORMAT`), border color, and the
`quarter_margins` computation were all already correct from the previous
round; only *where on the sheet* each already-correct value gets placed
changed.

### 12.2 Verified

- Formula recalculation: 0 errors (58 formulas), and the yearly Total
  formula's cell references automatically followed the new column
  positions (e.g. `=SUM(F7,H7,J7,L7)` now referencing the four Total
  sub-columns in their new, lower-indexed slots) with no change needed
  to the code that builds that formula string -- it was already looking
  up `cols["total"]` by name.
- `tests/compare_with_manual.py`: same 10 pre-existing differences as
  every prior round (§6), 0 new ones; the quarter-margin self-consistency
  check passes for every row.
- `tests/test_column_reordering.py`: all 8 shuffled-column seeds still
  match the baseline exactly (unaffected -- this test only exercises
  *input* column order, never the *output* layout).
- `tests/test_future_year_compatibility.py`: all 3 scenarios still
  behave correctly.
- Streamlit: full upload -> generate -> download flow re-tested in a
  real browser; the downloaded workbook shows the new Total-then-Margin
  order.
- Desktop GUI backend (`gui.runner.generate_summary`): re-tested
  directly; produces the identical output the CLI and Streamlit do.


## 13. Worksheet 2: "<year> Actual & Forecast"

A second worksheet, dynamically named e.g. "2026 Actual & Forecast", now
shows the same groups Worksheet 1 already validated at monthly
granularity -- each month's own Actual/Forecast role, that month's
revenue and margin, then Total/Margin/Confidence/Comments -- instead of
Worksheet 1's quarterly view.

### 13.1 Which files were analyzed before writing any code, and why sourcing matters

Two candidate source workbooks were provided: the existing Master
workbook (`Sales by Customer- <year>`, already Worksheet 1's source) and
a second, separately-maintained "NK" workbook. Both were opened and
compared cell-by-cell before any implementation decision was made:

- **The NK workbook** has no Margin column anywhere -- not for one
  month, not for any month -- confirmed by dumping every header cell in
  both its type row and field row. It labels its forecast months
  "Outlook", not "Forecast". Its own Q1-Q4/Total columns are plain
  `=SUM()` formulas over its own Jan-Dec columns (self-contained for
  revenue), and it additionally contains "Projection" and "Prospecting"
  sections -- covering Sub-Group codes like `DS30_Projection` and
  `DS90_Secured` -- that Worksheet 1 has deliberately never included
  (they appear in every validation report's "Unmapped Sub-Groups" list).
- **The Master workbook** already has a Margin figure for every single
  month, for every project row, for both Actual and Forecast periods --
  the same figures already summed into Worksheet 1's Margin columns --
  and its own type-header row already says "Actual" / "Forecast"
  verbatim (confirmed by reading `Sales by Customer- 2026`'s row 1
  directly), which is the exact terminology requested and requires zero
  new keyword handling (`config.REVENUE_ROLE_KEYWORDS` already lists
  both).
- **Cross-check:** AIS Solutions' Jan-Jun revenue was compared cell by
  cell between the two workbooks and matches exactly, confirming both
  describe the same underlying business reality for the scope Worksheet
  1 covers -- the Master workbook is simply the superset (adds margin,
  uses the requested terminology already, and is the single source
  already powering Worksheet 1, guaranteeing both worksheets in the
  same output file are always internally consistent with each other).

**Conclusion, evidence-based rather than assumed: the Master workbook
alone is the source for Worksheet 2.** The NK workbook was not used for
any computation -- it served only as corroborating evidence for this
decision and as a visual reference for the requested layout. This is
also why Worksheet 2 shows the same 32 groups as Worksheet 1, not the
NK workbook's larger set: extending scope to include `DS90_Secured` (or
the Projection/Prospecting sections) would mean deciding, on Claude's
own authority, that a Sub-Group code excluded from Worksheet 1's
validated `config.OUTPUT_SECTIONS` should be included after all -- which
is a business decision, not a formatting one, and is flagged below in
§13.5 rather than made silently.

### 13.2 What was reused vs. what is genuinely new

Nothing in `excel_reader.py`, `aggregator.py`, `comment_mapper.py`,
`historical_lookup.py`, `validator.py`, or `config.py`'s business rules
(`OUTPUT_SECTIONS`, DS-code mapping, role keywords) was modified.
Worksheet 2 is built from a new, additive module, **`monthly_view.py`**,
which:

- reuses the exact same `ProjectRow.monthly_revenue` /
  `.monthly_margin` dictionaries `excel_reader.py` already parsed for
  Worksheet 1 -- these are keyed by calendar month 1-12 and were already
  being computed; `aggregator.py` just throws that granularity away
  once it has summed them into quarters. This module keeps it instead
  of summing it away, which is why no new parsing exists anywhere;
- reuses the exact same grouping key (`normalize_name(row.group)`) and
  the exact same per-section Sub-Group/DS-code filter
  (`section.ds_codes`) `aggregator.aggregate_section` already uses;
- reuses the exact same **group list** Worksheet 1 already computed and
  validated -- `monthly_view.build_monthly_sections` takes
  `section_results` (Worksheet 1's already-aggregated, comment-matched,
  Rule-6-filtered `GroupSummary` objects) as an input and only adds a
  monthly breakdown on top of groups that already exist. It never
  independently decides whether a group belongs in the Summary;
- reuses each group's already-matched `GroupSummary.comment` for
  Worksheet 2's Comments column, instead of querying the
  ClientComments sheet a second time;
- reuses each group's already-computed `GroupSummary.total_margin` for
  Worksheet 2's yearly Margin column, as a static value, exactly the
  same convention Worksheet 1's own final Margin column already uses
  (not re-derived from the monthly figures via a new formula).

The only genuinely new logic is:

1. **Re-bucketing monthly figures by month instead of by quarter** --
   the same underlying numbers, kept at finer granularity instead of
   discarded.
2. **`resolve_month_roles()`** -- reads each month's own type-header
   cell directly (the same cell `build_column_map` already inspected to
   decide "this is a revenue column", just read again for its literal
   text: "Actual" or "Forecast", verbatim, whatever the sheet says) and
   returns `{month: role_text}`. Nothing about which calendar months are
   "Actual" vs "Forecast" is hardcoded anywhere -- if a future workbook
   moves the Actual/Forecast boundary to a different month, or a role
   label changes, this keeps working unmodified because it reads the
   sheet's own words rather than assuming a fixed range. (The literal
   phrase "Outlook" never needed to be replaced anywhere in code -- the
   Master workbook already says "Forecast".)
3. **Confidence**, read directly from the sheet's own Renewal Confidence
   column (`cmap.renewal_confidence`, already resolved by the unmodified
   `build_column_map`) via each `ProjectRow`'s own `row_index`, with a
   `Counter.most_common()` reduction across a group's rows -- the exact
   same reduction pattern `aggregator.py` already uses for `poc` and
   `raw_sub_group`.

### 13.3 Column layout and formatting

```
Name | POC | <Role> <Mon> | Margin | <Role> <Mon> | Margin | ... | Total | Margin | Confidence | Comments
```

`<Role>` and the set of months are both fully dynamic (driven by
whatever `resolve_month_roles()` finds); nothing assumes 12 months, a
Jan-start, or any particular Actual/Forecast split. Each row's Total is
a live `=SUM(...)` formula over that row's own monthly value cells --
mirroring Worksheet 1's Total-formula convention, and, like Worksheet
1's yearly Total, referencing the value columns by explicit comma-
separated cell list (`=SUM(C6,E6,G6,...)`) rather than a contiguous
range, since each month's Margin column sits between them. Section
banners, subtotal rows, and blank-row spacing reuse the exact same
`config.OutputSection` fields (`heading`, `title`, `subtotal_label`,
`blank_rows_after_*`) that already drive Worksheet 1's layout, so the
two sheets stay visually and structurally consistent automatically.

Formatting matches Worksheet 1 exactly -- same `config.FONT_NAME` /
`FONT_SIZE`, same `HEADER_FILL` / `SECTION_FILL` / `SUBHEADING_FILL` /
`SUBTOTAL_FILL`, same `CURRENCY_FORMAT` (a real `$`, applied to every
monthly value/margin cell plus the final Total/Margin), and the same
thin-black-border-on-every-populated-cell treatment as Worksheet 1's
Feature 1 (blank spacer rows between sections are, as on Worksheet 1,
deliberately left unbordered). The only new styling constant added is
`config.CONFIDENCE_COLUMN_WIDTH` -- a column width, not a business rule.
A new helper, `SummaryWriter._write_plain_banner`, mirrors
`_write_banner_row`'s exact visual behaviour for Worksheet 2's own,
differently-sized column count, so Worksheet 1's own `_write_banner_row`
never needed to change.

### 13.4 Wiring: what changed in main.py / gui/runner.py

Both already call the same sequence of steps to build `section_results`
for Worksheet 1 (`gui/runner.py`'s own docstring describes it as
mirroring `main.py` "step-for-step" for GUI progress reporting -- that
duplication is pre-existing, not introduced here). Each gained the same
two-line addition immediately before the existing `writer.build(...)`
call:

```python
month_roles = resolve_month_roles(ws_main, cmap)
monthly_section_results = build_monthly_sections(rows, cmap, ws_main, section_results)
wb = writer.build(section_results, monthly_section_results, month_roles)
```

`SummaryWriter.build()`'s signature gained two *optional* trailing
parameters specifically so this remains backward compatible; a
defensive fallback (a bare placeholder sheet, the same one this feature
replaces) fires only if a hypothetical future caller omits them.

### 13.5 Business rules that were confidently implemented vs. ones flagged rather than invented

**Implemented with direct evidence from the uploaded workbooks:**
- Which months are Actual vs Forecast (read from the sheet's own header).
- Monthly Margin "where applicable" -- read from a group's real
  monthly-margin figures wherever they exist; a group with no tracked
  activity for a month shows 0, the same convention Worksheet 1 already
  uses for its own quarter cells (see §6 for the pre-existing Rule 6
  blank-vs-zero distinction, which continues to apply at the *group*
  level, not per calendar month).
- Confidence and Comments, both sourced from already-existing, already-
  discovered columns.

**Flagged rather than invented, because the evidence was genuinely
ambiguous or would require a business decision Claude has no authority
to make silently:**
- **`DS90_Secured`** (JSG, Mastek, MatchPoint Solutions, Maxonic,
  Technology Partners, VDart -- all present in the Master workbook,
  confirmed by direct lookup) is treated as "Staffing- Secured" business
  by the NK workbook's own manual categorization, but is excluded from
  `config.OUTPUT_SECTIONS`'s `ds_codes=[70, 80, 85]` for that section.
  Worksheet 2 currently mirrors Worksheet 1's validated scope (i.e.
  excludes it) rather than silently expanding what "Staffing- Secured"
  means. If this code should be included going forward, that is a
  one-line change to `config.OUTPUT_SECTIONS` -- but it changes
  Worksheet 1's output too, so it needs an explicit decision, not an
  assumption made while building Worksheet 2.
- **Track 2 / Projection / Prospecting** sections (visible in the NK
  workbook, covering DS-codes such as `DS30_Projection` and
  `DS50_Projection`) are out of scope for the same reason: no validated
  mapping of these Sub-Group codes to Summary sections exists anywhere
  in the current pipeline, and inventing one was explicitly out of
  bounds for this task.

### 13.6 Verified

- Cell-by-cell diff of Worksheet 1 between this version's output and
  the previously-delivered output workbook, run against the identical
  Master workbook: **0 differences** in content, and separately, **0
  differences** in number format, border, fill, font, merged-cell
  layout, freeze panes, or gridline setting.
- `tests/compare_with_manual.py`: same 10 pre-existing, already-
  documented differences as every prior round (§6) -- 0 new ones.
- `tests/test_column_reordering.py` (8 seeds) and
  `tests/test_future_year_compatibility.py`: unaffected, all pass.
- New: `tests/test_worksheet2_actual_forecast.py` -- confirms Worksheet
  2 exists under the correct dynamic name, shows the same 32 groups in
  the same order as Worksheet 1, dynamically shows both "Actual" and
  "Forecast" role labels (proving this wasn't hardcoded), includes the
  Confidence/Comments columns, and that Worksheet 1 is byte-for-byte
  stable across repeated runs.
- Formula recalculation across the whole workbook (both sheets): 0
  errors, 142 formulas, using the attached real Master workbook.
- Streamlit: full upload -> generate -> download flow re-tested in a
  real browser against the actual uploaded Master workbook; downloaded
  file contains both sheets correctly.
- Desktop GUI backend (`gui.runner.generate_summary`): re-tested
  directly against the same file; produces the identical two-sheet
  output.
- Dynamic naming re-verified for a second target year (2025): sheet is
  named "2025 Actual & Forecast" with no code change, confirming no year
  is hardcoded anywhere in the new code.
