# NormalizeIdentifierColumns bulk-write rewrite

Branch `feat/normalize-bulk`, base `76062be`. Commit `6d1ae16`
`perf(NormalizeIdentifierColumns): bulk read/compute/write`.

## Fix round 1 (review finding: blanket write over unchanged cells)

Commit `808fec4` `fix(NormalizeIdentifierColumns): drop blanket write`.

### The bug

`NicSweepColumn`'s "no formula cells" branch did:
```
ElseIf Not anyFormula Then
    targetRange.NumberFormat = "@"
    targetRange.Value = targetValues
```
over the *entire* `2..lastRow` range -- every currently-configured tenant column takes this
branch, since none declares a mixed-formula column today. `targetValues` held each unchanged
cell's own current value (so the *value* round-tripped to the same thing), but the
`NumberFormat = "@"` and `Value =` calls still executed against WARRANTY markers, blanks, and
already-normalized text cells that never needed either write. That defeats guarantee #4
(already-normalized text is not rewritten, specifically to keep the co-author-visible diff
minimal on a live shared workbook) at the COM-call level, even though the value never visibly
changes. The checked-in tests didn't catch it because `Set-Cell` force-set `NumberFormat = "@"`
on every string fixture cell *before* the sweep ran, so a spurious re-set to the same "@" was
invisible to a post-sweep value-only assertion.

### The fix

Deleted the blanket-write branch entirely. `NicSweepColumn` now has exactly two write paths:
`filterActive -> NicWritePerCell` (unchanged from the first round) and, otherwise,
`NicWriteRuns` unconditionally -- the same run-based writer previously reserved for the
mixed-formula case, now used for every non-filtered write. Its run-inclusion test was already
`(Not rowIsFormula(r)) And rowNeedsWrite(r)`, i.e. keyed on "needs a write", not "isn't a
formula" -- so routing the common case through it required no logic change to `NicWriteRuns`
itself, only removing the special-cased branch and the now-unused `anyFormula` tracking
variable. On a first sweep of an unnormalized column every row needs writing, so the whole
column is one contiguous run and collapses to exactly one bulk write -- the original
performance win is unchanged for that case. On a later Refresh, once most cells are already
normalized, it naturally shrinks to writing only the runs that changed, which is guarantee #4
working as designed rather than a special case.

Confirmed the branch is actually gone, not just unreachable:
```
$ grep -n "Not anyFormula\|anyFormula\|targetRange.Value = targetValues\|targetRange.NumberFormat" NormalizeIdentifierColumns.vb
(no output)
```

### Test changes

`Set-Cell` no longer force-sets `NumberFormat = "@"` on every string fixture cell. It now only
does so for a numeric-looking string (needed so Excel doesn't re-numericize it on assignment --
a real setup requirement, not a mask); a non-numeric string (WARRANTY-style) and a blank cell
are left with whatever format a fresh cell already has, so a test can tell whether the sweep
touched the cell's format at all.

Added format-precondition assertions in two places:
- **Ordinary sweep test** (has a formula cell, so it exercises the run-based writer either way):
  planted `NumberFormat = "0.00"` on the already-normalized-text cell (`'1166552'`) after
  Set-Cell's own `"@"` setup step, and recorded the natural pre-sweep format of the WARRANTY and
  blank cells. Asserted all three are unchanged after the sweep.
- **Larger-scale (300-row) test** (zero formula cells -- the exact shape that took the buggy
  blanket-write branch): same treatment on one already-clean-text row and one blank row. This
  is the assertion that actually caught the bug (see RED below); the ordinary-sweep test's
  formula cell meant it always used the run-based writer even in the buggy code, so it could
  not have caught this specific regression on its own.

### TDD evidence

**RED against the buggy blanket-write code** (`git show 6d1ae16:NormalizeIdentifierColumns.vb`
restored temporarily), first attempt: `passed: 49 failed: 0` -- the new assertions in the
ordinary-sweep test didn't fire, because that fixture has a formula cell (row 3, `=1+1`), so
`anyFormula` was already `True` and the buggy code took the run-based path anyway, same as the
fix. This matches the review's diagnosis exactly: the bug only manifests when a column has *no*
formula cells. Added the equivalent assertions to the 300-row (zero-formula) test instead:

```
FAIL an already-clean text row gets no format write in the no-formula bulk path either -- expected '0.00', got '@'
FAIL a blank row gets no format write in the no-formula bulk path either -- expected 'General', got '@'

passed: 49  failed: 2
```
This is genuine RED: both failures show the buggy code stamped `"@"` onto cells that never
needed a write, exactly the defect described.

**GREEN** (fix restored): `passed: 51  failed: 0` / `ALL PASS`.

### Full suite re-run (post-fix)

All exit 0:

| Suite | Result |
|---|---|
| Test-FiscalCalendar.ps1 | passed: 643 failed: 0 |
| Test-DocumentModule.ps1 | passed: 15 failed: 0 |
| Test-ReqJoinKey.ps1 | passed: 2 failed: 0 |
| Test-HashFile.ps1 | passed: 6 failed: 0 |
| Test-ProcessedBatchLog.ps1 | passed: 15 failed: 0 |
| Test-MoveProcessedFile.ps1 | passed: 10 failed: 0 |
| Test-BillCodeScreen.ps1 | passed: 13 failed: 0 |
| Test-MirrorInvoiceBlock.ps1 | passed: 42 failed: 0 |
| Test-NormalizeIdentifierColumns.ps1 | passed: 51 failed: 0 |

`git -C ../SecuritasAutomation status --short` and `git -C ../JCI-invoice-tracker status
--short` both returned nothing -- untouched.

### README

Rewrote the write-path description to describe run-based writes unconditionally (one run per
contiguous stretch of rows needing a change, collapsing to a single write on a first, fully-
unnormalized sweep), with an explicit note on why there is no separate whole-range branch and
what happened when there was one. The filtered-write fallback description is unchanged.

### Item 5 (optional): extending the filter test to multiple interleaved hidden/visible rows

Not done, per the coordinator's "don't spend much time on this if it's fiddly." Skipping it is
also lower-value now than it would have been before this fix: the filtered-write fallback
(`NicWritePerCell`) is used *unconditionally* whenever `ws.FilterMode` is `True`, regardless of
how many rows are hidden or how they're interleaved -- there is no code path left that could
attempt a multi-row array write across a filtered range in production code, so the dramatic
"wrong visible row" failure mode from the original COM probe is structurally excluded, not just
untested at larger scale. The existing single-hidden-row fixture still fully exercises the code
path actually in use (one write per changed cell, unconditionally, under a filter); more
interleaving would exercise the same one `If`/`Else` branch repeatedly rather than covering new
logic.

### Commit

`808fec4` -- `fix(NormalizeIdentifierColumns): drop blanket write`. Not pushed, not merged.

## Before/after shape

**Before:** `NormalizeIdentifierColumnsOn` looped `r = 2 To colLastRow` and, per row, called
`ws.Cells(r, colLetter).HasFormula` then (if not a formula) `.Value` -- two COM round trips per
cell, ~2 * 3446 * 4 ≈ 27k calls on Securitas's sheet across its four declared columns every
`Refresh`. Writes (`NumberFormat` + `Value`) were also per-cell, only issued for rows that
actually changed.

**After:** one bulk `.Value` read and one bulk `.Formula` read per column (rows 2..lastRow),
with everything else decided in VBA memory:

- A row is a formula row when its bulk-read `.Formula` text starts with `"="`
  (`NicIsFormulaText`), guarded against `IsError`/`IsNull` so `CStr` never raises.
- Blank/Null/Error rows and already-normalized string rows are left as-is in the in-memory
  target array (their own current value); numeric rows and string rows whose trimmed form
  differs are marked `needsWrite` and counted.
- Write-back (`NicSweepColumn`), in priority order:
  1. **Active `AutoFilter` (`ws.FilterMode`) -> `NicWritePerCell`.** One `NumberFormat`/`Value`
     write per changed, non-formula row -- i.e. exactly the old per-cell behavior, minus the
     old per-cell reads (already done in bulk). See "Filter-write bug found during TDD" below
     for why this fallback exists.
  2. **No formula cells, no active filter -> one blanket write.** `targetRange.NumberFormat =
     "@"` then `targetRange.Value = targetValues` over the whole range, identical in shape to
     `ConvertStoreNumbersOn`. This is the path every currently-configured tenant column takes.
  3. **Formula cell mixed in, no active filter -> `NicWriteRuns`.** One bulk write per maximal
     contiguous run of changed, non-formula rows, so a formula cell is never inside a written
     range and a long clean stretch still costs one call.

`LastNonEmptyRow` (bulk `UsedRange` read, filter-proof) is unchanged.

Guarantees preserved: never writes a formula cell (all three write paths explicitly exclude
formula rows); filter-proof last row (`LastNonEmptyRow` untouched, still exercised by the
filtered-sheet test); WARRANTY/blank/Null/Error cells get no format or value write (they keep
their own value in the array and are never in a `needsWrite` run); already-normalized text is
not rewritten (same reason); `NumberFormat = "@"` always precedes the value write in all three
paths; no `Err.Raise`/`MsgBox` anywhere in the new code; return value is the exact count of
cells actually changed.

## TDD evidence

Added two cases to `tests/Test-NormalizeIdentifierColumns.ps1`:
1. A formula cell (`=5+5`) placed in the middle of column A, with numeric/string cells needing
   normalization both before and after it, plus a WARRANTY marker and a blank after it.
   Asserts `HasFormula` stays `True`, the formula text and calculated value are unchanged, every
   literal cell that needed normalizing was normalized, and the returned count is exactly 4.
2. A 300-row column cycling numeric / already-clean-text / blank, asserting the returned count
   equals exactly the number of numeric rows, spot-checking a handful of rows including the
   first and last numeric rows, and asserting the sweep call completes in under 10 seconds
   (`System.Diagnostics.Stopwatch`) as a regression guard against reintroducing a per-cell/
   quadratic sweep.

**RED (against the old per-cell implementation, `git show HEAD:NormalizeIdentifierColumns.vb`
restored temporarily):** all 46 assertions passed (`passed: 46  failed: 0`). This is expected
and not a false negative: the two new tests encode *correctness* guarantees the old per-cell
code already satisfied (it processed every row individually, so formula-in-the-middle and
mixed-shape correctness were never actually broken by the defect -- only its call count was).
The timing ceiling was deliberately set generously enough (task instruction) that it does not
fail against a merely-slower-but-linear old implementation either; it exists to catch a future
*quadratic* reintroduction, not to distinguish "slow" from "fast" at 300 rows. So this pair of
tests is a genuine regression guard for the new bulk code paths (`NicSweepColumn`,
`NicWriteRuns`) going forward, even though it did not RED on correctness against the code being
replaced.

**RED that did fire (against my first draft of the new implementation):** running the full
suite after switching to the bulk implementation failed 2 of 46:
```
FAIL the hidden row is normalized to text -- expected 'String', got 'Double'
FAIL the hidden row gets @ format too -- expected '@', got 'General'
```
This is the pre-existing filtered-sheet test (trailing hidden row), not one of my two additions
-- see "Filter-write bug found during TDD" below.

**GREEN (after the filter fix):** `passed: 46  failed: 0` / `ALL PASS`.

## Filter-write bug found during TDD

The RED failure above led to a direct COM probe (outside the VBA harness, plain PowerShell +
`Excel.Application`) of what a multi-row `Range.Value2 = array` write actually does when the
sheet has an active `AutoFilter` with hidden rows in the target range:

- 2-row range (A2 visible, A3 hidden), 2-element array: A2 got the correct element, A3 kept its
  old value entirely -- the hidden row was silently skipped.
- 5-row range (rows 3 and 5 hidden, 2 and 4 and 6 visible), 5-element array: **every visible
  row received the array's first element**, not its own -- A2/A4/A6 (all three visible cells)
  all got `"V2"` instead of their respective `"V2"/"V4"/"V6"`. Hidden rows again kept their old
  values untouched.
- The same experiment with rows hidden manually (`.Rows(n).Hidden = True`, `FilterMode` False)
  wrote correctly in the expected order -- the corruption is specifically tied to
  `ws.FilterMode = True`, not to `Hidden` in general.

So a multi-row bulk write is unsafe -- not merely slow, but capable of writing the *wrong*
value into a visible cell -- whenever the sheet has an active AutoFilter hiding rows inside the
written range. `NormalizeIdentifierColumnsOn` now reads `ws.FilterMode` once per call and routes
every column through `NicWritePerCell` (one write per changed cell, no bulk range write) when
it's `True`, which is exactly this module's pre-rewrite behavior and therefore already proven
correct by the original filtered-sheet test. This does forgo the bulk-write speedup while a
filter is active, but that is not the case a scheduled `Refresh` normally runs in, and
correctness under a live filter is an explicit, tested guarantee (#2) that could not be traded
away.

## Test suite results

All run via `powershell -NoProfile -ExecutionPolicy Bypass -File tests\<name>.ps1`, all exit 0:

| Suite | Result |
|---|---|
| Test-FiscalCalendar.ps1 | passed: 643 failed: 0 |
| Test-DocumentModule.ps1 | passed: 15 failed: 0 |
| Test-ReqJoinKey.ps1 | passed: 2 failed: 0 |
| Test-HashFile.ps1 | passed: 6 failed: 0 |
| Test-ProcessedBatchLog.ps1 | passed: 15 failed: 0 |
| Test-MoveProcessedFile.ps1 | passed: 10 failed: 0 |
| Test-BillCodeScreen.ps1 | passed: 13 failed: 0 |
| Test-MirrorInvoiceBlock.ps1 | passed: 42 failed: 0 |
| Test-NormalizeIdentifierColumns.ps1 | passed: 46 failed: 0 |

`git -C ../SecuritasAutomation status --short` and `git -C ../JCI-invoice-tracker status
--short` both returned nothing -- untouched, as required.

## README

Updated the "Normalizing identifier columns" section to describe the bulk read/compute/write
shape (mirroring `ConvertStoreNumbers.vb`), the three write paths, and the AutoFilter write
hazard/fallback, replacing wording that implied a per-cell sweep.

## Commit

`6d1ae16` -- `perf(NormalizeIdentifierColumns): bulk read/compute/write`. Not pushed, not merged.

## Concerns / what wasn't exercised at real scale

- **The mixed-formula, no-filter path (`NicWriteRuns`) is only exercised by the new
  formula-in-the-middle unit test (9 rows).** No currently-configured tenant declares a mixed
  identifier column, so this path has no production data to validate against and has never run
  against a few-thousand-row sheet. The run-splitting logic itself is straightforward (split on
  formula/no-need boundaries), but it has not been load-tested.
- **The AutoFilter fallback (`NicWritePerCell`) is only exercised by the pre-existing
  9-row/1-hidden-row filtered-sheet test.** It is byte-for-byte the old per-cell write pattern,
  so I'm confident in its correctness, but it means the bulk-write performance win this task
  exists to deliver does **not** apply while a user has an active filter on the tracker sheet
  during `Refresh`. If Securitas (or another tenant) routinely runs `Refresh` with a filter
  active on a ~3500-row sheet, that specific run would still pay close to the old per-cell cost.
  This wasn't in scope to fix further (the task's own guarantee #2 required correctness under a
  filter over speed), but it's worth knowing about if filtered-during-Refresh turns out to be
  common rather than rare.
- The AutoFilter write corruption itself was not previously documented anywhere in this repo
  (only the read-side `End(xlUp)`/`Find` trap was). It's now documented in the
  `NicSweepColumn`/`NicWritePerCell` comments and here; worth a mention if `MirrorInvoiceBlock.vb`
  or any other module is ever changed to do bulk multi-row writes near filterable data.
