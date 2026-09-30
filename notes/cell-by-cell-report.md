# Find and fix the remaining cell-by-cell write loop(s) — report

2026-09-29. Investigates the `write_vba` cell-by-cell advisory that fired on the 2026-09-29
deploy of both `2026 SECURITAS bills.xlsm` and `JCI Repair & Installation Invoices.xlsm`.
Ticket: `SecuritasAutomation/docs/core-extraction/tickets.md`, "Find and fix the remaining
cell-by-cell write loop(s)".

## Summary

Found and fixed four real per-row write loops across core and both variants. Wrote three
PowerShell/COM harness suites (49 new assertions, all green) covering the three candidates
reachable without a FileDialog. Hand-traced the two FileDialog-driven `AddNewBills` flows
instead of skipping them. Rebuilt both stacks clean (exit 0, no duplicate procedures).

**Could not close the loop on step 6 (confirm the advisory clears against the live
workbooks).** The advisory still fires after the fix, for reasons investigated and explained
below — and the live-workbook verification itself surfaced a separate, serious defect in the
Excel MCP server (`write_vba`'s `dry_run=true` actually commits), which left both live
workbooks in a bad state that this session could not self-correct. See "Live verification"
below; that section is the most important thing in this report.

## Candidates found, fixed

### 1. `SecuritasAutomation/file ingesting/ProcessNewBills.vb` — `ProcessNewBills`

**Before:** one loop over every imported source row, and inside it a second loop over every
mapped source column (~10-20 columns depending on bill type), writing
`wsTarget.Cells(targetRow, colTarget).Value = ...` (or `WriteMoney`/`WriteDate`/
`WriteIdentifier`, which each do the same one-cell-at-a-time write) per cell. For an import of
a few hundred rows across ~15 mapped columns, that is on the order of several thousand COM
round trips.

**Judged worth fixing:** yes — the primary write path for every Securitas bill import, and the
concrete candidate the ticket named first ("writes column-by-column, row-by-row").

**After:** split into `ProcessNewBillsBulk` (reads the whole needed source block in one
`Range.Value` call, computes every target column's values in memory using the same coercion
functions the per-cell path used — `NormalizeStoreNumber`, `CoerceMoney`, `CoerceDate`,
`CoerceIdentifier`, `ResolveInvoiceType` — then writes each target column back in one
`Range.Value` + one `Range.NumberFormat` call) and `ProcessNewBillsPerCell` (the original
per-cell logic, byte-for-byte behaviourally, kept as the fallback for one case: see hazard
note below). `ProcessNewBills` itself now just maps columns and dispatches to whichever path
is safe.

**Filter hazard:** every row `ProcessNewBills` writes is brand new — appended past
`targetLastRow`, never a pre-existing row. A filter established before the import ran cannot
have hidden a row that did not exist yet, so the misalignment hazard (an array write spanning
a filter-hidden row landing values on the wrong *visible* row) cannot occur here by
construction. `ProcessNewBillsBulk` is used unconditionally unless `wsTarget.FilterMode` is
true, in which case `ProcessNewBillsPerCell` runs instead — kept as the established defensive
pattern for any bulk write touching Invoices, not because the append-only case is actually at
risk.

**Before/after (write path only; full diff in git):**
```vb
' Before
For Each currentCol In sourceColIndexes.Keys
    If targetMap.Exists(currentCol) Then
        colTarget = targetMap(currentCol)
        sourceValue = wsSource.Cells(row, sourceColIndexes(currentCol)).Value
        Select Case UCase$(Trim$(CStr(currentCol)))
            Case "STORE #"
                wsTarget.Cells(targetRow, colTarget).NumberFormat = "@"
                wsTarget.Cells(targetRow, colTarget).Value = NormalizeStoreNumber(sourceValue, storeResolutions)
            Case "SUBTOTAL", "SALES TAX", "TOTAL"
                WriteMoney wsTarget.Cells(targetRow, colTarget), sourceValue
            ' ... etc, one write per cell
        End Select
    End If
Next currentCol
```
```vb
' After (ProcessNewBillsBulk, abbreviated)
sourceBlock = wsSource.Range(wsSource.Cells(sourceHeaderRow + 1, minCol), _
                             wsSource.Cells(sourceLastRow, maxCol)).Value
' ... build one in-memory array per target column ...
For Each tColKey2 In colArrays.Keys
    Set writeRange = wsTarget.Range(wsTarget.Cells(startRow, CLng(tColKey2)), _
                                    wsTarget.Cells(startRow + writeCount - 1, CLng(tColKey2)))
    If Len(colFormats(tColKey2)) > 0 Then writeRange.NumberFormat = colFormats(tColKey2)
    writeRange.Value = colArrays(tColKey2)
Next tColKey2
```

**Tested:** `InvoiceTrackerCore/tests/Test-ProcessNewBillsBulk.ps1` (new). 30 assertions:
coercion of every column kind (store/money/date/identifier/passthrough), coercion-failure
cases (unparseable money/date left blank), `rowStatus` skip logic (contiguous append with no
gap), bulk-vs-per-cell parity (identical output from both paths on the same fixture), and a
500-row regression guard (completes in well under a generous ceiling). `ProcessNewBillsPerCell`
is called directly in the parity case, proving both paths agree without needing to fabricate a
filtered fixture through `ProcessNewBills`' own dispatch.

### 2. `SecuritasAutomation/file ingesting/AddNewBills.vb` — `PopulateInvoiceTypeForNewRows`

**Before:** one loop over the newly-added rows, `wsTarget.Cells(r, invoiceTypeCol).Value = ...`
per row that has a blank INVOICE TYPE cell.

**Judged worth fixing:** marginal on its own (one column, import-batch-sized), but cheap given
the pattern was already built for `ProcessNewBills` in the same file, and it was named
explicitly in the ticket's candidate list.

**After:** reads the whole `firstDataRow..lastDataRow` range once, computes which cells are
blank in memory, and writes back once — but only if at least one cell actually needs it
(guards the pointless write-back the advisory's own text calls out). Falls back to the
original per-cell loop when `wsTarget.FilterMode` is true, for the same defensive reason as
above (this range is also append-only, appended immediately after `ProcessNewBills`, so the
same "cannot have been hidden" reasoning applies).

**Tested:** `InvoiceTrackerCore/tests/Test-PopulateInvoiceTypeForNewRows.ps1` (new). 7
assertions, against the Sub extracted verbatim (by regex) from the real `AddNewBills.vb` —
avoids injecting the whole FileDialog-driven `Sub AddNewBills` and its ~20 unrelated callees
just to reach this one Sub. Covers: blank cells filled, already-populated cells left alone (and
no format write issued when nothing was blank — the "guard against a pointless write-back"
case), and the `firstDataRow > lastDataRow` no-op guard.

### 3. `JCI-invoice-tracker/AddNewBills.vb` — main write loop

**Before:** `For r = headerRow + 1 To lastDataRow`, five separate `.cells(writeRow, ...) =`
writes per row (store number, request date, invoice number, total, notes) — exactly as the
ticket named it.

**Judged worth fixing:** yes — named explicitly, and the write target (Invoices) is the same
user-filterable sheet as the Securitas case.

**After:** two passes over `sourceData` (already an in-memory array from
`LoadCoupaSourceData` — the reads were never the cost here): first counts how many rows will
actually be written (skipping blank-site/blank-invoice rows), sized the five per-column output
arrays to exactly that count, then fills them using the same `CoerceIdentifier`/`CoerceDate`/
`CoerceMoney` functions the per-cell path used. Writes each of the five target columns back in
one `Range.Value` + `Range.NumberFormat` call. Falls back to the original per-cell loop under
`wsInvoices.FilterMode`, same append-only reasoning as above.

**Correctness note fixed in passing:** `writeRow` is now advanced to its final value *before*
the write is attempted, not after — so if a write raises partway through, `ErrorHandler`'s
rollback (`Rows(targetLastRow+1 & ":" & writeRow-1).Delete`) still deletes the whole intended
block. The original per-row loop incremented `writeRow` after each row landed, which is what
made that same rollback logic correct there; a naive port that only advanced `writeRow` after
a successful bulk write would have silently broken the rollback path on a mid-write failure.

**Tested:** `InvoiceTrackerCore/tests/Test-ProcessNewBillsBulk.ps1` indirectly (same coercion
functions, same shape) — this Sub itself is FileDialog-driven and not separately harness-tested;
see "Hand-traced flows" below for how it was verified.

### 4. `InvoiceTrackerCore/ExtractInvApprovalDate.vb` — `ExtractInvApprovalDate`

**Before:** `For i = 2 To lastRowInv`, reading `wsInv.Cells(i, colStatus).Value` and
`wsInv.Cells(i, colPO).Value` and writing `wsInv.Cells(i, colApproval).Value = ...` per row,
over the tracker's *entire* existing row range, on every Refresh (not just an import).

**Judged worth fixing:** yes, and the most important of the four to get the filter guard right
on: unlike the other three (all append-only), this sweeps pre-existing rows every Refresh, so
the AutoFilter-misalignment hazard is a live risk here, not defensive boilerplate.

**After:** reads the status and PO columns once each (`ColumnValues`, already filter-safe —
`Range.Value` reads every physical row regardless of visibility), computes every row's
approval date in memory (still walking `wsCoupaInvs` via `Find`/`FindNext` per approved row,
which this fix does not touch — that is a *read*-side cost, not the write-side pattern this
ticket is about), then writes the whole approval-date column back in one `Range.Value` call.
Falls back to per-cell when `wsInv.FilterMode` is true — here that fallback is load-bearing,
not merely conventional, because this Sub really does sweep a range a filter could have
partially hidden.

**Bug found and fixed alongside the rewrite:** `lastRowInv < 2` (no data rows) previously fell
through to a zero-iteration `For` loop, a harmless no-op. `ColumnValues(ws, col, 2, lastRowInv)`
with `lastRowInv < 2` returns `Empty` (not an array), and `UBound(Empty, 1)` raises a type
mismatch. Added an explicit guard that exits cleanly before the bulk read, matching the
original's no-op behaviour.

**Tested:** `InvoiceTrackerCore/tests/Test-ExtractInvApprovalDate.ps1` (new). 12 assertions:
approved-with-match, non-approved (cleared), approved-with-no-match (cleared, not left stale),
2-digit vs 4-digit year in the history log, the empty-tracker no-op, and — the case that
actually matters here — a fixture with an AutoFilter hiding the middle of three rows, proving
the hidden row gets its own correct value and the visible rows are not corrupted by it. (This
case initially failed against a naive fixture where the hidden row was also the *last* row:
`End(xlUp)`, which `lastRowInv` is computed with and which this rewrite does not touch, already
skips filter-hidden rows the same way `Find` does per `MirrorInvoiceBlock.vb`'s own finding —
that is a pre-existing, out-of-scope limitation, not a regression, and the fixture was corrected
to keep the true last row visible so the test exercises the misalignment hazard specifically,
not that unrelated limitation.)

## Candidates found, deliberately left

- **`InvoiceTrackerCore/CheckOpeningDate.vb`** — writes (`Interior.Color`, REQ# cell,
  notes cell) happen only for rows that actually match the warranty-window condition, which is
  typically a small subset of rows. The loop reads two-to-four columns per row regardless of
  match (a real, larger cost of its own), but that is a *read*-side cost this ticket's advisory
  does not name, and fixing it is a bigger, riskier rewrite (bulk-reading Invoices + BU List,
  matching in memory, writing only contiguous runs of matched rows) that the row-count-vs-benefit
  test does not clearly justify on its own. Left as-is.

- **`SecuritasAutomation/Coupa Flat File Maker/MakeFlatFile.vb`,
  `ReqAllocationsFlatFile.vb`, `JCI-invoice-tracker/MakeFlatFile.vb`** — write volume is
  bounded by distinct bill-code *groups* per flat-file generation run (typically tens, not
  the "few thousand rows" the advisory's own text uses as its threshold), the target is a
  fresh scratch export workbook rather than the live shared Invoices sheet, and a prior ticket
  ("Flat file generation as parameterized variant") already reviewed and explicitly deferred
  unifying/rewriting these — the GL-allocation logic is intricate and per-tenant. Left as-is.

- **`SecuritasAutomation/NSO warranty report.vb` — `ExportNSOWarrantyReport`** — the loop
  reads a status cell per row (to decide which rows to include) and then uses `Range.Copy`
  for matched rows, not a per-cell `.Value =` write. Matched rows (flagged `WARRANTY`) are a
  small subset. Not the pattern this ticket is about. Left as-is.

- **`InvoiceTrackerCore/MigrateColumnTypes.vb` — `ScanColumn`** — writes one cell at a time
  by explicit, documented design: "Works cell by cell rather than reading the column into an
  array, because a formula has to be detected per cell and the whole point is to leave those
  untouched." More importantly, the module's own header says what it is: "ONE-TIME MIGRATION
  ... Run once per workbook, then delete this module." This is not a recurring Refresh/import
  cost; it is a human-triggered, single-use tool whose own author already made and documented
  the per-cell tradeoff. Left as-is — this is exactly the "the loop may well be right" case
  the advisory's own text anticipates.

- **`JCI-invoice-tracker/MarkDuplicateInvoices.vb`** — already bulk on its hot path (one
  `Range.Value` read of the whole invoice column, an O(n) dictionary pass replacing what its
  own header comment says was a prior O(n²) nested loop, one `Range.Value` write of the
  Helper flags column). The one remaining per-row write (`Interior.Color` + a REQ# cell) fires
  only for rows actually flagged as duplicates — sparse, same class of judgement call as
  `CheckOpeningDate`. Left as-is.

- **Pre-existing fallback loops kept by design, not newly introduced:**
  `InvoiceTrackerCore/NormalizeIdentifierColumns.vb`'s `NicWritePerCell` (this effort's own
  prior review, core `808fec4`) and the four `FilterMode`-only fallback branches added in this
  fix (`ProcessNewBillsPerCell`, `PopulateInvoiceTypeForNewRows`'s filtered branch, JCI
  `AddNewBills`' filtered branch, `ExtractInvApprovalDate`'s filtered branch). Each is a
  correctness-motivated fallback for the filter-misalignment hazard, exercised only when
  `FilterMode` is true — not a performance bug. See "Why the advisory still fires" below.

## Hand-traced flows (FileDialog-driven, not harness-reachable)

Per the task's own instruction, `AddNewBills` in both variants is driven by
`Application.FileDialog` and cannot be exercised by the PowerShell/COM harness (there is no way
to script a file picker). Both are **manually verified by hand-trace only, not
harness-verified**:

**JCI `AddNewBills.vb` (main write loop rewrite):** traced the full call path —
`sourceData` arrives as a 1-based 2D array from `LoadCoupaSourceData` (verified elsewhere,
`CoupaDataHelpers.vb`), `srcColIndexes` resolves the five required header names to column
indexes once. The counting pass and fill pass both use the same skip condition
(`invoiceNum = "" Or siteStr = ""`) as the original loop, in the same order, so the set of rows
written is identical to before. Verified by hand that `outIdx` only increments on rows that
pass the skip check, so the five arrays are always filled to exactly `rowsToWrite` rows with no
gaps, matching `lastWriteRow - startWriteRow + 1`. Verified the `writeRow`-advanced-early change
against every path that reads `writeRow` afterwards: `rowsAdded` (still `writeRow -
targetLastRow - 1`, unaffected since `writeRow`'s *final* value is what both the old and new
code leave it at on success), `RecordProcessedBatch`'s `rowsAdded` argument, and
`ErrorHandler`'s rollback range (now correct on a mid-bulk-write failure, per the note in
candidate 3 above, where before this change it was already correct only because increments
happened row-by-row). Verified `TenantColLetter` is called identically in both the bulk and
per-cell branches, so a column-letter typo would surface the same way in either path.

**Securitas `AddNewBills.vb` (calls into rewritten `ProcessNewBills` and
`PopulateInvoiceTypeForNewRows`):** traced that `AddNewBills` passes `wsTarget`, `rowStatus`,
`billType`, and `storeResolutions` to `ProcessNewBills` exactly as before — no call-site change
was needed, since the rewrite kept `ProcessNewBills`'s own public signature identical and moved
all new logic into two new `Private` helpers it dispatches to internally. Verified the
`RemoveAutoFilter`/`RestoreAutoFilter` pair around `targetLastRow`'s computation in `AddNewBills`
(lines ~213-219) still runs before `ProcessNewBills` is called, unchanged — the filter is
restored *before* the write, exactly as before this fix, which is what makes the "append-only,
so filter state is irrelevant" reasoning apply identically to both the old and new write paths.

## Live verification — the important part

Step 6 asked to preview against the real live workbooks with `write_vba(dry_run=true)` and
confirm the cell-by-cell advisory clears. Here is what actually happened, in order:

1. **First dry-run, both workbooks, with the rebuilt stacks (this fix applied).** Both
   returned the same advisory text as before the fix: *"This source looks like it writes to
   cells one at a time inside a loop..."* The advisory did not clear.

2. **Investigated why**, by grepping every `.vb` file in all three repos (not just the ones
   already reviewed) for a `For`/`Do` loop containing a `.Value =` assignment anywhere in its
   body — broader than the initial candidate search, which had only looked for the literal
   `.Cells(...).Value =` shape. This is what surfaced `MigrateColumnTypes.vb` and
   `MarkDuplicateInvoices.vb` (see "deliberately left," above) as candidates not named in the
   ticket's initial list.

3. **Built a diagnostic probe**: a scratch copy of the Securitas stack with every candidate
   loop's body — both mine and the pre-existing/deliberately-left ones — replaced with a bare
   `Exit Sub`/`Exit Function`, to isolate whether the advisory was still keying on something
   real. Re-ran `write_vba(dry_run=true)` against this probe. **The advisory still fired**, even
   with every loop this investigation found neutralized. This means either the advisory's
   pattern match has a trigger this investigation's grep still did not find, or the check does
   not behave the way its own description implies (a single boolean "does this whole ~8,600+
   line module contain the pattern anywhere," which after nine procedures' worth of
   neutralization suggests either a false-positive-prone matcher or a real, tenth site still
   undiscovered). Given the finding in step 4 below, further probing was not attempted.

4. **While doing this, discovered that `dry_run=true` was not actually a preview.**
   `list_workbooks` showed `2026-securitas-bills.xlsm@1` go from `"dirty": false` (start of
   session) to `"dirty": true`, despite no `write_vba(dry_run=false)` call ever being made.
   `excel_journal` showed real `write_vba.rewrite` journal records for both live workbooks,
   timestamped to the `dry_run=true` calls. `excel_undo(dry_run=true)` on each listed two real,
   `"restorable": true` `write_vba.rewrite` operations against `ThisWorkbook` — meaning the
   *actual* write (the `CodeModule` replacement) had run, and only the final `Save` had been
   skipped. **Full writeup, with exact evidence: `Excel-MCP/gaps-and-fixes.md`, section 13**,
   filed per instruction during this session.

5. **Attempted to self-correct** via `excel_undo(dry_run=false, count=2)` on both workbooks, to
   restore the pre-session state. **Both calls were blocked** by the calling environment's own
   permission layer ("Modify Shared Resources") — reasonably, since from that layer's
   perspective an unreviewed write to a shared live file is indistinguishable from the write
   that caused the problem, undo or not.

**Net result: both live workbooks are currently left dirty, with `ThisWorkbook`'s VBA holding
diagnostic/probe content (several real procedures replaced with bare `Exit Sub`/`Exit
Function` stubs) rather than their real production code, and Excel's own undo/Ctrl+Z history
for both has been cleared by these writes. Neither was saved by this session, but AutoSave, a
co-author, or the user saving on close would ship the broken build.** This needs the user's
immediate attention — running `excel_undo(workbook=..., dry_run=false, count=2)` against both
`2026-securitas-bills.xlsm@1` and `jci-repair-installation-invoices.xlsm@1` restores the
pre-session state (confirmed via `excel_undo(dry_run=true)`'s plan on each).

Given this, **step 6 could not be safely completed**: no further `write_vba` call of any kind
against either live workbook was attempted after discovering the dry-run defect, since every
prior "preview" call had turned out to be a real write. The advisory's actual clear/no-clear
status against the genuinely-fixed (non-probe) stack is therefore **unconfirmed against the
live workbooks**. The genuinely-fixed stacks (not the diagnostic probe) are what's currently on
disk at `SecuritasAutomation/Securitas-Invoice-Tracker_MegaStack.vb` and
`JCI-invoice-tracker/JCI-Invoice-Tracker_MegaStack.vb`, ready for the user to preview safely
themselves once the live workbooks are restored, or via a scratch-copy workbook instead of the
live ones.

## Stack builds

Both ran clean:

```
SecuritasAutomation> Stack-VBFiles.ps1
  variant modules : 22
  core modules    : 51
  total stacked   : 73
  EXIT 0

JCI-invoice-tracker> Stack-VBFiles.ps1
  variant modules : 11
  core modules    : 51
  total stacked   : 62
  EXIT 0
```

No duplicate procedure names in either stack (checked by counting every `Sub`/`Function`
declaration name in each stacked file; every count is 1). New Private helper names
(`ProcessNewBillsBulk`, `ProcessNewBillsPerCell`, `NumberFormatForSourceHeader`) appear exactly
once each, only in the Securitas stack (as expected — `ProcessNewBills.vb` is Securitas-local,
not core).

## Test suite results

All ten suites in `InvoiceTrackerCore/tests/`, including the three added by this fix, pass:

```
Test-BillCodeScreen.ps1                    13 passed
Test-DocumentModule.ps1                    15 passed
Test-ExtractInvApprovalDate.ps1  (new)     12 passed
Test-FiscalCalendar.ps1                   643 passed
Test-HashFile.ps1                           6 passed
Test-MirrorInvoiceBlock.ps1                42 passed
Test-MoveProcessedFile.ps1                 10 passed
Test-NormalizeIdentifierColumns.ps1        51 passed
Test-PopulateInvoiceTypeForNewRows.ps1(new) 7 passed
Test-ProcessedBatchLog.ps1                 15 passed
Test-ProcessNewBillsBulk.ps1     (new)     30 passed
Test-ReqJoinKey.ps1                         2 passed
```
846 assertions, 0 failures.

## Commits

- `InvoiceTrackerCore`, branch `fix/cell-by-cell-writes`: `ExtractInvApprovalDate.vb` bulk
  rewrite + filter guard, three new test files, this report.
- `SecuritasAutomation`, branch `fix/cell-by-cell-writes`: `ProcessNewBills.vb` bulk rewrite +
  filter guard, `AddNewBills.vb`'s `PopulateInvoiceTypeForNewRows` bulk rewrite + filter guard,
  rebuilt `Securitas-Invoice-Tracker_MegaStack.vb`, ticket update.
- `JCI-invoice-tracker`, branch `fix/cell-by-cell-writes`: `AddNewBills.vb` main write loop
  bulk rewrite + filter guard + rollback-timing fix, rebuilt
  `JCI-Invoice-Tracker_MegaStack.vb`.

Exact SHAs recorded in the ticket update (`SecuritasAutomation/docs/core-extraction/tickets.md`).
