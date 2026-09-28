# Batch/Workbook Tracking, Monitoring-Aware Duplicate Check, All-Years Mirror — Design

Status: implemented (see plan and tickets.md)
Repos touched: `InvoiceTrackerCore` (shared), `SecuritasAutomation` (variant), `JCI-invoice-tracker` (variant, partial)
`JCI-invoice-tracker` gets the hash gate/record/move (Feature 1) only, via its own TenantConfig
declarations and `AddNewBills.vb` changes; it has no All-Years archive (Feature 3) and its
`MarkDuplicateInvoices` (its analogue of Securitas's monitoring-aware duplicate check, Feature 2)
is unchanged.

## Problem

Three related gaps in the bill-ingestion flow (`AddNewBills` → `ProcessNewBills` →
`CheckNewBillsForDuplicates`, all in `SecuritasAutomation/file ingesting/`):

1. **No tracking of whether a source workbook/file has already been processed.**
   Only bill/invoice numbers are tracked, inside the target sheet. Nothing records
   that a given source file was run through `AddNewBills`, so an accidental
   re-run of the same file isn't caught until (and unless) its bill codes collide.

2. **The existing duplicate check is wrong for monitoring bills.** Monitoring
   bills legitimately repeat the same BILL CODE across multiple rows (one
   invoice covering several stores). `CheckNewBillsForDuplicates`
   (`file ingesting/CheckNewBillsForDuplicates.vb`) runs `CountIf` over the
   whole BILL CODE column including the rows just inserted, so every monitoring
   row after the first gets wrongly marked `DUPLICATE`. Repair and installation
   bills are correctly one-bill-code-per-row and should keep being rejected on
   any repeat, whether against existing rows or within the same file.

3. **The All-Years archive is a frozen snapshot, and it is wrong for monitoring
   bills.** `Securitas All-Years Invoices - Consolidated.xlsm` (same SharePoint
   folder as the working book) was built once, 2026-08-31 to 2026-09-03, by
   consolidating eight legacy sources, deduped to one row per bill code
   (newest source wins), pasted values-only. Nothing has written to it since.
   Because it kept one row per bill code, every monitoring bill lost all but one
   of its store rows: BILL CODE `6005463095` has 9+ rows in the working book and
   1 in All-Years (verified 2026-09-26). The only record of that migration is
   `Excel-MCP/gaps-and-fixes.md`.

## Scope

- Applies to Securitas and JCI equally for workbook/batch tracking (goes in
  core, gated by tenant config).
- The All-Years mirror (§3) is Securitas-only: a pure routine in core, called by
  a thin Securitas wrapper. JCI has no All-Years workbook yet.
- The monitoring-aware duplicate-check fix is Securitas-only — JCI has no
  monitoring bills and keeps its own `MarkDuplicateInvoices` unchanged.
- Duplicate and already-processed checks (§1, §2) read the **working tracker
  workbook only** (e.g. `2026 SECURITAS bills.xlsm`), never All-Years. All-Years
  is written by §3, never read for decisions.
- Out of scope: any change to JCI's duplicate logic, any change to the PDF
  ingestion sidecar, any change to how Coupa exports are parsed. (§3 adds a step
  *after* each `UpdateCoupaData` branch; it does not change the import.)
- Out of scope, follow-on tickets (see end): redoing the consolidation so
  monitoring rows are preserved; building JCI's All-Years workbook.

## Design

### 1. Workbook/batch tracking (core, both tenants)

**Purpose:** (a) block an accidental re-run of a file already fully processed,
and (b) keep an audit trail of what batches ran.

**Mechanism:** content hash is the source of truth; moving the file afterward
is cosmetic only and never gates logic.

- `HashFile(path As String) As String` (new core module) — shells to
  `certutil -hashfile "<path>" SHA256`, parses the hash from output. Called on
  the picked file's path before `Workbooks.Open`.
- `ProcessedBatchLog` (new core module):
  - `IsAlreadyProcessed(hash As String) As Boolean`
  - `RecordProcessedBatch(hash, sourceFilename, tenant, billsAdded, billsSkipped)`
  - Both operate on a hidden sheet inside the **working tracker workbook**,
    one row per batch: hash, source filename, tenant, timestamp, bills added,
    bills skipped. If the sheet doesn't exist yet, create it on first write —
    no manual migration step.
- New `Tenant*()` accessors, defined per-variant in each `TenantConfig.vb`
  (Securitas and JCI):
  - `TenantProcessedBatchSheet() As String` — hidden sheet name.
  - `TenantProcessedFolder() As String` — folder to move the source file into
    after a successful run. Empty string disables the move (tracking still
    happens; only the cosmetic move is skipped).

**Flow change in `AddNewBills`:**

1. User picks file → `HashFile` computes hash before opening.
2. `IsAlreadyProcessed(hash)` — if true, stop before opening the workbook.
   Message names the prior run's timestamp and bill count. No partial state.
3. Otherwise proceed with open/layout-detection/header-lookup as today.
4. After the duplicate check (§2) and row writes complete, call
   `RecordProcessedBatch`.
5. If `TenantProcessedFolder()` is non-empty, attempt to move/rename the
   source file there. Failure (locked file, permissions) is logged/warned but
   does **not** roll back the batch record — the batch already happened.

### 2. Monitoring-aware duplicate check (Securitas only)

Replace the post-write, whole-column `CountIf` in
`CheckNewBillsForDuplicates.vb` with a pre-write check that knows the bill
layout (repair / installation / monitoring, already detected by
`FindBillHeaderRow`, `file ingesting/AddNewBills.vb:272`).

For each candidate bill code in the incoming file, decide before any rows are
written:

| Check | Repair / Installation | Monitoring |
|---|---|---|
| Code already present in rows `1…targetLastRow` of the working sheet | reject | reject |
| Code repeats across rows within the incoming file itself | reject | **allowed** |

- "Reject" means: that row (repair/installation) or that whole group of rows
  sharing the code (monitoring "bill") is skipped — not written, not marked
  red, not left for manual cleanup.
- Rows that pass are written normally.
- No red-cell/`REQ #` = `"DUPLICATE"` marking anymore — rejected rows simply
  don't get inserted. `CheckNewBillsForDuplicates.vb` is deleted. The
  `failureState` rollback at `AddNewBills.vb:199-217` **stays**: it undoes a
  failed import, which is unrelated to duplicates.
- A repair or installation code that repeats within the file skips **every**
  row carrying it, because the tool cannot tell which copy is right.
- The classification is a pure function in core, `ClassifyBillCodes` in
  `BillCodeScreen.vb`. It takes arrays in, returns a status per row, and knows
  nothing about bill types. Only Securitas calls it, and the decision that
  monitoring may repeat stays in `AddNewBills`. Keeping it pure lets the core
  harness test it.
- End-of-run summary (`MsgBox`, matching the existing Coupa-warnings report
  style) lists bill codes added vs. skipped, and why each skip happened
  (already processed vs. repeat-in-file).

JCI's `MarkDuplicateInvoices.vb` is unchanged.

### 3. All-Years mirror (Securitas only)

**Purpose:** All-Years is a backup of the working book and the cross-year
lookup surface. It must receive every new bill and every Coupa-column change.

**Mechanism: a whole-block mirror, not row matching.** Each sync replaces
All-Years' entire current-year block with a fresh values-only copy of the
working book's `Invoices` sheet. No row is matched to another row.

Row matching was considered and rejected. BILL CODE is not unique for
monitoring bills (N store rows share one code in both workbooks), so a
key-based refresh would write one store's row over all N. There is also no
"rows just updated" set to match from: most Coupa columns (T–AC) are live
`XLOOKUP`s (`WriteFormulas_Tracker.vb:51-59`) that recalculate on every row
whenever a Coupa sheet is replaced. A whole-block rewrite avoids both problems.
Because each sync produces the same output from the same input, it can safely
be re-run, and it recovers from any earlier failed sync.

**Block ownership.** A row in All-Years belongs to the mirror when its
`Source File` value is in the owned-tag set:

- `"2026:Invoices"`: the tag the mirror writes (follows the `2025:Invoices`
  convention).
- `"2026Model:Invoices"`: the one-off label the migration script gave to rows
  taken from the 2026 working book. The mirror claims these too, so the first
  sync replaces them rather than duplicating them.

Rows with any other tag (`2025:Invoices`, `RootNoYear:Invoices`, …) are never
touched.

**Algorithm** — `MirrorInvoiceBlock(src As Worksheet, dst As Worksheet,
writeTag As String, ownedTags As Variant) As Long`, a new core module:

1. Compare header rows by name. `dst` must have every `src` header, in the same
   order, plus `Source File`. On mismatch, raise nothing, return `-1`, and write
   nothing. The caller warns. Find `Source File` by header name, not position.
2. Find each sheet's real last row with `Cells(Rows.Count, <BILL CODE
   col>).End(xlUp)`, never `UsedRange`. All-Years' `UsedRange` is stale: it ends
   at row 6611, while the data ends near 6288.
3. Delete every `dst` row whose tag is in `ownedTags`. Work bottom-up, one
   `Rows("a:b").Delete` per contiguous run. Owned rows are interleaved with
   legacy rows (the migration did not sort by source), so this is many small
   deletes, not one.
4. Append every `src` data row below the new last row. First set each
   destination column's `NumberFormat` from the source column's first data
   row. Then write `src.Value` as one array and fill `Source File` with
   `writeTag`.
5. Return the number of mirrored rows.

Legacy rows are never read back and rewritten. That is deliberate: writing a
text value such as `0175` into a General-format cell turns it into the number
175, and legacy cells were pasted without a guaranteed format. `.Value` is used
rather than `.Value2` so dates and currency keep their types.

The routine shows no `MsgBox` and calls no `Err.Raise`, so the test harness can
drive it.

**Values only.** Columns B–H and T–AC in the working book are formulas pointing
at `'BU List'`, `'Coupa Reqs'`, `Helper!` and so on. The mirror copies their
evaluated values, never their formulas. The migration built All-Years the same
way. All-Years' current-year rows are not recalculated there.

**The mirrored rows are read-only.** Edits belong in the working book, and any
hand edit to an owned row in All-Years is overwritten by the next sync. `PAID`
values in the working book reach All-Years as ordinary values.

**Why `PAID` matters.** Much of 2023–2024 was dirty: some invoices were paid
under different invoice numbers, so no lookup can find their payment. The user
investigated those by hand and replaced the lookup formulas with the literal
`PAID`. Those cells are the only record of the investigation. They sit in
non-owned rows, so the mirror never touches them. The separate, still-open
request to stop `WriteFormulas_*` from overwriting `PAID` cells is not part of
this spec.

**Securitas wrapper** — `SyncAllYears(Optional announce As Boolean)`, in
`SecuritasAutomation`:

1. Resolve the path by joining `ThisWorkbook.Path` with `TenantAllYearsWorkbookName()`
   (`/` for a cloud path, `\` otherwise) -- the archive lives in the same folder as
   the working book, so this is derived, never a hardcoded location; a scratch
   copy of the working book then only ever reaches an archive beside itself. If
   All-Years is already open in this instance, use that copy; otherwise open it
   with `Application.EnableEvents = False`, because All-Years carries a copy of
   the stack and its `Workbook_Open` must not fire.
2. If `wb.ReadOnly`, skip: the file is locked elsewhere, and writes would only
   reach memory.
3. Call `MirrorInvoiceBlock` with no `MsgBox` or other user prompt between
   reading and writing. That keeps the window for a concurrent edit as short as
   possible.
4. Save explicitly. Close only if the wrapper opened the file.
5. On success, stamp the time in the working book at
   `TenantAllYearsSyncStampCell()` (a `Helper` cell, same pattern as
   `TenantImportTimestampCell`).
6. On any failure, show one warning and leave the stamp unchanged. Never roll
   back or block the operation that called the sync.

**Call sites:**
- End of `AddNewBills`, after the `failureState` rollback check
  (`AddNewBills.vb:199-217`) and after the Feature 2 skip logic. The mirror
  reads the sheet, so it only ever sees rows that were actually written.
- End of each `UpdateCoupaData` branch: `requisitions`, `invoices` **and
  `orders`**. The `orders` branch changes ORDER DATE and PO STATUS through
  formulas.
- End of `Refresh`, on success. It rewrites REQ # (`LookupReqs`), payment
  numbers and every formula column.
- A manual `SyncAllYears` macro, so a person can re-sync after an outage.

**How missed syncs are made visible.** Warn-and-skip is acceptable only because
missed syncs are visible and recover on the next success. At the start of
`AddNewBills` and `UpdateCoupaData`, if the sync stamp is older than
`TenantAllYearsStaleDays()` (Securitas: 3), show one warning naming the last
successful sync. Any later successful sync fixes all the drift, because it
rewrites the whole block.

**Concurrency: known limit.** Excel can't lock a co-authored workbook. If a
person edits All-Years between the read and the write, or two users sync at
once, one write can be lost at cell level. The mirror limits the damage: the
next sync rewrites the owned block exactly. Edits a person makes to non-owned
rows during a sync could still be lost. That risk is accepted, because
All-Years' non-owned rows are historical and not edited in normal use.

## Error Handling

- Hash failure (file locked, `certutil` unavailable/blocked) → abort before
  `Workbooks.Open`; clear message; nothing opened or written.
- Missing BILL CODE / REQ # columns → same guard `CheckNewBillsForDuplicates`
  already has today, just runs earlier in the sequence.
- File-move failure after a successful batch → warn, continue; batch is
  already correctly recorded regardless of whether the move succeeded.
- Hidden log sheet absent (older tracker workbook copy, first rollout) →
  created on first write.
- All-Years unreachable, read-only, or with mismatched headers → one warning;
  the calling operation completes; the sync stamp stays stale, so the staleness
  warning keeps showing until a sync succeeds.
- Feature 1 records a batch even if the All-Years sync after it fails. That is
  safe, because the next successful sync mirrors every row, including that
  batch's.

## Testing

- Core (`InvoiceTrackerCore/tests/`, existing `VbaHarness.psm1` pattern,
  hidden Excel over COM, run via
  `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-X.ps1`):
  - `Test-HashFile.ps1`
  - `Test-ProcessedBatchLog.ps1`
- Securitas: no test harness exists for `AddNewBills`'s path today. Add one
  covering the corrected duplicate-check table above (repair/installation
  reject-on-repeat, monitoring allow-repeat-within-file, both reject
  already-processed) since this change touches exactly that path.
- `Test-MirrorInvoiceBlock.ps1` (core, **required**): two sheets in one hidden
  workbook act as source and destination, so no two-workbook harness is needed.
  Cases:
  - a monitoring group of N rows sharing a code arrives as N rows;
  - non-owned rows are preserved, in order;
  - both owned tags are replaced;
  - re-running gives identical output;
  - shrinking the source clears leftover rows;
  - a header mismatch returns -1 and writes nothing;
  - a formula source cell arrives as its value.
- The Securitas wrapper (open, `ReadOnly`, events, save) can't be simulated
  against SharePoint. It is verified by a written manual checklist run against a
  **copy** of All-Years, never the live file.
- Per the existing harness constraint, tests must not exercise any path that
  can call `Err.Raise` (blocks the hidden COM instance on the error dialog).

## Tenant Config Additions

`TenantConfig.vb` in each variant gains:

```
TenantProcessedBatchSheet() As String   ' e.g. "ProcessedBatches" (hidden)
TenantProcessedFolder() As String       ' "" disables the cosmetic move
```

Both Securitas and JCI define these. Core ships no default (consistent with
ADR-0003 — a stack without a `TenantConfig.vb` doesn't compile).

Securitas only, read by its `SyncAllYears` wrapper rather than by core, so JCI
needs no change:

```
TenantAllYearsWorkbookName() As String   ' e.g. "Securitas All-Years Invoices - Consolidated.xlsm"
                                          ' (archive lives beside the working book; SyncAllYears
                                          ' derives the full path from ThisWorkbook.Path)
TenantAllYearsSyncStampCell() As String  ' "E6" on Helper (E2-E5 hold import stamps)
TenantAllYearsStaleDays() As Long        ' 3
TenantAllYearsOwnedTags() As Variant     ' Array("2026:Invoices", "2026Model:Invoices")
TenantAllYearsWriteTag() As String       ' "2026:Invoices"
```

The tags are year-bound. At year rollover, the new working book declares the
new year's write tag, and the previous year's block remains as non-owned
history.

## Documentation follow-up (doc-sync-core)

Adapter mappings already exist in both repos
(`SecuritasAutomation/CLAUDE.md:13-23`, `InvoiceTrackerCore/CLAUDE.md:7-18`).
On closing this work:

- Update `SecuritasAutomation/docs/core-extraction/tickets.md` (Coordination
  Board) to close out the tracking/dedup ticket(s).
- Update each repo's `README.md` (Capability Doc) in the same commit as the
  module changes it describes — required by existing repo convention.
- No new ADR expected — this doesn't change the seam design covered by
  ADR-0001–0005, but re-check that assumption once the plan is written; if
  the hidden-sheet log location turns out to need its own contract decision,
  add an ADR entry in `SecuritasAutomation/docs/adr/` (shared ADR series —
  do not start a separate series in core, per existing convention).

- Record the 2026-08-31 to 2026-09-03 All-Years consolidation on the
  Coordination Board, including its dedup key (one row per bill code, newest
  source wins) and the monitoring-row loss that key caused. Its only current
  record is `Excel-MCP/gaps-and-fixes.md`.

## Follow-on Tickets (not in this plan)

1. **Redo the All-Years consolidation** once §3 has shipped. Rebuild the
   non-2026 history from the eight legacy sources, keeping every monitoring
   store row: dedup on the whole bill (BILL CODE + source), not on BILL CODE
   alone. Then run one sync to lay the 2026 block on top. **Must carry across
   every hand-set `PAID` value** from the current All-Years and the sources.
   They record manual investigation that has no other record. Before replacing
   anything, diff the `PAID` count between the old and new builds.
2. **JCI All-Years workbook.** Inventory JCI's legacy sources, consolidate, then
   move the `SyncAllYears` wrapper into core and give JCI the accessors.

## Commit Attribution

No AI attribution / Co-Authored-By lines, per
`SecuritasAutomation/docs/core-extraction/tickets.md` standing rule and the
user's global instruction.
