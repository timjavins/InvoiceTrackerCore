# Batch/Workbook Tracking + Monitoring-Aware Duplicate Check — Design

Status: approved by user, ready for implementation plan
Repos touched: `InvoiceTrackerCore` (shared), `SecuritasAutomation` (variant)
Not touched: `JCI-invoice-tracker` (tenant config additions only; no behavior change)

## Problem

Two related gaps in the bill-ingestion flow (`AddNewBills` → `ProcessNewBills` →
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

## Scope

- Applies to Securitas and JCI equally for workbook/batch tracking (goes in
  core, gated by tenant config).
- The monitoring-aware duplicate-check fix is Securitas-only — JCI has no
  monitoring bills and keeps its own `MarkDuplicateInvoices` unchanged.
- Checks run against the **working tracker workbook only** (e.g.
  `2026 SECURITAS bills.xlsm`), never the All-Years archive.
- Out of scope: any change to JCI's duplicate logic, any change to the PDF
  ingestion sidecar, any change to Coupa export handling.

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
  don't get inserted, so there's nothing to delete bottom-up afterward
  (removes the `:199-217` cleanup path in `AddNewBills.vb`).
- End-of-run summary (`MsgBox`, matching the existing Coupa-warnings report
  style) lists bill codes added vs. skipped, and why each skip happened
  (already processed vs. repeat-in-file).

JCI's `MarkDuplicateInvoices.vb` is unchanged.

## Error Handling

- Hash failure (file locked, `certutil` unavailable/blocked) → abort before
  `Workbooks.Open`; clear message; nothing opened or written.
- Missing BILL CODE / REQ # columns → same guard `CheckNewBillsForDuplicates`
  already has today, just runs earlier in the sequence.
- File-move failure after a successful batch → warn, continue; batch is
  already correctly recorded regardless of whether the move succeeded.
- Hidden log sheet absent (older tracker workbook copy, first rollout) →
  created on first write.

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

## Commit Attribution

No AI attribution / Co-Authored-By lines, per
`SecuritasAutomation/docs/core-extraction/tickets.md` standing rule and the
user's global instruction.
