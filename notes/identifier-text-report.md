# Identifier text normalization -- implementation report

Implements the design from `notes/identifier-types-investigation.md`: identifier columns
(REQ #, invoice numbers, payment #) are now swept to text on every `Refresh`, the same way
`ConvertStoreNumbers` zero-pads STORE #.

## Changes, by file:line

### InvoiceTrackerCore (`feat/identifier-text`)

- `NormalizeIdentifierColumns.vb` (new file, 197 lines)
  - `NormalizeIdentifier(v)` -- line 23. Pure. Whole-number numeric -> `Format$(v, "0")`;
    non-whole numeric -> `CStr(v)`; string -> `Trim$`; Empty/Null/error -> unchanged.
  - `NormalizeIdentifierColumnsOn(ws, cols)` -- line 58. Sweeps an array of column letters,
    finds each column's last row via the full `UsedRange` extent (filter-proof), skips
    `HasFormula` cells, sets `NumberFormat = "@"` before rewriting a changed/numeric cell,
    leaves already-normalized text untouched, returns the count changed. `On Error GoTo Done`
    at the top means any runtime error returns whatever was already normalized rather than
    raising.
  - `LastNonEmptyRow(ws, colLetter, usedLastRow)` -- line 129 (private helper).
  - `NormalizeIdentifierColumns(announce, manageProtection)` -- line 168. Resolves
    `TenantSheetName("tracker")` and `TenantIdentifierColumns()` (mapped through
    `TenantColLetter`), wraps in `UnprotectSheet`/`ProtectSheet` when `manageProtection`, is a
    no-op for an empty accessor array.
- `LookupReqs.vb:150` -- added `wsTracker.Range(colTrackerReq & "2:" & colTrackerReq &
  lastRow).NumberFormat = "@"` immediately before the existing `.Value = trackerReqNums`
  write (line 151), with a comment explaining why: a General cell re-numericizes a
  numeric-looking string on write, matching `CopyPaymentNums.vb` and `UpdateSearchValues.vb`.
- `README.md:188` -- new "Normalizing identifier columns" section, next to the fiscal-calendar
  section (there was no pre-existing `ConvertStoreNumbers` section to sit beside; this is
  the nearest identifier-related section, "The requisition join key", immediately above it).
  States every tenant must declare `TenantIdentifierColumns()` (ADR-0003).
- `tests/Test-NormalizeIdentifierColumns.ps1` (new file) -- see TDD evidence below.

### SecuritasAutomation (`feat/identifier-text`)

- `TenantConfig.vb:127-131` -- `TenantIdentifierColumns()` returns
  `Array("submitted-invoice-number", "coupa-invoice-number", "requisition-number",
  "payment-number")`. `purchase-order-number` (U) deliberately omitted: it is a pure formula
  column in `WriteFormulas_Tracker.vb`, and the sweep already skips formula cells, so listing
  it would do nothing.
- `Refresh.vb:41-46` -- `Call NormalizeIdentifierColumns(False, True)` added right after the
  existing `ConvertStoreNumbers(False, True)` call (line 39), i.e. after `LookupReqs` (line 30)
  and before `WriteFormulas_Tracker` (line 45 pre-edit / 48 post-edit). `manageProtection:=True`
  matches how `ConvertStoreNumbers` is passed in this same `Refresh`.

### JCI-invoice-tracker (`feat/identifier-text`)

- `TenantConfig.vb:119-121` -- `TenantIdentifierColumns()` returns
  `Array("submitted-invoice-number", "requisition-number")`. `coupa-invoice-number` not listed
  separately -- column J doubles as both concepts here (ADR-0001), listing it twice would sweep
  the same column twice for nothing. PO # (N) and PAYMENT (X) are pure formula columns, skipped
  for the same reason as Securitas's U.
- `Refresh.vb:35-41` -- `Call NormalizeIdentifierColumns(False, False)` added between
  `LookupReqs(True, False)` (line 33) and `WriteFormulas_Tracker(ActiveSheet)` (line 43).
  `manageProtection:=False` matches how `ConvertStoreNumbers` is passed later in this same
  `Refresh` (line 47) -- the sheet is already unprotected once around the whole sequence.
- `WriteFormulas_Tracker.vb:48-53` -- formula fix, separate commit (see below).

**Exact before/after of the JCI formula (`WriteFormulas_Tracker.vb`, column M / REQ STATUS):**

Before:
```
formulasM(i - 1, 1) = "=XLOOKUP(L" & i & ",'Coupa Reqs'!A:A,'Coupa Reqs'!D:D,IF(ISNUMBER(S" & i & "),""DELETED?"",""""))"
```

After:
```
formulasM(i - 1, 1) = "=XLOOKUP(L" & i & ",'Coupa Reqs'!A:A,'Coupa Reqs'!D:D,IF(ISNUMBER(L" & i & "),""DELETED?"",""""))"
```

Only the `ISNUMBER(S...)` -> `ISNUMBER(L...)` reference changed. Per the investigation
(sections 3 and 6), JCI's REQ # column is `L`; `S` is `Coupa INV Date`, an unrelated column left
over from copying Securitas's formula (where the equivalent fallback correctly tests `S`,
Securitas's own REQ # column -- confirmed directly:
`SecuritasAutomation/file ingesting/WriteFormulas_Tracker.vb:51`,
`=XLOOKUP(S...,IF(ISNUMBER(S...),"DELETED?",""))`).

### NordGuardsTracker (`po-key-and-core-adoption`)

- `TenantConfig.vb:109-115` -- `TenantIdentifierColumns()` returns `Array()` with the comment
  "no identifier normalization configured yet", so this stack compiles against core's new
  requirement.

## TDD evidence (RED then GREEN)

`tests/Test-NormalizeIdentifierColumns.ps1` was written alongside the implementation. To
produce a genuine RED, the `cell.NumberFormat = "@"` line in the numeric-write branch of
`NormalizeIdentifierColumnsOn` was temporarily removed and the suite re-run:

```
FAIL A2 becomes a text-typed value -- expected 'String', got 'Double'
FAIL A2 gets @ format before the write -- expected '@', got 'General'
FAIL C2 becomes text too -- multiple columns are swept -- expected 'String', got 'Double'
FAIL the hidden row is normalized to text -- expected 'String', got 'Double'
FAIL the hidden row gets @ format too -- expected '@', got 'General'

passed: 20  failed: 5
```

This is exactly the bug class the whole task exists to prevent: writing a numeric-looking
string into a General-formatted cell re-numericizes it. The line was restored (diffed
byte-identical against the pre-break copy) and the suite re-run, GREEN:

```
passed: 25  failed: 0
ALL PASS
```

## Suite outputs (all core suites, InvoiceTrackerCore)

```
=== tests/Test-BillCodeScreen.ps1 ===
passed: 13  failed: 0
ALL PASS
exit=0
=== tests/Test-DocumentModule.ps1 ===
passed: 15  failed: 0
ALL PASS
exit=0
=== tests/Test-FiscalCalendar.ps1 ===
passed: 643  failed: 0
ALL PASS
exit=0
=== tests/Test-HashFile.ps1 ===
passed: 6  failed: 0
ALL PASS
exit=0
=== tests/Test-MirrorInvoiceBlock.ps1 ===
passed: 42  failed: 0
ALL PASS
exit=0
=== tests/Test-MoveProcessedFile.ps1 ===
passed: 10  failed: 0
ALL PASS
exit=0
=== tests/Test-NormalizeIdentifierColumns.ps1 ===
passed: 25  failed: 0
ALL PASS
exit=0
=== tests/Test-ProcessedBatchLog.ps1 ===
passed: 15  failed: 0
ALL PASS
exit=0
=== tests/Test-ReqJoinKey.ps1 ===
passed: 2  failed: 0
ALL PASS
exit=0
```

All nine suites exit 0, including `Test-ReqJoinKey.ps1` (confirms the `LookupReqs.vb` edit did
not disturb `ReqJoinKeyColumn`) and the new suite.

## Stack builds

All four stacks rebuilt with each repo's own `Stack-VBFiles.ps1` (`< /dev/null` to satisfy the
trailing `Read-Host`). All exited 0, no duplicate-filename errors, no shadowing warnings related
to this work.

- **SecuritasAutomation**: 22 variant + 51 core = 73 modules stacked, exit 0.
- **JCI-invoice-tracker**: 11 variant + 51 core = 62 modules stacked, exit 0.
- **NordGuardsTracker**: 13 variant + 51 core = 64 modules stacked, exit 0.

Confirmed each new name is defined exactly once per assembled stack (`grep -c` against each
`*MegaStack*.vb`):

| Name | Securitas | JCI | NordGuards |
|---|---|---|---|
| `NormalizeIdentifier` | 1 | 1 | 1 |
| `NormalizeIdentifierColumnsOn` | 1 | 1 | 1 |
| `NormalizeIdentifierColumns` (Sub) | 1 | 1 | 1 |
| `TenantIdentifierColumns` | 1 | 1 | 1 |

`*MegaStack*.vb` files are `.gitignore`d in all three variant repos (confirmed via `grep -i
megastack .gitignore`), so none were staged or committed.

## Commits per repo

**InvoiceTrackerCore** (`feat/identifier-text`):
- `41e1565` `feat: normalize tracker identifier columns to text` -- `NormalizeIdentifierColumns.vb`, `tests/Test-NormalizeIdentifierColumns.ps1`, `README.md`
- `76062be` `fix(LookupReqs): set text format before writing REQ #` -- `LookupReqs.vb`

**SecuritasAutomation** (`feat/identifier-text`):
- `05fe28f` `feat: wire NormalizeIdentifierColumns into Refresh` -- `Refresh.vb`, `TenantConfig.vb`

**JCI-invoice-tracker** (`feat/identifier-text`):
- `316752d` `feat: wire NormalizeIdentifierColumns into Refresh` -- `Refresh.vb`, `TenantConfig.vb`
- `32bcce3` `fix(WriteFormulas_Tracker): DELETED? tests REQ # column` -- `WriteFormulas_Tracker.vb` (separate commit, as required)

**NordGuardsTracker** (`po-key-and-core-adoption`):
- `2e47438` `chore(TenantConfig): declare TenantIdentifierColumns` -- `TenantConfig.vb`

Nothing pushed, nothing merged, no branch switches.

## Concerns

- **JCI's DELETED? fix is untested against live data.** The formula change is mechanically
  correct (confirmed against the investigation notes and Securitas's original), but it was not
  re-verified against the live JCI workbook -- the harness suite has no coverage for
  `WriteFormulas_Tracker`'s formula strings (they are plain string concatenation, not something
  the VBA test harness exercises), and I did not open or write to the live cloud workbook, per
  the safety constraint. Once this ships, the ~124 `DELETED?` cells the investigation flagged
  should be spot-checked again after the next `Refresh`.
- **`NormalizeIdentifierColumns` runs unconditionally on every `Refresh`, unlike
  `MigrateColumnTypes`, which is one-time and manual.** This is by design (the investigation's
  explicit recommendation, to close the gap where `LookupReqs` could keep reintroducing numeric
  cells even after a cleanup), but it means every `Refresh` now does one more full-column sweep
  per declared identifier concept. Given the column sizes in the investigation (JCI ~458 rows,
  Securitas ~3446 rows), this is inexpensive, but it is a new steady-state cost that did not
  exist before.
- **NordGuardsTracker's `TenantIdentifierColumns()` is a placeholder** (`Array()`), per the
  task's explicit instruction -- this repo's live data was never audited for the same
  numeric-drift problem. If it ever gains real identifier columns, someone needs to revisit this
  the way `TenantFormulaFormats()` already flags as future work in that same file.
- **Neither the new module nor `LookupReqs.vb`'s fix was exercised against a live or
  representative-sized workbook.** All verification is the harness's scratch-workbook fixtures
  (small, synthetic) plus the stack builds' pure compile/duplicate-name check. Per the safety
  constraint, no workbook -- live or otherwise -- was opened, attached to, or written by this
  work; only the core test harness's own hidden Excel instances ran any of this code.
