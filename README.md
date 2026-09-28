# InvoiceTrackerCore

Shared modules for the invoice tracker family. Consumed by variant repos
(`SecuritasAutomation`, `JCI-invoice-tracker`) at assembly time.

This repo holds no per-supplier facts. Everything that varies between suppliers is
declared by each variant in its own `TenantConfig.vb`.

## How consumption works

VBA has no linker or package manager. `Stack-VBFiles.ps1` is the module system: it
concatenates core `.vb` files and the variant's own `.vb` files into a single stacked
module the workbook imports.

Each variant keeps a copy of `Stack-VBFiles.ps1` in its root and runs it from there.
The canonical copy lives here — when it changes, copy it out to the variants.

Expected layout:

```
<parent>/
├── InvoiceTrackerCore/       ← this repo
├── SecuritasAutomation/
└── JCI-invoice-tracker/
```

Override the default with `-CorePath` if your layout differs.

## The shadow rule

When a variant file and a core file share a filename, **the variant wins** and the core
file is dropped from the stack. That is how a variant overrides a core module.

VBA has no namespaces, so stacking two files that both define `Sub Foo()` fails to
compile. Shadowing is the only safe override mechanism.

A shadow forks the module — it stops receiving core improvements. Prefer parameterizing
through `TenantConfig.vb`. Keep shadows rare and document them in the variant.

**Shadow paired modules together.** Some core modules share module-level state with a
sibling (`PauseThinking.vb` holds the saved state `RestoreThinking.vb` restores).
Shadowing only one half leaves the pair inconsistent — the override won't maintain the
state its sibling still depends on. Shadow both, or neither.

## Conventions for core modules

- **No `Option Explicit` outside `Header.vb`.** Every module is concatenated into one
  file, and VBA requires module-level options to precede all procedures. `Header.vb` is
  pinned first and owns them; an `Option Explicit` anywhere else lands mid-file and
  fails to compile.
- **Module-level declarations are hoisted for you.** VBA requires *every* module-level
  declaration before the first procedure, not just `Option Explicit` -- otherwise the build
  fails with "Only comments may appear after End Sub, End Function, or End Property". A
  module in the middle of the stack cannot satisfy that alone, so `Stack-VBFiles.ps1` lifts
  declarations into a header block at the top of the assembled file and reports which
  modules it moved. Declaring state in a core module is fine.
- **The stack is pasted into `ThisWorkbook`, a class module.** Avoid what VBA forbids as
  public members of an object module: `Public Const`, public fixed-size arrays,
  fixed-length strings, and `Declare`. `Private Const` and procedure-local `Const` are fine.
- **One module's filename must match what variants use**, since the shadow rule keys on
  filename. Splitting a core module into differently-named files defeats a variant's
  ability to override it.
- **Don't depend on the active sheet.** Take a worksheet parameter, or resolve one
  through `TenantSheetName()`.

## The Coupa Users roster

`CoupaUsers.vb` verifies every requester email against the active-Coupa-user report before it
reaches a flat file, because Coupa rejects an upload row whose requester is not a live user.

The roster is a **shared resource**, so it reaches each workbook the same way `BU List` does —
a Power Query against a SharePoint list on the Asset Protection PMO site, landing on a sheet
named by `TenantSheetName("coupa-users")` (currently `Coupa Users` in both variants).

Wiring up a variant workbook:

1. Publish the active-users report as a list on the Asset Protection PMO SharePoint site, so
   both trackers read one copy.
2. In the workbook: **Data → Get Data → From Online Services → From SharePoint Online List**,
   point it at `https://nordstrom.sharepoint.com/sites/AssetProtectionPMO`, pick the list, and
   load it to a new sheet named to match `TenantSheetName("coupa-users")`. `BU List`'s existing
   query is the template — same site, same `SharePoint.Tables` source, `ApiVersion = 15`.
3. Nothing else. Core finds the columns itself.

Core resolves columns **by header name**, not position, because a refresh reorders and renames
freely. It accepts several spellings for each of the two columns that matter — see
`EmailHeaderNames` and `StatusHeaderNames` in `CoupaUsers.vb`. A roster with no recognizable
status column is treated as already filtered to active users, which is what an "active users
report" normally is.

If the sheet is absent, empty, or unreadable, verification is **unavailable** and every address
passes, with the reason reported once in the run's summary. A workbook not yet wired up must
still be able to export. See ADR-0005.

## The Site RPs sheet

`RequesterEmail.vb` resolves a requisition's requester from the sheet named by
`TenantSheetName("site-rps")`, walking `Email` → `Tier 1 email` → `Tier 2 email` →
`Tier 3 email` and taking the first that is an active Coupa user.

That sheet is built by `docs/power-query/site-rps.m` — paste it into each tracker's Power
Query Advanced Editor and name the query `Site RPs`. It joins two shared sources:

- **`BU List`** (the SharePoint list, already queried in both trackers) — the store list,
  store name, and BU
- **`Store AP Staff Directory.xlsx` / `AP POCs`** on the same SharePoint site — tier 1/2/3
  contacts

`Vertical` is derived as BU `200` → `NS`, BU `250` → `NR`, else blank. Tier 1 is suppressed
for `NS` stores: for a full-line store the directory's tier 1 is the store-level AP person,
who is not the right requisition requester — tier 2 is. **That suppression is load-bearing.**
Remove it and every NS store starts raising requisitions for a different person.

### Why it is a query

It used to be a hand-typed store list plus XLOOKUP formulas into the directory. The links
drifted: JCI's resolved to the SharePoint copy while Securitas's was hardcoded to
`C:\Users\p4bn\Downloads\`, so the two trackers disagreed about the responsible party for
**100 stores** — invisibly, since nothing compared them. Shared facts fed by per-workbook
formulas will drift; shared facts fed by a query against a shared source cannot.

Adding a store now needs no worksheet editing in either tracker.

## The requester warnings report

`RequesterWarnings.vb` collects the data-quality problems a run meets while working out who each
requisition belongs to, and writes them to:

```
%USERPROFILE%\Downloads\Coupa requester warnings YYYYMMDD.xlsx
```

Same folder and same date shape as `Coupa uploads YYYYMMDD.csv`, so both artifacts of one export sit
together. One sheet, eight columns:

```
Tracker Row | Store | Bill Ref | Issue | Requester Used | Tier | Inactive Skipped | Detail
```

with a frozen header row and AutoFilter across the used range. `Issue` is a closed set so it is
worth filtering on:

| Issue | Meaning |
| --- | --- |
| `NoSiteRPsRow` | the store is not on `Site RPs`, so the bill was skipped |
| `NoContactAtAnyTier` | listed, but no email in any of the four columns |
| `AllTiersInactive` | addresses exist, but none is an active Coupa user |
| `EscalatedPastInactive` | exported fine, but not under the site's own responsible party |
| `StoreNotInBUList` | the store has no GL code, so the bill was skipped |

**A clean run writes no file at all**, and `WriteRequesterWarningsReport` returns `""`. That is
deliberate: an empty report, or a file left from last week, would be read as this run's verdict.
No file means nothing to fix.

These warnings used to be concatenated into the generators' "Error Summary" `MsgBox`. That dialog
now carries a single line naming the count and the path. Tenant-specific warnings — Securitas's
blank bill codes, `INVOICE TYPE` problems, and the capitalize-or-expense and project-detail answers
— stay in the dialog, because they are decisions the operator was just asked about rather than rows
missing from a shared list.

`RecordRequesterOutcome` decides whether an outcome is a warning at all and what to label it, so
both tenants classify identically. It classifies from `TryResolveSiteRequester`'s `outcome` code,
never from its `reason` prose — otherwise rewording a warning would silently change the
classification.

## The requisition join key

`LookupReqs.vb` backfills the tracker's requisition number once Coupa has assigned one, joining
tracker rows to the "Coupa Reqs" import sheet. `ReqJoinKeyColumn()` resolves which tracker column
carries the value to join on: it prefers `TenantColLetter("req-join-key")` and falls back to
`TenantColLetter("submitted-invoice-number")` when a tenant declares no `req-join-key`.

`req-join-key` is for a tenant that raises a purchase order before any supplier document exists,
and so has no submitted invoice # to join on — it mints its own key instead. Reusing
`submitted-invoice-number` for that key would recreate the identity ambiguity ADR-0001 exists to
prevent.

The fallback is the compatibility guarantee: `SecuritasAutomation` and `JCI-invoice-tracker` need
no change to keep joining on the column they always used. Both now declare `req-join-key`
explicitly anyway, each naming its own existing submitted-invoice-number column — `O` for
Securitas, `J` for JCI — so the declared value and the fallback agree for both today.

A tenant whose `req-join-key` entry is missing or broken degrades quietly to the
`submitted-invoice-number` fallback rather than failing loudly: every tenant's `TenantColLetter`
raises the same error for an unknown concept, and core has no way to tell that apart from a real
defect in the tenant's config.

## Normalizing identifier columns

`NormalizeIdentifierColumns.vb` keeps REQ #, invoice numbers, and payment # as text on the
tracker sheet, the same way `ConvertStoreNumbers.vb` zero-pads STORE #. Every Coupa-side sheet
(`Coupa Reqs`, `Coupa Invs`, `Coupa POs`) stores its key column as text; XLOOKUP/VLOOKUP/Match
never coerce a numeric key to match a text value, so a tracker identifier cell that lands as a
Number silently breaks every formula keyed on it. See
`notes/identifier-types-investigation.md` for the live-workbook evidence and root cause.

`NormalizeIdentifierColumnsOn(ws, cols)` sweeps a sheet/column list with a bulk read/compute/
write per column, the same shape `ConvertStoreNumbers.vb` uses: one read of `.Value` and one
read of `.Formula` over the whole column, everything decided in memory (a cell is a formula
cell when its `.Formula` text starts with `"="`), then written back with one `NumberFormat =
"@"` plus `Value = array` write per maximal contiguous run of rows that actually need a change.
Formula cells are never marked as needing a change, so they're never inside a written range;
an already-normalized text cell, a WARRANTY-style marker, and a blank/Null/Error cell are the
same — none of them gets a format or value write, not just an unchanged one. On a first sweep
of an unnormalized column every row needs writing, so the whole column is one run and this
still collapses to a single bulk write; later sweeps naturally shrink to writing only what
changed. There is deliberately no separate "whole range in one shot, no formula cells" branch,
because that would write `NumberFormat`/`Value` to every cell in the range including ones that
never needed it — an earlier draft of this rewrite did exactly that, and it was caught in
review because a co-author-visible format write is exactly what guarantee #4 exists to prevent.
It finds each column's last row by reading the whole `UsedRange` extent and scanning for the
last non-blank cell in memory,
because both `End(xlUp)` and `Find` walk visible cells only and would miss a row an AutoFilter
is hiding — the same trap `MirrorInvoiceBlock.vb` documents. That same AutoFilter hiding also
makes a multi-row bulk *write* unreliable (confirmed directly: a written array can land on the
wrong visible row while a hidden row keeps its stale value), so whenever the sheet has an
active filter the sweep falls back to one write per changed cell instead of a bulk range write.
It never raises and never shows a dialog: on error it returns whatever it already normalized.

`NormalizeIdentifierColumns(announce, manageProtection)` is the entry point Refresh calls: it
resolves the tracker sheet via `TenantSheetName("tracker")` and the columns via
`TenantIdentifierColumns()`, mapping each through `TenantColLetter`. **Every tenant that stacks
core must declare `TenantIdentifierColumns()`** (ADR-0003) — an empty array is a valid,
explicit no-op for a tenant that has none configured yet. `manageProtection` follows
`ConvertStoreNumbers`'s convention exactly, and each `Refresh` passes it the same way.

## The fiscal calendar

`FiscalCalendar.vb` computes Nordstrom's 4-5-4 retail fiscal calendar -- fiscal year
boundaries, week counts, month starts, and the reverse lookups from a date to its fiscal
year/week/month. It is pure computation: no worksheet, no workbook, no file, no network. That
matters because the calendar used to be looked up in a Finance-owned SharePoint workbook, which
meant every consumer inherited that file's availability and layout. This module instead derives
the whole calendar from one anchor rule -- fiscal year N ends on the Saturday closest to 31
January of year N+1 -- plus a repeating 4,5,4 week cycle per quarter, so nothing at run time
needs to read anything.

The rule is verified, not assumed: it reproduces all 46 published years, FY2005-FY2050, exactly,
checked against `fiscal-calendar-fixture.csv`. That fixture is a **test oracle, never a runtime
input** -- the module never reads it outside the test suite, so a missing or corrupted fixture
file cannot break a workbook in production, only weaken the tests that guard this module's next
change.

Run the two test suites from the repo root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-DocumentModule.ps1
```

The first asserts the calendar logic itself against the fixture, injected into a throwaway
standard module. The second re-injects the same module into `ThisWorkbook`, a document (class)
module -- the kind every tenant stack actually pastes it into -- because a document module
forbids things a standard module allows, and this module could otherwise pass every logic
assertion and still fail to compile in production. See `tests/README.md` for why there are two
scripts and what each one does and does not prove.

## Batch tracking

`HashFile.vb`, `ProcessedBatchLog.vb` and `MoveProcessedFile.vb` let a tenant's `AddNewBills`
refuse a bill file it has already processed. The file is identified by SHA-256 of its bytes
(`certutil`), so a renamed or re-sent copy is still recognised. See
`docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md`, section 1.

`HashFile` returns `""` on any failure and never raises. Callers treat `""` as "stop before
opening the file".

`ProcessedBatchLog.vb` keeps the log on a hidden sheet inside the working tracker, named by
`TenantProcessedBatchSheet()`, one row per processed file: hash, file name, tenant, time, rows
added, rows skipped. The sheet is created on first write. Every tenant that stacks core must
declare `TenantProcessedBatchSheet()` and `TenantProcessedFolder()` (ADR-0003).

`MoveProcessedFile.vb` moves a finished file into the subfolder `TenantProcessedFolder()` names,
beside the file. A name collision gets a timestamp suffix. The move is cosmetic: an empty folder
name disables it, and a failed move is reported but changes nothing, since the hash log is the
record.

## Screening bill codes before import

`BillCodeScreen.vb`'s `ClassifyBillCodes` marks each incoming row `already-processed` (its code
is in the tracker's existing rows), `repeat-in-file` (its code is on more than one incoming row,
and the caller disallows that), or `""` (write it). It runs **before** any row is written, so
skipped rows never land and nothing has to be marked red or deleted. It knows nothing about bill
types: Securitas passes `allowRepeatsInFile = True` for monitoring bills, which legitimately put
one bill code on several store rows.

## Mirroring a tracker into an archive

`MirrorInvoiceBlock.vb` replaces the rows of an archive sheet that belong to the current year,
those whose trailing `Source File` tag is in `ownedTags`, with a values-only copy of a tracker
sheet, and never touches any other row. It refuses, writing nothing and returning `-1`, when:
- the headers differ, or the archive has a column after `Source File`;
- the source has no `BILL CODE` header;
- the source is empty;
- either sheet has an active filter (`FilterMode = True`) -- see "Filters" below.

A `-1` from one of those refusals means the archive is untouched, but `-1` can also come from an
error raised mid-way through the delete-then-append (an error during the row deletes or the
append itself). That case can leave the archive with owned rows already deleted and their
replacements not yet appended -- a partial write, not a no-op. Callers cannot tell the two apart
from the return value alone and must not assume `-1` means nothing changed; the next successful
run repairs any partial write, because it always replaces the whole owned block again.

Owned rows are deleted and re-appended rather than the sheet being rewritten, because writing
a legacy text value such as `0175` back into a General cell converts it to a number. The
Securitas `SyncAllYears` wrapper supplies the SharePoint side; see the batch-tracking spec,
section 3.

**Filters.** A sheet with an active filter (`FilterMode = True`, meaning rows are currently
hidden by filter criteria) is refused outright, before any delete -- neither sheet is read,
deleted, or written. A filter is never cleared or changed to work around this: the archive is
shared, so its view must stay exactly as the caller left it.

In practice this refusal is a backstop, not the normal path: the Securitas `SyncAllYears`
wrapper clears AutoFilter criteria on both sheets before calling this routine, the same way
`Refresh` does, so `FilterMode` is normally already false by the time `MirrorInvoiceBlock` runs.

That refusal exists because the more surgical fix was tried first and measured, not assumed.
`End(xlUp)` walks like Ctrl+Up and returns the last *visible* row on a filtered sheet, which
would truncate a filtered source or leave filtered-out legacy rows in the archive unreplaced
(with their replacements appended on top of them). Swapping in
`Columns(col).Find(What:="*", ..., SearchDirection:=xlPrevious)` looked like the fix, on the
assumption that `Find` walks actual cell order and ignores hidden rows. Testing against a real
AutoFilter (`tests/Test-MirrorInvoiceBlock.ps1`) showed that assumption is wrong: `Find` skips
rows an AutoFilter is hiding just like `End(xlUp)` does, so it reproduced the same corruption
instead of fixing it. (`Find` *does* find a row hidden manually, via Format > Hide, unlike
`End(xlUp)` -- which is why `MibLastRow` still uses `Find` rather than reverting, even though it
cannot be trusted alone under an AutoFilter.) Refusing on `FilterMode` is the safe alternative
to a fix that only looks correct.

## Design docs

Architecture decisions and the domain glossary live in `SecuritasAutomation`:

- `CONTEXT.md` — domain glossary
- `docs/adr/0001` — the two invoice-number identities
- `docs/adr/0002` — why core is an independent sibling folder
- `docs/adr/0003` — TenantConfig as the single narrow interface
- `docs/adr/0004` — the shadow rule
- `docs/adr/0005` — the requester is the site's responsible party

Notes kept in this repo:

- `docs/selenium-late-binding.md` — why the `Selenium` reference should go away, and the
  "Selenium is not installed" guard that cannot currently fire
