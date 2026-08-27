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

## Design docs

Architecture decisions and the domain glossary live in `SecuritasAutomation`:

- `CONTEXT.md` — domain glossary
- `docs/adr/0001` — the two invoice-number identities
- `docs/adr/0002` — why core is an independent sibling folder
- `docs/adr/0003` — TenantConfig as the single narrow interface
- `docs/adr/0004` — the shadow rule
- `docs/adr/0005` — the requester is the site's responsible party
