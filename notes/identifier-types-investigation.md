# Identifier column type investigation (REQ # and siblings)

Read-only investigation. Nothing was changed in code or in the live workbooks. Scope: design
a normalization pattern for identifier columns on the `Invoices` sheet, modeled on
`ConvertStoreNumbers`, to be wired into `Refresh`.

Sources: `SecuritasAutomation` (main), `JCI-invoice-tracker` (main), `InvoiceTrackerCore` (main),
plus the two live cloud workbooks `2026-securitas-bills.xlsm` and
`jci-repair-installation-invoices.xlsm`, read via the Excel MCP tools only.

---

## 1. Invoices sheet column inventory

### Securitas (`2026-securitas-bills.xlsm`, sheet `Invoices`, header row 1, data rows 2-3447
(used range extends to 6611 but rows 3448+ are blank padding))

| Header | Col | Formula/Literal | Writer | Intended kind |
|---|---|---|---|---|
| STORE # | A | Literal | AddNewBills/ProcessNewBills (`NormalizeStoreNumber`), fixed up by `ConvertStoreNumbers` | identifier |
| BU | B | Formula: `VLOOKUP(A,'BU List'!...)` | WriteFormulas_Tracker | identifier (BU code) |
| REGION | C | Formula: `VLOOKUP(A,'BU List'!...)` | WriteFormulas_Tracker | free text |
| STORE NAME/ADDRESS/CITY/ST/ZIP | D-H | Formula: `VLOOKUP(A,'Store Directory'!...)` | WriteFormulas_Tracker | free text |
| INV DATE | I | Literal | AddNewBills/ProcessNewBills (`WriteDate`) | date |
| REQUEST DATE | J | Literal | AddNewBills/ProcessNewBills (`WriteDate`) | date |
| INVOICE # | K | Literal | AddNewBills/ProcessNewBills (`WriteIdentifier`) — this is the "coupa-invoice-number" concept | identifier |
| INVOICE TYPE | L | Literal | AddNewBills (`ResolveInvoiceType`/`PopulateInvoiceTypeForNewRows`) | free text (canonical enum) |
| TRANSACTION DETAILS | M | Literal | AddNewBills/ProcessNewBills | free text |
| DUE DATE | N | Literal | AddNewBills/ProcessNewBills (`WriteDate`) | date |
| BILL CODE | O | Literal | AddNewBills/ProcessNewBills (`WriteIdentifier`) — this is the "submitted-invoice-number" / req-join-key concept | identifier |
| SUBTOTAL/SALES TAX/TOTAL | P-R | Literal | AddNewBills/ProcessNewBills (`WriteMoney`) | money |
| REQ # | S | Literal (backfilled) | `LookupReqs` only fills blanks; also manual entry; markers `WARRANTY`/`DUPLICATE` | identifier |
| REQ STATUS | T | Formula: `XLOOKUP(S,'Coupa Reqs'!A:A,'Coupa Reqs'!D:D, IF(ISNUMBER(S),"DELETED?",""))` | WriteFormulas_Tracker | free text/status |
| PO # | U | Formula: `IF(OR(T="Draft",T="Pending Approval"),"", XLOOKUP(S,'Coupa Reqs'!A:A,'Coupa Reqs'!F:F,""))` — keyed on S (REQ #), not on U itself | WriteFormulas_Tracker | identifier |
| ORDER DATE | V | Formula: `XLOOKUP(U,'Coupa POs'!A:A,'Coupa POs'!B:B,"")` | WriteFormulas_Tracker | date |
| PO STATUS | W | Formula (Cancelled test + `XLOOKUP(U,'Coupa POs'!A:A,'Coupa POs'!C:C,"")`) | WriteFormulas_Tracker | free text |
| Total Invoiced | X | Formula: `XLOOKUP(Helper!B,'Coupa Invs'!B:B,'Coupa Invs'!E:E,"")` | WriteFormulas_Tracker | money |
| INV STATUS | Y | Formula (Cancelled test + `XLOOKUP(Helper!B,...F:F, XLOOKUP(U,...F:F,""))`) | WriteFormulas_Tracker | free text |
| Coupa INV Date | Z | Formula: `XLOOKUP(Helper!B,...G:G, XLOOKUP(U,...G:G,""))` | WriteFormulas_Tracker | date |
| INV Approval Date | AA | Literal | `ExtractInvApprovalDate` | date |
| Expected Pay | AB | Formula: `XLOOKUP(Helper!B,...J:J, XLOOKUP(U,...J:J,""))` | WriteFormulas_Tracker | date |
| PAY Date | AC | Formula: `XLOOKUP(Helper!B,...I:I, XLOOKUP(U,...I:I,""))` | WriteFormulas_Tracker | date |
| PAYMENT # | AD | Literal | `CopyPaymentNums` (`Application.Match(Helper!B, 'Coupa Invs'!B, 0)` → `'Coupa Invs'!K`) | identifier |
| NOTES/COMMENTS | AE | Literal | manual | free text |
| Relevant email/name | AF-AG | Literal | `RequesterEmail`/blocker logic | free text |

### JCI (`jci-repair-installation-invoices.xlsm`, sheet `Invoices`, header row 1, data rows 2-459)

| Header | Col | Formula/Literal | Writer | Intended kind |
|---|---|---|---|---|
| STORE # | A | Literal | AddNewBills (`WriteIdentifier`), fixed up by `ConvertStoreNumbers` | identifier |
| BU | B | Formula: `VLOOKUP(A,'BU List'!...)` | WriteFormulas_Tracker | identifier (BU code) |
| REGION, STORE NAME/ADDRESS/CITY/ST/ZIP | C-H | Formula (`VLOOKUP`) | WriteFormulas_Tracker | free text |
| INV DATE | I | Literal | AddNewBills (`WriteDate`) | date |
| INVOICE # | J | Literal | AddNewBills (`WriteIdentifier`) — this is submitted-invoice-number, coupa-invoice-number, AND req-join-key all at once (ADR-0001) | identifier |
| SUBTOTAL | K | Literal | AddNewBills (`WriteMoney`) | money |
| REQ # | L | Literal (backfilled) | `LookupReqs` fills blanks; markers `WARRANTY` | identifier |
| REQ STATUS | M | Formula: `XLOOKUP(L,'Coupa Reqs'!A:A,'Coupa Reqs'!D:D, IF(ISNUMBER(S),"DELETED?",""))` | WriteFormulas_Tracker | free text/status |
| PO # | N | Formula: `IF(L2="WARRANTY",L2, IF(OR(L="",M="Draft",M="Pending Approval"),"", XLOOKUP(L,'Coupa Reqs'!A:A,'Coupa Reqs'!F:F,"")))` | WriteFormulas_Tracker | identifier |
| ORDER DATE | O | Formula: `XLOOKUP(N,'Coupa POs'!A:A,'Coupa POs'!B:B,"")` | WriteFormulas_Tracker | date |
| PO STATUS | P | Formula (Cancelled test + `XLOOKUP(N,...C:C,"")`) | WriteFormulas_Tracker | free text |
| Uninvoiced | Q | Formula: `XLOOKUP(N,'Coupa POs'!A:A,'Coupa POs'!G:G,"")` | WriteFormulas_Tracker | money |
| INV STATUS | R | Formula (Cancelled test + `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!F:F,"")`) | WriteFormulas_Tracker | free text |
| Coupa INV Date | S | Formula: `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!G:G,"")` | WriteFormulas_Tracker | date |
| INV Total | T | Formula: `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!E:E,"")` | WriteFormulas_Tracker | money |
| INV Approval Date | U | Literal | `ExtractInvApprovalDate` | date |
| Expected Pay | V | Formula: `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!J:J,"")` | WriteFormulas_Tracker | date |
| PAY Date | W | Formula: `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!I:I,"")` | WriteFormulas_Tracker | date |
| PAYMENT | X | Formula: `XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!K:K,"")` | WriteFormulas_Tracker | identifier (JCI writes this by formula, unlike Securitas' literal `CopyPaymentNums`) |
| NOTES/COMMENTS, Relevant email/name | Y-AA | Literal | manual | free text |

Important asymmetry: JCI has **no separate BILL CODE column**. Column J (`INVOICE #`) plays all
three roles Securitas splits across K (coupa-invoice-number) and O (submitted-invoice-number /
req-join-key) — see `TenantConfig.vb` comments citing ADR-0001. JCI's PAYMENT (X) is a *formula*
column, not a literal copied by a `CopyPaymentNums`-style routine as in Securitas.

---

## 2. Stored-type counts for identifier columns (live workbooks, `read_range mode="values"`)

Counted by inspecting whether each value came back as a JSON number (numeric-typed cell) or a
JSON string (text-typed cell), confirmed against a known-numeric column (TOTAL) and a known-text
column (STORE #-style) to validate the tool distinguishes them correctly.

### Securitas — `Invoices`, real data rows 2-3447 (3446 rows; rows 3448-6611 are blank padding)

| Column | Number | Text | Blank | Notes |
|---|---|---|---|---|
| BILL CODE (O) | 0 | 3446 | 0 (of real rows) | all text |
| INVOICE # (K) | 0 | 3446 | 0 | all text |
| REQ # (S) | 0 | 3446 | 0 | all text, including markers like `WARRANTY`, `LP Tech` |
| PO # (U) | 0 | 3286 | 160 | all non-blank are text |
| PAYMENT # (AD) | 0 | 1400 | 2046 | all non-blank are text (some blanks are a literal `" "` space, from `ReplaceNullValues`/`CopyPaymentNums` semantics) |
| BU (B) | 0 | 3446 | 0 | formula result, text (e.g. `"250"`) |

**Securitas' `Invoices` identifier columns currently hold zero numeric contamination.** This
looks like the result of `MigrateColumnTypes.vb`, a one-time migration in core (see section 5)
already having been run against this workbook: it explicitly converts `store-number`,
`submitted-invoice-number`, `coupa-invoice-number`, `requisition-number`, `payment-number`, etc.
to text, cell-by-cell, always setting `NumberFormat = "@"` before assigning the value (the
detail that matters — see section 4).

`Coupa Reqs!A` ("Req #"), 5870 rows: **0 numbers, 5870 text.** Confirms
`WriteCoupaColumnFromSourceData`'s `IsIdentifierHeader` coercion (commit `f78dc69`, "fix: write
Coupa identifier columns as text on import", 2026-08-03) is working correctly on the source side
for this tenant.

### JCI — `Invoices`, rows 2-459 (458 rows, all populated)

| Column | Number | Text | Blank | Notes |
|---|---|---|---|---|
| REQ # (L) | **384** | 74 (54 real numbers-as-text + 20 `WARRANTY` markers) | 0 | **mixed** |
| INVOICE # (J) | **280** | 178 | 0 | **mixed** |
| PO # (N) | 0 | 54 | 404 | all non-blank are text (formula result) |
| PAYMENT (X) | 0 | 77 | 381 | all non-blank are text (formula result) |
| BU (B) | 0 | 458 | 0 | formula result, text |

Both mixed columns show the same shape: a contiguous **older block stored as Number**, and a
**contiguous newest block (the last ~54-178 rows) stored as Text**. The transition lines up with
the same 2026-08-03 core fix that cleaned Securitas' source-side import — but for JCI it only
protected *new* rows going forward; the ~280-384 rows already in the sheet before that date were
never converted, because `MigrateColumnTypes.vb` (the tool that fixed Securitas) is a manual,
opt-in script and there is no evidence it has been run against the JCI workbook.

Examples of the minority type:
- REQ # (L): row 268 (`914751`, Number) vs. row 434 (`"1166552"`, Text).
- INVOICE # (J): row 2 (`8059053607`, Number) vs. row 258 (`"8059136798"`, Text).

`Coupa Reqs!A` ("Req #"), 1666 rows: visually 100% text (every value quoted, including
compound/annotated values like `"1140207, 1140208"` and `"1129550 (Cancelled)"` on the sibling
`PO Number` column `F`, which by definition can only be text). `Coupa Reqs!F` ("PO Number"),
1666 rows: same — 100% text, confirmed by direct read.

---

## 3. Every lookup keyed on an identifier

### WriteFormulas_Tracker (both tenants) — formula text and what each key matches against

**Securitas** (`SecuritasAutomation/file ingesting/WriteFormulas_Tracker.vb:44-59`):
```
B: =IFERROR(VLOOKUP(A,'BU List'!$A:$D,4,0), VLOOKUP(A,'BU List'!$B:$D,3,0))
C: =VLOOKUP(A,'BU List'!A:E,5,0)
D-H: =VLOOKUP(A,'Store Directory'!$A:$M,n,0)
T: =XLOOKUP(S,'Coupa Reqs'!A:A,'Coupa Reqs'!D:D, IF(ISNUMBER(S),"DELETED?",""))
U: =IF(OR(T="Draft",T="Pending Approval"),"", XLOOKUP(S,'Coupa Reqs'!A:A,'Coupa Reqs'!F:F,""))
V: =XLOOKUP(U,'Coupa POs'!A:A,'Coupa POs'!B:B,"")
W: =IF(ISNUMBER(SEARCH("Cancelled",U)),"Cancelled", XLOOKUP(U,'Coupa POs'!A:A,'Coupa POs'!C:C,""))
X: =XLOOKUP(Helper!B,'Coupa Invs'!B:B,'Coupa Invs'!E:E,"")
Y: =IF(W="Cancelled","PO Cancelled", XLOOKUP(Helper!B,'Coupa Invs'!B:B,'Coupa Invs'!F:F, XLOOKUP(U,'Coupa Invs'!B:B,'Coupa Invs'!F:F,"")))
Z, AB, AC: same double-XLOOKUP shape as Y, against columns G/J/I of 'Coupa Invs'
```

**JCI** (`JCI-invoice-tracker/WriteFormulas_Tracker.vb:41-58`):
```
B-H: same VLOOKUP shape as Securitas, keyed on A
M: =XLOOKUP(L,'Coupa Reqs'!A:A,'Coupa Reqs'!D:D, IF(ISNUMBER(S),"DELETED?",""))   <- note: fallback tests S, not L (see below)
N: =IF(L2="WARRANTY",L2, IF(OR(L="",M="Draft",M="Pending Approval"),"", XLOOKUP(L,'Coupa Reqs'!A:A,'Coupa Reqs'!F:F,"")))
O: =XLOOKUP(N,'Coupa POs'!A:A,'Coupa POs'!B:B,"")
P: =IF(ISNUMBER(SEARCH("Cancelled",N)),"Cancelled", XLOOKUP(N,'Coupa POs'!A:A,'Coupa POs'!C:C,""))
Q: =XLOOKUP(N,'Coupa POs'!A:A,'Coupa POs'!G:G,"")
R: =IF(P="Cancelled","PO Cancelled", IF(J="","", XLOOKUP(J,'Coupa Invs'!B:B,'Coupa Invs'!F:F,"")))
S, T, V, W, X: same IF(J="","",XLOOKUP(J,'Coupa Invs'!B:B,...)) shape against columns G/E/J/I/K
```

Both trackers' `M`/`T` formula carries a fallback `IF(ISNUMBER(S),"DELETED?","")` that in
Securitas correctly tests the *same* identifier column being looked up (S = REQ #). In JCI this
is a stale copy from the Securitas formula string: JCI's own REQ # column is L, not S — column S
in JCI is `Coupa INV Date`. So JCI's "DELETED?" fallback is accidentally testing an unrelated
date column, not the REQ # it is nominally reporting on. This is a pre-existing bug independent
of the type-mismatch issue, noted here because it makes the `DELETED?` values on the JCI sheet
partly *misleading* on top of being partly a mismatch symptom (see section 6).

Key/type summary for every lookup:

| Formula col | Keys on | Target sheet!col | Target stored type | Coerces key? |
|---|---|---|---|---|
| B (both) | store-number (A) | 'BU List'!A/B | text (BU List presumed text, not directly audited) | no |
| T/M (both) | requisition-number (S/L) | 'Coupa Reqs'!A ("Req #") | **text** (confirmed, both tenants) | no |
| U/N (both) | requisition-number (S/L) | 'Coupa Reqs'!F ("PO Number") | **text** (confirmed, both tenants) | no |
| V/O (both) | tracker's own PO# (U/N) | 'Coupa POs'!A ("PO Number") | text (pinned via same `IsIdentifierHeader`) | no |
| X.../R.../etc | coupa-invoice-number key (Helper!B for Securitas, J directly for JCI) | 'Coupa Invs'!B ("Invoice #") | **text** (confirmed via `find`, e.g. JCI B581) | no — Helper!B is pre-coerced to text by `UpdateSearchValues`, but JCI's direct use of J is not pre-coerced |

**No formula anywhere coerces its key with `&""` or `TEXT()`.** XLOOKUP/VLOOKUP/Match in Excel do
not implicitly convert a numeric key to match a text value in the lookup array (or vice versa),
so every one of these lookups silently fails whenever the *tracker-side* key cell (S/L, U/N, or J)
is numeric while the *Coupa-side* target column is text — which, per section 2, is exactly the
condition affecting JCI's legacy rows.

### Helper sheet formulas (Securitas only; JCI has no Helper-based search key)

`WriteFormulas_Helper.vb` writes an `IFS(...)` status formula into `Helper!A`, unrelated to
identifier typing. `Update SEARCH column.vb` (`UpdateSearchValues`) populates `Helper!B` via
`Application.Match(O_or_K, 'Coupa Invs'!B:B, 0)`, explicitly writing the result as text
(`NumberFormat = "@"` set *before* `.Value = resultsArr`, `UpdateSearchValues.vb:85-90`) — this is
exactly the pattern this design should generalize.

### VBA dictionary/Find joins

- **`LookupReqs.vb`** (core): builds `reqLookup` keyed on `Trim$(CStr(coupaPartNums(i,1)))` (a
  string cast, so the dictionary key side is always text) against `wsTracker`'s
  `req-join-key` column (`submitted-invoice-number`/`coupa-invoice-number`, cast the same way via
  `Trim$(CStr(...))`). Both sides of *this* join are string-cast before comparison, so the
  Bill-Code/Invoice#-to-Coupa-Req join itself is not type-sensitive. **But the *value* written
  back — `trackerReqNums(i,1) = reqLookup(invoiceNum)`, where the dictionary value is the raw,
  uncast `coupaReqNums(i,1)`** — is written straight into the tracker's REQ # column with **no
  coercion and no `NumberFormat = "@"` set on the destination range** (`LookupReqs.vb:146`).
- **`CopyPaymentNums.vb`** (Securitas): `Application.Match(helperVal, coupaColB, 0)` against
  `'Coupa Invs'!B` (text, per above) using `Helper!B` (text, per `UpdateSearchValues`) — a clean
  text/text join. Correctly sets `NumberFormat = "@"` before writing the result
  (`CopyPaymentNums.vb:48-49`).
- **`UpdateSearchValues`** (Securitas): `Application.Match(O_or_K, coupaColB, 0)` — O and K are
  both text (section 2), `coupaColB` (`Coupa Invs!B`) is text. Clean join, and the write path sets
  `NumberFormat = "@"` first.

---

## 4. Where numbers get into REQ # — root cause

Two independent mechanisms, confirmed against the live JCI data:

**(a) Historical import.** Before commit `f78dc69` ("fix: write Coupa identifier columns as text
on import", 2026-08-03), `ProcessCoupaExport` → `CopyData` → `WriteCoupaColumnFromSourceData` had
no `IsIdentifierHeader` coercion, so `Coupa Reqs!A` ("Req #") held whatever type the Coupa export
produced — numeric, for a plain numeric ID. `LookupReqs` copies that column's raw values straight
through (`trackerReqNums(i,1) = reqLookup(invoiceNum)`, `LookupReqs.vb:139`) with no cast. Every
row backfilled before that date landed as a Number in the tracker. This explains the bulk of
JCI's 384 numeric REQ # rows: confirmed live that `Coupa Reqs!A` today holds these same
requisition numbers **as text** (e.g. `914751` → `Coupa Reqs!A527` = `"914751"`, `810492` →
`Coupa Reqs!A617` = `"810492"`) — the source has since been cleaned up (or re-imported), but the
already-written tracker cells were never revisited.

**(b) `LookupReqs` still has no `NumberFormat = "@"` guard, so the same bug can recur even now.**
Every other writer that puts a text-shaped value into a cell first sets
`target.NumberFormat = "@"` and only then assigns the value:
`ConvertStoreNumbersOn` (`ConvertStoreNumbers.vb:71-72`), `CopyPaymentNums`
(`CopyPaymentNums.vb:48-49`), `UpdateSearchValues` (`Update SEARCH column.vb:87-89`), and the
one-time `MigrateColumnTypes.ScanColumn` (`MigrateColumnTypes.vb:238-239`). This matters because
in VBA, assigning a numeric-looking `String` to `Range.Value` on a cell whose `NumberFormat` is
still `General` makes Excel parse it exactly as if a person had typed it in — silently
re-numericizing it — while a cell already set to `"@"` keeps it as text. `LookupReqs.vb:146`
(`wsTracker.Range(colTrackerReq & "2:" & colTrackerReq & lastRow).Value = trackerReqNums`) does
**not** set the format first. New rows appended by `AddNewBills`/`ProcessNewBills` never write
REQ # at all (it isn't a source column — it only gets filled later by `LookupReqs`), so a freshly
appended row's REQ # cell can easily still be `General`-formatted when `LookupReqs` later fills
it — meaning even a perfectly clean, already-text `Coupa Reqs!A` (which is the case today, per
section 2) does not guarantee LookupReqs's *next* backfill stays text. This is a live, ongoing
risk, not just historical residue.

**Does `ProcessCoupaExport` (CSV import) store Req # as number or text today?** Text — confirmed
by direct read of `Coupa Reqs!A` on both live workbooks (5870/1666 rows, 0 numbers). The
`IsIdentifierHeader`/`CoerceIdentifier` fix in `CoupaDataHelpers.vb` is doing its job on the
import side for both formats (CSV and XLSX route through the same
`WriteCoupaColumnFromSourceData`).

---

## 5. `ConvertStoreNumbers` — the pattern to follow

- **Structure** (`InvoiceTrackerCore/ConvertStoreNumbers.vb`): a thin `ConvertStoreNumbers(announce, manageProtection)` entry point resolves the sheet/column from `TenantConfig` (`TenantSheetName("tracker")`, `TenantColLetter("store-number")`) and delegates to `ConvertStoreNumbersOn(ws, colLetter, resolutions)`, which is sheet/column-agnostic and independently testable. `ConvertStoreNumbersOn` reads the whole column into an array, calls the pure function `NormalizeStoreNumber` per cell (4-digit zero-pad via `Format$(value, "0000")` when `IsNumeric`, otherwise passthrough or a human-supplied resolution from a `resolutions` dictionary), **sets `targetRange.NumberFormat = "@"` before** writing `targetRange.Value = updated` in one bulk assignment, and returns a count of non-empty conversions.
- **Wired into `Refresh`**: Securitas calls `ConvertStoreNumbers(False, True)` (self-manages protection, since Securitas's `Refresh` never opens its own unprotect window); JCI calls `ConvertStoreNumbers(False, False)` (protection already open around the whole `Refresh` sequence). Also called from each tenant's `AddNewBills` after new rows land, and from the one-time `MigrateColumnTypes`.
- **Configured through `TenantConfig`**: only two facts are needed per tenant — `TenantSheetName("tracker")` and `TenantColLetter("store-number")` — so the same core code runs unmodified against Securitas's column A and JCI's column A (same letter here, but the mechanism doesn't assume that).
- **manageProtection** parameter lets each tenant's differing `Refresh` protection style (Securitas: protect only inside each step; JCI: unprotect once around the whole sequence) drive whether the callee toggles protection itself — this is the seam a new "ConvertIdentifiers"-style routine needs to reuse to fit into both `Refresh` scripts without change to their differing protection idioms.

---

## 6. XLOOKUPs currently failing due to type mismatch

**JCI, `REQ STATUS` (M) and everything chained off it — confirmed at scale.** `find(query="DELETED?", sheet="Invoices")` on the live JCI workbook returns two blocks:

- Rows **282-405** (~120 of 124 scanned cells; a handful of gaps show blank instead, see below) literally show the text `"DELETED?"` — the formula's own fallback for "not found in Coupa Reqs AND the (mistargeted) `S` column happens to be numeric."
- The remaining rows in that same numeric-REQ# span (roughly rows 3-281) show **blank** `""` instead, because the fallback's `ISNUMBER(S)` test is false there (Coupa INV Date not yet populated) — same underlying XLOOKUP failure, just a different fallback branch.

Confirmed concretely with `find`, not just inferred from row position:
- `L268 = 914751` (Number) — `914751` **does** exist in `Coupa Reqs!A527` as text `"914751"`. `XLOOKUP(914751, 'Coupa Reqs'!A:A, ...)` cannot match a number against that text, so M268 is wrong (not a real deletion).
- `L377 = 1047860` (Number) — exists as `Coupa Reqs!A377 = "1047860"`. Same failure.
- `L` (many rows) `= 810492` (Number) — exists as `Coupa Reqs!A617 = "810492"`. Same failure, and `810492` recurs dozens of times in L, so this one mismatch alone silently blanks REQ STATUS for a large share of rows.

**JCI, `INV STATUS`/`Coupa INV Date`/`INV Total`/`Expected Pay`/`PAY Date`/`PAYMENT` (R, S, T, V, W, X) — same root cause, confirmed on one row.** All key on `J` directly (no Helper-style pre-coercion in JCI). `J2 = 8059053607` (Number); it exists as `Coupa Invs!B581 = "8059053607"` (text). `XLOOKUP(J2, 'Coupa Invs'!B:B, ...)` fails, and live `R2` is blank rather than a real invoice status. Since J has 280 Number-typed rows (section 2), this single mismatch pattern plausibly affects most or all of them, cascading through six formula columns per row rather than one.

**Count, honestly stated:** at minimum, the ~124 explicit `"DELETED?"` cells in M are a hard floor for confirmed-wrong REQ STATUS values (some fraction of the 124 could be genuine Coupa deletions rather than mismatches — I did not check all 124 individually, only 3 spot checks, all 3 of which were mismatches, not deletions). Beyond that block, the remainder of the 384 Number-typed L rows and 280 Number-typed J rows are the population *at risk*, and every spot check against Coupa Reqs/Coupa Invs came back "the value exists as text, the tracker just holds it as a number" — i.e., not a single spot check found a genuine not-found case. A full count would require joining all 384/280 numeric keys against the Coupa-side text lists cell by cell, which is straightforward for the eventual fix's dry-run/report step but wasn't done exhaustively here to stay within a read-only investigation.

**Securitas: no evidence of live mismatch failures.** `find(query="DELETED?", sheet="Invoices")` on the Securitas workbook returns zero formula-driven `DELETED?` hits (the only matches are unrelated free-text NOTES entries containing the word "deleted"). Consistent with section 2: every Securitas identifier column sampled is already 100% text.

---

## Recommended normalization (for the design, not implemented here)

- Add a core `ConvertIdentifiers`/`NormalizeIdentifierColumns`-style routine, shaped exactly like `ConvertStoreNumbersOn`: read column → run each cell through `CoerceIdentifier` (already in `CoerceValues.vb`, used by ingestion) → **set `NumberFormat = "@"` before** the bulk `.Value =` write → return a changed-count. Drive it from a list of concepts (`store-number` already covered; add `submitted-invoice-number`, `coupa-invoice-number`, `requisition-number`, `payment-number`, and JCI's `req-join-key`/`purchase-order-number` where distinct) resolved through `TenantColLetter`, mirroring `MigrateColumnTypes`'s `textConcepts` list but made idempotent and cheap enough to run on every `Refresh`, not just once.
- Wire it into both `Refresh.vb` scripts the same way `ConvertStoreNumbers` is wired in: after `LookupReqs` (so freshly backfilled REQ # values get swept too) and before `WriteFormulas_Tracker` (so the XLOOKUPs recalculate against clean keys in the same run).
- Fix `LookupReqs.vb:146` to set `NumberFormat = "@"` on the destination range before assigning `trackerReqNums` — this is the one write path in the whole audited call graph that skips the guard every sibling writer uses, and it is the mechanism that can keep reintroducing numeric REQ # cells even after a cleanup.
- **The matching Coupa-side sheets also need to stay text, and today they do** (`Coupa Reqs!A`/`F` and `Coupa Invs!B` are already 100% text on both live workbooks) — so the fix is one-sided: only the tracker's own S/L, U/N-source, and J columns need sweeping. No change is needed on the Coupa import side; `ProcessCoupaExport`'s `IsIdentifierHeader` coercion already covers it. Confirm this stays true after the next few imports, since it's inferred from a single snapshot rather than guaranteed by a test.
- Optional, separate from typing: JCI's `WriteFormulas_Tracker.vb` formula for M has a stale fallback that tests column `S` (Coupa INV Date) instead of `L` (REQ #) for the `"DELETED?"` label — worth flagging to whoever owns that formula, since it will keep mislabeling rows even after the type fix.
