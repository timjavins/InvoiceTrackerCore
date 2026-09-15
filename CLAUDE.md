# InvoiceTrackerCore

Shared VBA modules for the invoice tracker family. Consumed at assembly time by variant repos
(`SecuritasAutomation`, `JCI-invoice-tracker`, and `NordGuardsTracker` once its migration lands).
Read `README.md` first — it carries the conventions that actually bind code here.

## Doc Sync adapter

Required role mapping for the global `doc-sync-core` skill. Two roles live in a sibling repo, which
is deliberate: this repo holds no per-supplier facts and no decision record of its own.

| Canonical role | File |
|---|---|
| Session Log | `notes/sessions/` (per the global `session-notes` skill; gitignored, local only) |
| Coordination Board | `../SecuritasAutomation/docs/core-extraction/tickets.md` |
| Capability Doc | `README.md` |
| Architecture Contract | `../SecuritasAutomation/docs/adr/` (ADRs 0001-0005) |
| Doc Index | not used |

**`README.md` is the Capability Doc, and that is a live obligation rather than a label.** When a
module is added, removed, or changes what it offers consumers, `README.md` changes in the same
commit. That rule exists because it was broken once: the fiscal-calendar effort shipped
`FiscalCalendar.vb` plus an entire `tests/` directory without touching `README.md`, and a consumer
reading it would not have learned either existed. The final review caught it.

**Why the Architecture Contract is cross-repo.** `README.md` already states it: architecture
decisions and the domain glossary live in `SecuritasAutomation` (`CONTEXT.md`, `docs/adr/`). ADR-0002
records why core is an independent sibling folder rather than a submodule, and ADR-0003 why
`TenantConfig` is the single narrow interface. Do not start a parallel ADR series here — add to that
one, and cite it by number rather than by a vague "see the ADR", which was another defect the final
review caught in production source.

## Testing

There is no VBA test framework, and one cannot live in this repo as VBA: `Stack-VBFiles.ps1` globs
`*.vb` **recursively** with no `tests/` exclusion, so a `.vb` file under `tests/` would be
concatenated into every tenant's production stack. PowerShell files are not swept, so the suites are
PowerShell driving Excel over COM. See `tests/README.md`, which explains the two layers and the
traps — including why a test must never call a procedure that can `Err.Raise`.

## Session documentation

Use the global `session-notes` skill. Notes live in `notes/sessions/` and are gitignored.
