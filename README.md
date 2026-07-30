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

## Design docs

Architecture decisions and the domain glossary live in `SecuritasAutomation`:

- `CONTEXT.md` — domain glossary
- `docs/adr/0001` — the two invoice-number identities
- `docs/adr/0002` — why core is an independent sibling folder
- `docs/adr/0003` — TenantConfig as the single narrow interface
- `docs/adr/0004` — the shadow rule
