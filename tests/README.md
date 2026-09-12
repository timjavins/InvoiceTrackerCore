# Tests

PowerShell, driving Excel over COM. There is no VBA test framework here and there cannot be a
convenient one: `Stack-VBFiles.ps1` sweeps `*.vb` recursively with no `tests/` exclusion, so any
VBA test module in this repo would be concatenated into every tenant's production stack. `.ps1`
files are not swept, so the tests live here instead.

Run them:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Each script exits 0 on all-pass, 1 on any failure, so they are CI-usable as-is.

## Requirements

- Excel installed.
- Excel's **Trust access to the VBA project object model** enabled (Trust Center > Macro
  Settings). The harness injects source into a throwaway workbook; without VBOM it cannot.

## What the harness does not do

It never attaches to a running Excel instance and never opens a workbook you have open. Each run
creates its own hidden `Excel.Application`, adds an empty workbook, injects the modules under test
as a **standard** module, calls functions with `Application.Run`, then closes without saving and
deletes its temp file.

The standard module is a deliberate simplification with a real cost. `Application.Run` cannot
resolve an unqualified name in a document module, and even when qualified as `ThisWorkbook.Proc` it
executes the procedure but **discards a `Function`'s return value** -- so assertions would be
impossible there. A standard module sidesteps both.

The cost is that a standard module allows things a document module forbids: `Public Const`, public
fixed-size arrays, fixed-length strings, `Declare`. **A module can pass every test here and still
fail to compile in the real stack.** That gap is closed separately, by compiling the assembled stack
in a scratch workbook and asserting against it through the Excel MCP, whose `run_macro` does return
values from document-module functions. Both layers are needed; neither substitutes for the other.
