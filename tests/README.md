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
as a **standard** module, calls functions with `Application.Run`, then closes without saving. The
workbook only ever exists in memory -- nothing calls `SaveAs` -- so there is no temp file on disk
to clean up.

The standard module is a deliberate simplification with a real cost. `Application.Run` cannot
resolve an unqualified name in a document module, and even when qualified as `ThisWorkbook.Proc` it
executes the procedure but **discards a `Function`'s return value** -- so assertions would be
impossible there. A standard module sidesteps both.

The cost is that a standard module allows things a document module forbids: `Public Const`, public
fixed-size arrays, fixed-length strings, `Declare`. **A module can pass every test here and still
fail to compile in the real stack.** That gap is closed separately, by compiling the assembled stack
in a scratch workbook and asserting against it through the Excel MCP, whose `run_macro` does return
values from document-module functions. Both layers are needed; neither substitutes for the other.

`CodeModule.AddFromString` does not validate VBA syntax -- it accepts the text unconditionally, and
a syntax error only surfaces later, as a confusing COM error at the first `Application.Run`. If a
red run looks like the harness itself is broken, check the injected source for a syntax error
before suspecting the harness.

## Never test a path that raises

A test must never call a VBA procedure that will `Err.Raise`. An unhandled VBA error under
`Application.Run` opens a modal End/Debug dialog; in the harness's hidden Excel instance nothing
can dismiss it, the COM call blocks forever, and the dialog can surface in the user's own Excel
session. Guard clauses are verified by code review, not by this suite.

## Two layers: standard module vs. document module

There are two scripts here for a reason, not by accident:

- **`Test-FiscalCalendar.ps1`** injects `FiscalCalendar.vb` into a **standard** module and calls
  its functions with `Application.Run`. This is where the calendar logic itself is tested --
  fiscal year boundaries, week counts, month starts -- against the fixture. A standard module
  is a deliberate simplification: it allows things a document module forbids (`Public Const`,
  public fixed-size arrays, fixed-length strings, `Declare`), and `Application.Run` only works
  at all because a standard module resolves unqualified names.

- **`Test-DocumentModule.ps1`** injects `Header.vb` then `FiscalCalendar.vb` into `ThisWorkbook`,
  a **document (class) module** -- the same kind of module, and the same injection order, that
  `Stack-VBFiles.ps1` pins in every tenant's real assembled stack. Production code lives in a
  document module, never a standard one, so this is the only script that exercises the
  restrictions that actually apply in production. It calls functions as COM methods on the
  workbook object (`$wb.FiscalYearWeeks(2026)`) rather than `Application.Run`, because
  `Application.Run` discards a document-module `Function`'s return value -- there would be
  nothing to assert against otherwise. Each call also forces lazy compilation of the module up
  to that point, so working through all 12 public functions is itself the compile check: a
  module can pass every assertion in `Test-FiscalCalendar.ps1` and still fail to compile here.

Neither script substitutes for the other. `Test-FiscalCalendar.ps1` proves the logic is correct;
`Test-DocumentModule.ps1` proves that same code still compiles and returns correct values once it
is a document-module member, which is the only place the class-module restrictions apply.
