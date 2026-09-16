# Tests

PowerShell, driving Excel over COM. There is no VBA test framework here and there cannot be a
convenient one: `Stack-VBFiles.ps1` sweeps `*.vb` recursively with no `tests/` exclusion, so any
VBA test module in this repo would be concatenated into every tenant's production stack. `.ps1`
files are not swept, so the tests live here instead.

Run them:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-DocumentModule.ps1
```

Each script exits 0 on all-pass, 1 on any failure, so they are CI-usable as-is.

## Requirements

- Excel installed.
- Excel's **Trust access to the VBA project object model** enabled (Trust Center > Macro
  Settings). The harness injects source into a throwaway workbook; without VBOM it cannot.

## What the harness does not do

It never attaches to a running Excel instance and never opens a workbook you have open. Each run
creates its own hidden `Excel.Application`, adds an empty workbook, injects the modules under test
as a **standard** module (or, for `Test-DocumentModule.ps1`, into `ThisWorkbook` -- see below),
calls functions with `Application.Run` (or as COM methods on the workbook object, for the
document-module case), then closes without saving. The workbook only ever exists in memory --
nothing calls `SaveAs` -- so there is no temp file on disk to clean up.

The standard module is a deliberate simplification with a real cost. `Application.Run` cannot
resolve an unqualified name in a document module, and even when qualified as `ThisWorkbook.Proc` it
executes the procedure but **discards a `Function`'s return value** -- so assertions would be
impossible there. A standard module sidesteps both.

The cost is that a standard module allows things a document module forbids: `Public Const`, public
fixed-size arrays, fixed-length strings, `Declare`. **A module can pass every test here and still
fail to compile in the real stack.** That gap is closed separately, by `Test-DocumentModule.ps1`
(below), which injects the module under test into `ThisWorkbook` -- a document module, same as
production -- over the same PowerShell/COM harness. Injecting an assembled tenant stack instead was
considered and rejected: Securitas's stack declares variables typed as its UserForms, which are
bound at compile time and cannot compile in a bare scratch workbook. Both layers are needed; neither
substitutes for the other.

`CodeModule.AddFromString` does not validate VBA syntax -- it accepts the text unconditionally, and
a syntax error only surfaces later, as a confusing COM error at the first `Application.Run`. If a
red run looks like the harness itself is broken, check the injected source for a syntax error
before suspecting the harness.

**VBA compiles per module, on demand -- not the whole project up front.** A module that gets
entered must have every name it references directly resolvable, including a call inside a branch
that never executes at runtime, because runtime reachability is irrelevant to compilation. A module
that is never entered, on the other hand, does not need its own callees resolvable, because it is
never compiled at all. Concretely: injecting `LookupReqs.vb` requires also injecting
`UnprotectSheet.vb` and `ProtectSheet.vb`, even when `manageProtection:=False` means those calls
never run, while `TenantSheetPassword` -- referenced only inside those two modules' own bodies --
needs no stub, because those bodies are never compiled unless something enters them. Getting this
backwards looks like a harness bug (an "undefined" error for a name that is clearly stubbed
elsewhere, or a missing stub that never seems to matter) when it is really about which module got
entered.

## Never test a path that raises

A test must never call a VBA procedure that will `Err.Raise`. An unhandled VBA error under
`Application.Run` opens a modal End/Debug dialog; in the harness's hidden Excel instance nothing
can dismiss it, the COM call blocks forever, and the dialog can surface in the user's own Excel
session. Guard clauses are verified by code review, not by this suite.

`Application.Visible = $false` does **not** suppress that dialog -- it renders on screen regardless
of instance visibility. The danger is not that nobody *can* dismiss a modal in a hidden instance; it
is that the **operator** must, in their own Excel session, because the blocked COM call never
returns on its own. This is not theoretical: it has happened, mid-run, and the operator had to press
OK by hand and abort. See "Timing out a stuck call" below for the harness's mitigation.

## Timing out a stuck call

`Invoke-VbaFunction` accepts `-TimeoutSeconds`, which defaults to `0` -- today's behaviour, wait
indefinitely -- so no existing caller changes. A positive value arms a watchdog for that one call:
if it has not returned in time, the watchdog kills **only the Excel process `New-VbaHost` captured
for that host** and `Invoke-VbaFunction` throws a specific error naming the procedure that blocked,
instead of leaving a hung COM call and a modal dialog for a human to find.

`New-VbaHost` captures that PID by a before/after set difference of running `EXCEL` process IDs,
taken immediately before and after it creates its own `Excel.Application`. If that difference is
anything other than exactly one new PID, it stores nothing, and every kill path -- the watchdog here
and `Remove-VbaHost`'s poll below -- becomes inert rather than guessing. That is deliberate: a wrong
guess could kill the operator's own Excel, including a live cloud-hosted workbook with unsaved
changes. Never hardcode or guess a PID to make a timeout "work".

`Remove-VbaHost` also polls for about two seconds after calling `Quit()` -- which returns before the
process has actually exited -- and kills the captured PID if it is still alive once that grace
period elapses.

## Two layers: standard module vs. document module

There are two scripts here for a reason, not by accident:

- **`Test-FiscalCalendar.ps1`** injects `FiscalCalendar.vb` into a **standard** module and calls
  its functions with `Application.Run`. This is where the calendar logic itself is tested --
  fiscal year boundaries, week counts, month starts -- against the fixture. A standard module
  is a deliberate simplification: it allows things a document module forbids (`Public Const`,
  public fixed-size arrays, fixed-length strings, `Declare`), and `Application.Run` only works
  at all because a standard module resolves unqualified names.

- **`Test-DocumentModule.ps1`** injects `FiscalCalendar.vb` into `ThisWorkbook`, a **document
  (class) module** -- the same kind of module `Stack-VBFiles.ps1` pins it into in every tenant's
  real assembled stack. Production code lives in a document module, never a standard one, so
  this is the only script that exercises the restrictions that actually apply in production. It
  does not also inject `Header.vb`: that module is comment-only in the Securitas tenant and has
  no declarations, so it would prove nothing here, and it would make this script depend on a
  sibling repo checkout that will not exist everywhere it runs. It calls functions as COM methods on the
  workbook object (`$wb.FiscalYearWeeks(2026)`) rather than `Application.Run`, because
  `Application.Run` discards a document-module `Function`'s return value -- there would be
  nothing to assert against otherwise. Each call also forces lazy compilation of the module up
  to that point, so working through all 12 public functions is itself the compile check: a
  module can pass every assertion in `Test-FiscalCalendar.ps1` and still fail to compile here.

Neither script substitutes for the other. `Test-FiscalCalendar.ps1` proves the logic is correct;
`Test-DocumentModule.ps1` proves that same code still compiles and returns correct values once it
is a document-module member, which is the only place the class-module restrictions apply.
