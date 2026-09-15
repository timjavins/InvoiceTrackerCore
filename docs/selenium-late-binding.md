# Selenium: switch to late binding

**Status:** recommended, not yet done. Raised 2026-09-15 while designing VBA write access for
Excel-MCP, from the question "can an agent turn on the Selenium reference library for a user?"

The answer turned out to be that nobody should have to. The reference requirement can be removed
outright — and removing it also fixes a guard that cannot currently fire.

## The bug this fixes

`AttachCoupaDocs` and `GetBlockerEmails` both check for Selenium before using it:

```vba
On Error Resume Next
seleniumInstalled = Not (CreateObject("Selenium.WebDriver") Is Nothing)
On Error GoTo 0

If Not seleniumInstalled Then
    MsgBox "Selenium is not installed. Please install Selenium Basic and the appropriate Chrome driver.", vbExclamation
    Exit Sub
End If
```

That check is correct and **unreachable on the machines it was written for.**

Twelve lines above it, `Dim driver As WebDriver` names a type that only exists when the `Selenium`
reference is present. VBA compiles before it runs. So on a machine without SeleniumBasic installed
and the reference ticked, the project fails to compile with "User-defined type not defined" and
execution never reaches the friendly message. The user gets a compile error pointing at a `Dim`
line instead of an instruction telling them what to install.

A guard that only runs where it is not needed is not a guard.

## The fix

Replace the early-bound types with `Object` and create the driver through the ProgID the check
already uses. Then the same `CreateObject` call both proves Selenium is present *and* produces the
object, instead of instantiating one only to throw it away:

```vba
Dim driver As Object

On Error Resume Next
Set driver = CreateObject("Selenium.WebDriver")
On Error GoTo 0

If driver Is Nothing Then
    MsgBox "Selenium is not installed. Please install Selenium Basic and the appropriate Chrome driver.", vbExclamation
    Exit Sub
End If
```

Note the guard also gets less fragile. The current version relies on `seleniumInstalled` still
holding the `False` that `Dim` gave it, because `On Error Resume Next` skips the failed assignment
rather than writing to it — it works by accident of default initialisation. Testing the object
reference directly does not depend on that.

## What has to change

Two files, three type names, roughly nineteen declarations — including function *parameters*, which
are easy to miss:

| File | Declarations |
|---|---|
| `AttachCoupaDocs.vb` | `driver As WebDriver` (local, and as a parameter in four function signatures), `WebElement` locals, `WebElements` locals, `Set driver = New WebDriver` |
| `GetBlockerEmails.vb` | `driver As WebDriver` (local, and one `Private Function` parameter), `Set driver = New WebDriver` |

Every `As WebDriver`, `As WebElement` and `As WebElements` becomes `As Object`. Every
`New WebDriver` becomes `CreateObject("Selenium.WebDriver")`.

## What it costs

IntelliSense and compile-time type checking on Selenium calls, in the VBE, for whoever maintains
these two files. That is the whole cost.

The usual second cost does not apply here: late binding also forces named enum constants to be
replaced with their literal values, and **neither file uses any** — no `By.`, no `Keys.`, no
`Selenium.` constants, only the `WebDriver` / `WebElement` / `WebElements` types. Checked, not
assumed.

## What it does not fix

SeleniumBasic and a version-matched Chrome driver still have to be installed on the machine. Late
binding removes the *reference* step and moves the failure from compile time to runtime where the
existing message can catch it. It does not remove the install.

Consumers therefore go from two setup steps to one, and the remaining step announces itself in
plain language instead of as a compile error.

## Why not have an agent tick the reference instead

It was considered, and it is worse:

* Ticking `Tools > References` by hand needs no special Excel setting. Having an agent do it
  programmatically needs "Trust access to the VBA project object model" — a per-machine macro
  security setting, off by default, that macro viruses use to spread. Asking every consumer to
  enable that so a tool can save them one checkbox is a bad trade.
* It would leave the compile-time failure in place for anyone who had not run the agent yet.
