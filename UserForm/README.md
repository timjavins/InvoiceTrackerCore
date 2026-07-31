# UserForms

Canonical source for the shared UserForms. **Excluded from the stack on purpose** — the
assembler skips any path containing `UserForm`.

## Why these can't be stacked

`Stack-VBFiles.ps1` concatenates module source into one blob that gets pasted into a single
standard module in the workbook. Form code does not live there: each UserForm is its own
object in the VBA project, with event handlers (`btnClose_Click`, `UserForm_Initialize`)
bound to controls that only exist on that form. Pasting it into the main module would
produce handlers with no controls to bind to, plus a stray `Option Explicit` in the middle
of the file.

A UserForm is also not just text. Excel stores it as a `.frm` — which holds the control
definitions **and** the code-behind — plus a binary `.frx` holding the visual layout.

## Where the forms live

Each variant owns its own `.frm`/`.frx` pair under its own `UserForm/` folder, because form
class names must carry the tenant prefix (see below). This folder holds no form sources: a
`.frm` is already self-contained, so keeping a separate copy of the code-behind here would
just be a second version to drift out of step.

    SecuritasAutomation/UserForm/ChromeDriver Error/SecuritasWebDriverErrorForm.frm + .frx
    SecuritasAutomation/UserForm/SheetPicker/SecuritasSheetPickerForm.frm + .frx
    JCI-invoice-tracker/UserForm/ChromeDriver Error/JCIWebDriverErrorForm.frm + .frx
    JCI-invoice-tracker/UserForm/SheetPicker/JCISheetPickerForm.frm + .frx

To move a form between workbooks: export the `.frm` (the `.frx` follows automatically) from
the VBA editor, import it into the target workbook, then rename it for that tenant.

## Per-tenant naming is required, not cosmetic

**Every form must carry its tenant's prefix.** Both tracker workbooks are routinely open in
the same Excel instance, and their VBA projects share the default project name
(`VBAProject`), so two forms with the same class name across the two projects collide.

- `SecuritasWebDriverErrorForm` / `JCIWebDriverErrorForm`
- `SecuritasSheetPickerForm` / `JCISheetPickerForm`

Core never names a form class. It calls a variant-owned shim — `ShowWebDriverError`,
`PickSheetName` — which does the instantiation, so name resolution stays in the repo whose
workbook actually holds the form.

### A shim cannot fall back when its form is missing

`Dim frm As SomeForm` is resolved at **compile time**, so a workbook without the form fails
to compile — an error handler cannot rescue it. A variant that does not have a form must not
mention the class at all; its shim returns empty and lets core take the fallback path. JCI's
`PickSheetName` is exactly that stub.

## Forms

| Form | Purpose |
|---|---|
| `<Tenant>WebDriverErrorForm` | Shown when the Selenium WebDriver fails to start, usually a ChromeDriver version mismatch. Offers the SeleniumBasic directory and the driver download URL. |
| `<Tenant>SheetPickerForm` | Lets the user pick a sheet when an imported workbook has no sheet matching the expected name. |

## Manual steps — done (2026-07-31)

Both workbooks' forms were renamed with tenant prefixes and exported into their repos, so
all four `.frm`/`.frx` pairs are now version-controlled. Nothing outstanding.

JCI's picker form exists in the repo but has not been wired up yet — see below.

### Optional: give JCI a sheet picker

Until this is done, core prompts with a list of the workbook's sheets, which works fine.

Two parts, and only the second is text you can copy:

**1. The form object — import it, do not retype it.** A UserForm is a `.frm` (control
definitions plus code-behind) *and* a binary `.frx` (the visual layout), so it cannot be
rebuilt by pasting code. Import
`JCI-invoice-tracker/UserForm/SheetPicker/JCISheetPickerForm.frm` in the VBA editor
(File → Import File); the `.frx` is picked up alongside it. Confirm the imported form is
named `JCISheetPickerForm`.

**2. The shim — replace the stub body** in `JCI-invoice-tracker/PickSheetName.vb` with this,
which is Securitas's version with the class name changed:

```vba
Public Function PickSheetName(ByVal wb As Workbook, ByVal expectedName As String) As String
    On Error GoTo ShowFailed

    Dim picker As JCISheetPickerForm
    Set picker = New JCISheetPickerForm

    picker.InitializeSheets wb, expectedName
    picker.Show vbModal

    PickSheetName = VBA.Trim$(picker.SelectedSheetName)

    Unload picker
    Set picker = Nothing
    Exit Function

ShowFailed:
    ' Could not display the picker: return empty so core prompts instead.
    PickSheetName = vbNullString
End Function
```

Do not add these lines before the form exists in the workbook — the type is resolved at
compile time, so naming a form that is not there breaks the build outright.

## About the GUID in the error form

The error form's code contains:

```vba
Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
```

That is Microsoft's registered CLSID for `MSForms.DataObject`, the clipboard helper — a fixed
Windows constant, the same on every machine. It is not per-workbook, not generated, and
nothing to keep unique: copy it verbatim. `MSForms.DataObject` cannot be created by name from
a standard module without a reference to the MSForms library, which is why the class ID is
used instead.

Form class names and CLSIDs are unrelated concerns. Class names collide because they are
*our* identifiers in two projects that share a default project name; a CLSID is Microsoft's
identifier for their own class and cannot collide.
