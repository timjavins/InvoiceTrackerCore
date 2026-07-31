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

A UserForm is also not just text. Excel stores it as a `.frm` **plus** a binary `.frx`
holding the visual layout, so it cannot be reconstructed from the `.vb` here alone.

## Consequence: forms are imported by hand, once per workbook

The `.vb` files here are version control for the *code behind* each form. To put a form
into a workbook, export it from a workbook that has it (`.frm` + `.frx`) and import it via
the VBA editor. Keep the code here in step by hand when it changes.

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
| `ChromeDriver Error/UserForm_WebDriverError.vb` | Shown when the Selenium WebDriver fails to start, usually a ChromeDriver version mismatch. Offers the SeleniumBasic directory and the driver download URL. |
| `SheetPicker/UserForm_SheetPicker.vb` | Lets the user pick a sheet when an imported workbook has no sheet matching the expected name. |

## Manual steps outstanding

All are one-time workbook edits in the VBA editor (Properties window → Name). The code
already expects the prefixed names.

1. **Rename Securitas's error form** `WebDriverErrorForm` → `SecuritasWebDriverErrorForm`.
2. **Rename Securitas's picker** `SheetPickerForm` → `SecuritasSheetPickerForm`.
3. **Export JCI's `JCIWebDriverErrorForm`** into `JCI-invoice-tracker/UserForm/` so it is
   under version control. It exists only inside the `.xlsm` today. Its code should match
   `ChromeDriver Error/UserForm_WebDriverError.vb` here.

Optional — give JCI a sheet picker: import `SheetPicker/UserForm_SheetPicker.vb` into JCI's
workbook, name it `JCISheetPickerForm`, then replace the stub body in
`JCI-invoice-tracker/PickSheetName.vb` with the instantiation Securitas uses. Until then core
prompts with a list of the workbook's sheets, which works.
