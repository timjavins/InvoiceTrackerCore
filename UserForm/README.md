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

## Per-tenant naming

Form class names are resolved at compile time — `New SomeForm` cannot take a name from a
config string. So each variant owns its own copy of a form, named with its tenant prefix:

- `SecuritasWebDriverErrorForm`
- `JCIWebDriverErrorForm`

Core never names a form class. It calls a variant-owned shim (`ShowWebDriverError`) that
does the instantiation, so name resolution stays in the repo where the form actually exists.

## Forms

| Form | Purpose |
|---|---|
| `ChromeDriver Error/UserForm_WebDriverError.vb` | Shown when the Selenium WebDriver fails to start, usually a ChromeDriver version mismatch. Offers the SeleniumBasic directory and the driver download URL. |
| `SheetPicker/UserForm_SheetPicker.vb` | Lets the user pick a sheet when an imported workbook has no sheet matching the expected name. |
