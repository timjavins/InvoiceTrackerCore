# Batch Tracking, Monitoring-Aware Duplicate Check, All-Years Mirror — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop re-importing bill files already processed, stop flagging monitoring bills as duplicates while still skipping true duplicates, and keep the All-Years archive mirrored from the working book.

**Architecture:** Four pure core modules do the logic and are tested in the existing PowerShell/COM harness:
- `HashFile`
- `ProcessedBatchLog`
- `BillCodeScreen`
- `MirrorInvoiceBlock`

A fifth core module, `MoveProcessedFile`, is small file I/O that is also harness-tested. Variant code only wires them in: `AddNewBills` in both tenants, and in Securitas also `UpdateCoupaData` and `Refresh`, plus a thin Securitas `SyncAllYears` wrapper that opens the SharePoint archive. Variant wiring is verified by driving the flows by hand, per the repos' standing rule.

**Tech Stack:** VBA (Excel, stacked into `ThisWorkbook`), PowerShell 5.1 test harness (`tests/VbaHarness.psm1`), `certutil`.

**Spec:** `docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md` (this repo). Read it before starting any task.

## Global Constraints

- **No `Option Explicit` outside `Header.vb`.** Every module is concatenated into one file.
- **No `Public Const`, public fixed-size arrays, fixed-length strings, or `Declare`.** The stack lives in `ThisWorkbook`, a class module. `Private Const` is fine.
- **Every procedure name is global to the stack, `Private` ones included.** The whole stack is one module, so two `Private Function Foo` in different files collide. Prefix private helpers per module: `Hf` (HashFile), `Pbl` (ProcessedBatchLog), `Mpf` (MoveProcessedFile), `Bcs` (BillCodeScreen), `Mib` (MirrorInvoiceBlock), `Ay` (SyncAllYears).
- **No VBA test code in any repo.** `Stack-VBFiles.ps1` sweeps `*.vb` recursively with no `tests/` exclusion. Test stubs and drivers are PowerShell here-strings written to `$env:TEMP`, the pattern `tests/Test-ReqJoinKey.ps1` uses.
- **Core functions under test never `Err.Raise` and never show a `MsgBox`.** An unhandled error under `Application.Run` opens a modal dialog that hangs the harness. Failures are return values: `""`, `False`, or `-1`.
- **`Invoke-VbaFunction` passes at most 2 arguments.** Tests call small driver functions that supply the rest.
- **Don't depend on the active sheet.** Take a `Worksheet`/`Workbook` parameter.
- **Never write to a workbook the user has open.** That includes `2026 SECURITAS bills.xlsm`, `Securitas All-Years Invoices - Consolidated.xlsm` and `JCI Repair & Installation Invoices.xlsm`. Tests use the harness's own hidden instance. Manual verification uses **copies**.
- **Commits:** Conventional Commits, subject ≤ 50 chars, body only when the why isn't obvious. **No AI attribution or Co-Authored-By lines.** Commit each repo separately. Push only when the user says to.
- **Core README changes in the same commit as the module it describes.**
- **Run tests from the repo root** with `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-X.ps1`. Exit code 0 means pass. Excel's "Trust access to the VBA project object model" must be on (it is on this machine).

## File Structure

`InvoiceTrackerCore` (new):
| File | Responsibility |
|---|---|
| `HashFile.vb` | SHA-256 of a file via `certutil`; `""` on failure |
| `ProcessedBatchLog.vb` | Hidden per-workbook log of processed file hashes |
| `MoveProcessedFile.vb` | Move a processed file into a sibling subfolder |
| `BillCodeScreen.vb` | Classify incoming bill codes: ok / already-processed / repeat-in-file |
| `MirrorInvoiceBlock.vb` | Replace the owned block of one sheet with a values-only copy of another |
| `tests/Test-HashFile.ps1`, `tests/Test-ProcessedBatchLog.ps1`, `tests/Test-MoveProcessedFile.ps1`, `tests/Test-BillCodeScreen.ps1`, `tests/Test-MirrorInvoiceBlock.ps1` | One harness script per module |

`SecuritasAutomation`:
| File | Change |
|---|---|
| `TenantConfig.vb` | + 2 batch accessors, + 5 All-Years accessors |
| `file ingesting/AddNewBills.vb` | hash gate, pre-write screen, record, move, summary, sync |
| `file ingesting/ProcessNewBills.vb` | optional `rowStatus` skips rows |
| `file ingesting/CheckNewBillsForDuplicates.vb` | **delete** |
| `SyncAllYears.vb` (new) | open archive, guard, mirror, save, stamp, warn |
| `file ingesting/UpdateCoupaData.vb`, `Refresh.vb` | staleness warning + sync call |

`JCI-invoice-tracker`:
| File | Change |
|---|---|
| `TenantConfig.vb` | + 2 batch accessors |
| `AddNewBills.vb` | hash gate, record, move |

---

### Task 1: `HashFile` (core)

**Files:**
- Create: `InvoiceTrackerCore/HashFile.vb`
- Create: `InvoiceTrackerCore/tests/Test-HashFile.ps1`

**Interfaces:**
- Consumes: nothing.
- Produces: `Public Function HashFile(ByVal filePath As String) As String`, which returns 64 lowercase hex characters, or `""` if the file is missing or unreadable, or if `certutil` fails.

- [ ] **Step 1: Write the failing test**

Create `tests/Test-HashFile.ps1`:

```powershell
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-HashFile.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$dir = Join-Path $env:TEMP ('hashfile-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$abc = Join-Path $dir 'abc.txt'
[IO.File]::WriteAllText($abc, 'abc')                        # no BOM, no newline
$spaced = Join-Path $dir 'bill file (2) & co.txt'
[IO.File]::WriteAllText($spaced, 'abc')
$missing = Join-Path $dir 'nope.txt'

# Known SHA-256 of the three bytes "abc" (FIPS 180-2 test vector).
$abcHash = 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'

$src = Join-Path $repo 'HashFile.vb'
$vba = New-VbaHost -SourceFiles @($src)
try {
    Assert-Equal -Expected $abcHash -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($abc) -TimeoutSeconds 30) `
                 -Because 'hash of "abc" matches the FIPS test vector'
    Assert-Equal -Expected $abcHash -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($spaced) -TimeoutSeconds 30) `
                 -Because 'a path with spaces, parentheses and & still hashes'
    Assert-Equal -Expected '' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($missing) -TimeoutSeconds 30) `
                 -Because 'a missing file returns "" rather than raising'
    Assert-Equal -Expected '' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @('https://example.com/x.xlsx') -TimeoutSeconds 30) `
                 -Because 'a URL returns "" rather than raising'
    $leftovers = @(Get-ChildItem -Path $env:TEMP -Filter 'hashfile_*.txt' -ErrorAction SilentlyContinue |
                   Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) })
    Assert-Equal -Expected 0 -Actual $leftovers.Count -Because 'HashFile deletes its temp output file'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Same module as a document-module member: proves it compiles where production puts it.
$vba = New-VbaHost -SourceFiles @($src) -DocumentModule
try {
    Assert-Equal -Expected $abcHash -Actual ($vba.Workbook.HashFile($abc)) `
                 -Because 'HashFile compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
Write-AssertSummary
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-HashFile.ps1`
Expected: throws `Source file not found: ...HashFile.vb`.

- [ ] **Step 3: Implement**

Create `HashFile.vb`:

```vb
' SHA-256 of a file's bytes, as 64 lowercase hex characters, or "" if it cannot be computed.
'
' Identifies a bill file by content rather than by name: a vendor re-sending the same file
' under a new name, or someone copying it back into the inbox, still hashes the same. See the
' batch-tracking spec, section 1.
'
' Shells to certutil, which ships with Windows, rather than calling CryptoAPI: that needs
' Declare, which a class module (ThisWorkbook, where the stack lives) forbids as a public
' member and which would tie this module to PtrSafe signatures. Output goes through a temp
' file because WScript.Shell.Exec always flashes a console window; Run with style 0 does not.
'
' "" is the failure signal, never Err.Raise: callers treat "" as "stop before opening the
' file", and the test harness cannot survive a raised error.
'
' Returns the hash only; it stores nothing. AddNewBills passes it to ProcessedBatchLog, which
' keeps it in column A of the hidden "Processed Batches" sheet in the tracker workbook.
Public Function HashFile(ByVal filePath As String) As String
    Dim outPath As String
    Dim wsh As Object
    Dim fileNum As Integer
    Dim textLine As String
    Dim candidate As String

    HashFile = vbNullString
    On Error GoTo Failed

    If LCase$(Left$(filePath, 4)) = "http" Then Exit Function
    If Len(Dir$(filePath)) = 0 Then Exit Function

    outPath = Environ$("TEMP") & "\hashfile_" & Format$(Now, "yyyymmddhhnnss") & "_" & _
              CStr(Int(Rnd * 1000000)) & ".txt"
    Set wsh = CreateObject("WScript.Shell")
    If wsh.Run("cmd /c certutil -hashfile """ & filePath & """ SHA256 > """ & outPath & """", 0, True) <> 0 Then
        GoTo Cleanup
    End If

    ' certutil prints a label line, the digest, then a status line. Older Windows puts a
    ' space between each byte of the digest, so strip spaces before testing the shape.
    fileNum = FreeFile
    Open outPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textLine
        candidate = LCase$(Replace(Trim$(textLine), " ", ""))
        If HfIsHexDigest(candidate) Then
            HashFile = candidate
            Exit Do
        End If
    Loop
    Close #fileNum
    fileNum = 0

Cleanup:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If Len(outPath) > 0 Then If Len(Dir$(outPath)) > 0 Then Kill outPath
    Exit Function

Failed:
    HashFile = vbNullString
    Resume Cleanup
End Function

Private Function HfIsHexDigest(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) <> 64 Then Exit Function
    For i = 1 To 64
        If InStr(1, "0123456789abcdef", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    HfIsHexDigest = True
End Function
```

- [ ] **Step 4: Run it and confirm it passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-HashFile.ps1`
Expected: `ALL PASS`, exit 0.

- [ ] **Step 5: Add a README section and commit**

Append to `README.md`, before `## Design docs`:

```markdown
## Batch tracking

`HashFile.vb`, `ProcessedBatchLog.vb` and `MoveProcessedFile.vb` let a tenant's `AddNewBills`
refuse a bill file it has already processed. The file is identified by SHA-256 of its bytes
(`certutil`), so a renamed or re-sent copy is still recognised. See
`docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md`, section 1.

`HashFile` returns `""` on any failure and never raises. Callers treat `""` as "stop before
opening the file".
```

```bash
git add HashFile.vb tests/Test-HashFile.ps1 README.md
git commit -m "feat: add HashFile for bill-file identity"
```

---

### Task 2: `ProcessedBatchLog` (core)

**Files:**
- Create: `InvoiceTrackerCore/ProcessedBatchLog.vb`
- Create: `InvoiceTrackerCore/tests/Test-ProcessedBatchLog.ps1`
- Modify: `InvoiceTrackerCore/README.md` (the "Batch tracking" section from Task 1)

**Interfaces:**
- Consumes: tenant accessor `TenantProcessedBatchSheet() As String` (defined per variant in Task 6 and Task 7).
- Produces:
  - `Public Function IsAlreadyProcessed(ByVal wb As Workbook, ByVal hash As String) As Boolean`
  - `Public Function ProcessedBatchSummary(ByVal wb As Workbook, ByVal hash As String) As String`, which returns `""` if the hash is not logged.
  - `Public Sub RecordProcessedBatch(ByVal wb As Workbook, ByVal hash As String, ByVal sourceFile As String, ByVal tenant As String, ByVal rowsAdded As Long, ByVal rowsSkipped As Long)`
  - The log sheet has headers `HASH | SOURCE FILE | TENANT | PROCESSED AT | ROWS ADDED | ROWS SKIPPED` and is created hidden (`xlSheetHidden`) on first write.

- [ ] **Step 1: Write the failing test**

Create `tests/Test-ProcessedBatchLog.ps1`:

```powershell
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ProcessedBatchLog.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$stub = Join-Path $env:TEMP 'StubPblTenant.vb'
@'
Public Function TenantProcessedBatchSheet() As String
    TenantProcessedBatchSheet = "Processed Batches"
End Function

Public Function T_IsProcessed(ByVal h As String) As Boolean
    T_IsProcessed = IsAlreadyProcessed(ThisWorkbook, h)
End Function

Public Function T_Record(ByVal h As String, ByVal f As String) As String
    RecordProcessedBatch ThisWorkbook, h, f, "TestTenant", 3, 1
    T_Record = "ok"
End Function

Public Function T_Summary(ByVal h As String) As String
    T_Summary = ProcessedBatchSummary(ThisWorkbook, h)
End Function

Public Function T_SheetVisibility() As Long
    T_SheetVisibility = ThisWorkbook.Worksheets("Processed Batches").Visible
End Function

Public Function T_LogRows() As Long
    With ThisWorkbook.Worksheets("Processed Batches")
        T_LogRows = .Cells(.Rows.Count, 1).End(xlUp).Row - 1
    End With
End Function

Public Function T_Header(ByVal c As Long) As String
    T_Header = CStr(ThisWorkbook.Worksheets("Processed Batches").Cells(1, c).Value)
End Function
'@ | Set-Content -LiteralPath $stub -Encoding UTF8

$src = Join-Path $repo 'ProcessedBatchLog.vb'
$h1 = 'a' * 64
$h2 = 'b' * 64

$vba = New-VbaHost -SourceFiles @($stub, $src)
try {
    $run = { param($n, $a) Invoke-VbaFunction -VbaHost $vba -Name $n -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @($h1)) -Because 'nothing is processed before the log sheet exists'
    Assert-Equal -Expected ''     -Actual (& $run 'T_Summary' @($h1))     -Because 'no summary for an unlogged hash'
    Assert-Equal -Expected 'ok'   -Actual (& $run 'T_Record' @($h1, 'bills-sept.xlsx')) -Because 'first record creates the sheet'
    Assert-Equal -Expected $true  -Actual (& $run 'T_IsProcessed' @($h1)) -Because 'a recorded hash is processed'
    Assert-Equal -Expected $true  -Actual (& $run 'T_IsProcessed' @($h1.ToUpper())) -Because 'hash match ignores case'
    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @($h2)) -Because 'a different hash is not processed'
    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @(''))  -Because 'an empty hash is never processed'
    Assert-Equal -Expected 0      -Actual (& $run 'T_SheetVisibility' @()) -Because 'the log sheet is hidden (xlSheetHidden = 0)'
    Assert-Equal -Expected 'HASH' -Actual (& $run 'T_Header' @(1)) -Because 'header A1'
    Assert-Equal -Expected 'ROWS SKIPPED' -Actual (& $run 'T_Header' @(6)) -Because 'header F1'

    $summary = [string](& $run 'T_Summary' @($h1))
    Assert-Equal -Expected $true -Actual ($summary -like '*bills-sept.xlsx*') -Because "summary names the file ($summary)"
    Assert-Equal -Expected $true -Actual ($summary -like '*3 added, 1 skipped*') -Because "summary gives the counts ($summary)"

    Assert-Equal -Expected 'ok' -Actual (& $run 'T_Record' @($h2, 'bills-oct.xlsx')) -Because 'second record reuses the sheet'
    Assert-Equal -Expected 2    -Actual (& $run 'T_LogRows' @()) -Because 'two batches, two rows, one sheet'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($stub, $src) -DocumentModule
try {
    Assert-Equal -Expected $false -Actual ($vba.Workbook.T_IsProcessed($h1)) `
                 -Because 'ProcessedBatchLog compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $stub -Force -ErrorAction SilentlyContinue
Write-AssertSummary
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ProcessedBatchLog.ps1`
Expected: throws `Source file not found: ...ProcessedBatchLog.vb`.

- [ ] **Step 3: Implement**

Create `ProcessedBatchLog.vb`:

```vb
' Which bill files this workbook has fully processed, keyed by content hash (HashFile).
'
' The log is a hidden sheet in the WORKING tracker workbook itself, not a file beside it:
' the tracker is what gets shared, copied and rolled over each year, and the log has to
' travel with it. The sheet is created on first write, so rolling this out needs no manual
' migration step. Sheet name comes from TenantProcessedBatchSheet(). See the batch-tracking
' spec, section 1.
'
' Row layout: HASH | SOURCE FILE | TENANT | PROCESSED AT | ROWS ADDED | ROWS SKIPPED

Private Const PBL_HEADERS As String = "HASH|SOURCE FILE|TENANT|PROCESSED AT|ROWS ADDED|ROWS SKIPPED"

Public Function IsAlreadyProcessed(ByVal wb As Workbook, ByVal hash As String) As Boolean
    IsAlreadyProcessed = (PblFindRow(wb, hash) > 0)
End Function

' One line describing the earlier run of this file, for the "already processed" message.
Public Function ProcessedBatchSummary(ByVal wb As Workbook, ByVal hash As String) As String
    Dim r As Long
    r = PblFindRow(wb, hash)
    If r = 0 Then Exit Function

    With wb.Worksheets(TenantProcessedBatchSheet())
        ProcessedBatchSummary = CStr(.Cells(r, 2).Value) & ", processed " & _
            Format$(.Cells(r, 4).Value, "mm/dd/yyyy h:nn AM/PM") & ": " & _
            CStr(.Cells(r, 5).Value) & " added, " & CStr(.Cells(r, 6).Value) & " skipped"
    End With
End Function

Public Sub RecordProcessedBatch(ByVal wb As Workbook, ByVal hash As String, _
                                ByVal sourceFile As String, ByVal tenant As String, _
                                ByVal rowsAdded As Long, ByVal rowsSkipped As Long)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = PblEnsureSheet(wb)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = LCase$(Trim$(hash))
    ws.Cells(r, 2).Value = sourceFile
    ws.Cells(r, 3).Value = tenant
    ws.Cells(r, 4).Value = Now
    ws.Cells(r, 5).Value = rowsAdded
    ws.Cells(r, 6).Value = rowsSkipped
End Sub

Private Function PblFindRow(ByVal wb As Workbook, ByVal hash As String) As Long
    Dim key As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    key = LCase$(Trim$(hash))
    If Len(key) = 0 Then Exit Function
    If Not PblSheetExists(wb, TenantProcessedBatchSheet()) Then Exit Function

    Set ws = wb.Worksheets(TenantProcessedBatchSheet())
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = key Then
            PblFindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function PblEnsureSheet(ByVal wb As Workbook) As Worksheet
    Dim name As String
    Dim ws As Worksheet
    Dim prior As Object
    Dim headers As Variant

    name = TenantProcessedBatchSheet()
    If PblSheetExists(wb, name) Then
        Set PblEnsureSheet = wb.Worksheets(name)
        Exit Function
    End If

    ' Adding a sheet activates it; put the user back where they were.
    Set prior = wb.ActiveSheet
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    headers = Split(PBL_HEADERS, "|")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).Value = headers
    ws.Visible = xlSheetHidden

    On Error Resume Next
    If Not prior Is Nothing Then prior.Activate
    On Error GoTo 0

    Set PblEnsureSheet = ws
End Function

Private Function PblSheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            PblSheetExists = True
            Exit Function
        End If
    Next ws
End Function
```

- [ ] **Step 4: Run it and confirm it passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ProcessedBatchLog.ps1`
Expected: `ALL PASS`, exit 0.

- [ ] **Step 5: Update the README and commit**

Append to the "Batch tracking" section of `README.md`:

```markdown
`ProcessedBatchLog.vb` keeps the log on a hidden sheet inside the working tracker, named by
`TenantProcessedBatchSheet()`, one row per processed file: hash, file name, tenant, time, rows
added, rows skipped. The sheet is created on first write. Every tenant that stacks core must
declare `TenantProcessedBatchSheet()` and `TenantProcessedFolder()` (ADR-0003).
```

```bash
git add ProcessedBatchLog.vb tests/Test-ProcessedBatchLog.ps1 README.md
git commit -m "feat: add ProcessedBatchLog hidden-sheet log"
```

---

### Task 3: `MoveProcessedFile` (core)

**Files:**
- Create: `InvoiceTrackerCore/MoveProcessedFile.vb`
- Create: `InvoiceTrackerCore/tests/Test-MoveProcessedFile.ps1`
- Modify: `InvoiceTrackerCore/README.md` (the "Batch tracking" section)

**Interfaces:**
- Consumes: `GetBaseFileName(path) As String` and `GetFileExt(path) As String`, where the extension has no dot (both are existing core functions, `GetBaseFileName.vb` and `GetFileExt.vb`). Also `FileNameFromPath(path) As String` (`GetBaseFileName.vb`).
- Produces: `Public Function MoveProcessedFile(ByVal filePath As String, ByVal subfolderName As String) As String`. It returns the new full path, or `""` if nothing moved: an empty subfolder name, a URL, a missing file, or any file-system error. If the target name is taken, it appends ` (yyyymmdd-hhnnss)`.

- [ ] **Step 1: Write the failing test**

Create `tests/Test-MoveProcessedFile.ps1`:

```powershell
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MoveProcessedFile.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$dir = Join-Path $env:TEMP ('mpf-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$first = Join-Path $dir 'bills.csv'
Set-Content -LiteralPath $first -Value 'one'

$srcFiles = @(
    (Join-Path $repo 'GetBaseFileName.vb'),
    (Join-Path $repo 'GetFileExt.vb'),
    (Join-Path $repo 'MoveProcessedFile.vb')
)
$vba = New-VbaHost -SourceFiles $srcFiles
try {
    $run = { param($a) Invoke-VbaFunction -VbaHost $vba -Name 'MoveProcessedFile' -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal -Expected '' -Actual (& $run @($first, '')) -Because 'an empty subfolder name disables the move'
    Assert-Equal -Expected $true -Actual (Test-Path -LiteralPath $first) -Because 'the file is untouched when the move is disabled'

    $expected = Join-Path $dir 'Processed\bills.csv'
    Assert-Equal -Expected $expected -Actual (& $run @($first, 'Processed')) -Because 'moves into a new sibling subfolder'
    Assert-Equal -Expected $false -Actual (Test-Path -LiteralPath $first) -Because 'the original is gone after the move'
    Assert-Equal -Expected $true  -Actual (Test-Path -LiteralPath $expected) -Because 'the file is at the returned path'

    Set-Content -LiteralPath $first -Value 'two'
    $second = [string](& $run @($first, 'Processed'))
    Assert-Equal -Expected $true -Actual ($second -match '\\Processed\\bills \(\d{8}-\d{6}\)\.csv$') `
                 -Because "a name collision gets a timestamp suffix ($second)"
    Assert-Equal -Expected 'one' -Actual ((Get-Content -LiteralPath $expected -Raw).Trim()) `
                 -Because 'the earlier processed file is not overwritten'

    Assert-Equal -Expected '' -Actual (& $run @((Join-Path $dir 'missing.csv'), 'Processed')) -Because 'a missing file returns ""'
    Assert-Equal -Expected '' -Actual (& $run @('https://example.com/x.csv', 'Processed')) -Because 'a URL returns ""'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles $srcFiles -DocumentModule
try {
    Assert-Equal -Expected '' -Actual ($vba.Workbook.MoveProcessedFile((Join-Path $dir 'missing.csv'), 'Processed')) `
                 -Because 'MoveProcessedFile compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
Write-AssertSummary
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MoveProcessedFile.ps1`
Expected: throws `Source file not found: ...MoveProcessedFile.vb`.

- [ ] **Step 3: Implement**

Create `MoveProcessedFile.vb`:

```vb
' Moves a processed bill file into a subfolder beside it, and returns the new path, or "" if
' it did not move.
'
' Cosmetic only: nothing decides anything from where a file sits. The content hash in
' ProcessedBatchLog is what stops a re-run. So every failure here is a quiet "" for the
' caller to mention, never an error. An empty subfolderName (TenantProcessedFolder) disables
' the move.
Public Function MoveProcessedFile(ByVal filePath As String, ByVal subfolderName As String) As String
    Dim folder As String
    Dim destFolder As String
    Dim target As String
    Dim ext As String

    MoveProcessedFile = vbNullString
    On Error GoTo Failed

    If Len(Trim$(subfolderName)) = 0 Then Exit Function
    If LCase$(Left$(filePath, 4)) = "http" Then Exit Function
    If Len(Dir$(filePath)) = 0 Then Exit Function

    folder = Left$(filePath, InStrRev(filePath, "\"))
    destFolder = folder & Trim$(subfolderName)
    If Len(Dir$(destFolder, vbDirectory)) = 0 Then MkDir destFolder

    target = destFolder & "\" & FileNameFromPath(filePath)
    If Len(Dir$(target)) > 0 Then
        ext = GetFileExt(filePath)
        target = destFolder & "\" & GetBaseFileName(filePath) & " (" & Format$(Now, "yyyymmdd-hhnnss") & ")" & _
                 IIf(Len(ext) > 0, "." & ext, "")
    End If

    Name filePath As target
    MoveProcessedFile = target
    Exit Function

Failed:
    MoveProcessedFile = vbNullString
End Function
```

- [ ] **Step 4: Run it and confirm it passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MoveProcessedFile.ps1`
Expected: `ALL PASS`, exit 0.

- [ ] **Step 5: Update the README and commit**

Append to the "Batch tracking" section of `README.md`:

```markdown
`MoveProcessedFile.vb` moves a finished file into the subfolder `TenantProcessedFolder()` names,
beside the file. A name collision gets a timestamp suffix. The move is cosmetic: an empty folder
name disables it, and a failed move is reported but changes nothing, since the hash log is the
record.
```

```bash
git add MoveProcessedFile.vb tests/Test-MoveProcessedFile.ps1 README.md
git commit -m "feat: add MoveProcessedFile for done bill files"
```

---

### Task 4: `BillCodeScreen` (core)

**Files:**
- Create: `InvoiceTrackerCore/BillCodeScreen.vb`
- Create: `InvoiceTrackerCore/tests/Test-BillCodeScreen.ps1`
- Modify: `InvoiceTrackerCore/README.md`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `Public Function ClassifyBillCodes(ByVal existingCodes As Variant, ByVal incomingCodes As Variant, ByVal allowRepeatsInFile As Boolean) As Variant`. It returns an array with the same bounds as `incomingCodes`. Each element is `""` (write it), `"already-processed"` or `"repeat-in-file"`. If `incomingCodes` is empty, it returns `Array()`.
  - `Public Function BillCodeColumnValues(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long, ByVal lastRow As Long) As Variant`, which returns a 0-based 1-D array (`Array()` if `lastRow < firstRow`).
- Rules:
  - Codes are compared trimmed and case-insensitive.
  - A blank code or an error cell is never flagged.
  - `already-processed` takes precedence over `repeat-in-file`.
  - When `allowRepeatsInFile = False`, **every** row carrying a repeated code is flagged, not just the second and later ones.

- [ ] **Step 1: Write the failing test**

Create `tests/Test-BillCodeScreen.ps1`:

```powershell
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-BillCodeScreen.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

# Drivers take "|"-joined strings so no array has to cross the COM boundary, and map ""
# to "ok" so expected values stay readable.
$driver = Join-Path $env:TEMP 'BcsDriver.vb'
@'
Private Function BcsTestRun(ByVal existingJoined As String, ByVal incomingJoined As String, ByVal allow As Boolean) As String
    Dim existing As Variant, incoming As Variant, result As Variant, i As Long, parts() As String
    existing = Split(existingJoined, "|")
    incoming = Split(incomingJoined, "|")
    result = ClassifyBillCodes(existing, incoming, allow)
    If UBound(result) < LBound(result) Then BcsTestRun = "(empty)": Exit Function
    ReDim parts(LBound(result) To UBound(result))
    For i = LBound(result) To UBound(result)
        parts(i) = IIf(result(i) = "", "ok", result(i))
    Next i
    BcsTestRun = Join(parts, "|")
End Function

Public Function T_Repair(ByVal e As String, ByVal n As String) As String
    T_Repair = BcsTestRun(e, n, False)
End Function

Public Function T_Monitoring(ByVal e As String, ByVal n As String) As String
    T_Monitoring = BcsTestRun(e, n, True)
End Function

Public Function T_Column(ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim v As Variant, i As Long, s As String
    v = BillCodeColumnValues(ThisWorkbook.Worksheets(1), 1, firstRow, lastRow)
    s = CStr(UBound(v) - LBound(v) + 1) & ":"
    For i = LBound(v) To UBound(v)
        s = s & CStr(v(i)) & ";"
    Next i
    T_Column = s
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

$src = Join-Path $repo 'BillCodeScreen.vb'
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $r = { param($n, $a) Invoke-VbaFunction -VbaHost $vba -Name $n -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('A|B', 'C|D')) -Because 'repair: new codes pass'
    Assert-Equal 'already-processed|ok' (& $r 'T_Repair' @('A|B', 'B|C')) -Because 'repair: a code already in the tracker is skipped'
    Assert-Equal 'repeat-in-file|repeat-in-file|ok' (& $r 'T_Repair' @('', 'C|C|D')) -Because 'repair: every row of a repeated code is skipped'
    Assert-Equal 'already-processed|already-processed' (& $r 'T_Repair' @('C', 'C|C')) -Because 'repair: already-processed wins over repeat-in-file'
    Assert-Equal 'ok|ok|ok' (& $r 'T_Monitoring' @('', 'C|C|D')) -Because 'monitoring: a shared code across store rows passes'
    Assert-Equal 'already-processed|already-processed|ok' (& $r 'T_Monitoring' @('C', 'C|C|D')) -Because 'monitoring: a whole already-processed group is skipped'
    Assert-Equal 'already-processed' (& $r 'T_Repair' @('  abc ', 'ABC')) -Because 'match is trimmed and case-insensitive'
    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('', '|X')) -Because 'a blank incoming code is never flagged'
    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('', '|')) -Because 'two blanks are not a repeat'

    $ws = $vba.Workbook.Worksheets(1)
    $ws.Range('A2').Value2 = [double]6005463095
    $ws.Range('A3').NumberFormat = '@'
    $ws.Range('A3').Value2 = '6005463095'
    Assert-Equal '2:6005463095;6005463095;' (& $r 'T_Column' @(2, 3)) -Because 'numeric and text codes read back as the same string'
    Assert-Equal '1:6005463095;' (& $r 'T_Column' @(2, 2)) -Because 'a one-cell range still returns an array'
    Assert-Equal '0:' (& $r 'T_Column' @(3, 2)) -Because 'an empty range returns an empty array'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($src, $driver) -DocumentModule
try {
    Assert-Equal 'ok|ok|ok' ($vba.Workbook.T_Monitoring('', 'C|C|D')) -Because 'BillCodeScreen compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-BillCodeScreen.ps1`
Expected: throws `Source file not found: ...BillCodeScreen.vb`.

- [ ] **Step 3: Implement**

Create `BillCodeScreen.vb`:

```vb
' Decides, before anything is written, which incoming bill rows to skip.
'
' Two different tests, because Securitas monitoring bills put one bill code on several store
' rows and repair/installation bills never do (see the batch-tracking spec, section 2):
'   already-processed  the code is already in the tracker (the rows BEFORE this import)
'   repeat-in-file     the code appears on more than one incoming row, and the caller said
'                      repeats are not allowed for this bill type
' The old check ran after the rows were written and counted the whole column, new rows
' included, so every monitoring row after the first looked like a duplicate of its siblings.
'
' Knows nothing about bill types: the caller decides allowRepeatsInFile. Pure over arrays,
' so the core harness can test it without a tracker.
'
' Every row of a repeated code is flagged, not just the later ones: the tool cannot tell
' which copy is right, so a person decides.
Public Function ClassifyBillCodes(ByVal existingCodes As Variant, ByVal incomingCodes As Variant, _
                                  ByVal allowRepeatsInFile As Boolean) As Variant
    Dim existing As Object
    Dim counts As Object
    Dim result() As Variant
    Dim i As Long
    Dim k As String

    If Not IsArray(incomingCodes) Then ClassifyBillCodes = Array(): Exit Function
    If UBound(incomingCodes) < LBound(incomingCodes) Then ClassifyBillCodes = Array(): Exit Function

    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    If IsArray(existingCodes) Then
        For i = LBound(existingCodes) To UBound(existingCodes)
            k = BcsKey(existingCodes(i))
            If Len(k) > 0 Then existing(k) = True
        Next i
    End If

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    For i = LBound(incomingCodes) To UBound(incomingCodes)
        k = BcsKey(incomingCodes(i))
        If Len(k) > 0 Then counts(k) = counts(k) + 1
    Next i

    ReDim result(LBound(incomingCodes) To UBound(incomingCodes))
    For i = LBound(incomingCodes) To UBound(incomingCodes)
        k = BcsKey(incomingCodes(i))
        If Len(k) = 0 Then
            result(i) = ""
        ElseIf existing.Exists(k) Then
            result(i) = "already-processed"
        ElseIf Not allowRepeatsInFile And counts(k) > 1 Then
            result(i) = "repeat-in-file"
        Else
            result(i) = ""
        End If
    Next i

    ClassifyBillCodes = result
End Function

' One column of a sheet as a 0-based 1-D array. Range.Value returns a bare value, not an array,
' for a single cell, which is the case this exists to absorb.
Public Function BillCodeColumnValues(ByVal ws As Worksheet, ByVal col As Long, _
                                     ByVal firstRow As Long, ByVal lastRow As Long) As Variant
    Dim raw As Variant
    Dim result() As Variant
    Dim i As Long

    If lastRow < firstRow Then BillCodeColumnValues = Array(): Exit Function

    raw = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value
    ReDim result(0 To lastRow - firstRow)
    If IsArray(raw) Then
        For i = 0 To lastRow - firstRow
            result(i) = raw(i + 1, 1)
        Next i
    Else
        result(0) = raw
    End If
    BillCodeColumnValues = result
End Function

Private Function BcsKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    BcsKey = Trim$(CStr(v))
End Function
```

- [ ] **Step 4: Run it and confirm it passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-BillCodeScreen.ps1`
Expected: `ALL PASS`, exit 0.

- [ ] **Step 5: Update the README and commit**

Append to `README.md`, before `## Design docs`:

```markdown
## Screening bill codes before import

`BillCodeScreen.vb`'s `ClassifyBillCodes` marks each incoming row `already-processed` (its code
is in the tracker's existing rows), `repeat-in-file` (its code is on more than one incoming row,
and the caller disallows that), or `""` (write it). It runs **before** any row is written, so
skipped rows never land and nothing has to be marked red or deleted. It knows nothing about bill
types: Securitas passes `allowRepeatsInFile = True` for monitoring bills, which legitimately put
one bill code on several store rows.
```

```bash
git add BillCodeScreen.vb tests/Test-BillCodeScreen.ps1 README.md
git commit -m "feat: add ClassifyBillCodes pre-write screen"
```

---

### Task 5: `MirrorInvoiceBlock` (core)

**Files:**
- Create: `InvoiceTrackerCore/MirrorInvoiceBlock.vb`
- Create: `InvoiceTrackerCore/tests/Test-MirrorInvoiceBlock.ps1`
- Modify: `InvoiceTrackerCore/README.md`

**Interfaces:**
- Consumes: nothing.
- Produces: `Public Function MirrorInvoiceBlock(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal writeTag As String, ByVal ownedTags As Variant) As Long`. It returns the number of rows mirrored, or `-1` (writing nothing) in any of these cases:
  - the headers differ;
  - `dst` has a column after `Source File`;
  - `src` has no `BILL CODE` header;
  - `src` has no data rows;
  - an error occurs.
- The contract:
  - Row 1 is the header on both sheets.
  - `dst` headers are exactly the `src` headers (trimmed, case-insensitive, same order) plus `Source File` immediately after.
  - `dst` rows whose `Source File` is in `ownedTags` are deleted.
  - All `src` data rows are appended as values, tagged `writeTag`.
  - Legacy rows are never rewritten.

- [ ] **Step 1: Write the failing test**

Create `tests/Test-MirrorInvoiceBlock.ps1`:

```powershell
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MirrorInvoiceBlock.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$driver = Join-Path $env:TEMP 'MibDriver.vb'
@'
Public Function T_Mirror() As Long
    T_Mirror = MirrorInvoiceBlock(ThisWorkbook.Worksheets("Src"), ThisWorkbook.Worksheets("Dst"), _
                                  "2026:Invoices", Array("2026:Invoices", "2026Model:Invoices"))
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

$xlUp = -4162

function New-Sheets($wb) {
    $dst = $wb.Worksheets.Add(); $dst.Name = 'Dst'
    $src = $wb.Worksheets.Add(); $src.Name = 'Src'
    @{ Src = $src; Dst = $dst }
}

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        $cell = $ws.Cells($row, $c + 1)
        if ($v -is [string] -and $v.StartsWith('=')) { $cell.Formula = $v }
        elseif ($v -is [string]) { $cell.NumberFormat = '@'; $cell.Value2 = $v }
        else { $cell.Value2 = $v }
    }
}

function Get-LastRow($ws, [int]$col) { $ws.Cells($ws.Rows.Count, $col).End($xlUp).Row }

function Build-Fixture($s) {
    Set-Row $s.Src 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES')
    # A monitoring bill: three store rows sharing one code.
    Set-Row $s.Src 2 @('0003', 'M1', 10, 'store a')
    Set-Row $s.Src 3 @('0071', 'M1', 20, 'store b')
    Set-Row $s.Src 4 @('0227', 'M1', 30, 'store c')
    Set-Row $s.Src 5 @('0414', 'R9', '=2*50', 'repair')

    Set-Row $s.Dst 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES', 'Source File')
    Set-Row $s.Dst 2 @('0175', 'L1', 1, 'legacy one', '2025:Invoices')
    Set-Row $s.Dst 3 @('0999', 'M1', 60, 'collapsed by migration', '2026Model:Invoices')
    Set-Row $s.Dst 4 @('0176', 'L2', 2, 'PAID', 'RootNoYear:Invoices')
    Set-Row $s.Dst 5 @('0998', 'OLD', 5, 'stale mirror row', '2026:Invoices')
    Set-Row $s.Dst 6 @('0177', 'L3', 3, 'legacy three', '2025:Invoices')
}

$src = Join-Path $repo 'MirrorInvoiceBlock.vb'
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $s = New-Sheets $vba.Workbook
    Build-Fixture $s
    $d = $s.Dst
    $mirror = { Invoke-VbaFunction -VbaHost $vba -Name 'T_Mirror' -TimeoutSeconds 60 }

    Assert-Equal 4 (& $mirror) -Because 'four source rows mirrored'
    Assert-Equal 8 (Get-LastRow $d 5) -Because '3 legacy + 4 mirrored rows, header in row 1'
    Assert-Equal 'L1|L2|L3' (($d.Cells(2,2).Value2, $d.Cells(3,2).Value2, $d.Cells(4,2).Value2) -join '|') `
                 -Because 'legacy rows keep their order, owned rows (both tags) are gone'
    Assert-Equal 'PAID' ($d.Cells(3,4).Value2) -Because 'a hand-set PAID on a legacy row survives'
    Assert-Equal 'M1|M1|M1|R9' (($d.Cells(5,2).Value2, $d.Cells(6,2).Value2, $d.Cells(7,2).Value2, $d.Cells(8,2).Value2) -join '|') `
                 -Because 'all three monitoring store rows arrive, not one'
    Assert-Equal '0003' ($d.Cells(5,1).Value2) -Because 'a text store number keeps its leading zeros'
    Assert-Equal '0175' ($d.Cells(2,1).Value2) -Because 'a legacy text store number is not rewritten to a number'
    Assert-Equal 100 ($d.Cells(8,3).Value2) -Because 'a formula arrives as its value'
    Assert-Equal $false ($d.Cells(8,3).HasFormula) -Because 'no formula is copied'
    Assert-Equal '2026:Invoices' ($d.Cells(8,5).Value2) -Because 'mirrored rows get the write tag'

    Assert-Equal 4 (& $mirror) -Because 'a re-run mirrors the same four rows'
    Assert-Equal 8 (Get-LastRow $d 5) -Because 'a re-run does not duplicate anything'

    $s.Src.Rows(5).Delete() | Out-Null
    Assert-Equal 3 (& $mirror) -Because 'a smaller source mirrors fewer rows'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'the dropped source row is gone from the mirror'

    $d.Cells(1,3).Value2 = 'AMOUNT'
    Assert-Equal -1 (& $mirror) -Because 'a header mismatch refuses'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'a refused mirror writes nothing'
    $d.Cells(1,3).Value2 = ' TOTAL '

    $s.Src.Range('A2:D10').ClearContents() | Out-Null
    Assert-Equal -1 (& $mirror) -Because 'an empty source refuses rather than wiping the owned block'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'the owned block survives an empty source'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($src, $driver) -DocumentModule
try {
    New-Sheets $vba.Workbook | Out-Null
    Assert-Equal -1 ($vba.Workbook.T_Mirror()) -Because 'MirrorInvoiceBlock compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MirrorInvoiceBlock.ps1`
Expected: throws `Source file not found: ...MirrorInvoiceBlock.vb`.

- [ ] **Step 3: Implement**

Create `MirrorInvoiceBlock.vb`:

```vb
' Replaces dst's owned block with a values-only copy of src, and returns the rows copied, or
' -1 if it wrote nothing.
'
' Built for the All-Years archive (batch-tracking spec, section 3): src is the working book's
' Invoices sheet, dst the archive's, which has the same columns plus a trailing "Source File"
' tag. Rows whose tag is in ownedTags belong to the mirror and are replaced wholesale on every
' run; every other row is history and is never touched.
'
' Whole-block rather than row matching, because no row key is unique: a monitoring bill puts
' one BILL CODE on several store rows in both workbooks. Replacing the block needs no key,
' and re-running it is always safe.
'
' Owned rows are deleted run by run and the new block appended, rather than reading the whole
' sheet and writing it back, because writing a legacy text value like "0175" back into a
' General cell would silently turn it into the number 175.
Public Function MirrorInvoiceBlock(ByVal src As Worksheet, ByVal dst As Worksheet, _
                                   ByVal writeTag As String, ByVal ownedTags As Variant) As Long
    Dim srcCols As Long, tagCol As Long, keyCol As Long, c As Long
    Dim srcLast As Long, dstLast As Long, n As Long, r As Long, runEnd As Long, newFirst As Long
    Dim tags As Variant, data As Variant, one() As Variant

    MirrorInvoiceBlock = -1
    On Error GoTo Failed

    ' 1. Headers must line up exactly.
    srcCols = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    tagCol = srcCols + 1
    For c = 1 To srcCols
        If MibHeader(src.Cells(1, c).Value) <> MibHeader(dst.Cells(1, c).Value) Then Exit Function
        If MibHeader(src.Cells(1, c).Value) = "BILL CODE" Then keyCol = c
    Next c
    If keyCol = 0 Then Exit Function
    If MibHeader(dst.Cells(1, tagCol).Value) <> "SOURCE FILE" Then Exit Function
    If dst.Cells(1, dst.Columns.Count).End(xlToLeft).Column <> tagCol Then Exit Function

    ' 2. Real last rows, never UsedRange (the archive's is stale from old deletions).
    srcLast = MibLastRow(src, 1, keyCol)
    If srcLast < 2 Then Exit Function          ' an empty source must not wipe the owned block
    n = srcLast - 1
    dstLast = MibLastRow(dst, 1, tagCol)

    ' 3. Delete owned rows, bottom-up, one delete per contiguous run. Tags are read once up
    '    front; deleting below r never shifts rows at or above r, so the indexes stay valid.
    If dstLast >= 2 Then
        tags = dst.Range(dst.Cells(2, tagCol), dst.Cells(dstLast, tagCol)).Value
        r = dstLast
        Do While r >= 2
            If MibOwned(MibAt(tags, r - 1), ownedTags) Then
                runEnd = r
                Do While r >= 2
                    If Not MibOwned(MibAt(tags, r - 1), ownedTags) Then Exit Do
                    r = r - 1
                Loop
                dst.Rows((r + 1) & ":" & runEnd).Delete
            Else
                r = r - 1
            End If
        Loop
    End If

    ' 4. Append src as values. Formats first, so text stays text when the values land.
    newFirst = MibLastRow(dst, 1, tagCol) + 1
    If newFirst < 2 Then newFirst = 2
    data = src.Range(src.Cells(2, 1), src.Cells(srcLast, srcCols)).Value
    ' A 1x1 range returns a bare value, not an array.
    If Not IsArray(data) Then ReDim one(1 To 1, 1 To 1): one(1, 1) = data: data = one
    For c = 1 To srcCols
        dst.Range(dst.Cells(newFirst, c), dst.Cells(newFirst + n - 1, c)).NumberFormat = _
            MibFormatFor(data, c, src.Cells(2, c).NumberFormat)
    Next c
    dst.Range(dst.Cells(newFirst, 1), dst.Cells(newFirst + n - 1, srcCols)).Value = data
    dst.Range(dst.Cells(newFirst, tagCol), dst.Cells(newFirst + n - 1, tagCol)).Value = writeTag

    MirrorInvoiceBlock = n
    Exit Function

Failed:
    MirrorInvoiceBlock = -1
End Function

Private Function MibHeader(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    MibHeader = UCase$(Trim$(CStr(v)))
End Function

Private Function MibLastRow(ByVal ws As Worksheet, ByVal colA As Long, ByVal colB As Long) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.Count, colA).End(xlUp).Row
    b = ws.Cells(ws.Rows.Count, colB).End(xlUp).Row
    MibLastRow = IIf(a > b, a, b)
End Function

Private Function MibAt(ByVal tags As Variant, ByVal i As Long) As Variant
    If IsArray(tags) Then MibAt = tags(i, 1) Else MibAt = tags
End Function

Private Function MibOwned(ByVal tag As Variant, ByVal ownedTags As Variant) As Boolean
    Dim i As Long
    If IsError(tag) Or IsEmpty(tag) Then Exit Function
    For i = LBound(ownedTags) To UBound(ownedTags)
        If StrComp(Trim$(CStr(tag)), Trim$(CStr(ownedTags(i))), vbTextCompare) = 0 Then
            MibOwned = True
            Exit Function
        End If
    Next i
End Function

' A column holding any numeric-looking text ("0003", "6005463095") is forced to Text, or the
' value write converts it to a number. Otherwise the source's own format carries over.
Private Function MibFormatFor(ByVal data As Variant, ByVal c As Long, ByVal fallback As String) As String
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If VarType(data(i, c)) = vbString Then
            If Len(data(i, c)) > 0 And IsNumeric(data(i, c)) Then
                MibFormatFor = "@"
                Exit Function
            End If
        End If
    Next i
    MibFormatFor = fallback
End Function
```

- [ ] **Step 4: Run it and confirm it passes**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MirrorInvoiceBlock.ps1`
Expected: `ALL PASS`, exit 0.

- [ ] **Step 5: Run every core test**

Run each: `Test-FiscalCalendar`, `Test-DocumentModule`, `Test-ReqJoinKey`, `Test-HashFile`, `Test-ProcessedBatchLog`, `Test-MoveProcessedFile`, `Test-BillCodeScreen`, `Test-MirrorInvoiceBlock`.
Expected: every one exits 0.

- [ ] **Step 6: Update the README and commit**

Append to `README.md`, before `## Design docs`:

```markdown
## Mirroring a tracker into an archive

`MirrorInvoiceBlock.vb` replaces the rows of an archive sheet that belong to the current year,
those whose trailing `Source File` tag is in `ownedTags`, with a values-only copy of a tracker
sheet, and never touches any other row. It refuses, writing nothing and returning `-1`, when:
- the headers differ, or the archive has a column after `Source File`;
- the source has no `BILL CODE` header;
- the source is empty.

Owned rows are deleted and re-appended rather than the sheet being rewritten, because writing
a legacy text value such as `0175` back into a General cell converts it to a number. The
Securitas `SyncAllYears` wrapper supplies the SharePoint side; see the batch-tracking spec,
section 3.
```

```bash
git add MirrorInvoiceBlock.vb tests/Test-MirrorInvoiceBlock.ps1 README.md
git commit -m "feat: add MirrorInvoiceBlock values-only mirror"
```

---

### Task 6: Securitas `AddNewBills` — hash gate, pre-write screen, record, move

**Files:**
- Modify: `SecuritasAutomation/TenantConfig.vb` (append before the `TenantSheetPassword` block)
- Modify: `SecuritasAutomation/file ingesting/AddNewBills.vb:54-58` (after the extension check), `:176-191` (screen before write), `:198-228` (record, move, summary)
- Modify: `SecuritasAutomation/file ingesting/ProcessNewBills.vb:6,44-87`
- Delete: `SecuritasAutomation/file ingesting/CheckNewBillsForDuplicates.vb`

**Interfaces:**
- Consumes (Tasks 1–4):
  - `HashFile`
  - `IsAlreadyProcessed(wb, hash)`
  - `ProcessedBatchSummary(wb, hash)`
  - `RecordProcessedBatch(wb, hash, sourceFile, tenant, rowsAdded, rowsSkipped)`
  - `MoveProcessedFile(path, subfolder)`
  - `ClassifyBillCodes(existing, incoming, allowRepeats)`
  - `BillCodeColumnValues(ws, col, first, last)`
- Produces:
  - `TenantProcessedBatchSheet() As String` = `"Processed Batches"`
  - `TenantProcessedFolder() As String` = `"Processed"`
  - `ProcessNewBills(..., Optional billType As String, Optional rowStatus As Variant)`, where `rowStatus` is the array returned by `ClassifyBillCodes` over rows `sourceHeaderRow+1 … sourceLastRow`.

- [ ] **Step 1: Add the tenant accessors**

In `SecuritasAutomation/TenantConfig.vb`, insert above the `' Password used to protect the tracker sheet.` comment:

```vb
' --- Batch tracking -------------------------------------------------------------

' Hidden sheet in this workbook logging every bill file already processed, by content hash.
' See InvoiceTrackerCore's ProcessedBatchLog.vb.
Public Function TenantProcessedBatchSheet() As String
    TenantProcessedBatchSheet = "Processed Batches"
End Function

' Subfolder, beside the bill file, that a processed file is moved into. "" disables the move.
' Cosmetic only -- the hash log is what stops a re-run.
Public Function TenantProcessedFolder() As String
    TenantProcessedFolder = "Processed"
End Function

```

- [ ] **Step 2: Let `ProcessNewBills` skip rows**

In `ProcessNewBills.vb`, change the signature on line 6 to add a trailing parameter:

```vb
Sub ProcessNewBills(wsSource As Worksheet, sourceHeaderRow As Long, sourceLastRow As Long, sourceColIndexes As Object, wsTarget As Worksheet, targetLastRow As Long, targetHeaderDict As Object, Optional storeResolutions As Object, Optional billType As String, Optional rowStatus As Variant)
```

Add to the header comment block (after line 5):

```vb
'
' rowStatus, when given, is ClassifyBillCodes' verdict for each source row from
' sourceHeaderRow + 1 on. A non-empty entry means skip that row: it is already in the tracker,
' or repeats within the file. Skipped rows are simply not written.
```

Replace the loop header and footer at lines 44 and 86-87, `For row = ...` through `Next row`, so the body is skipped for flagged rows:

```vb
    For row = sourceHeaderRow + 1 To sourceLastRow
        If Not IsMissing(rowStatus) Then
            If Len(rowStatus(LBound(rowStatus) + row - sourceHeaderRow - 1)) > 0 Then GoTo NextSourceRow
        End If
```

…leave the existing column loop unchanged…

```vb
        targetRow = targetRow + 1
NextSourceRow:
    Next row
```

- [ ] **Step 3: Add the hash gate to `AddNewBills`**

In `AddNewBills.vb`, declare with the other `Dim`s at the top of the sub:

```vb
    Dim fileHash As String
    Dim rowStatus As Variant
    Dim rowsToAdd As Long
    Dim rowsSkipped As Long
    Dim skippedList As String
    Dim importFailed As Boolean
    Dim moveNote As String
    Dim i As Long
```

Insert immediately after the extension check (after line 58, `End If`):

```vb
    ' Identify the file by content before opening it, so a re-run of a file already processed
    ' stops here with nothing opened or written. Batch-tracking spec, section 1.
    fileHash = HashFile(selectedFile)
    If Len(fileHash) = 0 Then
        MsgBox "Could not read this file to check whether it was already processed:" & vbCrLf & vbCrLf & _
               selectedFile & vbCrLf & vbCrLf & "Nothing was imported.", vbExclamation
        Exit Sub
    End If
    If IsAlreadyProcessed(ThisWorkbook, fileHash) Then
        MsgBox "This file was already processed:" & vbCrLf & vbCrLf & _
               ProcessedBatchSummary(ThisWorkbook, fileHash) & vbCrLf & vbCrLf & _
               "Nothing was imported.", vbInformation
        Exit Sub
    End If
```

- [ ] **Step 4: Screen bill codes before writing**

Replace line 190-191 (the `ProcessNewBills` comment and call) with:

```vb
            ' Decide which incoming rows to skip before anything is written. Monitoring bills
            ' put one bill code on several store rows, so only they may repeat within the file.
            ' Batch-tracking spec, section 2.
            If Not targetHeaderDict.Exists("BILL CODE") Then
                wb.Close False
                RestoreThinking
                MsgBox "The 'BILL CODE' column was not found in the " & wsTarget.Name & " sheet.", vbExclamation
                Exit Sub
            End If
            rowStatus = ClassifyBillCodes( _
                BillCodeColumnValues(wsTarget, targetHeaderDict("BILL CODE"), 2, targetLastRow), _
                BillCodeColumnValues(wsSource, sourceColIndexes("BILL CODE"), sourceHeaderRow + 1, sourceLastRow), _
                (billType = "monitoring"))
            For i = LBound(rowStatus) To UBound(rowStatus)
                If Len(rowStatus(i)) = 0 Then
                    rowsToAdd = rowsToAdd + 1
                Else
                    rowsSkipped = rowsSkipped + 1
                    If rowsSkipped <= 25 Then
                        skippedList = skippedList & vbCrLf & "  " & _
                            CStr(wsSource.Cells(sourceHeaderRow + 1 + i - LBound(rowStatus), sourceColIndexes("BILL CODE")).Value) & _
                            " (" & rowStatus(i) & ")"
                    End If
                End If
            Next i
            If rowsSkipped > 25 Then skippedList = skippedList & vbCrLf & "  ...and " & (rowsSkipped - 25) & " more"

            ' ProcessNewBills writes only the rows the screen passed.
            If rowsToAdd > 0 Then
                ProcessNewBills wsSource, sourceHeaderRow, sourceLastRow, sourceColIndexes, wsTarget, _
                                targetLastRow, targetHeaderDict, storeResolutions, billType, rowStatus
            End If
```

- [ ] **Step 5: Record, move and summarise; remove the old duplicate check**

Wrap `PopulateInvoiceTypeForNewRows` (lines 194-196) so it only runs when rows were added. Change its `If` to:

```vb
            If Not failureState And Not hasInvoiceTypeCol And rowsToAdd > 0 Then
```

Inside the `If failureState Then` rollback block (after `withErrors = True`, line 215), add:

```vb
                importFailed = True
```

Directly after that `End If` (line 217), add a check that the rows really landed. `ProcessNewBills` has two early exits, a column-mapping error and an empty map, that show a message without setting `failureState`. Without this check, a file that wrote nothing would still be recorded as processed:

```vb
            ' Recording a file as processed must mean its rows are actually here.
            If Not importFailed And rowsToAdd > 0 Then
                If wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row - targetLastRow <> rowsToAdd Then
                    importFailed = True
                End If
            End If
```

Delete lines 219-220 (the `CheckNewBillsForDuplicates` comment and call).

Replace the tail from `wb.Close False` (line 222) through `RestoreThinking` (line 228) with:

```vb
            wb.Close False

            ' Self-manage protection: this sub protects the sheet before its post-import steps.
            Call CheckOpeningDate(False, True)
            Call UpdateSearchValues

            ' Record only a file that fully imported. A rolled-back file can be re-run as is.
            ' The move comes after wb.Close, which releases the file.
            If Not importFailed Then
                RecordProcessedBatch ThisWorkbook, fileHash, FileNameFromPath(selectedFile), _
                                     TenantSupplierKeyword(), rowsToAdd, rowsSkipped
                If Len(TenantProcessedFolder()) > 0 Then
                    If Len(MoveProcessedFile(selectedFile, TenantProcessedFolder())) = 0 Then
                        moveNote = vbCrLf & vbCrLf & "The file could not be moved into the '" & _
                                   TenantProcessedFolder() & "' folder. It is still recorded as processed."
                    End If
                End If
            End If

            RestoreThinking

            If Not importFailed Then
                MsgBox "Added " & rowsToAdd & " row(s)." & _
                       IIf(rowsSkipped > 0, vbCrLf & vbCrLf & "Skipped " & rowsSkipped & " row(s):" & skippedList, "") & _
                       moveNote, IIf(rowsSkipped > 0, vbExclamation, vbInformation)
            End If
```

Then delete the old module:

```bash
git -C SecuritasAutomation rm "file ingesting/CheckNewBillsForDuplicates.vb"
```

- [ ] **Step 6: Build the stack and check that nothing still calls the old routine**

```powershell
cd C:\Users\p4bn\Documents\SecuritasAutomation
powershell -NoProfile -ExecutionPolicy Bypass -File .\Stack-VBFiles.ps1
Select-String -Path .\*MegaStack*.vb -Pattern 'CheckNewBillsForDuplicates' | Measure-Object | % Count
Select-String -Path .\*MegaStack*.vb -Pattern '^(Public |Private )?(Function|Sub) (HashFile|IsAlreadyProcessed|RecordProcessedBatch|MoveProcessedFile|ClassifyBillCodes|BillCodeColumnValues|MirrorInvoiceBlock)\b' | Measure-Object | % Count
```
Expected: the stack builds with exit 0. The first count is `0`. The second is `7`, one definition each.

- [ ] **Step 7: Verify by hand, on a scratch copy (the user drives this)**

The user runs these steps. They can't be automated here: the stack is pasted into a workbook, and the live file must not be touched.
1. Download a copy of `2026 SECURITAS bills.xlsm` to a local folder and open **the copy**.
2. In the VBE, replace `ThisWorkbook`'s code with the new `Securitas-Invoice-Tracker_MegaStack.vb`. Run `Debug > Compile VBAProject`; it should raise no error.
3. Copy a real recent **monitoring** bill file into a scratch folder and run `AddNewBills` on it. Expect: every store row added, none red, and no `DUPLICATE` in REQ #. The summary says "Added N row(s)", and the file has moved into `Processed\`.
4. Copy the same file back from `Processed\` into the scratch folder and run `AddNewBills` again. Expect "This file was already processed: …", with nothing written.
5. Take a **repair** bill file, duplicate one data row in it, save it under a new name, and run `AddNewBills`. Expect both copies of that code skipped as `repeat-in-file`, and every other row added.
6. Take a repair file whose bill codes are already in the tracker (for example, copy three existing rows into a new file). Expect every row skipped as `already-processed`, and "Added 0 row(s)".
7. Unhide the `Processed Batches` sheet: it should have one row per successful run.

Note for this and every later verification step that runs `AddNewBills` (which syncs to
All-Years on success): this scratch copy of the working book only ever syncs to an archive
sitting in ITS OWN folder (Task 8, ruling I1) -- there is no archive there yet at this point in
the plan, so the sync fails safe with a warning, and that warning is expected here, not a
regression.

Record what was run and what happened in the commit body.

- [ ] **Step 8: Commit**

```bash
cd C:\Users\p4bn\Documents\SecuritasAutomation
git add TenantConfig.vb "file ingesting/AddNewBills.vb" "file ingesting/ProcessNewBills.vb"
git commit -m "feat(AddNewBills): hash gate + pre-write screen" -m "Monitoring bills share one bill code across store rows; only they may
repeat within a file. Replaces the post-write CheckNewBillsForDuplicates.
Verified on a scratch copy: <what was driven>."
```

---

### Task 7: JCI `AddNewBills` — hash gate, record, move

**Files:**
- Modify: `JCI-invoice-tracker/TenantConfig.vb` (append before the `TenantSheetPassword` block)
- Modify: `JCI-invoice-tracker/AddNewBills.vb:57-60` (gate), `:150-165` (record, move)

**Interfaces:**
- Consumes: `HashFile`, `IsAlreadyProcessed`, `ProcessedBatchSummary`, `RecordProcessedBatch`, `MoveProcessedFile` (Tasks 1–3).
- Produces: JCI's `TenantProcessedBatchSheet()` = `"Processed Batches"` and `TenantProcessedFolder()` = `"Processed"`.

- [ ] **Step 1: Add the tenant accessors**

In `JCI-invoice-tracker/TenantConfig.vb`, insert above `' Password used to protect the tracker sheet.`:

```vb
' --- Batch tracking -------------------------------------------------------------

' Hidden sheet in this workbook logging every bill file already processed, by content hash.
' See InvoiceTrackerCore's ProcessedBatchLog.vb.
Public Function TenantProcessedBatchSheet() As String
    TenantProcessedBatchSheet = "Processed Batches"
End Function

' Subfolder, beside the bill file, that a processed file is moved into. "" disables the move.
' Cosmetic only -- the hash log is what stops a re-run.
Public Function TenantProcessedFolder() As String
    TenantProcessedFolder = "Processed"
End Function

```

- [ ] **Step 2: Add the hash gate**

Declare with the other `Dim`s:

```vb
    Dim fileHash As String
    Dim moveNote As String
```

Insert after line 57 (`If answer = vbNo Then GoTo ChooseFile`):

```vb
    ' Identify the file by content before reading it, so a re-run of a file already processed
    ' stops here with nothing written. InvoiceTrackerCore batch-tracking spec, section 1.
    fileHash = HashFile(selectedFile)
    If Len(fileHash) = 0 Then
        MsgBox "Could not read this file to check whether it was already processed:" & vbCrLf & vbCrLf & _
               selectedFile & vbCrLf & vbCrLf & "Nothing was imported.", vbExclamation
        Exit Sub
    End If
    If IsAlreadyProcessed(ThisWorkbook, fileHash) Then
        MsgBox "This file was already processed:" & vbCrLf & vbCrLf & _
               ProcessedBatchSummary(ThisWorkbook, fileHash) & vbCrLf & vbCrLf & _
               "Nothing was imported.", vbInformation
        Exit Sub
    End If
```

- [ ] **Step 3: Record and move after a successful import**

Replace line 165 (`MsgBox "Import complete. " ...`) with:

```vb
    ' Only this success path records: ErrorHandler rolls rows back, so a failed file can be
    ' re-run as is. LoadCoupaSourceData has already closed the file, so it can be moved.
    RecordProcessedBatch ThisWorkbook, fileHash, FileNameFromPath(selectedFile), _
                         TenantSupplierKeyword(), rowsAdded, (lastDataRow - headerRow) - rowsAdded
    If Len(TenantProcessedFolder()) > 0 Then
        If Len(MoveProcessedFile(selectedFile, TenantProcessedFolder())) = 0 Then
            moveNote = vbCrLf & vbCrLf & "The file could not be moved into the '" & _
                       TenantProcessedFolder() & "' folder. It is still recorded as processed."
        End If
    End If

    MsgBox "Import complete. " & rowsAdded & " new bill(s) added." & moveNote, vbInformation
```

Before relying on it, confirm that the loader releases the file. Open `InvoiceTrackerCore/CoupaDataHelpers.vb` and check that `LoadCoupaSourceData` closes both paths: the CSV `Open … For Input` is followed by `Close #fileNum`, and the XLSX `Workbooks.Open(… ReadOnly:=True)` at `:96` is followed by a `.Close`. If the XLSX path does not close, stop and report it instead of adding a close here.

- [ ] **Step 4: Build the stack**

```powershell
cd C:\Users\p4bn\Documents\JCI-invoice-tracker
powershell -NoProfile -ExecutionPolicy Bypass -File .\Stack-VBFiles.ps1
```
Expected: exit 0.

- [ ] **Step 5: Verify by hand on a scratch copy (the user drives this)**

On a copy of `JCI Repair & Installation Invoices.xlsm`:
1. Paste the stack and compile.
2. Import a real JCI file. Expect it imported and moved into `Processed\`.
3. Copy it back and import it again. Expect "already processed" and nothing written.

`MarkDuplicateInvoices` is unchanged.

- [ ] **Step 6: Commit**

```bash
cd C:\Users\p4bn\Documents\JCI-invoice-tracker
git add TenantConfig.vb AddNewBills.vb
git commit -m "feat(AddNewBills): refuse already-processed files"
```

---

### Task 8: Securitas `SyncAllYears` and its call sites

**Files:**
- Create: `SecuritasAutomation/SyncAllYears.vb`
- Modify: `SecuritasAutomation/TenantConfig.vb` (All-Years accessors)
- Modify: `SecuritasAutomation/file ingesting/AddNewBills.vb` (staleness warning at the start, sync at the end)
- Modify: `SecuritasAutomation/file ingesting/UpdateCoupaData.vb:8,89` (warning, sync)
- Modify: `SecuritasAutomation/Refresh.vb:63` (sync on success)

**Interfaces:**
- Consumes:
  - `MirrorInvoiceBlock` (Task 5)
  - existing core helpers: `PauseThinking`, `RestoreThinking`, `UnprotectSheetOn(ws)`, `ProtectSheetOn(ws)` (nesting-counted; restores the protection state it found), `TenantSheetName(role)`
- Produces:
  - `Public Function SyncAllYears(Optional ByVal announce As Boolean = False) As Boolean`
  - `Public Sub SyncAllYearsNow()` (the manual macro)
  - `Public Sub WarnIfAllYearsStale()`
  - `TenantAllYearsWorkbookName()`, `TenantAllYearsSyncStampCell()`, `TenantAllYearsStaleDays()`, `TenantAllYearsOwnedTags()`, `TenantAllYearsWriteTag()`

- [ ] **Step 1: Add the tenant accessors**

In `SecuritasAutomation/TenantConfig.vb`, add after the batch-tracking block from Task 6:

```vb
' --- All-Years archive ----------------------------------------------------------
'
' Read by this variant's SyncAllYears.vb only, not by core, so no other tenant has to declare
' these. The archive is a separate cloud workbook that lives in the SAME FOLDER as this
' working book; see the batch-tracking spec, section 3, and ruling I1 in the final-fix report.

' Name only, not a path -- SyncAllYears derives the full path from ThisWorkbook.Path, so a
' scratch copy of the working book only ever reaches an archive in ITS OWN folder.
Public Function TenantAllYearsWorkbookName() As String
    TenantAllYearsWorkbookName = "Securitas All-Years Invoices - Consolidated.xlsm"
End Function

' Helper cell stamped with the last successful sync. E2-E5 already hold Coupa import stamps.
Public Function TenantAllYearsSyncStampCell() As String
    TenantAllYearsSyncStampCell = "E6"
End Function

' A sync older than this many days triggers a warning at the start of an import.
Public Function TenantAllYearsStaleDays() As Long
    TenantAllYearsStaleDays = 3
End Function

' Source File tags the mirror owns and replaces on every sync. "2026Model:Invoices" is the
' label the one-time Sept 2026 consolidation gave rows taken from this workbook; claiming it
' means the first sync replaces those rows instead of duplicating them. Year-bound: the 2027
' workbook declares its own tags, and 2026's block becomes history.
Public Function TenantAllYearsOwnedTags() As Variant
    TenantAllYearsOwnedTags = Array("2026:Invoices", "2026Model:Invoices")
End Function

Public Function TenantAllYearsWriteTag() As String
    TenantAllYearsWriteTag = "2026:Invoices"
End Function

```

- [ ] **Step 2: Write `SyncAllYears.vb`**

```vb
' Mirrors this workbook's Invoices sheet into the All-Years archive (a separate cloud workbook)
' by calling core's MirrorInvoiceBlock. Batch-tracking spec, section 3.
'
' A backup step, never a blocker: every failure shows one warning and returns False. Nothing
' here raises into the caller, and nothing rolls back what the caller already did. Missed
' syncs stay visible through the Helper stamp (WarnIfAllYearsStale), and the next successful
' sync repairs all drift, because it rewrites the whole owned block.
'
' Events stay off for the whole sync, not just the open: the archive carries a copy of this
' stack, and its Workbook_Open and sheet handlers must not fire while its rows are deleted.

Public Sub SyncAllYearsNow()
    SyncAllYears True
End Sub

Public Function SyncAllYears(Optional ByVal announce As Boolean = False) As Boolean
    Dim path As String
    Dim target As Workbook
    Dim openedHere As Boolean
    Dim src As Worksheet
    Dim dst As Worksheet
    Dim mirrored As Long
    Dim problem As String
    Dim priorEvents As Boolean
    Dim priorAlerts As Boolean
    Dim paused As Boolean
    Dim unprotected As Boolean

    path = AyArchivePath()   ' joins ThisWorkbook.Path with TenantAllYearsWorkbookName() -- see ruling I1
    If Len(path) = 0 Then Exit Function

    priorEvents = Application.EnableEvents
    priorAlerts = Application.DisplayAlerts
    On Error GoTo Failed

    PauseThinking
    paused = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set target = AyFindOpen(path)
    If target Is Nothing Then
        ' UpdateLinks:=0: the archive has external links, and the "Update links?" prompt would
        ' otherwise stop the sync.
        Set target = Workbooks.Open(Filename:=path, UpdateLinks:=0)
        openedHere = True
    End If

    If target.ReadOnly Then
        problem = "it opened read-only, so it is probably locked by someone else"
        GoTo Finish
    End If

    Set src = ThisWorkbook.Worksheets(TenantSheetName("tracker"))
    Set dst = target.Worksheets(TenantSheetName("tracker"))

    ' PauseThinking set calculation to manual. Recalculate the source sheet so the mirrored
    ' formula values are current even when a caller is still inside its own pause.
    src.Calculate

    ' Read and write with no prompt in between, to keep the window for a co-author's edit short.
    UnprotectSheetOn dst
    unprotected = True
    mirrored = MirrorInvoiceBlock(src, dst, TenantAllYearsWriteTag(), TenantAllYearsOwnedTags())
    ProtectSheetOn dst
    unprotected = False

    If mirrored < 0 Then
        problem = "its column headers no longer match this workbook's, or this workbook's " & _
                  TenantSheetName("tracker") & " sheet is empty"
        GoTo Finish
    End If

    target.Save
    AyStamp
    SyncAllYears = True

Finish:
    On Error Resume Next
    If unprotected Then ProtectSheetOn dst
    If openedHere And Not target Is Nothing Then target.Close SaveChanges:=False
    Application.DisplayAlerts = priorAlerts
    Application.EnableEvents = priorEvents
    If paused Then RestoreThinking
    On Error GoTo 0

    If SyncAllYears Then
        If announce Then MsgBox "All-Years archive updated: " & mirrored & " row(s) mirrored.", vbInformation
    Else
        MsgBox "The All-Years archive was not updated: " & problem & "." & vbCrLf & vbCrLf & _
               "Your work in this workbook is not affected. Run SyncAllYearsNow to retry." & _
               vbCrLf & "Last successful sync: " & AyLastSyncText(), vbExclamation
    End If
    Exit Function

Failed:
    problem = Err.Description
    Resume Finish
End Function

' Warns once when the last successful sync is older than TenantAllYearsStaleDays().
Public Sub WarnIfAllYearsStale()
    Dim stamp As Variant
    On Error Resume Next
    stamp = ThisWorkbook.Worksheets(TenantSheetName("helper")).Range(TenantAllYearsSyncStampCell()).Value
    On Error GoTo 0
    If IsDate(stamp) Then
        If Now - CDate(stamp) <= TenantAllYearsStaleDays() Then Exit Sub
    End If
    MsgBox "The All-Years archive was last updated: " & AyLastSyncText() & "." & vbCrLf & vbCrLf & _
           "It updates automatically after this step. If this warning keeps appearing, run " & _
           "SyncAllYearsNow and read its message.", vbExclamation
End Sub

Private Function AyFindOpen(ByVal path As String) As Workbook
    Dim wb As Workbook
    Dim leaf As String
    leaf = Mid$(path, InStrRev(Replace(path, "\", "/"), "/") + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Or StrComp(wb.Name, leaf, vbTextCompare) = 0 Then
            Set AyFindOpen = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub AyStamp()
    Dim helper As Worksheet
    Set helper = ThisWorkbook.Worksheets(TenantSheetName("helper"))
    UnprotectSheetOn helper
    helper.Range(TenantAllYearsSyncStampCell()).Offset(0, -1).Value = "all-years sync"
    helper.Range(TenantAllYearsSyncStampCell()).Value = Now
    ProtectSheetOn helper
End Sub

Private Function AyLastSyncText() As String
    Dim stamp As Variant
    On Error Resume Next
    stamp = ThisWorkbook.Worksheets(TenantSheetName("helper")).Range(TenantAllYearsSyncStampCell()).Value
    On Error GoTo 0
    If IsDate(stamp) Then
        AyLastSyncText = Format$(CDate(stamp), "mm/dd/yyyy h:nn AM/PM")
    Else
        AyLastSyncText = "never"
    End If
End Function
```

- [ ] **Step 3: Wire up the call sites**

`AddNewBills.vb`: just after the `targetHeaderDict Is Nothing` guard (after line 19), add:

```vb
    WarnIfAllYearsStale
```

In the success tail from Task 6 Step 5, place the sync between `RestoreThinking` and the summary `MsgBox`:

```vb
            If Not importFailed Then SyncAllYears
```

`UpdateCoupaData.vb`: add `WarnIfAllYearsStale` on the line before `ChooseFile:` (line 8). After the `RestoreThinking` that follows `End Select` (line 89), add:

```vb
            ' Every Coupa branch changes columns the archive mirrors (orders too: ORDER DATE
            ' and PO STATUS are formulas over Coupa POs). Case Else never reaches this line.
            SyncAllYears
```

`Refresh.vb`: after `Call RestoreThinking` (line 63), add:

```vb
    ' Refresh rewrites REQ #, payment numbers and every formula column the archive mirrors.
    SyncAllYears
```

- [ ] **Step 4: Build the stack**

```powershell
cd C:\Users\p4bn\Documents\SecuritasAutomation
powershell -NoProfile -ExecutionPolicy Bypass -File .\Stack-VBFiles.ps1
```
Expected: exit 0.

- [ ] **Step 5: Verify by hand against copies (the user drives this)**

No path swap is needed, and none should be made: `TenantAllYearsWorkbookName()` is a bare file
name, and `SyncAllYears` derives the full path from `ThisWorkbook.Path`. So put the archive
copy in the SAME FOLDER as the scratch working-book copy from Task 6 -- that alone makes the
scratch working book resolve to the scratch archive, and makes it impossible for either copy to
reach the live shared archive by accident. Nothing in `TenantConfig.vb` needs editing or
restoring before or after this step.

Before checking anything else:
- [ ] the archive copy's tracker sheet is actually named `Invoices` (`TenantSheetName("tracker")`) --
      a rename here fails the header check silently (`MirrorInvoiceBlock` returns -1);
- [ ] `Helper!D6` on the scratch working book is empty -- it was empty on 2026-09-26, and
      `TenantAllYearsSyncStampCell()` (`E6`) sits next to it; a leftover value in D6 would be
      overwritten unnoticed by `AyStamp`'s label write to `E6`'s Offset(0, -1).
- [ ] before running this against anything that could become the LIVE sync, count `PAID` values
      in the archive's rows tagged `2026Model:Invoices` -- these rows are owned and get replaced
      wholesale on the first sync. If any are found, copy them into the working book first (see
      "Why PAID matters" in the spec); once replaced, that hand-investigation record is gone.

1. On the scratch copy of the working book from Task 6, paste the stack and compile.
2. Run `SyncAllYearsNow`. Expect "N row(s) mirrored", with N equal to the working book's data rows.
3. In the archive copy:
   - no `2026Model:Invoices` rows remain;
   - bill code `6005463095` has as many rows as in the working book (9 or more), not 1;
   - the legacy row count is unchanged;
   - at least one `PAID` value on a 2023/2024 row is still `PAID`.
4. Run `SyncAllYearsNow` again. The archive row count doesn't change.
5. Open the archive copy normally in a **second** Excel instance, so that instance holds the lock. Then run the sync from the first instance. Expect the "opened read-only" warning, and no change to the copy.
6. Clear `Helper!E6` and run `AddNewBills`. Expect the staleness warning first, then a sync at the end.
7. Run `UpdateCoupaData` with an orders export. Expect a silent sync, with `Helper!E6` updated.
8. Confirm the archive copy opened by the sync has AutoSave off (File > the AutoSave toggle), and that its Invoices sheet password matches TenantSheetPassword() ("Formulas").
9. Open the archive copy yourself, edit a cell without saving, then run SyncAllYearsNow. Expect the "open with unsaved changes" refusal and no change to the copy.

- [ ] **Step 6: Commit**

```bash
cd C:\Users\p4bn\Documents\SecuritasAutomation
git add SyncAllYears.vb TenantConfig.vb "file ingesting/AddNewBills.vb" "file ingesting/UpdateCoupaData.vb" Refresh.vb
git commit -m "feat: mirror Invoices into All-Years archive" -m "Whole-block, values-only mirror via core MirrorInvoiceBlock. Runs after
AddNewBills, every UpdateCoupaData branch and Refresh; warns when stale.
Verified against local copies: <what was driven>."
```

---

### Task 9: Docs and the Coordination Board

**Files:**
- Modify: `SecuritasAutomation/README.md` (Capability Doc)
- Modify: `JCI-invoice-tracker/README.md` if one exists (otherwise skip)
- Modify: `SecuritasAutomation/docs/core-extraction/tickets.md` (Coordination Board)
- Modify: `InvoiceTrackerCore/docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md` (status line)

- [ ] **Step 1: Update the Securitas Capability Doc**

Add a section to `SecuritasAutomation/README.md`:

```markdown
## Importing bills

`AddNewBills` refuses a file it has already processed. It recognises the file by content hash
(the hidden `Processed Batches` sheet), so a renamed or re-sent copy is recognised too. It
checks bill codes **before** writing:
- a code already in the tracker is skipped;
- a code repeated within the file is skipped on every row, **except** for monitoring bills,
  which put one bill code on several store rows.

Skipped rows are listed in the closing message, never written or marked red. A processed file
moves into a `Processed` subfolder beside it.

## All-Years archive

`Securitas All-Years Invoices - Consolidated.xlsm` receives a values-only mirror of this
workbook's `Invoices` sheet, tagged `2026:Invoices`, after every `AddNewBills`,
`UpdateCoupaData` and `Refresh`, and whenever you run `SyncAllYearsNow`. The mirrored rows are
read-only copies: edit here, not there. A failed sync warns and never blocks. `Helper!E6`
records the last success, and imports warn when it is more than 3 days old.
```

- [ ] **Step 2: Add tickets to the Coordination Board**

Append to `SecuritasAutomation/docs/core-extraction/tickets.md`:

```markdown
---

# Batch tracking and All-Years mirror (2026-09-26)

Spec: `InvoiceTrackerCore/docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md`.
Plan: `InvoiceTrackerCore/docs/superpowers/plans/2026-09-26-batch-workbook-tracking.md`.

## Batch tracking, bill-code screen, All-Years mirror — DONE

- Core: `HashFile`, `ProcessedBatchLog`, `MoveProcessedFile`, `BillCodeScreen`,
  `MirrorInvoiceBlock`, each with a harness test.
- Securitas: `AddNewBills` hash gate + pre-write screen (`CheckNewBillsForDuplicates` deleted),
  `SyncAllYears` after AddNewBills / UpdateCoupaData / Refresh.
- JCI: `AddNewBills` hash gate. `MarkDuplicateInvoices` unchanged.
- Verified: <record what the user drove in Tasks 6-8>.

## Record: the All-Years consolidation (2026-08-31 to 2026-09-03)

Never documented at the time. It built `Securitas All-Years Invoices - Consolidated.xlsm` from
eight sources:
- pre-2024 Oracle export;
- 2024;
- 2024 uninvoiced;
- Securitas Expense Tracker (Old Process);
- the undated root Repair_Installation file;
- 2025;
- SECURITAS OLD BILL CHECKER;
- the then-current 2026 workbook (tagged `2026Model:Invoices`).

It deduped to **one row per bill code, newest source wins**, and pasted values only. That key
**collapsed every monitoring bill to one store row**: `6005463095` has 9+ rows in the 2026
workbook and had 1 in the archive (verified 2026-09-26). The only other record of the
migration is `Excel-MCP/gaps-and-fixes.md`.

## Redo the All-Years consolidation — NOT STARTED

**Blocked by:** Batch tracking, bill-code screen, All-Years mirror.

- [ ] Rebuild non-2026 history from the eight sources, deduping on the whole bill
      (bill code + source), not on bill code alone, so monitoring store rows survive
- [ ] Carry across every hand-set `PAID` value, from the current archive and from the sources.
      Much of 2023-2024 was paid under different invoice numbers with no lookup trail; `PAID`
      is the only record of that manual investigation. Diff the `PAID` count between the old
      and new builds before replacing anything.
- [ ] Run one `SyncAllYearsNow` to lay the 2026 block on top

## JCI All-Years archive — NOT STARTED

- [ ] Inventory JCI's legacy sources and consolidate them into a JCI archive
- [ ] Move `SyncAllYears` into core and give JCI the `TenantAllYears*` accessors

## Stop WriteFormulas_* overwriting PAID cells — NOT STARTED

Requested during the consolidation session and still undone. Hand-set `PAID` replaces a lookup
formula for invoices paid under another number; a formula rewrite destroys that record.
```

- [ ] **Step 3: Mark the spec done and commit each repo**

In the spec, change the `Status:` line to `Status: implemented (see plan and tickets.md)`.

```bash
cd C:\Users\p4bn\Documents\SecuritasAutomation
git add README.md docs/core-extraction/tickets.md
git commit -m "docs: batch tracking, All-Years mirror, tickets"

cd C:\Users\p4bn\Documents\InvoiceTrackerCore
git add docs/superpowers/specs/2026-09-26-batch-workbook-tracking-design.md
git commit -m "docs(spec): mark batch tracking implemented"
```

- [ ] **Step 4: Ask the user before pushing**

List each repo's unpushed commits (`git log --oneline @{u}..`) and ask whether to push all three.
