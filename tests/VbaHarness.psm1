# Runs VBA from InvoiceTrackerCore against a throwaway workbook in its own hidden
# Excel instance, so tests never touch a workbook the user has open.
#
# Requires Excel's "Trust access to the VBA project object model" (VBOM). Without it
# $wb.VBProject throws, and the message Excel gives is unhelpful, so we catch and explain.

Set-StrictMode -Version Latest

function New-VbaHost {
    param([Parameter(Mandatory)][string[]] $SourceFiles)

    # Validate every source path BEFORE any COM object exists. Nothing can leak a hidden
    # Excel process if nothing was created yet -- and this validation is the one most likely
    # to fail (a typo'd path, a task run before its .vb file exists), so it must not need
    # cleanup at all. Do not move this below New-Object "for tidiness": that reintroduces the
    # leak this check exists to prevent.
    foreach ($file in $SourceFiles) {
        if (-not (Test-Path -LiteralPath $file)) { throw "Source file not found: $file" }
    }

    # Initialise both to $null before the try so the catch's cleanup can safely test for
    # either one -- under Set-StrictMode, referencing an unset variable throws, which would
    # otherwise mask the real error with an unrelated "variable has not been set" failure.
    $excel = $null
    $wb = $null

    try {
        # New-Object -ComObject creates a SEPARATE instance. Never use GetActiveObject here:
        # that would attach to the user's Excel and run test code beside live workbooks.
        # This whole block -- construction, property assignment, workbook creation, VBProject
        # access, module injection -- is inside one try so nothing between "Excel exists" and
        # "the host is fully built" can leak a hidden process on failure.
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Add()

        try {
            $project = $wb.VBProject
        } catch {
            throw "Cannot reach the VBA project. Enable Excel > File > Options > Trust Center > " +
                  "Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', " +
                  "then re-run. Excel said: $($_.Exception.Message)"
        }

        # 1 = vbext_ct_StdModule. A standard module (not ThisWorkbook) so Application.Run
        # resolves unqualified names -- a class module would require 'ThisWorkbook.Proc'.
        $module = $project.VBComponents.Add(1)

        foreach ($file in $SourceFiles) {
            $source = Get-Content -LiteralPath $file -Raw -Encoding UTF8
            $module.CodeModule.AddFromString($source)
        }
    } catch {
        if ($excel) { try { $excel.Quit() } catch { } }
        foreach ($obj in $wb, $excel) {
            if ($obj) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { } }
        }
        throw
    }

    @{ Excel = $excel; Workbook = $wb }
}

function Invoke-VbaFunction {
    param(
        [Parameter(Mandatory)][hashtable] $VbaHost,
        [Parameter(Mandatory)][string] $Name,
        [object[]] $Arguments = @()
    )
    switch ($Arguments.Count) {
        0 { $VbaHost.Excel.Run($Name) }
        1 { $VbaHost.Excel.Run($Name, $Arguments[0]) }
        2 { $VbaHost.Excel.Run($Name, $Arguments[0], $Arguments[1]) }
        default { throw "Invoke-VbaFunction supports up to 2 arguments; got $($Arguments.Count)." }
    }
}

function Remove-VbaHost {
    param([Parameter(Mandatory)][hashtable] $VbaHost)

    try { $VbaHost.Workbook.Close($false) } catch { }
    try { $VbaHost.Excel.Quit() } catch { }
    foreach ($key in 'Workbook', 'Excel') {
        if ($VbaHost[$key]) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($VbaHost[$key]) } catch { }
        }
    }
    [GC]::Collect()
    $null
}

function Import-FiscalFixture {
    param([Parameter(Mandatory)][string] $Path)

    # Import-Csv cannot skip comment lines, so strip them before parsing.
    Get-Content -LiteralPath $Path |
        Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } |
        ConvertFrom-Csv |
        ForEach-Object {
            [pscustomobject]@{
                fiscal_year = [int]    $_.fiscal_year
                week1_start = [datetime]::ParseExact($_.week1_start, 'yyyy-MM-dd', $null)
                year_end    = [datetime]::ParseExact($_.year_end,    'yyyy-MM-dd', $null)
                weeks       = [int]    $_.weeks
            }
        }
}

function Reset-AssertCounters {
    $script:Passed = 0
    $script:Failed = 0
    $script:Failures = New-Object System.Collections.Generic.List[string]
}

function Assert-Equal {
    param($Expected, $Actual, [Parameter(Mandatory)][string] $Because)

    # Normalise dates to yyyy-MM-dd so a COM DateTime with a zero time component
    # compares equal to a fixture date.
    if ($Expected -is [datetime]) { $Expected = $Expected.ToString('yyyy-MM-dd') }
    if ($Actual   -is [datetime]) { $Actual   = $Actual.ToString('yyyy-MM-dd') }

    if ("$Expected" -eq "$Actual") {
        $script:Passed++
    } else {
        $script:Failed++
        $script:Failures.Add("FAIL $Because -- expected '$Expected', got '$Actual'")
    }
}

function Write-AssertSummary {
    foreach ($f in $script:Failures) { Write-Host $f -ForegroundColor Red }
    Write-Host ""
    Write-Host ("passed: {0}  failed: {1}" -f $script:Passed, $script:Failed)
    if ($script:Failed -gt 0) { exit 1 }
    Write-Host "ALL PASS" -ForegroundColor Green
    exit 0
}

Export-ModuleMember -Function New-VbaHost, Invoke-VbaFunction, Remove-VbaHost,
                              Import-FiscalFixture, Reset-AssertCounters, Assert-Equal,
                              Write-AssertSummary
