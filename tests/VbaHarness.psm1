# Runs VBA from InvoiceTrackerCore against a throwaway workbook in its own hidden
# Excel instance, so tests never touch a workbook the user has open.
#
# Requires Excel's "Trust access to the VBA project object model" (VBOM). Without it
# $wb.VBProject throws, and the message Excel gives is unhelpful, so we catch and explain.
#
# A hidden instance is not a SAFE instance. CodeModule.AddFromString (below) does not
# validate VBA syntax or resolve names, so bad injected code fails later, at the first call
# into it -- and an unhandled VBA compile or runtime error under automation opens a MODAL
# dialog. Application.Visible = $false does NOT suppress that dialog; it renders on screen
# regardless of instance visibility, and the blocked COM call never returns on its own. That
# is not theoretical: it happened in the live operator's own Excel session, who had to press
# OK by hand and abort the run, and a wall-clock timeout around the whole script did not
# help, because the human reached the dialog first and a blocked COM call does not return.
#
# So the harness cannot prevent the dialog, but it can make it die on its own. New-VbaHost
# captures the PID of the Excel process IT just created (see below); Invoke-VbaFunction's
# -TimeoutSeconds and Remove-VbaHost's post-Quit poll both use that captured PID, and ONLY
# that PID, to kill a stuck instance instead of leaving it for a human. This is
# safety-critical: the operator has real Excel sessions of their own open, including one
# holding a live, unsaved, cloud-hosted workbook, plus an unrelated pre-existing headless
# Excel process. Killing by process NAME, or "all EXCEL processes", could destroy either.
# Every kill path below is guarded to no-op when the captured PID is unknown -- a harness
# that cannot prove which process is its own child must not guess.

Set-StrictMode -Version Latest

function New-VbaHost {
    param(
        [Parameter(Mandatory)][string[]] $SourceFiles,

        # Injects into ThisWorkbook (a document/class module) instead of adding a standard
        # module. A standard module lets Application.Run resolve an unqualified name, which is
        # why Invoke-VbaFunction and the 643-assertion suite depend on the default (unset)
        # path. ThisWorkbook always exists on a workbook and forbids what a standard module
        # allows (Public Const, public fixed-size arrays, fixed-length strings, Declare), so it
        # must be looked up rather than Added -- and its procedures are reachable only as COM
        # methods on the workbook object ($wb.Func(...)), not through Application.Run.
        [switch] $DocumentModule
    )

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
    $excelPid = $null

    try {
        # New-Object -ComObject creates a SEPARATE instance. Never use GetActiveObject here:
        # that would attach to the user's Excel and run test code beside live workbooks.
        # This whole block -- construction, property assignment, workbook creation, VBProject
        # access, module injection -- is inside one try so nothing between "Excel exists" and
        # "the host is fully built" can leak a hidden process on failure.
        #
        # New-Object does not hand back a PID, and $excel.Hwnd is not a safe substitute for
        # one -- a hidden instance can lack a usable window handle. So capture it by process-ID
        # set difference instead: snapshot every running EXCEL pid immediately before
        # construction, snapshot again immediately after, and the one new id is ours. If that
        # is not exactly one new id -- none appeared yet, or more than one Excel process
        # started at the same moment -- there is no safe way to tell which is ours, so store
        # nothing. Every kill path in this module (Invoke-VbaFunction -TimeoutSeconds,
        # Remove-VbaHost) treats a missing pid as "do not act" rather than guessing; a wrong
        # guess here could kill the operator's own Excel, including a live cloud-hosted
        # workbook with unsaved changes.
        $pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Id)

        $excel = New-Object -ComObject Excel.Application

        $pidsAfter = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Id)
        $newPids = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
        if ($newPids.Count -eq 1) { $excelPid = $newPids[0] }

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

        if ($DocumentModule) {
            $module = $project.VBComponents('ThisWorkbook')
        } else {
            # 1 = vbext_ct_StdModule. A standard module (not ThisWorkbook) so Application.Run
            # resolves unqualified names -- a class module would require 'ThisWorkbook.Proc'.
            $module = $project.VBComponents.Add(1)
        }

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

    @{ Excel = $excel; Workbook = $wb; DocumentModule = [bool]$DocumentModule; ExcelProcessId = $excelPid }
}

function Invoke-VbaFunction {
    param(
        [Parameter(Mandatory)][hashtable] $VbaHost,
        [Parameter(Mandatory)][string] $Name,
        [object[]] $Arguments = @(),

        # 0 (default) preserves every existing caller's behaviour exactly: wait for
        # Application.Run indefinitely, same as before this parameter existed.
        #
        # A positive value arms a watchdog for this one call. If $Name has not returned
        # within $TimeoutSeconds, the watchdog kills ONLY the PID New-VbaHost captured for
        # THIS host -- never by process name, never "all EXCEL" -- and this function throws a
        # specific error naming $Name instead of leaving a hung COM call and a modal dialog
        # for a human to find (see the module header for why Visible = $false does not help).
        #
        # The watchdog runs as a separate background job (its own process), not a thread
        # inside this one. That is deliberate: $VbaHost.Excel is an STA COM proxy, and handing
        # that same proxy to another thread to poke at is its own hazard. A background job
        # needs nothing but the plain integer pid, so it can only ever act on the one pid it
        # was given -- it never touches the COM object at all.
        #
        # If New-VbaHost could not identify its own pid, this timeout is inert: the watchdog
        # has nothing safe to kill, so a blocked call still hangs, exactly as it would with
        # $TimeoutSeconds = 0. Never guess a pid to make the timeout "work".
        [int] $TimeoutSeconds = 0
    )

    # A document-module host's procedures are members of the workbook COM object, not
    # names Application.Run can resolve: unqualified, Run cannot see into ThisWorkbook at
    # all, and even qualified as 'ThisWorkbook.Proc' it executes the procedure but discards
    # a Function's return value -- so this would either error opaquely or silently return
    # $null. Callers against a -DocumentModule host must call the workbook object directly,
    # e.g. $VbaHost.Workbook.FiscalYearWeeks(2026).
    if ($VbaHost.DocumentModule) {
        throw "Invoke-VbaFunction cannot call '$Name' on a -DocumentModule host: " +
              "Application.Run cannot resolve an unqualified name in a document module, and " +
              "even qualified it discards a Function's return value. Call it as a COM method " +
              "on the workbook object instead, e.g. `$VbaHost.Workbook.$Name(...)."
    }

    if ($Arguments.Count -gt 2) {
        throw "Invoke-VbaFunction supports up to 2 arguments; got $($Arguments.Count)."
    }

    if ($TimeoutSeconds -le 0) {
        switch ($Arguments.Count) {
            0 { return $VbaHost.Excel.Run($Name) }
            1 { return $VbaHost.Excel.Run($Name, $Arguments[0]) }
            2 { return $VbaHost.Excel.Run($Name, $Arguments[0], $Arguments[1]) }
        }
    }

    $targetPid = $VbaHost.ExcelProcessId
    $watchdog = $null
    if ($targetPid) {
        $watchdog = Start-Job -ScriptBlock {
            param($Seconds, $TargetPid)
            Start-Sleep -Seconds $Seconds
            # Guarded a second time here, inside the job itself: this scriptblock closes over
            # nothing but the two plain values it was handed, so there is no path by which it
            # could act on any pid other than the one $VbaHost.ExcelProcessId captured.
            if ($TargetPid) { Stop-Process -Id $TargetPid -Force -ErrorAction SilentlyContinue }
        } -ArgumentList $TimeoutSeconds, $targetPid
    }

    $result = $null
    $callError = $null
    try {
        $result = switch ($Arguments.Count) {
            0 { $VbaHost.Excel.Run($Name) }
            1 { $VbaHost.Excel.Run($Name, $Arguments[0]) }
            2 { $VbaHost.Excel.Run($Name, $Arguments[0], $Arguments[1]) }
        }
    } catch {
        $callError = $_
    } finally {
        # Whether the call returned, threw, or is still technically "returning" right as the
        # process dies underneath it, the job's own state tells us whether IT already ran the
        # kill (State 'Completed' means the sleep elapsed and Stop-Process was invoked) versus
        # still waiting (State 'Running', harmlessly cancelled below).
        $fired = $false
        if ($watchdog) {
            $fired = ($watchdog.State -eq 'Completed')
            Stop-Job $watchdog -ErrorAction SilentlyContinue
            Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
        }
    }

    if ($fired) {
        throw "Invoke-VbaFunction: '$Name' did not return within ${TimeoutSeconds}s, so the " +
              "harness killed its own Excel process (PID $targetPid). A modal VBA compile or " +
              "runtime dialog is the likely cause -- Application.Visible = `$false does not " +
              "suppress it (see this module's header)."
    }
    if ($callError) { throw $callError }
    $result
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

    # Quit() is asynchronous: it returns before the process has actually exited, and cleanup
    # today has already once left a stray EXCEL.EXE that had to be found and killed by hand.
    # So poll briefly (~2s is the figure established elsewhere in this project) for OUR
    # captured pid -- never any other EXCEL process -- to disappear on its own, and only kill
    # it if it is still there once that grace period elapses. A missing pid means there is
    # nothing safe to act on, so this is a no-op in that case, same as everywhere else.
    $targetPid = $VbaHost.ExcelProcessId
    if ($targetPid) {
        $deadline = (Get-Date).AddSeconds(2)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $targetPid -ErrorAction SilentlyContinue)) {
            Start-Sleep -Milliseconds 200
        }
        if (Get-Process -Id $targetPid -ErrorAction SilentlyContinue) {
            try { Stop-Process -Id $targetPid -Force -ErrorAction SilentlyContinue } catch { }
        }
    }

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
