<#
.SYNOPSIS
    Assembles a variant's VBA modules plus the shared InvoiceTrackerCore modules into a
    single stacked .vb file for import into an Excel workbook.

.DESCRIPTION
    VBA has no linker or package manager, so this script IS the module system.
    It concatenates .vb source files into one artifact.

    SHADOW RULE (see ADR-0004): when a variant file and a core file share a filename,
    the variant wins and the core file is excluded. This is how a variant overrides a
    core module. VBA has no namespaces, so including both would define the same Sub
    twice and fail to compile.

    Note that shadowing forks the module -- the shadow stops receiving core
    improvements. Prefer parameterizing through TenantConfig.vb instead.

.PARAMETER CorePath
    Path to the InvoiceTrackerCore checkout. Defaults to a sibling folder.

.PARAMETER OutputFile
    Name of the stacked output file, written into the variant root.

.PARAMETER SkipCore
    Stack variant modules only, ignoring core. Diagnostic use.
#>
[CmdletBinding()]
param(
    [string]$CorePath,
    [string]$OutputFile,
    [switch]$SkipCore
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is not available while binding parameter defaults, so resolve the
# default core location here instead.
if (-not $CorePath) {
    $CorePath = Join-Path $PSScriptRoot '..\InvoiceTrackerCore'
}

# Modules pinned to the front of the stack, in this order.
# Header declares globals; Refresh is the primary entry point.
$priorityFiles = @('Header.vb', 'Refresh.vb')

function Get-StackableFiles {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$OutputLeaf
    )

    Get-ChildItem -Path $Root -Recurse -File -Filter '*.vb' |
        Where-Object {
            $_.Name -ne $OutputLeaf -and
            $_.Name -notlike 'DEPRECATED*' -and
            $_.FullName -notlike '*\UserForm\*' -and
            $_.FullName -notlike '*\node_modules\*'
        }
}

# --- Resolve the variant root and output name ------------------------------------

$variantRoot = $PSScriptRoot

if (-not $OutputFile) {
    # Reuse the existing stacked artifact's name if one is already present, so the
    # workbook's import target does not change.
    $existing = Get-ChildItem -Path $variantRoot -File -Filter '*MegaStack*.vb' |
        Select-Object -First 1
    if ($existing) {
        $OutputFile = $existing.Name
    }
    else {
        $OutputFile = '{0}_MegaStack.vb' -f (Split-Path $variantRoot -Leaf)
    }
}

$outputPath = Join-Path $variantRoot $OutputFile
$outputLeaf = Split-Path $outputPath -Leaf

# --- Gather variant modules -------------------------------------------------------

$variantFiles = @(Get-StackableFiles -Root $variantRoot -OutputLeaf $outputLeaf)

# --- Gather core modules, applying the shadow rule --------------------------------

$coreFiles = @()
$shadowed = @()

if (-not $SkipCore) {
    if (-not (Test-Path -LiteralPath $CorePath)) {
        throw @"
Core path not found: $CorePath

InvoiceTrackerCore must be checked out for this variant to build.
Clone it as a sibling folder:

    git clone https://github.com/timjavins/InvoiceTrackerCore.git

...or pass an explicit path:

    .\$($MyInvocation.MyCommand.Name) -CorePath <path-to-InvoiceTrackerCore>

To stack variant modules only (diagnostic), pass -SkipCore.
"@
    }

    $resolvedCore = (Resolve-Path -LiteralPath $CorePath).Path
    $variantNames = @($variantFiles | ForEach-Object { $_.Name })

    foreach ($file in Get-StackableFiles -Root $resolvedCore -OutputLeaf $outputLeaf) {
        if ($variantNames -contains $file.Name) {
            $shadowed += $file.Name
        }
        else {
            $coreFiles += $file
        }
    }
}

# --- Order: priority files first, then everything else alphabetically -------------

$allFiles = @($variantFiles) + @($coreFiles)

$orderedFiles = @()
foreach ($priorityFile in $priorityFiles) {
    $orderedFiles += $allFiles | Where-Object { $_.Name -ieq $priorityFile }
}
$orderedFiles += $allFiles |
    Where-Object { $priorityFiles -notcontains $_.Name } |
    Sort-Object Name

# --- Guard: duplicate filenames would mean duplicate definitions -----------------

$duplicates = $orderedFiles | Group-Object Name | Where-Object { $_.Count -gt 1 }
if ($duplicates) {
    $detail = ($duplicates | ForEach-Object {
        "  {0}`n{1}" -f $_.Name, (($_.Group | ForEach-Object { "      $($_.FullName)" }) -join "`n")
    }) -join "`n"

    throw @"
Duplicate module filenames in the stack -- this would fail to compile in VBA:

$detail

Each module name must be unique across the variant and core combined.
"@
}

# --- Warn: a local variable that shadows a procedure ------------------------------

# VBA identifiers are case-insensitive, so `Dim tierLabels` and `Private Function TierLabels()`
# are the SAME name. The Dim wins for the whole procedure it appears in, and calling the
# function from there compiles cleanly -- `tierLabels()` is a valid array access -- then fails
# at run time with "Subscript out of range".
#
# This exact bug shipped into a live workbook and cost a full debugging session, because the
# symptom (error 9 from a flat-file generator) points nowhere near the cause. Only the
# assembler sees every module at once, so only the assembler can catch the next one.
#
# A warning rather than an error: shadowing is not always a mistake, and a hard failure here
# could block a build over a local that never calls its namesake.

$procRegex = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'
$endRegex = '^\s*End\s+(?:Sub|Function|Property)\b'
$dimRegex = '^\s*(?:Dim|Static)\s+([\w\s,]+?)(?:\s+As\s|\s*$)'

$procNames = @{}
foreach ($file in $orderedFiles) {
    foreach ($line in (Get-Content -LiteralPath $file.FullName -Encoding UTF8)) {
        if ($line -match $procRegex) { $procNames[$matches[1].ToLower()] = $matches[1] }
    }
}

$shadows = @()
foreach ($file in $orderedFiles) {
    $currentProc = ''
    $inProcedure = $false
    $lineNo = 0
    foreach ($line in (Get-Content -LiteralPath $file.FullName -Encoding UTF8)) {
        $lineNo++
        if (-not $inProcedure -and $line -match $procRegex) {
            $currentProc = $matches[1]
            $inProcedure = $true
            continue
        }
        if ($inProcedure -and $line -match $endRegex) {
            $inProcedure = $false
            $currentProc = ''
            continue
        }
        if ($inProcedure -and $line -match $dimRegex) {
            foreach ($declared in ($matches[1] -split ',')) {
                $name = ($declared.Trim() -split '\s+')[0]
                if (-not $name) { continue }
                $key = $name.ToLower()
                if ($procNames.ContainsKey($key) -and $key -ne $currentProc.ToLower()) {
                    $shadows += [pscustomobject]@{
                        File     = $file.Name
                        Line     = $lineNo
                        Proc     = $currentProc
                        Local    = $name
                        Shadowed = $procNames[$key]
                    }
                }
            }
        }
    }
}

# --- Split each module into declarations and procedures ---------------------------

# VBA requires every module-level declaration to appear before the first procedure:
# anything else raises "Only comments may appear after End Sub, End Function, or End
# Property". Since the stack concatenates many modules, a module that legitimately declares
# its own state lands mid-file and breaks the build.
#
# So hoist declarations instead of asking every module author to remember the rule. Each
# module's module-level lines are collected into a header block, its procedures follow, and
# the two are emitted in the required order.
function Split-VbModule {
    param([Parameter(Mandatory)][string]$Text)

    $decls = New-Object System.Collections.Generic.List[string]
    $procs = New-Object System.Collections.Generic.List[string]

    $inProc = $false
    $seenProc = $false

    foreach ($line in ($Text -split "`r?`n")) {
        $trimmed = $line.Trim()

        if (-not $inProc -and $trimmed -match '^(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+\w') {
            $inProc = $true
            $seenProc = $true
            $procs.Add($line)
            continue
        }

        if ($inProc) {
            $procs.Add($line)
            if ($trimmed -match '^End\s+(Sub|Function|Property)\b') { $inProc = $false }
            continue
        }

        # Outside a procedure.
        if ($trimmed -match '^(Option\s|Public\s|Private\s|Friend\s|Dim\s|Const\s|Declare\s|Type\s|Enum\s|Implements\s)') {
            # A declaration. Carry any comment lines immediately above it along, so the
            # explanation travels with the thing it explains.
            $lead = New-Object System.Collections.Generic.List[string]
            while ($procs.Count -gt 0 -and $procs[$procs.Count - 1].Trim() -match '^(''|$)') {
                $lead.Insert(0, $procs[$procs.Count - 1])
                $procs.RemoveAt($procs.Count - 1)
            }
            foreach ($c in $lead) { $decls.Add($c) }
            $decls.Add($line)
        }
        else {
            # Comments and blanks. Leave them where they are -- a module's preamble belongs
            # with its procedures, not hoisted into the header block.
            $procs.Add($line)
        }
    }

    [pscustomobject]@{
        Declarations = ($decls -join "`r`n").Trim()
        Procedures   = ($procs -join "`r`n").Trim()
    }
}

$declBlocks = New-Object System.Collections.Generic.List[string]
$procBlocks = New-Object System.Collections.Generic.List[string]
$hoistedFrom = New-Object System.Collections.Generic.List[string]

foreach ($file in $orderedFiles) {
    $raw = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    $part = Split-VbModule -Text $raw

    if ($part.Declarations) {
        $declBlocks.Add(("' ===== from {0} =====`r`n{1}" -f $file.Name, $part.Declarations))

        # Report only modules whose declarations had to move, i.e. not the first module.
        if ($procBlocks.Count -gt 0) { $hoistedFrom.Add($file.Name) }
    }
    if ($part.Procedures) {
        $procBlocks.Add($part.Procedures)
    }
}

# --- Write ------------------------------------------------------------------------

$merged = (@($declBlocks) + @($procBlocks)) -join "`r`n`r`n"

# Encoding matters in both directions, and getting either wrong corrupts comments silently.
#
# Read: every Get-Content above passes -Encoding UTF8. Windows PowerShell 5.1 otherwise decodes a
# BOM-less file as the system code page, so a UTF-8 em-dash arrived as three cp1252 characters and
# was re-encoded on write -- once per stack, compounding. Sources are UTF-8 without BOM.
#
# Write: UTF8Encoding($false), not [Text.Encoding]::UTF8, which emits a BOM. The VBE shows a BOM as
# a stray character on the first line of the pasted module.
[System.IO.File]::WriteAllText($outputPath, $merged, (New-Object System.Text.UTF8Encoding($false)))

# --- Report -----------------------------------------------------------------------

Write-Host ""
Write-Host "Wrote $outputPath"
Write-Host ""
Write-Host ("  variant modules : {0}" -f $variantFiles.Count)
if ($SkipCore) {
    Write-Host "  core modules    : (skipped)"
}
else {
    Write-Host ("  core modules    : {0}" -f $coreFiles.Count)
}
Write-Host ("  total stacked   : {0}" -f $orderedFiles.Count)

if ($hoistedFrom.Count -gt 0) {
    Write-Host ""
    Write-Host ("  Hoisted module-level declarations from {0} module(s):" -f $hoistedFrom.Count)
    foreach ($name in ($hoistedFrom | Sort-Object -Unique)) {
        Write-Host "    $name"
    }
    Write-Host "  VBA requires these before the first procedure, so they move to the top."
}

if ($shadows.Count -gt 0) {
    Write-Host ""
    Write-Host ("  WARNING -- {0} local variable(s) shadow a procedure of the same name:" -f $shadows.Count)
    foreach ($s in $shadows) {
        Write-Host ("    {0}:{1}  in {2}()  local '{3}' shadows '{4}'" -f `
            $s.File, $s.Line, $s.Proc, $s.Local, $s.Shadowed)
    }
    Write-Host "  VBA names are case-insensitive, so the Dim wins for that whole procedure."
    Write-Host "  Calling the procedure from there compiles, then fails at run time with"
    Write-Host "  'Subscript out of range'. Rename one of them."
}

if ($shadowed.Count -gt 0) {
    Write-Host ""
    Write-Host ("  Shadowed core modules ({0}) -- variant copy wins:" -f $shadowed.Count)
    foreach ($name in ($shadowed | Sort-Object)) {
        Write-Host "    $name"
    }
    Write-Host ""
    Write-Host "  A shadow forks the module: it will not receive core improvements."
}

Write-Host ""
Write-Host "Code in any UserForm folder is intentionally excluded."
Write-Host ""
