<#
.SYNOPSIS
    Batch-tests all CMPackager recipe XML files in a folder using Test-RecipeInstallation.ps1.

.DESCRIPTION
    Iterates over every non-template recipe XML in the target folder, runs
    Test-RecipeInstallation.ps1 for each one, and collects the outcome.
    Produces a detailed console report and a timestamped CSV file suitable
    for archival and compliance purposes.

    Each test gets its own workspace folder so logs are preserved for
    review regardless of the outcome.

    Installers are downloaded automatically from each recipe's <URL> or
    <PrefetchScript> and deleted after testing unless -KeepInstallers is set.

.PARAMETER RecipesPath
    Folder containing recipe XML files to test.
    Defaults to the Disabled\ folder two levels above this script.

.PARAMETER OutputPath
    Directory where the results CSV is written.
    Defaults to the current working directory.

.PARAMETER TimeoutMinutes
    Per-test timeout passed to Test-RecipeInstallation.ps1.
    Defaults to 30.

.PARAMETER KeepInstallers
    When set, downloaded installers are not deleted after each test.

.PARAMETER Filter
    Wildcard pattern to restrict which recipe files are tested.
    E.g. -Filter 'Mozilla*' or -Filter '7-Zip.xml'.
    Defaults to '*.xml' (all recipes).

.EXAMPLE
    .\Test-RecipeBatch.ps1

.EXAMPLE
    .\Test-RecipeBatch.ps1 -RecipesPath ..\..\Recipes -Filter 'Adobe*' -TimeoutMinutes 45

.NOTES
    Requires Windows Sandbox and PowerShell 7+.
    Must be run on a Windows host.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$RecipesPath = (Join-Path $PSScriptRoot '..\..' 'Disabled'),

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [int]$TimeoutMinutes = 30,

    [Parameter(Mandatory = $false)]
    [switch]$KeepInstallers,

    [Parameter(Mandatory = $false)]
    [string]$Filter = '*.xml'
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ── Resolve paths ─────────────────────────────────────────────────────────────

$testScript = Join-Path $PSScriptRoot 'Test-RecipeInstallation.ps1'
if (-not (Test-Path $testScript)) {
    Write-Error "Test-RecipeInstallation.ps1 not found at: $testScript"
    exit 1
}

$RecipesPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $RecipesPath))
if (-not (Test-Path $RecipesPath -PathType Container)) {
    Write-Error "Recipes folder not found: $RecipesPath"
    exit 1
}

if (-not (Test-Path $OutputPath -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# ── Discover recipes ───────────────────────────────────────────────────────────

$recipeFiles = Get-ChildItem -Path $RecipesPath -Filter $Filter |
    Where-Object { $_.Name -notlike '_*' } |
    Sort-Object Name

$total = $recipeFiles.Count
if ($total -eq 0) {
    Write-Host "No recipes matched '$Filter' in: $RecipesPath" -ForegroundColor Yellow
    exit 0
}

# ── Setup ─────────────────────────────────────────────────────────────────────

$sessionTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$batchWorkspaceRoot = Join-Path $env:TEMP 'CMPackagerBatchTest'
$csvPath = Join-Path $OutputPath "RecipeTestResults_$sessionTimestamp.csv"

Write-Host ''
Write-Host '  CMPackager Recipe Batch Tester' -ForegroundColor White
Write-Host '══════════════════════════════════════════════════════════' -ForegroundColor White
Write-Host "  Recipes folder : $RecipesPath" -ForegroundColor Gray
Write-Host "  Recipes found  : $total" -ForegroundColor Gray
Write-Host "  Timeout/test   : $TimeoutMinutes min" -ForegroundColor Gray
Write-Host "  Results CSV    : $csvPath" -ForegroundColor Gray
Write-Host '══════════════════════════════════════════════════════════' -ForegroundColor White
Write-Host ''

$allResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$i = 0

foreach ($recipeFile in $recipeFiles) {
    $i++

    # ── Parse recipe for display ───────────────────────────────────────────────

    $appName     = $recipeFile.BaseName
    $depTypeName = ''
    $xml         = $null
    try {
        [xml]$xml = Get-Content $recipeFile.FullName -Raw
        if ($xml.ApplicationDef.Application.Name) {
            $appName = $xml.ApplicationDef.Application.Name
        }
        $firstDep = $xml.ApplicationDef.DeploymentTypes.DeploymentType | Select-Object -First 1
        if ($firstDep) { $depTypeName = $firstDep.Name }
    } catch {
        # will be marked ERROR below
    }

    # ── Banner ─────────────────────────────────────────────────────────────────

    $indexLabel = "[$i/$total]"
    Write-Host ''
    Write-Host '──────────────────────────────────────────────────────────' -ForegroundColor DarkGray
    Write-Host "  $indexLabel $appName" -ForegroundColor Cyan
    Write-Host "  Recipe: $($recipeFile.Name)" -ForegroundColor DarkGray
    Write-Host ''


    # ── Per-test workspace ─────────────────────────────────────────────────────

    $safeName = $recipeFile.BaseName -replace '[^\w\-]', '_'
    $workspace = Join-Path $batchWorkspaceRoot $safeName

    # ── Run test ───────────────────────────────────────────────────────────────

    $testStart  = Get-Date
    $extraNotes = ''

    $testArgs = @{
        RecipePath     = $recipeFile.FullName
        WorkspacePath  = $workspace
        TimeoutMinutes = $TimeoutMinutes
    }
    if (-not $KeepInstallers) { $testArgs['CleanupInstaller'] = $true }

    # Wait for any lingering sandbox from the previous test to fully exit before
    # launching the next one (only one instance of Windows Sandbox is allowed).
    # On Windows 11 24H2+ wsb.exe is used for a precise check; on older systems
    # we fall back to process names.
    $wsbCliPath     = "$env:SystemRoot\System32\wsb.exe"
    $useWsbCliLocal = Test-Path $wsbCliPath
    $sbWaitDeadline = (Get-Date).AddMinutes(3)
    $sbWaitLogged   = $false
    while ((Get-Date) -lt $sbWaitDeadline) {
        $sandboxRunning = if ($useWsbCliLocal) {
            $listOut = & $wsbCliPath list 2>&1
            [bool]($listOut | Select-String -Pattern '[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}')
        } else {
            [bool](Get-Process -Name 'WindowsSandbox', 'WindowsSandboxClient' -ErrorAction SilentlyContinue)
        }
        if (-not $sandboxRunning) { break }
        if (-not $sbWaitLogged) {
            Write-Host '  Waiting for previous sandbox to shut down...' -ForegroundColor DarkYellow
            $sbWaitLogged = $true
        }
        Start-Sleep -Seconds 3
    }
    if ($sbWaitLogged) { Write-Host '' }

    try {
        & $testScript @testArgs
    } catch {
        $extraNotes = "Unhandled exception: $_"
    }

    $durationMin = [math]::Round(((Get-Date) - $testStart).TotalMinutes, 1)

    # ── Determine outcome from workspace artefacts ─────────────────────────────

    $resultsJson = Join-Path $workspace 'results.json'
    $sandboxLog  = Join-Path $workspace 'sandbox.log'

    $outcome                = 'ERROR'
    $notes                  = $extraNotes
    $installExitCode        = $null
    $installSuccess         = $null
    $detectionAfterInstall  = $null
    $uninstallExitCode      = $null
    $uninstallSuccess       = $null
    $detectionAfterUninstall = $null

    if (Test-Path $resultsJson) {
        try {
            $r = Get-Content $resultsJson -Raw | ConvertFrom-Json
            $outcome                 = $r.OverallResult
            $depTypeName             = if ($r.DeploymentType) { $r.DeploymentType } else { $depTypeName }
            $installExitCode         = $r.InstallExitCode
            $installSuccess          = $r.InstallSuccess
            $detectionAfterInstall   = $r.DetectionAfterInstall?.Detected
            $uninstallExitCode       = $r.UninstallExitCode
            $uninstallSuccess        = $r.UninstallSuccess
            $detectionAfterUninstall = $r.DetectionAfterUninstall?.Detected
            if ($r.Notes -and $r.Notes.Count -gt 0) {
                $notes = ($r.Notes -join '; ')
            }
        } catch {
            $outcome = 'ERROR'
            $notes   = "Failed to parse results.json: $_"
        }
    } elseif (Test-Path $sandboxLog) {
        $outcome = 'TIMEOUT'
        $notes   = "Sandbox started but no results produced within $TimeoutMinutes minutes"
    } else {
        $outcome = 'SKIPPED'
        $notes   = if ($notes) { $notes } else { 'Sandbox did not start — recipe may lack a downloadable URL' }
    }

    # ── Record result ──────────────────────────────────────────────────────────

    $row = [PSCustomObject]@{
        Timestamp               = $testStart.ToString('yyyy-MM-dd HH:mm:ss')
        Application             = $appName
        RecipeFile              = $recipeFile.Name
        DeploymentType          = $depTypeName
        Result                  = $outcome
        DurationMinutes         = $durationMin
        InstallExitCode         = $installExitCode
        InstallSuccess          = $installSuccess
        DetectionAfterInstall   = $detectionAfterInstall
        UninstallExitCode       = $uninstallExitCode
        UninstallSuccess        = $uninstallSuccess
        DetectionAfterUninstall = $detectionAfterUninstall
        Notes                   = $notes
        WorkspacePath           = $workspace
    }
    $allResults.Add($row)

    # Write CSV after every test so partial results survive an interrupted run
    $allResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}

# ── Final report ───────────────────────────────────────────────────────────────

$counts = @{
    Pass    = ($allResults | Where-Object { $_.Result -like 'PASS*' }).Count
    Fail    = ($allResults | Where-Object { $_.Result -eq 'FAIL' }).Count
    Timeout = ($allResults | Where-Object { $_.Result -eq 'TIMEOUT' }).Count
    Skipped = ($allResults | Where-Object { $_.Result -eq 'SKIPPED' }).Count
    Error   = ($allResults | Where-Object { $_.Result -eq 'ERROR' }).Count
}

Write-Host ''
Write-Host '══════════════════════════════════════════════════════════' -ForegroundColor White
Write-Host '  BATCH TEST REPORT' -ForegroundColor White
Write-Host '══════════════════════════════════════════════════════════' -ForegroundColor White
Write-Host ''

foreach ($r in $allResults) {
    $resultColor = switch -Wildcard ($r.Result) {
        'PASS*'   { 'Green'      }
        'FAIL'    { 'Red'        }
        'TIMEOUT' { 'Yellow'     }
        'SKIPPED' { 'DarkYellow' }
        default   { 'Gray'       }
    }

    $appCol      = $r.Application.PadRight(36)
    $resultCol   = $r.Result.PadRight(10)
    $durationCol = "$($r.DurationMinutes) min"

    Write-Host "  $appCol $resultCol  $durationCol" -ForegroundColor $resultColor

    # Show detection details for non-passing results
    if ($r.Result -notlike 'PASS*' -and $r.Result -notin @('SKIPPED', 'TIMEOUT')) {
        if ($null -ne $r.InstallSuccess) {
            $installLabel = if ($r.InstallSuccess) { 'Install OK' } else { "Install FAILED (exit $($r.InstallExitCode))" }
            Write-Host "    ├ $installLabel" -ForegroundColor DarkGray
        }
        if ($null -ne $r.DetectionAfterInstall) {
            $detLabel = if ($r.DetectionAfterInstall) { 'Detected after install' } else { 'NOT detected after install' }
            Write-Host "    ├ $detLabel" -ForegroundColor DarkGray
        }
        if ($null -ne $r.UninstallSuccess) {
            $uninstLabel = if ($r.UninstallSuccess) { 'Uninstall OK' } else { "Uninstall FAILED (exit $($r.UninstallExitCode))" }
            Write-Host "    ├ $uninstLabel" -ForegroundColor DarkGray
        }
        if ($null -ne $r.DetectionAfterUninstall) {
            $detUnLabel = if (-not $r.DetectionAfterUninstall) { 'Not detected after uninstall (correct)' } else { 'Still detected after uninstall' }
            Write-Host "    ├ $detUnLabel" -ForegroundColor DarkGray
        }
    }
    if ($r.Notes) {
        Write-Host "    └ $($r.Notes)" -ForegroundColor DarkGray
    }
}

Write-Host ''
Write-Host '──────────────────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host ("  {0} tested  |  {1} PASS  |  {2} FAIL  |  {3} TIMEOUT  |  {4} SKIPPED  |  {5} ERROR" -f `
    $total, $counts.Pass, $counts.Fail, $counts.Timeout, $counts.Skipped, $counts.Error) -ForegroundColor White
Write-Host '══════════════════════════════════════════════════════════' -ForegroundColor White
Write-Host ''
Write-Host "  CSV: $csvPath" -ForegroundColor Gray
Write-Host ''
