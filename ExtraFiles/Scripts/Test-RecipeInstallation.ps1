<#
.SYNOPSIS
    Tests CMPackager recipe install and uninstall commands using Windows Sandbox.

.DESCRIPTION
    Parses a CMPackager recipe XML file, extracts install/uninstall commands and detection
    methods for one deployment type, then orchestrates a Windows Sandbox session to:
      1. Install the application
      2. Validate the detection method confirms installation
      3. Uninstall the application
      4. Validate the detection method confirms removal

    Results are written to a JSON file in the temp workspace and displayed on completion.

    Requirements:
      - Windows Sandbox feature must be enabled
        (Settings > Optional Features > Windows Sandbox)
      - Run this script on a Windows host (not on macOS/Linux)
      - Installer file must be provided via -InstallerPath, or the recipe must contain
        a <URL> or <PrefetchScript> element so the installer can be downloaded automatically

.PARAMETER RecipePath
    Path to the CMPackager recipe XML file to test.

.PARAMETER DeploymentTypeName
    The Name attribute of the DeploymentType element to test (e.g. "DeploymentType1").
    Defaults to the first DeploymentType in the recipe.

.PARAMETER InstallerPath
    Path to the already-downloaded installer file on the host.
    If omitted, the installer is downloaded automatically using the recipe's <URL> or
    <PrefetchScript>. The downloaded file is placed in the sandbox workspace folder.

.PARAMETER CleanupInstaller
    Delete the installer file from the workspace after the test completes.
    Useful when downloading automatically to avoid leaving large files behind.

.PARAMETER WorkspacePath
    Host folder used as the sandbox mapped drive.
    Defaults to %TEMP%\CMPackagerSandboxTest.

.PARAMETER TimeoutMinutes
    Maximum minutes to wait for the sandbox test to complete before giving up.
    Defaults to 30.

.EXAMPLE
    .\Test-RecipeInstallation.ps1 -RecipePath "..\..\Recipes\7-Zip.xml" `
        -InstallerPath "C:\Temp\7Zipx64.msi"

.EXAMPLE
    .\Test-RecipeInstallation.ps1 -RecipePath "..\..\Disabled\MozillaFirefox.xml" `
        -DeploymentTypeName "DeploymentType1"

.EXAMPLE
    .\Test-RecipeInstallation.ps1 -RecipePath "..\..\Disabled\VLC.xml" -CleanupInstaller

.NOTES
    Author: CMPackager project — ExtraFiles\Scripts
    Requires: Windows 10/11 Pro or Enterprise with Windows Sandbox enabled
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) { throw "Recipe file not found: $_" }
        return $true
    })]
    [string]$RecipePath,

    [Parameter(Mandatory = $false)]
    [string]$DeploymentTypeName,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) { throw "Installer file not found: $_" }
        return $true
    })]
    [string]$InstallerPath,

    [Parameter(Mandatory = $false)]
    [string]$WorkspacePath = (Join-Path $env:TEMP "CMPackagerSandboxTest"),

    [Parameter(Mandatory = $false)]
    [int]$TimeoutMinutes = 30,

    [Parameter(Mandatory = $false)]
    [switch]$CleanupInstaller
)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

#region ── Helpers ──────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Cyan
}

function Write-Pass {
    param([string]$Message)
    Write-Host "  [PASS] $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  [FAIL] $Message" -ForegroundColor Red
}

function Write-Info {
    param([string]$Message)
    Write-Host "  [INFO] $Message" -ForegroundColor Gray
}

# Map SCCM hive names to PowerShell PSDrive paths
function Convert-HiveName {
    param([string]$Hive)
    switch ($Hive) {
        'LocalMachine'  { return 'HKLM:' }
        'CurrentUser'   { return 'HKCU:' }
        'ClassesRoot'   { return 'HKCR:' }
        'Users'         { return 'HKU:'  }
        'CurrentConfig' { return 'HKCC:' }
        default         { return 'HKLM:' }
    }
}

#endregion

#region ── Validate Prerequisites ───────────────────────────────────────────────

Write-Host "`nCMPackager Recipe Installation Tester" -ForegroundColor White
Write-Host "======================================`n" -ForegroundColor White

# Check Windows Sandbox is available
$sandboxExe = 'WindowsSandbox.exe'
$sandboxPath = "$env:SystemRoot\System32\WindowsSandbox.exe"
if (-not (Test-Path $sandboxPath)) {
    Write-Error @"
Windows Sandbox executable not found at: $sandboxPath
Enable it via: Settings > Optional Features > Windows Sandbox
(Requires Windows 10/11 Pro or Enterprise, build 18305+)
"@
    exit 1
}

#endregion

#region ── Parse Recipe ──────────────────────────────────────────────────────────

Write-Step "Parsing recipe: $RecipePath"
[xml]$recipe = Get-Content -Path $RecipePath -Raw

$appName = $recipe.ApplicationDef.Application.Name

# Select deployment type
$allDepTypes = $recipe.ApplicationDef.DeploymentTypes.DeploymentType
if ($PSBoundParameters.ContainsKey('DeploymentTypeName')) {
    $depType = $allDepTypes | Where-Object { $_.Name -eq $DeploymentTypeName } | Select-Object -First 1
    if (-not $depType) {
        Write-Error "DeploymentType '$DeploymentTypeName' not found in recipe. Available: $(($allDepTypes | ForEach-Object { $_.Name }) -join ', ')"
        exit 1
    }
} else {
    $depType = $allDepTypes | Select-Object -First 1
}

$depTypeName     = $depType.Name
$installType     = $depType.InstallationType   # MSI or Script
$installProgram  = $depType.InstallProgram
$installationMSI = $depType.InstallationMSI
$uninstallCmd    = $depType.UninstallCmd
$detMethodType   = $depType.DetectionMethodType  # MSI, Custom, CustomScript

Write-Info "Application     : $appName"
Write-Info "Deployment type : $depTypeName ($installType)"
Write-Info "Detection method: $detMethodType"

# Build the install command
if ([string]::IsNullOrWhiteSpace($installProgram)) {
    if ($installType -eq 'MSI' -and -not [string]::IsNullOrWhiteSpace($installationMSI)) {
        $installProgram = "msiexec.exe /i `"$installationMSI`" /qn /l*v install.log"
    } else {
        Write-Error "No InstallProgram and no InstallationMSI found for deployment type '$depTypeName'."
        exit 1
    }
}

# Build uninstall command fallback for pure MSI types
if ([string]::IsNullOrWhiteSpace($uninstallCmd)) {
    if ($installType -eq 'MSI' -and -not [string]::IsNullOrWhiteSpace($installationMSI)) {
        $uninstallCmd = "msiexec.exe /x `"$installationMSI`" /qn /l*v uninstall.log"
        Write-Info "No UninstallCmd in recipe; will use: $uninstallCmd"
    } else {
        Write-Warning "No uninstall command found. Uninstall step will be skipped."
        $uninstallCmd = ''
    }
}

# Find linked download for installer filename
$linkedDownload = $recipe.ApplicationDef.Downloads.Download |
    Where-Object { $_.DeploymentType -eq $depTypeName } |
    Select-Object -First 1

$installerFileName = if ($linkedDownload -and -not [string]::IsNullOrWhiteSpace($linkedDownload.DownloadFileName)) {
    $linkedDownload.DownloadFileName
} elseif (-not [string]::IsNullOrWhiteSpace($installationMSI)) {
    $installationMSI
} else {
    'installer'
}

# TODO: support ExtraCopyFunctions by staging the extra files into the sandbox workspace.
#       Until then, recipes that use it cannot be tested reliably and are skipped.
if ($linkedDownload -and -not [string]::IsNullOrWhiteSpace($linkedDownload.ExtraCopyFunctions)) {
    Write-Warning "Recipe '$appName' uses ExtraCopyFunctions, which is not yet supported by this tester. Skipping."
    exit 2
}

#endregion

#region ── Resolve Installer File ────────────────────────────────────────────────

if ($PSBoundParameters.ContainsKey('InstallerPath')) {
    $resolvedInstallerPath = $InstallerPath
    Write-Step "Using provided installer: $resolvedInstallerPath"
} else {
    $directUrl     = if ($linkedDownload) { $linkedDownload.URL }           else { $null }
    $prefetchScript = if ($linkedDownload) { $linkedDownload.PrefetchScript } else { $null }

    # Resolve a download URL — either directly from <URL> or by running <PrefetchScript>
    $resolvedUrl = $null

    if (-not [string]::IsNullOrWhiteSpace($directUrl)) {
        $resolvedUrl = $directUrl.Trim()
        Write-Step "Downloading installer from recipe URL: $resolvedUrl"

    } elseif (-not [string]::IsNullOrWhiteSpace($prefetchScript)) {
        Write-Step "Resolving installer URL via PrefetchScript..."

        # Load Get-InstallerURLfromWinget from CMPackager.ps1 so PrefetchScripts can call it
        $cmPackagerPath = Join-Path $PSScriptRoot '..' '..' 'CMPackager.ps1'
        $cmPackagerPath = [System.IO.Path]::GetFullPath($cmPackagerPath)
        if (Test-Path $cmPackagerPath) {
            try {
                $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                    $cmPackagerPath, [ref]$null, [ref]$null)
                $funcDef = $ast.FindAll({
                    $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $args[0].Name -eq 'Get-InstallerURLfromWinget'
                }, $false) | Select-Object -First 1
                if ($funcDef) {
                    Invoke-Expression $funcDef.Extent.Text
                    Write-Info "Loaded Get-InstallerURLfromWinget from CMPackager.ps1"
                } else {
                    Write-Warning "Get-InstallerURLfromWinget not found in CMPackager.ps1 — PrefetchScript may fail"
                }
            } catch {
                Write-Warning "Could not parse CMPackager.ps1: $_ — PrefetchScript may fail"
            }
        } else {
            Write-Warning "CMPackager.ps1 not found at $cmPackagerPath — PrefetchScript may fail"
        }

        # Execute PrefetchScript — it is expected to set $URL in the current scope
        $URL = $null
        try {
            Invoke-Expression $prefetchScript | Out-Null
        } catch {
            Write-Error "PrefetchScript execution failed: $_"
            exit 1
        }

        if ([string]::IsNullOrWhiteSpace($URL)) {
            Write-Error "PrefetchScript did not produce a URL. Check the recipe's <PrefetchScript> block."
            exit 1
        }

        $resolvedUrl = $URL.Trim()
        Write-Info "PrefetchScript resolved URL: $resolvedUrl"

    } else {
        Write-Error "No -InstallerPath provided and the recipe has neither a <URL> nor a <PrefetchScript>."
        exit 1
    }

    # Download to a local temp directory; the workspace copy step will place it in the shared folder
    $downloadDir = Join-Path $env:TEMP 'CMPackagerDownload'
    New-Item -ItemType Directory -Path $downloadDir -Force | Out-Null
    $resolvedInstallerPath = Join-Path $downloadDir $installerFileName

    Write-Info "Downloading to: $resolvedInstallerPath"
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $resolvedUrl -OutFile $resolvedInstallerPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'
        Write-Info "Download complete: $([math]::Round((Get-Item $resolvedInstallerPath).Length / 1MB, 1)) MB"
    } catch {
        Write-Error "Failed to download installer from $resolvedUrl`: $_"
        exit 1
    }
}

#endregion

#region ── Build Detection Clause Data ───────────────────────────────────────────

# Serialise detection clauses into a PowerShell literal that will be embedded
# verbatim into the generated RunTest.ps1 inside the sandbox.

$detectionClauseLiterals = @()

switch ($detMethodType) {
    'MSI' {
        # Product code is read from the MSI file at test time inside the sandbox
        $detectionClauseLiterals += "@{ Type='MSI'; InstallerFile='$installerFileName' }"
    }
    { $_ -in 'Custom', 'CustomScript' } {
        $clauses = $depType.CustomDetectionMethods.DetectionClause
        if ($null -eq $clauses) {
            Write-Warning "DetectionMethodType is Custom but no DetectionClauses found. Detection step will be skipped."
        } else {
            foreach ($clause in $clauses) {
                $clauseType = $clause.DetectionClauseType
                switch -Wildcard ($clauseType) {
                    'File' {
                        $filePath    = $clause.Path   -replace "'", "''"
                        $fileName    = $clause.Name   -replace "'", "''"
                        $propType    = $clause.PropertyType
                        $expectedVal = $clause.ExpectedValue -replace "'", "''"
                        $operator    = $clause.ExpressionOperator
                        $checkValue  = $clause.Value  # 'True' means check property, not just existence
                        $is64Bit     = $clause.Is64Bit
                        $detectionClauseLiterals += "@{ Type='File'; FilePath='$filePath'; FileName='$fileName'; PropertyType='$propType'; ExpectedValue='$expectedVal'; Operator='$operator'; CheckValue='$checkValue'; Is64Bit='$is64Bit' }"
                    }
                    'RegistryKey*' {
                        $hive    = Convert-HiveName $clause.Hive
                        $keyName = $clause.KeyName  -replace "'", "''"
                        $valName = $clause.ValueName -replace "'", "''"
                        $propType    = $clause.PropertyType
                        $expectedVal = $clause.ExpectedValue -replace "'", "''"
                        $operator    = $clause.ExpressionOperator
                        $checkValue  = $clause.Value
                        $detectionClauseLiterals += "@{ Type='Registry'; Hive='$hive'; KeyName='$keyName'; ValueName='$valName'; PropertyType='$propType'; ExpectedValue='$expectedVal'; Operator='$operator'; CheckValue='$checkValue' }"
                    }
                    'MSI' {
                        $productCode = $clause.ProductCode -replace "'", "''"
                        $detectionClauseLiterals += "@{ Type='MSI'; ProductCode='$productCode' }"
                    }
                    default {
                        Write-Warning "Unknown DetectionClauseType '$clauseType' — skipping this clause."
                    }
                }
            }
        }
    }
    default {
        Write-Warning "DetectionMethodType '$detMethodType' is not supported. Detection step will be skipped."
    }
}

# Render as a PowerShell array literal for embedding in the generated script
$detectionArrayLiteral = if ($detectionClauseLiterals.Count -gt 0) {
    "@(`n    " + ($detectionClauseLiterals -join ",`n    ") + "`n)"
} else {
    '@()'
}

#endregion

#region ── Prepare Workspace ─────────────────────────────────────────────────────

Write-Step "Preparing sandbox workspace: $WorkspacePath"

# Clean previous run
if (Test-Path $WorkspacePath) {
    Remove-Item $WorkspacePath -Recurse -Force
}
New-Item -ItemType Directory -Path $WorkspacePath -Force | Out-Null

# Copy installer into workspace
$sandboxInstallerPath = Join-Path $WorkspacePath $installerFileName
Copy-Item -Path $resolvedInstallerPath -Destination $sandboxInstallerPath -Force
Write-Info "Copied installer: $installerFileName"

#endregion

#region ── Generate Inner RunTest.ps1 ────────────────────────────────────────────

Write-Step "Generating sandbox test script"

$innerScript = @"
# ============================================================
# RunTest.ps1 -- Generated by Test-RecipeInstallation.ps1
# Runs inside Windows Sandbox. Do not edit manually.
# ============================================================
`$ErrorActionPreference = 'Continue'

# ── Injected values ──────────────────────────────────────────
`$AppName          = '$($appName -replace "'", "''")'
`$DepTypeName      = '$depTypeName'
`$InstallerFile    = '$installerFileName'
`$InstallationType = '$installType'
`$InstallCmd       = '$($installProgram -replace "'", "''" -replace '"', '`"')'
`$UninstallCmd     = '$($uninstallCmd   -replace "'", "''" -replace '"', '`"')'
`$DetectionClauses = $detectionArrayLiteral
`$ResultsPath      = 'C:\TestFiles\results.json'
`$LogPath          = 'C:\TestFiles\sandbox.log'

# ── Logging ──────────────────────────────────────────────────
function Write-Log {
    param([string]`$Message)
    `$ts = Get-Date -Format 'HH:mm:ss'
    "`$ts `$Message" | Tee-Object -FilePath `$LogPath -Append | Write-Host
}

# ── Execute a command string robustly via a temp batch file ──
function Invoke-RecipeCommand {
    param(
        [string]`$Command,
        [string]`$WorkingDir = 'C:\Temp\CMPackagerTest'
    )
    # Unescape PowerShell backtick-quotes used in recipe XML
    `$cmd = `$Command -replace '``"', '"'
    # Expand environment variables
    `$cmd = [System.Environment]::ExpandEnvironmentVariables(`$cmd)
    Write-Log "Running: `$cmd"

    `$batchFile = [System.IO.Path]::Combine(`$env:TEMP, 'recipe_cmd.cmd')
    @(
        '@echo off',
        ('cd /d "' + `$WorkingDir + '"'),
        `$cmd,
        'exit /b %ERRORLEVEL%'
    ) | Set-Content -Path `$batchFile -Encoding ASCII

    `$psi = New-Object System.Diagnostics.ProcessStartInfo
    `$psi.FileName  = 'cmd.exe'
    `$psi.Arguments = '/c "' + `$batchFile + '"'
    `$psi.UseShellExecute = `$false
    `$psi.CreateNoWindow  = `$true
    `$proc = [System.Diagnostics.Process]::Start(`$psi)
    `$proc.WaitForExit()
    `$exitCode = `$proc.ExitCode
    Remove-Item `$batchFile -ErrorAction SilentlyContinue
    return `$exitCode
}

# ── Wait for MsiInstaller Application Event Log entry ────────
# EventID 1033 = install completed, 1034 = uninstall completed
function Wait-MsiInstallerEvent {
    param(
        [string]`$ProductName,
        [datetime]`$After,
        [int[]]`$EventIds,
        [int]`$TimeoutMinutes = 20
    )
    Write-Log "Waiting for MsiInstaller event (IDs: `$(`$EventIds -join ',')) for product '`$ProductName'..."
    `$deadline = (Get-Date).AddMinutes(`$TimeoutMinutes)
    while ((Get-Date) -lt `$deadline) {
        `$events = Get-EventLog -LogName Application -Source MsiInstaller -Newest 20 -ErrorAction SilentlyContinue |
            Where-Object { `$_.TimeGenerated -gt `$After -and `$_.EventID -in `$EventIds -and `$_.Message -like "*`$ProductName*" }
        if (`$events) {
            Write-Log "  MsiInstaller event found (ID `$(`$events[0].EventID)) -- installer transaction complete."
            return
        }
        Start-Sleep -Seconds 5
    }
    Write-Log "  WARNING: timed out waiting for MsiInstaller event after `$TimeoutMinutes minutes."
}

# ── Extract ProductCode from an MSI file (COM) ───────────────
function Get-MSIProductCode {
    param([string]`$MsiPath)
    try {
        `$installer = New-Object -ComObject WindowsInstaller.Installer
        `$db = `$installer.GetType().InvokeMember(
            'OpenDatabase', 'InvokeMethod', `$null, `$installer,
            @(`$MsiPath, 0))
        `$view = `$db.GetType().InvokeMember(
            'OpenView', 'InvokeMethod', `$null, `$db,
            @("SELECT Value FROM Property WHERE Property='ProductCode'"))
        `$view.GetType().InvokeMember('Execute', 'InvokeMethod', `$null, `$view, `$null) | Out-Null
        `$record = `$view.GetType().InvokeMember('Fetch', 'InvokeMethod', `$null, `$view, `$null)
        if (`$null -eq `$record) {
            Write-Log "WARNING: MSI Property table returned no ProductCode row"
            return `$null
        }
        return `$record.GetType().InvokeMember('StringData', 'GetProperty', `$null, `$record, @(1))
    } catch {
        Write-Log "WARNING: Could not read ProductCode from MSI: `$_"
        return `$null
    }
}

# ── Check if a MSI ProductCode is registered as installed ────
function Test-MSIProductCodeInstalled {
    param([string]`$ProductCode)
    if ([string]::IsNullOrWhiteSpace(`$ProductCode)) { return `$false }
    `$paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`$ProductCode",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\`$ProductCode"
    )
    return (`$paths | Where-Object { Test-Path `$_ }).Count -gt 0
}

# ── Expand common environment variable placeholders in paths ─
function Expand-RecipePath {
    param([string]`$Path)
    `$Path = `$Path -replace '%ProgramFiles%',      `$env:ProgramFiles
    `$Path = `$Path -replace '%ProgramFiles\(x86\)%', `${env:ProgramFiles(x86)}
    `$Path = `$Path -replace '%SystemRoot%',          `$env:SystemRoot
    `$Path = `$Path -replace '%SystemDrive%',         `$env:SystemDrive
    `$Path = `$Path -replace '%WinDir%',              `$env:WinDir
    `$Path = `$Path -replace '%CommonProgramFiles%',  `$env:CommonProgramFiles
    `$Path = [System.Environment]::ExpandEnvironmentVariables(`$Path)
    return `$Path
}

# ── Version comparison helper ────────────────────────────────
function Compare-Versions {
    param([string]`$Actual, [string]`$Expected, [string]`$Operator)
    if ([string]::IsNullOrWhiteSpace(`$Expected) -or `$Expected -like '`$*') {
        return `$null  # Cannot compare — expected value is a variable placeholder
    }
    try {
        `$a = [System.Version]`$Actual
        `$e = [System.Version]`$Expected
        `$cmp = `$a.CompareTo(`$e)
        `$result = switch (`$Operator) {
            'Equals'         { `$cmp -eq 0 }
            'NotEquals'      { `$cmp -ne 0 }
            'GreaterThan'    { `$cmp -gt 0 }
            'GreaterEquals'  { `$cmp -ge 0 }
            'LessThan'       { `$cmp -lt 0 }
            'LessEquals'     { `$cmp -le 0 }
            default          { `$cmp -ge 0 }
        }
        return `$result
    } catch {
        return `$null  # Version parse failed — treat as existence-only check
    }
}

# ── Evaluate all detection clauses, return structured result ─
function Test-DetectionClauses {
    param([array]`$Clauses)

    if (`$Clauses.Count -eq 0) {
        return @{ Detected = `$null; Details = 'No detection clauses configured' }
    }

    `$clauseResults = @()
    foreach (`$clause in `$Clauses) {
        `$r = @{ Type = `$clause.Type; Detected = `$false; Detail = '' }

        switch (`$clause.Type) {
            'File' {
                `$expandedPath = Expand-RecipePath `$clause.FilePath
                `$fullPath = Join-Path `$expandedPath `$clause.FileName
                `$exists = Test-Path `$fullPath -PathType Leaf

                if (-not `$exists) {
                    `$r.Detail = "File not found: `$fullPath"
                } else {
                    `$r.Detected = `$true
                    `$r.Detail   = "File found: `$fullPath"

                    # Optional version check
                    if (`$clause.CheckValue -eq 'True' -and `$clause.PropertyType -eq 'Version') {
                        `$actualVer = (Get-Item `$fullPath -ErrorAction SilentlyContinue).VersionInfo.FileVersion
                        `$versionOk = Compare-Versions `$actualVer `$clause.ExpectedValue `$clause.Operator
                        if (`$null -ne `$versionOk) {
                            `$r.Detail += " | Version `$actualVer `$(if (`$versionOk) { 'satisfies' } else { 'does NOT satisfy' }) `$(`$clause.Operator) `$(`$clause.ExpectedValue)"
                            `$r.Detected = `$r.Detected -and `$versionOk
                        } else {
                            `$r.Detail += " | Version `$actualVer (expected value is a placeholder -- existence check only)"
                        }
                    }
                }
            }

            'Registry' {
                `$regPath = "`$(`$clause.Hive)\`$(`$clause.KeyName)"
                `$keyExists = Test-Path `$regPath

                if (-not `$keyExists) {
                    `$r.Detail = "Registry key not found: `$regPath"
                } elseif ([string]::IsNullOrWhiteSpace(`$clause.ValueName)) {
                    `$r.Detected = `$true
                    `$r.Detail   = "Registry key exists: `$regPath"
                } else {
                    `$actualVal = (Get-ItemProperty `$regPath -Name `$clause.ValueName -ErrorAction SilentlyContinue).(`$clause.ValueName)
                    if (`$null -eq `$actualVal) {
                        `$r.Detail = "Registry value not found: `$regPath\`$(`$clause.ValueName)"
                    } else {
                        `$r.Detected = `$true
                        `$r.Detail   = "Registry value: `$regPath\`$(`$clause.ValueName) = `$actualVal"

                        if (`$clause.CheckValue -eq 'True' -and `$clause.PropertyType -eq 'Version') {
                            `$versionOk = Compare-Versions ([string]`$actualVal) `$clause.ExpectedValue `$clause.Operator
                            if (`$null -ne `$versionOk) {
                                `$r.Detail += " | Version `$actualVal `$(if (`$versionOk) { 'satisfies' } else { 'does NOT satisfy' }) `$(`$clause.Operator) `$(`$clause.ExpectedValue)"
                                `$r.Detected = `$r.Detected -and `$versionOk
                            } else {
                                `$r.Detail += " | (expected value is a placeholder -- existence check only)"
                            }
                        }
                    }
                }
            }

            'MSI' {
                `$productCode = `$clause.ProductCode
                if ([string]::IsNullOrWhiteSpace(`$productCode)) {
                    # Extract from MSI file
                    `$msiFullPath = "C:\TestFiles\`$(`$clause.InstallerFile)"
                    Write-Log "Extracting ProductCode from `$msiFullPath"
                    `$pc = Get-MSIProductCode `$msiFullPath
                    `$productCode = if (`$null -ne `$pc) { `$pc.Trim() } else { '' }
                }

                if ([string]::IsNullOrWhiteSpace(`$productCode)) {
                    `$r.Detail = "Could not determine MSI ProductCode"
                } else {
                    `$r.Detected = Test-MSIProductCodeInstalled `$productCode
                    `$r.Detail   = "MSI ProductCode `$productCode `$(if (`$r.Detected) { 'found' } else { 'NOT found' }) in Uninstall registry"
                }
            }
        }

        Write-Log "Detection [`$(`$clause.Type)]: `$(`$r.Detail)"
        `$clauseResults += `$r
    }

    # Overall: all clauses must detect (AND logic — matches SCCM default)
    `$allDetected = (`$clauseResults | Where-Object { -not `$_.Detected }).Count -eq 0
    return @{
        Detected       = `$allDetected
        ClauseResults  = `$clauseResults
    }
}

# ════════════════════════════════════════════════════════════
# MAIN TEST SEQUENCE
# ════════════════════════════════════════════════════════════
Write-Log "=== Test starting: `$AppName / `$DepTypeName ==="
Set-Location 'C:\TestFiles'

# Brief pause for the sandbox to finish its own startup sequence.
# The Windows Installer service (msiserver) starts independently of the shell
# and is available well before explorer.exe. A short fixed wait is sufficient.
Write-Log "Waiting 15s for sandbox startup to settle..."
Start-Sleep -Seconds 15
Write-Log "System ready."

# Copy the installer to a fully local directory.
# The mapped folder C:\TestFiles is accessible to the logged-in user process
# but NOT to the Windows Installer service (msiserver), which runs as SYSTEM.
# SYSTEM cannot read from the WSB mapped share, causing msiexec to abort
# silently after "File will have security applied from OpCode."
`$localWorkDir = 'C:\Temp\CMPackagerTest'
New-Item -ItemType Directory -Path `$localWorkDir -Force | Out-Null
Copy-Item -Path "C:\TestFiles\`$InstallerFile" -Destination "`$localWorkDir\`$InstallerFile" -Force
Write-Log "Installer copied to local dir: `$localWorkDir\`$InstallerFile"

`$results = [ordered]@{
    Application          = `$AppName
    DeploymentType       = `$DepTypeName
    Timestamp            = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    InstallExitCode      = `$null
    InstallSuccess       = `$false
    DetectionAfterInstall  = `$null
    UninstallExitCode    = `$null
    UninstallSuccess     = `$false
    DetectionAfterUninstall = `$null
    OverallResult        = 'FAIL'
    Notes                = @()
}

# ── 1. Install ───────────────────────────────────────────────
Write-Log "--- Step 1: Install ---"
try {
    `$installStart = Get-Date
    `$exitCode = Invoke-RecipeCommand -Command `$InstallCmd
    `$results.InstallExitCode = `$exitCode
    # 0 = success, 3010 = success + reboot required, 1641 = success + reboot initiated
    `$results.InstallSuccess = `$exitCode -in @(0, 3010, 1641)
    Write-Log "Install exit code: `$exitCode (`$(if (`$results.InstallSuccess) { 'SUCCESS' } else { 'FAILURE' }))"
} catch {
    `$results.Notes += "Install exception: `$_"
    Write-Log "ERROR during install: `$_"
}

# MSI only: wait for Windows Installer to commit the transaction before detecting
if (`$InstallationType -eq 'MSI') {
    if (`$null -eq `$installStart) { `$installStart = (Get-Date).AddMinutes(-1) }
    Wait-MsiInstallerEvent -ProductName `$AppName -After `$installStart -EventIds @(1033)
}
Start-Sleep -Seconds 3

# ── 2. Detect after install ──────────────────────────────────
Write-Log "--- Step 2: Detection after install ---"
`$detInstall = Test-DetectionClauses -Clauses `$DetectionClauses
`$results.DetectionAfterInstall = `$detInstall

# ── 3. Uninstall ─────────────────────────────────────────────
Write-Log "--- Step 3: Uninstall ---"
if ([string]::IsNullOrWhiteSpace(`$UninstallCmd)) {
    `$results.Notes += 'No uninstall command -- uninstall step skipped'
    `$results.UninstallSuccess = `$null
    Write-Log "No uninstall command configured -- skipping"
} else {
    try {
        `$uninstallStart = Get-Date
        `$exitCode = Invoke-RecipeCommand -Command `$UninstallCmd
        `$results.UninstallExitCode = `$exitCode
        `$results.UninstallSuccess  = `$exitCode -in @(0, 3010, 1641)
        Write-Log "Uninstall exit code: `$exitCode (`$(if (`$results.UninstallSuccess) { 'SUCCESS' } else { 'FAILURE' }))"
    } catch {
        `$results.Notes += "Uninstall exception: `$_"
        Write-Log "ERROR during uninstall: `$_"
    }

    # MSI only: wait for Windows Installer to commit the transaction before detecting
    if (`$InstallationType -eq 'MSI') {
        if (`$null -eq `$uninstallStart) { `$uninstallStart = (Get-Date).AddMinutes(-1) }
        Wait-MsiInstallerEvent -ProductName `$AppName -After `$uninstallStart -EventIds @(1034)
    }
    Start-Sleep -Seconds 3
}

# ── 4. Detect after uninstall ────────────────────────────────
Write-Log "--- Step 4: Detection after uninstall ---"
`$detUninstall = Test-DetectionClauses -Clauses `$DetectionClauses
`$results.DetectionAfterUninstall = `$detUninstall

# ── 5. Overall result ────────────────────────────────────────
`$installOk   = `$results.InstallSuccess -eq `$true
`$detAfterOk  = `$detInstall.Detected    -eq `$true
`$detAfterUn  = `$detUninstall.Detected  -eq `$false  # should NOT be detected after uninstall
`$uninstallOk = (`$results.UninstallSuccess -eq `$true) -or (`$null -eq `$results.UninstallSuccess)

if (`$installOk -and `$detAfterOk -and `$uninstallOk -and `$detAfterUn) {
    `$results.OverallResult = 'PASS'
} elseif (`$installOk -and `$detAfterOk -and [string]::IsNullOrWhiteSpace(`$UninstallCmd)) {
    `$results.OverallResult = 'PASS (install only -- no uninstall configured)'
} else {
    `$results.OverallResult = 'FAIL'
}

Write-Log "=== Overall result: `$(`$results.OverallResult) ==="

# ── 6. Write results ─────────────────────────────────────────
`$results | ConvertTo-Json -Depth 10 | Set-Content -Path `$ResultsPath -Encoding UTF8
Write-Log "Results written to `$ResultsPath"

# ── 7. Copy installer logs to mapped folder ──────────────────
foreach (`$logName in @('install.log', 'uninstall.log')) {
    `$src = Join-Path 'C:\Temp\CMPackagerTest' `$logName
    if (Test-Path `$src) {
        Copy-Item `$src 'C:\TestFiles\' -Force -ErrorAction SilentlyContinue
        Write-Log "Copied `$logName to C:\TestFiles\"
    }
}

# ── 8. Shut down sandbox ─────────────────────────────────────
Write-Log "Shutting down sandbox..."
Start-Sleep -Seconds 2
& shutdown /s /t 0
"@

$innerScriptPath = Join-Path $WorkspacePath 'RunTest.ps1'
$innerScript | Set-Content -Path $innerScriptPath -Encoding Unicode
Write-Info "Generated: RunTest.ps1"

# Generate a .cmd launcher so the LogonCommand does not rely on WSB honouring
# the -ExecutionPolicy flag when invoking .ps1 files directly.
$cmdLauncherPath = Join-Path $WorkspacePath 'Run.cmd'
@"
@echo off
powershell.exe -ExecutionPolicy Bypass -NonInteractive -File "C:\TestFiles\RunTest.ps1"
"@ | Set-Content -Path $cmdLauncherPath -Encoding ASCII
Write-Info "Generated: Run.cmd"

#endregion

#region ── Generate .wsb Config ──────────────────────────────────────────────────

Write-Step "Generating Windows Sandbox configuration"

$wsbContent = @"
<Configuration>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$WorkspacePath</HostFolder>
      <SandboxFolder>C:\TestFiles</SandboxFolder>
      <ReadOnly>false</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
    <Command>C:\TestFiles\Run.cmd</Command>
  </LogonCommand>
</Configuration>
"@

$wsbPath = Join-Path $WorkspacePath 'SandboxTest.wsb'
$wsbContent | Set-Content -Path $wsbPath -Encoding UTF8
Write-Info "Generated: SandboxTest.wsb"

#endregion

#region ── Launch Sandbox and Wait ───────────────────────────────────────────────

$resultsFile = Join-Path $WorkspacePath 'results.json'

Write-Host "`n" -NoNewline
Write-Step "Launching Windows Sandbox..."
Write-Info "The sandbox window will open. A PowerShell window will run the test automatically."
Write-Info "The sandbox will shut down when the test completes."
Write-Info "Timeout: $TimeoutMinutes minutes"
Write-Host ""

Start-Process -FilePath $sandboxPath -ArgumentList $wsbPath

# Poll for results file
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$dotCount  = 0
Write-Host "  Waiting for results" -NoNewline -ForegroundColor Cyan

while (-not (Test-Path $resultsFile) -and (Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    Write-Host '.' -NoNewline -ForegroundColor Cyan
    $dotCount++
    if ($dotCount % 12 -eq 0) {
        $remaining = [math]::Round(($deadline - (Get-Date)).TotalMinutes, 1)
        Write-Host " ($remaining min remaining)" -NoNewline -ForegroundColor Gray
    }
}
Write-Host ""

if (-not (Test-Path $resultsFile)) {
    Write-Host "`n══════════════════════════════════════" -ForegroundColor White
    Write-Host "  TIMED OUT after $TimeoutMinutes minutes" -ForegroundColor Red
    Write-Host "══════════════════════════════════════`n" -ForegroundColor White
    Write-Host "  The sandbox did not produce a results file within the allowed time." -ForegroundColor Yellow
    Write-Host "  The sandbox may still be running. Check the log for the last known state:`n" -ForegroundColor Yellow
    $sandboxLog = Join-Path $WorkspacePath 'sandbox.log'
    if (Test-Path $sandboxLog) {
        Write-Info "Sandbox log: $sandboxLog"
        Write-Host ""
        Write-Host "  Last 20 lines:" -ForegroundColor Gray
        Get-Content $sandboxLog -Tail 20 | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
    } else {
        Write-Host "  No sandbox.log found — the sandbox may not have started correctly." -ForegroundColor Red
    }
    foreach ($logName in @('install.log', 'uninstall.log')) {
        $logPath = Join-Path $WorkspacePath $logName
        if (Test-Path $logPath) {
            Write-Host ""
            Write-Info "${logName}: $logPath"
        }
    }
    Write-Host ""
    exit 1
}

# Give the sandbox a moment to finish writing before reading
Start-Sleep -Seconds 2

#endregion

#region ── Display Results ───────────────────────────────────────────────────────

Write-Host "`n══════════════════════════════════════" -ForegroundColor White
Write-Host "  TEST RESULTS" -ForegroundColor White
Write-Host "══════════════════════════════════════`n" -ForegroundColor White

try {
    $results = Get-Content $resultsFile -Raw | ConvertFrom-Json
} catch {
    Write-Error "Failed to parse results file: $_"
    exit 1
}

Write-Host "  Application    : $($results.Application)"
Write-Host "  Deployment type: $($results.DeploymentType)"
Write-Host "  Timestamp      : $($results.Timestamp)"
Write-Host ""

# Install result
$installColor = if ($results.InstallSuccess) { 'Green' } else { 'Red' }
Write-Host "  Install" -ForegroundColor White
Write-Host "    Exit code : $($results.InstallExitCode)" -ForegroundColor $installColor
Write-Host "    Success   : $($results.InstallSuccess)"  -ForegroundColor $installColor

# Detection after install
Write-Host "`n  Detection after install" -ForegroundColor White
$detInstall = $results.DetectionAfterInstall
$detInstallColor = if ($detInstall.Detected) { 'Green' } else { 'Red' }
Write-Host "    Detected  : $($detInstall.Detected)" -ForegroundColor $detInstallColor
foreach ($cr in $detInstall.ClauseResults) {
    Write-Host "    [$($cr.Type)] $($cr.Detail)" -ForegroundColor Gray
}

# Uninstall result
Write-Host "`n  Uninstall" -ForegroundColor White
if ($null -eq $results.UninstallExitCode -and $results.UninstallSuccess -eq $null) {
    Write-Host "    Skipped (no uninstall command configured)" -ForegroundColor Yellow
} else {
    $uninstallColor = if ($results.UninstallSuccess) { 'Green' } else { 'Red' }
    Write-Host "    Exit code : $($results.UninstallExitCode)" -ForegroundColor $uninstallColor
    Write-Host "    Success   : $($results.UninstallSuccess)"  -ForegroundColor $uninstallColor
}

# Detection after uninstall
Write-Host "`n  Detection after uninstall" -ForegroundColor White
$detUninstall = $results.DetectionAfterUninstall
# For uninstall check: NOT detected is the good outcome
$detUninstallColor = if (-not $detUninstall.Detected) { 'Green' } else { 'Red' }
Write-Host "    Detected  : $($detUninstall.Detected) $(if (-not $detUninstall.Detected) { '(correct — app is gone)' } else { '(UNEXPECTED — app still detected)' })" -ForegroundColor $detUninstallColor
foreach ($cr in $detUninstall.ClauseResults) {
    Write-Host "    [$($cr.Type)] $($cr.Detail)" -ForegroundColor Gray
}

# Notes
if ($results.Notes -and $results.Notes.Count -gt 0) {
    Write-Host "`n  Notes" -ForegroundColor White
    foreach ($note in $results.Notes) {
        Write-Host "    $note" -ForegroundColor Yellow
    }
}

# Overall
Write-Host ""
Write-Host "══════════════════════════════════════" -ForegroundColor White
$overallColor = if ($results.OverallResult -like 'PASS*') { 'Green' } else { 'Red' }
Write-Host "  OVERALL: $($results.OverallResult)" -ForegroundColor $overallColor
Write-Host "══════════════════════════════════════`n" -ForegroundColor White

# Show sandbox log location
$sandboxLog = Join-Path $WorkspacePath 'sandbox.log'
if (Test-Path $sandboxLog) {
    Write-Info "Sandbox log: $sandboxLog"
}
Write-Info "Results JSON: $resultsFile"
foreach ($logName in @('install.log', 'uninstall.log')) {
    $logPath = Join-Path $WorkspacePath $logName
    if (Test-Path $logPath) {
        Write-Info "${logName}: $logPath"
    }
}
Write-Host ""

if ($CleanupInstaller) {
    $installerInWorkspace = Join-Path $WorkspacePath $installerFileName
    if (Test-Path $installerInWorkspace) {
        Remove-Item $installerInWorkspace -Force
        Write-Info "Installer removed: $installerInWorkspace"
    }
}

#endregion
