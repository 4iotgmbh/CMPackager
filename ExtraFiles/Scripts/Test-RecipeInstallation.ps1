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
$script:testStartTime  = Get-Date

#region ── Helpers ──────────────────────────────────────────────────────────────

function Get-ElapsedPrefix {
    $elapsed = (Get-Date) - $script:testStartTime
    $m = [int]$elapsed.TotalMinutes
    $s = $elapsed.Seconds
    return '[{0}:{1:d2}]' -f $m, $s
}

function Write-Step {
    param([string]$Message)
    Write-Host "  $(Get-ElapsedPrefix) $Message" -ForegroundColor Cyan
}

function Write-Pass {
    param([string]$Message)
    Write-Host "  $(Get-ElapsedPrefix) [PASS] $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  $(Get-ElapsedPrefix) [FAIL] $Message" -ForegroundColor Red
}

function Write-Info {
    param([string]$Message)
    Write-Host "  $(Get-ElapsedPrefix) [INFO] $Message" -ForegroundColor Gray
}

# Stub for CMPackager's Add-LogContent. Recipe scripts (DownloadVersionCheck,
# ExtraCopyFunctions) call it for diagnostic logging. We write to the console and
# append to sandbox.log in the workspace when the workspace directory already exists.
function Add-LogContent {
    param([string]$Content)
    $ts = Get-Date -Format 'HH:mm:ss'
    Write-Host "  [$ts] [LOG] $Content" -ForegroundColor DarkGray
    if ($WorkspacePath -and (Test-Path $WorkspacePath -PathType Container)) {
        Add-Content -Path (Join-Path $WorkspacePath 'sandbox.log') -Value "[$ts] $Content" -ErrorAction SilentlyContinue
    }
}

# Extract ProductCode from an MSI file using the Windows Installer COM object (host-side).
# Called before sandbox launch so Detect.ps1 only needs a registry lookup (no COM in sandbox).
function Get-MsiProductCodeFromFile {
    param([string]$MsiPath)
    if (-not (Test-Path $MsiPath)) { return $null }
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($MsiPath, 0))
        $view = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db,
            @("SELECT Value FROM Property WHERE Property='ProductCode'"))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if ($null -eq $record) { return $null }
        return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, @(1))
    } catch {
        Write-Warning "Could not extract ProductCode from MSI on host: $_"
        return $null
    }
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

# Detect wsb.exe (Windows 11 24H2+ CLI) vs WindowsSandbox.exe (legacy launcher).
# wsb.exe lets us launch with a stable sandbox ID and stop it by ID — avoiding the
# "only one instance" problem that plagues the legacy launcher path.
# Use Get-Command so PATH is searched; wsb.exe may not always be in System32 directly.
$wsbCliCmd     = Get-Command 'wsb.exe' -ErrorAction SilentlyContinue
$wsbCliPath    = if ($wsbCliCmd) { $wsbCliCmd.Source } else { "$env:SystemRoot\System32\wsb.exe" }
$wsbLaunchPath = "$env:SystemRoot\System32\WindowsSandbox.exe"
$useWsbCli     = [bool]$wsbCliCmd

if (-not $useWsbCli -and -not (Test-Path $wsbLaunchPath)) {
    Write-Error @"
Windows Sandbox not found.
  Checked (24H2+ CLI)  : $wsbCliPath
  Checked (legacy)     : $wsbLaunchPath
Enable it via: Settings > Optional Features > Windows Sandbox
(Requires Windows 10/11 Pro or Enterprise, build 18305+)
"@
    exit 1
}

if ($useWsbCli) {
    Write-Info "Sandbox control: wsb.exe CLI (Windows 11 24H2+)"
} else {
    Write-Info "Sandbox control: WindowsSandbox.exe (legacy mode)"
}

#endregion

#region ── Parse Recipe ──────────────────────────────────────────────────────────

Write-Step "Parsing recipe: $RecipePath"
try {
    [xml]$recipe = Get-Content -Path $RecipePath -Raw
} catch {
    Write-Host "  [ERROR] Recipe XML is invalid and cannot be parsed: $_" -ForegroundColor Red
    Write-Host "          Check for unescaped characters in PrefetchScript or other text nodes." -ForegroundColor Red
    Write-Host "          Use &lt; for <, &gt; for >, &amp; for & in XML content." -ForegroundColor Red
    exit 1
}

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


#endregion

#region ── Resolve Installer File ────────────────────────────────────────────────

# These are populated by DownloadVersionCheck when downloading automatically;
# they remain null when InstallerPath is provided by the caller.
$Version = $null; $FullVersion = $null

if ($PSBoundParameters.ContainsKey('InstallerPath')) {
    $resolvedInstallerPath = $InstallerPath
    Write-Step "Using provided installer: $resolvedInstallerPath"
} else {
    # When a node contains a CDATA section, PowerShell's XML adapter returns an XmlElement
    # rather than a plain string.  Extract .InnerText in that case to get the raw content.
    $directUrl = if ($linkedDownload) {
        $n = $linkedDownload.URL
        if ($n -is [System.Xml.XmlElement]) { $n.InnerText } else { [string]$n }
    } else { $null }
    $prefetchScript = if ($linkedDownload) {
        $n = $linkedDownload.PrefetchScript
        if ($n -is [System.Xml.XmlElement]) { $n.InnerText } else { [string]$n }
    } else { $null }
    $downloadVersionCheck = if ($linkedDownload) {
        $n = $linkedDownload.DownloadVersionCheck
        if ($n -is [System.Xml.XmlElement]) { $n.InnerText } else { [string]$n }
    } else { $null }

    # Load helper functions from CMPackager.ps1.
    # Get-InstallerURLfromWinget is needed by PrefetchScripts; Get-MSIInfo is needed by
    # DownloadVersionCheck scripts (e.g. 7-Zip.xml, GoogleChrome.xml).
    $cmPackagerPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..' '..' 'CMPackager.ps1'))
    if (Test-Path $cmPackagerPath) {
        try {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $cmPackagerPath, [ref]$null, [ref]$null)
            foreach ($funcName in @('Get-InstallerURLfromWinget', 'Get-MSIInfo')) {
                $funcDef = $ast.FindAll({
                    $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $args[0].Name -eq $funcName
                }, $false) | Select-Object -First 1
                if ($funcDef) {
                    Invoke-Expression $funcDef.Extent.Text
                    Write-Info "Loaded $funcName from CMPackager.ps1"
                } else {
                    Write-Warning "$funcName not found in CMPackager.ps1"
                }
            }
        } catch {
            Write-Warning "Could not parse CMPackager.ps1: $_ — PrefetchScript/DownloadVersionCheck may fail"
        }
    } else {
        Write-Warning "CMPackager.ps1 not found at $cmPackagerPath — PrefetchScript/DownloadVersionCheck may fail"
    }

    # Resolve a download URL — either directly from <URL> or by running <PrefetchScript>
    $resolvedUrl = $null

    if (-not [string]::IsNullOrWhiteSpace($directUrl)) {
        $resolvedUrl = $directUrl.Trim()
        Write-Step "Downloading installer from recipe URL: $resolvedUrl"

    } elseif (-not [string]::IsNullOrWhiteSpace($prefetchScript)) {
        Write-Step "Resolving installer URL via PrefetchScript..."

        # Execute PrefetchScript — it is expected to set $URL in the current scope.
        # $PSScriptRoot is an automatic variable that is not inherited inside Invoke-Expression.
        # Expose the script directory as $CMPackagerScriptRoot so PrefetchScripts that call helper
        # scripts (e.g. via powershell.exe -File (Join-Path $CMPackagerScriptRoot ...)) can resolve
        # paths correctly without relying on $PSScriptRoot.
        $CMPackagerScriptRoot = $PSScriptRoot
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
        Invoke-WebRequest -Uri $resolvedUrl -OutFile $resolvedInstallerPath -UseBasicParsing -AllowInsecureRedirect -ErrorAction Stop
        $ProgressPreference = 'Continue'
        Write-Info "Download complete: $([math]::Round((Get-Item $resolvedInstallerPath).Length / 1MB, 1)) MB"
    } catch {
        Write-Error "Failed to download installer from $resolvedUrl`: $_"
        exit 1
    }

    # Run DownloadVersionCheck to extract $Version and $FullVersion from the downloaded file.
    # These are set by the recipe script and used later for CustomScript detection substitution.
    if (-not [string]::IsNullOrWhiteSpace($downloadVersionCheck)) {
        Write-Step "Running DownloadVersionCheck to extract version..."
        $DownloadFile = $resolvedInstallerPath                                       # consumed inside Invoke-Expression
        $TempDir      = $downloadDir                                               # consumed inside Invoke-Expression
        $ScriptRoot   = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))   # consumed inside Invoke-Expression
        $null = $DownloadFile, $TempDir, $ScriptRoot                               # suppress PSUseDeclaredVarsMoreThanAssignments
        try {
            Invoke-Expression $downloadVersionCheck | Out-Null
            if ($Version) {
                $versionLabel = $Version
                if ($FullVersion -and $FullVersion -ne $Version) { $versionLabel += " (full: $FullVersion)" }
                Write-Info "Detected version: $versionLabel"
            } else {
                Write-Warning "DownloadVersionCheck ran but did not set `$Version."
            }
        } catch {
            Write-Warning "DownloadVersionCheck failed: $_"
        }
    }
}

#endregion

#region ── Prepare Workspace ─────────────────────────────────────────────────────

Write-Step "Preparing sandbox workspace: $WorkspacePath"

# Stop any lingering sandbox before touching the workspace — a running sandbox holds
# an open handle to the mapped folder and will cause Remove-Item to fail.
if ($useWsbCli) {
    $staleIds = & $wsbCliPath list 2>&1 | Where-Object { $_ -match '^[0-9a-fA-F]{8}-' }
    foreach ($staleId in $staleIds) {
        Write-Info "Stopping stale sandbox $staleId before workspace cleanup..."
        & $wsbCliPath stop --id $staleId 2>&1 | Out-Null
    }
    if ($staleIds) { Start-Sleep -Seconds 3 }
} else {
    $staleProcs = Get-Process -Name 'WindowsSandbox', 'WindowsSandboxClient' -ErrorAction SilentlyContinue
    if ($staleProcs) {
        Write-Info "Stopping stale sandbox processes before workspace cleanup..."
        $staleProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
}

# Clean previous run
if (Test-Path $WorkspacePath) {
    Remove-Item $WorkspacePath -Recurse -Force
}
New-Item -ItemType Directory -Path $WorkspacePath -Force | Out-Null

# Copy installer into workspace
$sandboxInstallerPath = Join-Path $WorkspacePath $installerFileName
Copy-Item -Path $resolvedInstallerPath -Destination $sandboxInstallerPath -Force
Write-Info "Copied installer: $installerFileName"

# Run ExtraCopyFunctions to stage any additional files into the workspace.
# This mirrors what CMPackager.ps1 does when preparing the SCCM content repository.
# Variables are mapped to their CMPackager equivalents so recipe scripts work unchanged.
$extraCopyFunctions = if ($linkedDownload) {
    $n = $linkedDownload.ExtraCopyFunctions
    if ($n -is [System.Xml.XmlElement]) { $n.InnerText } else { [string]$n }
} else { $null }

if (-not [string]::IsNullOrWhiteSpace($extraCopyFunctions)) {
    Write-Step "Running ExtraCopyFunctions to stage additional files into workspace..."
    $DownloadFile    = $resolvedInstallerPath                                         # CMPackager: path to downloaded file
    $TempDir         = if ($downloadDir) { $downloadDir } else { $env:TEMP }         # CMPackager: temp/download directory
    $DestinationPath = $WorkspacePath                                                  # CMPackager: content distribution path
    $ScriptRoot      = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))   # CMPackager: repo root (where 7za.exe lives)
    $Recipe          = $recipe                                                         # CMPackager: full parsed recipe XML
    $null = $DownloadFile, $TempDir, $DestinationPath, $ScriptRoot, $Recipe           # suppress PSUseDeclaredVarsMoreThanAssignments
    try {
        Invoke-Expression $extraCopyFunctions | Out-Null
        Write-Info "ExtraCopyFunctions complete. Workspace contents:"
        Get-ChildItem $WorkspacePath -Recurse | ForEach-Object {
            Write-Info "  $($_.FullName.Substring($WorkspacePath.Length + 1))"
        }
    } catch {
        Write-Warning "ExtraCopyFunctions failed: $_. Extra files may be missing from the workspace."
    }
}

#endregion

#region ── Build Detection Clause Data ───────────────────────────────────────────
# Placed after ExtraCopyFunctions so MSI files staged by it are available for
# ProductCode extraction (e.g. Citrix where a wrapper EXE extracts RIInstaller.msi).

$detectionClauseLiterals = @()

switch ($detMethodType) {
    'MSI' {
        # Extract ProductCode on the host so Detect.ps1 only needs a registry lookup.
        # Using the Windows Installer COM object inside the sandbox (SYSTEM context via wsb exec)
        # hangs for ~60 s — the COM server is unavailable to SYSTEM in the sandbox session.
        #
        # If the recipe's main installer is a wrapper EXE, ExtraCopyFunctions will have already
        # staged the real MSI into the workspace. Try the main installer first; if it is not an
        # MSI (or ProductCode extraction fails), scan the workspace for any staged .msi files.
        $hostPc = $null
        $msiFileUsed = $installerFileName
        $hostPc = Get-MsiProductCodeFromFile $resolvedInstallerPath
        if (-not $hostPc) {
            # Main file is not an MSI or ProductCode could not be read — look for staged MSIs
            $stagedMsi = Get-ChildItem -Path $WorkspacePath -Filter '*.msi' -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($stagedMsi) {
                Write-Info "Main installer is not an MSI; trying staged MSI: $($stagedMsi.Name)"
                $hostPc = Get-MsiProductCodeFromFile $stagedMsi.FullName
                $msiFileUsed = $stagedMsi.Name
            }
        }
        if ($hostPc) {
            Write-Info "MSI ProductCode (from host, file: $msiFileUsed): $($hostPc.Trim())"
            $detectionClauseLiterals += "@{ Type='MSI'; ProductCode='$($hostPc.Trim() -replace "'", "''")' }"
        } else {
            Write-Warning "Could not extract ProductCode from MSI on host — falling back to in-sandbox extraction (may be slow)"
            $detectionClauseLiterals += "@{ Type='MSI'; InstallerFile='$msiFileUsed' }"
        }
    }
    'Custom' {
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
    'CustomScript' {
        # Detection is a PowerShell script embedded in <DetectionMethod>.
        # Extract it now; the script will be written as CustomDetect.ps1 in the workspace
        # and wrapped by a generated Detect.ps1 that captures its exit code and emits JSON.
        $n = $depType.DetectionMethod
        $customDetectRaw = if ($n -is [System.Xml.XmlElement]) { $n.InnerText } else { [string]$n }
        # Substitute $Version / $FullVersion exactly as CMPackager.ps1 does at deploy time.
        $customDetectScript = $customDetectRaw `
            -replace [regex]::Escape('$Version'),     ($Version     ?? '') `
            -replace [regex]::Escape('$FullVersion'), ($FullVersion ?? '')
        $scriptLines = ($customDetectScript -split '\n').Count
        Write-Info "CustomScript detection: $scriptLines-line script extracted$(if ($Version) { " (version substituted: $Version)" })"
        # No structured clauses needed — detection is driven entirely by the custom script.
    }
    default {
        Write-Warning "DetectionMethodType '$detMethodType' is not supported. Detection step will be skipped."
    }
}

# Render as a PowerShell array literal for embedding in the generated legacy script
$detectionArrayLiteral = if ($detectionClauseLiterals.Count -gt 0) {
    "@(`n    " + ($detectionClauseLiterals -join ",`n    ") + "`n)"
} else {
    '@()'
}

#endregion

#region ── Generate Test Artifacts ──────────────────────────────────────────────

Write-Step "Generating sandbox test artifacts"

if ($useWsbCli) {

    # ── wsb.exe path: small independent scripts driven via wsb exec ──────────────
    # No LogonCommand — the outer script orchestrates every step directly.

    # Write detection clause data as JSON so Detect.ps1 needs no embedded variables
    if ($detectionClauseLiterals.Count -gt 0) {
        $clauseObjects = $detectionClauseLiterals | ForEach-Object { Invoke-Expression $_ }
        $clauseObjects | ConvertTo-Json -Depth 5 |
            Set-Content (Join-Path $WorkspacePath 'detection_clauses.json') -Encoding UTF8
    } else {
        '[]' | Set-Content (Join-Path $WorkspacePath 'detection_clauses.json') -Encoding UTF8
    }
    Write-Info "Generated: detection_clauses.json"

    # sandbox_setup.cmd — run once after sandbox ready, before any installation.
    # Windows 11 24H2 Sandbox performs code-integrity safety checks on MSI files
    # that cause msiexec to stall for minutes. Disabling VerifiedAndReputablePolicyState
    # and refreshing the CI policy via CiTool resolves this.
    # See https://github.com/microsoft/Windows-Sandbox/issues/68#issuecomment-2684473932
    # xcopy /E copies all workspace files (installer, Detect.ps1, detection_clauses.json,
    # CustomDetect.ps1 if present, and any files staged by ExtraCopyFunctions) so no
    # individual copy lines are needed and extra files are handled automatically.
    @'
@echo off
REG add HKLM\SYSTEM\CurrentControlSet\Control\CI\Policy /v VerifiedAndReputablePolicyState /t REG_DWORD /d 0 /f
CiTool -r <NUL
mkdir C:\Temp\CMPackagerTest 2>nul
xcopy /E /Y /I "C:\TestFiles" "C:\Temp\CMPackagerTest" >nul 2>&1
'@ | Set-Content (Join-Path $WorkspacePath 'sandbox_setup.cmd') -Encoding ASCII
    Write-Info "Generated: sandbox_setup.cmd"

    # Detect.ps1 / CustomDetect.ps1 generation:
    # For clause-based detection (MSI / Custom): a single Detect.ps1 reads detection_clauses.json
    #   and emits a JSON result to stdout.
    # For CustomScript detection: CustomDetect.ps1 contains the recipe's detection script (with
    #   $Version already substituted); Detect.ps1 is a thin wrapper that runs it as a subprocess,
    #   captures its exit code, and emits the same JSON result format to stdout.
    # In both cases, detect_after_*.cmd runs Detect.ps1 unchanged — no branching needed there.
    #   stdout  → JSON result object (captured by detect_after_*.cmd via > redirect)
    #   stderr  → diagnostic log lines (captured by detect_after_*.cmd via 2> redirect)
    # PowerShell Set-Content/Add-Content cannot write from SYSTEM context (wsb exec), so
    # all output goes via the subprocess streams that cmd.exe can redirect to C:\TestFiles\.
    if ($detMethodType -eq 'CustomScript') {
        # Write the user's detection script with $Version already substituted.
        # UTF-8 encoding so multi-byte characters in the recipe script are preserved.
        $customDetectScript | Set-Content (Join-Path $WorkspacePath 'CustomDetect.ps1') -Encoding UTF8
        Write-Info "Generated: CustomDetect.ps1"
        # Write a thin Detect.ps1 wrapper — ASCII only (PS 5.1 encoding rule).
        @'
# Detect.ps1 - CustomScript detection wrapper
# Generated by Test-RecipeInstallation.ps1
# Runs CustomDetect.ps1 as a subprocess and emits a JSON result to stdout.
# stdout: JSON result   stderr: diagnostic log
# NOTE: ASCII only - encoding-safe for PowerShell 5.1 in sandbox
$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    [Console]::Error.WriteLine("$ts $Message")
}

$customScript = 'C:\Temp\CMPackagerTest\CustomDetect.ps1'
if (-not (Test-Path $customScript)) {
    Write-Log "ERROR: CustomDetect.ps1 not found at $customScript"
    [PSCustomObject]@{ Detected = $false; Output = ''; ExitCode = -1 } | ConvertTo-Json -Compress
    exit 1
}

Write-Log "Running custom detection script: $customScript"
$output = (& powershell.exe -NonInteractive -ExecutionPolicy Bypass -File $customScript 2>&1) | Out-String
$exitCode = $LASTEXITCODE
Write-Log "Custom detection exit code: $exitCode"
Write-Log "Custom detection output: $($output.Trim())"

$detected = ($exitCode -eq 0)
[PSCustomObject]@{ Detected = $detected; Output = $output.Trim(); ExitCode = $exitCode } | ConvertTo-Json -Compress
if ($detected) { exit 0 } else { exit 1 }
'@ | Set-Content (Join-Path $WorkspacePath 'Detect.ps1') -Encoding ASCII
        Write-Info "Generated: Detect.ps1 (CustomScript wrapper)"
    } else {
    @'
# Detect.ps1 - generated by Test-RecipeInstallation.ps1
# Exits: 0 = application detected, 1 = not detected
# stdout: JSON result   stderr: diagnostic log
# NOTE: No non-ASCII characters in this file - encoding-safe for PowerShell 5.1 in sandbox
$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    # stderr - captured by the cmd wrapper 2> redirect into detect_after_*.log
    [Console]::Error.WriteLine("$ts $Message")
}

function Get-MSIProductCode {
    param([string]$MsiPath)
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($MsiPath, 0))
        $view = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db,
            @("SELECT Value FROM Property WHERE Property='ProductCode'"))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if ($null -eq $record) { Write-Log "WARNING: MSI Property table returned no ProductCode row"; return $null }
        return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, @(1))
    } catch {
        Write-Log "WARNING: Could not read ProductCode from MSI: $_"
        return $null
    }
}

function Test-MSIProductCodeInstalled {
    param([string]$ProductCode)
    if ([string]::IsNullOrWhiteSpace($ProductCode)) { return $false }
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode"
    )
    return ($paths | Where-Object { Test-Path $_ }).Count -gt 0
}

function Expand-RecipePath {
    param([string]$Path)
    $Path = $Path -replace '%ProgramFiles%',        $env:ProgramFiles
    $Path = $Path -replace '%ProgramFiles\(x86\)%', ${env:ProgramFiles(x86)}
    $Path = $Path -replace '%SystemRoot%',           $env:SystemRoot
    $Path = $Path -replace '%SystemDrive%',          $env:SystemDrive
    $Path = $Path -replace '%WinDir%',               $env:WinDir
    $Path = $Path -replace '%CommonProgramFiles%',   $env:CommonProgramFiles
    $Path = [System.Environment]::ExpandEnvironmentVariables($Path)
    return $Path
}

function Compare-Versions {
    param([string]$Actual, [string]$Expected, [string]$Operator)
    if ([string]::IsNullOrWhiteSpace($Expected) -or $Expected -like '$*') { return $null }
    try {
        $a = [System.Version]$Actual; $e = [System.Version]$Expected
        $cmp = $a.CompareTo($e)
        # Explicit return per case -- 'return switch (...)' is PS 7+ only, not PS 5.1
        switch ($Operator) {
            'Equals'        { return ($cmp -eq 0) }
            'NotEquals'     { return ($cmp -ne 0) }
            'GreaterThan'   { return ($cmp -gt 0) }
            'GreaterEquals' { return ($cmp -ge 0) }
            'LessThan'      { return ($cmp -lt 0) }
            'LessEquals'    { return ($cmp -le 0) }
            default         { return ($cmp -ge 0) }
        }
    } catch { return $null }
}

function Test-DetectionClauses {
    param([array]$Clauses)
    if ($Clauses.Count -eq 0) { return @{ Detected = $null; ClauseResults = @(); Details = 'No detection clauses configured' } }
    $clauseResults = @()
    foreach ($clause in $Clauses) {
        $r = @{ Type = $clause.Type; Detected = $false; Detail = '' }
        switch ($clause.Type) {
            'File' {
                $expandedPath = Expand-RecipePath $clause.FilePath
                $fullPath = Join-Path $expandedPath $clause.FileName
                $exists = Test-Path $fullPath -PathType Leaf
                if (-not $exists) {
                    $r.Detail = "File not found: $fullPath"
                } else {
                    $r.Detected = $true; $r.Detail = "File found: $fullPath"
                    if ($clause.CheckValue -eq 'True' -and $clause.PropertyType -eq 'Version') {
                        $actualVer = (Get-Item $fullPath -ErrorAction SilentlyContinue).VersionInfo.FileVersion
                        $versionOk = Compare-Versions $actualVer $clause.ExpectedValue $clause.Operator
                        if ($null -ne $versionOk) {
                            $r.Detail += " | Version $actualVer $(if ($versionOk) { 'satisfies' } else { 'does NOT satisfy' }) $($clause.Operator) $($clause.ExpectedValue)"
                            $r.Detected = $r.Detected -and $versionOk
                        } else {
                            $r.Detail += " | Version $actualVer (expected value is a placeholder - existence check only)"
                        }
                    }
                }
            }
            'Registry' {
                $regPath = "$($clause.Hive)\$($clause.KeyName)"
                $keyExists = Test-Path $regPath
                if (-not $keyExists) {
                    $r.Detail = "Registry key not found: $regPath"
                } elseif ([string]::IsNullOrWhiteSpace($clause.ValueName)) {
                    $r.Detected = $true; $r.Detail = "Registry key exists: $regPath"
                } else {
                    $actualVal = (Get-ItemProperty $regPath -Name $clause.ValueName -ErrorAction SilentlyContinue).($clause.ValueName)
                    if ($null -eq $actualVal) {
                        $r.Detail = "Registry value not found: $regPath\$($clause.ValueName)"
                    } else {
                        $r.Detected = $true; $r.Detail = "Registry value: $regPath\$($clause.ValueName) = $actualVal"
                        if ($clause.CheckValue -eq 'True' -and $clause.PropertyType -eq 'Version') {
                            $versionOk = Compare-Versions ([string]$actualVal) $clause.ExpectedValue $clause.Operator
                            if ($null -ne $versionOk) {
                                $r.Detail += " | Version $actualVal $(if ($versionOk) { 'satisfies' } else { 'does NOT satisfy' }) $($clause.Operator) $($clause.ExpectedValue)"
                                $r.Detected = $r.Detected -and $versionOk
                            } else {
                                $r.Detail += " | (expected value is a placeholder - existence check only)"
                            }
                        }
                    }
                }
            }
            'MSI' {
                $productCode = $clause.ProductCode
                if ([string]::IsNullOrWhiteSpace($productCode)) {
                    $msiFullPath = "C:\Temp\CMPackagerTest\$($clause.InstallerFile)"
                    Write-Log "Extracting ProductCode from $msiFullPath"
                    $pc = Get-MSIProductCode $msiFullPath
                    $productCode = if ($null -ne $pc) { $pc.Trim() } else { '' }
                }
                if ([string]::IsNullOrWhiteSpace($productCode)) {
                    $r.Detail = "Could not determine MSI ProductCode"
                } else {
                    $r.Detected = Test-MSIProductCodeInstalled $productCode
                    $r.Detail   = "MSI ProductCode $productCode $(if ($r.Detected) { 'found' } else { 'NOT found' }) in Uninstall registry"
                }
            }
        }
        Write-Log "Detection [$($clause.Type)]: $($r.Detail)"
        $clauseResults += $r
    }
    $allDetected = ($clauseResults | Where-Object { -not $_.Detected }).Count -eq 0
    return @{ Detected = $allDetected; ClauseResults = $clauseResults }
}

# ── Load and normalise detection clauses from JSON ───────────────────────────
$clausesRaw = Get-Content 'C:\Temp\CMPackagerTest\detection_clauses.json' -Raw | ConvertFrom-Json
$DetectionClauses = $clausesRaw | ForEach-Object {
    $ht = @{}
    $_.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
    $ht
}

$result = Test-DetectionClauses -Clauses $DetectionClauses
Write-Log "Detection overall: $($result.Detected)"
# Output JSON to stdout - the cmd wrapper captures this with > to C:\TestFiles\detect_after_*.json
$result | ConvertTo-Json -Depth 10
if ($result.Detected) { exit 0 } else { exit 1 }
'@ | Set-Content (Join-Path $WorkspacePath 'Detect.ps1') -Encoding ASCII
        Write-Info "Generated: Detect.ps1 (clause-based)"
    } # end if CustomScript / else

    # WaitMsiEvent.ps1 — static, polls Application event log for MsiInstaller events
    if ($installType -eq 'MSI') {
        @'
param(
    [int[]]$EventIds,
    [int]$TimeoutMinutes = 15
)
$ErrorActionPreference = 'Continue'
$LogPath = 'C:\TestFiles\sandbox.log'
function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    "$ts $Message" | Tee-Object -FilePath $LogPath -Append | Write-Host
}
Write-Log "Waiting for MsiInstaller event (IDs: $($EventIds -join ','))..."
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
while ((Get-Date) -lt $deadline) {
    $events = Get-EventLog -LogName Application -Source MsiInstaller -Newest 20 -ErrorAction SilentlyContinue |
        Where-Object { $_.EventID -in $EventIds }
    if ($events) {
        Write-Log "  MsiInstaller event found (ID $($events[0].EventID)) -- installer transaction complete."
        exit 0
    }
    Start-Sleep -Seconds 5
}
Write-Log "  WARNING: timed out waiting for MsiInstaller event after $TimeoutMinutes minutes."
exit 1
'@ | Set-Content (Join-Path $WorkspacePath 'WaitMsiEvent.ps1') -Encoding UTF8
        Write-Info "Generated: WaitMsiEvent.ps1"
    }

    # install.cmd — copies installer to a fully local path (msiexec/SYSTEM cannot
    # read from the WSB mapped folder share), then runs the install command
    $installCmdBatch = $installProgram -replace '`"', '"'
    @"
@echo off
mkdir C:\Temp\CMPackagerTest 2>nul
xcopy /Y "C:\TestFiles\$installerFileName" "C:\Temp\CMPackagerTest\" >nul 2>&1
cd /d C:\Temp\CMPackagerTest
echo [%TIME%] Running install: $installCmdBatch >>C:\TestFiles\sandbox.log
$installCmdBatch
set EXIT=%ERRORLEVEL%
if exist install.log copy /Y install.log "C:\TestFiles\install.log" >nul 2>&1
echo [%TIME%] Install exit code: %EXIT% >>C:\TestFiles\sandbox.log
echo %EXIT%> "C:\TestFiles\install.exitcode"
exit /b %EXIT%
"@ | Set-Content (Join-Path $WorkspacePath 'install.cmd') -Encoding ASCII
    Write-Info "Generated: install.cmd"

    # uninstall.cmd — runs the uninstall command in the same working directory
    if (-not [string]::IsNullOrWhiteSpace($uninstallCmd)) {
        $uninstallCmdBatch = $uninstallCmd -replace '`"', '"'
        @"
@echo off
cd /d C:\Temp\CMPackagerTest
echo [%TIME%] Running uninstall: $uninstallCmdBatch >>C:\TestFiles\sandbox.log
$uninstallCmdBatch
set EXIT=%ERRORLEVEL%
if exist uninstall.log copy /Y uninstall.log "C:\TestFiles\uninstall.log" >nul 2>&1
echo [%TIME%] Uninstall exit code: %EXIT% >>C:\TestFiles\sandbox.log
echo %EXIT%> "C:\TestFiles\uninstall.exitcode"
exit /b %EXIT%
"@ | Set-Content (Join-Path $WorkspacePath 'uninstall.cmd') -Encoding ASCII
        Write-Info "Generated: uninstall.cmd"
    }

    # Detect wrapper cmds — PowerShell cannot write to ANY path (local OR mapped folder)
    # from SYSTEM context via wsb exec. Detect.ps1 outputs JSON to stdout and diagnostics
    # to stderr; cmd.exe captures both directly to C:\TestFiles\ via > and 2> redirects.
    foreach ($detectStep in @('install', 'uninstall')) {
        # Detect.ps1 outputs JSON to stdout and diagnostic log lines to stderr.
        # cmd.exe captures stdout directly to C:\TestFiles\detect_after_*.json with >
        # and stderr to C:\TestFiles\detect_after_*.log with 2>.
        # This works because cmd.exe CAN write to the mapped folder; PowerShell cannot.
        @"
@echo off
echo [%TIME%] detect_after_$detectStep.cmd starting >>C:\TestFiles\sandbox.log
if not exist "C:\Temp\CMPackagerTest\Detect.ps1" (
  echo [%TIME%] ERROR: Detect.ps1 not found at C:\Temp\CMPackagerTest\ >>C:\TestFiles\sandbox.log
  exit /b 2
)
powershell.exe -ExecutionPolicy Bypass -NonInteractive -File C:\Temp\CMPackagerTest\Detect.ps1 > "C:\TestFiles\detect_after_$detectStep.json" 2>"C:\TestFiles\detect_after_$detectStep.log"
set EXIT=%ERRORLEVEL%
echo [%TIME%] Detect.ps1 exit code: %EXIT% >>C:\TestFiles\sandbox.log
exit /b %EXIT%
"@ | Set-Content (Join-Path $WorkspacePath "detect_after_$detectStep.cmd") -Encoding ASCII
    }
    Write-Info "Generated: detect_after_install.cmd, detect_after_uninstall.cmd"

    # .wsb config — mapped folder only, no LogonCommand (orchestrated via wsb exec)
    $wsbContent = @"
<Configuration>
  <MemoryInMB>8192</MemoryInMB>
  <Networking>Disable</Networking>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$WorkspacePath</HostFolder>
      <SandboxFolder>C:\TestFiles</SandboxFolder>
      <ReadOnly>false</ReadOnly>
    </MappedFolder>
  </MappedFolders>
</Configuration>
"@
    $wsbPath = Join-Path $WorkspacePath 'SandboxTest.wsb'
    $wsbContent | Set-Content -Path $wsbPath -Encoding UTF8
    Write-Info "Generated: SandboxTest.wsb (no LogonCommand)"

} else {

    # ── Legacy path: full inner RunTest.ps1 + LogonCommand ───────────────────────
    # Render detection clauses as a PowerShell array literal for embedding
    $detectionArrayLiteral = if ($detectionClauseLiterals.Count -gt 0) {
        "@(`n    " + ($detectionClauseLiterals -join ",`n    ") + "`n)"
    } else { '@()' }

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
        [datetime]`$After,
        [int[]]`$EventIds,
        [int]`$TimeoutMinutes = 20
    )
    Write-Log "Waiting for MsiInstaller event (IDs: `$(`$EventIds -join ','))..."
    `$deadline = (Get-Date).AddMinutes(`$TimeoutMinutes)
    while ((Get-Date) -lt `$deadline) {
        `$events = Get-EventLog -LogName Application -Source MsiInstaller -Newest 20 -ErrorAction SilentlyContinue |
            Where-Object { `$_.TimeGenerated -gt `$After -and `$_.EventID -in `$EventIds }
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
    Wait-MsiInstallerEvent -After `$installStart -EventIds @(1033)
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
        Wait-MsiInstallerEvent -After `$uninstallStart -EventIds @(1034)
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

    $wsbContent = @"
<Configuration>
  <MemoryInMB>8192</MemoryInMB>
  <Networking>Disable</Networking>
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
    Write-Info "Generated: SandboxTest.wsb (with LogonCommand)"

} # end if ($useWsbCli) / else

#endregion

#region ── Launch Sandbox and Wait ───────────────────────────────────────────────

$resultsFile = Join-Path $WorkspacePath 'results.json'
$sandboxLog  = Join-Path $WorkspacePath 'sandbox.log'

Write-Host "`n" -NoNewline
Write-Step "Launching Windows Sandbox..."
if ($useWsbCli) {
    Write-Info "Running headless via wsb.exe — no sandbox window will appear."
}
Write-Info "Timeout: $TimeoutMinutes minutes"
Write-Host ""

$deadline    = (Get-Date).AddMinutes($TimeoutMinutes)
$dotCount    = 0
$sandboxId   = $null
$exitedEarly = $false
$timedOut    = $false

if ($useWsbCli) {
    # ── wsb.exe path (Windows 11 24H2+): exec-based orchestration ────────────────
    # Each test step is driven from the host via wsb exec -r System.
    # Results are collected directly — no polling for a results file.
    # All scripts run in System context, matching ConfigMgr deployment behaviour.

    # ── Helpers ───────────────────────────────────────────────────────────────────

    # Pre-initialise result variables so Write-TimeoutResult can always serialise them.
    $installExitCode   = $null;  $installSuccess   = $null
    $detAfterInstall   = [ordered]@{ Detected = $null; ClauseResults = @() }
    $uninstallExitCode = $null;  $uninstallSuccess = $null
    $detAfterUninstall = [ordered]@{ Detected = $null; ClauseResults = @() }

    # Invoke-SandboxExec — runs a wsb exec step and enforces the per-test deadline.
    # Uses Start-Job + Wait-Job so the & operator handles argument quoting correctly
    # (Start-Process -ArgumentList splits multi-word --command values on spaces).
    # Returns the process exit code, or -999 when the deadline is reached.
    function Invoke-SandboxExec {
        param([string]$StepLabel, [string]$Command)
        $remaining = [int]([datetime]$deadline - (Get-Date)).TotalSeconds
        if ($remaining -le 5) {
            Write-Host "  $(Get-ElapsedPrefix) [TIMEOUT] Deadline reached before '$StepLabel'." -ForegroundColor Yellow
            return -999
        }
        # Capture variables for the job scope.
        $wsbP = $wsbCliPath; $sbId = $sandboxId
        $job = Start-Job -ScriptBlock {
            param($p, $id, $cmd)
            $null = & $p exec --id $id -c $cmd -r System 2>&1
            $LASTEXITCODE
        } -ArgumentList $wsbP, $sbId, $Command

        $done = $job | Wait-Job -Timeout $remaining
        if ($done) {
            $rc = [int](($job | Receive-Job) | Select-Object -Last 1)
            $job | Remove-Job -Force
            return $rc
        }
        $job | Stop-Job; $job | Remove-Job -Force
        Write-Host "  $(Get-ElapsedPrefix) [TIMEOUT] '$StepLabel' exceeded the per-test deadline." -ForegroundColor Yellow
        return -999
    }

    # Write-TimeoutResult — persists a TIMEOUT result and stops the sandbox.
    function Write-TimeoutResult {
        param([string]$Reason)
        Write-Host "  $(Get-ElapsedPrefix) [TIMEOUT] $Reason" -ForegroundColor Yellow
        "$((Get-Date -Format 'HH:mm:ss')) TIMEOUT: $Reason" | Add-Content $sandboxLog
        [ordered]@{
            Application             = $appName
            DeploymentType          = $depTypeName
            Timestamp               = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            InstallExitCode         = $installExitCode
            InstallSuccess          = $installSuccess
            DetectionAfterInstall   = $detAfterInstall
            UninstallExitCode       = $uninstallExitCode
            UninstallSuccess        = $uninstallSuccess
            DetectionAfterUninstall = $detAfterUninstall
            OverallResult           = 'TIMEOUT'
            Notes                   = @($Reason)
        } | ConvertTo-Json -Depth 10 | Set-Content $resultsFile -Encoding UTF8
        & $wsbCliPath stop --id $sandboxId 2>&1 | Out-Null
        exit 1
    }

    # ── Stop any stale sandboxes ──────────────────────────────────────────────────
    $staleIds = & $wsbCliPath list 2>&1 | Where-Object { $_ -match '^[0-9a-fA-F]{8}-' }
    foreach ($staleId in $staleIds) {
        Write-Info "Stopping stale sandbox $staleId before launch..."
        & $wsbCliPath stop --id $staleId 2>&1 | Out-Null
    }
    if ($staleIds) { Start-Sleep -Seconds 3 }

    # ── Launch ───────────────────────────────────────────────────────────────────
    # wsb start --config takes the WSB XML content as an inline string, not a file path.
    # Collapse to a single line — some builds of wsb.exe reject multi-line arguments.
    # wsb start outputs:  "Windows Sandbox environment started successfully:\nId: <guid>"
    $wsbContentOneLine = $wsbContent -replace '\r?\n\s*', ''
    $startOutput = & $wsbCliPath start --config $wsbContentOneLine 2>&1
    $sandboxId   = ($startOutput | Where-Object { $_ -match '^Id:\s' }) -replace '^Id:\s*'

    if ($sandboxId) {
        Write-Info "Sandbox ID: $sandboxId"
    } else {
        Write-Host "`n  [ERROR] Failed to start sandbox." -ForegroundColor Red
        Write-Host "          wsb output: $($startOutput -join ' ')" -ForegroundColor Red
        exit 1
    }

    # ── Wait for sandbox to be ready ─────────────────────────────────────────────
    # Poll until wsb exec succeeds; System context must be available for testing.
    Write-Host "  Waiting for sandbox to initialize" -NoNewline -ForegroundColor Cyan
    $readyDeadline = (Get-Date).AddMinutes(5)
    $sandboxReady  = $false
    while (-not $sandboxReady -and (Get-Date) -lt $readyDeadline) {
        Start-Sleep -Seconds 5
        $null = & $wsbCliPath exec --id $sandboxId -c 'cmd /c exit 0' -r System 2>&1
        if ($LASTEXITCODE -eq 0) { $sandboxReady = $true }
        else { Write-Host '.' -NoNewline -ForegroundColor Cyan }
    }
    Write-Host ""
    if (-not $sandboxReady) {
        Write-Host "`n  [ERROR] Sandbox did not become ready within 5 minutes." -ForegroundColor Red
        & $wsbCliPath stop --id $sandboxId 2>&1 | Out-Null
        exit 1
    }
    Write-Info "Sandbox ready. Starting test sequence."
    "$((Get-Date -Format 'HH:mm:ss')) === Test starting: $appName / $depTypeName ===" |
        Set-Content $sandboxLog -Encoding UTF8

    # ── Sandbox setup: disable CI safety checks that stall MSI installs on 24H2 ──
    Write-Step "Applying sandbox CI policy fix (disables MSI install stalling on 24H2)"
    $null = Invoke-SandboxExec 'Sandbox setup' 'C:\TestFiles\sandbox_setup.cmd'

    # ── Step 1: Install ──────────────────────────────────────────────────────────
    Write-Step "Step 1: Install"
    "$((Get-Date -Format 'HH:mm:ss')) --- Step 1: Install ---" | Add-Content $sandboxLog
    $null = Invoke-SandboxExec 'Step 1: Install' 'C:\TestFiles\install.cmd'
    # wsb exec always returns 0 — read actual exit code from the file written by install.cmd.
    $installExitCodeFile = Join-Path $WorkspacePath 'install.exitcode'
    if (Test-Path $installExitCodeFile) {
        $installExitCode = [int]((Get-Content $installExitCodeFile -Raw).Trim())
    } else {
        Write-TimeoutResult 'Timed out during Step 1: Install (no install.exitcode produced)'
    }
    $installSuccess = $installExitCode -in @(0, 3010, 1641)
    Write-Info "Install exit code: $installExitCode ($(if ($installSuccess) { 'SUCCESS' } else { 'FAILURE' }))"

    # MSI only: wsb exec is synchronous and msiexec /qn is synchronous, so the
    # registry is committed before exec returns. A brief sleep ensures any
    # transient caches flush before detection. (WaitMsiEvent was removed because
    # MsiInstaller event 1033/1034 does not appear reliably in the sandbox event log.)
    if ($installType -eq 'MSI') { Start-Sleep -Seconds 10 }

    # ── Step 2: Detect after install ─────────────────────────────────────────────
    Write-Step "Step 2: Detection after install"
    "$((Get-Date -Format 'HH:mm:ss')) --- Step 2: Detection after install ---" | Add-Content $sandboxLog
    $detectInstallRc  = Invoke-SandboxExec 'Step 2: Detect after install' 'C:\TestFiles\detect_after_install.cmd'
    if ($detectInstallRc -eq -999) { Write-TimeoutResult 'Timed out during Step 2: Detection after install' }

    # wsb exec returns 0 regardless of the command's exit code — use the JSON file as
    # the source of truth.  A missing file means Detect.ps1 did not produce output.
    $detAfterInstall = [ordered]@{ Detected = $null; ClauseResults = @() }
    $detAfterInstallJson = Join-Path $WorkspacePath 'detect_after_install.json'
    if (Test-Path $detAfterInstallJson) {
        try { $detAfterInstall = Get-Content $detAfterInstallJson -Raw | ConvertFrom-Json } catch {}
    } else {
        Write-Info "detect_after_install.json not produced (wsb exec rc: $detectInstallRc) — check sandbox.log"
    }
    Write-Info "Detected after install: $($detAfterInstall.Detected)"

    # ── Steps 3 & 4: Uninstall + detect after uninstall ─────────────────────────
    # Only run if install succeeded AND the app was detected after install.
    # If either precondition fails the test is already a FAIL; skipping uninstall
    # avoids side-effects on a partially installed app and keeps the result clear.
    $uninstallExitCode = $null
    $uninstallSuccess  = $null
    $detAfterUninstall = [ordered]@{ Detected = $null; ClauseResults = @() }

    if ($installSuccess -ne $true -or $detAfterInstall.Detected -ne $true) {
        $skipReason = if ($installSuccess -ne $true) {
            "install failed (exit code $installExitCode)"
        } else {
            "app not detected after install"
        }
        Write-Info "Skipping Steps 3 & 4 — $skipReason."
        "$((Get-Date -Format 'HH:mm:ss')) --- Steps 3 & 4 skipped: $skipReason ---" | Add-Content $sandboxLog
    } else {
        # ── Step 3: Uninstall ────────────────────────────────────────────────────
        if ([string]::IsNullOrWhiteSpace($uninstallCmd)) {
            Write-Info "No uninstall command — skipping uninstall step."
            "$((Get-Date -Format 'HH:mm:ss')) --- Step 3: Uninstall skipped (no command) ---" | Add-Content $sandboxLog
        } else {
            Write-Step "Step 3: Uninstall"
            "$((Get-Date -Format 'HH:mm:ss')) --- Step 3: Uninstall ---" | Add-Content $sandboxLog
            $null = Invoke-SandboxExec 'Step 3: Uninstall' 'C:\TestFiles\uninstall.cmd'
            # wsb exec always returns 0 — read actual exit code from the file written by uninstall.cmd.
            $uninstallExitCodeFile = Join-Path $WorkspacePath 'uninstall.exitcode'
            if (Test-Path $uninstallExitCodeFile) {
                $uninstallExitCode = [int]((Get-Content $uninstallExitCodeFile -Raw).Trim())
            } else {
                Write-TimeoutResult 'Timed out during Step 3: Uninstall (no uninstall.exitcode produced)'
            }
            $uninstallSuccess = $uninstallExitCode -in @(0, 3010, 1641)
            Write-Info "Uninstall exit code: $uninstallExitCode ($(if ($uninstallSuccess) { 'SUCCESS' } else { 'FAILURE' }))"

            # Brief pause for the Windows Installer service to commit registry cleanup.
            # Manual testing shows the uninstall key vanishes immediately; 5 s is ample.
            if ($installType -eq 'MSI') { Start-Sleep -Seconds 5 }
        }

        # ── Step 4: Detect after uninstall ───────────────────────────────────────
        Write-Step "Step 4: Detection after uninstall"
        "$((Get-Date -Format 'HH:mm:ss')) --- Step 4: Detection after uninstall ---" | Add-Content $sandboxLog
        $detectUninstallRc  = Invoke-SandboxExec 'Step 4: Detect after uninstall' 'C:\TestFiles\detect_after_uninstall.cmd'
        if ($detectUninstallRc -eq -999) { Write-TimeoutResult 'Timed out during Step 4: Detection after uninstall' }

        $detAfterUninstallJson = Join-Path $WorkspacePath 'detect_after_uninstall.json'
        if (Test-Path $detAfterUninstallJson) {
            try { $detAfterUninstall = Get-Content $detAfterUninstallJson -Raw | ConvertFrom-Json } catch {}
        } else {
            Write-Info "detect_after_uninstall.json not produced (wsb exec rc: $detectUninstallRc) — check sandbox.log"
        }
        Write-Info "Detected after uninstall: $($detAfterUninstall.Detected)"
    }

    # ── Step 5: Compute overall result and write results.json ────────────────────
    $installOk   = $installSuccess -eq $true
    $detAfterOk  = $detAfterInstall.Detected   -eq $true
    $detAfterUn  = $detAfterUninstall.Detected -eq $false   # NOT detected = good
    $uninstallOk = ($uninstallSuccess -eq $true) -or ($null -eq $uninstallSuccess)

    $overallResult = if ($installOk -and $detAfterOk -and $uninstallOk -and $detAfterUn) {
        'PASS'
    } elseif ($installOk -and $detAfterOk -and [string]::IsNullOrWhiteSpace($uninstallCmd)) {
        'PASS (install only — no uninstall configured)'
    } else {
        'FAIL'
    }
    Write-Info "Overall result: $overallResult"
    "$((Get-Date -Format 'HH:mm:ss')) === Overall result: $overallResult ===" | Add-Content $sandboxLog

    [ordered]@{
        Application             = $appName
        DeploymentType          = $depTypeName
        Timestamp               = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        InstallExitCode         = $installExitCode
        InstallSuccess          = $installSuccess
        DetectionAfterInstall   = $detAfterInstall
        UninstallExitCode       = $uninstallExitCode
        UninstallSuccess        = $uninstallSuccess
        DetectionAfterUninstall = $detAfterUninstall
        OverallResult           = $overallResult
        Notes                   = @()
    } | ConvertTo-Json -Depth 10 | Set-Content $resultsFile -Encoding UTF8

    # Stop the sandbox — test is complete
    & $wsbCliPath stop --id $sandboxId 2>&1 | Out-Null

} else {
    # ── Legacy WindowsSandbox.exe path ───────────────────────────────────────────
    # Wait up to 30 s for any lingering sandbox from a previous test to fully exit
    # before attempting to launch a new one (only one instance is allowed).
    $preWaitEnd    = (Get-Date).AddSeconds(30)
    $preWaitLogged = $false
    while ((Get-Date) -lt $preWaitEnd) {
        $stale = Get-Process -Name 'WindowsSandbox', 'WindowsSandboxClient' -ErrorAction SilentlyContinue
        if (-not $stale) { break }
        if (-not $preWaitLogged) {
            Write-Info "Waiting for previous sandbox to finish shutting down..."
            $preWaitLogged = $true
        }
        Start-Sleep -Seconds 3
    }

    Start-Process -FilePath $wsbLaunchPath -ArgumentList $wsbPath

    # Poll purely on the results file + timeout.
    # The legacy launcher exits immediately and WindowsSandboxClient can take time
    # to appear; reliable process-death detection is not possible here, so we simply
    # wait for the results file or the deadline.
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

    # Force-close any remaining sandbox processes so the next test can start cleanly
    Get-Process -Name 'WindowsSandbox', 'WindowsSandboxClient' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path $resultsFile)) { $timedOut = $true }

if ($exitedEarly) {
    Write-Host "`n══════════════════════════════════════" -ForegroundColor White
    Write-Host "  SANDBOX CLOSED UNEXPECTEDLY" -ForegroundColor Red
    Write-Host "══════════════════════════════════════`n" -ForegroundColor White
    Write-Host "  The sandbox exited before the test completed." -ForegroundColor Yellow
    Write-Host "  Check the log for the last known state:`n" -ForegroundColor Yellow
} elseif ($timedOut) {
    Write-Host "`n══════════════════════════════════════" -ForegroundColor White
    Write-Host "  TIMED OUT after $TimeoutMinutes minutes" -ForegroundColor Red
    Write-Host "══════════════════════════════════════`n" -ForegroundColor White
    Write-Host "  The sandbox did not produce a results file within the allowed time." -ForegroundColor Yellow
    Write-Host "  Check the log for the last known state:`n" -ForegroundColor Yellow
}

if ($exitedEarly -or $timedOut) {
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
if ($null -eq $results.UninstallExitCode -and $null -eq $results.UninstallSuccess) {
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
foreach ($logName in @('install.log', 'uninstall.log', 'detect_after_install.log', 'detect_after_uninstall.log')) {
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
