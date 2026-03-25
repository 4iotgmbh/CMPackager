# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CMPackager is a PowerShell automation tool for SCCM/MEM ConfigMgr that downloads, packages, distributes, and deploys applications using XML "recipe" files. The main script is `CMPackager.ps1` (~2,173 lines).

## Running the Script

```powershell
# Standard run (processes all recipes in Recipes/ folder)
.\CMPackager.ps1

# Custom preference file
.\CMPackager.ps1 -PreferenceFile "path\to\CMPackager.prefs"

# Run a single recipe (dynamic parameter, tab-completable from Recipes/)
.\CMPackager.ps1 -SingleRecipe AdobeReader.xml

# Setup mode (WPF GUI for initial configuration)
.\CMPackager.ps1 -Setup
```

**Prerequisites**: SCCM 1906+ with ConfigMgr Console installed, PowerShell 3.0+.

**Configuration**: Copy `CMPackager.prefs.template` to `CMPackager.prefs` and edit the XML settings. The prefs file is gitignored.

There are no automated tests or linting configured for this project.

## Architecture

### Main Execution Pipeline

The script processes each recipe XML file through a sequential pipeline (see MAIN section ~line 2070):

1. **`Start-ApplicationDownload`** — Downloads the app (via URL or PrefetchScript in recipe)
2. **`Invoke-ApplicationCreation`** — Creates the application in SCCM
3. **`Add-DeploymentType`** — Adds deployment types (MSI, Script, etc.) with detection methods
4. **`Invoke-ApplicationDistribution`** — Distributes content to distribution points/groups
5. **`Invoke-ApplicationSupersedence`** — Sets up supersedence relationships with previous versions
6. **`Invoke-ApplicationDeployment`** — Deploys to SCCM collections
7. **`Invoke-ApplicationCleanup`** — Cleans up temporary files

### Recipe System

Recipes are XML files following the schema in `Disabled/_RecipeSchema.xsd`. Each recipe defines:

- **`<Application>`** — App metadata (name, publisher, icon, etc.)
- **`<Downloads>`** — Download sources with version checking; each `<Download>` is linked to a DeploymentType by the `DeploymentType` attribute
- **`<DeploymentTypes>`** — Install commands, detection methods (Registry, File, MSI), requirements rules, dependencies
- **`<Distribution>`** — Content distribution settings
- **`<Supersedence>`** — Version supersedence behavior
- **`<Deployment>`** — Collection targeting and scheduling

Active recipes go in `Recipes/`. Disabled/template recipes are in `Disabled/` (~98 examples).

### WinGet Integration for Recipe Creation

CMPackager includes tools to leverage the Windows Package Manager (WinGet) repository for recipe creation:

**Get-WingetInfo.ps1 Script** (`ExtraFiles/Scripts/Get-WingetInfo.ps1`):

- PowerShell 7 utility for researching applications when creating new recipes
- Searches the WinGet repository and provides interactive selection via `Out-ConsoleGridView`
- Retrieves comprehensive package details from WinGet manifests (YAML files on GitHub)
- Returns: installer URLs, version info, publisher, description, silent install switches, product codes, architecture, etc.
- Requires modules: `Microsoft.WinGet.Client`, `Microsoft.PowerShell.ConsoleGuiTools`, `powershell-yaml`, `cobalt`
- Usage: `.\Get-WingetInfo.ps1 -ApplicationName "Adobe Reader", "7-Zip" -Output List`

**New-ScaffoldRecipe.ps1 Script** (`ExtraFiles/Scripts/New-ScaffoldRecipe.ps1`):

- Accepts pipeline input from Get-WingetInfo.ps1 to automatically scaffold new recipe files
- Selects the appropriate template (`_MSIRecipeTemplate.xml` or `_EXERecipeTemplate.xml`) based on installer type
- Creates recipe file named after the application (spaces removed)
- Populates XML fields: Name, Description, Publisher, UserDocumentation, PrefetchScript, install commands, etc.
- Outputs to `Recipes/` folder by default (customizable via `-OutputPath` parameter)
- Usage: `.\Get-WingetInfo.ps1 -ApplicationName "7-Zip" | .\New-ScaffoldRecipe.ps1`
- Note: Generated recipes require customization (detection methods, icons, testing) before deployment

**Dynamic URL Functions in CMPackager.ps1**:

- **`Get-InstallerURLfromWinget`** ([CMPackager.ps1:301](CMPackager.ps1#L301)) — Queries GitHub API for WinGet package manifests, parses the latest version's installer YAML, and returns the installer download URL. Parameters: `-apiUrl` (required), `-InstallerType` (`msi`, `exe`, or `zip`, required), `-Architecture` (`x64`, `x86`, `arm64`, `arm`, optional), `-Scope` (`machine` or `user`, optional). When multiple installers exist in the manifest, Architecture and Scope narrow the selection. For `-InstallerType zip`, returns a `PSCustomObject` with `.Url` (the zip download URL) and `.NestedFilePath` (the `RelativeFilePath` of the nested installer inside the zip); for `msi`/`exe`, returns a plain string URL.

These functions are called within recipe `<PrefetchScript>` blocks to dynamically fetch current download URLs at runtime:

```powershell
# Example PrefetchScript usage in a recipe:
$apiUrl = "https://api.github.com/repos/microsoft/winget-pkgs/contents/manifests/a/Adobe/Acrobat/Reader/64-bit"
$DownloadURL = Get-InstallerURLfromWinget -apiUrl $apiUrl -InstallerType msi -Architecture x64 -Scope machine
```

**Complete Recipe Creation Workflow**:

```powershell
# Step 1: Research the application and gather WinGet information
cd ExtraFiles\Scripts
.\Get-WingetInfo.ps1 -ApplicationName "7-Zip"

# Step 2: Pipe the output to New-ScaffoldRecipe.ps1 to create a new recipe
.\Get-WingetInfo.ps1 -ApplicationName "7-Zip" | .\New-ScaffoldRecipe.ps1

# Step 3: Customize the generated recipe (add icon, configure detection method, test)

# Step 4: Run CMPackager with the new recipe
cd ..\..
.\CMPackager.ps1 -SingleRecipe 7-Zip.xml
```

### Key Supporting Functions

- **`Add-DetectionMethodClause`** — Builds SCCM detection logic (Registry, File, MSI clauses)
- **`Add-RequirementsRule`** — Adds OS, global condition, or custom requirements
- **`Copy-CMDeploymentTypeRule`** — Copies requirement rules between deployment types (adapted from external source)
- **`Connect-ConfigMgr`** — Establishes SCCM site connection and imports the ConfigurationManager module
- **`Add-LogContent`** — Centralized logging with rotation (`$Global:LogPath`)
- **`Get-MSIInfo`** / **`Get-MSISourceFileVersion`** — Extract metadata from MSI files using COM
- **`Get-InstallerURLfromWinget`** — Queries WinGet package manifests from GitHub to retrieve the current installer download URL; used in recipe PrefetchScripts with `-InstallerType msi|exe` and optional `-Architecture`/`-Scope` to select among multiple installers

### Global State

All shared state uses `$Global:` prefix variables initialized from the prefs file: `$Global:CMSite`, `$Global:SiteServer`, `$Global:ContentLocationRoot`, `$Global:TempDir`, `$Global:LogPath`, etc.

## Coding Conventions

- **Function naming**: PowerShell verb-noun PascalCase (`Invoke-ApplicationCreation`, `Add-DeploymentType`)
- **Parameter validation**: Uses `ValidateScript`, `ValidateSet`, `ValidateNotNullOrEmpty`
- **Error handling**: Try-catch blocks with `Add-LogContent` for detailed logging
- **Dynamic parameters**: Used for the `-SingleRecipe` parameter (populated from Recipes/ folder at runtime)
- **XML entity encoding**: PrefetchScript blocks in recipes use `&amp;` for `&`, etc.

## Recipe Testing Infrastructure

Two scripts in `ExtraFiles/Scripts/` implement automated install/uninstall testing of recipes in Windows Sandbox:

- **`Test-RecipeInstallation.ps1`** — single-recipe test orchestrator
- **`Test-RecipeBatch.ps1`** — batch wrapper; iterates a recipes folder, collects CSV results

**Purpose**: Validate that recipe install/uninstall commands and detection methods actually work before deploying via SCCM.

### Running tests

```powershell
# Test a single recipe
.\ExtraFiles\Scripts\Test-RecipeInstallation.ps1 -RecipePath .\Disabled\7-Zip.xml

# Batch test all recipes in a folder
.\ExtraFiles\Scripts\Test-RecipeBatch.ps1 -RecipesPath .\Disabled -Filter '7-Zip.xml'
```

**Prerequisites**: Windows 11 with Windows Sandbox enabled, PowerShell 7+. Must run on Windows.

### Dual-mode sandbox control

The script auto-detects the sandbox launcher at startup:

- **wsb.exe path** (Windows 11 24H2+): detected via `Get-Command 'wsb.exe'` (PATH search — never hardcode `System32\wsb.exe`). Runs headlessly; each step is driven from the host via `wsb exec`.
- **Legacy path** (older Windows): falls back to `WindowsSandbox.exe` with a `<LogonCommand>` that runs a self-contained `RunTest.ps1` inside the sandbox.

### wsb.exe CLI (confirmed behaviour)

- `wsb start --config <xml-string>` — XML content inline, not a file path. Must be single-line; collapse with `-replace '\r?\n\s*', ''` before passing.
- `wsb start` outputs `"Windows Sandbox environment started successfully:\nId: <guid>"` on success.
- `wsb list` — bare GUIDs, one per line. Match running sandboxes with `'^[0-9a-fA-F]{8}-'`.
- `wsb stop --id <guid>` — targeted teardown.
- `wsb exec --id <guid> --command <cmd> --run-as System` — use long-form flags `--command` and `--run-as` (short forms unconfirmed).
- **wsb exec always returns exit code 0** regardless of the child process exit code. Never use the wsb exec return value as success/failure signal — use output files or JSON instead.
- `Write-Error` with `$ErrorActionPreference = 'Stop'` throws a terminating exception caught silently by the batch caller. Use `Write-Host ... -ForegroundColor Red` + `exit 1` for fatal errors instead.

### CRITICAL: PowerShell cannot write from SYSTEM context (wsb exec)

PowerShell running as SYSTEM via `wsb exec --run-as System` **cannot write to any path** — not to the WSB mapped folder (`C:\TestFiles\`) and not to local paths like `C:\Temp\` either. `Set-Content`, `Add-Content`, `Copy-Item` all fail silently. PowerShell CAN read local files fine.

`cmd.exe` running as SYSTEM via `wsb exec` CAN write to both mapped and local paths.

**Solution — use stdout/stderr instead of file I/O:**
- Detect.ps1 outputs JSON result to stdout (`$result | ConvertTo-Json -Depth 10`), diagnostics to stderr (`[Console]::Error.WriteLine()`)
- `detect_after_*.cmd` captures: `powershell.exe ... > "C:\TestFiles\detect.json" 2>"C:\TestFiles\detect.log"`
- install/uninstall progress: `echo [%TIME%] ... >>C:\TestFiles\sandbox.log` (cmd.exe echo)
- sandbox_setup.cmd copies files using cmd.exe `copy`, not PowerShell `Copy-Item`

### Windows 11 24H2 CI policy fix

MSI installs stall ~2 minutes on 24H2 sandbox due to code-integrity safety checks. Applied as the very first exec step in `sandbox_setup.cmd`:

```batch
REG add HKLM\SYSTEM\CurrentControlSet\Control\CI\Policy /v VerifiedAndReputablePolicyState /t REG_DWORD /d 0 /f
CiTool -r <NUL
```

Drops install time from ~2 minutes to ~2 seconds. Reference: https://github.com/microsoft/Windows-Sandbox/issues/68

### MSI ProductCode extraction

`New-Object -ComObject WindowsInstaller.Installer` hangs ~60 seconds in SYSTEM context inside the sandbox. **Always extract the MSI ProductCode on the HOST** before sandbox launch using `Get-MsiProductCodeFromFile` (defined in the Helpers region of `Test-RecipeInstallation.ps1`), then embed it directly into `detection_clauses.json`.

### Detect.ps1 authoring rules (PS 5.1 in sandbox)

The sandbox runs PowerShell 5.1 (Windows built-in). Two rules must be followed when editing the Detect.ps1 template in `Test-RecipeInstallation.ps1`:

1. **No non-ASCII characters** — Em dashes (`—`) and other Unicode cause encoding corruption when PS 5.1 reads the file (UTF-8 bytes E2 80 94 are misread as Windows-1252, where byte 0x94 is a RIGHT DOUBLE QUOTATION MARK — a string delimiter in PS 5.1, breaking string literals). Save Detect.ps1 with `-Encoding ASCII` and use plain hyphens.

2. **No `return switch`** — `return switch (...) { ... }` is PowerShell 7+ syntax. In PS 5.1 use explicit `return` inside each switch arm:
   ```powershell
   # Wrong (PS 7+ only):
   return switch ($op) { 'Equals' { $cmp -eq 0 } }
   # Correct (PS 5.1 compatible):
   switch ($op) { 'Equals' { return ($cmp -eq 0) } }
   ```

### Test sequence (wsb.exe path)

1. Stop any stale sandboxes via `wsb list` + `wsb stop`
2. `wsb start --config $wsbContentOneLine` → capture sandbox ID
3. Poll `wsb exec --command 'cmd /c exit 0' --run-as System` until exit 0 (sandbox ready)
4. `wsb exec ... sandbox_setup.cmd` → apply CI fix, copy scripts to local drive (`C:\Temp\CMPackagerTest\`)
5. `wsb exec ... install.cmd` → install (log copied back to `C:\TestFiles\`)
6. `Start-Sleep 10` for MSI (registry committed; MsiInstaller event log unreliable in sandbox)
7. `wsb exec ... detect_after_install.cmd` → Detect.ps1 stdout captured by cmd.exe `>` to `C:\TestFiles\detect_after_install.json`
8. If install succeeded AND app detected: run Steps 3 & 4 (uninstall + detect after uninstall); otherwise skip and record FAIL
9. `wsb exec ... uninstall.cmd` → uninstall (log copied back)
10. `Start-Sleep 5` for MSI
11. `wsb exec ... detect_after_uninstall.cmd` → same pattern as step 7
12. Build `results.json` on host from JSON files; `wsb stop --id $sandboxId`

Note: WaitMsiEvent was removed — MsiInstaller event 1033/1034 does not appear reliably in the sandbox event log. Fixed-duration sleeps are used instead. Also do not filter MsiInstaller events by product name — the sandbox runs exactly one install so any matching event ID is ours.

### Sandbox configuration

```xml
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
```

### Artifacts

Generated before launch: `detection_clauses.json`, `Detect.ps1`, `install.cmd`, `uninstall.cmd`, `detect_after_install.cmd`, `detect_after_uninstall.cmd`, `sandbox_setup.cmd`, `SandboxTest.wsb`

Produced by test run: `results.json`, `sandbox.log`, `install.log`, `uninstall.log`, `detect_after_install.json`, `detect_after_install.log`, `detect_after_uninstall.json`, `detect_after_uninstall.log`

### Batch CSV columns

`Timestamp, Application, RecipeFile, DeploymentType, Result, DurationMinutes, InstallExitCode, InstallSuccess, DetectionAfterInstall, UninstallExitCode, UninstallSuccess, DetectionAfterUninstall, Notes, WorkspacePath`

Result values: `PASS`, `PASS (install only — no uninstall configured)`, `FAIL`, `TIMEOUT`, `SKIPPED`, `ERROR`

### Known limitations / TODOs

- Custom/CustomScript detection methods are not yet supported — recipes using them are skipped with exit 2 (shows as SKIPPED in CSV)

## File Layout

| Path | Purpose |
| ------ | --------- |
| `CMPackager.ps1` | Main script (all core logic) |
| `CMPackager.prefs.template` | Configuration template (XML) |
| `GlobalConditions.xml` | Pre-defined SCCM global conditions |
| `Recipes/` | Active recipe XMLs (gitignored except Template.xml) |
| `Disabled/` | Example/template recipes and XSD schema |
| `ExtraFiles/icons/` | Application icons for SCCM |
| `ExtraFiles/Scripts/` | Helper scripts: `Get-WingetInfo.ps1` (recipe research), `New-ScaffoldRecipe.ps1` (recipe scaffolding), driver recipe generators, setup GUI XAML |
| `7za.exe` | Bundled 7-Zip for extraction tasks |
