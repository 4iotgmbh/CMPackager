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

**GetWingetInfo.ps1 Script** (`ExtraFiles/Scripts/GetWingetInfo.ps1`):

- PowerShell 7 utility for researching applications when creating new recipes
- Searches the WinGet repository and provides interactive selection via `Out-ConsoleGridView`
- Retrieves comprehensive package details from WinGet manifests (YAML files on GitHub)
- Returns: installer URLs, version info, publisher, description, silent install switches, product codes, architecture, etc.
- Requires modules: `Microsoft.WinGet.Client`, `Microsoft.PowerShell.ConsoleGuiTools`, `powershell-yaml`, `cobalt`
- Usage: `.\GetWingetInfo.ps1 -ApplicationName "Adobe Reader", "7-Zip" -Output List`

**ScaffoldRecipe.ps1 Script** (`ExtraFiles/Scripts/ScaffoldRecipe.ps1`):

- Accepts pipeline input from GetWingetInfo.ps1 to automatically scaffold new recipe files
- Selects the appropriate template (`_MSIRecipeTemplate.xml` or `_EXERecipeTemplate.xml`) based on installer type
- Creates recipe file named after the application (spaces removed)
- Populates XML fields: Name, Description, Publisher, UserDocumentation, PrefetchScript, install commands, etc.
- Outputs to `Recipes/` folder by default (customizable via `-OutputPath` parameter)
- Usage: `.\GetWingetInfo.ps1 -ApplicationName "7-Zip" | .\ScaffoldRecipe.ps1`
- Note: Generated recipes require customization (detection methods, icons, testing) before deployment

**Dynamic URL Functions in CMPackager.ps1**:

- **`Get-MSIInstallerURLfromWinget`** ([CMPackager.ps1:301](CMPackager.ps1#L301)) — Queries GitHub API for WinGet package manifests, parses the latest version's installer YAML, and returns the MSI download URL
- **`Get-ExeInstallerURLfromWinget`** ([CMPackager.ps1:371](CMPackager.ps1#L371)) — Same functionality for EXE installers

These functions are called within recipe `<PrefetchScript>` blocks to dynamically fetch current download URLs at runtime:

```powershell
# Example PrefetchScript usage in a recipe:
$apiUrl = "https://api.github.com/repos/microsoft/winget-pkgs/contents/manifests/a/Adobe/Acrobat/Reader/64-bit"
$DownloadURL = Get-MSIInstallerURLfromWinget -apiUrl $apiUrl
```

**Complete Recipe Creation Workflow**:

```powershell
# Step 1: Research the application and gather WinGet information
cd ExtraFiles\Scripts
.\GetWingetInfo.ps1 -ApplicationName "7-Zip"

# Step 2: Pipe the output to ScaffoldRecipe.ps1 to create a new recipe
.\GetWingetInfo.ps1 -ApplicationName "7-Zip" | .\ScaffoldRecipe.ps1

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
- **`Get-MSIInstallerURLfromWinget`** / **`Get-ExeInstallerURLfromWinget`** — Query WinGet package manifests from GitHub to retrieve current installer download URLs; used in recipe PrefetchScripts to dynamically obtain the latest version download links from the winget-pkgs repository

### Global State

All shared state uses `$Global:` prefix variables initialized from the prefs file: `$Global:CMSite`, `$Global:SiteServer`, `$Global:ContentLocationRoot`, `$Global:TempDir`, `$Global:LogPath`, etc.

## Coding Conventions

- **Function naming**: PowerShell verb-noun PascalCase (`Invoke-ApplicationCreation`, `Add-DeploymentType`)
- **Parameter validation**: Uses `ValidateScript`, `ValidateSet`, `ValidateNotNullOrEmpty`
- **Error handling**: Try-catch blocks with `Add-LogContent` for detailed logging
- **Dynamic parameters**: Used for the `-SingleRecipe` parameter (populated from Recipes/ folder at runtime)
- **XML entity encoding**: PrefetchScript blocks in recipes use `&amp;` for `&`, etc.

## File Layout

| Path | Purpose |
| ------ | --------- |
| `CMPackager.ps1` | Main script (all core logic) |
| `CMPackager.prefs.template` | Configuration template (XML) |
| `GlobalConditions.xml` | Pre-defined SCCM global conditions |
| `Recipes/` | Active recipe XMLs (gitignored except Template.xml) |
| `Disabled/` | Example/template recipes and XSD schema |
| `ExtraFiles/icons/` | Application icons for SCCM |
| `ExtraFiles/Scripts/` | Helper scripts: `GetWingetInfo.ps1` (recipe research), `ScaffoldRecipe.ps1` (recipe scaffolding), driver recipe generators, setup GUI XAML |
| `7za.exe` | Bundled 7-Zip for extraction tasks |
