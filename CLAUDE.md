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

### Key Supporting Functions

- **`Add-DetectionMethodClause`** — Builds SCCM detection logic (Registry, File, MSI clauses)
- **`Add-RequirementsRule`** — Adds OS, global condition, or custom requirements
- **`Copy-CMDeploymentTypeRule`** — Copies requirement rules between deployment types (adapted from external source)
- **`Connect-ConfigMgr`** — Establishes SCCM site connection and imports the ConfigurationManager module
- **`Add-LogContent`** — Centralized logging with rotation (`$Global:LogPath`)
- **`Get-MSIInfo`** / **`Get-MSISourceFileVersion`** — Extract metadata from MSI files using COM

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
|------|---------|
| `CMPackager.ps1` | Main script (all core logic) |
| `CMPackager.prefs.template` | Configuration template (XML) |
| `GlobalConditions.xml` | Pre-defined SCCM global conditions |
| `Recipes/` | Active recipe XMLs (gitignored except Template.xml) |
| `Disabled/` | Example/template recipes and XSD schema |
| `ExtraFiles/icons/` | Application icons for SCCM |
| `ExtraFiles/Scripts/` | Helper scripts (driver recipe generators, setup GUI XAML) |
| `7za.exe` | Bundled 7-Zip for extraction tasks |
