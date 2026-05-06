# CMPackager - 4IoT Fork

A PowerShell automation tool for SCCM/MEM ConfigMgr that downloads, packages, distributes, and deploys applications using XML "recipe" files. The goal is to package any frequently-updating application with little to no ongoing work after the initial recipe is created.

## Getting Started

1. Clone or download the project.
2. Copy `CMPackager.prefs.template` to `CMPackager.prefs` and fill in your SCCM site details.
3. Browse the ~75 example recipes in the `Disabled/` folder, adjust them to your environment, and move them into `Recipes/`.
4. Run `CMPackager.ps1` - every recipe in `Recipes/` is processed if a newer version is available.

```powershell
# Standard run
.\CMPackager.ps1

# Run a single recipe (tab-completable from Recipes/ folder)
.\CMPackager.ps1 -SingleRecipe 7-Zip.xml

# First-time setup
.\CMPackager.ps1 -Setup
```

### Prerequisites

- SCCM 1906+ / MEM ConfigMgr (tested on SCCM 2509); console must have been opened at least once.
- PowerShell 5.1+ for the main script.
- PowerShell 7+ for helper scripts (`Get-WingetInfo.ps1`, `Test-RecipeInstallation.ps1`, `Test-RecipeBatch.ps1`).

## What This Fork Adds

### WinGet-Backed Download URLs

Most recipes now use `Get-InstallerURLfromWinget` - a function built into `CMPackager.ps1` - inside their `<PrefetchScript>` block. Instead of scraping vendor websites, the function queries the official [WinGet package manifest repository](https://github.com/microsoft/winget-pkgs) on GitHub to resolve the current installer URL at runtime.

```powershell
# Example PrefetchScript in a recipe
$DownloadURL = Get-InstallerURLfromWinget `
    -apiUrl "https://api.github.com/repos/microsoft/winget-pkgs/contents/manifests/7/7zip/7zip" `
    -InstallerType msi `
    -Architecture x64 `
    -Scope machine
```

Supported parameters: `-InstallerType msi|exe|zip`, `-Architecture x64|x86|arm64|arm`, `-Scope machine|user`.

Set `<GitHubToken>` in your prefs file (or `$env:GITHUB_TOKEN`) to avoid GitHub API rate limits.

### Recipe Research and Scaffolding Workflow

Two helper scripts in `ExtraFiles/Scripts/` speed up new recipe creation:

**Step 1 - Research the package:**

```powershell
cd ExtraFiles\Scripts
.\Get-WingetInfo.ps1 -ApplicationName "7-Zip"
```

Opens an interactive grid to select the package, then pulls full metadata from the WinGet YAML manifests: installer URLs, product codes, publisher, silent switches, architectures, and more. Requires PS 7+ and the `Microsoft.WinGet.Client`, `Microsoft.PowerShell.ConsoleGuiTools`, `powershell-yaml`, and `cobalt` modules.

**Step 2 - Scaffold the recipe file:**

```powershell
.\Get-WingetInfo.ps1 -ApplicationName "7-Zip" | .\New-ScaffoldRecipe.ps1
```

Pipes the metadata into `New-ScaffoldRecipe.ps1`, which picks the right template (`_MSIRecipeTemplate.xml` or `_EXERecipeTemplate.xml`), creates the recipe file in `Recipes/`, and pre-fills all fields it can derive from WinGet. You still need to add the icon, verify the detection method, and test before deploying.

### Three-Ring Deployment Model

All recipes normalised to include three deployment collections:

| Ring | Purpose |
|------|---------|
| Test | IT / early testers, immediate deployment |
| Pilot | Selected power users, short deadline |
| GA | All users / general availability, extended deadline |

### Automated Recipe Testing (Windows Sandbox)

Two scripts validate that a recipe's install, detection, and uninstall commands actually work before you push anything to SCCM:

```powershell
# Test one recipe
.\ExtraFiles\Scripts\Test-RecipeInstallation.ps1 -RecipePath .\Disabled\7-Zip.xml

# Batch-test a whole folder and write a CSV report
.\ExtraFiles\Scripts\Test-RecipeBatch.ps1 -RecipesPath .\Disabled
```

Each test spins up a clean Windows Sandbox instance, runs install, detection, uninstall, and detection again, then reports `PASS / FAIL / TIMEOUT / SKIPPED`. Results are saved as `results.json` and a timestamped CSV. Requires Windows 11 with Windows Sandbox enabled and PS 7+.

### Web UI

A browser-based dashboard for managing and monitoring CMPackager without touching the command line.

```powershell
# Start on the default port (8080)
powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1

# Use a different port
powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1 -Port 9090

# Enable verbose server-side logging
powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1 -DebugMode
```

Then open `http://localhost:8080/` in a browser. The server reads `CMPackager.prefs` automatically; a warning banner appears in the UI if prefs are missing.

The UI has four tabs:

| Tab | What it does |
|-----|-------------|
| **Recipes** | Two-column view of enabled and disabled recipes. Enable/disable a recipe by clicking the arrow button (moves the file between `Recipes/` and `Disabled/`). Run a single recipe directly from the card. Set per-recipe Windows Task Scheduler schedules (daily / weekly / monthly) without leaving the browser - the server auto-assigns staggered start times to avoid conflicts. |
| **Output** | Live terminal output streamed via Server-Sent Events while CMPackager is running. Optionally show a live tail of the log file alongside process output. Auto-scroll toggle included. |
| **Test Results** | Loads the latest `RecipeTestResults_*.csv` produced by `Test-RecipeBatch.ps1` and renders it as a sortable, filterable table with colour-coded PASS / FAIL / TIMEOUT / SKIPPED badges. |
| **SCCM Status** | Connects to your SCCM site and shows the current application version and deployment statistics (targeted, success, errors, in-progress) for every active recipe. |

The header also has a global **Run All** button and a **Stop** button (visible while a run is active), and a live status pill that pulses green while CMPackager is running.

The server is pure PowerShell with no external dependencies - it uses `System.Net.HttpListener` and a runspace pool for concurrent request handling.

### Surface Driver Packaging

1. Add `MicrosoftSurfaceDriversRecipe.xml` from `Disabled/` to `Recipes/`.
2. Edit `ExtraFiles/Scripts/MicrosoftDrivers.csv` - remove any models you don't want packaged.
3. Run CMPackager normally; the first run creates per-model recipes; subsequent runs download updated drivers.

## Recipe Library

| Location | Count | Purpose |
|----------|-------|---------|
| `Recipes/` | Active | Recipes processed on each run (gitignored except `Template.xml`) |
| `Disabled/` | ~75 | Examples and templates to start from |

Notable apps covered: 7-Zip, Adobe Reader, AzureDataStudio, Chrome, Firefox, Git, JetBrains Toolbox, Notepad++, PowerShell, Python, VLC, VS Code, Wireshark, Zoom, and many more.

## Contributing

Pull requests and issue reports are welcome. Recipes are the easiest contribution - see `Disabled/_MSIRecipeTemplate.xml` and `Disabled/_EXERecipeTemplate.xml` as starting points, or use the scaffold workflow above.

## Authors

- **Andrew Jimenez** - *Original Author* - [asjimene](https://github.com/asjimene)
- **Mirko Schnellbach** - *Fork Maintainer* - [4IoTMirko](https://github.com/4IoTMirko)

See also the [contributors list](https://github.com/4IoTGmbH/CMPackager/graphs/contributors).

## Acknowledgments

Code adapted from:

- Janik von Rots - [Copy-CMDeploymentTypeRule](https://janikvonrotz.ch/2017/10/20/configuration-manager-configure-requirement-rules-for-deployment-types-with-powershell/)
- Jaap Brasser - [Get-ExtensionAttribute](http://www.jaapbrasser.com)
- Nickolaj Andersen - [Get-MSIInfo](http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/)

## License Notice

This project does not distribute applications. Recipes provide links to vendor download URLs. Downloading and packaging software with this tool does not grant you a license for that software. Ensure you are properly licensed for everything you package and distribute.
