<#
.SYNOPSIS
    Scaffolds new CMPackager recipe XML files from WinGet package information.

.DESCRIPTION
    This script accepts pipeline input from GetWingetInfo.ps1 and creates new recipe XML files
    by copying the appropriate template (_MSIRecipeTemplate.xml or _EXERecipeTemplate.xml) and
    populating it with information from the WinGet package manifest.

.PARAMETER InputObject
    Pipeline input from GetWingetInfo.ps1 containing package information.

.PARAMETER OutputPath
    Directory where the new recipe files should be created. Defaults to the Recipes folder.

.EXAMPLE
    .\GetWingetInfo.ps1 -ApplicationName "7-Zip" | .\ScaffoldRecipe.ps1

.EXAMPLE
    .\GetWingetInfo.ps1 -ApplicationName "Adobe Reader", "VLC" | .\ScaffoldRecipe.ps1 -OutputPath "C:\Temp\Recipes"

.NOTES
    Requires: PowerShell 5.1+
#>

[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [PSCustomObject[]]$InputObject,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (Test-Path $_ -PathType Container) {
            $true
        } else {
            throw "Output path '$_' does not exist or is not a directory."
        }
    })]
    [string]$OutputPath
)

begin {
    # Set default output path to Recipes folder if not specified
    if (-not $PSBoundParameters.ContainsKey('OutputPath')) {
        $scriptRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $OutputPath = Join-Path $scriptRoot "Recipes"
    }

    # Define template paths
    $scriptRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $msiTemplatePath = Join-Path $scriptRoot "Disabled\_MSIRecipeTemplate.xml"
    $exeTemplatePath = Join-Path $scriptRoot "Disabled\_EXERecipeTemplate.xml"

    # Validate template files exist
    if (-not (Test-Path $msiTemplatePath)) {
        throw "MSI template not found at: $msiTemplatePath"
    }
    if (-not (Test-Path $exeTemplatePath)) {
        throw "EXE template not found at: $exeTemplatePath"
    }

    Write-Verbose "Output path: $OutputPath"
    Write-Verbose "MSI Template: $msiTemplatePath"
    Write-Verbose "EXE Template: $exeTemplatePath"

    $processedCount = 0
}

process {
    foreach ($package in $InputObject) {
        try {
            Write-Host "`nProcessing $($package.Name)..." -ForegroundColor Cyan

            # Validate required properties
            if (-not $package.Name) {
                Write-Warning "Package missing 'Name' property. Skipping."
                continue
            }
            if (-not $package.'Installer Type') {
                Write-Warning "Package '$($package.Name)' missing 'Installer Type' property. Skipping."
                continue
            }

            # Determine which template to use based on installer type
            $installerType = $package.'Installer Type'.ToLower()
            $templatePath = $null
            $installationType = $null

            if ($installerType -eq 'msi') {
                $templatePath = $msiTemplatePath
                $installationType = 'MSI'
                Write-Verbose "Using MSI template"
            } elseif ($installerType -match '^(exe|inno|nullsoft|burn|wix)$') {
                $templatePath = $exeTemplatePath
                $installationType = 'Script'
                Write-Verbose "Using EXE/Script template for installer type: $installerType"
            } else {
                Write-Warning "Unsupported installer type '$installerType' for package '$($package.Name)'. Skipping."
                continue
            }

            # Generate output filename (remove spaces and special characters)
            $sanitizedName = $package.Name -replace '[^\w\s-]', '' -replace '\s+', ''
            $outputFileName = "$sanitizedName.xml"
            $outputFilePath = Join-Path $OutputPath $outputFileName

            # Check if file already exists
            if (Test-Path $outputFilePath) {
                Write-Warning "Recipe file '$outputFileName' already exists. Skipping."
                continue
            }

            # Load the template XML
            [xml]$recipeXml = Get-Content -Path $templatePath -Raw

            # Populate Application section
            $recipeXml.ApplicationDef.Application.Name = $package.Name
            $recipeXml.ApplicationDef.Application.Description = if ($package.Description -and $package.Description -ne "Not found") {
                $package.Description
            } else {
                $package.Name
            }
            $recipeXml.ApplicationDef.Application.Publisher = if ($package.Publisher -and $package.Publisher -ne "Not found") {
                $package.Publisher
            } else {
                "Unknown"
            }
            $recipeXml.ApplicationDef.Application.AutoInstall = "True"
            $recipeXml.ApplicationDef.Application.UserDocumentation = if ($package.HomePage -and $package.HomePage -ne "Not found") {
                $package.HomePage
            } else {
                ""
            }

            # Icon - use sanitized name + .png extension
            $iconName = "$sanitizedName.png"
            $recipeXml.ApplicationDef.Application.Icon = $iconName
            Write-Host "  Note: You'll need to add icon file: ExtraFiles/icons/$iconName" -ForegroundColor Yellow

            # Populate Downloads section
            $download = $recipeXml.ApplicationDef.Downloads.Download

            # Build PrefetchScript using the WinGet API URL and appropriate function
            if ($package.APIUrl -and $package.APIUrl -ne "Not found") {
                if ($installationType -eq 'MSI') {
                    $prefetchScript = "`$URL = Get-MSIInstallerURLfromWinget -apiUrl `"$($package.APIUrl)`""
                } else {
                    $prefetchScript = "`$URL = Get-ExeInstallerURLfromWinget -apiUrl `"$($package.APIUrl)`""
                }
                $download.PrefetchScript = $prefetchScript
                $download.URL = ""
            } else {
                # Fallback to direct URL if API URL is not available
                $download.PrefetchScript = ""
                if ($package.'Installer URL' -and $package.'Installer URL' -ne "Not found") {
                    $download.URL = $package.'Installer URL'
                } else {
                    $download.URL = ""
                }
            }

            # Set download filename
            if ($package.'Installer URL' -and $package.'Installer URL' -ne "Not found") {
                $urlFileName = [System.IO.Path]::GetFileName([Uri]$package.'Installer URL'.AbsolutePath)
                $download.DownloadFileName = $urlFileName
            } else {
                $download.DownloadFileName = "$sanitizedName.$installerType"
            }

            # Version check
            if ($installationType -eq 'MSI') {
                $download.DownloadVersionCheck = "[String]`$Version = ([String](Get-MSIInfo -Path `$DownloadFile -Property ProductVersion)).TrimStart().TrimEnd()"
            } else {
                $download.DownloadVersionCheck = ""
            }

            $download.Version = ""
            $download.FullVersion = ""
            $download.ExtraCopyFunctions = ""

            # Populate DeploymentType section
            $deploymentType = $recipeXml.ApplicationDef.DeploymentTypes.DeploymentType
            $deploymentType.DeploymentTypeName = "$($package.Name) Install"
            $deploymentType.InstallationType = $installationType

            # Set installation behavior defaults
            $deploymentType.CacheContent = "False"
            $deploymentType.BranchCache = "True"
            $deploymentType.ContentFallback = "True"
            $deploymentType.OnSlowNetwork = "Download"
            $deploymentType.InstallationBehaviorType = "InstallForSystem"
            $deploymentType.LogonReqType = "WhetherOrNotUserLoggedOn"
            $deploymentType.UserInteractionMode = "Hidden"
            $deploymentType.EstRuntimeMins = "15"
            $deploymentType.MaxRuntimeMins = "30"
            $deploymentType.RebootBehavior = "BasedOnExitCode"

            # Build install command based on installer type
            if ($installationType -eq 'MSI') {
                # For MSI, add InstallationMSI element if it doesn't exist
                $installationMsiNode = $deploymentType.SelectSingleNode("InstallationMSI")
                if (-not $installationMsiNode) {
                    $installationMsiNode = $recipeXml.CreateElement("InstallationMSI")
                    $deploymentType.InsertBefore($installationMsiNode, $deploymentType.InstallProgram) | Out-Null
                }
                $installationMsiNode.InnerText = $download.DownloadFileName

                $deploymentType.InstallProgram = "msiexec.exe /i $($download.DownloadFileName) /qn /norestart /l*v install.log"
                $deploymentType.UninstallCmd = "msiexec.exe /x $($download.DownloadFileName) /qn /norestart /l*v uninstall.log"
                $deploymentType.DetectionMethodType = "Custom"
            } else {
                # For EXE installers, use silent switches if available
                if ($package.'Silent switches' -and $package.'Silent switches' -ne "Not found") {
                    $silentSwitches = $package.'Silent switches'
                } elseif ($package.'SilentWithProgress switches' -and $package.'SilentWithProgress switches' -ne "Not found") {
                    $silentSwitches = $package.'SilentWithProgress switches'
                } else {
                    $silentSwitches = "/S"
                }

                $deploymentType.InstallProgram = "$($download.DownloadFileName) $silentSwitches"
                $deploymentType.UninstallCmd = ""
                $deploymentType.DetectionMethodType = "Custom"
            }

            # Set detection method placeholder
            # User will need to customize this based on the application
            $deploymentType.CustomDetectionMethods.DetectionClause.DetectionClauseType = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.Name = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.Path = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.PropertyType = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.ExpectedValue = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.ExpressionOperator = ""
            $deploymentType.CustomDetectionMethods.DetectionClause.Value = ""

            # Save the populated XML
            $xmlSettings = New-Object System.Xml.XmlWriterSettings
            $xmlSettings.Indent = $true
            $xmlSettings.IndentChars = "`t"
            $xmlSettings.NewLineChars = "`r`n"
            $xmlSettings.Encoding = [System.Text.Encoding]::UTF8

            $xmlWriter = [System.Xml.XmlWriter]::Create($outputFilePath, $xmlSettings)
            $recipeXml.Save($xmlWriter)
            $xmlWriter.Close()

            Write-Host "  ✓ Created recipe: $outputFileName" -ForegroundColor Green

            # Output summary of what needs to be customized
            Write-Host "`n  Required customizations for $outputFileName`:" -ForegroundColor Yellow
            Write-Host "    • Add icon file: ExtraFiles/icons/$iconName" -ForegroundColor Gray
            Write-Host "    • Configure detection method in DeploymentType section" -ForegroundColor Gray
            if ($installationType -eq 'Script') {
                Write-Host "    • Verify/customize install command and switches" -ForegroundColor Gray
                Write-Host "    • Add uninstall command if needed" -ForegroundColor Gray
            }
            Write-Host "    • Review and test the recipe before deployment`n" -ForegroundColor Gray

            $processedCount++

        } catch {
            Write-Error "Error processing package '$($package.Name)': $_"
            continue
        }
    }
}

end {
    Write-Host "`nProcessed $processedCount recipe(s) successfully." -ForegroundColor Cyan
    if ($processedCount -gt 0) {
        Write-Host "Recipe files created in: $OutputPath" -ForegroundColor Cyan
    }
}
