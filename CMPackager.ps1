<#	
	.NOTES
	===========================================================================
	 Created on:   		1/9/2018 11:34 AM
	 Last Updated:  	02/26/2026
	 Original Author:	Andrew Jimenez (asjimene) - https://github.com/asjimene/
	 Fork Maintainer: Mirko Schnellbach (4IoTMirko) - 4IoT GmbH - https://github.com/4iotgmbh
	 Filename:     		CMPackager.ps1
	 Fork URL:				https://github.com/4iotgmbh/CMPackager
	===========================================================================
	.DESCRIPTION
		Packages Applications for ConfigMgr using XML Based Recipe Files

	Uses Scripts and Functions Sourced from the Following:
		Copy-CMDeploymentTypeRule - https://janikvonrotz.ch/2017/10/20/configuration-manager-configure-requirement-rules-for-deployment-types-with-powershell/
		Get-ExtensionAttribute - Jaap Brasser - http://www.jaapbrasser.com
		Get-MSIInfo - Nickolaj Andersen - http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
	
	7-Zip Application is Redistributed for Ease of Use:
		7-Zip Binary - Igor Pavlov - https://www.7-zip.org/
#>

[CmdletBinding()]
param (
	[switch]$Setup = $false,
	[switch]$WebServer = $false,

	[ValidateScript({
		if (-not ($_ | Resolve-Path | Test-Path -PathType Leaf)) {
			throw "File doesn't exist or a file wasn't specified."
		}
		return $true
	})]
	[System.IO.FileInfo]$PreferenceFile = "$PSScriptRoot\CMPackager.prefs",
	
	[ValidateScript({
		if (-not ($_ | Resolve-Path | Test-Path -PathType Container)) {
			throw "Directory doesn't exist or a directory wasn't specified."
		}
		return $true
	})]
	[System.IO.DirectoryInfo]$RecipePath = "$PSScriptRoot\Recipes"
)
DynamicParam {
	# If RecipePath is specified populate list of available recipes from custom recipe location  
	if ($PSBoundParameters['RecipePath']) {
		$configurationFileNames = Get-ChildItem *.xml -Path $($PSBoundParameters['RecipePath']) | Select-Object -ExpandProperty Name
	}
	else 
	{
	# If RecipePath is note specified, check to see if running from the CMPackager directory and populate from standard recipe directory
		if ((test-path .\Recipes) -and (test-path .\cmpackager.ps1)) {
			$configurationFileNames = Get-ChildItem *.xml -Path .\Recipes | Select-Object -ExpandProperty Name
		}
	}
	# Make SingleRecipe parameter availabe only if possible recipes are found in the above checked custom or standard dirs
	if ($configurationFileNames) {
		$ParamAttrib = New-Object System.Management.Automation.ParameterAttribute
		$ParamAttrib.Mandatory = $false
		$ParamAttrib.ParameterSetName = '__AllParameterSets'
		$AttribColl = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
		$AttribColl.Add($ParamAttrib)
		$AttribColl.Add((New-Object System.Management.Automation.AliasAttribute('Recipe')))
		$AttribColl.Add((New-Object System.Management.Automation.ValidateSetAttribute($configurationFileNames)))
		$RuntimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('SingleRecipe', [string[]], $AttribColl)
		$RuntimeParamDic = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$RuntimeParamDic.Add('SingleRecipe', $RuntimeParam)
		return  $RuntimeParamDic
	}
}

process {

	$Global:ScriptVersion = "26.02.26.0 - 4IoT Gmbh fork"

	$Global:ScriptRoot = $PSScriptRoot

	Write-Host "Preference file: $($PreferenceFile), Recipe path: $($RecipePath)"

	if (-not (Test-Path $PreferenceFile -ErrorAction SilentlyContinue)) {
		$Setup = $true
	}
	## Global Variables (Only load if not setup)
	# Import the Prefs file
	if (-not ($Setup)) {
		[xml]$PackagerPrefs = Get-Content $PreferenceFile

		# Packager Vars
		$Global:TempDir = $PackagerPrefs.PackagerPrefs.TempDir
		$Global:LogPath = $PackagerPrefs.PackagerPrefs.LogPath
		$Global:MaxLogSize = 1000kb

		# Package Location Vars
		$Global:ContentLocationRoot = $PackagerPrefs.PackagerPrefs.ContentLocationRoot
		$Global:ContentFolderPattern = $PackagerPrefs.PackagerPrefs.ContentFolderPattern
		$Global:IconRepo = $PackagerPrefs.PackagerPrefs.IconRepo

		# CM Vars
		$Global:CMSite = $PackagerPrefs.PackagerPrefs.CMSite
		$Global:SiteCode = ($Global:CMSite).Replace(':', '')
		$Global:SiteServer = $PackagerPrefs.PackagerPrefs.SiteServer
		$Global:RequirementsTemplateAppName = $PackagerPrefs.PackagerPrefs.RequirementsTemplateAppName
		$Global:PreferredDistributionLoc = $PackagerPrefs.PackagerPrefs.PreferredDistributionLoc
		$Global:PreferredDeployCollection = $PackagerPrefs.PackagerPrefs.PreferredDeployCollection
		$Global:NoVersionInSWCenter = [System.Convert]::ToBoolean($PackagerPrefs.PackagerPrefs.NoVersionInSWCenter)
		$Global:CMPSModulePath = $PackagerPrefs.PackagerPrefs.CMPSModulePath


		# Email Vars
		[string[]]$Global:EmailTo = [string[]]$PackagerPrefs.PackagerPrefs.EmailTo
		$Global:EmailFrom = $PackagerPrefs.PackagerPrefs.EmailFrom
		$Global:EmailServer = $PackagerPrefs.PackagerPrefs.EmailServer
		$Global:SendEmailPreference = [System.Convert]::ToBoolean($PackagerPrefs.PackagerPrefs.SendEmailPreference)
		$Global:NotifyOnDownloadFailure = [System.Convert]::ToBoolean($PackagerPrefs.PackagerPrefs.NotifyOnDownloadFailure)

		$Global:EmailSubject = "CMPackager Report [$Global:CMSite] - $(Get-Date -format d)"
		$Global:EmailBody = "[$Global:CMSite] New Application Updates Packaged on $(Get-Date -Format d)`n`n"

		#This gets switched to True if Applications are Packaged
		$Global:SendEmail = $false
		$Global:TemplateApplicationCreatedFlag = $false

		# GitHub API token — raises quota from 60 to 5,000 req/hr
		$Global:GitHubToken = $PackagerPrefs.PackagerPrefs.GitHubToken

		# Web server settings
		$Global:WebServerPort         = $PackagerPrefs.PackagerPrefs.WebServerPort
		$Global:WebServerRequiredRole = $PackagerPrefs.PackagerPrefs.WebServerRequiredRole
	}

	$Global:ConfigMgrConnection = $false

	$Global:OperatorsLookup = @{ And = 'And'; Or = 'Or'; Other = 'Other'; IsEquals = 'Equals'; NotEquals = 'Not equal to'; GreaterThan = 'Greater than'; LessThan = 'Less than'; Between = 'Between'; NotBetween = 'Not Between'; GreaterEquals = 'Greater than or equal to'; LessEquals = 'Less than or equal to'; BeginsWith = 'Begins with'; NotBeginsWith = 'Does not begin with'; EndsWith = 'Ends with'; NotEndsWith = 'Does not end with'; Contains = 'Contains'; NotContains = 'Does not contain'; AllOf = 'All of'; OneOf = 'OneOf'; NoneOf = 'NoneOf'; SetEquals = 'Set equals'; SubsetOf = 'Subset of'; ExcludesAll = 'Exludes all' }
	## Functions
	function Add-LogContent {
		param
		(
			[parameter(Mandatory = $false)]
			[switch]$Load,
			[parameter(Mandatory = $true)]
			$Content
		)
		$line = "$(Get-Date -Format G) - $Content`r`n"
		if ($Load -and (Get-Item $LogPath -ErrorAction SilentlyContinue).length -gt $MaxLogSize) {
			[System.IO.File]::WriteAllText($LogPath, $line, [System.Text.Encoding]::UTF8)
		}
		else {
			[System.IO.File]::AppendAllText($LogPath, $line, [System.Text.Encoding]::UTF8)
		}
	}

	function Get-ExtensionAttribute {
		<#
.Synopsis
Retrieves extension attributes from files or folder

.DESCRIPTION
Uses the dynamically generated parameter -ExtensionAttribute to select one or multiple extension attributes and display the attribute(s) along with the FullName attribute

.NOTES   
Name: Get-ExtensionAttribute.ps1
Author: Jaap Brasser
Version: 1.0
DateCreated: 2015-03-30
DateUpdated: 2015-03-30
Blog: http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.PARAMETER FullName
The path to the file or folder of which the attributes should be retrieved. Can take input from pipeline and multiple values are accepted.

.PARAMETER ExtensionAttribute
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

.EXAMPLE   
. .\Get-ExtensionAttribute.ps1
    
Description 
-----------     
This command dot sources the script to ensure the Get-ExtensionAttribute function is available in your current PowerShell session

.EXAMPLE
Get-ExtensionAttribute -FullName C:\Music -ExtensionAttribute Size,Length,Bitrate

Description
-----------
Retrieves the Size,Length,Bitrate and FullName of the contents of the C:\Music folder, non recursively

.EXAMPLE
Get-ExtensionAttribute -FullName C:\Music\Song2.mp3,C:\Music\Song.mp3 -ExtensionAttribute Size,Length,Bitrate

Description
-----------
Retrieves the Size,Length,Bitrate and FullName of Song.mp3 and Song2.mp3 in the C:\Music folder

.EXAMPLE
Get-ChildItem -Recurse C:\Video | Get-ExtensionAttribute -ExtensionAttribute Size,Length,Bitrate,Totalbitrate

Description
-----------
Uses the Get-ChildItem cmdlet to provide input to the Get-ExtensionAttribute function and retrieves selected attributes for the C:\Videos folder recursively

.EXAMPLE
Get-ChildItem -Recurse C:\Music | Select-Object FullName,Length,@{Name = 'Bitrate' ; Expression = { Get-ExtensionAttribute -FullName $_.FullName -ExtensionAttribute Bitrate | Select-Object -ExpandProperty Bitrate } }

Description
-----------
Combines the output from Get-ChildItem with the Get-ExtensionAttribute function, selecting the FullName and Length properties from Get-ChildItem with the ExtensionAttribute Bitrate
#>
		[CmdletBinding()]
		Param (
			[Parameter(ValueFromPipeline = $true,
				ValueFromPipelineByPropertyName = $true,
				Position = 0)]
			[string[]]$FullName
		)
		DynamicParam {
			$Attributes = New-Object System.Management.Automation.ParameterAttribute
			$Attributes.ParameterSetName = "__AllParameterSets"
			$Attributes.Mandatory = $false
			$AttributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
			$AttributeCollection.Add($Attributes)
			$Values = @($Com = (New-Object -ComObject Shell.Application).NameSpace('C:\'); 1 .. 400 | ForEach-Object { $com.GetDetailsOf($com.Items, $_) } | Where-Object { $_ } | ForEach-Object { $_ -replace '\s' })
			$AttributeValues = New-Object System.Management.Automation.ValidateSetAttribute($Values)
			$AttributeCollection.Add($AttributeValues)
			$DynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ExtensionAttribute", [string[]], $AttributeCollection)
			$ParamDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
			$ParamDictionary.Add("ExtensionAttribute", $DynParam1)
			$ParamDictionary
		}
	
		begin {
			$ShellObject = New-Object -ComObject Shell.Application
			$DefaultName = $ShellObject.NameSpace('C:\')
			$ExtList = 0 .. 400 | ForEach-Object {
				($DefaultName.GetDetailsOf($DefaultName.Items, $_)).ToUpper().Replace(' ', '')
			}
		}
	
		process {
			foreach ($Object in $FullName) {
				# Check if there is a fullname attribute, in case pipeline from Get-ChildItem is used
				if ($Object.FullName) {
					$Object = $Object.FullName
				}
			
				# Check if the path is a single file or a folder
				if (-not (Test-Path -Path $Object -PathType Container)) {
					$CurrentNameSpace = $ShellObject.NameSpace($(Split-Path -Path $Object))
					$CurrentNameSpace.Items() | Where-Object {
						$_.Path -eq $Object
					} | ForEach-Object {
						$HashProperties = @{
							FullName = $_.Path
						}
						foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute) {
							$HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
						}
						New-Object -TypeName PSCustomObject -Property $HashProperties
					}
				}
				elseif (-not $input) {
					$CurrentNameSpace = $ShellObject.NameSpace($Object)
					$CurrentNameSpace.Items() | ForEach-Object {
						$HashProperties = @{
							FullName = $_.Path
						}
						foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute) {
							$HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
						}
						New-Object -TypeName PSCustomObject -Property $HashProperties
					}
				}
			}
		}
	
		end {
			Remove-Variable -Force -Name DefaultName
			Remove-Variable -Force -Name CurrentNameSpace
			Remove-Variable -Force -Name ShellObject
		}
	}

function Get-GitHubAuthHeaders {
    # Canonical copy -- keep in sync with sibling scripts in ExtraFiles\Scripts\.
    # Precedence: CMPackager.prefs <GitHubToken> > $env:GITHUB_TOKEN > anonymous.
    param([string]$PrefsToken = '')
    $token = if ($PrefsToken) { $PrefsToken } elseif ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } else { $null }
    $h = @{ 'User-Agent' = 'CMPackager' }
    if ($token) { $h['Authorization'] = "Bearer $token" }
    return $h
}

function Get-InstallerURLfromWinget {
  param (
    [parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$apiUrl,

    [parameter(Mandatory = $true)]
    [ValidateSet("msi", "exe", "zip", "burn", "wix")]
    [string]$InstallerType,

    [parameter(Mandatory = $false)]
    [ValidateSet("x64", "x86", "arm64", "arm")]
    [string]$Architecture,

    [parameter(Mandatory = $false)]
    [ValidateSet("machine", "user")]
    [string]$Scope
  )
  # Reliably determine the current installer download URL
  # Method: Query the winget (Windows Package Manager) manifest from GitHub
  # This is publicly accessible, machine-readable, and always up to date.
  # Use -Architecture and -Scope to select among multiple installers in the manifest.

  # Map installer types that share a file extension with another type
  $fileExtension = switch ($InstallerType) { "burn" { "exe" } "wix" { "msi" } default { $InstallerType } }

  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  # Resolve wrapper download URLs to a directly downloadable CDN URL.
  # Some manifests (e.g. WinSCP via SourceForge) list a /download page rather
  # than a direct file link.  SourceForge serves an HTML countdown page for GET
  # requests; the real CDN URL is embedded inside that HTML.  HEAD requests do
  # not redirect to the file, so we must fetch the page and parse it out.
  function Resolve-InstallerUrl ([string]$Url) {
      if ($Url -notmatch '/download$') { return $Url }
      try {
          $html = (Invoke-WebRequest -Uri $Url -UseBasicParsing -MaximumRedirection 10 `
                      -TimeoutSec 30 -ErrorAction Stop).Content
          # SourceForge embeds the CDN URL in a <meta http-equiv="refresh"> tag and
          # the "Problems Downloading?" button's data-release-url attribute.
          $m = [regex]::Match($html, 'https://downloads\.sourceforge\.net/project/[^\s"''<>\\]+')
          if ($m.Success) {
              return [System.Net.WebUtility]::HtmlDecode($m.Value)
          }
      } catch [System.Net.WebException] {
          if ($_.Exception.Response) {
              $final = $_.Exception.Response.ResponseUri.AbsoluteUri
              $_.Exception.Response.Close()
              if ($final -and $final -ne $Url) { return $final }
          }
      } catch {}
      return $Url
  }

  try {
      $headers = Get-GitHubAuthHeaders -PrefsToken $Global:GitHubToken
      $headers['Accept'] = 'application/vnd.github.v3+json'

      # Step 1: List version folders to find the latest
      $versions = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop

      # Sort versions properly (semantic versioning only, skip the rest) and pick the latest
      $latestVersion = $versions |
          Where-Object { $_.type -eq "dir" } |
          Where-Object { $_.name -notmatch '[A-Za-z]'} |
          Sort-Object { [version]($_.name -replace '[^0-9.]', '') } -Descending |
          Select-Object -First 1

      Write-Verbose "Latest version found: $($latestVersion.name)"

      # Step 2: Get the installer manifest
      $versionUrl = "$apiUrl/$($latestVersion.name)"
      $files = Invoke-RestMethod -Uri $versionUrl -Headers $headers -ErrorAction Stop

      $installerFile = $files | Where-Object { $_.name -match "installer" }

      if ($installerFile) {
          # Step 3: Download and parse the installer YAML
          $yamlContent = Invoke-RestMethod -Uri $installerFile.download_url -Headers $headers -ErrorAction Stop

          # Parse the Installers: list into individual blocks for filtering
          $installerBlocks = @()
          $lines = $yamlContent -split '\r?\n'
          $inInstallers = $false
          $currentBlock = $null

          foreach ($line in $lines) {
              if ($line -match '^Installers:\s*$') {
                  $inInstallers = $true
                  continue
              }
              if ($inInstallers) {
                  if ($line -match '^- ') {
                      if ($null -ne $currentBlock) { $installerBlocks += $currentBlock }
                      $currentBlock = $line.Substring(2) + "`n"
                  } elseif ($line -match '^  ') {
                      if ($null -ne $currentBlock) { $currentBlock += $line.TrimStart() + "`n" }
                  } elseif ($line -notmatch '^$') {
                      break
                  }
              }
          }
          if ($null -ne $currentBlock) { $installerBlocks += $currentBlock }

          # Fall back to full YAML if parsing yielded nothing
          if ($installerBlocks.Count -eq 0) {
              $urlMatch = [regex]::Match($yamlContent, "(?i)InstallerUrl:\s*(https://[^\s]+\.$fileExtension(?:/download)?)")
              if ($urlMatch.Success) {
                  if ($InstallerType -eq "zip") {
                      $nestedMatch = [regex]::Match($yamlContent, "(?i)RelativeFilePath:\s*(.+)")
                      $nestedPath = if ($nestedMatch.Success) { $nestedMatch.Groups[1].Value.Trim() } else { $null }
                      [PSCustomObject]@{ Url = (Resolve-InstallerUrl $urlMatch.Groups[1].Value); NestedFilePath = $nestedPath }
                  } else {
                      Resolve-InstallerUrl $urlMatch.Groups[1].Value
                  }
              }
              else { Write-Warning "Could not parse $InstallerType URL from installer manifest." }
              return
          }

          # Filter by Architecture if specified
          $filtered = $installerBlocks
          if ($Architecture) {
              $archFiltered = $filtered | Where-Object { $_ -match "(?i)Architecture:\s*$Architecture\b" }
              if ($archFiltered) { $filtered = $archFiltered }
              else { Write-Warning "No installer found for Architecture '$Architecture', trying all entries." }
          }

          # Filter by Scope if specified
          if ($Scope) {
              $scopeFiltered = $filtered | Where-Object { $_ -match "(?i)Scope:\s*$Scope\b" }
              if ($scopeFiltered) { $filtered = $scopeFiltered }
              else { Write-Warning "No installer found for Scope '$Scope', trying all filtered entries." }
          }

          # Filter by InstallerType if blocks declare it explicitly (handles per-installer InstallerType)
          $typeFiltered = $filtered | Where-Object { $_ -match "(?i)InstallerType:\s*$InstallerType\b" }
          if ($typeFiltered) { $filtered = $typeFiltered }

          $targetBlock = $filtered | Select-Object -First 1

          if ($targetBlock) {
              $urlMatch = [regex]::Match($targetBlock, "(?i)InstallerUrl:\s*(https://[^\s]+\.$fileExtension(?:/download)?)")
              if ($urlMatch.Success) {
                  if ($InstallerType -eq "zip") {
                      $nestedMatch = [regex]::Match($targetBlock, "(?i)RelativeFilePath:\s*(.+)")
                      $nestedPath = if ($nestedMatch.Success) { $nestedMatch.Groups[1].Value.Trim() } else { $null }
                      [PSCustomObject]@{ Url = (Resolve-InstallerUrl $urlMatch.Groups[1].Value); NestedFilePath = $nestedPath }
                  } else {
                      Resolve-InstallerUrl $urlMatch.Groups[1].Value
                  }
              } else {
                  Write-Warning "Could not parse $InstallerType URL from installer manifest."
                  Write-Verbose $targetBlock
              }
          } else {
              Write-Warning "No installer matching the specified criteria found."
              Write-Verbose $yamlContent
          }
      } else {
          Write-Warning "Installer manifest file not found in version folder."
      }
  }
  catch {
      $statusCode = $null
      if ($_.Exception.Response) {
          $statusCode = [int]$_.Exception.Response.StatusCode
      } elseif ($_.Exception -is [Microsoft.PowerShell.Commands.HttpResponseException]) {
          $statusCode = [int]$_.Exception.StatusCode
      }
      if ($statusCode -eq 401) {
          throw "GitHub API returned 401 Unauthorized. The configured token is invalid or expired.`n  ACTION: Generate a new personal access token at https://github.com/settings/tokens and set it as <GitHubToken> in CMPackager.prefs (or in the GITHUB_TOKEN environment variable)."
      } elseif ($statusCode -eq 403) {
          throw "GitHub API returned 403 Forbidden. The anonymous rate limit (60 req/hr) has been reached.`n  ACTION: Add a GitHub personal access token as <GitHubToken> in CMPackager.prefs (or GITHUB_TOKEN env var) to raise the limit to 5,000 req/hr."
      } else {
          throw "GitHub API error: $($_.Exception.Message)"
      }
  }
}

	function Get-MSIInfo {
		param (
			[parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[System.IO.FileInfo]$Path,
			[parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion", "InstallPrerequisites")]
			[string]$Property
		)
	
		Process {
			try {
				# Read property from MSI database
				$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
				$MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
				$Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
				$View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
				$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
				$Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
				$Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
			
				# Commit database and close view
				$MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
				$View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)
				$MSIDatabase = $null
				$View = $null
			
				# Return the value
				return $Value
			}
			catch {
				Write-Warning -Message $_.Exception.Message; break
			}
		}
		End {
			# Run garbage collection and release ComObject
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
			[System.GC]::Collect()
		}
	}

	function Get-MSISourceFileVersion {
		<#
		.SYNOPSIS
			Get the version of a file from an MSI's File Table
		.DESCRIPTION
			Search a Windows Installer database's File Table for a file name and return the version.
		.EXAMPLE
			PS C:\> Get-MSISourceFileVersion -Msi "C:\Program Files\Microsoft Configuration Manager\tools\ConsoleSetup\AdminConsole.msi" -FileName 'ConBlder.exe|AdminUI.ConsoleBuilder.exe'
			Get the version of the file 'ConBlder.exe|AdminUI.ConsoleBuilder.exe'
		.NOTES
			https://docs.microsoft.com/en-us/windows/win32/msi/file-table
		#>
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)][ValidateScript({Test-Path $_})][Alias('Installer')]
			$Msi, # The MSI to query
			[Parameter(Mandatory)][ValidateNotNullOrEmpty()]
			$FileName # The file to find the version of. Must be an exact match, in the Windows Installer's format including the shortname https://docs.microsoft.com/en-us/windows/win32/msi/filename.
		)

		begin {
			$windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
		}

		process {
			try {
				$database = $windowsInstaller.GetType().InvokeMember(
						"OpenDatabase", "InvokeMethod", $null,
						$windowsInstaller, @((Get-Item $Msi).FullName, 0)
					)

				$query = "SELECT FileName,Version FROM File WHERE FileName = '$filename'"
				$view = $database.GetType().InvokeMember(
						"OpenView", "InvokeMethod", $null, $database, $query
					)

				$view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null

				$record = $view.GetType().InvokeMember(
						"Fetch", "InvokeMethod", $null, $view, $null
					)

				while ($record -ne $null) {
					$fileName = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
					$version = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 2)

					Write-Output ([version]$version)

					$record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
				}

			} finally {
				$view.GetType().InvokeMember("Close", "InvokeMethod", $null, $view, $null) | Out-Null
			}
		}

	} # Get-MSISourceFileVersion

	function Invoke-VersionCheck {
		## Contact CM and determine if the Application Version is New
		[CmdletBinding()]
		param (
			[Parameter()]
			[String]
			$ApplicationName,
			[Parameter()]
			[String]
			$ApplicationSWVersion,
			[Parameter()]
			[Switch]
			# Require versions that can be parsed as a version or int to be higher than currently in CM as well as not previously added
			$RequireHigherVersion
		)

		Push-Location
		Set-Location $Global:CMSite
		If ($RequireHigherVersion -and ($ApplicationSWVersion -as [version])) {
			# Use [version] for proper sorting
			Add-LogContent "Requiring new version numbers to be higher than current"
			$currentHighest = Get-CMApplication -Name "$ApplicationName*" |
				Select-Object -ExpandProperty SoftwareVersion -ErrorAction SilentlyContinue |
				ForEach-Object {$_ -as [version]} |
				Sort-Object -Descending |
				Select-Object -First 1
			$newApp = ($ApplicationSWVersion -as [version]) -gt $currentHighest
			if ($newApp) {Add-LogContent "$ApplicationSWVersion is a new and higher version"}
			else {Add-LogContent "$ApplicationSWVersion is not new and higher - Moving to next application"}
		}
		ElseIf ($RequireHigherVersion -and ($ApplicationSWVersion -as [int])) {
			# Try [int]
			Add-LogContent "Requiring new version numbers to be higher than current"
			$currentHighest = Get-CMApplication -Name "$ApplicationName*" |
				Select-Object -ExpandProperty SoftwareVersion -ErrorAction SilentlyContinue |
				ForEach-Object {$_ -as [int]} |
				Sort-Object -Descending |
				Select-Object -First 1
			$newApp = ($ApplicationSWVersion -as [int]) -gt $currentHighest
			if ($newApp) {Add-LogContent "$ApplicationSWVersion is a new and higher version"}
			else {Add-LogContent "$ApplicationSWVersion is not new and higher - Moving to next application"}
		}
		ElseIf ((-not (Get-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -Fast)) -and (-not ([System.String]::IsNullOrEmpty($ApplicationSWVersion)))) {
			$newApp = $true			
			Add-LogContent "$ApplicationSWVersion is a new Version"
		}
		Else {
			$newApp = $false
			Add-LogContent "$ApplicationSWVersion is not a new Version - Moving to next application"
		}
        
		# If SkipPackaging is specified, return that the app is up-to-date.
		if ($ApplicationSWVersion -eq "SkipPackaging") {
			$newApp = $false
		}

		Pop-Location
		Write-Output $newApp
	}

	Function Start-ApplicationDownload {
		Param (
			$Recipe
		)
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		$ApplicationPublisher = $Recipe.ApplicationDef.Application.Publisher

		ForEach ($Download In $Recipe.ApplicationDef.Downloads.ChildNodes) {
			## Set Variables
			$newApp = $false
			$DownloadFileName = $Download.DownloadFileName
			$URL = $Download.URL
			$DownloadVersionCheck = $Download.DownloadVersionCheck
			$DownloadFile = "$TempDir\$DownloadFileName"
			$AppRepoFolder = $Download.AppRepoFolder
			$ExtraCopyFunctions = $Download.ExtraCopyFunctions
			$RequireHigherVersion = [System.Convert]::ToBoolean($Download.RequireHigherVersion)

			## Run the prefetch script if it exists, the prefetch script can be used to determine the location of the download URL, and optionally provide
			## the software version before the download occurs
			$PrefetchScript = $Download.PrefetchScript
			If (-not ([String]::IsNullOrEmpty($PrefetchScript))) {
				$ProgressPreference = 'SilentlyContinue'
				try {
					Invoke-Expression $PrefetchScript | Out-Null
				} catch {
					$errMsg = $_.Exception.Message
					Add-LogContent "ERROR: PrefetchScript failed for $($ApplicationName)"
					Add-LogContent "ERROR: $errMsg"
					Write-Host "ERROR: PrefetchScript failed for $($ApplicationName)" -ForegroundColor Red
					Write-Host $errMsg -ForegroundColor Red
					if ($Global:NotifyOnDownloadFailure) {
						$Global:SendEmail = $true; $Global:SendEmail | Out-Null
						$Global:EmailBody += "   - PrefetchScript failed for $($ApplicationName): $errMsg`n"
					}
					continue
				}
			}

			if (-not ([System.String]::IsNullOrEmpty($Download.Version))) {
				## Version Check after prefetch script (skip download if possible)
				## To Set the Download Version in the Prefetch Script, Simply set the variable $Download.Version to the [String]Version of the Application
				$ApplicationSWVersion = $Download.Version
				Add-LogContent "Prefetch Script Provided a Download Version of: $ApplicationSWVersion"
				$newApp = Invoke-VersionCheck -ApplicationName $ApplicationName -ApplicationSWVersion ([string]$ApplicationSWVersion) -RequireHigherVersion:$RequireHigherVersion
			}
			else {
				$newApp = $true
			}

			Add-LogContent "Version Check after prefetch script is $newapp"
			if ($newApp) {
				Add-LogContent "$ApplicationName will be downloaded"
			}
			else {
				Add-LogContent "$ApplicationName will not be downloaded"
			}

			## Download the Application
			If ((-not ([String]::IsNullOrEmpty($URL))) -and ($newapp)) {
				Add-LogContent "Downloading $ApplicationName from $URL"
				$ProgressPreference = 'SilentlyContinue'
				$iwrParams = @{ Uri = "$URL"; OutFile = $DownloadFile }
				if ($HTTPheaders) { $iwrParams['Headers'] = $HTTPheaders }
				if ($PSVersionTable.PSVersion.Major -ge 7) { $iwrParams['AllowInsecureRedirect'] = $true }
				Invoke-WebRequest @iwrParams | Out-Null
				Add-LogContent "Completed Downloading $ApplicationName"

				## Run the Version Check Script and record the Version and FullVersion
				If (-not ([String]::IsNullOrEmpty($DownloadVersionCheck))) {
					Invoke-Expression $DownloadVersionCheck | Out-Null
					$Download.Version = [string]$Version
					$Download.FullVersion = [string]$FullVersion
				}

				$ApplicationSWVersion = $Download.Version
				Add-LogContent "Found Version $ApplicationSWVersion from Download FullVersion: $FullVersion"
			}
			else {
				if (-not $newApp) {
					Add-LogContent "$Version was found in ConfigMgr, Skipping Download"
				}
				if ([String]::IsNullOrEmpty($URL)) {
					Add-LogContent "URL Not Specified, Skipping Download"
				}
			}

			## Determine if the Download Failed or if an Application Version was not detected, and add the Failure to the email if the Flag is set
			if (((-not (Test-Path $DownloadFile)) -and $newApp) -or ([System.String]::IsNullOrEmpty($ApplicationSWVersion))) {
				Add-LogContent "ERROR: Failed to Download or find the Version for $ApplicationName"
				if ($Global:NotifyOnDownloadFailure) {
					$Global:SendEmail = $true; $Global:SendEmail | Out-Null
					$Global:EmailBody += "   - Failed to Download: $ApplicationName`n"
				}
			}
		
			$newApp = Invoke-VersionCheck -ApplicationName $ApplicationName -ApplicationSWVersion $ApplicationSWVersion -RequireHigherVersion:$RequireHigherVersion
		
			## Create the Application folders and copy the download if the Application is New
			If ($newapp) {
				## Create Application Share Folder
				$ContentPath = "$Global:ContentLocationRoot\$ApplicationName\Packages\$Version"
				if ($Global:ContentFolderPattern) {	
					$ContentFolderPatternReplace = $Global:ContentFolderPattern -Replace '\$ApplicationName',$ApplicationName -Replace '\$Publisher',$ApplicationPublisher -Replace '\$Version',$Version
					$ContentPath = "$Global:ContentLocationRoot\$ContentFolderPatternReplace"
				}

				If ([String]::IsNullOrEmpty($AppRepoFolder)) {
					$DestinationPath = $ContentPath
					Add-LogContent "Destination Path set as $DestinationPath"
				}
				Else {
					$DestinationPath = "$ContentPath\$AppRepoFolder"
					Add-LogContent "Destination Path set as $DestinationPath"
				}
				New-Item -ItemType Directory -Path $DestinationPath -Force
			
				## Copy to Download to Application Share
				Add-LogContent "Copying downloads to $DestinationPath"
				Copy-Item -Path $DownloadFile -Destination $DestinationPath -Force
			
				## Extra Copy Functions If Required
				If (-not ([String]::IsNullOrEmpty($ExtraCopyFunctions))) {
					Add-LogContent "Performing Extra Copy Functions"
					Invoke-Expression $ExtraCopyFunctions | Out-Null
				}
			}
		}
	
		## Return True if All Downloaded Applications were new Versions
		Return $NewApp
	}

	Function Invoke-ApplicationCreation {
		Param (
			$Recipe
		)
	
		## Set Variables
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		$ApplicationPublisher = $Recipe.ApplicationDef.Application.Publisher
		$ApplicationDescription = $Recipe.ApplicationDef.Application.Description
		$ApplicationAdminDescription = $Recipe.ApplicationDef.Application.AdminDescription
		$ApplicationDocURL = $Recipe.ApplicationDef.Application.UserDocumentation
		$ApplicationOptionalReference = $Recipe.ApplicationDef.Application.OptionalReference
		$ApplicationLinkText = $Recipe.ApplicationDef.Application.LinkText
		$ApplicationPrivacyUrl = $Recipe.ApplicationDef.Application.PrivacyUrl
		$ApplicationFolderPath = $Recipe.ApplicationDef.Application.FolderPath
		$ApplicationOwner = $Recipe.ApplicationDef.Application.Owner
		$ApplicationSupportContact = $Recipe.ApplicationDef.Application.SupportContact
		$ApplicationKeywords = $Recipe.ApplicationDef.Application.Keywords
		$ApplicationUserCategories = $Recipe.ApplicationDef.Application.UserCategories
		$ApplicationAdminCategories = $Recipe.ApplicationDef.Application.AdminCategories
		$ApplicationIcon = $Recipe.ApplicationDef.Application.Icon
		$LocalizedName = $Recipe.ApplicationDef.Application.LocalizedName
		$ApplicationAutoInstall = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Application.AutoInstall)
		$ApplicationDisplaySupersedence = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Application.DisplaySupersedence)
		$ApplicationIsFeatured = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Application.FeaturedApplication)
		$AppCreated = $true
	
		ForEach ($Download In ($Recipe.ApplicationDef.Downloads.Download)) {
			If (-not ([System.String]::IsNullOrEmpty($Download.Version))) {
				$ApplicationSWVersion = $Download.Version		
			}
		}
	
		## Create the Application
		Push-Location
		Set-Location $Global:CMSite
		Add-LogContent "Creating Application: $ApplicationName $ApplicationSWVersion"

		# Change the SW Center Display Name based on Setting
		$ApplicationDisplayName = if ($LocalizedName) {$LocalizedName} else {$ApplicationName}
		if (!$Global:NoVersionInSWCenter) { $ApplicationDisplayName += " $ApplicationSWVersion"}

		Add-LogContent "Building application import command"

		# Because I (also) hate the yellow squiggly lines
		Write-Output $ApplicationDisplayName, $ApplicationPublisher, $ApplicationAutoInstall, $ApplicationDisplaySupersedence, $ApplicationIsFeatured | Out-Null

		# Reference: https://docs.microsoft.com/en-us/powershell/module/configurationmanager/new-cmapplication
		$NewAppCommand = 'New-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -LocalizedName "$ApplicationDisplayName" -SoftwareVersion "$ApplicationSWVersion" -ReleaseDate $(Get-Date) -AutoInstall $ApplicationAutoInstall -DisplaySupersedenceInApplicationCatalog $ApplicationDisplaySupersedence -IsFeatured $ApplicationIsFeatured'
		$CmdSwitches = ''
	
		## Build the rest of the command based on values in the xml
		If (-not ([System.String]::IsNullOrEmpty($ApplicationPublisher)))  {
			$CmdSwitches += ' -Publisher "$ApplicationPublisher"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationDescription)))  {
			$CmdSwitches += ' -LocalizedDescription "$ApplicationDescription"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationIcon)))  {
			if (Test-Path "$Global:IconRepo\$ApplicationIcon") {
				$CmdSwitches += " -IconLocationFile ""$Global:IconRepo\$ApplicationIcon"""
			} elseif (Test-Path "$ScriptRoot\ExtraFiles\Icons\$ApplicationIcon") {
				$CmdSwitches += " -IconLocationFile ""$ScriptRoot\ExtraFiles\Icons\$ApplicationIcon"""
			} else {
				Add-LogContent "ERROR: Unable to find icon $ApplicationIcon, creating application without icon"
			}
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationDocURL)))  {
			$CmdSwitches += ' -UserDocumentation "$ApplicationDocURL"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationOptionalReference)))  {
			$CmdSwitches += ' -OptionalReference "$ApplicationOptionalReference"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationAdminDescription)))  {
			$CmdSwitches += ' -Description "$ApplicationAdminDescription"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationOwner)))  {
			$CmdSwitches += ' -Owner "$ApplicationOwner"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationSupportContact)))  {
			$CmdSwitches += ' -SupportContact "$ApplicationSupportContact"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationKeywords)))  {
			$CmdSwitches += ' -Keyword "$ApplicationKeywords"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationLinkText)))  {
			$CmdSwitches += ' -LinkText "$ApplicationLinkText"'
		}

		If (-not ([System.String]::IsNullOrEmpty($ApplicationPrivacyUrl)))  {
			$CmdSwitches += ' -PrivacyUrl "$ApplicationPrivacyUrl"'
		}
	
		## Run the New-CMApplication Command
		$NewAppCommandFull = "$NewAppCommand$CmdSwitches"
		Add-LogContent "Command: $NewAppCommandFull"
		Try {
			Invoke-Expression $NewAppCommandFull | Out-Null
			Add-LogContent "Application Created"
		}
		Catch {
			$AppCreated = $false
			$ErrorMessage = $_.Exception.Message
			$FullyQualified = $_.Exeption.FullyQualifiedErrorID
			Add-LogContent "ERROR: Creating Application Failed!"
			Add-LogContent "ERROR: $ErrorMessage"
			Add-LogContent "ERROR: $FullyQualified"
			Add-LogContent "ERROR: $($_.CategoryInfo.Category): $($_.CategoryInfo.Reason)"
		}

		# Apply categories if supplied. This was not availabe during application creation
		if ($AppCreated) {
			Try {
				## Set user categories that display in Software Center
				If (-not ([System.String]::IsNullOrEmpty($ApplicationUserCategories))) {
					## Create list to store user categories
					$AppUserCatList = New-Object System.Collections.ArrayList
					foreach ($ApplicationUserCategory in ($ApplicationUserCategories).Split(",")) {
						if (-not (($AppUserCatObj = Get-CMCategory -Name $ApplicationUserCategory | Where-Object {$_.CategoryTypeName -eq "CatalogCategories"}))) {
							## Create if not found and add to list
							Add-LogContent "$ApplicationUserCategory category was supplied in recipe, but does not exist. Creating user category"
							$null = $AppUserCatList.Add((New-CMCategory -CategoryType "CatalogCategories" -Name $ApplicationUserCategory))
						} else {
							## Add to list
							$null = $AppUserCatList.Add($AppUserCatObj)
						}
					}
				}

				## Set administrative categories that display in admin console
				If (-not ([System.String]::IsNullOrEmpty($ApplicationAdminCategories))) {
					## Create list to store admin categories
					$AppAdminCatList = New-Object System.Collections.ArrayList
					foreach ($ApplicationAdminCategory in ($ApplicationAdminCategories).Split(",")) {
						if (-not (($AppAdminCatObj = Get-CMCategory -Name $ApplicationAdminCategory | Where-Object {$_.CategoryTypeName -eq "AppCategories"}))) {
							## Create if not found and add to list
							Add-LogContent "$ApplicationAdminCategory category was supplied in recipe, but does not exist. Creating admin category"
							$null = $AppAdminCatList.Add((New-CMCategory -CategoryType "AppCategories" -Name $ApplicationAdminCategory))
						} else {
							## Add to list
							$null = $AppAdminCatList.Add($AppAdminCatObj)
						}
					}
				}

				## Run Set-CMApplication depending on which types of categories exist
				## Reference: https://docs.microsoft.com/en-us/powershell/module/configurationmanager/set-cmapplication
				if (($AppUserCatList) -and ($AppAdminCatList)) {
					Set-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -AddAppCategory $AppAdminCatList -AddUserCategory $AppUserCatList
				} elseif ($AppUserCatList) {
					Set-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -AddUserCategory $AppUserCatList
				} elseif ($AppAdminCatList) {
					Set-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -AddAppCategory $AppAdminCatList
				}
			}
			Catch { 
				$AppCreated = $false
				$ErrorMessage = $_.Exception.Message
				$FullyQualified = $_.Exception.FullyQualifiedErrorID
				Add-LogContent "ERROR: Setting Application Categories Failed!"
				Add-LogContent "ERROR: $ErrorMessage"
				Add-LogContent "ERROR: $FullyQualified"
				Add-LogContent "ERROR: $($_.CategoryInfo.Category): $($_.CategoryInfo.Reason)"
			}
		}

		# Move the Application to folder path if supplied
		If ($AppCreated) {
			Try {
				If (-not ([System.String]::IsNullOrEmpty($ApplicationFolderPath))) {
					# Create the folder if it does not exist
					if (-not (Test-Path ".\Application\$ApplicationFolderPath")) {
						New-Item -ItemType Directory -Path ".\Application\$ApplicationFolderPath" -ErrorAction SilentlyContinue
					}
					Add-LogContent "Command: Move-CMObject -InputObject (Get-CMApplication -Name ""$ApplicationName $ApplicationSWVersion"") -FolderPath "".\Application\$ApplicationFolderPath"""
					Move-CMObject -InputObject (Get-CMApplication -Name "$ApplicationName $ApplicationSWVersion") -FolderPath ".\Application\$ApplicationFolderPath"
				}
			}
			Catch { 
				$AppCreated = $false
				$ErrorMessage = $_.Exception.Message
				$FullyQualified = $_.Exception.FullyQualifiedErrorID
				Add-LogContent "ERROR: Application Move Failed!"
				Add-LogContent "ERROR: $ErrorMessage"
				Add-LogContent "ERROR: $FullyQualified"
				Add-LogContent "ERROR: $($_.CategoryInfo.Category): $($_.CategoryInfo.Reason)"
			}
		}

		## Send an Email if an Application was successfully Created and record the Application Name and Version for the Email
		If ($AppCreated) {
			$Global:SendEmail = $true; $Global:SendEmail | Out-Null
			$Global:EmailBody += "   - $ApplicationName $ApplicationSWVersion`n"
		}
		Pop-Location
	
		## Return True if the Application was Created Successfully
		Return $AppCreated
	}

	Function Add-DetectionMethodClause {
		Param (
			$DetectionMethod,
			$AppVersion,
			$AppFullVersion
		)
	
		$detMethodDetectionClauseType = $DetectionMethod.DetectionClauseType
		Add-LogContent "Adding Detection Method Clause Type $detMethodDetectionClauseType"
		Switch ($detMethodDetectionClauseType) {
			Directory {
				$detMethodCommand = "New-CMDetectionClauseDirectory"
				If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Name))) {
					$DetectionMethod.Name = ($DetectionMethod.Name).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
					$detMethodCommand += " -DirectoryName `'$($DetectionMethod.Name)`'"
				}
			}
			File {
				$detMethodCommand = "New-CMDetectionClauseFile"
				If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Name))) {
					$DetectionMethod.Name = ($DetectionMethod.Name).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
					$detMethodCommand += " -FileName `'$($DetectionMethod.Name)`'"
				}
			}
			RegistryKey {
				$detMethodCommand = "New-CMDetectionClauseRegistryKey"
			}
			RegistryKeyValue {
				$detMethodCommand = "New-CMDetectionClauseRegistryKeyValue"
			
			}
			WindowsInstaller {
				$detMethodCommand = "New-CMDetectionClauseWindowsInstaller"
			}
		}
		If (([System.Convert]::ToBoolean($DetectionMethod.Existence)) -and (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Existence)))) {
			$detMethodCommand += " -Existence"
		}
		If (([System.Convert]::ToBoolean($DetectionMethod.Is64Bit)) -and (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Is64Bit)))) {
			$detMethodCommand += " -Is64Bit"
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Path))) {
			$DetectionMethod.Path = ($DetectionMethod.Path).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			$detMethodCommand += " -Path `'$($DetectionMethod.Path)`'"
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.PropertyType))) {
			$detMethodCommand += " -PropertyType $($DetectionMethod.PropertyType)"
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.ExpectedValue))) {
			$DetectionMethod.ExpectedValue = ($DetectionMethod.ExpectedValue).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			$detMethodCommand += " -ExpectedValue `"$($DetectionMethod.ExpectedValue)`""
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.ExpressionOperator))) {
			$detMethodCommand += " -ExpressionOperator $($DetectionMethod.ExpressionOperator)"
		}
		If (([System.Convert]::ToBoolean($DetectionMethod.Value)) -and (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Value)))) {
			$detMethodCommand += " -Value"
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.Hive))) {
			$detMethodCommand += " -Hive $($DetectionMethod.Hive)"
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.KeyName))) {
			$DetectionMethod.KeyName = ($DetectionMethod.KeyName).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			$detMethodCommand += " -KeyName `"$($DetectionMethod.KeyName)`""
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.ValueName))) {
			$detMethodCommand += " -ValueName `"$($DetectionMethod.ValueName)`""
		}
		If (-not ([System.String]::IsNullOrEmpty($DetectionMethod.ProductCode))) {
			$detMethodCommand += " -ProductCode `"$($DetectionMethod.ProductCode)`""
		}
		Add-LogContent "$detMethodCommand"
	
		## Run the Detection Method Command as Created by the Logic Above
	
		Push-Location
		Set-Location $CMSite
		Try {
			$DepTypeDetectionMethod += Invoke-Expression $detMethodCommand
		}
		Catch {
			$ErrorMessage = $_.Exception.Message
			$FullyQualified = $_.Exeption.FullyQualifiedErrorID
			Add-LogContent "ERROR: Creating Detection Method Clause Failed!"
			Add-LogContent "ERROR: $ErrorMessage"
			Add-LogContent "ERROR: $FullyQualified"
		}
		Pop-Location
	
		## Return the Detection Method Variable
		Return $DepTypeDetectionMethod
	}

	Function Copy-CMDeploymentTypeRule {
		<#
	Function taken from https://janikvonrotz.ch/2017/10/20/configuration-manager-configure-requirement-rules-for-deployment-types-with-powershell/ and modified
 	
     #>
		Param (
			[System.String]$SourceApplicationName,
			[System.String]$DestApplicationName,
			[System.String]$DestDeploymentTypeName,
			[System.String]$RuleName
		)
		Push-Location
		Set-Location $CMSite
		$DestDeploymentTypeIndex = 0
 
		# get the applications
		$SourceApplication = Get-CMApplication -Name $SourceApplicationName | ConvertTo-CMApplication
		$DestApplication = Get-CMApplication -Name $DestApplicationName | ConvertTo-CMApplication
	
		# Get DestDeploymentTypeIndex by finding the Title
		$DestDeploymentTypeIndex = $DestApplication.DeploymentTypes.Title.IndexOf($DestDeploymentTypeName)
    
		$Available = ($SourceApplication.DeploymentTypes[0].Requirements).Name
		Add-LogContent "Available Requirements to chose from:`r`n $($Available -Join ', ')"
    
		# get requirement rules from source application
		$Requirements = $SourceApplication.DeploymentTypes[0].Requirements | Where-Object { (($_.Name).TrimStart().TrimEnd()) -eq (($RuleName).TrimStart().TrimEnd()) }
		if ([System.String]::IsNullOrEmpty($Requirements)) {
			Add-LogContent "No Requirement rule was an exact match for $RuleName"
			$Requirements = $SourceApplication.DeploymentTypes[0].Requirements | Where-Object { $_.Name -match $RuleName }
		}
		if ([System.String]::IsNullOrEmpty($Requirements)) {
			Add-LogContent "No Requirement rule was matched, tring one more thing for $RuleName"
			$Requirements = $SourceApplication.DeploymentTypes[0].Requirements | Where-Object { $_.Name -like $RuleName }
		}
		Add-LogContent "$($Requirements.Name) will be added"

		# apply requirement rules
		$Requirements | ForEach-Object {
     
			$RuleExists = $DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Requirements | Where-Object { $_.Name -match $RuleName }
			if ($RuleExists) {
 
				Add-LogContent "WARN: The rule `"$($_.Name)`" already exists in target application deployment type"
 
			}
			else {
         
				Add-LogContent "Apply rule `"$($_.Name)`" on target application deployment type"
 
				# create new rule ID
				$_.RuleID = "Rule_$( [guid]::NewGuid())"
 
				$DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Requirements.Add($_)
			}
		}
 
		# push changes
		$CMApplication = ConvertFrom-CMApplication -Application $DestApplication
		$CMApplication.Put()
		Pop-Location
	}

	function Add-RequirementsRule {
		[CmdletBinding()]
		param (
			[Parameter(Mandatory)]
			[ValidateSet('Value', 'Existential', 'OperatingSystem')]
			[String]
			$ReqRuleType,
			[Parameter()]
			[ValidateSet( 'And', 'Or', 'Other', 'IsEquals', 'NotEquals', 'GreaterThan', 'LessThan', 'Between', 'NotBetween', 'GreaterEquals', 'LessEquals', 'BeginsWith', 'NotBeginsWith', 'EndsWith', 'NotEndsWith', 'Contains', 'NotContains', 'AllOf', 'OneOf', 'NoneOf', 'SetEquals', 'SubsetOf', 'ExcludesAll')]
			$ReqRuleOperator,
			[Parameter(Mandatory)]
			[String[]]
			$ReqRuleValue,
			[Parameter()]
			[String[]]
			$ReqRuleValue2,
			[Parameter()]
			[String]
			$ReqRuleGlobalConditionName,
			[Parameter(Mandatory)]
			[String]
			$ReqRuleApplicationName,
			[Parameter(Mandatory)]
			[String]
			$ReqRuleApplicationDTName
		)
		
		Push-Location
		Set-Location $Global:CMSite
		Write-Host "`"$ReqRuleType of $ReqRuleGlobalConditionName $ReqRuleOperator $ReqRuleValue`" is being added"

		if (-not ([System.String]::IsNullOrEmpty($ReqRuleValue))) {
			$ReqRuleValueName = $ReqRuleValue
			#if (($ReqRuleOperator -eq 'Oneof') -or ($ReqRuleOperator -eq 'Noneof') -or ($ReqRuleOperator -eq 'Allof') -or ($ReqRuleOperator -eq 'Subsetof') -or ($ReqRuleOperator -eq 'ExcludesAll')) {
			if ($ReqRuleValue[1]) {
				$ReqRuleVal = $ReqRuleValue
				$ReqRuleValueName = "{ $($ReqRuleVal -join ", ") }"
			}
			if ([system.string]::IsNullOrEmpty($ReqRuleVal)) {
				$ReqRuleVal = $ReqRuleValue[0]
			}
		}
		
		if (-not ([System.String]::IsNullOrEmpty($ReqRuleValue2))) {
			if ($ReqRuleValue2[1]) {
				$ReqRuleVal2 = $ReqRuleValue2
				$ReqRuleValue2Name = "{ $($ReqRuleVal2 -join ", ") }"
			}
			if ([system.string]::IsNullOrEmpty($ReqRuleVal)) {
				$ReqRuleVal2 = $ReqRuleValue2[0]
			}
		}

		switch ($ReqRuleType) {
			Existential {
				Add-LogContent "Existential Rule $ReqRuleVal"
				$CMGlobalCondition = Get-CMGlobalCondition -Name $ReqRuleGlobalConditionName
				if ([System.Convert]::ToBoolean($ReqRuleVal)) {
					$rule = $CMGlobalCondition | New-CMRequirementRuleExistential -Existential $([System.Convert]::ToBoolean($($ReqRuleVal | Select-Object -first 1)))
					$rule.Name = "Existential of $ReqRuleGlobalConditionName Not equal to 0"
				}
				else {
					$rule = $CMGlobalCondition | New-CMRequirementRuleExistential -Existential $([System.Convert]::ToBoolean($($ReqRuleVal | Select-Object -first 1)))
					$rule.Name = "Existential of $ReqRuleGlobalConditionName Equals 0"
				}
			}
			OperatingSystem {
				Add-LogContent "Operating System $ReqRuleOperator `"$ReqruleVal`""
				# Only supporting Windows Operating Systems at this time
				$GlobalCondition = Get-CMGlobalCondition -name "Operating System" | Where-Object PlatformType -eq 1
				$rule = $GlobalCondition | New-CMRequirementRuleOperatingSystemValue -RuleOperator $ReqRuleOperator -PlatformStrings $ReqRuleVal
				$rule.Name = "Operating System $Global:OperatorsLookup $ReqRuleValueName"
			}
			Default {
				# DEFAULT TO VALUE
				Add-LogContent "Value $ReqRuleOperator `"$ReqRuleVal`""
				$CMGlobalCondition = Get-CMGlobalCondition -Name $ReqRuleGlobalConditionName
				if ([System.String]::IsNullOrEmpty($ReqRuleValue2)) {
					$rule = $CMGlobalCondition | New-CMRequirementRuleCommonValue -Value1 $ReqRuleVal -RuleOperator $ReqRuleOperator
					$rule.Name = "$ReqRuleGlobalConditionName $Global:OperatorsLookup $ReqRuleValueName"
				}
				else {
					$rule = $CMGlobalCondition | New-CMRequirementRuleCommonValue -Value1 $ReqRuleVal -RuleOperator $ReqRuleOperator -Value2 $ReqRuleVal2
					$rule.Name = "$ReqRuleGlobalConditionName $Global:OperatorsLookup $ReqRuleValueName $ReqRuleValue2Name"
				}
			}
		}

		Add-LogContent "Adding Requirement to $ReqRuleApplicationName, $ReqRuleApplicationDTName"
		Get-CMDeploymentType -ApplicationName $ReqRuleApplicationName -DeploymentTypeName $ReqRuleApplicationDTName | Set-CMDeploymentType -AddRequirement $rule
		Pop-Location
	}


	Function Add-CMDeploymentTypeProcessDetection {
		# Creates a Deployment Type Process Detection "Install Behavior tab in Deployment types".
		Param (
			[System.String]$DestApplicationName,
			[System.String]$DestDeploymentTypeName,
			[System.String]$ProcessDetectionDisplayName,
			[System.String]$ProcessDetectionExecutable
		)
		Push-Location
		Set-Location $CMSite
		$DestDeploymentTypeIndex = 0
 
		# get the applications
		$DestApplication = Get-CMApplication -Name $DestApplicationName | ConvertTo-CMApplication
	
		# Get DestDeploymentTypeIndex by finding the Title
		$DestDeploymentTypeIndex = $DestApplication.DeploymentTypes.Title.IndexOf($DestDeploymentTypeName)
    
		# Create Process Detection and set variables
		$ProcessInfo = [Microsoft.ConfigurationManagement.ApplicationManagement.ProcessInformation]::new()
		$ProcessInfo.DisplayInfo.Add(@{"DisplayName" = $ProcessDetectionDisplayName; Language = $NULL })
		$ProcessInfo.Name = $ProcessDetectionExecutable
 
		# push changes
		$DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Installer.InstallProcessDetection.ProcessList.Add($ProcessInfo)
		$CMApplication = ConvertFrom-CMApplication -Application $DestApplication
		$CMApplication.Put()
		Pop-Location
	}

	Function New-CMDeploymentTypeProcessRequirement {
		# Creates a Deployment Type Process Requirement "Install Behavior tab in Deployment types" by copying an existing Process Requirement.
		# LEGACY
		Param (
			[System.String]$SourceApplicationName,
			[System.String]$DestApplicationName,
			[System.String]$DestDeploymentTypeName,
			[System.String]$ProcessDetectionDisplayName,
			[System.String]$ProcessDetectionExecutable
		)
		Push-Location
		Set-Location $CMSite
		$DestDeploymentTypeIndex = 0
 
		# get the applications
		$SourceApplication = Get-CMApplication -Name $SourceApplicationName | ConvertTo-CMApplication
		$DestApplication = Get-CMApplication -Name $DestApplicationName | ConvertTo-CMApplication
	
		# Get DestDeploymentTypeIndex by finding the Title
		$DestDeploymentTypeIndex = $DestApplication.DeploymentTypes.Title.IndexOf($DestDeploymentTypeName)
    
		# Get requirement rules from source application
		$ProcessRequirementsList = $SourceApplication.DeploymentTypes[0].Installer.InstallProcessDetection.ProcessList[0]
		$ProcessRequirementsList
		if (-not ([System.String]::IsNullOrEmpty($ProcessRequirementsList))) {
			$ProcessRequirementsList.Name = $ProcessDetectionExecutable
			$ProcessRequirementsList.DisplayInfo[0].DisplayName = $ProcessDetectionDisplayName
			$ProcessRequirementsList
			$DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Installer.InstallProcessDetection.ProcessList.Add($ProcessRequirementsList)
		}
 
		# push changes
		$CMApplication = ConvertFrom-CMApplication -Application $DestApplication
		$CMApplication.Put()
		Pop-Location
	}

	Function Add-DeploymentType {
		Param (
			$Recipe
		)
	
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		$ApplicationPublisher = $Recipe.ApplicationDef.Application.Publisher
		$ApplicationDescription = $Recipe.ApplicationDef.Application.Description
		$ApplicationDocURL = $Recipe.ApplicationDef.Application.UserDocumentation
	
		## Set Return Value to True, It will toggle to False if something Fails
		$DepTypeReturn = $true
	
		## Loop through each Deployment Type and Add them to the Application as needed
		ForEach ($DeploymentType In $Recipe.ApplicationDef.DeploymentTypes.ChildNodes) {
			$DepTypeName = $DeploymentType.Name
			$DepTypeDeploymentTypeName = $DeploymentType.DeploymentTypeName
			Add-LogContent "New DeploymentType - $DepTypeDeploymentTypeName"
		
			$AssociatedDownload = $Recipe.ApplicationDef.Downloads.Download | Where-Object DeploymentType -eq $DepTypeName
			$ApplicationSWVersion = $AssociatedDownload.Version
			$Version = $AssociatedDownload.Version
			If (-not ([String]::IsNullOrEmpty($AssociatedDownload.FullVersion))) {
				$FullVersion = $AssociatedDownload.FullVersion
				$AppFullVersion = $AssociatedDownload.FullVersion
			}
		
			# General
			$DepTypeApplicationName = "$ApplicationName $ApplicationSWVersion"
			$DepTypeInstallationType = $DeploymentType.InstallationType
			Add-LogContent "Deployment Type Set as: $DepTypeInstallationType"
		
			$stDepTypeComment = $DeploymentType.Comments
			$DepTypeLanguage = $DeploymentType.Language
		
			# Content Settings
			# Content Location
			$ContentPath = "$Global:ContentLocationRoot\$ApplicationName\Packages\$Version"
			if ($Global:ContentFolderPattern) {	
				$ContentFolderPatternReplace = $Global:ContentFolderPattern -Replace '\$ApplicationName',$ApplicationName -Replace '\$Publisher',$ApplicationPublisher -Replace '\$Version',$Version
				$ContentPath = "$Global:ContentLocationRoot\$ContentFolderPatternReplace"
			}

			If ([String]::IsNullOrEmpty($AssociatedDownload.AppRepoFolder)) {
				$DepTypeContentLocation = $ContentPath
			}
			Else {
				$DepTypeContentLocation = "$ContentPath\$($AssociatedDownload.AppRepoFolder)"
			}
			$swDepTypeCacheContent = [System.Convert]::ToBoolean($DeploymentType.CacheContent)
			$swDepTypeEnableBranchCache = [System.Convert]::ToBoolean($DeploymentType.BranchCache)
			$swDepTypeContentFallback = [System.Convert]::ToBoolean($DeploymentType.ContentFallback)
			$stDepTypeSlowNetworkDeploymentMode = $DeploymentType.OnSlowNetwork
			$stDepTypeUninstallOption = $DeploymentType.UninstallOption
			$stDepTypeUninstallContentLocation = $DeploymentType.UninstallContentLocation
		
			# Programs
			if (-not ([System.String]::IsNullOrEmpty($DeploymentType.InstallProgram))) {
				$stDepTypeInstallCommand = ($DeploymentType.InstallProgram).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			}
			
			if (-not ([System.String]::IsNullOrEmpty($DeploymentType.UninstallCmd))) {
				$stDepTypeUninstallationProgram = $DeploymentType.UninstallCmd
				$stDepTypeUninstallationProgram = ($stDepTypeUninstallationProgram).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			}

			if (-not ([System.String]::IsNullOrEmpty($DeploymentType.RepairCmd))) {
				$stDepTypeRepairCommand = ($DeploymentType.RepairCmd).replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
			}
			$swDepTypeForce32Bit = [System.Convert]::ToBoolean($DeploymentType.Force32bit)
		
			# User Experience
			$stDepTypeInstallationBehaviorType = $DeploymentType.InstallationBehaviorType
			$stDepTypeLogonRequirementType = $DeploymentType.LogonReqType
			$stDepTypeUserInteractionMode = $DeploymentType.UserInteractionMode
			$swDepTypeRequireUserInteraction = [System.Convert]::ToBoolean($DeploymentType.ReqUserInteraction)
			$stDepTypeEstimatedRuntimeMins = $DeploymentType.EstRuntimeMins
			$stDepTypeMaximumRuntimeMins = $DeploymentType.MaxRuntimeMins
			$stDepTypeRebootBehavior = $DeploymentType.RebootBehavior
		
			# Because I hate the yellow squiggly lines
			Write-Output $ApplicationPublisher, $ApplicationDescription, $ApplicationDocURL, $DepTypeLanguage, $stDepTypeComment, $swDepTypeCacheContent, $swDepTypeEnableBranchCache, $swDepTypeContentFallback, $stDepTypeSlowNetworkDeploymentMode, $swDepTypeForce32Bit, $stDepTypeInstallationBehaviorType, $stDepTypeLogonRequirementType, $stDepTypeUserInteractionMode$swDepTypeRequireUserInteraction, $stDepTypeEstimatedRuntimeMins, $stDepTypeMaximumRuntimeMins, $stDepTypeRebootBehavior | Out-Null

			$DepTypeDetectionMethodType = $DeploymentType.DetectionMethodType
			Add-LogContent "Detection Method Type Set as $DepTypeDetectionMethodType"
			$DepTypeAddDetectionMethods = $false
		
			If (($DepTypeDetectionMethodType -eq "Custom") -and (-not ([System.String]::IsNullOrEmpty($DeploymentType.CustomDetectionMethods.ChildNodes)))) {
				$DepTypeDetectionMethods = @()
				$DepTypeAddDetectionMethods = $true
				$DepTypeDetectionClauseConnector = @()
				Add-LogContent "Adding Detection Method Clauses"
				ForEach ($DetectionMethod In $($DeploymentType.CustomDetectionMethods.ChildNodes | Where-Object Name -NE "DetectionClauseExpression")) {
					Add-LogContent "New Detection Method Clause $Version $FullVersion"
					$DepTypeDetectionMethods += Add-DetectionMethodClause -DetectionMethod $DetectionMethod -AppVersion $Version -AppFullVersion $FullVersion
				}
				if (-not [System.string]::IsNullOrEmpty($($DeploymentType.CustomDetectionMethods.ChildNodes | Where-Object Name -EQ "DetectionClauseExpression"))) {
					$CustomDetectionMethodExpression = ($DeploymentType.CustomDetectionMethods.ChildNodes | Where-Object Name -EQ "DetectionClauseExpression").ChildNodes
				}
				ForEach ($DetectionMethodExpression In $CustomDetectionMethodExpression) {
					if ($DetectionMethodExpression.Name -eq "DetectionClauseConnector") {
						Add-LogContent "New Detection Clause Connector $($DetectionMethodExpression.ConnectorClause),$($DetectionMethodExpression.ConnectorClauseConnector)"
						$DepTypeDetectionClauseConnector += @{"LogicalName" = $DepTypeDetectionMethods[$DetectionMethodExpression.ConnectorClause].Setting.LogicalName; "Connector" = "$($DetectionMethodExpression.ConnectorClauseConnector)" }
					}
					if ($DetectionMethodExpression.Name -eq "DetectionClauseGrouping") {
						Add-LogContent "New Detection Clause Grouping Statement Found - NOT READY YET"
					}
				}
			}
		
			Switch ($DepTypeInstallationType) {
				Script {
					Write-Host "Script Deployment"
					$DepTypeCommand = "Add-CMScriptDeploymentType -ApplicationName `"$DepTypeApplicationName`" -ContentLocation `"$DepTypeContentLocation`" -DeploymentTypeName `"$DepTypeDeploymentTypeName`""
					$CmdSwitches = ""
				
					## Build the Rest of the command based on values in the xml
					## Switch type Arguments
					ForEach ($DepTypeVar In $(Get-Variable | Where-Object {
								$_.Name -like "swDepType*"
							})) {
						If (([System.Convert]::ToBoolean($deptypevar.Value)) -and (-not ([System.String]::IsNullOrEmpty($DepTypeVar.Value)))) {
							$CmdSwitch = "-$($($DepTypeVar.Name).Replace("swDepType", ''))"
							$CmdSwitches += " $CmdSwitch"
						}
					}
				
					## String Type Arguments
					ForEach ($DepTypeVar In $(Get-Variable | Where-Object {
								$_.Name -like "stDepType*"
							})) {
						If (-not ([System.String]::IsNullOrEmpty($DepTypeVar.Value))) {
							$CmdSwitch = "-$($($DepTypeVar.Name).Replace("stDepType", '')) `'$($DepTypeVar.Value)`'"
							$CmdSwitches += " $CmdSwitch"
						}
					}
				
					If ($DepTypeDetectionMethodType -eq "CustomScript") {
						$DepTypeScriptLanguage = $DeploymentType.ScriptLanguage
						If (-not ([string]::IsNullOrEmpty($DepTypeScriptLanguage))) {
							$CMDSwitch = "-ScriptLanguage `"$DepTypeScriptLanguage`""
							$CmdSwitches += " $CmdSwitch"
						}
					
						$DepTypeScriptText = ($DeploymentType.DetectionMethod).Replace('$Version', $Version).replace('$FullVersion', $AppFullVersion)
						If (-not ([string]::IsNullOrEmpty($DepTypeScriptText))) {
							$CMDSwitch = "-ScriptText `'$DepTypeScriptText`'"
							$CmdSwitches += " $CmdSwitch"
						}
					}
				
					$DepTypeForce32BitDetection = $DeploymentType.ScriptDetection32Bit
					If (([System.Convert]::ToBoolean($DepTypeForce32BitDetection)) -and (-not ([System.String]::IsNullOrEmpty($DepTypeForce32BitDetection)))) {
						$CmdSwitches += " -ForceScriptDetection32Bit"
					}
				
					## Run the Add-CMApplicationDeployment Command
					$DeploymentTypeCommand = "$DepTypeCommand$CmdSwitches"
					If ($DepTypeAddDetectionMethods) {
						$DeploymentTypeCommand += " -ScriptType Powershell -ScriptText `"write-output 0`""
					}
					Add-LogContent "Creating DeploymentType"
					Add-LogContent "Command: $DeploymentTypeCommand"
					Push-Location
					Set-Location $CMSite
					Try {
						Invoke-Expression $DeploymentTypeCommand | Out-Null
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						$FullyQualified = $_.Exeption.FullyQualifiedErrorID
						Add-LogContent "ERROR: Creating Deployment Type Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						Add-LogContent "ERROR: $FullyQualified"
						$DepTypeReturn = $false
					}
				
					## Add Detection Methods if required for this Deployment Type
					If ($DepTypeAddDetectionMethods) {
						Add-LogContent "Adding Detection Methods"
					
						Add-LogContent "Number of Detection Methods: $($DepTypeDetectionMethods.Count)"
						if ($DepTypeDetectionMethods.Count -eq 1) {
					
							Add-LogContent "Set-CMScriptDeploymentType -ApplicationName $DepTypeApplicationName -DeploymentTypeName $DepTypeDeploymentTypeName -AddDetectionClause $($DepTypeDetectionMethods[0].DataType.Name)"
							Try {
								Set-CMScriptDeploymentType -ApplicationName "$DepTypeApplicationName" -DeploymentTypeName "$DepTypeDeploymentTypeName" -AddDetectionClause $DepTypeDetectionMethods
							}
							Catch {
								Write-Host $_
								$ErrorMessage = $_.Exception.Message
								$FullyQualified = $_.Exeption.FullyQualifiedErrorID
								Add-LogContent "ERROR: Adding Detection Method Failed!"
								Add-LogContent "ERROR: $ErrorMessage"
								Add-LogContent "ERROR: $FullyQualified"
								$DepTypeReturn = $false
							}
						} 
						Else {
							Add-LogContent "Set-CMScriptDeploymentType -ApplicationName $DepTypeApplicationName -DeploymentTypeName $DepTypeDeploymentTypeName -AddDetectionClause $($DepTypeDetectionMethods[0].DataType.Name) -DetectionClauseConnector $DepTypeDetectionClauseConnector"
							Try {	
								Set-CMScriptDeploymentType -ApplicationName "$DepTypeApplicationName" -DeploymentTypeName "$DepTypeDeploymentTypeName" -AddDetectionClause $DepTypeDetectionMethods -DetectionClauseConnector $DepTypeDetectionClauseConnector
							}
							Catch {
								Write-Host $_
								$ErrorMessage = $_.Exception.Message
								$FullyQualified = $_.Exeption.FullyQualifiedErrorID
								Add-LogContent "ERROR: Adding Detection Method Failed!"
								Add-LogContent "ERROR: $ErrorMessage"
								Add-LogContent "ERROR: $FullyQualified"
								$DepTypeReturn = $false
							}	
						}		
					}
					Pop-Location	
				}
				MSI {
					Write-Host "MSI Deployment"
					$DepTypeInstallationMSI = $DeploymentType.InstallationMSI
					$DepTypeCommand = "Add-CMMsiDeploymentType -ApplicationName `"$DepTypeApplicationName`" -ContentLocation `"$DepTypeContentLocation\$DepTypeInstallationMSI`" -DeploymentTypeName `"$DepTypeDeploymentTypeName`""
					$CmdSwitches = ""

					## Build the Rest of the command based on values in the xml
					ForEach ($DepTypeVar In $(Get-Variable | Where-Object {
								$_.Name -like "swDepType*"
							})) {
						If (([System.Convert]::ToBoolean($deptypevar.Value)) -and (-not ([System.String]::IsNullOrEmpty($DepTypeVar.Value)))) {
							$CmdSwitch = "-$($($DepTypeVar.Name).Replace("swDepType", ''))"
							$CmdSwitches += " $CmdSwitch"
						}
					}
				
					ForEach ($DepTypeVar In $(Get-Variable | Where-Object {
								$_.Name -like "stDepType*"
							})) {
						If (-not ([System.String]::IsNullOrEmpty($DepTypeVar.Value))) {
							$CmdSwitch = "-$($($DepTypeVar.Name).Replace("stDepType", '')) `"$($DepTypeVar.Value)`""
							$CmdSwitches += " $CmdSwitch"
						}
					}
				
					## Special Arguments based on Detection Method
					Switch ($DepTypeDetectionMethodType) {
						MSI {
							$DepTypeProductCode = $DeploymentType.ProductCode
							If (-not ([string]::IsNullOrEmpty($DepTypeProductCode))) {
								$CMDSwitch = "-ProductCode `"$DepTypeProductCode`""
								$CmdSwitches += " $CmdSwitch"
							}
						}
						CustomScript {
							$DepTypeScriptLanguage = $DeploymentType.ScriptLanguage
							If (-not ([string]::IsNullOrEmpty($DepTypeScriptLanguage))) {
								$CMDSwitch = "-ScriptLanguage `"$DepTypeScriptLanguage`""
								$CmdSwitches += " $CmdSwitch"
							}
						
							$DepTypeForce32BitDetection = $DeploymentType.ScriptDetection32Bit
							If (([System.Convert]::ToBoolean($DepTypeForce32BitDetection)) -and (-not ([System.String]::IsNullOrEmpty($DepTypeForce32BitDetection)))) {
								$CmdSwitches += " -ForceScriptDetection32Bit"
							}
						
							$DepTypeScriptText = ($DeploymentType.DetectionMethod).Replace("REPLACEMEWITHTHEAPPVERSION", $($AssociatedDownload.Version))
							If (-not ([string]::IsNullOrEmpty($DepTypeScriptText))) {
								$CMDSwitch = "-ScriptText `'$DepTypeScriptText`'"
								$CmdSwitches += " $CmdSwitch"
							}
						}
					}
				
					## Run the Add-CMApplicationDeployment Command
					Push-Location
					Set-Location $CMSite
					$DeploymentTypeCommand = "$DepTypeCommand$CmdSwitches -Force"
					Add-LogContent "Creating DeploymentType"
					Add-LogContent "Command: $DeploymentTypeCommand"
					Try {
						Invoke-Expression $DeploymentTypeCommand | Out-Null
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						$FullyQualified = $_.Exeption.FullyQualifiedErrorID
						Add-LogContent "ERROR: Adding MSI Deployment Type Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						Add-LogContent "ERROR: $FullyQualified"
						$DepTypeReturn = $false
					}
					If ($DepTypeAddDetectionMethods) {
						if ($DepTypeDetectionMethodType -eq "Custom") {
							Add-LogContent "Removing MSI Detection Method before adding new Detection Method"
							Push-Location
							Set-Location $CMSite
							Set-CMMsiDeploymentType -ApplicationName "$DepTypeApplicationName" -DeploymentTypeName "$DepTypeDeploymentTypeName" -ScriptText "Write-Output 0" -ScriptType PowerShell
							Pop-Location
						}
						Add-LogContent "Adding Detection Methods"
					
						Add-LogContent "Number of Detection Methods: $($DepTypeDetectionMethods.Count)"
						if ($DepTypeDetectionMethods.Count -eq 1) {
							Add-LogContent "Set-CMMsiDeploymentType -ApplicationName $DepTypeApplicationName -DeploymentTypeName $DepTypeDeploymentTypeName -AddDetectionClause $($DepTypeDetectionMethods[0].DataType.Name)"
							Try {
								Set-CMMsiDeploymentType -ApplicationName "$DepTypeApplicationName" -DeploymentTypeName "$DepTypeDeploymentTypeName" -AddDetectionClause $DepTypeDetectionMethods
							}
							Catch {
								$ErrorMessage = $_.Exception.Message
								$FullyQualified = $_.Exeption.FullyQualifiedErrorID
								Add-LogContent "ERROR: Adding Detection Method Failed!"
								Add-LogContent "ERROR: $ErrorMessage"
								Add-LogContent "ERROR: $FullyQualified"
								$DepTypeReturn = $false
							}
						}
						else {
							Add-LogContent "Set-CMMsiDeploymentType -ApplicationName $DepTypeApplicationName -DeploymentTypeName $DepTypeDeploymentTypeName -AddDetectionClause $($DepTypeDetectionMethods[0].DataType.Name) -"
							Try {
								Set-CMMsiDeploymentType -ApplicationName "$DepTypeApplicationName" -DeploymentTypeName "$DepTypeDeploymentTypeName" -AddDetectionClause $DepTypeDetectionMethods -DetectionClauseConnector $DepTypeDetectionClauseConnector
							}
							Catch {
								$ErrorMessage = $_.Exception.Message
								$FullyQualified = $_.Exeption.FullyQualifiedErrorID
								Add-LogContent "ERROR: Adding Detection Method Failed!"
								Add-LogContent "ERROR: $ErrorMessage"
								Add-LogContent "ERROR: $FullyQualified"
								$DepTypeReturn = $false
							}
						}
					}
					Pop-Location
				}			
				MSIX {
					# SOON(TM)
				}
				Default {
					$DepTypeReturn = $false
				}
			}

		
			## Add LEGACY Requirements for Deployment Type if they exist
			If (-not [System.String]::IsNullOrEmpty($DeploymentType.Requirements)) {
				Add-LogContent "Adding Requirements to $DepTypeDeploymentTypeName"
				$DepTypeRules = $DeploymentType.Requirements.RuleName
				ForEach ($DepTypeRule In $DepTypeRules) {
					Copy-CMDeploymentTypeRule -SourceApplicationName $RequirementsTemplateAppName -DestApplicationName $DepTypeApplicationName -DestDeploymentTypeName $DepTypeDeploymentTypeName -RuleName $DepTypeRule
				}
			}

			## Add NEW Requirements for Deployment Type is Necessary
			if (-not [System.String]::IsNullOrEmpty($DeploymentType.RequirementsRules)) {
				Add-LogContent "Adding Requirements to $DepTypeDeploymentTypeName"
				$DepTypeReqRules = $DeploymentType.RequirementsRules.RequirementsRule
				ForEach ($DepTypeReqRule In $DepTypeReqRules) {
					$addRequirementsRuleSplat = @{
						ReqRuleApplicationName   = $DepTypeApplicationName
						ReqRuleApplicationDTName = $DepTypeDeploymentTypeName
						ReqRuleValue             = @($DepTypeReqRule.RequirementsRuleValue.RuleValue)
						ReqRuleType              = $DepTypeReqRule.RequirementsRuleType
					}
					
					if (-not ([system.string]::IsNullOrEmpty($DepTypeReqRule.RequirementsRuleGlobalCondition))) {
						$addRequirementsRuleSplat.Add("ReqRuleGlobalConditionName", $DepTypeReqRule.RequirementsRuleGlobalCondition)
					}

					if (-not ([system.string]::IsNullOrEmpty($DepTypeReqRule.RequirementsRuleOperator))) {
						$addRequirementsRuleSplat.Add("ReqRuleOperator", $DepTypeReqRule.RequirementsRuleOperator)
					}

					if (-not ([system.string]::IsNullOrEmpty($DepTypeReqRule.RequirementsRuleValue2))) {
						$addRequirementsRuleSplat.Add("ReqRuleValue2", $DepTypeReqRule.ReqRuleValue2.RuleValue)
					}
					Write-Output "Add-RequirementsRule $addRequirementsRuleSplat"
					Add-RequirementsRule @addRequirementsRuleSplat
				}
			}
        
			## Add Install Behavior for Deployment Type if they exist
			If (-not [System.String]::IsNullOrEmpty($DeploymentType.InstallBehavior)) {
				Add-LogContent "Adding Install Behavior to $DepTypeDeploymentTypeName"
				$DepTypeInstallBehaviorProcesses = $DeploymentType.InstallBehavior.InstallBehaviorProcess
				ForEach ($DepTypeInstallBehavior In $DepTypeInstallBehaviorProcesses) {
					$newCMDeploymentTypeProcessRequirementSplat = @{
						ProcessDetectionDisplayName = $DepTypeInstallBehavior.DisplayName
						DestApplicationName         = $DepTypeApplicationName
						ProcessDetectionExecutable  = $DepTypeInstallBehavior.InstallBehaviorExe
						DestDeploymentTypeName      = $DepTypeDeploymentTypeName
					}
					Add-CMDeploymentTypeProcessDetection @newCMDeploymentTypeProcessRequirementSplat
				}
			}
		
			## Add Dependencies for Deployment Type if they exist
			if (-not [System.String]::IsNullOrEmpty($DeploymentType.Dependencies)) {
				Add-LogContent "Adding Dependencies to $DepTypeDeploymentTypeName"
				$DepTypeDependencyGroups = $DeploymentType.Dependencies.DependencyGroup
				foreach ($DepTypeDependencyGroup in $DepTypeDependencyGroups) {
					Add-LogContent "Creating Dependency Group $($DepTypeDependencyGroup.GroupName) on $DepTypeDeploymentTypeName"
					Push-Location
					Set-Location $CMSite
					$DependencyGroup = Get-CMDeploymentType -ApplicationName $DepTypeApplicationName -DeploymentTypeName $DepTypeDeploymentTypeName | New-CMDeploymentTypeDependencyGroup -GroupName $DepTypeDependencyGroup.GroupName
					$DepTypeDependencyGroupApps = $DepTypeDependencyGroup.DependencyGroupApp
					foreach ($DepTypeDependencyGroupApp in $DepTypeDependencyGroupApps) {
						$DependencyGroupAppAutoInstall = [System.Convert]::ToBoolean($DepTypeDependencyGroupApp.DependencyAutoInstall)
						$DependencyAppName = ((Get-CMApplication $DepTypeDependencyGroupApp.AppName | Sort-Object -Property Version -Descending | Select-Object -First 1).LocalizedDisplayName)
						if (-not [System.String]::IsNullOrEmpty($DepTypeDependencyGroupApp.DependencyDepType)) {
							Add-LogContent "Selecting Deployment Type for App Dependency: $($DepTypeDependencyGroupApp.DependencyDepType)"
							$DependencyAppObject = Get-CMDeploymentType -ApplicationName $DependencyAppName -DeploymentTypeName "$($DepTypeDependencyGroupApp.DependencyDepType)"
						}
						else {
							$DependencyAppObject = Get-CMDeploymentType -ApplicationName $DependencyAppName
						}
						$DependencyGroup | Add-CMDeploymentTypeDependency -DeploymentTypeDependency $DependencyAppObject -IsAutoInstall $DependencyGroupAppAutoInstall
					}
					Pop-Location
				}
			}
		}
		Return $DepTypeReturn
	}

	Function Invoke-ApplicationDistribution {
		Param (
			$Recipe
		)
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		ForEach ($Download In ($Recipe.ApplicationDef.Downloads.Download)) {
			If (-not ([System.String]::IsNullOrEmpty($Download.Version))) {
				$ApplicationSWVersion = $Download.Version
			}
		}
		$Success = $true
		## Distributes the Content for the Created Application based on the Information in the Recipe XML under the Distribution Node
		Push-Location
		Set-Location $CMSite
		$DistContent = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Distribution.DistributeContent)
		If ($DistContent) {
			If (-not ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Distribution.DistributeToGroup))) {
				$DistributionGroups = ($Recipe.ApplicationDef.Distribution.DistributeToGroup).Split(",")
				Add-LogContent "Distributing Content for $ApplicationName $ApplicationSWVersion to $($Recipe.ApplicationDef.Distribution.DistributeToGroup)"
				ForEach ($DistributionGroup In $DistributionGroups) {
					Try {
						Start-CMContentDistribution -ApplicationName "$ApplicationName $ApplicationSWVersion" -DistributionPointGroupName $DistributionGroup -ErrorAction Stop
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						Add-LogContent "ERROR: Content Distribution Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						$Success = $false
					}
				}
			}
			If (-not ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Distribution.DistributeToDPs))) {
				Add-LogContent "Distributing Content to $($Recipe.ApplicationDef.Distribution.DistributeToDPs)"
				$DistributionDPs = ($Recipe.ApplicationDef.Distribution.DistributeToDPs).Split(",")
				ForEach ($DistributionPoint In $DistributionDPs) {
					Try {
						Start-CMContentDistribution -ApplicationName "$ApplicationName $ApplicationSWVersion" -DistributionPointName $DistributionPoint -ErrorAction Stop
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						Add-LogContent "ERROR: Content Distribution Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						$Success = $false
					}
				}
			}
			If ((([string]::IsNullOrEmpty($Recipe.ApplicationDef.Distribution.DistributeToDPs)) -and ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Distribution.DistributeToGroup))) -and (-not ([String]::IsNullOrEmpty($Global:PreferredDistributionLoc)))) {
				$DistributionGroups = ($Global:PreferredDistributionLoc).Split(",")
				Add-LogContent "Distribution was set to True but No Distribution Points or Groups were Selected, Using Preferred Distribution Group(s): $Global:PreferredDistributionLoc"
				ForEach ($DistributionGroup In $DistributionGroups) {
					Try {
						Start-CMContentDistribution -ApplicationName "$ApplicationName $ApplicationSWVersion" -DistributionPointGroupName $DistributionGroup -ErrorAction Stop
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						Add-LogContent "ERROR: Content Distribution Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						$Success = $false
					}
				}
			}
		}
		Pop-Location
		Return $Success
	}

	Function Invoke-ApplicationDeployment {
		Param (
			$Recipe
		)
	
		$Success = $true
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		ForEach ($Download In ($Recipe.ApplicationDef.Downloads.Download)) {
			If (-not ([System.String]::IsNullOrEmpty($Download.Version))) {
				$ApplicationSWVersion = $Download.Version
			}
		}
	
		## Deploys the Created application based on the Information in the Recipe XML under the Deployment Node
		Push-Location
		Set-Location $CMSite
		foreach ($deployment in $Recipe.ApplicationDef.Deployment) 
		{
			If ([System.Convert]::ToBoolean($Deployment.DeploySoftware)) {
				$DeploymentSplat = @{
					Name = "$ApplicationName $ApplicationSWVersion"
					DeployAction = 'Install'
					UserNotification = 'DisplaySoftwareCenterOnly'
					UpdateSupersedence = [System.Convert]::ToBoolean($Deployment.UpdateSuperseded)
					AllowRepairApp = [System.Convert]::ToBoolean($Deployment.AllowRepair)
					ErrorAction = 'Stop'
				}

				if (-not ([string]::IsNullOrEmpty($Deployment.Purpose))) {
					$DeploymentSplat['DeployPurpose'] = $Deployment.Purpose
				}

				if (-not ([string]::IsNullOrEmpty($Deployment.AvailableOffset))) {
					$DeploymentSplat['AvailableDateTime'] = (Get-Date) + $Deployment.AvailableOffset
				}

				if (-not ([string]::IsNullOrEmpty($Deployment.DeadlineOffset))) {
					$DeploymentSplat['DeadlineDateTime'] = (Get-Date) + $Deployment.DeadlineOffset
				}

				if (-not ([string]::IsNullOrEmpty($Deployment.TimeBaseOn))) {
					# Only 'LocalTime' or 'UTC' are accepted values, but let CM error.
					$DeploymentSplat['TimeBaseOn'] = $Deployment.TimeBaseOn
				}

				$DeploymentCollections = If (
					-not ([string]::IsNullOrEmpty($Deployment.DeploymentCollection))
					) {
					$Deployment.DeploymentCollection
				} elseIf (-not ([String]::IsNullOrEmpty($Global:PreferredDeployCollection))) {
					$Global:PreferredDeployCollection
				}

				Foreach ($DeploymentCollection in $DeploymentCollections) {
					## Check if the deployment collection exists in the ConfigMgr site, create it if missing
					$ExistingCollection = Get-CMDeviceCollection -Name $DeploymentCollection -ErrorAction SilentlyContinue
					If (-not $ExistingCollection) {
						$WarningMessage = "Collection '$DeploymentCollection' does not exist in site '$($Global:SiteCode)'. Creating collection."
						Write-Warning $WarningMessage
						Add-LogContent "WARNING: $WarningMessage"
						Try {
							New-CMDeviceCollection -Name $DeploymentCollection -LimitingCollectionName "All Systems" -ErrorAction Stop | Out-Null
							Add-LogContent "Successfully created collection '$DeploymentCollection' with limiting collection 'All Systems'"
						}
						Catch {
							$ErrorMessage = $_.Exception.Message
							Add-LogContent "ERROR: Failed to create collection '$DeploymentCollection'"
							Add-LogContent "ERROR: $ErrorMessage"
							$Success = $false
							Continue
						}
					}

					Try {
						Add-LogContent "Deploying $ApplicationName $ApplicationSWVersion to $DeploymentCollection"
						If ($DeploymentSplat.UpdateSupersedence) { Add-LogContent "UpdateSuperseded enabled, new package will automatically upgrade previous version" }
						New-CMApplicationDeployment -CollectionName $DeploymentCollection @DeploymentSplat
					}
					Catch {
						$ErrorMessage = $_.Exception.Message
						Add-LogContent "ERROR: Deployment Failed!"
						Add-LogContent "ERROR: $ErrorMessage"
						$Success = $false
					}
				}
			}
		}
		Pop-Location
		Return $Success
	}

	function Invoke-ApplicationSupersedence {
		param (
			$Recipe
		)

		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		$ApplicationPublisher = $Recipe.ApplicationDef.Application.Publisher
		If (-not ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Supersedence.Supersedence))) {
			$SupersedenceEnabled = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Supersedence.Supersedence)
		}
		else {
			$SupersedenceEnabled = $false
		}

		If (-not ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Supersedence.Uninstall))) {
			$UninstallOldApp = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Supersedence.Uninstall)
		}
		else {
			$UninstallOldApp = $false
		}

		Write-Host "Supersedence is $SupersedenceEnabled"
		if ($SupersedenceEnabled) {
			# Get the Previous Application Deployment Type
			Push-Location
			Set-Location $CMSite
			$Latest2Apps = Get-CMApplication -Name "$ApplicationName*" -Fast | Where-Object { ($_.Manufacturer -eq $ApplicationPublisher) -and ($_.IsExpired -eq $false) -and ($_.IsSuperseded -eq $false) } | Sort-Object DateCreated -Descending | Select-Object -first 2
			Write-Host "Latest 2 apps = $($Latest2Apps.LocalizedDisplayName)"
			if ($Latest2Apps.Count -eq 2) {
				$NewApp = $Latest2Apps | Select-Object -First 1
				$OldApp = $Latest2Apps | Select-Object -last 1
				Write-Host "Old: $($oldapp.LocalizedDisplayName) New: $($newapp.LocalizedDisplayName)"

				# Check that the DeploymentTypes and Deployment Type Names Match if not, skip supersedence
				$NewAppDeploymentTypes = Get-CMDeploymentType -InputObject $NewApp | Sort-Object LocalizedDisplayName
				$OldAppDeploymentTypes = Get-CMDeploymentType -InputObject $OldApp | Sort-Object LocalizedDisplayName

				Foreach ($DeploymentType in $NewAppDeploymentTypes) {
					Write-Host "Superseding $($DeploymentType.LocalizedDisplayName)"
					$SupersededDeploymentType = $OldAppDeploymentTypes | Where-Object LocalizedDisplayName -eq $DeploymentType.LocalizedDisplayName
					Set-CMApplicationSupersedence -InputObject $NewApp -CurrentDeploymentType $DeploymentType -SupersededApplication $OldApp -OldDeploymentType $SupersededDeploymentType -IsUninstall $UninstallOldApp | Out-Null
				}
			}
			Pop-Location
		}
		Write-Output $true
	}

	function Invoke-ApplicationCleanup {
		param (
			$Recipe
		)
		If (-not ([string]::IsNullOrEmpty($Recipe.ApplicationDef.Supersedence.CleanupSuperseded))) {
			$CleanupEnabled = [System.Convert]::ToBoolean($Recipe.ApplicationDef.Supersedence.CleanupSuperseded)
		}
		else {
			$CleanupEnabled = $false
		}
		$ApplicationName = $Recipe.ApplicationDef.Application.Name
		$CleanupEnabled = $Recipe.ApplicationDef.Supersedence.CleanupSuperseded
		$keep = $Recipe.ApplicationDef.Supersedence.KeepSuperseded

		Write-Output "Cleanup is $CleanupEnabled"
		if ($CleanupEnabled) {
			
			Push-Location
			Set-Location $CMSite
			Write-Host "Keeping $Keep superseded revisions of $ApplicationName"
			$Applications = Get-CMApplication -Name "$ApplicationName*" | Where-Object IsSuperseded -eq $true | Sort-Object DateCreated
			If ($Applications.Count -le $keep) {
				Write-Host "Number of superseded applications ($($Applications.Count)) is less than or equal to the number to keep ($keep), no applications will be removed"
			}
			else {
				Write-Host "Number of superseded applications ($($Applications.Count)) is more than or equal to the number to keep ($keep), oldest applications will be removed"
				$Applications = $Applications | Select-Object -First ($Applications.Count - $keep)
				ForEach ($Application in $Applications) {
					# Get the content location and remove it
					Write-Host "Cleaning up $($Application.LocalizedDisplayName)"
					Pop-Location
					$ApplicationXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($Application.SDMPackageXML, $true)
					$Location = $ApplicationXML.DeploymentTypes[0].Installer.Contents | Select-Object -ExpandProperty Location # BUGBUG: Get all the deployment locations and remove them
					Remove-Item -LiteralPath $Location -Recurse
					Add-LogContent "Removed application content from $Location`n"
					# Remove the deployments and app itself
					Push-Location
					Set-Location $CMSite
					$Application | Get-CMApplicationDeployment | Remove-CMApplicationDeployment -Force
					Get-CMApplication $Application.LocalizedDisplayName | Remove-CMApplication -Force
					## Send an Email if an Application was successfully cleaned up and record the Application Name and Version for the Email
					$Global:SendEmail = $true; $Global:SendEmail | Out-Null
					$Global:EmailBody += "      - Removed $($Application.LocalizedDisplayName) `n"
					Add-LogContent "Removed $($Application.LocalizedDisplayName) $($Application.SoftwareVersion)`n"
				}
			}
			Pop-Location
		}
		Write-Output $true
	}	

	Function Send-EmailMessage {
		Add-LogContent "Sending Email"
		$Global:EmailBody += "`n`nThis message was automatically generated"
		Try {
			Send-MailMessage -To $EmailTo -Subject $EmailSubject -From $EmailFrom -Body $Global:EmailBody -SmtpServer $EmailServer -ErrorAction Stop
		}
		Catch {
			$ErrorMessage = $_.Exception.Message
			Add-LogContent "ERROR: Sending Email Failed!"
			Add-LogContent "ERROR: $ErrorMessage"
		}
	}

	Function Connect-ConfigMgr {
		$Global:ConfigMgrConnection = $true
		$Global:ConfigMgrConnection | Out-Null
		if (-not (Get-Module ConfigurationManager)) {
			try {
				Add-LogContent "Importing ConfigurationManager Module"
				if ($Global:CMPSModulePath) {
					Import-Module (Join-Path $Global:CMPSModulePath ConfigurationManager.psd1) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				} else {
					Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				}
			} 
			catch {
				$ErrorMessage = $_.Exception.Message
				Add-LogContent "ERROR: Importing ConfigurationManager Module Failed!"
				Add-LogContent "ERROR: $ErrorMessage"
				if (-not $Setup) {
					Exit 1
				}
				else {
					$Global:ConfigMgrConnection = $false
				}
			}
		}
	
		if ($null -eq (Get-PSDrive -Name $Global:SiteCode -ErrorAction SilentlyContinue)) {
			try {
				New-PSDrive -Name $Global:SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $Global:SiteServer -Scope Script
			}
			catch {
				Add-LogContent "ERROR - The CM PSDrive could not be loaded. Exiting..."
				Add-LogContent "ERROR: $ErrorMessage"
				if (-not $Setup) {
					Exit 1
				}
				else {
					$Global:ConfigMgrConnection = $false
				}
			}
		}
	}

	function Test-SetupInputs {
		param([hashtable]$Settings)
		$e = New-Object System.Collections.ArrayList
		if ([string]::IsNullOrWhiteSpace($Settings.TempDir))             { [void]$e.Add('Working Directory is required.') }
		if ([string]::IsNullOrWhiteSpace($Settings.ContentLocationRoot)) { [void]$e.Add('Content Root is required.') }
		if ($Settings.CMSite -notmatch '^[A-Z0-9]{3}:?$')               { [void]$e.Add("Site Code must be 3 uppercase alphanumeric characters, optionally followed by ':'.") }
		if ([string]::IsNullOrWhiteSpace($Settings.SiteServer))          { [void]$e.Add('Site Server is required.') }
		if ($Settings.SendEmailPreference -eq 'True') {
			if ([string]::IsNullOrWhiteSpace($Settings.EmailTo)   -or $Settings.EmailTo   -notmatch '@') { [void]$e.Add('Email To must be a valid email address.') }
			if ([string]::IsNullOrWhiteSpace($Settings.EmailFrom) -or $Settings.EmailFrom -notmatch '@') { [void]$e.Add('Email From must be a valid email address.') }
			if ([string]::IsNullOrWhiteSpace($Settings.EmailServer))                                     { [void]$e.Add('SMTP Server is required when email reports are enabled.') }
		}
		if ($Settings.ContainsKey('WebServerPort') -and -not [string]::IsNullOrWhiteSpace($Settings.WebServerPort)) {
			$portInt = 0
			if (-not ([int]::TryParse($Settings.WebServerPort, [ref]$portInt)) -or $portInt -lt 1 -or $portInt -gt 65535) {
				[void]$e.Add('Web Server Port must be a number between 1 and 65535.')
			}
		}
		return ,$e
	}

	function Save-SetupPrefs {
		param([hashtable]$Settings, [string]$PreferenceFile, [string]$TemplatePath)
		if (Test-Path $PreferenceFile -ErrorAction SilentlyContinue) {
			[xml]$xml = Get-Content $PreferenceFile
		} else {
			[xml]$xml = Get-Content $TemplatePath
		}
		foreach ($key in $Settings.Keys) {
			$xml.PackagerPrefs.$key = [string]$Settings[$key]
		}
		$xml.PackagerPrefs.LogPath = "$(Split-Path $Settings.TempDir -Parent)\CMPackager.log"
		$parentDir = Split-Path $PreferenceFile -Parent
		if ($parentDir -and -not (Test-Path $parentDir -ErrorAction SilentlyContinue)) {
			New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
		}
		$xml.Save($PreferenceFile)
	}

	function Invoke-InteractiveSetup {
		param([hashtable]$Defaults)

		function prompt-field {
			param([string]$label, [string]$default, [scriptblock]$validate, [string]$errMsg, [switch]$optional)
			do {
				$hint   = if ($default) { " [$default]" } else { '' }
				$answer = Read-Host "$label$hint"
				if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $default }
				if ($optional -and [string]::IsNullOrWhiteSpace($answer)) { return $answer }
				if ($validate -and -not (& $validate $answer)) {
					Write-Host "  $errMsg" -ForegroundColor Red
					$answer = $null
				}
			} while ([string]::IsNullOrWhiteSpace($answer))
			return $answer
		}

		function prompt-bool {
			param([string]$label, [bool]$default)
			$hint   = if ($default) { '[Y/n]' } else { '[y/N]' }
			$answer = Read-Host "$label $hint"
			if ([string]::IsNullOrWhiteSpace($answer)) { return $default }
			return $answer -match '^[yY]'
		}

		$s = @{}
		Write-Host ''
		Write-Host '=== CMPackager Setup Wizard ===' -ForegroundColor Cyan
		Write-Host ''
		Write-Host '-- Required Settings --' -ForegroundColor Yellow
		$s.TempDir             = prompt-field 'Working Directory'             $Defaults.TempDir             { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
		$s.ContentLocationRoot = prompt-field 'Content Root'                  $Defaults.ContentLocationRoot { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
		$s.CMSite              = (prompt-field 'Site Code (e.g. PS1 or PS1:)' $Defaults.CMSite              { param($v) $v.ToUpper() -match '^[A-Z0-9]{3}:?$' } "Must be 3 uppercase alphanumeric characters, optionally followed by ':'.").ToUpper()
		$s.SiteServer          = prompt-field 'Site Server FQDN'              $Defaults.SiteServer          { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
		Write-Host ''
		Write-Host '-- Optional Settings --' -ForegroundColor Yellow
		$s.IconRepo            = prompt-field 'Icon Repository (leave blank to skip)' $Defaults.IconRepo -optional
		$s.NoVersionInSWCenter = (prompt-bool 'Hide version in Software Center display names?' ($Defaults.NoVersionInSWCenter -eq 'True')).ToString()
		Write-Host ''
		Write-Host '-- Email Reporting --' -ForegroundColor Yellow
		$sendEmail             = prompt-bool 'Enable email reports?' ($Defaults.SendEmailPreference -eq 'True')
		$s.SendEmailPreference = $sendEmail.ToString()
		if ($sendEmail) {
			$s.EmailTo                 = prompt-field 'Email To'    $Defaults.EmailTo     { param($v) $v -match '@' } 'Must contain @.'
			$s.EmailFrom               = prompt-field 'Email From'  $Defaults.EmailFrom   { param($v) $v -match '@' } 'Must contain @.'
			$s.EmailServer             = prompt-field 'SMTP Server' $Defaults.EmailServer { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
			$s.NotifyOnDownloadFailure = (prompt-bool 'Notify on download failure?' ($Defaults.NotifyOnDownloadFailure -eq 'True')).ToString()
		} else {
			$s.EmailTo                 = $Defaults.EmailTo
			$s.EmailFrom               = $Defaults.EmailFrom
			$s.EmailServer             = $Defaults.EmailServer
			$s.NotifyOnDownloadFailure = $Defaults.NotifyOnDownloadFailure
		}
		Write-Host ''
		Write-Host '-- SCCM Defaults (optional) --' -ForegroundColor Yellow
		$s.PreferredDistributionLoc  = prompt-field 'Preferred Distribution Point Group' $Defaults.PreferredDistributionLoc  -optional
		$s.PreferredDeployCollection = prompt-field 'Preferred Deployment Collection'    $Defaults.PreferredDeployCollection -optional
		$s.ContentFolderPattern      = $Defaults.ContentFolderPattern
		$s.CMPSModulePath            = $Defaults.CMPSModulePath
		$s.GitHubToken               = $Defaults.GitHubToken
		Write-Host ''
		Write-Host '-- Web Server --' -ForegroundColor Yellow
		$s.WebServerPort         = prompt-field 'Web Server Port' $Defaults.WebServerPort { param($v) $v -match '^\d+$' -and [int]$v -ge 1 -and [int]$v -le 65535 } 'Must be a number between 1 and 65535.'
		$s.WebServerRequiredRole = prompt-field 'Required SCCM Role (leave blank for any admin)' $Defaults.WebServerRequiredRole -optional

		Write-Host ''
		Write-Host '-- Review --' -ForegroundColor Yellow
		$reviewItems = [ordered]@{
			'Working Directory'         = $s.TempDir
			'Content Root'              = $s.ContentLocationRoot
			'Site Code'                 = $s.CMSite
			'Site Server'               = $s.SiteServer
			'Icon Repository'           = if ($s.IconRepo) { $s.IconRepo } else { '(empty)' }
			'Hide Version in SW Center' = $s.NoVersionInSWCenter
			'Email Reports'             = $s.SendEmailPreference
			'Email To'                  = if ($s.EmailTo) { $s.EmailTo } else { '(empty)' }
			'Email From'                = if ($s.EmailFrom) { $s.EmailFrom } else { '(empty)' }
			'SMTP Server'               = if ($s.EmailServer) { $s.EmailServer } else { '(empty)' }
			'Notify on Failure'         = $s.NotifyOnDownloadFailure
			'Distribution Point Group'  = if ($s.PreferredDistributionLoc) { $s.PreferredDistributionLoc } else { '(empty)' }
			'Deployment Collection'     = if ($s.PreferredDeployCollection) { $s.PreferredDeployCollection } else { '(empty)' }
			'Web Server Port'           = $s.WebServerPort
			'Required SCCM Role'        = if ($s.WebServerRequiredRole) { $s.WebServerRequiredRole } else { '(any admin)' }
		}
		foreach ($pair in $reviewItems.GetEnumerator()) {
			Write-Host ("  {0,-30} {1}" -f "$($pair.Key):", $pair.Value)
		}
		Write-Host ''
		$confirm = Read-Host 'Save these settings? [Y/n]'
		if ($confirm -match '^[nN]') { return $null }
		return $s
	}

	function Invoke-SpectreSetup {
		param([hashtable]$Defaults)

		function read-required {
			param([string]$prompt, [string]$default, [scriptblock]$validate, [string]$errMsg)
			do {
				$val = Read-SpectreText -Message $prompt -DefaultAnswer $default
				if ([string]::IsNullOrWhiteSpace($val)) { $val = $default }
				if ($validate -and -not (& $validate $val)) {
					Write-Host "  $errMsg" -ForegroundColor Red
					$val = $null
				}
			} while ([string]::IsNullOrWhiteSpace($val))
			return $val
		}

		$s = @{}

		Write-SpectreRule -Title 'CMPackager Setup Wizard' -Color 'Blue'
		Write-Host ''

		Write-SpectreRule -Title 'Required Settings' -Color 'Grey'
		$s.TempDir             = read-required 'Working Directory'              $Defaults.TempDir             { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
		$s.ContentLocationRoot = read-required 'Content Root'                   $Defaults.ContentLocationRoot { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
		$s.CMSite              = (read-required 'Site Code (e.g. PS1 or PS1:)' $Defaults.CMSite              { param($v) $v.ToUpper() -match '^[A-Z0-9]{3}:?$' } "Must be 3 uppercase alphanumeric characters, optionally followed by ':'.").ToUpper()
		$s.SiteServer          = read-required 'Site Server FQDN'               $Defaults.SiteServer          { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'

		Write-Host ''
		Write-SpectreRule -Title 'Optional Settings' -Color 'Grey'
		$s.IconRepo            = Read-SpectreText    -Message 'Icon Repository (leave blank to skip)' -DefaultAnswer $Defaults.IconRepo -AllowEmpty
		$noVersion             = Read-SpectreConfirm -Message 'Hide version in Software Center display names?' -DefaultAnswer (if ($Defaults.NoVersionInSWCenter -eq 'True') { 'y' } else { 'n' })
		$s.NoVersionInSWCenter = $noVersion.ToString()

		Write-Host ''
		Write-SpectreRule -Title 'Email Reporting' -Color 'Grey'
		$sendEmail             = Read-SpectreConfirm -Message 'Enable email reports?' -DefaultAnswer (if ($Defaults.SendEmailPreference -eq 'True') { 'y' } else { 'n' })
		$s.SendEmailPreference = $sendEmail.ToString()
		if ($sendEmail) {
			$s.EmailTo                 = read-required 'Email To'    $Defaults.EmailTo     { param($v) $v -match '@' } 'Must contain @.'
			$s.EmailFrom               = read-required 'Email From'  $Defaults.EmailFrom   { param($v) $v -match '@' } 'Must contain @.'
			$s.EmailServer             = read-required 'SMTP Server' $Defaults.EmailServer { param($v) -not [string]::IsNullOrWhiteSpace($v) } 'Value is required.'
			$notifyFail                = Read-SpectreConfirm -Message 'Notify on download failure?' -DefaultAnswer (if ($Defaults.NotifyOnDownloadFailure -eq 'True') { 'y' } else { 'n' })
			$s.NotifyOnDownloadFailure = $notifyFail.ToString()
		} else {
			$s.EmailTo                 = $Defaults.EmailTo
			$s.EmailFrom               = $Defaults.EmailFrom
			$s.EmailServer             = $Defaults.EmailServer
			$s.NotifyOnDownloadFailure = $Defaults.NotifyOnDownloadFailure
		}

		Write-Host ''
		Write-SpectreRule -Title 'SCCM Defaults (optional)' -Color 'Grey'
		$s.PreferredDistributionLoc  = Read-SpectreText -Message 'Preferred Distribution Point Group' -DefaultAnswer $Defaults.PreferredDistributionLoc  -AllowEmpty
		$s.PreferredDeployCollection = Read-SpectreText -Message 'Preferred Deployment Collection'    -DefaultAnswer $Defaults.PreferredDeployCollection -AllowEmpty
		$s.ContentFolderPattern      = $Defaults.ContentFolderPattern
		$s.CMPSModulePath            = $Defaults.CMPSModulePath
		$s.GitHubToken               = $Defaults.GitHubToken

		Write-Host ''
		Write-SpectreRule -Title 'Web Server' -Color 'Grey'
		$s.WebServerPort         = read-required 'Web Server Port' $Defaults.WebServerPort { param($v) $v -match '^\d+$' -and [int]$v -ge 1 -and [int]$v -le 65535 } 'Must be a number between 1 and 65535.'
		$s.WebServerRequiredRole = Read-SpectreText -Message 'Required SCCM Role (leave blank for any admin)' -DefaultAnswer $Defaults.WebServerRequiredRole -AllowEmpty

		Write-Host ''
		Write-SpectreRule -Title 'Review' -Color 'Grey'
		$tableData = @(
			[pscustomobject]@{ Setting = 'Working Directory';         Value = $s.TempDir }
			[pscustomobject]@{ Setting = 'Content Root';              Value = $s.ContentLocationRoot }
			[pscustomobject]@{ Setting = 'Site Code';                 Value = $s.CMSite }
			[pscustomobject]@{ Setting = 'Site Server';               Value = $s.SiteServer }
			[pscustomobject]@{ Setting = 'Icon Repository';           Value = if ($s.IconRepo) { $s.IconRepo } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'Hide Version in SW Center'; Value = $s.NoVersionInSWCenter }
			[pscustomobject]@{ Setting = 'Email Reports';             Value = $s.SendEmailPreference }
			[pscustomobject]@{ Setting = 'Email To';                  Value = if ($s.EmailTo) { $s.EmailTo } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'Email From';                Value = if ($s.EmailFrom) { $s.EmailFrom } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'SMTP Server';               Value = if ($s.EmailServer) { $s.EmailServer } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'Notify on Failure';         Value = $s.NotifyOnDownloadFailure }
			[pscustomobject]@{ Setting = 'Distribution Point Group';  Value = if ($s.PreferredDistributionLoc) { $s.PreferredDistributionLoc } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'Deployment Collection';     Value = if ($s.PreferredDeployCollection) { $s.PreferredDeployCollection } else { '(empty)' } }
			[pscustomobject]@{ Setting = 'Web Server Port';           Value = $s.WebServerPort }
			[pscustomobject]@{ Setting = 'Required SCCM Role';        Value = if ($s.WebServerRequiredRole) { $s.WebServerRequiredRole } else { '(any admin)' } }
		)
		Format-SpectreTable -Data $tableData -Color 'Grey'

		$save = Read-SpectreConfirm -Message 'Save these settings?' -DefaultAnswer 'y'
		if (-not $save) { return $null }
		return $s
	}

	function Invoke-Setup {
		param([string]$PreferenceFile)

		if (Test-Path $PreferenceFile -ErrorAction SilentlyContinue) {
			[xml]$xml = Get-Content $PreferenceFile
		} else {
			[xml]$xml = Get-Content "$PSScriptRoot\CMPackager.prefs.template"
		}
		$defaults = @{
			TempDir                   = $xml.PackagerPrefs.TempDir
			ContentLocationRoot       = $xml.PackagerPrefs.ContentLocationRoot
			IconRepo                  = $xml.PackagerPrefs.IconRepo
			CMSite                    = $xml.PackagerPrefs.CMSite
			SiteServer                = $xml.PackagerPrefs.SiteServer
			NoVersionInSWCenter       = $xml.PackagerPrefs.NoVersionInSWCenter
			EmailTo                   = $xml.PackagerPrefs.EmailTo
			EmailFrom                 = $xml.PackagerPrefs.EmailFrom
			EmailServer               = $xml.PackagerPrefs.EmailServer
			SendEmailPreference       = $xml.PackagerPrefs.SendEmailPreference
			NotifyOnDownloadFailure   = $xml.PackagerPrefs.NotifyOnDownloadFailure
			PreferredDistributionLoc  = $xml.PackagerPrefs.PreferredDistributionLoc
			PreferredDeployCollection = $xml.PackagerPrefs.PreferredDeployCollection
			ContentFolderPattern      = $xml.PackagerPrefs.ContentFolderPattern
			CMPSModulePath            = $xml.PackagerPrefs.CMPSModulePath
			GitHubToken               = $xml.PackagerPrefs.GitHubToken
			WebServerPort             = $xml.PackagerPrefs.WebServerPort
			WebServerRequiredRole     = $xml.PackagerPrefs.WebServerRequiredRole
		}

		$useSpectre = $false
		if (Get-Module -ListAvailable -Name PwshSpectreConsole -ErrorAction SilentlyContinue) {
			try {
				Import-Module PwshSpectreConsole -ErrorAction Stop
				$useSpectre = $true
			} catch {
				Write-Host "Failed to import PwshSpectreConsole ($_). Falling back to interactive prompts." -ForegroundColor Yellow
			}
		} else {
			Write-Host ''
			Write-Host 'PwshSpectreConsole is not installed. It provides a nicer setup experience.'
			$answer = Read-Host 'Install it now? (requires internet access) [y/N]'
			if ($answer -match '^[yY]') {
				try {
					Install-Module PwshSpectreConsole -Scope CurrentUser -Force -ErrorAction Stop
					Import-Module PwshSpectreConsole -ErrorAction Stop
					$useSpectre = $true
				} catch {
					Write-Host "Installation failed ($_). Falling back to interactive prompts." -ForegroundColor Yellow
				}
			}
		}

		$settings = if ($useSpectre) {
			Invoke-SpectreSetup -Defaults $defaults
		} else {
			Invoke-InteractiveSetup -Defaults $defaults
		}

		if ($null -eq $settings) {
			Write-Host 'Setup cancelled. No changes were saved.' -ForegroundColor Yellow
			return
		}

		$errors = Test-SetupInputs -Settings $settings
		if ($errors.Count -gt 0) {
			Write-Host ''
			Write-Host 'Cannot save - validation errors:' -ForegroundColor Red
			foreach ($e in $errors) { Write-Host "  - $e" -ForegroundColor Red }
			return
		}

		Save-SetupPrefs -Settings $settings -PreferenceFile $PreferenceFile -TemplatePath "$PSScriptRoot\CMPackager.prefs.template"
		Write-Host ''
		Write-Host "Configuration saved to $PreferenceFile" -ForegroundColor Green
	}

	function Get-WebServerPrefixes {
		param([int]$Port)
		$prefixes = New-Object System.Collections.Generic.List[string]
		[void]$prefixes.Add("http://localhost:$Port/")
		try {
			$netProfiles = Get-NetConnectionProfile -ErrorAction SilentlyContinue |
				Where-Object { $_.NetworkCategory -in @('Private', 'DomainAuthenticated') }
			foreach ($netProfile in $netProfiles) {
				$addrs = Get-NetIPAddress -InterfaceIndex $netProfile.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
					Where-Object { $_.IPAddress -notmatch '^169\.254\.' }
				foreach ($addr in $addrs) {
					$prefix = "http://$($addr.IPAddress):$Port/"
					if (-not $prefixes.Contains($prefix)) { [void]$prefixes.Add($prefix) }
				}
			}
		} catch {
			Add-LogContent "WebServer: could not enumerate network interfaces - $($_.Exception.Message)"
		}
		return [string[]]$prefixes
	}

	function Register-WebServerUrlAcl {
		param([int]$Port)
		$url     = "http://+:$Port/"
		$pattern = [regex]::Escape($url)
		$existing = (& netsh http show urlacl url=$url 2>&1) | Out-String
		if ($existing -notmatch $pattern) {
			$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
			Add-LogContent "WebServer: registering URL ACL for $url (user: $currentUser)"
			$cmdArgs = @('http', 'add', 'urlacl', "url=$url", "user=$currentUser")
			$result  = (& netsh @cmdArgs 2>&1) | Out-String
			Add-LogContent "WebServer: URL ACL result - $($result.Trim())"
			if ($result -match 'Error' -or $result -match 'Access is denied') {
				Write-Host "WARNING: URL ACL registration failed ($($result.Trim())). Run as administrator if the server fails to start." -ForegroundColor Yellow
			}
		} else {
			Add-LogContent "WebServer: URL ACL already registered for $url"
		}
	}

	function Test-SCCMAdminAccess {
		param([string]$LogonName)
		if ([string]::IsNullOrWhiteSpace($Global:SiteCode) -or [string]::IsNullOrWhiteSpace($Global:SiteServer)) {
			Add-LogContent "WebServer: SCCM site not configured - cannot verify admin access for '$LogonName'"
			return $null
		}
		try {
			$safeName = $LogonName.Replace('\', '\\').Replace("'", "''")
			$admins = Get-WmiObject -Namespace "root\SMS\site_$Global:SiteCode" `
				-ComputerName $Global:SiteServer `
				-Query "SELECT LogonName FROM SMS_Admin WHERE LogonName = '$safeName'" `
				-ErrorAction Stop
			return ($null -ne $admins)
		} catch {
			Add-LogContent "WebServer: SCCM admin check failed for '$LogonName' - $($_.Exception.Message)"
			return $null
		}
	}

	function Get-RecipeSummaries {
		param([string]$RecipesFolder)
		$summaries = New-Object System.Collections.Generic.List[psobject]
		$xmlFiles = Get-ChildItem -Path $RecipesFolder -Filter '*.xml' -ErrorAction SilentlyContinue |
			Where-Object { $_.Name -ne 'Template.xml' } |
			Sort-Object Name
		foreach ($f in $xmlFiles) {
			try {
				[xml]$r = Get-Content $f.FullName
				$app = $r.ApplicationDef.Application
				$dts = @($r.ApplicationDef.DeploymentTypes.DeploymentType)
				$hasPrefetch = $false
				foreach ($dl in @($r.ApplicationDef.Downloads.Download)) {
					if (-not [string]::IsNullOrWhiteSpace($dl.PrefetchScript)) { $hasPrefetch = $true }
				}
				[void]$summaries.Add([pscustomobject]@{
					Name        = [string]$app.Name
					Publisher   = [string]$app.Publisher
					Description = [string]$app.Description
					FileName    = $f.Name
					DtCount     = $dts.Count
					DtNames     = ($dts | ForEach-Object { $_.Name }) -join ', '
					HasPrefetch = $hasPrefetch
					Distribute  = [string]$r.ApplicationDef.Distribution.DistributeContent
					Deploy      = [string]$r.ApplicationDef.Deployment.DeploySoftware
				})
			} catch {
				[void]$summaries.Add([pscustomobject]@{
					Name        = $f.BaseName
					Publisher   = ''
					Description = "(Parse error: $($_.Exception.Message))"
					FileName    = $f.Name
					DtCount     = 0
					DtNames     = ''
					HasPrefetch = $false
					Distribute  = ''
					Deploy      = ''
				})
			}
		}
		return $summaries
	}

	function New-RecipesPageHtml {
		param([string]$RecipesFolder)
		$enc       = [System.Net.WebUtility]
		$recipes   = Get-RecipeSummaries -RecipesFolder $RecipesFolder
		$rows = foreach ($r in $recipes) {
			$urlBadge  = if ($r.HasPrefetch) { '<span class="badge bdyn">Dynamic</span>' } else { '<span class="badge bsta">Static</span>' }
			$distBadge = if ($r.Distribute -eq 'True') { '<span class="badge byes">Yes</span>' } else { '<span class="badge bno">No</span>' }
			$depBadge  = if ($r.Deploy -eq 'True') { '<span class="badge byes">Yes</span>' } else { '<span class="badge bno">No</span>' }
			$dtHtml    = if ($r.DtCount -gt 0) { "<span class='cnt'>$($r.DtCount)</span> $($enc::HtmlEncode($r.DtNames))" } else { '<em>none</em>' }
			"<tr><td><strong>$($enc::HtmlEncode($r.Name))</strong><br><small>$($enc::HtmlEncode($r.FileName))</small></td><td>$($enc::HtmlEncode($r.Publisher))</td><td>$($enc::HtmlEncode($r.Description))</td><td>$dtHtml</td><td>$urlBadge</td><td>$distBadge</td><td>$depBadge</td></tr>"
		}
		$rowsHtml  = $rows -join "`n"
		$count     = $recipes.Count
		$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
		return @"
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>CMPackager - Recipes</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;font-size:14px;background:#f3f2f1;color:#323130;min-height:100vh}
.bar{background:#0078d4;color:#fff;padding:12px 24px;display:flex;align-items:center;gap:12px}
.bar h1{font-size:18px;font-weight:600}.bar .sub{font-size:13px;opacity:.8}
.wrap{padding:24px;max-width:1400px;margin:0 auto}
.meta{margin-bottom:16px;font-size:13px;color:#605e5c}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:4px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
thead tr{background:#f8f8f8}
th{padding:10px 14px;text-align:left;font-weight:600;font-size:12px;color:#605e5c;text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid #edebe9}
td{padding:10px 14px;border-bottom:1px solid #edebe9;vertical-align:top}
tr:last-child td{border-bottom:none}tr:hover td{background:#f3f2f1}
small{color:#605e5c;font-family:'Cascadia Code',Consolas,monospace;font-size:11px}
.cnt{display:inline-block;background:#e1dfdd;color:#323130;border-radius:10px;padding:1px 7px;font-size:11px;font-weight:600;margin-right:4px}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}
.bdyn{background:#deecf9;color:#004e8c}.bsta{background:#e8e8e8;color:#444}
.byes{background:#dff6dd;color:#107c10}.bno{background:#fde7e9;color:#a4262c}
</style></head>
<body>
<div class="bar"><h1>CMPackager</h1><span class="sub">Recipe Browser</span></div>
<div class="wrap">
<p class="meta">$count recipe(s) | loaded $timestamp</p>
<table>
<thead><tr><th>Application</th><th>Publisher</th><th>Description</th><th>Deployment Types</th><th>URL</th><th>Distribute</th><th>Deploy</th></tr></thead>
<tbody>
$rowsHtml
</tbody></table></div></body></html>
"@
	}

	function New-AccessDeniedPageHtml {
		param([string]$LogonName, [string]$RequiredRole)
		$enc      = [System.Net.WebUtility]
		$roleNote = if ($RequiredRole) {
			"<p>Access requires the SCCM role: <strong>$($enc::HtmlEncode($RequiredRole))</strong></p>"
		} else {
			'<p>Access requires SCCM Administrative User permissions.</p>'
		}
		return @"
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>CMPackager - Access Denied</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f3f2f1;color:#323130;display:flex;align-items:center;justify-content:center;min-height:100vh}
.card{background:#fff;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.12);padding:40px 48px;max-width:480px;text-align:center}
h1{font-size:22px;color:#a4262c;margin-bottom:12px}
p{color:#605e5c;margin-bottom:8px;font-size:14px}
.acct{font-family:'Cascadia Code',Consolas,monospace;background:#f3f2f1;padding:4px 10px;border-radius:4px;display:inline-block;margin-top:8px}
</style></head>
<body>
<div class="card">
<h1>Access Denied</h1>
$roleNote
<p>Signed in as:</p>
<span class="acct">$($enc::HtmlEncode($LogonName))</span>
</div></body></html>
"@
	}

	function New-AuthRequiredPageHtml {
		param([string]$ServerHost)
		$enc      = [System.Net.WebUtility]
		$safeHost = $enc::HtmlEncode($ServerHost)
		return @"
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>CMPackager - Sign In</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f3f2f1;color:#323130;display:flex;align-items:center;justify-content:center;min-height:100vh}
.card{background:#fff;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.12);padding:40px 48px;max-width:560px;width:100%}
h1{font-size:20px;color:#0078d4;margin-bottom:8px}
.lead{color:#605e5c;font-size:14px;margin-bottom:24px;line-height:1.5}
h2{font-size:13px;font-weight:600;text-transform:uppercase;letter-spacing:.04em;color:#605e5c;margin-bottom:8px}
ol{padding-left:20px;font-size:14px;color:#323130;line-height:1.9;margin-bottom:20px}
code{font-family:'Cascadia Code',Consolas,monospace;background:#f3f2f1;padding:2px 6px;border-radius:3px;font-size:12px}
.cmd{background:#1e1e1e;color:#d4d4d4;padding:12px 16px;border-radius:4px;font-family:'Cascadia Code',Consolas,monospace;font-size:12px;margin:8px 0 20px;white-space:pre-wrap;word-break:break-all}
.reload{display:inline-block;margin-top:4px;color:#0078d4;font-size:14px;text-decoration:none}
.reload:hover{text-decoration:underline}
</style></head>
<body>
<div class="card">
<h1>Windows Authentication Required</h1>
<p class="lead">CMPackager uses Windows Integrated Authentication (NTLM/Kerberos). Edge and Chrome require the server to be explicitly trusted before they send credentials automatically.</p>
<h2>Option A &mdash; Internet Options (all browsers)</h2>
<ol>
<li>Open <code>inetcpl.cpl</code></li>
<li>Security &rarr; Local intranet &rarr; Sites &rarr; Advanced</li>
<li>Add <code>http://$safeHost</code> and click OK</li>
<li>Restart the browser and reload this page</li>
</ol>
<h2>Option B &mdash; Registry (run as Administrator)</h2>
<div class="cmd">reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$safeHost" /v http /t REG_DWORD /d 1 /f</div>
<a class="reload" href="javascript:location.reload()">Reload after configuring</a>
</div>
</body></html>
"@
	}

	function New-ErrorPageHtml {
		param([int]$StatusCode, [string]$Message)
		$enc = [System.Net.WebUtility]
		return @"
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>CMPackager - $StatusCode</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f3f2f1;color:#323130;display:flex;align-items:center;justify-content:center;min-height:100vh}
.card{background:#fff;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.12);padding:40px 48px;max-width:480px;text-align:center}
h1{font-size:48px;font-weight:700;color:#0078d4;margin-bottom:8px}
p{color:#605e5c;font-size:14px}
</style></head>
<body>
<div class="card"><h1>$StatusCode</h1><p>$($enc::HtmlEncode($Message))</p></div>
</body></html>
"@
	}

	function Send-WebResponse {
		param($Response, [int]$StatusCode, [string]$Body)
		$bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
		$Response.StatusCode      = $StatusCode
		$Response.ContentType     = 'text/html; charset=utf-8'
		$Response.ContentLength64 = $bytes.Length
		$Response.OutputStream.Write($bytes, 0, $bytes.Length)
		$Response.OutputStream.Close()
	}

	function Invoke-WebServerRequest {
		param($Context, [string]$RecipesFolder)
		$req      = $Context.Request
		$resp     = $Context.Response
		$identity = $Context.User.Identity

		# If the request arrived anonymously, issue a Negotiate challenge and return.
		# The browser will re-issue the request with credentials in the Authorization header.
		$isAuthenticated = $identity -and $identity.IsAuthenticated -and $identity.Name
		if (-not $isAuthenticated) {
			Add-LogContent "WebServer: anonymous request from $($req.RemoteEndPoint) - issuing Negotiate challenge"
			$resp.AddHeader('WWW-Authenticate', 'Negotiate')
			$body = New-AuthRequiredPageHtml -ServerHost $req.Url.Host
			Send-WebResponse -Response $resp -StatusCode 401 -Body $body
			return
		}

		$logonName = $identity.Name
		Add-LogContent "WebServer: $($req.HttpMethod) $($req.Url.LocalPath) from $($req.RemoteEndPoint) user=$logonName"

		$isAdmin = Test-SCCMAdminAccess -LogonName $logonName
		if ($isAdmin -ne $true) {
			$body = New-AccessDeniedPageHtml -LogonName $logonName -RequiredRole $Global:WebServerRequiredRole
			Send-WebResponse -Response $resp -StatusCode 403 -Body $body
			return
		}

		$path = $req.Url.LocalPath.TrimEnd('/')
		if ($path -eq '' -or $path -eq '/') {
			try {
				$body = New-RecipesPageHtml -RecipesFolder $RecipesFolder
				Send-WebResponse -Response $resp -StatusCode 200 -Body $body
			} catch {
				Add-LogContent "WebServer: error building recipes page - $($_.Exception.Message)"
				$body = New-ErrorPageHtml -StatusCode 503 -Message 'Failed to load recipe data.'
				Send-WebResponse -Response $resp -StatusCode 503 -Body $body
			}
		} else {
			$body = New-ErrorPageHtml -StatusCode 404 -Message 'Page not found.'
			Send-WebResponse -Response $resp -StatusCode 404 -Body $body
		}
	}

	function Register-SPNForWebServer {
		param([string]$Fqdn, [int]$Port)
		# Kerberos requires an HTTP SPN so the KDC can issue tickets for this service.
		# Without it the Negotiate handshake falls back to NTLM, which still works but
		# is slower and does not support delegation.
		$spn = "HTTP/$Fqdn`:$Port"
		$existing = & setspn -Q $spn 2>&1 | Out-String
		if ($existing -match [regex]::Escape($spn)) {
			Add-LogContent "WebServer: SPN already registered: $spn"
			return
		}
		$account = "$env:USERDOMAIN\$env:COMPUTERNAME$"
		Add-LogContent "WebServer: registering SPN $spn for $account"
		$result = & setspn -A $spn $account 2>&1 | Out-String
		Add-LogContent "WebServer: setspn result - $($result.Trim())"
		if ($result -match 'Updated object') {
			Write-Host "  SPN registered: $spn" -ForegroundColor Green
		} else {
			Write-Host "  WARNING: SPN registration failed - Kerberos will fall back to NTLM ($($result.Trim()))" -ForegroundColor Yellow
		}
	}

	function Start-CMPackagerWebServer {
		$portRef = [ref]0
		$port    = if ($Global:WebServerPort -and [int]::TryParse([string]$Global:WebServerPort, $portRef) -and $portRef.Value -ge 1 -and $portRef.Value -le 65535) { $portRef.Value } else { 8080 }
		$recipesFolder = Join-Path $PSScriptRoot 'Recipes'

		$fqdn = try { [System.Net.Dns]::GetHostEntry('localhost').HostName } catch { $env:COMPUTERNAME }

		$prefixes = Get-WebServerPrefixes -Port $port
		Register-WebServerUrlAcl -Port $port
		Register-SPNForWebServer -Fqdn $fqdn -Port $port

		$listener = New-Object System.Net.HttpListener
		# Negotiate-only: HTTP.sys owns the complete NTLM/Kerberos handshake (Type1->Challenge->
		# Type3) before queuing to GetContext(). Anonymous must NOT be included — when both are
		# set, HTTP.sys hands each message to the app individually instead of managing the exchange
		# itself, causing every NTLM Type-1 to arrive as an anonymous request and loop forever.
		$listener.AuthenticationSchemes = [System.Net.AuthenticationSchemes]::Negotiate

		# Use the strong wildcard prefix so HTTP.sys routes all requests on this port to our
		# queue regardless of the Host header (IP, short name, FQDN). Specific-hostname
		# prefixes only match exact Host headers and caused 503 when they did not match.
		$listenerPrefix = "http://+:$port/"
		$listener.Prefixes.Add($listenerPrefix)
		Write-Host ''
		Write-Host "Recommended URL (automatic Kerberos for domain members):" -ForegroundColor Yellow
		Write-Host "  http://$fqdn`:$port/" -ForegroundColor Green
		Write-Host ''
		Write-Host "Also accessible at:" -ForegroundColor Cyan
		foreach ($p in $prefixes) {
			Write-Host "  $p" -ForegroundColor Cyan
		}

		try {
			$listener.Start()
			Write-Host "CMPackager web server started. Press Ctrl+C to stop." -ForegroundColor Green
			Add-LogContent "WebServer started on port $port"

			while ($listener.IsListening) {
				$context = $null
				try {
					$context = $listener.GetContext()
				} catch [System.Net.HttpListenerException] {
					break
				}
				if ($null -eq $context) { continue }
				try {
					Invoke-WebServerRequest -Context $context -RecipesFolder $recipesFolder
				} catch {
					Add-LogContent "WebServer: unhandled request error - $($_.Exception.Message)"
					try {
						$errBody = New-ErrorPageHtml -StatusCode 503 -Message 'Internal server error.'
						Send-WebResponse -Response $context.Response -StatusCode 503 -Body $errBody
					} catch {}
				}
			}
		} catch {
			Add-LogContent "WebServer: fatal error - $($_.Exception.Message)"
			Write-Host "WebServer error: $($_.Exception.Message)" -ForegroundColor Red
		} finally {
			if ($listener.IsListening) { $listener.Stop() }
			$listener.Close()
			Add-LogContent "WebServer stopped."
			Write-Host "CMPackager web server stopped." -ForegroundColor Yellow
		}
	}

	function Test-GitHubApiAccess {
		if ($Global:GitHubToken) {
			Add-LogContent 'GitHub API: authenticated via prefs token (5,000 req/hr)'
		} elseif ($env:GITHUB_TOKEN) {
			Add-LogContent 'GitHub API: authenticated via GITHUB_TOKEN env var (5,000 req/hr)'
		} else {
			Add-LogContent 'GitHub API: unauthenticated (60 req/hr cap)'
		}

		$headers = Get-GitHubAuthHeaders -PrefsToken $Global:GitHubToken
		$headers['Accept'] = 'application/vnd.github.v3+json'

		try {
			$rateLimit = Invoke-RestMethod -Uri 'https://api.github.com/rate_limit' -Headers $headers -ErrorAction Stop
			$core      = $rateLimit.resources.core
			$remaining = [int]$core.remaining
			$limit     = [int]$core.limit
			$resetAt   = [DateTimeOffset]::FromUnixTimeSeconds([long]$core.reset).LocalDateTime.ToString('HH:mm:ss')

			if ($remaining -eq 0) {
				$msg = "GitHub API rate limit exhausted (0/$limit requests remaining, resets at $resetAt).`n  ACTION: Add or refresh a GitHub token as <GitHubToken> in CMPackager.prefs (or the GITHUB_TOKEN environment variable) to raise the limit to 5,000 req/hr."
				Add-LogContent "ERROR: $msg"
				Write-Output "ERROR: $msg"
				return $false
			}

			$statusMsg = "GitHub API: $remaining/$limit requests remaining (resets at $resetAt)."
			Add-LogContent $statusMsg
			Write-Output $statusMsg
			return $true
		}
		catch {
			$statusCode = $null
			if ($_.Exception.Response) {
				$statusCode = [int]$_.Exception.Response.StatusCode
			} elseif ($_.Exception -is [Microsoft.PowerShell.Commands.HttpResponseException]) {
				$statusCode = [int]$_.Exception.StatusCode
			}

			if ($statusCode -eq 401) {
				$msg = "GitHub API returned 401 Unauthorized. The configured token is invalid or expired.`n  ACTION: Generate a new personal access token at https://github.com/settings/tokens and set it as <GitHubToken> in CMPackager.prefs (or in the GITHUB_TOKEN environment variable)."
				Add-LogContent "ERROR: $msg"
				Write-Output "ERROR: $msg"
				return $false
			} elseif ($statusCode -eq 403) {
				$msg = "GitHub API returned 403 Forbidden. The anonymous rate limit has been reached.`n  ACTION: Add a GitHub personal access token as <GitHubToken> in CMPackager.prefs (or GITHUB_TOKEN env var) to raise the limit to 5,000 req/hr."
				Add-LogContent "ERROR: $msg"
				Write-Output "ERROR: $msg"
				return $false
			} else {
				$msg = "GitHub API connectivity check failed ($($_.Exception.Message)). Proceeding - recipes without GitHub URLs will still run."
				Add-LogContent "WARNING: $msg"
				Write-Output "WARNING: $msg"
				return $true
			}
		}
	}

	################################### MAIN ########################################
	## Startup
	if ($Setup) {
		Invoke-Setup -PreferenceFile $PreferenceFile
		exit
	}

	if ($WebServer) {
		Start-CMPackagerWebServer
		exit
	}

	Add-LogContent "--- Starting CMPackager Version $($Global:ScriptVersion) ---" -Load
	Connect-ConfigMgr

	## Create the Temp Folder if needed
	Add-LogContent "Creating CMPackager Temp Folder"
	if (-not (Test-Path $Global:TempDir)) {
		New-Item -ItemType Container -Path "$Global:TempDir" -Force -ErrorAction SilentlyContinue | Out-Null
	}

	## Allow all Cookies to download (Prevents Script from Freezing)
	Add-LogContent "Allowing All Cookies to Download (This prevents the script from freezing on a download)"
	reg add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3" /t REG_DWORD /v 1A10 /f /d 0
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

	## Create Global Conditions as defined in GlobalConditions.xml
	$GlobalConditionsXML = (([xml](Get-Content "$ScriptRoot\GlobalConditions.xml")).GlobalConditions.GlobalCondition | Where-Object Name -NE "Template" )
	Foreach ($GlobalCondition in $GlobalConditionsXML) {
		$NewGCArguments = @{ }
		$GlobalCondition.ChildNodes | ForEach-Object { if ($_.Name -ne "GCType") { $NewGCArguments[$_.Name] = $_.'#text' } }
		Push-Location
		Set-Location $Global:CMSite
		if (-not (Get-CMGlobalCondition -Name $GlobalCondition.Name)) {
			switch ($GlobalCondition.GCType) {	
				WqlQuery { 
					Add-LogContent "Creating New WQL Global Condition"
					Add-LogContent "New-CMGlobalConditionWqlQuery $NewGCArguments"
					New-CMGlobalConditionWqlQuery @NewGCArguments
				}
				Script { 
					Add-LogContent "Creating New Script Global Condition"
					Add-LogContent "New-CMGlobalConditionScript $NewGCArguments"
					New-CMGlobalConditionScript @NewGCArguments
				}
				Default {
					Add-LogContent "ERROR: Please specify a valid Global Condition Type of either WqlQuery or Script"
				}
			}
		}
		Pop-Location
	}

	## Get the Recipes
	$RecipeList = Get-ChildItem $RecipePath | Select-Object -Property Name -ExpandProperty Name | Where-Object -Property Name -NE "Template.xml" | Sort-Object -Property Name
	if (-not (Test-GitHubApiAccess)) {
		Add-LogContent 'Aborting: recipe processing will not start due to GitHub API access failure.'
		Write-Output 'Aborting: recipe processing will not start. See the log for details and corrective action.'
		exit 1
	}
	Add-LogContent -Content "All Recipes: $RecipeList"
	if (-not ([System.String]::IsNullOrEmpty($PSBoundParameters.SingleRecipe))) {
		$RecipeList = $RecipeList | Where-Object { $_ -in $PSBoundParameters.SingleRecipe }
	}
	## Begin Looping through all the Recipes 
	ForEach ($Recipe In $RecipeList) {
		## Reset All Variables
		$Download = $false
		$ApplicationCreation = $false
		$DeploymentTypeCreation = $false
		$ApplicationDistribution = $false
		$ApplicationSupersedence = $false
		$ApplicationDeployment = $false
		$ApplicationCleanup = $false
		
	
		try {
			## Import Recipe
			Add-LogContent "Importing Content for $Recipe"
			Write-Output "Begin Processing: $Recipe"
			[xml]$ApplicationRecipe = Get-Content $(Join-Path -Path $RecipePath -ChildPath $Recipe)
		
			## Perform Packaging Tasks
			Write-Output "Download"
			$Download = Start-ApplicationDownload -Recipe $ApplicationRecipe
			Add-LogContent "Continue to ApplicationCreation: $Download"
			If ($Download) {
				Write-output "Application Creation"
				$ApplicationCreation = Invoke-ApplicationCreation -Recipe $ApplicationRecipe
				Add-LogContent "Continue to DeploymentTypeCreation: $ApplicationCreation"
			}
			If ($ApplicationCreation) {
				Write-Output "Application Deployment Type Creation"
				$DeploymentTypeCreation = Add-DeploymentType -Recipe $ApplicationRecipe
				Add-LogContent "Continue to ApplicationDistribution: $DeploymentTypeCreation"
			}
			If ($DeploymentTypeCreation) {
				Write-Output "Application Distribution"
				$ApplicationDistribution = Invoke-ApplicationDistribution -Recipe $ApplicationRecipe
				Add-LogContent "Continue to Application Supersedence: $ApplicationDistribution"
			}
			If ($ApplicationDistribution) {
				Write-Output "Application Supersedence"
				$ApplicationSupersedence = Invoke-ApplicationSupersedence -Recipe $ApplicationRecipe
				Add-LogContent "Continue to Application Deployment: $ApplicationSupersedence"
			}
			If ($ApplicationSupersedence) {
				Write-Output "Application Deployment"
				$ApplicationDeployment = Invoke-ApplicationDeployment -Recipe $ApplicationRecipe
				Add-logContent "Completed Processing of $Recipe"
			}
			If ($ApplicationDeployment) {
				Write-Output "Application Cleanup"
				$ApplicationDeployment = Invoke-ApplicationCleanup -Recipe $ApplicationRecipe
				Add-logContent "Completed Processing of $Recipe"
			}
			if ($Global:TemplateApplicationCreatedFlag -eq $true) {
				Add-LogContent "WARN (LEGACY): The Requirements Application has been created, please do the following:`r`n1. Add an `"Install Behavior`" entry to the `"Templates`" deployment type of the $RequirementsTemplateAppName Application`r`n2. Run the CMPackager again to finish prerequisite setup and begin packaging software.`r`nExiting."
				Add-LogContent "THE REQUIREMENTS TEMPLATE APPLICTION IS NO LONGER NEEDED"
				Exit 0
			}
		} catch {
			Add-LogContent "Error processing ${Recipe}: $_ $($_.ScriptStackTrace)"
			Write-Error $_
		}
	}



	If ($Global:SendEmail -and $SendEmailPreference) {
		Send-EmailMessage
	}

	Add-LogContent "Cleaning Up Temp Directory $TempDir"
	Remove-Item -Path $TempDir -Recurse -Force

	## Reset all Cookies to download (Prevents Script from Freezing)
	Add-LogContent "Clearing All Cookies Download Setting"
	reg delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3" /v 1A10 /f

	Add-LogContent "--- End Of CMPackager ---"
}
