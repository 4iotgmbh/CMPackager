#Requires -Version 5.1
<#
.SYNOPSIS
    Returns the direct dl.dell.com CDN URL for Dell Command Update (Classic/Win32).

.DESCRIPTION
    Two-hop approach:
      1. Fetch KB article 000177325, extract every
         /support/home/en-us/drivers/driversdetails?driverid=XXXX link.
      2. For each driver-details link (Classic only, UWP excluded), fetch that
         page and extract the dl.dell.com/FOLDER…/filename.EXE CDN URL.
         Then follow any remaining HTTP redirects to confirm the terminal blob URL.
    Outputs ONLY the resolved CDN URL — no other text.

.NOTES
    PowerShell 5.1 on Windows — no extra modules required.
    Called from DellCommandMonitor.xml PrefetchScript via powershell.exe -File.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Web

# ── Configuration ──────────────────────────────────────────────────────────────
$KbUrl      = 'https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update'
$UA         = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
$TimeoutSec = 30

$BrowserHeaders = @{
    'User-Agent'                = $UA
    'Accept'                    = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
    'Accept-Language'           = 'en-US,en;q=0.9'
    'Cache-Control'             = 'no-cache'
    'Pragma'                    = 'no-cache'
    'Sec-Ch-Ua'                 = '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"'
    'Sec-Ch-Ua-Mobile'          = '?0'
    'Sec-Ch-Ua-Platform'        = '"Windows"'
    'Sec-Fetch-Dest'            = 'document'
    'Sec-Fetch-Mode'            = 'navigate'
    'Sec-Fetch-Site'            = 'none'
    'Upgrade-Insecure-Requests' = '1'
}

# Regex for the terminal CDN blob (case-insensitive; handles .EXE and .exe)
$CdnRx    = [regex]'(?i)https?://dl\.dell\.com/FOLDER\w+/\d+/[^\s"''<>\\]+'

# Regex to find driver-details links in any text (HTML or JSON)
$DetailRx = [regex]'(?i)https?://www\.dell\.com/support/home/[^"''<>\s\\]+/drivers/driversdetails\?driverid=[A-Za-z0-9]+'

# UWP / Store / non-Win32 exclusion
$SkipRx   = [regex]'(?i)uwp|windowsapp|msix|appxbundle|appinstaller|xbox'

# ── Helper: follow HTTP redirects, return terminal URL ─────────────────────────
function Resolve-Redirects ([string]$Uri) {
    $req = [System.Net.HttpWebRequest]::Create($Uri)
    $req.Method                       = 'HEAD'
    $req.UserAgent                    = $UA
    $req.AllowAutoRedirect            = $true
    $req.MaximumAutomaticRedirections = 20
    $req.Timeout                      = $TimeoutSec * 1000
    try {
        $r   = $req.GetResponse()
        $out = $r.ResponseUri.AbsoluteUri
        $r.Close()
        return $out
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            $out = $_.Exception.Response.ResponseUri.AbsoluteUri
            $_.Exception.Response.Close()
            return $out
        }
        return $Uri
    }
    catch { return $Uri }
}

# ── Helper: fetch a page, return raw HTML (null on failure) ───────────────────
function Invoke-Page ([string]$Uri, [string]$Referer = '') {
    $h = $BrowserHeaders.Clone()
    if ($Referer) {
        $h['Referer']       = $Referer
        $h['Sec-Fetch-Site'] = 'same-origin'
    }
    try {
        (Invoke-WebRequest -Uri $Uri -Headers $h -UseBasicParsing `
            -MaximumRedirection 10 -TimeoutSec $TimeoutSec).Content
    }
    catch { $null }
}

# ── Helper: extract first CDN URL from text, skipping UWP artefacts ──────────
function Find-CdnUrl ([string]$Text) {
    foreach ($m in $CdnRx.Matches($Text)) {
        $url = [System.Web.HttpUtility]::HtmlDecode(($m.Value -replace '\\/', '/'))
        if (-not $SkipRx.IsMatch($url)) { return $url }
    }
    return $null
}

# ── Step 1: fetch KB article ──────────────────────────────────────────────────
$kbHtml = Invoke-Page -Uri $KbUrl
if (-not $kbHtml) { Write-Error "Failed to fetch KB article: $KbUrl"; return }

# ── Step 2: collect driver-details links ─────────────────────────────────────
# Search both the raw HTML and any embedded __NEXT_DATA__ JSON blob so that
# links nested inside Next.js server props are also found.
$searchText = $kbHtml
$blob = [regex]::Match($kbHtml, '(?s)<script[^>]+__NEXT_DATA__[^>]*>(.*?)</script>', 'IgnoreCase')
if ($blob.Success) { $searchText += $blob.Groups[1].Value }

$detailLinks = [System.Collections.Generic.List[string]]::new()
$seen        = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

foreach ($m in $DetailRx.Matches($searchText)) {
    $url = [System.Web.HttpUtility]::HtmlDecode(($m.Value -replace '\\/', '/'))
    if (-not $SkipRx.IsMatch($url) -and $seen.Add($url)) {
        $detailLinks.Add($url) | Out-Null
    }
}

# Also scan parsed <a> tags (PS 5.1 UseBasicParsing builds these via regex)
$tmpPage = Invoke-WebRequest -Uri $KbUrl -Headers $BrowserHeaders -UseBasicParsing -MaximumRedirection 10
foreach ($link in $tmpPage.Links) {
    $hp = $link.PSObject.Properties['href']
    if (-not $hp) { continue }
    $href = $hp.Value
    if ($href -notmatch '^https?://') { $href = "https://www.dell.com$href" }
    if ($DetailRx.IsMatch($href) -and -not $SkipRx.IsMatch($href) -and $seen.Add($href)) {
        $detailLinks.Add($href) | Out-Null
    }
}

if ($detailLinks.Count -eq 0) {
    Write-Error 'No driversdetails links found on the KB page.'
    return
}

# ── Step 3: for each driver-details page, find the CDN URL ───────────────────
$result = $null

foreach ($detailUrl in $detailLinks) {

    # 3a. HTTP-redirect shortcut — some detail links redirect straight to CDN
    $resolved = Resolve-Redirects -Uri $detailUrl
    if ($CdnRx.IsMatch($resolved) -and -not $SkipRx.IsMatch($resolved)) {
        $result = $CdnRx.Match($resolved).Value
        break
    }

    # 3b. Fetch the driver-details page and mine it for the CDN URL
    $detailHtml = Invoke-Page -Uri $detailUrl -Referer $KbUrl
    if (-not $detailHtml) { continue }

    # Search plain HTML
    $cdnUrl = Find-CdnUrl -Text $detailHtml
    if ($cdnUrl) { $result = $cdnUrl; break }

    # Search __NEXT_DATA__ blob inside the detail page
    $blob2 = [regex]::Match($detailHtml, '(?s)<script[^>]+__NEXT_DATA__[^>]*>(.*?)</script>', 'IgnoreCase')
    if ($blob2.Success) {
        $cdnUrl = Find-CdnUrl -Text $blob2.Groups[1].Value
        if ($cdnUrl) { $result = $cdnUrl; break }
    }

    # Search all other <script> blocks
    foreach ($sm in [regex]::Matches($detailHtml, '(?s)<script[^>]*>(.*?)</script>', 'IgnoreCase')) {
        $cdnUrl = Find-CdnUrl -Text $sm.Groups[1].Value
        if ($cdnUrl) { $result = $cdnUrl; break }
    }
    if ($result) { break }
}

# ── Step 4: final redirect-chase and output ───────────────────────────────────
if ($result) {
    $terminal = Resolve-Redirects -Uri $result
    if ($CdnRx.IsMatch($terminal)) { $result = $CdnRx.Match($terminal).Value }
    Write-Output $result
} else {
    Write-Error 'Could not locate the CDN download URL on any driversdetails page.'
}
