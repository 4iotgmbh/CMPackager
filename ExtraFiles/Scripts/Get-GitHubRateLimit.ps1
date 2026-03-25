<#
.SYNOPSIS
    Reports the current GitHub API rate limit quota and reset time.

.DESCRIPTION
    Queries https://api.github.com/rate_limit and displays the remaining
    request quota and the UTC reset time for each resource group.

    Unauthenticated callers are capped at 60 requests/hour for the core API.
    Supplying a personal access token raises this to 5,000/hour.

.PARAMETER Token
    Optional GitHub personal access token (classic or fine-grained).
    When provided, requests are authenticated and receive a higher quota.

.EXAMPLE
    .\Get-GitHubRateLimit.ps1

.EXAMPLE
    .\Get-GitHubRateLimit.ps1 -Token $env:GITHUB_TOKEN
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Token
)

$resolvedToken = if ($Token) { $Token } elseif ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } else { $null }
$headers = @{ 'User-Agent' = 'CMPackager-RateLimitCheck' }
if ($resolvedToken) { $headers['Authorization'] = "Bearer $resolvedToken" }

try {
    $rl = Invoke-RestMethod -Uri 'https://api.github.com/rate_limit' -Headers $headers -ErrorAction Stop
} catch {
    Write-Host "ERROR: Failed to query GitHub rate limit API: $_" -ForegroundColor Red
    exit 1
}

$now = [DateTimeOffset]::UtcNow

Write-Host ''
Write-Host '  GitHub API Rate Limit' -ForegroundColor White
Write-Host '══════════════════════════════════════════════' -ForegroundColor White

$resources = [ordered]@{
    'core'              = $rl.resources.core
    'search'            = $rl.resources.search
    'graphql'           = $rl.resources.graphql
    'integration_manifest' = $rl.resources.integration_manifest
    'code_search'       = $rl.resources.code_search
}

foreach ($name in $resources.Keys) {
    $r = $resources[$name]
    if (-not $r) { continue }

    $resetAt  = [DateTimeOffset]::FromUnixTimeSeconds($r.reset)
    $waitSec  = [int](($resetAt - $now).TotalSeconds)
    $resetStr = $resetAt.LocalDateTime.ToString('HH:mm:ss')

    if ($r.remaining -eq 0) {
        $color  = 'Red'
        $status = "EXHAUSTED — resets at $resetStr (in ${waitSec}s)"
    } elseif ($r.remaining -le [math]::Max(5, $r.limit * 0.05)) {
        $color  = 'Yellow'
        $status = "$($r.remaining) / $($r.limit) — resets at $resetStr (in ${waitSec}s)"
    } else {
        $color  = 'Green'
        $status = "$($r.remaining) / $($r.limit) — resets at $resetStr"
    }

    Write-Host ("  {0,-24} {1}" -f $name, $status) -ForegroundColor $color
}

Write-Host '══════════════════════════════════════════════' -ForegroundColor White

$authMode = if ($resolvedToken) {
    if ($Token) { 'authenticated (explicit -Token)' } else { 'authenticated (GITHUB_TOKEN env var)' }
} else { 'unauthenticated (60 req/hr cap)' }
Write-Host "  Mode: $authMode" -ForegroundColor DarkGray
Write-Host ''
