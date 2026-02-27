$files = Get-ChildItem 'c:\git\cmpackager\Disabled\*.xml' -Recurse
$withWinget = ($files | Where-Object { (Get-Content $_.FullName -Raw) -match 'Get-InstallerURLfromWinget' }).Count
$total = $files.Count
Write-Host "Total recipe files: $total"
Write-Host "Files with Get-InstallerURLfromWinget: $withWinget"
# Check for remaining old-style scraping patterns
$oldStyle = $files | Where-Object {
    $content = Get-Content $_.FullName -Raw
    $content -match 'Invoke-WebRequest.*fossies|EvergreenApp|data\.services\.jetbrains|tenable\.com/downloads/api|appstreaming\.autodesk|SoftwareHome|dl\.google\.com|corretto\.aws/downloads/latest|mozilla\.org/\?product=firefox'
}
if ($oldStyle) {
    Write-Host "Files still using old web scraping:"
    $oldStyle | ForEach-Object { Write-Host "  $($_.Name)" }
} else {
    Write-Host "No files using old web scraping patterns."
}
