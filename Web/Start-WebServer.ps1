<#
.SYNOPSIS
    CMPackager Web UI — pure-PowerShell HTTP server.
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1
    powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1 -Port 9090
    powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1 -DebugMode
#>
param(
    [int]$Port = 8080,
    [switch]$DebugMode,
    [string]$AuditLogPath = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$prefsFile   = Join-Path $projectRoot 'CMPackager.prefs'

# ─── Shared state (all keys pre-created before runspace pool opens) ──────────
$shared = [hashtable]::Synchronized(@{
    CMProcess      = $null
    Running        = $false
    OutputBuffer   = [System.Collections.Generic.List[string]]::new()
    OutputLock     = [object]::new()
    LogPath        = $null
    CMSite         = $null
    CMPSModulePath = $null
    PrefsExists    = $false
    PrefsFile      = $prefsFile
    AuditLogPath   = ''
    ProjectRoot    = $projectRoot
    WebRoot        = $PSScriptRoot
    StartTime      = $null
    DebugMode      = $DebugMode.IsPresent
    ReaderRS       = $null   # Keep reader runspaces alive (GC prevention)
    ReaderPS       = $null
})

# ─── Load prefs ──────────────────────────────────────────────────────────────
function Initialize-SharedState {
    if (Test-Path $prefsFile) {
        try {
            [xml]$prefs = Get-Content $prefsFile -Raw
            $shared.LogPath        = $prefs.PackagerPrefs.LogPath
            $shared.CMSite         = $prefs.PackagerPrefs.CMSite -replace ':$', ''
            $shared.CMPSModulePath = $prefs.PackagerPrefs.CMPSModulePath
            $shared.AuditLogPath   = $prefs.PackagerPrefs.AuditLogPath
            $shared.PrefsExists    = $true
        } catch {
            Write-Warning "Could not parse CMPackager.prefs: $_"
        }
    }
}

Initialize-SharedState
# -AuditLogPath CLI parameter overrides prefs value
if ($AuditLogPath) { $shared.AuditLogPath = $AuditLogPath }

if ($DebugMode) {
    Write-Host "[DEBUG] DebugMode on  |  ProjectRoot: $projectRoot" -ForegroundColor DarkCyan
    Write-Host "[DEBUG] PrefsFile: $prefsFile  |  PrefsExists: $($shared.PrefsExists)" -ForegroundColor DarkCyan
    Write-Host "[DEBUG] LogPath: $($shared.LogPath)  |  CMSite: $($shared.CMSite)" -ForegroundColor DarkCyan
    Write-Host "[DEBUG] CMPSModulePath: $($shared.CMPSModulePath)" -ForegroundColor DarkCyan
    Write-Host "[DEBUG] AuditLogPath: $($shared.AuditLogPath)" -ForegroundColor DarkCyan
}

# ─── Handler scriptblock (runs inside each runspace) ─────────────────────────
$handlerScript = {
    param($ctx, $shared)

    # ── Debug logger — Write-Host writes to the shared PSHost from pool runspaces ──
    function Write-Dbg($msg) {
        if ($shared.DebugMode) {
            Write-Host "[DEBUG $(Get-Date -Format 'HH:mm:ss')] $msg" -ForegroundColor DarkCyan
        }
    }

    # ── Helpers ──────────────────────────────────────────────────────────────
    function Send-Json($ctx, $obj, [int]$status = 200) {
        $json  = $obj | ConvertTo-Json -Depth 10 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $ctx.Response.StatusCode      = $status
        $ctx.Response.ContentType     = 'application/json; charset=utf-8'
        $ctx.Response.ContentLength64 = $bytes.Length
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $ctx.Response.OutputStream.Close()
    }

    function Send-File($ctx, $path, $mime) {
        if (-not (Test-Path $path)) { $ctx.Response.StatusCode = 404; $ctx.Response.Close(); return }
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $ctx.Response.ContentType     = $mime
        $ctx.Response.ContentLength64 = $bytes.Length
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $ctx.Response.OutputStream.Close()
    }

    function Read-JsonBody($ctx) {
        $reader = [System.IO.StreamReader]::new($ctx.Request.InputStream, [System.Text.Encoding]::UTF8)
        return $reader.ReadToEnd() | ConvertFrom-Json
    }

    function Get-SafeFilename($name) {
        return [System.IO.Path]::GetFileName($name)
    }

    function Parse-RecipeMeta($file, $state) {
        try {
            [xml]$x = Get-Content $file.FullName -Raw
            $app = $x.ApplicationDef.Application
            [PSCustomObject]@{
                file      = $file.Name
                appName   = if ($app.Name)      { $app.Name }      else { $file.BaseName }
                publisher = if ($app.Publisher) { $app.Publisher } else { '' }
                state     = $state
            }
        } catch {
            [PSCustomObject]@{ file = $file.Name; appName = $file.BaseName; publisher = ''; state = $state }
        }
    }

    # ── Handlers ─────────────────────────────────────────────────────────────
    function Handle-Status($ctx) {
        # Sync running state with actual process state
        if ($shared.Running -and $shared.CMProcess -and $shared.CMProcess.HasExited) {
            $shared.Running   = $false
            $shared.CMProcess = $null
        }

        $lines = $null
        [System.Threading.Monitor]::Enter($shared.OutputLock)
        try   { $lines = @($shared.OutputBuffer.ToArray()) }
        finally { [System.Threading.Monitor]::Exit($shared.OutputLock) }

        $last50 = if ($lines.Count -gt 50) { $lines[($lines.Count - 50)..($lines.Count - 1)] } else { $lines }
        Send-Json $ctx @{
            running     = $shared.Running
            prefsExists = $shared.PrefsExists
            logPath     = $shared.LogPath
            lastLines   = $last50
            totalLines  = $lines.Count
            startTime   = if ($shared.StartTime) { $shared.StartTime.ToString('o') } else { $null }
        }
    }

    function Handle-Recipes($ctx) {
        $root = $shared.ProjectRoot

        $enabled = @(Get-ChildItem "$root\Recipes\*.xml" -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name -notlike '_*' -and $_.Name -ne 'Template.xml' } |
                     ForEach-Object { Parse-RecipeMeta $_ 'enabled' })

        # Build a set of enabled filenames so we can exclude them from disabled
        $enabledNames = @{}
        $enabled | ForEach-Object { $enabledNames[$_.file] = $true }

        $disabled = @(Get-ChildItem "$root\Disabled\*.xml" -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -notlike '_*' -and -not $enabledNames.ContainsKey($_.Name) } |
                      ForEach-Object { Parse-RecipeMeta $_ 'disabled' })

        Write-Dbg "Recipes: $($enabled.Count) enabled, $($disabled.Count) disabled"
        Send-Json $ctx @{ enabled = $enabled; disabled = $disabled }
    }

    function Handle-Enable($ctx) {
        $body = Read-JsonBody $ctx
        $file = Get-SafeFilename $body.file
        $src  = Join-Path $shared.ProjectRoot "Disabled\$file"
        $dst  = Join-Path $shared.ProjectRoot "Recipes\$file"
        Write-Dbg "Enable: $src -> $dst"
        if (-not (Test-Path $src)) { Send-Json $ctx @{ error = 'File not found in Disabled/' } 404; return }
        if (Test-Path $dst)        { Send-Json $ctx @{ error = 'Already exists in Recipes/' } 409; return }
        Move-Item $src $dst -Force
        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Disable($ctx) {
        $body = Read-JsonBody $ctx
        $file = Get-SafeFilename $body.file
        $src  = Join-Path $shared.ProjectRoot "Recipes\$file"
        $dst  = Join-Path $shared.ProjectRoot "Disabled\$file"
        Write-Dbg "Disable: $src -> $dst"
        if (-not (Test-Path $src)) { Send-Json $ctx @{ error = 'File not found in Recipes/' } 404; return }
        if (Test-Path $dst)        { Send-Json $ctx @{ error = 'Already exists in Disabled/' } 409; return }
        Move-Item $src $dst -Force
        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Run($ctx) {
        if ($shared.Running) { Send-Json $ctx @{ error = 'Already running' } 409; return }
        if (-not $shared.PrefsExists) { Send-Json $ctx @{ error = 'CMPackager.prefs not found' } 412; return }

        $body   = Read-JsonBody $ctx
        $mode   = $body.mode
        $recipe = if ($body.recipe) { Get-SafeFilename $body.recipe } else { '' }

        $scriptPath  = Join-Path $shared.ProjectRoot 'CMPackager.ps1'
        $recipesPath = Join-Path $shared.ProjectRoot 'Recipes'
        $prefsArg    = " -PreferenceFile `"$($shared.PrefsFile)`" -RecipePath `"$recipesPath`""
        $psArgs = if ($mode -eq 'single' -and $recipe) {
            "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`"$prefsArg -SingleRecipe `"$recipe`""
        } else {
            "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`"$prefsArg"
        }

        Write-Dbg "Run: powershell.exe $psArgs"

        $psi = [System.Diagnostics.ProcessStartInfo]::new('powershell.exe', $psArgs)
        $psi.WorkingDirectory       = $shared.ProjectRoot
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $proc = [System.Diagnostics.Process]::new()
        $proc.StartInfo = $psi

        # Clear output buffer before starting
        [System.Threading.Monitor]::Enter($shared.OutputLock)
        try { $shared.OutputBuffer.Clear() }
        finally { [System.Threading.Monitor]::Exit($shared.OutputLock) }

        $proc.Start() | Out-Null
        Write-Dbg "Process started, PID $($proc.Id)"

        $shared.CMProcess = $proc
        $shared.Running   = $true
        $shared.StartTime = [datetime]::Now

        # ── Dedicated reader runspaces — avoids .NET delegate/runspace-context issues ──
        # Each runspace does a blocking ReadLine() loop and writes to the shared buffer.
        # References are stored in $shared so they are not garbage-collected while running.

        $readerScript = {
            param($stream, $shared, $prefix)
            try {
                while ($true) {
                    $line = $stream.ReadLine()
                    if ($null -eq $line) { break }
                    $ts = "[$(Get-Date -Format 'HH:mm:ss')]$prefix $line"
                    [System.Threading.Monitor]::Enter($shared.OutputLock)
                    try {
                        if ($shared.OutputBuffer.Count -gt 5000) {
                            $shared.OutputBuffer.RemoveRange(0, 500)
                        }
                        $shared.OutputBuffer.Add($ts)
                    } finally {
                        [System.Threading.Monitor]::Exit($shared.OutputLock)
                    }
                }
            } catch {}
        }

        $rs1 = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs1.Open()
        $ps1 = [System.Management.Automation.PowerShell]::Create()
        $ps1.Runspace = $rs1
        $ps1.AddScript($readerScript).AddArgument($proc.StandardOutput).AddArgument($shared).AddArgument('') | Out-Null
        $ps1.BeginInvoke() | Out-Null

        $rs2 = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs2.Open()
        $ps2 = [System.Management.Automation.PowerShell]::Create()
        $ps2.Runspace = $rs2
        $ps2.AddScript($readerScript).AddArgument($proc.StandardError).AddArgument($shared).AddArgument(' [ERR]') | Out-Null
        $ps2.BeginInvoke() | Out-Null

        # Keep references alive — without this the GC can collect the PS/RS objects
        # while the async operation is still running
        $shared.ReaderRS = @($rs1, $rs2)
        $shared.ReaderPS = @($ps1, $ps2)

        Send-Json $ctx @{ ok = $true; pid = $proc.Id; mode = $mode; recipe = $recipe }
    }

    function Handle-Stop($ctx) {
        $proc = $shared.CMProcess
        if ($proc -and -not $proc.HasExited) {
            try { $proc.Kill() } catch {}
            Write-Dbg "Killed PID $($proc.Id)"
        }
        $shared.Running   = $false
        $shared.CMProcess = $null

        # Close reader runspaces
        if ($shared.ReaderRS) {
            foreach ($rs in $shared.ReaderRS) { try { $rs.Close() } catch {} }
            $shared.ReaderRS = $null
            $shared.ReaderPS = $null
        }

        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Stream($ctx) {
        $resp = $ctx.Response
        $resp.ContentType = 'text/event-stream; charset=utf-8'
        $resp.SendChunked = $true
        $resp.Headers.Add('Cache-Control', 'no-cache')
        $resp.Headers.Add('X-Accel-Buffering', 'no')
        $resp.Headers.Add('Access-Control-Allow-Origin', '*')

        $fromParam = $ctx.Request.QueryString['from']
        $lastIndex = if ($fromParam -match '^\d+$') { [int]$fromParam } else { 0 }

        $writer = [System.IO.StreamWriter]::new($resp.OutputStream, [System.Text.Encoding]::UTF8)
        $writer.AutoFlush = $true
        $writer.NewLine   = "`n"

        $lastLogOffset = 0
        $heartbeatTick = 0

        try {
            while ($true) {
                # ── New process output lines ──────────────────────────────────
                $lines = $null
                [System.Threading.Monitor]::Enter($shared.OutputLock)
                try   { $lines = @($shared.OutputBuffer.ToArray()) }
                finally { [System.Threading.Monitor]::Exit($shared.OutputLock) }

                for ($i = $lastIndex; $i -lt $lines.Count; $i++) {
                    $escaped = $lines[$i] -replace "`n", ' '
                    $writer.WriteLine("data: $escaped")
                    $writer.WriteLine('')
                }
                if ($lines.Count -gt $lastIndex) {
                    $lastIndex = $lines.Count
                    $writer.WriteLine("event: index")
                    $writer.WriteLine("data: $lastIndex")
                    $writer.WriteLine('')
                }

                # ── Tail log file ─────────────────────────────────────────────
                $logPath = $shared.LogPath
                if ($logPath -and (Test-Path $logPath)) {
                    try {
                        $fs = [System.IO.FileStream]::new(
                            $logPath,
                            [System.IO.FileMode]::Open,
                            [System.IO.FileAccess]::Read,
                            [System.IO.FileShare]::ReadWrite
                        )
                        $fs.Seek($lastLogOffset, [System.IO.SeekOrigin]::Begin) | Out-Null
                        $sr = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
                        $newContent = $sr.ReadToEnd()
                        $lastLogOffset = $fs.Position
                        $sr.Close(); $fs.Close()

                        if ($newContent.Length -gt 0) {
                            foreach ($logLine in ($newContent -split "`r?`n")) {
                                if ($logLine.Trim()) {
                                    $escaped = $logLine -replace "`n", ' '
                                    $writer.WriteLine("event: log")
                                    $writer.WriteLine("data: $escaped")
                                    $writer.WriteLine('')
                                }
                            }
                        }
                    } catch { $lastLogOffset = 0 }
                }

                # ── Heartbeat every ~15 s ─────────────────────────────────────
                $heartbeatTick++
                if ($heartbeatTick -ge 30) {
                    $writer.WriteLine(': heartbeat')
                    $writer.WriteLine('')
                    $heartbeatTick = 0
                }

                Start-Sleep -Milliseconds 500
            }
        } catch [System.IO.IOException] {
            # Client disconnected — normal
        } catch [System.Exception] {
            # Other disconnect variants (ObjectDisposedException, etc.)
        } finally {
            try { $writer.Close() } catch {}
            try { $resp.OutputStream.Close() } catch {}
        }
    }

    function Handle-Tests($ctx) {
        $searchPaths = @($shared.ProjectRoot)
        if ($shared.AuditLogPath -and (Test-Path $shared.AuditLogPath -ErrorAction SilentlyContinue)) {
            $searchPaths += $shared.AuditLogPath
        }
        $csvFiles = @(
            foreach ($p in $searchPaths) {
                Get-ChildItem "$p\RecipeTestResults_*.csv" -ErrorAction SilentlyContinue
            }
        ) | Sort-Object LastWriteTime -Descending
        Write-Dbg "Tests: searched $($searchPaths -join ', '), found $($csvFiles.Count) file(s)"
        if (-not $csvFiles) {
            Send-Json $ctx @{ rows = @(); file = $null; available = $false }
            return
        }
        $latest = $csvFiles[0]
        Write-Dbg "Tests: using $($latest.Name)"
        try {
            $rows = @(Import-Csv $latest.FullName)
            Send-Json $ctx @{ rows = $rows; file = $latest.Name; available = $true }
        } catch {
            Send-Json $ctx @{ rows = @(); file = $latest.Name; available = $false; error = $_.Exception.Message }
        }
    }

    function Handle-SCCM($ctx) {
        $siteCode = $shared.CMSite
        if (-not $siteCode) {
            Send-Json $ctx @{ available = $false; message = 'CMSite not configured in CMPackager.prefs.' }
            return
        }

        # Resolve ConfigurationManager module path (prefs → PSModulePath → SMS_ADMIN_UI_PATH env var)
        $modulePath = $null

        if ($shared.CMPSModulePath) {
            $candidate = Join-Path $shared.CMPSModulePath 'ConfigurationManager.psd1'
            if (Test-Path $candidate) { $modulePath = $candidate }
            Write-Dbg "SCCM: CMPSModulePath candidate: $candidate  found: $([bool]$modulePath)"
        }

        if (-not $modulePath) {
            $m = Get-Module ConfigurationManager -ListAvailable -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($m) { $modulePath = $m.Path }
            Write-Dbg "SCCM: PSModulePath search: $modulePath"
        }

        if (-not $modulePath -and $env:SMS_ADMIN_UI_PATH) {
            $smsItem = Get-Item $env:SMS_ADMIN_UI_PATH -ErrorAction SilentlyContinue
            if ($smsItem) {
                $candidate = Join-Path $smsItem.Parent.FullName 'ConfigurationManager.psd1'
                if (Test-Path $candidate) { $modulePath = $candidate }
            }
            Write-Dbg "SCCM: SMS_ADMIN_UI_PATH candidate: $candidate  found: $([bool]$modulePath)"
        }

        if (-not $modulePath) {
            Send-Json $ctx @{ available = $false; message = 'ConfigurationManager module not found. Install SCCM console or set CMPSModulePath in CMPackager.prefs.' }
            return
        }

        try {
            Import-Module $modulePath -ErrorAction Stop
            Push-Location
            Set-Location "${siteCode}:" -ErrorAction Stop
            Write-Dbg "SCCM: connected to ${siteCode}:"

            $recipes = @(Get-ChildItem "$($shared.ProjectRoot)\Recipes\*.xml" -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -notlike '_*' -and $_.Name -ne 'Template.xml' })

            $results = @(foreach ($r in $recipes) {
                try {
                    [xml]$x = Get-Content $r.FullName -Raw
                    $appName = $x.ApplicationDef.Application.Name
                    Write-Dbg "SCCM: querying '$appName*'"

                    # CMPackager names apps "$Name $Version" — use wildcard to find any version
                    $apps = @(Get-CMApplication -Name "$appName*" -Fast -ErrorAction SilentlyContinue)
                    # Pick the newest (by DateCreated or SoftwareVersion) — not superseded/expired
                    $app = $apps | Where-Object { -not $_.IsExpired -and -not $_.IsSuperseded } |
                                   Sort-Object DateCreated -Descending | Select-Object -First 1
                    if (-not $app) { $app = $apps | Sort-Object DateCreated -Descending | Select-Object -First 1 }

                    $deps = @()
                    if ($app) {
                        $deps = @(Get-CMDeployment -SoftwareName $app.LocalizedDisplayName -ErrorAction SilentlyContinue |
                                  Select-Object CollectionName, AssignmentAction,
                                                NumberTargeted, NumberSuccess, NumberErrors, NumberInProgress, NumberOther, NumberUnknown)
                    }

                    [PSCustomObject]@{
                        recipe       = $r.Name
                        appName      = $appName
                        sccmName     = if ($app) { $app.LocalizedDisplayName } else { $null }
                        found        = [bool]$app
                        version      = if ($app) { $app.SoftwareVersion } else { $null }
                        allVersions  = $apps.Count
                        deployments  = $deps
                    }
                } catch {
                    Write-Dbg "SCCM: error for $($r.Name): $_"
                    [PSCustomObject]@{ recipe = $r.Name; appName = ''; sccmName = $null; found = $false; version = $null; allVersions = 0; deployments = @() }
                }
            })

            Pop-Location
            Send-Json $ctx @{ available = $true; apps = $results }
        } catch {
            try { Pop-Location -ErrorAction SilentlyContinue } catch {}
            Write-Dbg "SCCM: exception: $_"
            Send-Json $ctx @{ available = $false; error = $_.Exception.Message }
        }
    }

    # ── Router ────────────────────────────────────────────────────────────────
    $req    = $ctx.Request
    $resp   = $ctx.Response
    $path   = $req.Url.AbsolutePath
    $method = $req.HttpMethod

    Write-Dbg "$method $path"

    if ($method -eq 'OPTIONS') {
        $resp.Headers.Add('Access-Control-Allow-Origin', '*')
        $resp.Headers.Add('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        $resp.Headers.Add('Access-Control-Allow-Headers', 'Content-Type')
        $resp.StatusCode = 204
        $resp.Close()
        return
    }

    try {
        if     ($method -eq 'GET'  -and $path -eq '/')            { Send-File $ctx (Join-Path $shared.WebRoot 'index.html') 'text/html; charset=utf-8' }
        elseif ($method -eq 'GET'  -and $path -eq '/app.js')      { Send-File $ctx (Join-Path $shared.WebRoot 'app.js') 'application/javascript; charset=utf-8' }
        elseif ($method -eq 'GET'  -and $path -eq '/api/status')  { Handle-Status  $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/recipes') { Handle-Recipes $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/enable')  { Handle-Enable  $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/disable') { Handle-Disable $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/run')     { Handle-Run     $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/stop')    { Handle-Stop    $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/stream')  { Handle-Stream  $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/tests')   { Handle-Tests   $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/sccm')    { Handle-SCCM    $ctx }
        else {
            $resp.StatusCode = 404
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('Not found')
            $resp.ContentLength64 = $bytes.Length
            $resp.OutputStream.Write($bytes, 0, $bytes.Length)
            $resp.OutputStream.Close()
        }
    } catch {
        try {
            $errJson  = @{ error = $_.Exception.Message } | ConvertTo-Json
            $errBytes = [System.Text.Encoding]::UTF8.GetBytes($errJson)
            $resp.StatusCode      = 500
            $resp.ContentType     = 'application/json'
            $resp.ContentLength64 = $errBytes.Length
            $resp.OutputStream.Write($errBytes, 0, $errBytes.Length)
            $resp.OutputStream.Close()
        } catch {}
        Write-Dbg "Handler exception [$method $path]: $($_.Exception.Message)"
    }
}

# ─── Listener setup with retry (port may linger after hard kill) ──────────────
$listener = $null
$pool      = $null

$maxRetries = 5
for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:$Port/")
    try {
        $listener.Start()
        break
    } catch {
        $listener.Close()
        if ($attempt -eq $maxRetries) {
            Write-Host "[ERROR] Cannot bind to port $Port after $maxRetries attempts: $_" -ForegroundColor Red
            Write-Host "Try: .\Start-WebServer.ps1 -Port 9090" -ForegroundColor Yellow
            exit 1
        }
        Write-Host "  Port $Port busy, retrying in 2s... ($attempt/$maxRetries)" -ForegroundColor Yellow
        Start-Sleep 2
        $listener = $null
    }
}

$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 8)
$pool.Open()

Write-Host ""
Write-Host "  CMPackager Web UI" -ForegroundColor Cyan
Write-Host "  http://localhost:$Port/" -ForegroundColor Green
Write-Host "  Project root : $projectRoot" -ForegroundColor DarkGray
Write-Host "  Prefs loaded : $($shared.PrefsExists)" -ForegroundColor DarkGray
if ($DebugMode) { Write-Host "  Debug mode   : ON" -ForegroundColor DarkCyan }
Write-Host "  Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

# ─── Main accept loop — try/finally guarantees cleanup on Ctrl+C ─────────────
try {
    while ($listener.IsListening) {
        try {
            $ctx = $listener.GetContext()
        } catch [System.Net.HttpListenerException] {
            break
        } catch {
            continue
        }

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $pool
        $ps.AddScript($handlerScript).AddArgument($ctx).AddArgument($shared) | Out-Null
        $ps.BeginInvoke() | Out-Null
    }
} finally {
    # Always runs on Ctrl+C, normal exit, or exception
    try { $listener.Stop()  } catch {}
    try { $listener.Close() } catch {}
    try { $pool.Close()     } catch {}

    # Kill any running CMPackager process
    if ($shared.CMProcess -and -not $shared.CMProcess.HasExited) {
        try { $shared.CMProcess.Kill() } catch {}
    }
    if ($shared.ReaderRS) {
        foreach ($rs in $shared.ReaderRS) { try { $rs.Close() } catch {} }
    }

    Write-Host "Server stopped." -ForegroundColor DarkGray
}
