<#
.SYNOPSIS
    CMPackager Web UI — pure-PowerShell HTTP server.
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1
    powershell.exe -ExecutionPolicy Bypass -File Web\Start-WebServer.ps1 -Port 9090
#>
param([int]$Port = 8080)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$prefsFile   = Join-Path $projectRoot 'CMPackager.prefs'

# ─── Shared state (all keys pre-created before runspace pool opens) ──────────
$shared = [hashtable]::Synchronized(@{
    CMProcess    = $null
    Running      = $false
    OutputBuffer = [System.Collections.Generic.List[string]]::new()
    OutputLock   = [object]::new()
    LogPath      = $null
    CMSite       = $null
    PrefsExists  = $false
    ProjectRoot  = $projectRoot
    WebRoot      = $PSScriptRoot
    StartTime    = $null
})

# ─── Load prefs ──────────────────────────────────────────────────────────────
function Initialize-SharedState {
    if (Test-Path $prefsFile) {
        try {
            [xml]$prefs = Get-Content $prefsFile -Raw
            $shared.LogPath    = $prefs.PackagerPrefs.LogPath
            $shared.CMSite     = $prefs.PackagerPrefs.CMSite -replace ':$', ''
            $shared.PrefsExists = $true
        } catch {
            Write-Warning "Could not parse CMPackager.prefs: $_"
        }
    }
}

Initialize-SharedState

# ─── Handler scriptblock (runs inside each runspace) ─────────────────────────
$handlerScript = {
    param($ctx, $shared)

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
        # strip all path separators — prevent traversal
        return [System.IO.Path]::GetFileName($name)
    }

    function Parse-RecipeMeta($file, $state) {
        try {
            [xml]$x = Get-Content $file.FullName -Raw
            $app = $x.ApplicationDef.Application
            [PSCustomObject]@{
                file      = $file.Name
                appName   = if ($app.Name) { $app.Name } else { $file.BaseName }
                publisher = if ($app.Publisher) { $app.Publisher } else { '' }
                state     = $state
            }
        } catch {
            [PSCustomObject]@{ file = $file.Name; appName = $file.BaseName; publisher = ''; state = $state }
        }
    }

    # ── Handlers ─────────────────────────────────────────────────────────────
    function Handle-Status($ctx) {
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
        $enabled  = @(Get-ChildItem "$root\Recipes\*.xml" -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -notlike '_*' -and $_.Name -ne 'Template.xml' } |
                       ForEach-Object { Parse-RecipeMeta $_ 'enabled' })
        $disabled = @(Get-ChildItem "$root\Disabled\*.xml" -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -notlike '_*' } |
                       ForEach-Object { Parse-RecipeMeta $_ 'disabled' })
        Send-Json $ctx @{ enabled = $enabled; disabled = $disabled }
    }

    function Handle-Enable($ctx) {
        $body = Read-JsonBody $ctx
        $file = Get-SafeFilename $body.file
        $src  = Join-Path $shared.ProjectRoot "Disabled\$file"
        $dst  = Join-Path $shared.ProjectRoot "Recipes\$file"
        if (-not (Test-Path $src)) { Send-Json $ctx @{ error = 'File not found in Disabled/' } 404; return }
        if (Test-Path $dst)        { Send-Json $ctx @{ error = 'File already exists in Recipes/' } 409; return }
        Move-Item $src $dst -Force
        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Disable($ctx) {
        $body = Read-JsonBody $ctx
        $file = Get-SafeFilename $body.file
        $src  = Join-Path $shared.ProjectRoot "Recipes\$file"
        $dst  = Join-Path $shared.ProjectRoot "Disabled\$file"
        if (-not (Test-Path $src)) { Send-Json $ctx @{ error = 'File not found in Recipes/' } 404; return }
        if (Test-Path $dst)        { Send-Json $ctx @{ error = 'File already exists in Disabled/' } 409; return }
        Move-Item $src $dst -Force
        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Run($ctx) {
        if ($shared.Running) { Send-Json $ctx @{ error = 'Already running' } 409; return }
        if (-not $shared.PrefsExists) { Send-Json $ctx @{ error = 'CMPackager.prefs not found' } 412; return }

        $body   = Read-JsonBody $ctx
        $mode   = $body.mode
        $recipe = if ($body.recipe) { Get-SafeFilename $body.recipe } else { '' }

        $scriptPath = Join-Path $shared.ProjectRoot 'CMPackager.ps1'
        $args = if ($mode -eq 'single' -and $recipe) {
            "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`" -SingleRecipe `"$recipe`""
        } else {
            "-NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
        }

        $psi = [System.Diagnostics.ProcessStartInfo]::new('powershell.exe', $args)
        $psi.WorkingDirectory       = $shared.ProjectRoot
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $proc = [System.Diagnostics.Process]::new()
        $proc.StartInfo             = $psi
        $proc.EnableRaisingEvents   = $true

        # Capture shared refs for .NET delegate closures
        $capturedShared = $shared

        $proc.OutputDataReceived.Add([System.Diagnostics.DataReceivedEventHandler]{
            param($s, $e)
            if ($null -ne $e.Data) {
                $line = "[$(Get-Date -Format 'HH:mm:ss')] $($e.Data)"
                [System.Threading.Monitor]::Enter($capturedShared.OutputLock)
                try {
                    if ($capturedShared.OutputBuffer.Count -gt 5000) {
                        $capturedShared.OutputBuffer.RemoveRange(0, 500)
                    }
                    $capturedShared.OutputBuffer.Add($line)
                } finally { [System.Threading.Monitor]::Exit($capturedShared.OutputLock) }
            }
        })

        $proc.ErrorDataReceived.Add([System.Diagnostics.DataReceivedEventHandler]{
            param($s, $e)
            if ($null -ne $e.Data) {
                $line = "[$(Get-Date -Format 'HH:mm:ss')] [ERR] $($e.Data)"
                [System.Threading.Monitor]::Enter($capturedShared.OutputLock)
                try {
                    if ($capturedShared.OutputBuffer.Count -gt 5000) {
                        $capturedShared.OutputBuffer.RemoveRange(0, 500)
                    }
                    $capturedShared.OutputBuffer.Add($line)
                } finally { [System.Threading.Monitor]::Exit($capturedShared.OutputLock) }
            }
        })

        $proc.Exited.Add([System.EventHandler]{
            $capturedShared.Running   = $false
            $capturedShared.CMProcess = $null
        })

        [System.Threading.Monitor]::Enter($shared.OutputLock)
        try { $shared.OutputBuffer.Clear() }
        finally { [System.Threading.Monitor]::Exit($shared.OutputLock) }

        $proc.Start() | Out-Null
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        $shared.CMProcess = $proc
        $shared.Running   = $true
        $shared.StartTime = [datetime]::Now

        Send-Json $ctx @{ ok = $true; pid = $proc.Id; mode = $mode; recipe = $recipe }
    }

    function Handle-Stop($ctx) {
        $proc = $shared.CMProcess
        if ($proc -and -not $proc.HasExited) {
            try { $proc.Kill() } catch {}
        }
        $shared.Running   = $false
        $shared.CMProcess = $null
        Send-Json $ctx @{ ok = $true }
    }

    function Handle-Stream($ctx) {
        $resp = $ctx.Response
        $resp.ContentType = 'text/event-stream; charset=utf-8'
        $resp.SendChunked = $true
        $resp.Headers.Add('Cache-Control', 'no-cache')
        $resp.Headers.Add('X-Accel-Buffering', 'no')
        $resp.Headers.Add('Access-Control-Allow-Origin', '*')

        # Parse ?from=N so reconnecting clients skip already-seen lines
        $fromParam = $ctx.Request.QueryString['from']
        $lastIndex = if ($fromParam -match '^\d+$') { [int]$fromParam } else { 0 }

        $writer = [System.IO.StreamWriter]::new($resp.OutputStream, [System.Text.Encoding]::UTF8)
        $writer.AutoFlush = $true
        $writer.NewLine   = "`n"

        $lastLogOffset  = 0
        $heartbeatTick  = 0

        try {
            while ($true) {
                # ── Send new process output lines ────────────────────────────
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
                    # Send updated index so client can reconnect correctly
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
                        $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
                        $newContent = $reader.ReadToEnd()
                        $lastLogOffset = $fs.Position
                        $reader.Close(); $fs.Close()

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
            # Other disconnect variants
        } finally {
            try { $writer.Close() } catch {}
            try { $resp.OutputStream.Close() } catch {}
        }
    }

    function Handle-Tests($ctx) {
        $root = $shared.ProjectRoot
        $csvFiles = @(Get-ChildItem "$root\RecipeTestResults_*.csv" -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending)
        if (-not $csvFiles) {
            Send-Json $ctx @{ rows = @(); file = $null; available = $false }
            return
        }
        $latest = $csvFiles[0]
        try {
            $rows = @(Import-Csv $latest.FullName)
            Send-Json $ctx @{ rows = $rows; file = $latest.Name; available = $true }
        } catch {
            Send-Json $ctx @{ rows = @(); file = $latest.Name; available = $false; error = $_.Exception.Message }
        }
    }

    function Handle-SCCM($ctx) {
        $cmModule = Get-Module ConfigurationManager -ListAvailable -ErrorAction SilentlyContinue |
                    Select-Object -First 1
        if (-not $cmModule) {
            Send-Json $ctx @{ available = $false; message = 'ConfigurationManager module not found on this machine.' }
            return
        }

        $siteCode = $shared.CMSite
        if (-not $siteCode) {
            Send-Json $ctx @{ available = $false; message = 'CMSite not configured in CMPackager.prefs.' }
            return
        }

        try {
            Import-Module ConfigurationManager -ErrorAction Stop
            Push-Location
            Set-Location "${siteCode}:" -ErrorAction Stop

            $recipes = @(Get-ChildItem "$($shared.ProjectRoot)\Recipes\*.xml" -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -notlike '_*' -and $_.Name -ne 'Template.xml' })

            $results = @(foreach ($r in $recipes) {
                try {
                    [xml]$x = Get-Content $r.FullName -Raw
                    $appName = $x.ApplicationDef.Application.Name
                    $app = Get-CMApplication -Name $appName -Fast -ErrorAction SilentlyContinue
                    $deps = if ($app) {
                        @(Get-CMApplicationDeployment -ApplicationName $appName -ErrorAction SilentlyContinue |
                          Select-Object CollectionName, AssignmentType, DesiredConfigType, NumberTotal, NumberSuccess, NumberErrors, NumberInProgress)
                    } else { @() }
                    [PSCustomObject]@{
                        recipe      = $r.Name
                        appName     = $appName
                        found       = [bool]$app
                        version     = if ($app) { $app.SoftwareVersion } else { $null }
                        deployments = $deps
                    }
                } catch {
                    [PSCustomObject]@{ recipe = $r.Name; appName = ''; found = $false; deployments = @() }
                }
            })

            Pop-Location
            Send-Json $ctx @{ available = $true; apps = $results }
        } catch {
            try { Pop-Location -ErrorAction SilentlyContinue } catch {}
            Send-Json $ctx @{ available = $false; error = $_.Exception.Message }
        }
    }

    # ── Router ────────────────────────────────────────────────────────────────
    $req    = $ctx.Request
    $resp   = $ctx.Response
    $path   = $req.Url.AbsolutePath
    $method = $req.HttpMethod

    # CORS preflight
    if ($method -eq 'OPTIONS') {
        $resp.Headers.Add('Access-Control-Allow-Origin', '*')
        $resp.Headers.Add('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
        $resp.Headers.Add('Access-Control-Allow-Headers', 'Content-Type')
        $resp.StatusCode = 204
        $resp.Close()
        return
    }

    try {
        if     ($method -eq 'GET'  -and $path -eq '/')             { Send-File $ctx (Join-Path $shared.WebRoot 'index.html') 'text/html; charset=utf-8' }
        elseif ($method -eq 'GET'  -and $path -eq '/app.js')       { Send-File $ctx (Join-Path $shared.WebRoot 'app.js') 'application/javascript; charset=utf-8' }
        elseif ($method -eq 'GET'  -and $path -eq '/api/status')   { Handle-Status  $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/recipes')  { Handle-Recipes $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/enable')   { Handle-Enable  $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/disable')  { Handle-Disable $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/run')      { Handle-Run     $ctx }
        elseif ($method -eq 'POST' -and $path -eq '/api/stop')     { Handle-Stop    $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/stream')   { Handle-Stream  $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/tests')    { Handle-Tests   $ctx }
        elseif ($method -eq 'GET'  -and $path -eq '/api/sccm')     { Handle-SCCM    $ctx }
        else {
            $resp.StatusCode = 404
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('Not found')
            $resp.ContentLength64 = $bytes.Length
            $resp.OutputStream.Write($bytes, 0, $bytes.Length)
            $resp.OutputStream.Close()
        }
    } catch {
        try {
            $errBytes = [System.Text.Encoding]::UTF8.GetBytes((@{ error = $_.Exception.Message } | ConvertTo-Json))
            $resp.StatusCode      = 500
            $resp.ContentType     = 'application/json'
            $resp.ContentLength64 = $errBytes.Length
            $resp.OutputStream.Write($errBytes, 0, $errBytes.Length)
            $resp.OutputStream.Close()
        } catch {}
    }
}

# ─── Listener + Runspace pool ─────────────────────────────────────────────────
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")

try {
    $listener.Start()
} catch {
    Write-Host "[ERROR] Could not start listener on port $Port : $_" -ForegroundColor Red
    Write-Host "Try a different port: .\Start-WebServer.ps1 -Port 9090" -ForegroundColor Yellow
    exit 1
}

$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 8)
$pool.Open()

# Clean shutdown on Ctrl+C
Register-EngineEvent PowerShell.Exiting -Action {
    try { $listener.Stop() } catch {}
    try { $pool.Close() } catch {}
} | Out-Null

Write-Host ""
Write-Host "  CMPackager Web UI" -ForegroundColor Cyan
Write-Host "  http://localhost:$Port/" -ForegroundColor Green
Write-Host "  Project root : $projectRoot" -ForegroundColor DarkGray
Write-Host "  Prefs loaded : $($shared.PrefsExists)" -ForegroundColor DarkGray
Write-Host "  Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

# ─── Main accept loop ─────────────────────────────────────────────────────────
while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
    } catch [System.Net.HttpListenerException] {
        break   # listener stopped
    } catch {
        continue
    }

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $pool
    $ps.AddScript($handlerScript).AddArgument($ctx).AddArgument($shared) | Out-Null
    $ps.BeginInvoke() | Out-Null
}

$listener.Close()
$pool.Close()
Write-Host "Server stopped." -ForegroundColor DarkGray
