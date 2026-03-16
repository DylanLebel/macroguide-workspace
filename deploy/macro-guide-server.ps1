#Requires -Version 5.1
<#
.SYNOPSIS
    Nordic Minesteel Macro Guide — Server
    Maintained by Dylan Lebel

.DESCRIPTION
    Serves MacroGuide.html and reads/writes changelog + ticket JSON
    on the shared drive. Uses file locking for concurrent access.
    Runs on any Windows machine — no installs required.

.NOTES
    Started automatically by "Open Macro Guide.vbs"
    Press Ctrl+C to stop manually.
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ════════════════════════════════════════════════════════════════════════

$Port = 8123

# Paths — auto-detect relative to script, fall back to known locations
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$DataDir = Join-Path $ScriptDir 'data'
if (-not (Test-Path $DataDir)) {
    $DataDir = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\data'
}

$HtmlFile = Join-Path $ScriptDir 'MacroGuide.html'
if (-not (Test-Path $HtmlFile)) {
    $HtmlFile = 'C:\AllMacros\MacroGuide.html'
}

$ClFile = Join-Path $DataDir 'changelog.json'
$TkFile = Join-Path $DataDir 'tickets.json'

# Email notifications — set $EmailEnabled to $true and configure SMTP
$EmailEnabled = $false
$EmailSmtp    = 'mail.nmtech.com'
$EmailPort    = 25
$EmailFrom    = 'macroguide@nmtech.com'
$EmailTo      = 'dlebel@nmtech.com'

# ════════════════════════════════════════════════════════════════════════
#  FILE LOCKING (cross-machine safe via lock files on shared drive)
# ════════════════════════════════════════════════════════════════════════

$LockStaleMs  = 10000
$MaxRetries   = 6
$RetryBaseMs  = 150

function Get-Lock {
    param([string]$FilePath)
    $lockPath = "$FilePath.lock"
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            $fs = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            $info = @{ pid = $PID; host = $env:COMPUTERNAME; time = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() } | ConvertTo-Json -Compress
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($info)
            $fs.Write($bytes, 0, $bytes.Length)
            $fs.Close()
            return $true
        } catch {
            if (Test-Path $lockPath) {
                try {
                    $raw = Get-Content $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    $age = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() - $raw.time
                    if ($age -gt $LockStaleMs) {
                        Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
                        continue
                    }
                } catch {
                    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
                    continue
                }
            }
            $wait = [math]::Min($RetryBaseMs * [math]::Pow(2, $i), 3000) + (Get-Random -Minimum 0 -Maximum 50)
            Start-Sleep -Milliseconds $wait
        }
    }
    return $false
}

function Release-Lock {
    param([string]$FilePath)
    $lockPath = "$FilePath.lock"
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
}

# ════════════════════════════════════════════════════════════════════════
#  SAFE JSON I/O
# ════════════════════════════════════════════════════════════════════════

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        $raw = Get-Content $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) { return @() }
        # Ensure it's always an array
        if ($data -isnot [System.Array]) { $data = @($data) }
        return $data
    } catch {
        $bak = "$Path.corrupt.$([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())"
        Copy-Item $Path $bak -Force -ErrorAction SilentlyContinue
        Write-Host "  WARNING: Corrupt JSON at $Path -> backed up" -ForegroundColor Yellow
        return @()
    }
}

function Write-JsonFile {
    param([string]$Path, [object]$Data)
    $tmp = "$Path.tmp.$PID"
    $json = ConvertTo-Json $Data -Depth 10 -Compress:$false
    # Fix PowerShell's Unicode escaping — write raw UTF8
    [System.IO.File]::WriteAllText($tmp, $json, [System.Text.Encoding]::UTF8)
    Move-Item $tmp $Path -Force
}

function Invoke-LockedMutate {
    param([string]$FilePath, [scriptblock]$Mutate)
    $locked = Get-Lock $FilePath
    if (-not $locked) {
        throw "File is busy - another user is saving right now. Try again in a moment."
    }
    try {
        $data = @(Read-JsonFile $FilePath)
        $data = & $Mutate $data
        Write-JsonFile $FilePath $data
        return $data
    } finally {
        Release-Lock $FilePath
    }
}

# ════════════════════════════════════════════════════════════════════════
#  EMAIL NOTIFICATIONS
# ════════════════════════════════════════════════════════════════════════

function Send-TicketEmail {
    param($Ticket)
    if (-not $EmailEnabled) { return }
    try {
        $subject = "New Macro Guide Ticket: $($Ticket.title)"
        $body = @"
A new ticket was submitted in the Macro Guide.

Title:    $($Ticket.title)
Type:     $($Ticket.type)
Priority: $($Ticket.priority)
Macro:    $($Ticket.macroName)
From:     $(if ($Ticket.createdBy) { $Ticket.createdBy } else { 'Anonymous' })
Date:     $($Ticket.date)

Description:
$(if ($Ticket.description) { $Ticket.description } else { '(no description)' })
"@
        Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $subject -Body $body -SmtpServer $EmailSmtp -Port $EmailPort -ErrorAction Stop
        Write-Host "  Email sent: $subject" -ForegroundColor Green
    } catch {
        Write-Host "  Email failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ════════════════════════════════════════════════════════════════════════
#  ENSURE DATA FILES EXIST
# ════════════════════════════════════════════════════════════════════════

function Initialize-DataFiles {
    if (-not (Test-Path $DataDir)) {
        try {
            New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
            Write-Host "  Created data directory: $DataDir"
        } catch {
            throw "Data directory not found and could not be created:`n  $DataDir`n`nMake sure the Y: drive is mapped."
        }
    }
    if (-not (Test-Path $ClFile)) {
        $seed = @(@{
            macroId = 'all'; macroName = 'All Macros'; macroIcon = [char]0x2699 + [char]0xFE0F; macroColor = '#4f8ef7'
            date = (Get-Date -Format 'yyyy-MM-dd'); type = 'doc'
            description = 'Macro Guide created - this interactive reference page launched for the team.'
            author = 'Dylan'
        })
        Write-JsonFile $ClFile $seed
        Write-Host "  Created changelog.json (seed entry)"
    }
    if (-not (Test-Path $TkFile)) {
        [System.IO.File]::WriteAllText($TkFile, '[]', [System.Text.Encoding]::UTF8)
        Write-Host "  Created tickets.json (empty)"
    }
}

# ════════════════════════════════════════════════════════════════════════
#  HTTP HELPERS
# ════════════════════════════════════════════════════════════════════════

function Send-JsonResponse {
    param($Response, [int]$StatusCode, $Data)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = 'application/json; charset=utf-8'
    $Response.AddHeader('Access-Control-Allow-Origin', '*')
    $Response.AddHeader('Access-Control-Allow-Methods', 'GET,POST,PATCH,OPTIONS')
    $Response.AddHeader('Access-Control-Allow-Headers', 'Content-Type')
    $json = ConvertTo-Json $Data -Depth 10 -Compress
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

function Send-HtmlResponse {
    param($Response, [string]$FilePath)
    $Response.StatusCode = 200
    $Response.ContentType = 'text/html; charset=utf-8'
    $Response.AddHeader('Cache-Control', 'no-cache')
    $Response.AddHeader('Access-Control-Allow-Origin', '*')
    $buffer = [System.IO.File]::ReadAllBytes($FilePath)
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $Response.OutputStream.Close()
}

function Read-RequestBody {
    param($Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
    $body = $reader.ReadToEnd()
    $reader.Close()
    return $body
}

# ════════════════════════════════════════════════════════════════════════
#  HTTP SERVER
# ════════════════════════════════════════════════════════════════════════

function Start-Server {
    Write-Host ""
    Write-Host "  Nordic Minesteel Macro Guide Server" -ForegroundColor Cyan
    Write-Host "  ===================================" -ForegroundColor Cyan
    Write-Host ""

    # Validate
    if (-not (Test-Path $HtmlFile)) {
        Write-Host "  ERROR: MacroGuide.html not found at $HtmlFile" -ForegroundColor Red
        return
    }

    Initialize-DataFiles

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:${Port}/")

    try {
        $listener.Start()
    } catch {
        Write-Host "  ERROR: Could not start on port $Port." -ForegroundColor Red
        Write-Host "  The Macro Guide may already be running. Try: http://localhost:$Port" -ForegroundColor Yellow
        Write-Host ""
        return
    }

    Write-Host "  Server running    : http://localhost:$Port" -ForegroundColor Green
    Write-Host "  HTML source       : $HtmlFile"
    Write-Host "  Data directory    : $DataDir"
    if ($EmailEnabled) {
        Write-Host "  Email alerts      : $EmailTo (via $EmailSmtp)" -ForegroundColor Green
    } else {
        Write-Host "  Email alerts      : disabled"
    }
    Write-Host ""
    Write-Host "  Press Ctrl+C to stop." -ForegroundColor DarkGray
    Write-Host ""

    $InactivityMinutes = 120
    $LastActivity = [DateTime]::UtcNow

    try {
        while ($listener.IsListening) {
            # Auto-shutdown after inactivity
            if (([DateTime]::UtcNow - $LastActivity).TotalMinutes -gt $InactivityMinutes) {
                Write-Host "  No activity for $InactivityMinutes minutes — shutting down." -ForegroundColor Yellow
                break
            }
            $context = $listener.GetContext()
            $LastActivity = [DateTime]::UtcNow
            $req = $context.Request
            $res = $context.Response
            $url = $req.Url.AbsolutePath
            $method = $req.HttpMethod

            try {
                # ── OPTIONS (CORS preflight) ────────────────────────
                if ($method -eq 'OPTIONS') {
                    $res.AddHeader('Access-Control-Allow-Origin', '*')
                    $res.AddHeader('Access-Control-Allow-Methods', 'GET,POST,PATCH,OPTIONS')
                    $res.AddHeader('Access-Control-Allow-Headers', 'Content-Type')
                    $res.StatusCode = 204
                    $res.OutputStream.Close()
                    continue
                }

                # ── Serve HTML ──────────────────────────────────────
                if ($method -eq 'GET' -and ($url -eq '/' -or $url -eq '/index.html')) {
                    Send-HtmlResponse $res $HtmlFile
                    continue
                }

                # ── GET /api/changelog ──────────────────────────────
                if ($method -eq 'GET' -and $url -eq '/api/changelog') {
                    $data = @(Read-JsonFile $ClFile)
                    Send-JsonResponse $res 200 $data
                    continue
                }

                # ── GET /api/tickets ────────────────────────────────
                if ($method -eq 'GET' -and $url -eq '/api/tickets') {
                    $data = @(Read-JsonFile $TkFile)
                    Send-JsonResponse $res 200 $data
                    continue
                }

                # ── POST /api/changelog/add ─────────────────────────
                if ($method -eq 'POST' -and $url -eq '/api/changelog/add') {
                    $body = Read-RequestBody $req
                    $entry = $body | ConvertFrom-Json
                    if (-not $entry.date -or -not $entry.description) {
                        Send-JsonResponse $res 400 @{ error = 'Missing required fields (date, description).' }
                        continue
                    }
                    $updated = Invoke-LockedMutate $ClFile {
                        param($data)
                        $data += $entry
                        return $data
                    }
                    Send-JsonResponse $res 200 @{ ok = $true; count = $updated.Count; entries = @($updated) }
                    continue
                }

                # ── POST /api/tickets/add ───────────────────────────
                if ($method -eq 'POST' -and $url -eq '/api/tickets/add') {
                    $body = Read-RequestBody $req
                    $ticket = $body | ConvertFrom-Json
                    if (-not $ticket.title) {
                        Send-JsonResponse $res 400 @{ error = 'Missing required field (title).' }
                        continue
                    }
                    $updated = Invoke-LockedMutate $TkFile {
                        param($data)
                        $data += $ticket
                        return $data
                    }
                    # Send email (non-blocking via job)
                    if ($EmailEnabled) {
                        Start-Job -ScriptBlock {
                            param($From, $To, $Smtp, $SmtpPort, $Ticket)
                            $subj = "New Macro Guide Ticket: $($Ticket.title)"
                            $body = "Title: $($Ticket.title)`nType: $($Ticket.type)`nPriority: $($Ticket.priority)`nFrom: $($Ticket.createdBy)`n`n$($Ticket.description)"
                            Send-MailMessage -From $From -To $To -Subject $subj -Body $body -SmtpServer $Smtp -Port $SmtpPort -ErrorAction SilentlyContinue
                        } -ArgumentList $EmailFrom, $EmailTo, $EmailSmtp, $EmailPort, $ticket | Out-Null
                    }
                    Send-JsonResponse $res 200 @{ ok = $true; count = $updated.Count; entries = @($updated) }
                    continue
                }

                # ── PATCH /api/tickets/{id}/status ──────────────────
                if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/status$') {
                    $id = [long]$Matches[1]
                    $body = Read-RequestBody $req
                    $parsed = $body | ConvertFrom-Json
                    $newStatus = $parsed.status
                    $validStatuses = @('open', 'in-progress', 'done', 'closed')
                    if ($newStatus -notin $validStatuses) {
                        Send-JsonResponse $res 400 @{ error = "Invalid status. Must be one of: $($validStatuses -join ', ')" }
                        continue
                    }
                    $found = $false
                    $updated = Invoke-LockedMutate $TkFile {
                        param($data)
                        foreach ($tk in $data) {
                            if ([string]$tk.id -eq [string]$id) {
                                $tk.status = $newStatus
                                Set-Variable -Name found -Value $true -Scope 2
                            }
                        }
                        return $data
                    }
                    if (-not $found) {
                        Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
                        continue
                    }
                    Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
                    continue
                }

                # ── PATCH /api/tickets/{id}/cancel ──────────────────
                if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/cancel$') {
                    $id = [long]$Matches[1]
                    $body = Read-RequestBody $req
                    $parsed = $body | ConvertFrom-Json
                    if (-not $parsed.reason -or [string]::IsNullOrWhiteSpace($parsed.reason)) {
                        Send-JsonResponse $res 400 @{ error = 'A cancellation reason is required.' }
                        continue
                    }
                    $found = $false
                    $updated = Invoke-LockedMutate $TkFile {
                        param($data)
                        foreach ($tk in $data) {
                            if ([string]$tk.id -eq [string]$id) {
                                $tk.status = 'canceled'
                                $tk | Add-Member -NotePropertyName cancelReason -NotePropertyValue $parsed.reason.Trim() -Force
                                $tk | Add-Member -NotePropertyName canceledBy -NotePropertyValue $(if ($parsed.canceledBy) { $parsed.canceledBy } else { '' }) -Force
                                $tk | Add-Member -NotePropertyName canceledDate -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd') -Force
                                Set-Variable -Name found -Value $true -Scope 2
                            }
                        }
                        return $data
                    }
                    if (-not $found) {
                        Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
                        continue
                    }
                    Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
                    continue
                }

                # ── PATCH /api/tickets/{id}/vote ────────────────────
                if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/vote$') {
                    $id = [long]$Matches[1]
                    $found = $false
                    $updated = Invoke-LockedMutate $TkFile {
                        param($data)
                        foreach ($tk in $data) {
                            if ([string]$tk.id -eq [string]$id) {
                                $cur = if ($tk.PSObject.Properties['votes']) { [int]$tk.votes } else { 0 }
                                $tk | Add-Member -NotePropertyName votes -NotePropertyValue ($cur + 1) -Force
                                Set-Variable -Name found -Value $true -Scope 2
                            }
                        }
                        return $data
                    }
                    if (-not $found) {
                        Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
                        continue
                    }
                    Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
                    continue
                }

                # ── GET /api/shutdown (localhost only) ───────────────
                if ($method -eq 'GET' -and $url -eq '/api/shutdown') {
                    Send-JsonResponse $res 200 @{ ok = $true; message = 'Server shutting down.' }
                    Write-Host "  Shutdown requested via browser." -ForegroundColor Yellow
                    $listener.Stop()
                    return
                }

                # ── 404 ─────────────────────────────────────────────
                Send-JsonResponse $res 404 @{ error = "Not found: $url" }

            } catch {
                Write-Host "  ERROR [$method $url]: $($_.Exception.Message)" -ForegroundColor Red
                try {
                    Send-JsonResponse $res 500 @{ error = $_.Exception.Message }
                } catch {}
            }
        }
    } finally {
        $listener.Stop()
        $listener.Close()
        Release-Lock $ClFile
        Release-Lock $TkFile
        Write-Host "  Server stopped." -ForegroundColor Yellow
    }
}

# ── Run ─────────────────────────────────────────────────────────────────
Start-Server
