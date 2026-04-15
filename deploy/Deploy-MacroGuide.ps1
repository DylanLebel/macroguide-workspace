<#
.SYNOPSIS
    Deploys MacroGuide files from C:\AllMacros to the shared Y: drive.

.DESCRIPTION
    Copies MacroGuide.html, macro-guide-server.ps1, and Open Macro Guide.vbs
    from the local source (C:\AllMacros) to:
        Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\

    Before overwriting, it backs up the current Y: copies to a timestamped
    subfolder so you can roll back if something goes wrong.

.NOTES
    Run by double-clicking Deploy-MacroGuide.bat (the wrapper)
    or directly:  powershell -ExecutionPolicy Bypass -File Deploy-MacroGuide.ps1
#>

$ErrorActionPreference = 'Stop'

$Source      = 'C:\AllMacros'
$DeploySrc   = 'C:\AllMacros\deploy'
$Target      = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide'
$BackupRoot  = Join-Path $Target ('_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))

# Files to deploy: local path + filename in target
$Files = @(
    @{ Local = Join-Path $Source    'MacroGuide.html';         Name = 'MacroGuide.html' },
    @{ Local = Join-Path $DeploySrc 'macro-guide-server.ps1';  Name = 'macro-guide-server.ps1' },
    @{ Local = Join-Path $DeploySrc 'Open Macro Guide.vbs';    Name = 'Open Macro Guide.vbs' }
)

Write-Host ""
Write-Host "  Macro Guide Deploy" -ForegroundColor Cyan
Write-Host "  ==================" -ForegroundColor Cyan
Write-Host "  Source : $Source  +  $DeploySrc"
Write-Host "  Target : $Target"
Write-Host ""

# 1. Sanity checks
if (-not (Test-Path $Target)) {
    Write-Host "  ERROR: Target folder not reachable." -ForegroundColor Red
    Write-Host "  Make sure the Y: drive is mapped, then try again." -ForegroundColor Yellow
    Write-Host "    $Target" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to close"
    exit 1
}

foreach ($f in $Files) {
    if (-not (Test-Path $f.Local)) {
        Write-Host "  ERROR: Source file missing: $($f.Local)" -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit 1
    }
}

# 2. Warn if server appears to be running (stale HTML until restarted)
$serverLive = $false
$livePort = 0
foreach ($p in 8123,8124,8125,8126) {
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:$p/" -UseBasicParsing -TimeoutSec 1
        if ($r.StatusCode -eq 200) { $serverLive = $true; $livePort = $p; break }
    } catch {}
}
if ($serverLive) {
    Write-Host "  NOTE: Macro Guide server is running on port $livePort." -ForegroundColor Yellow
    Write-Host "  Users will see the new version AFTER the server picks up new files." -ForegroundColor Yellow
    Write-Host "  Click 'Restart Server' in the admin panel, or close + re-open the guide." -ForegroundColor Yellow
    Write-Host ""
}

# 3. Back up current Y: copies (only those that already exist)
Write-Host "  Backing up current Y: copies..." -ForegroundColor Gray
$backupNeeded = $false
foreach ($f in $Files) {
    $existing = Join-Path $Target $f.Name
    if (Test-Path $existing) {
        if (-not $backupNeeded) {
            New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
            $backupNeeded = $true
        }
        Copy-Item $existing (Join-Path $BackupRoot $f.Name) -Force
        Write-Host "    backed up: $($f.Name)" -ForegroundColor DarkGray
    }
}
if ($backupNeeded) {
    Write-Host "  Backup folder: $BackupRoot" -ForegroundColor DarkGray
} else {
    Write-Host "  (nothing to back up - first deploy)" -ForegroundColor DarkGray
}
Write-Host ""

# 4. Copy files
Write-Host "  Copying files..." -ForegroundColor Gray
$copied = 0
$failed = @()
foreach ($f in $Files) {
    $dest = Join-Path $Target $f.Name
    try {
        Copy-Item $f.Local $dest -Force
        $size = (Get-Item $dest).Length
        Write-Host ("    OK  {0,-30}  {1,8:N0} bytes" -f $f.Name, $size) -ForegroundColor Green
        $copied++
    } catch {
        Write-Host ("    FAIL {0}  ({1})" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
        $failed += $f.Name
    }
}

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "  Deploy complete: $copied file(s) updated." -ForegroundColor Green
    if ($serverLive) {
        Write-Host "  Remember to restart the server so users see the new HTML." -ForegroundColor Yellow
    }
} else {
    Write-Host "  Deploy finished with errors. Failed: $($failed -join ', ')" -ForegroundColor Red
    Write-Host "  Original files were backed up to:" -ForegroundColor Yellow
    Write-Host "    $BackupRoot" -ForegroundColor DarkGray
}
Write-Host ""
Read-Host "Press Enter to close"
