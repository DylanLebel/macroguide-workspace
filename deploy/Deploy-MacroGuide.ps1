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
$PdmTarget   = 'C:\NMT_PDM\Libraries\Macro'   # launcher also goes here so the team can open the guide from PDM
$IconName    = 'NMTMacroGuide.ico'                              # custom icon, matches MacroGuide.html aesthetic
$IconSrc     = Join-Path $DeploySrc $IconName                   # built by Build-MacroGuideIcon.ps1
$ShortcutName = 'NMT Macro Guide.lnk'
$BackupRoot  = Join-Path $Target ('_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))

# Create a .lnk that opens the .vbs via wscript.exe with a custom icon. We
# target wscript explicitly (rather than the .vbs directly) so the shortcut
# is bulletproof even if a user has cscript registered as the .vbs handler -
# cscript would pop up a console window. Returns $true on success.
function New-MacroGuideShortcut {
    param(
        [Parameter(Mandatory)] [string]$LnkPath,
        [Parameter(Mandatory)] [string]$VbsPath,
        [string]$IconPath = ''
    )
    try {
        # Remove existing first; some Windows builds refuse to overwrite a .lnk
        # in-place if its target was deleted/changed.
        if (Test-Path $LnkPath) { Remove-Item $LnkPath -Force -ErrorAction SilentlyContinue }
        $wsh = New-Object -ComObject WScript.Shell
        $lnk = $wsh.CreateShortcut($LnkPath)
        $lnk.TargetPath       = "$env:SystemRoot\System32\wscript.exe"
        $lnk.Arguments        = "`"$VbsPath`""
        $lnk.WorkingDirectory = Split-Path -Parent $VbsPath
        $lnk.Description      = 'Open the NMT Macro Guide'
        if ($IconPath -and (Test-Path $IconPath)) {
            $lnk.IconLocation = "$IconPath,0"
        }
        $lnk.Save()
        return $true
    } catch {
        Write-Host "    Could not build shortcut at ${LnkPath}: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

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

# 5. Remove the old "Open Macro Guide.vbs" from PDM if it's still there.
#    Users now click "NMT Macro Guide.lnk" (added below) which runs the .vbs
#    from Y: - no local PDM copy needed, and fewer scary-looking files in the
#    PDM folder.
Write-Host "  Checking for old Open Macro Guide.vbs in PDM..." -ForegroundColor Gray
$oldPdmVbs = Join-Path $PdmTarget 'Open Macro Guide.vbs'
if (-not (Test-Path $PdmTarget)) {
    Write-Host "    SKIP  PDM folder not found: $PdmTarget" -ForegroundColor DarkGray
} elseif (-not (Test-Path $oldPdmVbs)) {
    Write-Host "    (not present - nothing to remove)" -ForegroundColor DarkGray
} else {
    $oldVbs = Get-Item $oldPdmVbs
    if ($oldVbs.Attributes -band [IO.FileAttributes]::ReadOnly) {
        Write-Host "    SKIP  Old Open Macro Guide.vbs in PDM is READ-ONLY (checked in)." -ForegroundColor Yellow
        Write-Host "          Check it out in PDM, then re-run Deploy-MacroGuide.bat." -ForegroundColor Yellow
        $failed += 'Open Macro Guide.vbs (PDM read-only, cannot remove)'
    } else {
        try {
            # Back up before deleting - same backup folder as other PDM items.
            if (-not (Test-Path $BackupRoot)) { New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null }
            Copy-Item $oldPdmVbs (Join-Path $BackupRoot 'Open Macro Guide.vbs.pdm.removed.bak') -Force
            Remove-Item $oldPdmVbs -Force
            Write-Host "    OK  Removed old Open Macro Guide.vbs from PDM." -ForegroundColor Green
            Write-Host "    REMINDER: Check this deletion in via PDM so it's removed for everyone." -ForegroundColor Yellow
        } catch {
            Write-Host ("    FAIL  Could not remove old .vbs from PDM ({0})" -f $_.Exception.Message) -ForegroundColor Red
            $failed += 'Open Macro Guide.vbs (PDM remove)'
        }
    }
}
Write-Host ""

# 6. Build the user-facing "NMT Macro Guide" shortcut in both locations.
#    Users see the shortcut (with a proper icon) instead of the scary .vbs.
#    The .vbs stays as the engine - the .lnk just calls it via wscript.
Write-Host "  Building NMT Macro Guide shortcut..." -ForegroundColor Gray

# 6a. Auto-build the icon if it's missing. Build-MacroGuideIcon.ps1 is the
#     source of truth - running it here means a fresh checkout / new machine
#     gets the right icon without an extra manual step.
if (-not (Test-Path $IconSrc)) {
    $iconBuilder = Join-Path $DeploySrc 'Build-MacroGuideIcon.ps1'
    if (Test-Path $iconBuilder) {
        Write-Host "    Icon missing - running Build-MacroGuideIcon.ps1..." -ForegroundColor DarkGray
        try {
            & powershell.exe -ExecutionPolicy Bypass -File $iconBuilder | Out-Null
        } catch {
            Write-Host ("    WARN  Icon build failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# 6b. Copy the icon to Y: so the Y: shortcut is self-contained. We deploy
#     under the same name we use locally so it's easy to swap by hand later.
$YIcon = Join-Path $Target $IconName
if (Test-Path $IconSrc) {
    try {
        Copy-Item $IconSrc $YIcon -Force
        Write-Host ("    OK  {0,-30}  copied to Y:" -f $IconName) -ForegroundColor DarkGray
    } catch {
        Write-Host ("    WARN  Could not copy icon to Y: ({0}) - Y: shortcut will use generic icon." -f $_.Exception.Message) -ForegroundColor Yellow
        $YIcon = ''
    }
} else {
    Write-Host "    WARN  Icon source not found: $IconSrc" -ForegroundColor Yellow
    Write-Host "          Run Build-MacroGuideIcon.ps1, then re-run this deploy." -ForegroundColor DarkGray
    $YIcon = ''
}

# 6c. Y: shortcut - points at the .vbs on Y:, uses icon on Y:
$YShortcut = Join-Path $Target $ShortcutName
$YVbs      = Join-Path $Target 'Open Macro Guide.vbs'
if (New-MacroGuideShortcut -LnkPath $YShortcut -VbsPath $YVbs -IconPath $YIcon) {
    Write-Host ("    OK  {0,-30}  -> {1}" -f $ShortcutName, $Target) -ForegroundColor Green
} else {
    $failed += "$ShortcutName (Y:)"
}

# 6d. PDM shortcut - points at the .vbs on Y: AND the icon on Y:.
#     IMPORTANT: Do NOT copy the icon into PDM. PDM files are vaulted by
#     default - users only get a local copy when they explicitly "Get Latest"
#     on that file. Without a local copy, the .lnk's IconLocation can't
#     resolve and Explorer falls back to a generic icon. Pointing IconLocation
#     at Y: avoids this entirely - Y: is mapped on every workstation that
#     uses macros anyway, so the icon renders immediately in PDM's folder
#     view with no per-user setup.
if (Test-Path $PdmTarget) {
    $PdmShortcutIcon = $YIcon

    $PdmShortcut = Join-Path $PdmTarget $ShortcutName
    $PdmVbs      = Join-Path $Target 'Open Macro Guide.vbs'   # Y: target, not PDM
    $skipPdm = $false
    if (Test-Path $PdmShortcut) {
        $existing = Get-Item $PdmShortcut
        if ($existing.Attributes -band [IO.FileAttributes]::ReadOnly) {
            Write-Host "    SKIP  $ShortcutName in PDM is READ-ONLY (checked in)." -ForegroundColor Yellow
            Write-Host "          Check it out in PDM, then re-run Deploy-MacroGuide.bat." -ForegroundColor Yellow
            $failed += "$ShortcutName (PDM read-only)"
            $skipPdm = $true
        }
    }
    if (-not $skipPdm) {
        # Back up existing PDM shortcut if one exists
        if (Test-Path $PdmShortcut) {
            if (-not (Test-Path $BackupRoot)) { New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null }
            Copy-Item $PdmShortcut (Join-Path $BackupRoot "$ShortcutName.pdm.bak") -Force -ErrorAction SilentlyContinue
        }
        if (New-MacroGuideShortcut -LnkPath $PdmShortcut -VbsPath $PdmVbs -IconPath $PdmShortcutIcon) {
            Write-Host ("    OK  {0,-30}  -> {1}" -f $ShortcutName, $PdmTarget) -ForegroundColor Green
            Write-Host "          (shortcut target: $PdmVbs)" -ForegroundColor DarkGray
            Write-Host "    REMINDER: Check this shortcut in via PDM so the rest of the team gets it." -ForegroundColor Yellow
        } else {
            $failed += "$ShortcutName (PDM)"
        }
    }
}
Write-Host ""

if ($failed.Count -eq 0) {
    Write-Host "  Deploy complete: $copied file(s) updated on Y:, launcher + shortcut dropped in PDM." -ForegroundColor Green
    Write-Host "  Users should click '$ShortcutName' (nice icon) instead of the .vbs." -ForegroundColor Cyan
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
