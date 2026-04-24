#Requires -Version 5.1
<#
.SYNOPSIS
    Builds a local "NMT Macro Guide (preview).lnk" so the new .ico can be
    inspected in Windows Explorer before the production deploy.

.DESCRIPTION
    Drops the .lnk in C:\AllMacros\deploy\. Target points at the local
    Open Macro Guide.vbs so the shortcut is also functional — double-click
    to confirm the launcher fires correctly with the new icon.

    Production shortcuts (Y: + PDM) are built later by Deploy-MacroGuide.ps1.
#>

$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LnkPath   = Join-Path $ScriptDir 'NMT Macro Guide (preview).lnk'
$VbsPath   = Join-Path $ScriptDir 'Open Macro Guide.vbs'
$IconPath  = Join-Path $ScriptDir 'NMTMacroGuide.ico'

if (-not (Test-Path $VbsPath))  { throw "Missing: $VbsPath" }
if (-not (Test-Path $IconPath)) { throw "Missing: $IconPath - run Build-MacroGuideIcon.ps1 first." }

# Some Windows builds refuse to overwrite a .lnk if its target was deleted
# or moved, so remove first.
if (Test-Path $LnkPath) { Remove-Item $LnkPath -Force }

$wsh = New-Object -ComObject WScript.Shell
$lnk = $wsh.CreateShortcut($LnkPath)
$lnk.TargetPath       = "$env:SystemRoot\System32\wscript.exe"
$lnk.Arguments        = "`"$VbsPath`""
$lnk.WorkingDirectory = $ScriptDir
$lnk.Description      = 'Open the NMT Macro Guide (preview build)'
$lnk.IconLocation     = "$IconPath,0"
$lnk.Save()

Write-Host ""
Write-Host "  Preview shortcut written:" -ForegroundColor Green
Write-Host "    $LnkPath" -ForegroundColor DarkGray
Write-Host "  Open the deploy folder in Explorer to see the icon." -ForegroundColor Cyan
Write-Host ""
