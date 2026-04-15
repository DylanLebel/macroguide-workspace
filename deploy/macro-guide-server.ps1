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
Add-Type -AssemblyName System.Web

# ════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ════════════════════════════════════════════════════════════════════════

$PortCandidates = @(8123, 8124, 8125, 8126)
$Port = $PortCandidates[0]  # actual chosen port set later when listener starts
$PortStatusFile = Join-Path $env:TEMP 'MacroGuide.port'

# Paths — auto-detect relative to script, fall back to known locations
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$DataDir = Join-Path $ScriptDir 'data'
if (-not (Test-Path $DataDir)) {
    $DataDir = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\data'
}
# Local fallback if shared data dir is unavailable (e.g. Y: not mapped).
# Initialize-DataFiles will switch to this if the primary fails.
$LocalDataDir = Join-Path $env:APPDATA 'MacroGuide\data'
$DataSource = 'shared'  # 'shared' or 'local' — set by Initialize-DataFiles

$HtmlFile = Join-Path $ScriptDir 'MacroGuide.html'
if (-not (Test-Path $HtmlFile)) {
    $HtmlFile = 'C:\AllMacros\MacroGuide.html'
}

$ClFile = Join-Path $DataDir 'changelog.json'
$TkFile = Join-Path $DataDir 'tickets.json'



# Email notifications — uses Outlook COM (no credentials needed, works with logged-in Outlook)
$EmailEnabled = $true
$EmailTo      = 'dlebel@nmtech.com'

# ════════════════════════════════════════════════════════════════════════
#  ERROR LOG (shared, so Dylan can see what's failing on any user's box)
# ════════════════════════════════════════════════════════════════════════

# Primary log on Y: so Dylan can see all users' errors in one place.
# Falls back to APPDATA when Y: isn't reachable (same policy as data dir).
$SharedErrorLog = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\data\errors.log'
$LocalErrorLog  = Join-Path $env:APPDATA 'MacroGuide\errors.log'

function Write-ErrorLog {
    param([string]$Context, [string]$Message, [string]$Detail = '')
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $user = $env:USERNAME
    $machine = $env:COMPUTERNAME
    $line = "[$ts] [$user@$machine] [$Context] $Message"
    if ($Detail) { $line += " | $Detail" }
    # Try shared first; fall back to local if Y: isn't writable.
    $target = $SharedErrorLog
    try {
        $dir = Split-Path -Parent $target
        if (-not (Test-Path $dir)) { throw "shared dir missing" }
        Add-Content -Path $target -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        try {
            $localDir = Split-Path -Parent $LocalErrorLog
            if (-not (Test-Path $localDir)) { New-Item -ItemType Directory -Path $localDir -Force | Out-Null }
            Add-Content -Path $LocalErrorLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

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

function Send-OutlookEmail {
    param([string]$Subject, [string]$HtmlBody, [string]$To = $EmailTo)
    if (-not $EmailEnabled) { return }
    try {
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $Subject
        $mail.HTMLBody = $HtmlBody
        $mail.Send()
        Write-Host "  Email sent to ${To}: $Subject" -ForegroundColor Green
    } catch {
        Write-Host "  Email failed ($To): $($_.Exception.Message)" -ForegroundColor Yellow
        # Most common cause: Outlook isn't running on this user's machine.
        # Log so Dylan can see who got missed and manually follow up.
        Write-ErrorLog 'email' "Outlook send failed to $To" "Subject: $Subject | $($_.Exception.Message)"
    }
}

function Get-EmailHtmlWrapper {
    param([string]$Title, [string]$InnerHtml)
    return @"
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if mso]>
<xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml>
<style>body,table,td,p,a,h2{font-family:'Segoe UI',Arial,sans-serif!important;}td{mso-line-height-rule:exactly;}</style>
<![endif]-->
</head>
<body style="margin:0;padding:0;background-color:#f0f0f0;font-family:'Segoe UI',-apple-system,'Helvetica Neue',Arial,sans-serif;line-height:1.6;-webkit-font-smoothing:antialiased;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f0f0f0;">
<tr><td align="center" style="padding:40px 16px;">
<!--[if mso]><table role="presentation" width="540" cellpadding="0" cellspacing="0" border="0" align="center" style="border:1px solid #e0e0e0;"><tr><td><![endif]-->
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:540px;background-color:#ffffff;border:1px solid #e5e5e5;">
  <tr><td style="height:4px;background-color:#4f8ef7;font-size:1px;line-height:1px;">&nbsp;</td></tr>
  <tr><td style="padding:28px 36px 0 36px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr><td style="font-size:11px;text-transform:uppercase;letter-spacing:2px;color:#b0b0b0;font-weight:600;font-family:'Segoe UI',Arial,sans-serif;padding-bottom:20px;">NORDIC MINESTEEL &bull; MACRO GUIDE</td></tr>
    <tr><td style="font-size:20px;font-weight:600;color:#1a1a1a;font-family:'Segoe UI',Arial,sans-serif;line-height:28px;padding-bottom:24px;">$Title</td></tr>
    </table>
  </td></tr>
  <tr><td style="padding:0 36px 28px 36px;">
    $InnerHtml
  </td></tr>
  <tr><td style="border-top:1px solid #eeeeee;padding:16px 36px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr><td style="font-size:11px;color:#c0c0c0;font-family:'Segoe UI',Arial,sans-serif;">Automated notification from Macro Guide</td></tr>
    </table>
  </td></tr>
</table>
<!--[if mso]></td></tr></table><![endif]-->
</td></tr></table>
</body></html>
"@
}

function Get-EmailRow {
    param([string]$Label, [string]$Value, [string]$ValueStyle = '')
    $vs = if ($ValueStyle) { $ValueStyle } else { 'color:#1a1a1a;font-size:14px;font-family:''Segoe UI'',Arial,sans-serif;' }
    return "<tr><td width=`"110`" style=`"padding:11px 12px 11px 0;border-bottom:1px solid #f0f0f0;color:#999999;font-size:13px;font-family:'Segoe UI',Arial,sans-serif;vertical-align:top;`">$Label</td><td style=`"padding:11px 0;border-bottom:1px solid #f0f0f0;$vs`">$Value</td></tr>"
}

function Get-EmailBadge {
    param([string]$Text, [string]$BgColor, [string]$TextColor)
    return @"
<!--[if mso]><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:$BgColor;padding:4px 14px;"><span style="color:$TextColor;font-size:12px;font-weight:600;font-family:'Segoe UI',Arial,sans-serif;">$Text</span></td></tr></table><![endif]-->
<!--[if !mso]><!--><span style="display:inline-block;padding:4px 14px;background-color:$BgColor;color:$TextColor;font-size:12px;font-weight:600;font-family:'Segoe UI',Arial,sans-serif;border-radius:20px;line-height:1.4;">$Text</span><!--<![endif]-->
"@
}

function Get-EmailDescBlock {
    param([string]$Text, [string]$AccentColor = '#4f8ef7')
    return @"
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">
<tr>
<td width="3" style="background-color:$AccentColor;font-size:1px;width:3px;">&nbsp;</td>
<td style="background-color:#fafafa;padding:14px 18px;font-size:14px;color:#444444;font-family:'Segoe UI',Arial,sans-serif;line-height:22px;">$Text</td>
</tr></table>
"@
}

function Get-EmailLink {
    param([string]$Text = 'Open Macro Guide')
    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;"><tr>
<td style="font-size:13px;color:#666666;font-family:'Segoe UI',Arial,sans-serif;padding:10px 14px;background-color:#f5f5f5;border-left:3px solid #4f8ef7;">
<strong style="color:#333333;">$Text</strong><br/>
Run <em>Open Macro Guide.vbs</em> from your NMT_PDM shortcut or find it at:<br/>
<span style="font-family:Consolas,'Courier New',monospace;font-size:12px;color:#4f8ef7;">Y:\Solidworks\Macros\Macro Data PDM\MacroGuide</span>
</td></tr></table>
"@
}

function Get-EmailSectionLabel {
    param([string]$Text)
    return "<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" border=`"0`" style=`"margin-bottom:10px;`"><tr><td style=`"font-size:11px;text-transform:uppercase;letter-spacing:1.5px;color:#999999;font-weight:600;font-family:'Segoe UI',Arial,sans-serif;padding-bottom:2px;`">$Text</td></tr></table>"
}

function Send-TicketEmail {
    param($Ticket)
    $subject = "New Macro Guide Request: $($Ticket.title)"
    $from = if ($Ticket.createdBy) { $Ticket.createdBy } else { 'Anonymous' }
    $desc = if ($Ticket.description) { [System.Web.HttpUtility]::HtmlEncode($Ticket.description) } else { 'No description provided' }
    $tTitle = [System.Web.HttpUtility]::HtmlEncode($Ticket.title)
    $tType = [System.Web.HttpUtility]::HtmlEncode($Ticket.type)
    $tPrio = [System.Web.HttpUtility]::HtmlEncode($Ticket.priority)
    $tMacro = [System.Web.HttpUtility]::HtmlEncode($Ticket.macroName)
    $tFrom = [System.Web.HttpUtility]::HtmlEncode($from)
    $tDate = [System.Web.HttpUtility]::HtmlEncode($Ticket.date)
    $rows = @(
        (Get-EmailRow 'Title' "<strong>$tTitle</strong>")
        (Get-EmailRow 'Type' $tType)
        (Get-EmailRow 'Priority' $tPrio)
        (Get-EmailRow 'Macro' $tMacro)
        (Get-EmailRow 'From' $tFrom)
        (Get-EmailRow 'Date' $tDate)
    ) -join "`n"
    $descBlock = Get-EmailDescBlock $desc '#4f8ef7'
    $link = Get-EmailLink
    $detailLabel = Get-EmailSectionLabel 'Details'
    $descLabel = Get-EmailSectionLabel 'Description'
    $inner = @"
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="font-size:14px;color:#666666;font-family:'Segoe UI',Arial,sans-serif;line-height:22px;padding-bottom:20px;">A new request was submitted in the Macro Guide.</td></tr></table>
    $detailLabel
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">$rows</table>
    $descLabel
    $descBlock
    $link
"@
    $htmlBody = Get-EmailHtmlWrapper -Title "New Request Submitted" -InnerHtml $inner
    Send-OutlookEmail -Subject $subject -HtmlBody $htmlBody
}

function Send-VoteEmail {
    param($Ticket, [int]$NewVoteCount)
    $subject = "Upvote on: $($Ticket.title) ($NewVoteCount votes)"
    $from = if ($Ticket.createdBy) { $Ticket.createdBy } else { 'Anonymous' }
    $tTitle = [System.Web.HttpUtility]::HtmlEncode($Ticket.title)
    $tStatus = [System.Web.HttpUtility]::HtmlEncode($Ticket.status)
    $tPrio = [System.Web.HttpUtility]::HtmlEncode($Ticket.priority)
    $tMacro = [System.Web.HttpUtility]::HtmlEncode($Ticket.macroName)
    $tFrom = [System.Web.HttpUtility]::HtmlEncode($from)
    $voteBadge = Get-EmailBadge $NewVoteCount '#eef4ff' '#4f8ef7'
    $rows = @(
        (Get-EmailRow 'Title' "<strong>$tTitle</strong>")
        (Get-EmailRow 'Votes' $voteBadge)
        (Get-EmailRow 'Status' $tStatus)
        (Get-EmailRow 'Priority' $tPrio)
        (Get-EmailRow 'Macro' $tMacro)
        (Get-EmailRow 'From' $tFrom)
    ) -join "`n"
    $link = Get-EmailLink
    $detailLabel = Get-EmailSectionLabel 'Details'
    $inner = @"
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="font-size:14px;color:#666666;font-family:'Segoe UI',Arial,sans-serif;line-height:22px;padding-bottom:20px;">A request was upvoted in the Macro Guide.</td></tr></table>
    $detailLabel
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">$rows</table>
    $link
"@
    $htmlBody = Get-EmailHtmlWrapper -Title "Request Upvoted" -InnerHtml $inner
    Send-OutlookEmail -Subject $subject -HtmlBody $htmlBody
}

function Send-StatusNotification {
    param($Ticket, [string]$NewStatus)
    # Collect recipient emails: creator + voters
    $recipients = @()
    $creatorUser = if ($Ticket.PSObject.Properties['windowsUser'] -and $Ticket.windowsUser) { $Ticket.windowsUser } else { $null }
    if ($creatorUser) { $recipients += "$creatorUser@nmtech.com" }
    if ($Ticket.PSObject.Properties['voters'] -and $Ticket.voters) {
        foreach ($voter in $Ticket.voters) { if ($voter) { $recipients += "$voter@nmtech.com" } }
    }
    $recipients = @($recipients | Select-Object -Unique)
    if ($recipients.Count -eq 0) {
        Write-Host "  No recipients for status notification (no windowsUser or voters) — skipping." -ForegroundColor DarkGray
        return
    }
    $recipientList = $recipients -join '; '
    $title = $Ticket.title
    if ($NewStatus -eq 'in-progress') {
        $subject = "Your request is being worked on: $title"
        $statusLabel = 'In Progress'
        $badgeBg = '#fef3e2'; $badgeColor = '#b8760a'
        $msg = "Dylan has started working on your request. You'll be notified again when it's complete."
    } elseif ($NewStatus -eq 'done') {
        $subject = "Your request is complete: $title"
        $statusLabel = 'Done'
        $badgeBg = '#e8f9ee'; $badgeColor = '#1a8a3a'
        $msg = "Your request has been completed. If you have questions or need follow-up, submit a new request or reach out to Dylan."
    } else { return }
    $tTitle = [System.Web.HttpUtility]::HtmlEncode($title)
    $tMacro = [System.Web.HttpUtility]::HtmlEncode($Ticket.macroName)
    $statusBadge = Get-EmailBadge $statusLabel $badgeBg $badgeColor
    $rows = @(
        (Get-EmailRow 'Title' "<strong>$tTitle</strong>")
        (Get-EmailRow 'Macro' $tMacro)
        (Get-EmailRow 'Status' $statusBadge)
    ) -join "`n"
    $link = Get-EmailLink
    $detailLabel = Get-EmailSectionLabel 'Details'
    $inner = @"
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="font-size:14px;color:#666666;font-family:'Segoe UI',Arial,sans-serif;line-height:22px;padding-bottom:20px;">$msg</td></tr></table>
    $detailLabel
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">$rows</table>
    $link
"@
    $htmlBody = Get-EmailHtmlWrapper -Title $subject -InnerHtml $inner
    Send-OutlookEmail -Subject $subject -HtmlBody $htmlBody -To $recipientList
}

function Send-PokeEmail {
    param($Ticket)
    $pokes = if ($Ticket.PSObject.Properties['pokes']) { [int]$Ticket.pokes } else { 1 }
    $subject = "POKE ($pokes): $($Ticket.title)"
    $from = if ($Ticket.createdBy) { $Ticket.createdBy } else { 'Anonymous' }
    $desc = if ($Ticket.description) { [System.Web.HttpUtility]::HtmlEncode($Ticket.description) } else { 'No description provided' }
    $tTitle = [System.Web.HttpUtility]::HtmlEncode($Ticket.title)
    $tStatus = [System.Web.HttpUtility]::HtmlEncode($Ticket.status)
    $tPrio = [System.Web.HttpUtility]::HtmlEncode($Ticket.priority)
    $tMacro = [System.Web.HttpUtility]::HtmlEncode($Ticket.macroName)
    $tFrom = [System.Web.HttpUtility]::HtmlEncode($from)
    $pokeBadge = Get-EmailBadge $pokes '#fef0f0' '#c0392b'
    $rows = @(
        (Get-EmailRow 'Title' "<strong>$tTitle</strong>")
        (Get-EmailRow 'Pokes' $pokeBadge)
        (Get-EmailRow 'Status' $tStatus)
        (Get-EmailRow 'Priority' $tPrio)
        (Get-EmailRow 'Macro' $tMacro)
        (Get-EmailRow 'From' $tFrom)
    ) -join "`n"
    $descBlock = Get-EmailDescBlock $desc '#c0392b'
    $link = Get-EmailLink
    $detailLabel = Get-EmailSectionLabel 'Details'
    $descLabel = Get-EmailSectionLabel 'Description'
    $inner = @"
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="font-size:14px;color:#666666;font-family:'Segoe UI',Arial,sans-serif;line-height:22px;padding-bottom:20px;">Someone poked a request in the Macro Guide &mdash; they're waiting on this.</td></tr></table>
    $detailLabel
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">$rows</table>
    $descLabel
    $descBlock
    $link
"@
    $htmlBody = Get-EmailHtmlWrapper -Title "Poke Reminder" -InnerHtml $inner
    Send-OutlookEmail -Subject $subject -HtmlBody $htmlBody
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
            # Shared drive unavailable — fall back to local APPDATA so the guide
            # still works read-only for this user. Their tickets/changelog edits
            # will NOT be visible to the rest of the team until Y: is restored.
            Write-Host "  WARNING: Shared data dir unavailable ($DataDir)." -ForegroundColor Yellow
            Write-Host "  Falling back to local-only mode: $LocalDataDir" -ForegroundColor Yellow
            Write-Host "  Tickets and changelog edits will NOT sync to the team." -ForegroundColor Yellow
            try {
                New-Item -ItemType Directory -Path $LocalDataDir -Force | Out-Null
            } catch {
                throw "Could not create local fallback data directory either:`n  $LocalDataDir"
            }
            $script:DataDir = $LocalDataDir
            $script:ClFile  = Join-Path $LocalDataDir 'changelog.json'
            $script:TkFile  = Join-Path $LocalDataDir 'tickets.json'
            $script:DataSource = 'local'
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
    $attachDir = Join-Path $DataDir 'attachments'
    if (-not (Test-Path $attachDir)) {
        New-Item -ItemType Directory -Path $attachDir -Force | Out-Null
        Write-Host "  Created attachments directory"
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
    $Response.AddHeader('Access-Control-Allow-Methods', 'GET,POST,PATCH,DELETE,OPTIONS')
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
#  REQUEST HANDLER (extracted to avoid PS 5.1 nested try/catch bug)
# ════════════════════════════════════════════════════════════════════════

function Handle-Request {
    param($req, $res)
    $url = $req.Url.AbsolutePath
    $method = $req.HttpMethod

    # ── OPTIONS (CORS preflight) ────────────────────────
    if ($method -eq 'OPTIONS') {
        $res.AddHeader('Access-Control-Allow-Origin', '*')
        $res.AddHeader('Access-Control-Allow-Methods', 'GET,POST,PATCH,DELETE,OPTIONS')
        $res.AddHeader('Access-Control-Allow-Headers', 'Content-Type')
        $res.StatusCode = 204
        $res.OutputStream.Close()
        return
    }

    # ── Serve HTML ──────────────────────────────────────
    if ($method -eq 'GET' -and ($url -eq '/' -or $url -eq '/index.html')) {
        Send-HtmlResponse $res $HtmlFile
        return
    }

    # ── GET /api/changelog ──────────────────────────────
    if ($method -eq 'GET' -and $url -eq '/api/changelog') {
        $data = @(Read-JsonFile $ClFile)
        Send-JsonResponse $res 200 $data
        return
    }

    # ── GET /api/tickets ────────────────────────────────
    if ($method -eq 'GET' -and $url -eq '/api/tickets') {
        $data = @(Read-JsonFile $TkFile)
        Send-JsonResponse $res 200 $data
        return
    }

    # ── POST /api/changelog/add ─────────────────────────
    if ($method -eq 'POST' -and $url -eq '/api/changelog/add') {
        $body = Read-RequestBody $req
        $entry = $body | ConvertFrom-Json
        if (-not $entry.date -or -not $entry.description) {
            Send-JsonResponse $res 400 @{ error = 'Missing required fields (date, description).' }
            return
        }
        $updated = Invoke-LockedMutate $ClFile {
            param($data)
            $maxId = 0
            foreach ($e in $data) {
                if ($e.PSObject.Properties['id'] -and [int]$e.id -gt $maxId) { $maxId = [int]$e.id }
            }
            $entry | Add-Member -NotePropertyName id -NotePropertyValue ($maxId + 1) -Force
            $data += $entry
            return $data
        }
        Send-JsonResponse $res 200 @{ ok = $true; count = $updated.Count; entries = @($updated) }
        return
    }

    # ── POST /api/tickets/add ───────────────────────────
    if ($method -eq 'POST' -and $url -eq '/api/tickets/add') {
        $body = Read-RequestBody $req
        $ticket = $body | ConvertFrom-Json
        if (-not $ticket.title) {
            Send-JsonResponse $res 400 @{ error = 'Missing required field (title).' }
            return
        }
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            $data += $ticket
            return $data
        }
        # Send email via Outlook COM
        Send-TicketEmail $ticket
        Send-JsonResponse $res 200 @{ ok = $true; count = $updated.Count; entries = @($updated) }
        return
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
            return
        }
        $found = $false
        $notifyTicket = $null
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    $tk.status = $newStatus
                    Set-Variable -Name found -Value $true -Scope 2
                    if ($newStatus -eq 'in-progress' -or $newStatus -eq 'done') {
                        Set-Variable -Name notifyTicket -Value $tk -Scope 2
                    }
                }
            }
            return $data
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        # Email creator + voters when status changes to in-progress or done
        if ($notifyTicket) {
            Send-StatusNotification $notifyTicket $newStatus
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/cancel ──────────────────
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/cancel$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        if (-not $parsed.reason -or [string]::IsNullOrWhiteSpace($parsed.reason)) {
            Send-JsonResponse $res 400 @{ error = 'A cancellation reason is required.' }
            return
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
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/vote ────────────────────
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/vote$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $null
        $voterUser = $null
        if ($body -and $body.Trim().Length -gt 0) {
            try { $parsed = $body | ConvertFrom-Json } catch {}
        }
        if ($parsed -and $parsed.PSObject.Properties['windowsUser'] -and $parsed.windowsUser) {
            $voterUser = $parsed.windowsUser
        }
        $found = $false
        $completed = $false
        $votedTicket = $null
        $newVoteCount = 0
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    # Block votes on completed tickets
                    if ($tk.status -in @('done', 'closed', 'canceled')) {
                        Set-Variable -Name completed -Value $true -Scope 2
                        return $data
                    }
                    $cur = if ($tk.PSObject.Properties['votes']) { [int]$tk.votes } else { 0 }
                    $tk | Add-Member -NotePropertyName votes -NotePropertyValue ($cur + 1) -Force
                    # Track voter in voters array
                    if ($voterUser) {
                        $existingVoters = @()
                        if ($tk.PSObject.Properties['voters'] -and $tk.voters) {
                            $existingVoters = @($tk.voters)
                        }
                        if ($voterUser -notin $existingVoters) {
                            $existingVoters += $voterUser
                        }
                        $tk | Add-Member -NotePropertyName voters -NotePropertyValue $existingVoters -Force
                    }
                    # Auto-escalate priority based on vote count
                    $newCount = $cur + 1
                    $prioOrder = @{ 'low' = 0; 'medium' = 1; 'high' = 2; 'critical' = 3 }
                    $newPrio = if ($newCount -ge 4) { 'critical' } elseif ($newCount -ge 3) { 'high' } elseif ($newCount -ge 2) { 'medium' } else { $null }
                    if ($newPrio -and $prioOrder[$newPrio] -gt $prioOrder[$tk.priority]) {
                        $tk.priority = $newPrio
                    }
                    Set-Variable -Name votedTicket -Value $tk -Scope 2
                    Set-Variable -Name newVoteCount -Value $newCount -Scope 2
                }
            }
            return $data
        }
        if ($completed) {
            Send-JsonResponse $res 400 @{ error = 'Cannot vote on a completed request.' }
            return
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        # Email on upvote
        if ($votedTicket) {
            Send-VoteEmail $votedTicket $newVoteCount
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/poke ─────────────────────
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/poke$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        $pokerUser = if ($parsed.PSObject.Properties['windowsUser'] -and $parsed.windowsUser) { $parsed.windowsUser } else { '' }
        $found = $false
        $completed = $false
        $pokedTicket = $null
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    # Block pokes on completed tickets
                    if ($tk.status -in @('done', 'closed', 'canceled')) {
                        Set-Variable -Name completed -Value $true -Scope 2
                        return $data
                    }
                    $cur = if ($tk.PSObject.Properties['pokes']) { [int]$tk.pokes } else { 0 }
                    $tk | Add-Member -NotePropertyName pokes -NotePropertyValue ($cur + 1) -Force
                    $tk | Add-Member -NotePropertyName lastPoke -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd') -Force
                    # Track who poked
                    if ($pokerUser) {
                        if ($tk.PSObject.Properties['pokers'] -and $tk.pokers) {
                            $existingPokers = @($tk.pokers)
                        } else {
                            $existingPokers = @()
                        }
                        if ($existingPokers -notcontains $pokerUser) {
                            $existingPokers += $pokerUser
                        }
                        $tk | Add-Member -NotePropertyName pokers -NotePropertyValue $existingPokers -Force
                    }
                    Set-Variable -Name pokedTicket -Value $tk -Scope 2
                }
            }
            return $data
        }
        if ($completed) {
            Send-JsonResponse $res 400 @{ error = 'Cannot poke a completed request.' }
            return
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        # Email poke reminder
        if ($pokedTicket) {
            Send-PokeEmail $pokedTicket
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/edit ─────────────────────
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/edit$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        $windowsUser = if ($parsed.PSObject.Properties['windowsUser']) { $parsed.windowsUser } else { '' }
        $found = $false
        $forbidden = $false
        $badStatus = $false
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    if (-not $windowsUser -or $tk.windowsUser -ne $windowsUser) {
                        Set-Variable -Name forbidden -Value $true -Scope 2
                        return $data
                    }
                    if ($tk.status -notin @('open', 'in-progress')) {
                        Set-Variable -Name badStatus -Value $true -Scope 2
                        return $data
                    }
                    $tk.title = $parsed.title
                    $tk.description = $parsed.description
                    $tk.priority = $parsed.priority
                    $tk.type = $parsed.type
                    $tk | Add-Member -NotePropertyName lastEdited -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd') -Force
                }
            }
            return $data
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        if ($forbidden) {
            Send-JsonResponse $res 403 @{ error = 'You can only edit your own requests.' }
            return
        }
        if ($badStatus) {
            Send-JsonResponse $res 400 @{ error = 'Cannot edit a completed request.' }
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── POST /api/tickets/{id}/attach ─────────────────────
    if ($method -eq 'POST' -and $url -match '^/api/tickets/(\d+)/attach$') {
        $id = [long]$Matches[1]
        $fileName = [System.Web.HttpUtility]::UrlDecode($req.QueryString['name'])
        if (-not $fileName) {
            Send-JsonResponse $res 400 @{ error = 'Missing file name query parameter (?name=...).' }
            return
        }
        # Sanitize filename — remove path traversal characters
        $fileName = [System.IO.Path]::GetFileName($fileName)
        $attachDir = Join-Path $DataDir "attachments\$id"
        if (-not (Test-Path $attachDir)) {
            New-Item -ItemType Directory -Path $attachDir -Force | Out-Null
        }
        $filePath = Join-Path $attachDir $fileName
        # Read binary body and save
        try {
            $ms = New-Object System.IO.MemoryStream
            $req.InputStream.CopyTo($ms)
            $bytes = $ms.ToArray()
            $ms.Close()
            # Limit to 25 MB
            if ($bytes.Length -gt 26214400) {
                Send-JsonResponse $res 400 @{ error = 'File too large (max 25 MB).' }
                return
            }
            [System.IO.File]::WriteAllBytes($filePath, $bytes)
            Write-Host "  Attachment saved: $filePath ($([math]::Round($bytes.Length/1024))KB)" -ForegroundColor Cyan
            Send-JsonResponse $res 200 @{ ok = $true; name = $fileName; size = $bytes.Length }
        } catch {
            Send-JsonResponse $res 500 @{ error = "Failed to save attachment: $($_.Exception.Message)" }
        }
        return
    }

    # ── GET /api/attachments/{id}/{filename} ─────────────
    if ($method -eq 'GET' -and $url -match '^/api/attachments/(\d+)/(.+)$') {
        $id = $Matches[1]
        $fileName = [System.Web.HttpUtility]::UrlDecode($Matches[2])
        $fileName = [System.IO.Path]::GetFileName($fileName)
        $filePath = Join-Path $DataDir "attachments\$id\$fileName"
        if (-not (Test-Path $filePath)) {
            Send-JsonResponse $res 404 @{ error = 'Attachment not found.' }
            return
        }
        # Determine content type
        $ext = [System.IO.Path]::GetExtension($fileName).ToLower()
        $mimeTypes = @{
            '.jpg'  = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.png' = 'image/png'
            '.gif'  = 'image/gif'; '.bmp' = 'image/bmp'; '.webp' = 'image/webp'
            '.svg'  = 'image/svg+xml'; '.pdf' = 'application/pdf'
            '.zip'  = 'application/zip'; '.txt' = 'text/plain'
            '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }
        $contentType = if ($mimeTypes.ContainsKey($ext)) { $mimeTypes[$ext] } else { 'application/octet-stream' }
        $res.StatusCode = 200
        $res.ContentType = $contentType
        $res.AddHeader('Access-Control-Allow-Origin', '*')
        $res.AddHeader('Cache-Control', 'public, max-age=3600')
        # For non-image files, suggest download
        $isImage = $ext -in @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg')
        if (-not $isImage) {
            $res.AddHeader('Content-Disposition', "attachment; filename=`"$fileName`"")
        }
        $buffer = [System.IO.File]::ReadAllBytes($filePath)
        $res.ContentLength64 = $buffer.Length
        $res.OutputStream.Write($buffer, 0, $buffer.Length)
        $res.OutputStream.Close()
        return
    }

    # ── POST /api/tickets/{id}/comment (add comment) ────
    if ($method -eq 'POST' -and $url -match '^/api/tickets/(\d+)/comment$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        $found = $false
        $forbidden = $false
        $duplicate = $false
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    # Check authorization: must be ticket creator or a voter
                    $isCreator = ($parsed.windowsUser -and $tk.windowsUser -eq $parsed.windowsUser)
                    $isVoter = $false
                    if ($parsed.windowsUser -and $tk.PSObject.Properties['voters'] -and $tk.voters) {
                        $isVoter = $parsed.windowsUser -in @($tk.voters)
                    }
                    if (-not $isCreator -and -not $isVoter) {
                        Set-Variable -Name forbidden -Value $true -Scope 2
                        return $data
                    }
                    # Check for existing comment from this user
                    if ($tk.PSObject.Properties['comments'] -and $tk.comments) {
                        foreach ($c in @($tk.comments)) {
                            if ($c.windowsUser -eq $parsed.windowsUser) {
                                Set-Variable -Name duplicate -Value $true -Scope 2
                                return $data
                            }
                        }
                    }
                    # Create comment
                    $comment = @{
                        windowsUser = $parsed.windowsUser
                        displayName = $parsed.displayName
                        text        = $parsed.text
                        gifUrl      = $parsed.gifUrl
                        date        = (Get-Date -Format 'yyyy-MM-dd')
                    }
                    # Append to comments array (create if needed)
                    $existingComments = @()
                    if ($tk.PSObject.Properties['comments'] -and $tk.comments) {
                        $existingComments = @($tk.comments)
                    }
                    $existingComments += [pscustomobject]$comment
                    $tk | Add-Member -NotePropertyName comments -NotePropertyValue $existingComments -Force
                }
            }
            return $data
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        if ($forbidden) {
            Send-JsonResponse $res 403 @{ error = 'You must vote on a request before commenting.' }
            return
        }
        if ($duplicate) {
            Send-JsonResponse $res 400 @{ error = 'You already have a comment. Use edit instead.' }
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/comment (edit comment) ──
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/comment$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        $found = $false
        $commentFound = $false
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    if ($tk.PSObject.Properties['comments'] -and $tk.comments) {
                        foreach ($c in @($tk.comments)) {
                            if ($c.windowsUser -eq $parsed.windowsUser) {
                                $c.text = $parsed.text
                                $c | Add-Member -NotePropertyName gifUrl -NotePropertyValue $parsed.gifUrl -Force
                                $c | Add-Member -NotePropertyName editedDate -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd') -Force
                                Set-Variable -Name commentFound -Value $true -Scope 2
                            }
                        }
                    }
                }
            }
            return $data
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        if (-not $commentFound) {
            Send-JsonResponse $res 404 @{ error = 'Comment not found.' }
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── DELETE /api/tickets/{id}/comment (admin delete comment) ──
    if ($method -eq 'DELETE' -and $url -match '^/api/tickets/(\d+)/comment$') {
        $id = [long]$Matches[1]
        $body = Read-RequestBody $req
        $parsed = $body | ConvertFrom-Json
        $targetUser = $parsed.windowsUser
        $found = $false
        $commentFound = $false
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                    if ($tk.PSObject.Properties['comments'] -and $tk.comments) {
                        $newComments = @($tk.comments | Where-Object { $_.windowsUser -ne $targetUser })
                        if ($newComments.Count -lt @($tk.comments).Count) {
                            Set-Variable -Name commentFound -Value $true -Scope 2
                        }
                        $tk | Add-Member -NotePropertyName comments -NotePropertyValue $newComments -Force
                    }
                }
            }
            return $data
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        if (-not $commentFound) {
            Send-JsonResponse $res 404 @{ error = 'Comment not found.' }
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── PATCH /api/tickets/{id}/delete (admin) ──────────
    if ($method -eq 'PATCH' -and $url -match '^/api/tickets/(\d+)/delete$') {
        $id = [long]$Matches[1]
        $found = $false
        $updated = Invoke-LockedMutate $TkFile {
            param($data)
            $newData = @()
            foreach ($tk in $data) {
                if ([string]$tk.id -eq [string]$id) {
                    Set-Variable -Name found -Value $true -Scope 2
                } else {
                    $newData += $tk
                }
            }
            return $newData
        }
        if (-not $found) {
            Send-JsonResponse $res 404 @{ error = 'Ticket not found.' }
            return
        }
        Send-JsonResponse $res 200 @{ ok = $true; entries = @($updated) }
        return
    }

    # ── GET /api/version ─────────────────────────────────
    # Returns a stable id for the current HTML file (last-write-time ticks).
    # The page polls this to detect when a deploy has updated the HTML so it
    # can offer the user a refresh. This is how users pick up content updates
    # without having to close and reopen the guide.
    if ($method -eq 'GET' -and $url -eq '/api/version') {
        $ver = ''
        try {
            if (Test-Path $HtmlFile) {
                $ver = (Get-Item $HtmlFile).LastWriteTimeUtc.Ticks.ToString()
            }
        } catch {}
        Send-JsonResponse $res 200 @{ version = $ver }
        return
    }

    # ── GET /api/shutdown (localhost only) ───────────────
    if ($method -eq 'GET' -and $url -eq '/api/shutdown') {
        Send-JsonResponse $res 200 @{ ok = $true; message = 'Server shutting down.' }
        Write-Host "  Shutdown requested via browser." -ForegroundColor Yellow
        return 'SHUTDOWN'
    }

    # ── GET /api/restart ───────────────────────────────
    # Stops this server and spawns a replacement, so the browser can reconnect
    # on the same port after a short retry.
    if ($method -eq 'GET' -and $url -eq '/api/restart') {
        Send-JsonResponse $res 200 @{ ok = $true; message = 'Server restarting.' }
        Write-Host "  Restart requested via browser." -ForegroundColor Yellow
        return 'RESTART'
    }

    # ── 404 ─────────────────────────────────────────────
    Send-JsonResponse $res 404 @{ error = "Not found: $url" }
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

    # Try each candidate port in order — falls through if port is busy
    $listener = $null
    foreach ($candidate in $PortCandidates) {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:${candidate}/")
        try {
            $listener.Start()
            $script:Port = $candidate
            break
        } catch {
            Write-Host "  Port $candidate busy, trying next..." -ForegroundColor DarkGray
            try { $listener.Close() } catch {}
            $listener = $null
        }
    }

    if ($null -eq $listener) {
        Write-Host "  ERROR: Could not bind to any port in: $($PortCandidates -join ', ')" -ForegroundColor Red
        Write-Host "  All ports are in use. Close other instances or contact Dylan." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    # Advertise chosen port so the VBS launcher / other clients can find us
    try {
        [System.IO.File]::WriteAllText($PortStatusFile, "$Port", [System.Text.Encoding]::ASCII)
    } catch {
        Write-Host "  (Could not write port status file: $($_.Exception.Message))" -ForegroundColor DarkGray
    }

    Write-Host "  Server running    : http://localhost:$Port" -ForegroundColor Green
    Write-Host "  HTML source       : $HtmlFile"
    Write-Host "  Data directory    : $DataDir"
    if ($EmailEnabled) {
        Write-Host "  Email alerts      : $EmailTo (via Outlook)" -ForegroundColor Green
    } else {
        Write-Host "  Email alerts      : disabled"
    }
    Write-Host ""
    Write-Host "  Press Ctrl+C to stop." -ForegroundColor DarkGray
    Write-Host ""

    $InactivityMinutes = 10
    $LastActivity = [DateTime]::UtcNow

    $running = $true
    while ($running -and $listener.IsListening) {
        # Auto-shutdown after inactivity
        if (([DateTime]::UtcNow - $LastActivity).TotalMinutes -gt $InactivityMinutes) {
            Write-Host "  No activity for $InactivityMinutes minutes — shutting down." -ForegroundColor Yellow
            break
        }

        $context = $null
        try { $context = $listener.GetContext() } catch { break }
        $LastActivity = [DateTime]::UtcNow

        try {
            $result = Handle-Request $context.Request $context.Response
            if ($result -eq 'SHUTDOWN') {
                $listener.Stop()
                $running = $false
            }
            elseif ($result -eq 'RESTART') {
                # Stop listener FIRST so the port is freed before the child tries to bind it.
                $listener.Stop()
                $listener.Close()
                # Small pause so Windows actually releases the port.
                Start-Sleep -Milliseconds 400
                # Spawn replacement (detached, hidden) running this same script.
                try {
                    $scriptPath = $PSCommandPath
                    if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
                    Start-Process -FilePath 'powershell.exe' `
                        -ArgumentList @('-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File', $scriptPath) `
                        -WindowStyle Hidden
                    Write-Host "  Replacement server launched." -ForegroundColor Green
                } catch {
                    Write-Host "  ERROR: Could not launch replacement: $($_.Exception.Message)" -ForegroundColor Red
                }
                $running = $false
            }
        } catch {
            $errUrl = $context.Request.Url.AbsolutePath
            $errMethod = $context.Request.HttpMethod
            Write-Host "  ERROR [$errMethod $errUrl]: $($_.Exception.Message)" -ForegroundColor Red
            Write-ErrorLog "$errMethod $errUrl" $_.Exception.Message $_.ScriptStackTrace
            try { Send-JsonResponse $context.Response 500 @{ error = $_.Exception.Message } } catch {}
        }
    }

    # Safe to call Stop/Close again even if already stopped (e.g. restart path).
    try { $listener.Stop() } catch {}
    try { $listener.Close() } catch {}
    Release-Lock $ClFile
    Release-Lock $TkFile
    # Only remove the port file if it still points at our port. On restart the
    # replacement child may have already written its own value — don't clobber.
    try {
        if (Test-Path $PortStatusFile) {
            $existing = (Get-Content $PortStatusFile -Raw -ErrorAction SilentlyContinue).Trim()
            if ($existing -eq "$Port") {
                Remove-Item $PortStatusFile -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {}
    Write-Host "  Server stopped." -ForegroundColor Yellow
}

# ── Run ─────────────────────────────────────────────────────────────────
Start-Server
