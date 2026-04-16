' =====================================================
'  Nordic Minesteel Macro Guide - Launcher
'  Maintained by Dylan Lebel
'
'  Double-click this file to open the Macro Guide.
'  Uses PowerShell (built into Windows - nothing to install).
'  The server runs in the background and shuts down automatically when idle.
' =====================================================

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

tmpDir = WshShell.ExpandEnvironmentStrings("%TEMP%")
logPath = tmpDir & "\MacroGuide-launcher.log"

Randomize
donePath = tmpDir & "\MacroGuide_splash_done_" & CStr(Int(Timer * 1000)) & "_" & CStr(Int(Rnd() * 1000000)) & ".flag"
On Error Resume Next
If fso.FileExists(donePath) Then fso.DeleteFile donePath, True
On Error GoTo 0

' Server location (always on Y: drive)
guideDir = "Y:\Solidworks\Macros\Macro Data PDM\MacroGuide"
serverScript = guideDir & "\macro-guide-server.ps1"

' Candidate ports (must match $PortCandidates in macro-guide-server.ps1)
Dim ports
ports = Array(8123, 8124, 8125, 8126)

Sub LogMessage(message)
    On Error Resume Next
    Dim lf
    Set lf = fso.OpenTextFile(logPath, 8, True)
    lf.WriteLine CStr(Now) & " " & message
    lf.Close
    On Error GoTo 0
End Sub

Function Q(value)
    Q = """" & value & """"
End Function

Sub MarkSplashDone()
    On Error Resume Next
    Dim df
    Set df = fso.CreateTextFile(donePath, True)
    df.WriteLine CStr(Now)
    df.Close
    On Error GoTo 0
End Sub

Sub CleanupLock()
    On Error Resume Next
    If createdLock Then
        If fso.FileExists(lockPath) Then fso.DeleteFile lockPath, True
    End If
    On Error GoTo 0
End Sub

Sub FailAndQuit(message)
    LogMessage "ERROR: " & Replace(message, vbCrLf, " | ")
    CleanupLock
    MarkSplashDone
    WScript.Sleep 400
    MsgBox message, vbExclamation, "Macro Guide"
    WScript.Quit 1
End Sub

Sub OpenGuide(port)
    LogMessage "Opening browser on port " & CStr(port)
    CleanupLock
    MarkSplashDone
    WScript.Sleep 350
    WshShell.Run "http://localhost:" & port & "/?user=" & winUser
    WScript.Quit 0
End Sub

Sub StartDelayedSplash()
    On Error Resume Next

    Dim splashScript, sf, cmd
    splashScript = tmpDir & "\MacroGuide_splash.ps1"

    Set sf = fso.CreateTextFile(splashScript, True)
    sf.WriteLine "param("
    sf.WriteLine "    [string]$DonePath,"
    sf.WriteLine "    [int]$DelayMs = 1000,"
    sf.WriteLine "    [string]$LogPath = ''"
    sf.WriteLine ")"
    sf.WriteLine "function Write-SplashLog {"
    sf.WriteLine "    param([string]$Message)"
    sf.WriteLine "    try {"
    sf.WriteLine "        if ($LogPath) {"
    sf.WriteLine "            $line = ""$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [splash] $Message"""
    sf.WriteLine "            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue"
    sf.WriteLine "        }"
    sf.WriteLine "    } catch {}"
    sf.WriteLine "}"
    sf.WriteLine "try {"
    sf.WriteLine "    Start-Sleep -Milliseconds $DelayMs"
    sf.WriteLine "    if (Test-Path -LiteralPath $DonePath) { exit 0 }"
    sf.WriteLine "    Add-Type -AssemblyName System.Windows.Forms"
    sf.WriteLine "    Add-Type -AssemblyName System.Drawing"
    sf.WriteLine "    if (Test-Path -LiteralPath $DonePath) { exit 0 }"
    sf.WriteLine "    $form = New-Object System.Windows.Forms.Form"
    sf.WriteLine "    $form.Text = 'Macro Guide'"
    sf.WriteLine "    $form.Size = New-Object System.Drawing.Size(390, 150)"
    sf.WriteLine "    $form.StartPosition = 'CenterScreen'"
    sf.WriteLine "    $form.TopMost = $true"
    sf.WriteLine "    $form.ShowInTaskbar = $true"
    sf.WriteLine "    $form.FormBorderStyle = 'FixedDialog'"
    sf.WriteLine "    $form.MaximizeBox = $false"
    sf.WriteLine "    $form.MinimizeBox = $false"
    sf.WriteLine "    $form.BackColor = [System.Drawing.Color]::White"
    sf.WriteLine "    $label = New-Object System.Windows.Forms.Label"
    sf.WriteLine "    $label.AutoSize = $false"
    sf.WriteLine "    $label.Text = ""Starting Macro Guide...`r`n`r`nYour browser will open automatically."""
    sf.WriteLine "    $label.Font = New-Object System.Drawing.Font('Segoe UI', 10)"
    sf.WriteLine "    $label.Location = New-Object System.Drawing.Point(24, 22)"
    sf.WriteLine "    $label.Size = New-Object System.Drawing.Size(335, 78)"
    sf.WriteLine "    $form.Controls.Add($label)"
    sf.WriteLine "    $timer = New-Object System.Windows.Forms.Timer"
    sf.WriteLine "    $timer.Interval = 250"
    sf.WriteLine "    $started = Get-Date"
    sf.WriteLine "    $timer.Add_Tick({"
    sf.WriteLine "        if ((Test-Path -LiteralPath $DonePath) -or (((Get-Date) - $started).TotalSeconds -ge 30)) {"
    sf.WriteLine "            $timer.Stop()"
    sf.WriteLine "            $form.Close()"
    sf.WriteLine "        }"
    sf.WriteLine "    })"
    sf.WriteLine "    $timer.Start()"
    sf.WriteLine "    [void]$form.ShowDialog()"
    sf.WriteLine "} catch {"
    sf.WriteLine "    Write-SplashLog $_.Exception.Message"
    sf.WriteLine "}"
    sf.Close

    If Err.Number = 0 Then
        cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Q(splashScript) & " -DonePath " & Q(donePath) & " -DelayMs 1000 -LogPath " & Q(logPath)
        WshShell.Run cmd, 0, False
        LogMessage "Delayed splash helper started"
    Else
        LogMessage "Delayed splash helper could not be written: " & Err.Description
    End If

    On Error GoTo 0
End Sub

' Probe only the Macro Guide server, not just any HTTP listener on the port.
' /api/version is lightweight and proves the app server is answering.
Function ProbePort(p)
    ProbePort = 0
    On Error Resume Next

    Dim h, body
    Set h = CreateObject("MSXML2.ServerXMLHTTP")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    h.setTimeouts 300, 300, 500, 700
    Err.Clear
    h.Open "GET", "http://localhost:" & p & "/api/version?launcher=" & CStr(Int(Timer * 1000)), False
    h.Send
    If Err.Number = 0 Then
        body = h.responseText
        If h.Status = 200 And InStr(1, body, "version", vbTextCompare) > 0 Then ProbePort = p
    End If

    Set h = Nothing
    On Error GoTo 0
End Function

' Require two successful probes so the browser is not opened on a stale
' port-file hint or a listener that briefly answered and then died.
Function ConfirmReady(p)
    ConfirmReady = False
    If ProbePort(p) = p Then
        WScript.Sleep 250
        If ProbePort(p) = p Then ConfirmReady = True
    End If
End Function

' Read port status file written by server.
Function ReadPortFile()
    ReadPortFile = 0
    On Error Resume Next

    Dim tmp, pf, txt
    tmp = tmpDir & "\MacroGuide.port"
    If fso.FileExists(tmp) Then
        Set pf = fso.OpenTextFile(tmp, 1)
        txt = Trim(pf.ReadAll)
        pf.Close
        If IsNumeric(txt) Then ReadPortFile = CInt(txt)
    End If

    On Error GoTo 0
End Function

Sub FindReadyPort(ByRef activePort)
    Dim i, hinted
    activePort = 0

    hinted = ReadPortFile()
    If hinted > 0 Then
        If ConfirmReady(hinted) Then activePort = hinted
    End If

    If activePort = 0 Then
        For i = 0 To UBound(ports)
            If ConfirmReady(ports(i)) Then
                activePort = ports(i)
                Exit For
            End If
        Next
    End If
End Sub

LogMessage "Launcher start"
StartDelayedSplash

winUser = WshShell.ExpandEnvironmentStrings("%USERNAME%")
createdLock = False
lockPath = tmpDir & "\MacroGuide.starting"

' Fast path: if the server is already running, this normally finishes before
' the delayed splash ever appears.
Dim activePort
FindReadyPort activePort
If activePort > 0 Then
    LogMessage "Existing server ready on port " & CStr(activePort)
    OpenGuide activePort
End If

' Slow path: no existing server answered. At this point the delayed splash is
' already running independently, so a cold Y: drive cannot block all feedback.
If Not fso.FileExists(serverScript) Then
    FailAndQuit "Server file not found:" & vbCrLf & serverScript & vbCrLf & vbCrLf & _
                "Please contact Dylan to fix the installation."
End If

' Anti-spam-click: if another instance started the server in the last 20s,
' skip the server-start step. This instance still waits and opens the browser.
skipStart = False
If fso.FileExists(lockPath) Then
    On Error Resume Next
    Dim lockFile, lockAge
    Set lockFile = fso.GetFile(lockPath)
    lockAge = DateDiff("s", lockFile.DateLastModified, Now)
    If Err.Number = 0 And lockAge < 20 Then skipStart = True
    On Error GoTo 0
End If

If Not skipStart Then
    On Error Resume Next
    Dim startLock
    Set startLock = fso.CreateTextFile(lockPath, True)
    startLock.WriteLine CStr(Now)
    startLock.Close
    If Err.Number = 0 Then createdLock = True
    On Error GoTo 0

    LogMessage "Starting server from " & serverScript
    WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Q(serverScript), 0, False
Else
    LogMessage "Another launcher instance is already starting the server"
End If

' Wait for the Macro Guide server to be ready. Do not open the browser unless
' /api/version confirms the app is actually answering.
attempts = 0
Do While attempts < 30 And activePort = 0
    WScript.Sleep 500
    attempts = attempts + 1
    FindReadyPort activePort
Loop

If activePort = 0 Then
    FailAndQuit "The server could not start." & vbCrLf & vbCrLf & _
                "This might happen if:" & vbCrLf & _
                "  - The Y: drive is not connected" & vbCrLf & _
                "  - All ports 8123-8126 are in use" & vbCrLf & _
                "  - PowerShell is restricted on this machine" & vbCrLf & vbCrLf & _
                "Please contact Dylan for help."
End If

LogMessage "Server ready on port " & CStr(activePort)
OpenGuide activePort

' Server runs in background and shuts itself down after inactivity.
