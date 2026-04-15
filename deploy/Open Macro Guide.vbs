' =====================================================
'  Nordic Minesteel Macro Guide — Launcher
'  Maintained by Dylan Lebel
'
'  Double-click this file to open the Macro Guide.
'  Uses PowerShell (built into Windows — nothing to install).
'  The server runs in the background and shuts down automatically when idle.
' =====================================================

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ── Server location (always on Y: drive) ────────────────────────────────
guideDir = "Y:\Solidworks\Macros\Macro Data PDM\MacroGuide"
serverScript = guideDir & "\macro-guide-server.ps1"

If Not fso.FileExists(serverScript) Then
    MsgBox "Server file not found:" & vbCrLf & serverScript & vbCrLf & vbCrLf & _
           "Please contact Dylan to fix the installation.", _
           vbExclamation, "Macro Guide"
    WScript.Quit
End If

' ── Candidate ports (must match $PortCandidates in macro-guide-server.ps1) ─
Dim ports
ports = Array(8123, 8124, 8125, 8126)

' Helper: probe a port — returns the port number if alive, 0 otherwise
Function ProbePort(p)
    ProbePort = 0
    On Error Resume Next
    Dim h
    Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "GET", "http://localhost:" & p & "/", False
    h.Send
    If Err.Number = 0 And h.Status = 200 Then ProbePort = p
    On Error GoTo 0
End Function

' Helper: read port status file written by server (fast path)
Function ReadPortFile()
    ReadPortFile = 0
    On Error Resume Next
    Dim tmp, f, txt
    tmp = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\MacroGuide.port"
    If fso.FileExists(tmp) Then
        Set f = fso.OpenTextFile(tmp, 1)
        txt = Trim(f.ReadAll)
        f.Close
        If IsNumeric(txt) Then ReadPortFile = CInt(txt)
    End If
    On Error GoTo 0
End Function

' ── Check if already running on any candidate port ──────────────────────
Dim activePort, i
activePort = 0

' Fast path: check the port file first
Dim hinted
hinted = ReadPortFile()
If hinted > 0 Then activePort = ProbePort(hinted)

If activePort = 0 Then
    For i = 0 To UBound(ports)
        activePort = ProbePort(ports(i))
        If activePort > 0 Then Exit For
    Next
End If

Dim winUser
winUser = WshShell.ExpandEnvironmentStrings("%USERNAME%")

If activePort > 0 Then
    ' Already running — open browser to the live port and exit
    WshShell.Run "http://localhost:" & activePort & "/?user=" & winUser
    WScript.Quit
End If

' ── Start server (hidden console window) ────────────────────────────────
' -ExecutionPolicy Bypass ensures it runs regardless of machine policy
' -WindowStyle Hidden keeps the PowerShell window invisible
Set proc = WshShell.Exec("powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & serverScript & """")

' Wait for server to be ready (up to 8 seconds) — probe all candidate ports
Dim attempts
attempts = 0
Do While attempts < 16 And activePort = 0
    WScript.Sleep 500
    attempts = attempts + 1
    hinted = ReadPortFile()
    If hinted > 0 Then activePort = ProbePort(hinted)
    If activePort = 0 Then
        For i = 0 To UBound(ports)
            activePort = ProbePort(ports(i))
            If activePort > 0 Then Exit For
        Next
    End If
Loop

If activePort = 0 Then
    MsgBox "The server could not start." & vbCrLf & vbCrLf & _
           "This might happen if:" & vbCrLf & _
           "  - The Y: drive is not connected" & vbCrLf & _
           "  - All ports 8123-8126 are in use" & vbCrLf & _
           "  - PowerShell is restricted on this machine" & vbCrLf & vbCrLf & _
           "Please contact Dylan for help.", _
           vbExclamation, "Macro Guide"
    On Error Resume Next
    proc.Terminate
    On Error GoTo 0
    WScript.Quit
End If

' ── Open browser with Windows username for identity tracking ─────────────
WshShell.Run "http://localhost:" & activePort & "/?user=" & winUser

' Server runs in background — shuts itself down after inactivity.
' No further action needed. You can close this window.
