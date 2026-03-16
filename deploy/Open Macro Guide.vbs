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

' ── Check if already running ────────────────────────────────────────────
On Error Resume Next
Dim http
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "http://localhost:8123/", False
http.Send
If http.Status = 200 Then
    ' Already running — just open the browser and exit silently
    Dim winUser2
    winUser2 = WshShell.ExpandEnvironmentStrings("%USERNAME%")
    WshShell.Run "http://localhost:8123/?user=" & winUser2
    WScript.Quit
End If
On Error GoTo 0

' ── Start server (hidden console window) ────────────────────────────────
' -ExecutionPolicy Bypass ensures it runs regardless of machine policy
' -WindowStyle Hidden keeps the PowerShell window invisible
Set proc = WshShell.Exec("powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & serverScript & """")

' Wait for server to be ready (up to 8 seconds)
Dim attempts
attempts = 0
Do While attempts < 16
    WScript.Sleep 500
    attempts = attempts + 1
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "http://localhost:8123/", False
    http.Send
    If http.Status = 200 Then Exit Do
    On Error GoTo 0
Loop

' Check if it started
On Error Resume Next
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "http://localhost:8123/", False
http.Send
If http.Status <> 200 Then
    MsgBox "The server could not start." & vbCrLf & vbCrLf & _
           "This might happen if:" & vbCrLf & _
           "  - The Y: drive is not connected" & vbCrLf & _
           "  - Another program is using port 8123" & vbCrLf & _
           "  - PowerShell is restricted on this machine" & vbCrLf & vbCrLf & _
           "Please contact Dylan for help.", _
           vbExclamation, "Macro Guide"
    proc.Terminate
    WScript.Quit
End If
On Error GoTo 0

' ── Open browser with Windows username for identity tracking ─────────────
Dim winUser
winUser = WshShell.ExpandEnvironmentStrings("%USERNAME%")
WshShell.Run "http://localhost:8123/?user=" & winUser

' Server runs in background — shuts itself down after inactivity.
' No further action needed. You can close this window.
