Attribute VB_Name = "ViewLogs"
' ====================================================================================
' ViewLogs.bas - Centralized Log Viewer
' ====================================================================================
'
' PURPOSE:
'   Allows admins/engineers to easily view macro logs from any user.
'   Lists all logs in the shared directory and allows opening the selected one.
'
' ====================================================================================

Option Explicit

' ====================================================================================
' ViewRecentLogs - Display recent logs from the shared folder
' ====================================================================================
Sub ViewRecentLogs()
    Dim sharedDir As String
    Dim fileName As String
    Dim logFiles As Object
    Dim fileList As String
    Dim fso As Object
    Dim fileCount As Integer
    Dim userChoice As String
    
    On Error GoTo ErrorHandler

    sharedDir = GetSharedLogDir()
    
    If Dir(sharedDir, vbDirectory) = "" Then
        MsgBox "Shared log directory not found: " & sharedDir, vbCritical
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFiles = CreateObject("Scripting.Dictionary")
    
    fileName = Dir(sharedDir & "*.log")
    
    ' Collect files and sort by date (newest first)
    Do While fileName <> ""
        logFiles.Add fileName, fso.GetFile(sharedDir & fileName).DateLastModified
        fileName = Dir()
    Loop
    
    If logFiles.Count = 0 Then
        MsgBox "No logs found in " & sharedDir, vbInformation
        Exit Sub
    End If

    ' Sort the dictionary by value (date) - simple bubble sort for small number of files
    Dim keys() As Variant
    keys = logFiles.keys
    
    Dim i As Integer, j As Integer
    Dim temp As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If logFiles(keys(i)) < logFiles(keys(j)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

    ' Build selection list (top 20 newest)
    fileList = "Select a log file to open (Newest first):" & vbCrLf & String(50, "-") & vbCrLf
    Dim limit As Integer
    limit = IIf(UBound(keys) > 19, 19, UBound(keys))
    
    For i = 0 To limit
        fileList = fileList & (i + 1) & ". " & keys(i) & " (" & logFiles(keys(i)) & ")" & vbCrLf
    Next i
    
    fileList = fileList & String(50, "-") & vbCrLf & "Enter the number (1-" & (limit + 1) & ") to open, or cancel:"

    userChoice = InputBox(fileList, "Shared Log Viewer", "1")
    
    If userChoice = "" Then Exit Sub
    
    If IsNumeric(userChoice) Then
        Dim idx As Integer
        idx = CInt(userChoice) - 1
        If idx >= 0 And idx <= limit Then
            ' Open the selected log file in Notepad
            shell "notepad.exe """ & sharedDir & keys(idx) & """", vbNormalFocus
        Else
            MsgBox "Invalid selection.", vbExclamation
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error viewing logs: " & Err.description, vbCritical
End Sub

' ====================================================================================
' OpenSharedLogFolder - Quick shortcut to open the folder in Explorer
' ====================================================================================
Sub OpenSharedLogFolder()
    Dim sharedDir As String
    sharedDir = GetSharedLogDir()
    
    If Dir(sharedDir, vbDirectory) <> "" Then
        shell "explorer.exe """ & sharedDir & """", vbNormalFocus
    Else
        MsgBox "Shared log directory not found.", vbExclamation
    End If
End Sub
