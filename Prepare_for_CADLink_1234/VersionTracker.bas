Attribute VB_Name = "VersionTracker"
' ====================================================================================
' VersionTracker.bas - User Version Tracking Module
' ====================================================================================
'
' PURPOSE:
'   Tracks which users are running the macro and what version they're using.
'   Maintains a shared log file in PDM that shows last version used by each user.
'
' FUNCTIONS:
'   - LogVersionUsage()     : Records current user and version to tracking file
'   - GetWindowsUsername()  : Gets the current Windows username
'
' ====================================================================================

Option Explicit

' Windows API declaration to get current username
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#End If

' ====================================================================================
' GetWindowsUsername - Get current Windows username
' ====================================================================================
Public Function GetWindowsUsername() As String
    Dim buffer As String
    Dim size As Long

    buffer = String(255, 0)
    size = Len(buffer)

    If GetUserName(buffer, size) <> 0 Then
        GetWindowsUsername = Left$(buffer, size - 1)
    Else
        GetWindowsUsername = Environ$("USERNAME") ' Fallback
    End If
End Function

' ====================================================================================
' LogVersionUsage - Record user and version to shared tracking file
' ====================================================================================
' This function creates/updates a tracking file in PDM showing who's using what version
' File format: Username,LastVersion,LastUsedDate,TotalRuns
' ====================================================================================
Public Sub LogVersionUsage()
    On Error Resume Next

    Dim username As String
    Dim trackingFilePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim userRecords As Object
    Dim recordParts() As String
    Dim currentDate As String
    Dim userFound As Boolean
    Dim recordKey As String
    Dim outputLines As String

    ' Get current user
    username = GetWindowsUsername()
    If username = "" Then Exit Sub

    ' Define tracking file path in hidden shared location (everyone can write, only you know about it)
    trackingFilePath = "Y:\Solidworks\Macros\Macro Data PDM\MacroVersionTracking.csv"

    ' Create dictionary to store user records
    Set userRecords = CreateObject("Scripting.Dictionary")
    userRecords.CompareMode = 1 ' Text compare, case insensitive

    currentDate = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    userFound = False

    ' Read existing tracking file if it exists
    fileNum = FreeFile
    If Dir(trackingFilePath) <> "" Then
        Open trackingFilePath For Input As #fileNum

        ' Skip header line if exists
        If Not EOF(fileNum) Then
            Line Input #fileNum, lineText
            If InStr(lineText, "Username") = 0 Then
                ' No header, process this line as data
                If Trim$(lineText) <> "" Then
                    recordParts = Split(lineText, ",")
                    If UBound(recordParts) >= 2 Then
                        recordKey = Trim$(recordParts(0))
                        userRecords(recordKey) = lineText
                    End If
                End If
            End If
        End If

        ' Read all user records
        Do While Not EOF(fileNum)
            Line Input #fileNum, lineText
            If Trim$(lineText) <> "" Then
                recordParts = Split(lineText, ",")
                If UBound(recordParts) >= 2 Then
                    recordKey = Trim$(recordParts(0))
                    userRecords(recordKey) = lineText
                End If
            End If
        Loop

        Close #fileNum
    End If

    ' Update or add current user's record
    If userRecords.exists(username) Then
        ' Parse existing record
        recordParts = Split(userRecords(username), ",")
        Dim totalRuns As Long
        totalRuns = 1
        If UBound(recordParts) >= 3 Then
            totalRuns = CLng(recordParts(3)) + 1
        End If
        ' Update record with new version and date
        userRecords(username) = username & "," & MacroVersion & "," & currentDate & "," & totalRuns
    Else
        ' Add new user record
        userRecords(username) = username & "," & MacroVersion & "," & currentDate & ",1"
    End If

    ' Write updated records back to file
    fileNum = FreeFile
    Open trackingFilePath For Output As #fileNum

    ' Write header
    Print #fileNum, "Username,LastVersion,LastUsedDate,TotalRuns"

    ' Write all records
    Dim key As Variant
    For Each key In userRecords.keys
        Print #fileNum, userRecords(key)
    Next key

    Close #fileNum

    ' Log to debug if enabled
    Logger.LogInfo "Version tracking updated: " & username & " using v" & MacroVersion

End Sub
