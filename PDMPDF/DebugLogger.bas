Attribute VB_Name = "DebugLogger"
' ============================================================================
' MODULE: DebugLogger
' Description: Centralized logging for debugging batch PDF operations
' Usage:
'   1. Set DEBUG_LOGGING_ENABLED = True to enable logging
'   2. Call LogInit at the start of a batch run
'   3. Call LogMessage throughout the code to log events
'   4. Call LogClose at the end of the batch run
'   5. Log file is created in the user's Documents folder
' ============================================================================
Option Explicit

' ============================================================================
' CONFIGURATION - Set to True to enable logging, False to disable
' ============================================================================
Public Const DEBUG_LOGGING_ENABLED As Boolean = True

' ============================================================================
' PRIVATE VARIABLES
' ============================================================================
Private mLogFileNum As Integer
Private mLogFilePath As String
Private mLogStartTime As Double
Private mIsLogOpen As Boolean
Private mIndentLevel As Long

' ============================================================================
' PUBLIC METHODS
' ============================================================================

' Initialize logging for a new batch run
Public Sub LogInit(Optional logName As String = "PDFBatch")
    Dim logFolder As String
    Dim timestamp As String

    If Not DEBUG_LOGGING_ENABLED Then Exit Sub

    ' Close any existing log
    If mIsLogOpen Then LogClose

    ' Create log folder in user's Documents
    logFolder = Environ$("USERPROFILE") & "\Documents\PDF PDM Logs"
    On Error Resume Next
    MkDir logFolder
    Err.Clear

    ' Create timestamped log file name
    timestamp = Format(Now, "yyyy-mm-dd_HHmmss")
    mLogFilePath = logFolder & "\" & logName & "_" & timestamp & ".log"

    ' Open the file for writing with error handling
    mLogFileNum = FreeFile
    Open mLogFilePath For Output As #mLogFileNum
    If Err.Number <> 0 Then
        ' Failed to open log file - disable logging for this session
        Debug.Print "DebugLogger: Failed to create log file: " & Err.Description
        mIsLogOpen = False
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    mIsLogOpen = True
    mLogStartTime = Timer
    mIndentLevel = 0

    ' Write header
    WriteLogLine "========================================================================"
    WriteLogLine "PDF PDM Batch Log"
    WriteLogLine "========================================================================"
    WriteLogLine "Started: " & Format(Now, "yyyy-mm-dd HH:mm:ss")
    WriteLogLine "User: " & Environ$("USERNAME")
    WriteLogLine "Computer: " & Environ$("COMPUTERNAME")
    WriteLogLine "Log File: " & mLogFilePath
    WriteLogLine "========================================================================"
    WriteLogLine ""
End Sub

' Log a message with timestamp
Public Sub LogMessage(message As String, Optional level As String = "INFO")
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    Dim elapsed As String
    Dim indent As String

    elapsed = FormatElapsed(Timer - mLogStartTime)
    indent = String(mIndentLevel * 2, " ")

    WriteLogLine "[" & elapsed & "] [" & level & "] " & indent & message
End Sub

' Log an error
Public Sub LogError(message As String, Optional errNum As Long = 0, Optional errDesc As String = "")
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    Dim fullMsg As String
    fullMsg = message

    If errNum <> 0 Then
        fullMsg = fullMsg & " (Error " & errNum & ": " & errDesc & ")"
    End If

    LogMessage fullMsg, "ERROR"
End Sub

' Log a warning
Public Sub LogWarning(message As String)
    LogMessage message, "WARN"
End Sub

' Log debug-level detail (for verbose output)
Public Sub LogDebug(message As String)
    LogMessage message, "DEBUG"
End Sub

' Log the start of a section (increases indent)
Public Sub LogSectionStart(sectionName As String)
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    LogMessage ">>> " & sectionName
    mIndentLevel = mIndentLevel + 1
End Sub

' Log the end of a section (decreases indent)
Public Sub LogSectionEnd(Optional result As String = "")
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    If mIndentLevel > 0 Then mIndentLevel = mIndentLevel - 1

    If result <> "" Then
        LogMessage "<<< " & result
    Else
        LogMessage "<<< Done"
    End If
End Sub

' Log file processing status
Public Sub LogFileProcess(fileNum As Long, totalFiles As Long, filePath As String, status As String)
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    LogMessage "File " & fileNum & "/" & totalFiles & ": " & fileName & " - " & status
End Sub

' Log a separator line
Public Sub LogSeparator()
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    WriteLogLine "------------------------------------------------------------------------"
End Sub

' Log batch summary
Public Sub LogBatchSummary(totalFiles As Long, successCount As Long, errorCount As Long, warningCount As Long)
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    Dim totalTime As Double
    totalTime = Timer - mLogStartTime

    WriteLogLine ""
    WriteLogLine "========================================================================"
    WriteLogLine "BATCH SUMMARY"
    WriteLogLine "========================================================================"
    WriteLogLine "Total Files Processed: " & totalFiles
    WriteLogLine "  Success: " & successCount
    WriteLogLine "  Errors:  " & errorCount
    WriteLogLine "  Warnings: " & warningCount
    WriteLogLine ""
    WriteLogLine "Total Time: " & FormatElapsed(totalTime)
    If totalFiles > 0 Then
        WriteLogLine "Average Time Per File: " & FormatElapsed(totalTime / totalFiles)
    End If
    WriteLogLine "========================================================================"
End Sub

' Close the log file
Public Sub LogClose()
    If Not DEBUG_LOGGING_ENABLED Then Exit Sub
    If Not mIsLogOpen Then Exit Sub

    WriteLogLine ""
    WriteLogLine "========================================================================"
    WriteLogLine "Log ended: " & Format(Now, "yyyy-mm-dd HH:mm:ss")
    WriteLogLine "========================================================================"

    Close #mLogFileNum
    mIsLogOpen = False

    ' Output location to Immediate window for easy access
    Debug.Print "Log file created: " & mLogFilePath
End Sub

' Get the current log file path (useful for showing to user)
Public Function GetLogFilePath() As String
    GetLogFilePath = mLogFilePath
End Function

' Check if logging is currently active
Public Function IsLoggingActive() As Boolean
    IsLoggingActive = mIsLogOpen And DEBUG_LOGGING_ENABLED
End Function

' ============================================================================
' PRIVATE HELPER METHODS
' ============================================================================

Private Sub WriteLogLine(text As String)
    On Error Resume Next
    Print #mLogFileNum, text

    ' Also output to Immediate window for real-time viewing
    Debug.Print text
    On Error GoTo 0
End Sub

Private Function FormatElapsed(secs As Double) As String
    Dim h As Long, m As Long, s As Long, ms As Long

    If secs < 0 Then secs = 0

    h = Int(secs / 3600)
    m = Int((secs - h * 3600) / 60)
    s = Int(secs - h * 3600 - m * 60)
    ms = Int((secs - Int(secs)) * 1000)

    If h > 0 Then
        FormatElapsed = h & ":" & Format(m, "00") & ":" & Format(s, "00") & "." & Format(ms, "000")
    ElseIf m > 0 Then
        FormatElapsed = m & ":" & Format(s, "00") & "." & Format(ms, "000")
    Else
        FormatElapsed = s & "." & Format(ms, "000") & "s"
    End If
End Function

' ============================================================================
' CONVENIENCE METHODS FOR COMMON LOGGING SCENARIOS
' ============================================================================

' Log PDM vault connection
Public Sub LogPDMConnection(vaultName As String, success As Boolean)
    If success Then
        LogMessage "Connected to PDM vault: " & vaultName
    Else
        LogError "Failed to connect to PDM vault: " & vaultName
    End If
End Sub

' Log file checkout status
Public Sub LogCheckoutStatus(filePath As String, isCheckedOut As Boolean, checkedOutBy As String)
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    If isCheckedOut Then
        LogDebug fileName & " is checked out by: " & checkedOutBy
    Else
        LogDebug fileName & " is not checked out"
    End If
End Sub

' Log PDF creation result
Public Sub LogPDFResult(drawingPath As String, pdfPath As String, success As Boolean, Optional errorMsg As String = "")
    Dim drawingName As String
    drawingName = Mid(drawingPath, InStrRev(drawingPath, "\") + 1)

    If success Then
        LogMessage "PDF created: " & drawingName & " -> " & pdfPath
    Else
        LogError "PDF failed: " & drawingName & " - " & errorMsg
    End If
End Sub

' Log folder collection
Public Sub LogFolderCollection(folderPath As String, fileCount As Long, includeSubfolders As Boolean)
    LogMessage "Collecting files from: " & folderPath
    LogMessage "  Include subfolders: " & IIf(includeSubfolders, "Yes", "No")
    LogMessage "  Files found: " & fileCount
End Sub

' Log filter results (for checked-out only mode)
Public Sub LogFilterResults(totalBefore As Long, totalAfter As Long, filterType As String)
    LogMessage "Filter applied: " & filterType
    LogMessage "  Before filter: " & totalBefore & " files"
    LogMessage "  After filter: " & totalAfter & " files"
    LogMessage "  Filtered out: " & (totalBefore - totalAfter) & " files"
End Sub
