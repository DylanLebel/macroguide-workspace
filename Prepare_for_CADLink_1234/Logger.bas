Attribute VB_Name = "Logger"
' ====================================================================================
' Logger.bas - Centralized Logging Module (OPTIMIZED)
' ====================================================================================
' OPTIMIZATION NOTES:
'   - Added ENABLE_LOGGING check at start of every public method (zero overhead when off)
'   - Cached log level to avoid repeated Environ() calls
'   - Batch file writing with persistent file handle during session
'   - Early exit patterns throughout
' ====================================================================================
Option Explicit
Option Private Module

' Log level constants
Public Enum LogLevel
    LOG_OFF = 0
    LOG_ERROR = 1
    LOG_WARN = 2
    LOG_INFO = 3
    LOG_DEBUG = 4
End Enum

' Module-level state
Private m_Initialized As Boolean
Private m_LogFilePath As String
Private m_LogFileNum As Integer
Private m_SessionStartTime As Date
Private m_FileHandleOpen As Boolean

' Cached values for performance
Private m_CachedLogLevel As LogLevel
Private m_LogLevelCached As Boolean
Private m_CachedLogToFile As Boolean
Private m_LogToFileCached As Boolean

' ====================================================================================
' PUBLIC LOGGING METHODS
' ====================================================================================

Public Sub LogError(message As String, Optional errNum As Long = 0, Optional errDesc As String = "")
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < LOG_ERROR Then Exit Sub

    Dim fullMsg As String
    If errNum <> 0 Then
        fullMsg = "[ERROR] " & message & " (Error #" & errNum & ": " & errDesc & ")"
    Else
        fullMsg = "[ERROR] " & message
    End If

    WriteLog fullMsg
End Sub

Public Sub LogWarn(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < LOG_WARN Then Exit Sub
    WriteLog "[WARN]  " & message
End Sub

Public Sub LogInfo(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < LOG_INFO Then Exit Sub
    WriteLog "[INFO]  " & message
End Sub

Public Sub LogDebug(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < LOG_DEBUG Then Exit Sub
    WriteLog "[DEBUG] " & message
End Sub

' Log with explicit function name prefix (useful for tracing)
Public Sub LogFunc(funcName As String, message As String, Optional level As LogLevel = LOG_DEBUG)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < level Then Exit Sub
    WriteLog "[" & GetLevelName(level) & "] [" & funcName & "] " & message
End Sub

' Log section headers for readability
Public Sub LogSection(sectionName As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    If GetLogLevelCached() < LOG_INFO Then Exit Sub
    WriteLog String(60, "=")
    WriteLog "=== " & sectionName & " ==="
    WriteLog String(60, "=")
End Sub

' ====================================================================================
' SESSION MANAGEMENT
' ====================================================================================

Public Sub StartSession(Optional sessionName As String = "Macro Session")
    ' Reset cached values at session start
    m_LogLevelCached = False
    m_LogToFileCached = False

    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub

    m_SessionStartTime = Now
    InitializeLogFile

    WriteLog String(70, "=")
    WriteLog "  " & sessionName
    WriteLog "  Started: " & Format(m_SessionStartTime, "yyyy-mm-dd hh:nn:ss")
    WriteLog "  Log Level: " & GetLevelName(GetLogLevelCached())
    WriteLog "  User: " & GetCurrentUser()
    WriteLog "  Logging Enabled: " & ENABLE_LOGGING
    WriteLog String(70, "=")
    WriteLog ""
End Sub

Public Sub EndSession(Optional modelCount As Long = 0)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then
        ' Still close file handle if it was somehow opened
        CloseLogFile
        Exit Sub
    End If

    Dim elapsed As Double
    elapsed = (Now - m_SessionStartTime) * 86400 ' Convert to seconds

    WriteLog ""
    WriteLog String(70, "=")
    WriteLog "  SESSION COMPLETE"
    If modelCount > 0 Then
        WriteLog "  Models Processed: " & modelCount
    End If
    WriteLog "  Duration: " & FormatElapsed(elapsed)
    WriteLog "  Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    WriteLog String(70, "=")

    CloseLogFile

    ' Sync to shared folder for central troubleshooting
    If ENABLE_SHARED_LOGGING Then
        SyncLogToSharedFolder
    End If
End Sub

' ====================================================================================
' SHARED LOGGING SYNC
' ====================================================================================

Private Sub SyncLogToSharedFolder()
    On Error Resume Next

    Dim localPath As String
    Dim sharedDir As String
    Dim sharedPath As String
    Dim fileName As String
    Dim fso As Object

    localPath = GetLogFilePath()
    If Dir(localPath) = "" Then Exit Sub

    sharedDir = GetSharedLogDir()

    ' Ensure shared directory exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sharedDir) Then
        ' Create folder hierarchy if needed
        CreateFolderRecursive sharedDir
    End If

    ' Format: YYYY-MM-DD_HHMM_Username.log
    fileName = Format(m_SessionStartTime, "yyyy-mm-dd_hhnn") & "_" & GetCurrentUser() & ".log"
    sharedPath = sharedDir & fileName

    ' Copy local log to shared location
    fso.CopyFile localPath, sharedPath, True

    On Error GoTo 0
End Sub

Private Sub CreateFolderRecursive(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(path) Then Exit Sub

    Dim parent As String
    parent = fso.GetParentFolderName(path)

    If Not fso.FolderExists(parent) Then
        CreateFolderRecursive parent
    End If

    fso.CreateFolder path
End Sub

' ====================================================================================
' UTILITY METHODS
' ====================================================================================

Public Function GetLogFilePath() As String
    If m_LogFilePath = "" Then
        m_LogFilePath = GetDefaultLogPath()
    End If
    GetLogFilePath = m_LogFilePath
End Function

Public Sub SetLogFilePath(path As String)
    m_LogFilePath = path
End Sub

' Force refresh of cached settings (call if you change config mid-session)
Public Sub RefreshLogSettings()
    m_LogLevelCached = False
    m_LogToFileCached = False
End Sub

' ====================================================================================
' PRIVATE HELPER METHODS
' ====================================================================================

Private Sub WriteLog(message As String)
    Dim timestamp As String
    timestamp = Format(Now, "hh:nn:ss") & " "

    ' Always write to Immediate window when logging enabled
    Debug.Print timestamp & message

    ' Write to file if enabled (using cached value)
    If GetLogToFileCached() Then
        WriteToFileBatched timestamp & message
    End If
End Sub

' OPTIMIZATION: Batch file writing - keeps file handle open during session
Private Sub WriteToFileBatched(message As String)
    On Error Resume Next

    ' Open file handle if not already open
    If Not m_FileHandleOpen Then
        If Not m_Initialized Then
            InitializeLogFile
        End If

        m_LogFileNum = FreeFile
        Open GetLogFilePath() For Append As #m_LogFileNum
        m_FileHandleOpen = True
    End If

    ' Write to open file handle
    If m_FileHandleOpen Then
        Print #m_LogFileNum, message
    End If

    On Error GoTo 0
End Sub

' Legacy single-write method (kept for compatibility but not used)
Private Sub WriteToFile(message As String)
    On Error Resume Next

    If Not m_Initialized Then
        InitializeLogFile
    End If

    Dim fileNum As Integer
    fileNum = FreeFile

    Open GetLogFilePath() For Append As #fileNum
    Print #fileNum, message
    Close #fileNum

    On Error GoTo 0
End Sub

Private Sub InitializeLogFile()
    On Error Resume Next

    Dim logPath As String
    logPath = GetLogFilePath()

    ' Ensure directory exists
    Dim folderPath As String
    folderPath = Left(logPath, InStrRev(logPath, "\"))

    ' Recursive MkDir not available in standard VBA, but we can try to create the leaf folder
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    ' Ensure file exists without clearing it (Append creates if not exists)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Close #fileNum

    m_Initialized = True

    On Error GoTo 0
End Sub

Private Sub CloseLogFile()
    On Error Resume Next

    ' Close batch file handle if open
    If m_FileHandleOpen Then
        Close #m_LogFileNum
        m_FileHandleOpen = False
    End If

    m_Initialized = False

    On Error GoTo 0
End Sub

' OPTIMIZATION: Cached log level to avoid repeated Environ() calls
Private Function GetLogLevelCached() As LogLevel
    If m_LogLevelCached Then
        GetLogLevelCached = m_CachedLogLevel
        Exit Function
    End If

    ' Calculate and cache
    m_CachedLogLevel = GetLogLevel()
    m_LogLevelCached = True
    GetLogLevelCached = m_CachedLogLevel
End Function

Private Function GetLogLevel() As LogLevel
    ' Check environment variable first
    Dim envLevel As String
    envLevel = Environ("CADLINK_LOG_LEVEL")

    If envLevel <> "" Then
        Select Case UCase(envLevel)
            Case "OFF", "0": GetLogLevel = LOG_OFF
            Case "ERROR", "1": GetLogLevel = LOG_ERROR
            Case "WARN", "2": GetLogLevel = LOG_WARN
            Case "INFO", "3": GetLogLevel = LOG_INFO
            Case "DEBUG", "4": GetLogLevel = LOG_DEBUG
            Case Else: GetLogLevel = LOG_INFO
        End Select
        Exit Function
    End If

    ' Fall back to ConfigConstants
    On Error Resume Next
    If DEBUG_MODE Then
        GetLogLevel = LOG_DEBUG
    Else
        GetLogLevel = LOG_INFO
    End If
    On Error GoTo 0
End Function

' OPTIMIZATION: Cached log-to-file setting
Private Function GetLogToFileCached() As Boolean
    If m_LogToFileCached Then
        GetLogToFileCached = m_CachedLogToFile
        Exit Function
    End If

    ' Calculate and cache
    m_CachedLogToFile = GetLogToFile()
    m_LogToFileCached = True
    GetLogToFileCached = m_CachedLogToFile
End Function

Private Function GetLogToFile() As Boolean
    ' Check environment variable first
    Dim envLogToFile As String
    envLogToFile = Environ("CADLINK_LOG_TO_FILE")

    If envLogToFile <> "" Then
        GetLogToFile = (UCase(envLogToFile) = "TRUE" Or envLogToFile = "1")
        Exit Function
    End If

    ' Fall back to ConfigConstants setting
    On Error Resume Next
    GetLogToFile = LOG_TO_FILE
    If Err.Number <> 0 Then
        GetLogToFile = False
    End If
    On Error GoTo 0
End Function

Private Function GetDefaultLogPath() As String
    ' Try to use ConfigConstants path first
    On Error Resume Next
    GetDefaultLogPath = DEBUG_LOG_PATH
    On Error GoTo 0

    ' Fallback to temp directory
    If GetDefaultLogPath = "" Then
        GetDefaultLogPath = Environ("TEMP") & "\CADLink_Macro.log"
    End If
End Function

Private Function GetCurrentUser() As String
    GetCurrentUser = Environ("USERNAME")
    If GetCurrentUser = "" Then GetCurrentUser = "Unknown"
End Function

Private Function GetLevelName(level As LogLevel) As String
    Select Case level
        Case LOG_OFF: GetLevelName = "OFF"
        Case LOG_ERROR: GetLevelName = "ERROR"
        Case LOG_WARN: GetLevelName = "WARN"
        Case LOG_INFO: GetLevelName = "INFO"
        Case LOG_DEBUG: GetLevelName = "DEBUG"
        Case Else: GetLevelName = "UNKNOWN"
    End Select
End Function

Private Function FormatElapsed(seconds As Double) As String
    If seconds < 60 Then
        FormatElapsed = Format(seconds, "0.0") & " seconds"
    ElseIf seconds < 3600 Then
        FormatElapsed = Format(seconds / 60, "0.0") & " minutes"
    Else
        FormatElapsed = Format(seconds / 3600, "0.0") & " hours"
    End If
End Function

' ====================================================================================
' BACKWARD COMPATIBILITY WRAPPERS
' ====================================================================================
' These allow existing code to work without immediate refactoring
' All have early-exit when ENABLE_LOGGING is False

Public Sub DebugLog(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogDebug message
End Sub

Public Sub DebugLogFunc(funcName As String, message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogFunc funcName, message, LOG_DEBUG
End Sub

Public Sub DebugPrint(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogDebug message
End Sub

Public Sub DebugPrintH(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogInfo message
End Sub

Public Sub DebugPrintE(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogError message
End Sub

Public Sub LogToFile(message As String)
    ' OPTIMIZATION: Early exit if logging disabled
    If Not ENABLE_LOGGING Then Exit Sub
    LogDebug message
End Sub
