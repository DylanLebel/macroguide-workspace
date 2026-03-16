Attribute VB_Name = "Module2"
' ====================================================================================
' ConfigConstants.bas - Centralized Configuration Constants
' ====================================================================================
'
' PURPOSE:
'   Single location for all configurable paths and settings.
'   Update these values when moving between environments.
'
' ENVIRONMENT VARIABLES (override constants):
'   PDM_ROOT           - Override PDM_VAULT_ROOT
'   CADLINK_LOG_LEVEL  - OFF, ERROR, WARN, INFO, DEBUG
'   CADLINK_LOG_TO_FILE - TRUE/FALSE
'   TEMP               - Used for log file fallback
'
' ====================================================================================
Option Explicit
Option Private Module
' ===================
' PDM Vault Paths
' ===================
' These can be overridden by environment variables for portability
Public Const PDM_VAULT_ROOT_DEFAULT As String = "C:\NMT_PDM\"
Public Const PDM_LIBRARIES_SUBFOLDER As String = "Libraries\"
Public Const PDM_MACRO_SUBFOLDER As String = "Libraries\Macro\Prepare for CADLink.swp"
Public Const PDM_MATERIAL_SUBFOLDER As String = "Libraries\Material\NMT MATERIALS.sldmat"
' ===================
' Debug Settings
' ===================
' Set to False for production use - significantly improves performance
Public Const DEBUG_MODE As Boolean = False
' Set to True to write logs to file (slower but persistent)
Public Const LOG_TO_FILE As Boolean = True
' Set to True to copy local logs to a shared directory for central troubleshooting
Public Const ENABLE_SHARED_LOGGING As Boolean = True

' ===================
' Performance Settings
' ===================
' ENABLE_LOGGING: Master switch for all debug logging
' Set to False for maximum performance (zero overhead from log calls)
' Set to True when you need to troubleshoot issues
Public Const ENABLE_LOGGING As Boolean = True
' ENABLE_VERBOSE_LOGGING: Controls detailed/verbose log messages
' Only has effect when ENABLE_LOGGING is True
Public Const ENABLE_VERBOSE_LOGGING As Boolean = True
' ===================
' Flat Bar Settings
' ===================
' Set to True to completely disable flat bar conversion prompts
Public Const DISABLE_FLAT_BAR_PROMPTS As Boolean = False
' ===================
' Output Paths
' ===================
Public Const PDF_OUTPUT_PATH As String = "C:\Output\"
' ===================
' PATH ACCESSOR FUNCTIONS
' ===================
' Use these functions instead of constants for environment-aware paths
Public Function GetPdmVaultRoot() As String
    Dim envRoot As String
    envRoot = Environ("PDM_ROOT")
    If envRoot <> "" Then
        ' Ensure trailing backslash
        If Right(envRoot, 1) <> "\" Then envRoot = envRoot & "\"
        GetPdmVaultRoot = envRoot
    Else
        GetPdmVaultRoot = PDM_VAULT_ROOT_DEFAULT
    End If
End Function
Public Function GetPdmLibrariesPath() As String
    GetPdmLibrariesPath = GetPdmVaultRoot() & PDM_LIBRARIES_SUBFOLDER
End Function
Public Function GetPdmMacroPath() As String
    GetPdmMacroPath = GetPdmVaultRoot() & PDM_MACRO_SUBFOLDER
End Function
Public Function GetPdmMaterialDbPath() As String
    GetPdmMaterialDbPath = GetPdmVaultRoot() & PDM_MATERIAL_SUBFOLDER
End Function
' ===================
' Logging Paths
' ===================
' Helper to ensure a directory structure exists
Private Sub CreateFolderStructure(ByVal path As String)
    On Error Resume Next
    Dim folders() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Remove trailing backslash for splitting
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    
    folders = Split(path, "\")
    currentPath = folders(0)
    
    For i = 1 To UBound(folders)
        currentPath = currentPath & "\" & folders(i)
        If Dir(currentPath, vbDirectory) = "" Then
            MkDir currentPath
        End If
    Next i
    On Error GoTo 0
End Sub

Public Function GetLogBaseDir() As String
    On Error Resume Next
    Dim networkPath As String
    networkPath = "Y:\Solidworks\Macros\Macro Data PDM\Prepareforcadlink\Logs\"
    
    ' Check if network path is accessible, fallback to temp if not
    If Dir(Left(networkPath, InStrRev(networkPath, "\", Len(networkPath) - 1)), vbDirectory) = "" Then
        Dim tempPath As String
        tempPath = Environ("TEMP")
        If tempPath = "" Then tempPath = "C:\Temp"
        If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
        GetLogBaseDir = tempPath
    Else
        GetLogBaseDir = networkPath
    End If
    On Error GoTo 0
End Function

Public Function GetUserLogFolder() As String
    Dim baseDir As String
    baseDir = GetLogBaseDir()
    
    Dim username As String
    username = Environ("USERNAME")
    If username = "" Then username = "UnknownUser"
    
    ' Organize by User \ Year \ Month
    Dim userFolder As String
    userFolder = baseDir & username & "\" & Format(Date, "yyyy") & "\" & Format(Date, "mm") & "\"
    
    ' Ensure the folder structure exists
    CreateFolderStructure userFolder
    
    GetUserLogFolder = userFolder
End Function

' Dynamic paths that work for any user
Public Function GetDebugLogPath() As String
    Dim userFolder As String
    userFolder = GetUserLogFolder()
    
    Dim dateStr As String
    dateStr = Format(Date, "yyyy-mm-dd")
    
    GetDebugLogPath = userFolder & "CADLink_" & dateStr & ".log"
End Function

Public Function GetSummaryLogPath() As String
    Dim userFolder As String
    userFolder = GetUserLogFolder()
    
    Dim dateStr As String
    dateStr = Format(Date, "yyyy-mm-dd")
    
    GetSummaryLogPath = userFolder & "Summary_" & dateStr & ".txt"
End Function
' Legacy constants for backward compatibility
' These now call the functions above
Public Function PDM_VAULT_ROOT() As String
    PDM_VAULT_ROOT = GetPdmVaultRoot()
End Function
Public Function PDM_LIBRARIES_PATH() As String
    PDM_LIBRARIES_PATH = GetPdmLibrariesPath()
End Function
Public Function PDM_MACRO_PATH() As String
    PDM_MACRO_PATH = GetPdmMacroPath()
End Function
Public Function PDM_MATERIAL_DB_PATH() As String
    PDM_MATERIAL_DB_PATH = GetPdmMaterialDbPath()
End Function
Public Function DEBUG_LOG_PATH() As String
    DEBUG_LOG_PATH = GetDebugLogPath()
End Function
Public Function SUMMARY_LOG_PATH() As String
    SUMMARY_LOG_PATH = GetSummaryLogPath()
End Function
Public Function GetSharedLogDir() As String
    ' Use a central location in PDM for shared logs
    ' Everyone needs write access to this folder
    GetSharedLogDir = "Y:\Solidworks\Macros\Macro Data PDM\Logs\"
End Function
Public Function GetModuleExportPath() As String
    GetModuleExportPath = "C:\Users\dlebel\Downloads\prepareforcadlinkmacro\feb25\"
End Function
Public Function MODULE_EXPORT_PATH() As String
    MODULE_EXPORT_PATH = GetModuleExportPath()
End Function
' ===================
' Centralized Debug Logging (Legacy Wrappers)
' ===================
' These now delegate to Logger.bas for backward compatibility
' New code should use Logger.LogInfo, Logger.LogDebug, etc.
' DebugLog/DebugLogFunc wrappers are provided by Logger.bas.
