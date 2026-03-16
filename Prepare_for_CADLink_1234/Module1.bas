Attribute VB_Name = "Module1"
' ====================================================================================
' ExportModules.bas
' Utility module for exporting/importing all VBA modules, classes, and forms
' Used for version control and backup purposes
' ====================================================================================
Option Explicit
Option Private Module
' Export all modules to disk
Sub ExportAllModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fileName As String
    Dim exportCount As Long

    On Error GoTo ErrHandler

    ' Set export folder - using centralized constant
    exportPath = MODULE_EXPORT_PATH()

    ' Create directory if it doesn't exist
    If Dir(exportPath, vbDirectory) = vbNullString Then MkDir exportPath

    Set vbProj = Application.VBE.ActiveVBProject
    exportCount = 0

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                fileName = exportPath & vbComp.Name & ".bas"
                vbComp.Export fileName
                exportCount = exportCount + 1

            Case vbext_ct_ClassModule
                fileName = exportPath & vbComp.Name & ".cls"
                vbComp.Export fileName
                exportCount = exportCount + 1

            Case vbext_ct_MSForm
                fileName = exportPath & vbComp.Name & ".frm"
                vbComp.Export fileName
                exportCount = exportCount + 1

            Case vbext_ct_Document
                ' These are ThisSW or ThisMacro - export but note they're document modules
                fileName = exportPath & vbComp.Name & ".cls"
                vbComp.Export fileName
                exportCount = exportCount + 1
        End Select
    Next vbComp

    MsgBox "Export complete!" & vbCrLf & _
           "Exported " & exportCount & " module(s) to:" & vbCrLf & exportPath, _
           vbInformation, "Export Successful"
    Exit Sub

ErrHandler:
    MsgBox "Error exporting modules: " & Err.description & vbCrLf & _
           "Error " & Err.Number, vbCritical, "Export Error"
End Sub
' Import all modules from disk
Sub ImportAllModules()
    Dim vbProj As VBIDE.VBProject
    Dim exportPath As String
    Dim fileName As String
    Dim fileExt As String
    Dim importCount As Long
    Dim file As String

    On Error GoTo ErrHandler

    exportPath = MODULE_EXPORT_PATH()

    ' Check if directory exists
    If Dir(exportPath, vbDirectory) = vbNullString Then
        MsgBox "Export folder not found:" & vbCrLf & exportPath, _
               vbExclamation, "Import Error"
        Exit Sub
    End If

    Set vbProj = Application.VBE.ActiveVBProject
    importCount = 0
    ' Import .bas files (Standard Modules)
    file = Dir(exportPath & "*.bas")
    Do While file <> vbNullString
        fileName = exportPath & file
        ' Remove existing module with same name to prevent duplicates
        Dim moduleName As String
        moduleName = Left(file, Len(file) - 4) ' Remove .bas extension
        
        ' CRITICAL FIX: Skip the management modules so they don't delete themselves while running
        If moduleName <> "Module1" And moduleName <> "Module2" And _
           moduleName <> "ExportModules" And moduleName <> "ConfigConstants" Then
            
            On Error Resume Next
            vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
            On Error GoTo ErrHandler
            
            vbProj.VBComponents.Import fileName
            importCount = importCount + 1
        End If
        
        file = Dir()
    Loop

        ' Skip ThisSW and ThisMacro (document modules)
        If LCase(file) <> "thissw.cls" And LCase(file) <> "thismacro.cls" And _
           LCase(file) <> "thislibrary.cls" And LCase(file) <> "thislibrary1.cls" Then
            
            fileName = exportPath & file
            ' Remove existing module with same name to prevent duplicates
            moduleName = Left(file, Len(file) - 4) ' Remove .cls extension
            
            ' Skip management modules
            If moduleName <> "Module1" And moduleName <> "Module2" And _
               moduleName <> "ExportModules" And moduleName <> "ConfigConstants" Then
                
                On Error Resume Next
                vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
                On Error GoTo ErrHandler
                
                vbProj.VBComponents.Import fileName
                importCount = importCount + 1
            End If
        End If

    ' Import .frm files (Forms)
    file = Dir(exportPath & "*.frm")
    Do While file <> vbNullString
        fileName = exportPath & file
        ' Remove existing form with same name to prevent duplicates
        moduleName = Left(file, Len(file) - 4) ' Remove .frm extension
        
        ' Skip management modules (unlikely for forms, but for safety)
        If moduleName <> "Module1" And moduleName <> "Module2" Then
            On Error Resume Next
            vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
            On Error GoTo ErrHandler
            
            vbProj.VBComponents.Import fileName
            importCount = importCount + 1
        End If
        
        file = Dir()
    Loop

    MsgBox "Import complete!" & vbCrLf & _
           "Imported " & importCount & " module(s) from:" & vbCrLf & exportPath, _
           vbInformation, "Import Successful"
    Exit Sub

ErrHandler:
    MsgBox "Error importing modules: " & Err.description & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           "Last file: " & file, vbCritical, "Import Error"
End Sub
' Remove all non-document modules
Sub RemoveAllModules()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim removeCount As Long
    Dim i As Long

    On Error GoTo ErrHandler

    ' Confirm with user
    If MsgBox("This will remove all modules, classes, and forms!" & vbCrLf & _
              "Document modules (ThisSW, ThisMacro) will be preserved." & vbCrLf & _
              "Utility modules (ExportModules, ConfigConstants) will be preserved." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbExclamation, "Confirm Remove All") = vbNo Then
        Exit Sub
    End If

    Set vbProj = Application.VBE.ActiveVBProject
    removeCount = 0

    ' Loop backwards to avoid index issues when removing
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set vbComp = vbProj.VBComponents.item(i)
        ' Only remove Standard Modules, Class Modules, and Forms
        ' Do NOT remove Document modules (ThisSW, ThisMacro, etc.)
        ' Do NOT remove ExportModules or ConfigConstants (utility modules)
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                ' Skip current refresh module and configuration module
                If vbComp.Name <> "Module1" And vbComp.Name <> "Module2" And _
                   vbComp.Name <> "ExportModules" And vbComp.Name <> "ConfigConstants" Then
                    vbProj.VBComponents.Remove vbComp
                    removeCount = removeCount + 1
                End If
        End Select
    Next i

    MsgBox "Removed " & removeCount & " module(s).", _
           vbInformation, "Remove Complete"
    Exit Sub

ErrHandler:
    MsgBox "Error removing modules: " & Err.description & vbCrLf & _
           "Error " & Err.Number, vbCritical, "Remove Error"
End Sub
' Refresh: Remove all modules then import from disk
Sub RefreshAllModules()
    Dim response As VbMsgBoxResult

    On Error GoTo ErrHandler

    ' Confirm with user
    response = MsgBox("This will:" & vbCrLf & _
                     "1. Remove all existing modules (except document modules)" & vbCrLf & _
                     "2. Import fresh copies from disk" & vbCrLf & vbCrLf & _
                     "Make sure you've exported recently!" & vbCrLf & vbCrLf & _
                     "Continue?", _
                     vbYesNo + vbQuestion, "Confirm Refresh")

    If response = vbNo Then Exit Sub

    ' Remove all modules
    Call RemoveAllModules

    ' Import all modules
    Call ImportAllModules

    MsgBox "Refresh complete!", vbInformation, "Success"
    Exit Sub

ErrHandler:
    MsgBox "Error during refresh: " & Err.description & vbCrLf & _
           "Error " & Err.Number, vbCritical, "Refresh Error"
End Sub
