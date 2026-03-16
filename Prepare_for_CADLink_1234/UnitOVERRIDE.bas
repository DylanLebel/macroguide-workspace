Attribute VB_Name = "UnitOVERRIDE"
Option Explicit
Option Private Module

' ========================================
' PATH-BASED UNIT OVERRIDE CONFIGURATION
' ========================================

Private Const FORCE_MM_PATHS As String = _
    "\NMT_PDM\Projects\Warranty\26896 - PTFI - Unloading Station Maintenance Platform|"
    '"\NMT_PDM\Projects\Capital\8546 - TestDL - testing\3 - Design\Models"
    ' "\Projects\MetricJob2|" & _
    ' "\\ServerName\SharedFolder\MetricProject"

' List of folder paths that should force dimensions to inches regardless of model units
Private Const FORCE_INCH_PATHS As String = _
"\NMT_PDM\Projects\Capital\26381 - BHP Olympic Dam Chutes|"

   ' "C:\Projects\ImperialJob1|" & _
   ' "C:\Projects\ImperialJob2|" & _
   ' "\\ServerName\SharedFolder\USProject"

' Function to check if current model path matches any special paths
Function GetPathBasedUnitOverride(model As ModelDoc2) As String
    ' Returns: "mm" = force millimeters, "in" = force inches, "" = no override
    
    If model Is Nothing Then
        GetPathBasedUnitOverride = ""
        Exit Function
    End If
    
    Dim modelPath As String
    modelPath = model.GetPathName
    
    If modelPath = "" Then
        DebugLog "Model not saved - no path-based override"
        GetPathBasedUnitOverride = ""
        Exit Function
    End If
    
    DebugLog "=== Checking Path-Based Unit Override ==="
    DebugLog "Model Path: " & modelPath
    
    ' Check FORCE_MM_PATHS
    Dim mmPaths() As String
    mmPaths = Split(FORCE_MM_PATHS, "|")
    
    Dim i As Integer
    For i = 0 To UBound(mmPaths)
        If mmPaths(i) <> "" Then
            If InStr(1, modelPath, mmPaths(i), vbTextCompare) > 0 Then
                DebugLog "MATCH FOUND: Forcing MM for path: " & mmPaths(i)
                GetPathBasedUnitOverride = "mm"
                Exit Function
            End If
        End If
    Next i
    
    ' Check FORCE_INCH_PATHS
    Dim inchPaths() As String
    inchPaths = Split(FORCE_INCH_PATHS, "|")
    
    For i = 0 To UBound(inchPaths)
        If inchPaths(i) <> "" Then
            If InStr(1, modelPath, inchPaths(i), vbTextCompare) > 0 Then
                DebugLog "MATCH FOUND: Forcing INCHES for path: " & inchPaths(i)
                GetPathBasedUnitOverride = "in"
                Exit Function
            End If
        End If
    Next i
    
    DebugLog "No path match - using standard logic"
    GetPathBasedUnitOverride = ""
End Function

