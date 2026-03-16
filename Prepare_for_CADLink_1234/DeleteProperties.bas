Attribute VB_Name = "DeleteProperties"
Option Explicit
Option Private Module

' =========================================================
' CleanDeleteProperties Module with Deep Diagnostics
' Updated with improved deletion logic
' =========================================================

' =========================
' Debug toggles - All controlled by DEBUG_MODE in ConfigConstants
' =========================
Public Const DEBUG_TRACE_OVERVIEW As Boolean = DEBUG_MODE          ' High level progress
Public Const DEBUG_TRACE_SCOPES As Boolean = DEBUG_MODE            ' Document, configs, cut lists discovered
Public Const DEBUG_TRACE_VALUES As Boolean = DEBUG_MODE            ' Show Get6 outputs for each examined property
Public Const DEBUG_TRACE_ASCII As Boolean = False                  ' Per character ASCII for names to reveal odd chars
Public Const DEBUG_TRACE_NAMES As Boolean = DEBUG_MODE             ' Dump full property name lists per scope
Public Const DEBUG_TRACE_SUMMARY As Boolean = DEBUG_MODE           ' One page outcome summary
Public Const DEBUG_TRACE_TIMING As Boolean = DEBUG_MODE            ' Simple timing stamps
Public Const DEBUG_TRACE_VERIFICATION As Boolean = DEBUG_MODE      ' Verify deletion worked

' =========================
' Configuration - Properties to delete
' =========================
Function GetPropertiesToDelete() As Object
    Dim props As Object
    Set props = CreateObject("Scripting.Dictionary")
    
    ' Always delete these properties regardless of content (False = always delete)
    props.Add "Project", False
    props.Add "StockSize", False
    props.Add "Purchased Assembly", False
    props.Add "Surface Treatment", False
    'props.Add "Mfg Comments", False
    props.Add "LengthWidthHeightUM", False
    props.Add "DrawnBy", False
    props.Add "Reference Catergory", False
    props.Add "stocksize", False           ' lowercase version
    props.Add "stockSize", False           ' camelCase version
    props.Add "Stocksize", False           ' Title case version
    props.Add "3d-bounding box thickness", False
    props.Add "3d-bounding box length", False
    props.Add "Total Length", False
    props.Add "DRAWING", False
    props.Add "GRADE", False
    props.Add "PURCHASED ASSEMBLY", False
    
    
    Set GetPropertiesToDelete = props
End Function

' =========================
' Entry point used by your pipeline
' =========================
Sub RemoveProperties(swModel As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    Dim t0 As Double: t0 = Timer
    
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "=== RemoveProperties start ==="
    If DEBUG_TRACE_TIMING Then DebugPrint "Start time: " & Format$(Now, "yyyy mm dd hh:nn:ss")
    
    ' Initial visibility checks
    If DEBUG_TRACE_OVERVIEW Then
        DebugPrintH "Initial dictionary probes for stock variants"
        DebugPrint "Has StockSize: " & propertiesToSet.exists("StockSize")
        DebugPrint "Has stocksize: " & propertiesToSet.exists("stocksize")
        DebugPrint "Has stockSize: " & propertiesToSet.exists("stockSize")
        DebugPrint "Has Stocksize: " & propertiesToSet.exists("Stocksize")
        DebugPrint "Has Reference Catergory: " & propertiesToSet.exists("Reference Catergory")
    End If
    
    Dim custPropMgr As customPropertyManager
    Set custPropMgr = swModel.Extension.customPropertyManager("")
    
    If DEBUG_TRACE_NAMES Or DEBUG_TRACE_VALUES Then
        DumpScopeProperties "Document", custPropMgr
    End If
    
    If DEBUG_TRACE_SCOPES Then
        DebugPrintH "Enumerating configurations"
    End If
    
    Dim vCfgs As Variant, i As Long
    vCfgs = swModel.GetConfigurationNames
    If Not IsEmpty(vCfgs) Then
        For i = LBound(vCfgs) To UBound(vCfgs)
            Dim cfgName As String
            cfgName = CStr(vCfgs(i))
            If DEBUG_TRACE_SCOPES Then DebugPrint "Config found: [" & cfgName & "]"
            If DEBUG_TRACE_NAMES Or DEBUG_TRACE_VALUES Then
                Dim cfgMgr As customPropertyManager
                Set cfgMgr = swModel.Extension.customPropertyManager(cfgName)
                DumpScopeProperties "Config " & cfgName, cfgMgr
            End If
        Next i
    ElseIf DEBUG_TRACE_SCOPES Then
        DebugPrint "No configurations returned by GetConfigurationNames"
    End If
    
    If swModel.GetType = swDocPART Then
        If DEBUG_TRACE_SCOPES Then DebugPrintH "Enumerating cut list folders"
        Dim f As Feature
        Set f = swModel.FirstFeature
        Do While Not f Is Nothing
            If DEBUG_TRACE_SCOPES Then DebugPrint "Feature " & f.Name & " type " & f.GetTypeName2
            If f.GetTypeName2 = "CutListFolder" Then
                Dim clMgr As customPropertyManager
                Set clMgr = f.customPropertyManager
                If DEBUG_TRACE_NAMES Or DEBUG_TRACE_VALUES Then
                    DumpScopeProperties "CutList " & f.Name, clMgr
                End If
            End If
            Set f = f.GetNextFeature
        Loop
    End If
    
    Dim propsToDelete As Object
    Set propsToDelete = GetPropertiesToDelete()
    
    If DEBUG_TRACE_OVERVIEW Then
        DebugPrintH "Dictionary before deletion count " & propertiesToSet.Count
    End If
    
    ' Delete from in memory dictionary first
    Dim propName As Variant
    For Each propName In propsToDelete.keys
        Dim deleteIfEmpty As Boolean
        deleteIfEmpty = propsToDelete(propName)
        
        If propertiesToSet.exists(propName) Then
            Dim propValue As String
            propValue = CStr(propertiesToSet(propName))
            Dim shouldDelete As Boolean
            shouldDelete = (Not deleteIfEmpty) Or (Len(Trim(propValue)) = 0)
            If shouldDelete Then
                propertiesToSet.Remove propName
                If DEBUG_TRACE_VALUES Then DebugPrint "Dict removed [" & CStr(propName) & "] value was [" & propValue & "]"
            Else
                If DEBUG_TRACE_VALUES Then DebugPrint "Dict kept [" & CStr(propName) & "] since deleteIfEmpty True and value not empty: [" & propValue & "]"
            End If
        Else
            If DEBUG_TRACE_VALUES Then DebugPrint "Dict did not contain [" & CStr(propName) & "]"
        End If
    Next propName
    
    If DEBUG_TRACE_OVERVIEW Then
        DebugPrintH "Dictionary after deletion count " & propertiesToSet.Count
    End If
    
    ' Delete from document and all scopes
    DeletePropertiesFromDocument swModel, propsToDelete
    
    ' Final verification snapshot for the common names
    If DEBUG_TRACE_OVERVIEW Then
        DebugPrintH "Final verification probes"
        VerifyOneNameAcrossScopes swModel, "StockSize"
        VerifyOneNameAcrossScopes swModel, "stocksize"
        VerifyOneNameAcrossScopes swModel, "stockSize"
        VerifyOneNameAcrossScopes swModel, "Stocksize"
        VerifyOneNameAcrossScopes swModel, "Reference Catergory"
    End If
    
    If DEBUG_TRACE_SUMMARY Then
        DebugPrintH "RemoveProperties summary"
        DebugPrint "Remaining dict has StockSize: " & propertiesToSet.exists("StockSize")
        DebugPrint "Remaining dict has stocksize: " & propertiesToSet.exists("stocksize")
        DebugPrint "Remaining dict has Reference Catergory: " & propertiesToSet.exists("Reference Catergory")
    End If
    
    If DEBUG_TRACE_TIMING Then
        DebugPrint "Elapsed seconds: " & Format$(Timer - t0, "0.000")
    End If
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "=== RemoveProperties end ==="
End Sub

' =========================
' Round Bar specific
' =========================
Sub DeleteRoundBarProperties(swModel As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    Dim rbProps As Object
    Set rbProps = CreateObject("Scripting.Dictionary")
    rbProps.Add "Mtl Part Number", False
    
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "DeleteRoundBarProperties begin"
    
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
        If DEBUG_TRACE_VALUES Then DebugPrint "Dict removed [Mtl Part Number]"
    Else
        If DEBUG_TRACE_VALUES Then DebugPrint "Dict not found [Mtl Part Number]"
    End If
    
    DeletePropertiesFromDocument swModel, rbProps
    
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "DeleteRoundBarProperties end"
End Sub

' =========================
' Core deletion across scopes
' =========================
Sub DeletePropertiesFromDocument(swModel As ModelDoc2, propertiesToDelete As Object)
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "DeletePropertiesFromDocument begin"
    
    Dim propName As Variant
    For Each propName In propertiesToDelete.keys
        Dim deleteIfEmpty As Boolean
        deleteIfEmpty = propertiesToDelete(propName)
        If DEBUG_TRACE_OVERVIEW Then DebugPrintH "Process name [" & CStr(propName) & "] deleteIfEmpty " & deleteIfEmpty
        
        ' Document level
        Dim docResult As Integer
        docResult = DeletePropertyFromLocation(swModel, "", CStr(propName), deleteIfEmpty)
        If DEBUG_TRACE_OVERVIEW Then DebugPrint "Document deletion result: " & docResult
        
        ' Configuration level
        Dim vCfgs As Variant, i As Long
        vCfgs = swModel.GetConfigurationNames
        If Not IsEmpty(vCfgs) Then
            For i = LBound(vCfgs) To UBound(vCfgs)
                Dim cfgResult As Integer
                cfgResult = DeletePropertyFromLocation(swModel, CStr(vCfgs(i)), CStr(propName), deleteIfEmpty)
                If DEBUG_TRACE_OVERVIEW Then DebugPrint "Config " & CStr(vCfgs(i)) & " deletion result: " & cfgResult
            Next i
        End If
        
        ' REMOVED: Cut list level deletion - only delete from model custom properties
        ' Cut lists should retain their properties
        
    Next propName
    
    If DEBUG_TRACE_OVERVIEW Then DebugPrintH "DeletePropertiesFromDocument end"
End Sub

' =========================
' Document or Configuration deletion with improved logic
' =========================
Function DeletePropertyFromLocation(swModel As ModelDoc2, configName As String, propName As String, deleteIfEmpty As Boolean) As Integer
    Dim scopeLabel As String
    scopeLabel = IIf(configName = "", "Document", "Config " & configName)
    
    On Error GoTo ErrorHandler
    
    Dim mgr As customPropertyManager
    Set mgr = swModel.Extension.customPropertyManager(configName)
    If mgr Is Nothing Then
        If DEBUG_TRACE_VALUES Then DebugPrintE "CustomPropertyManager is Nothing for " & scopeLabel
        DeletePropertyFromLocation = -1
        Exit Function
    End If
    
    ' First check if property exists in the names list
    Dim names As Variant
    Dim propertyExists As Boolean
    propertyExists = False
    
    names = mgr.GetNames
    If Not IsEmpty(names) Then
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If StrComp(CStr(names(i)), propName, vbBinaryCompare) = 0 Then
                propertyExists = True
                Exit For
            End If
        Next i
    End If
    
    If Not propertyExists Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " property not found in names list [" & propName & "]"
        DeletePropertyFromLocation = 0
        Exit Function
    End If
    
    ' Get the property value to check if we should delete it
    Dim val As String
    Dim resolved As String
    Dim wasResolved As Boolean
    Dim linkTo As Boolean
    Dim ret As Long
    
    ret = mgr.Get6(propName, False, val, resolved, wasResolved, linkTo)
    
    If DEBUG_TRACE_VALUES Then
        DebugPrint scopeLabel & " Get6 ret=" & ret & " wasResolved=" & wasResolved & " val=[" & val & "] resolved=[" & resolved & "]"
    End If
    
    ' Determine if we should delete based on the deleteIfEmpty flag
    Dim shouldDelete As Boolean
    If deleteIfEmpty Then
        ' Only delete if the resolved value is empty or whitespace
        shouldDelete = (Len(Trim(resolved)) = 0)
    Else
        ' Always delete regardless of value
        shouldDelete = True
    End If
    
    If Not shouldDelete Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " kept [" & propName & "] - deleteIfEmpty=True and value not empty: [" & resolved & "]"
        DeletePropertyFromLocation = 0
        Exit Function
    End If
    
    ' Attempt to delete the property
    Dim deleteResult As Boolean
    deleteResult = mgr.Delete2(propName)
    
    If deleteResult Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " Delete2 succeeded for [" & propName & "]"
        
        ' Verify deletion worked
        If DEBUG_TRACE_VERIFICATION Then
            Dim verifyNames As Variant
            verifyNames = mgr.GetNames
            Dim stillExists As Boolean
            stillExists = False
            If Not IsEmpty(verifyNames) Then
                Dim j As Long
                For j = LBound(verifyNames) To UBound(verifyNames)
                    If StrComp(CStr(verifyNames(j)), propName, vbBinaryCompare) = 0 Then
                        stillExists = True
                        Exit For
                    End If
                Next j
            End If
            If stillExists Then
                DebugPrintE scopeLabel & " VERIFICATION FAILED: [" & propName & "] still exists after Delete2 reported success"
                DeletePropertyFromLocation = -1
            Else
                DebugPrint scopeLabel & " VERIFICATION SUCCESS: [" & propName & "] confirmed deleted"
                DeletePropertyFromLocation = 1
            End If
        Else
            DeletePropertyFromLocation = 1
        End If
    Else
        If DEBUG_TRACE_VALUES Then DebugPrintE scopeLabel & " Delete2 returned False for [" & propName & "]"
        DeletePropertyFromLocation = -1
    End If
    
    Exit Function

ErrorHandler:
    DebugPrintE "Error in DeletePropertyFromLocation " & scopeLabel & " [" & propName & "]: " & Err.description
    DeletePropertyFromLocation = -1
End Function

' =========================
' Cut list folder deletion with improved logic
' =========================
Function DeleteCutlistProperty(swFeature As SldWorks.Feature, propName As String, deleteIfEmpty As Boolean) As Integer
    Dim scopeLabel As String
    scopeLabel = "CutList " & swFeature.Name
    
    On Error GoTo ErrorHandler
    
    Dim mgr As customPropertyManager
    Set mgr = swFeature.customPropertyManager
    If mgr Is Nothing Then
        If DEBUG_TRACE_VALUES Then DebugPrintE "CustomPropertyManager is Nothing for " & scopeLabel
        DeleteCutlistProperty = -1
        Exit Function
    End If
    
    ' Check if property exists in the names list
    Dim names As Variant
    Dim propertyExists As Boolean
    propertyExists = False
    
    names = mgr.GetNames
    If Not IsEmpty(names) Then
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If StrComp(CStr(names(i)), propName, vbBinaryCompare) = 0 Then
                propertyExists = True
                Exit For
            End If
        Next i
    End If
    
    If Not propertyExists Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " property not found in names list [" & propName & "]"
        DeleteCutlistProperty = 0
        Exit Function
    End If
    
    ' Get the property value
    Dim val As String
    Dim resolved As String
    Dim wasResolved As Boolean
    Dim linkTo As Boolean
    Dim ret As Long
    
    ret = mgr.Get6(propName, False, val, resolved, wasResolved, linkTo)
    
    If DEBUG_TRACE_VALUES Then
        DebugPrint scopeLabel & " Get6 ret=" & ret & " wasResolved=" & wasResolved & " val=[" & val & "] resolved=[" & resolved & "]"
    End If
    
    ' Determine if we should delete
    Dim shouldDelete As Boolean
    If deleteIfEmpty Then
        shouldDelete = (Len(Trim(resolved)) = 0)
    Else
        shouldDelete = True
    End If
    
    If Not shouldDelete Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " kept [" & propName & "] - deleteIfEmpty=True and value not empty: [" & resolved & "]"
        DeleteCutlistProperty = 0
        Exit Function
    End If
    
    ' Delete the property
    Dim deleteResult As Boolean
    deleteResult = mgr.Delete2(propName)
    
    If deleteResult Then
        If DEBUG_TRACE_VALUES Then DebugPrint scopeLabel & " Delete2 succeeded for [" & propName & "]"
        
        ' Verify deletion worked
        If DEBUG_TRACE_VERIFICATION Then
            Dim verifyNames As Variant
            verifyNames = mgr.GetNames
            Dim stillExists As Boolean
            stillExists = False
            If Not IsEmpty(verifyNames) Then
                Dim j As Long
                For j = LBound(verifyNames) To UBound(verifyNames)
                    If StrComp(CStr(verifyNames(j)), propName, vbBinaryCompare) = 0 Then
                        stillExists = True
                        Exit For
                    End If
                Next j
            End If
            If stillExists Then
                DebugPrintE scopeLabel & " VERIFICATION FAILED: [" & propName & "] still exists after Delete2 reported success"
                DeleteCutlistProperty = -1
            Else
                DebugPrint scopeLabel & " VERIFICATION SUCCESS: [" & propName & "] confirmed deleted"
                DeleteCutlistProperty = 1
            End If
        Else
            DeleteCutlistProperty = 1
        End If
    Else
        If DEBUG_TRACE_VALUES Then DebugPrintE scopeLabel & " Delete2 returned False for [" & propName & "]"
        DeleteCutlistProperty = -1
    End If
    
    Exit Function

ErrorHandler:
    DebugPrintE "Error in DeleteCutlistProperty " & scopeLabel & " [" & propName & "]: " & Err.description
    DeleteCutlistProperty = -1
End Function

' =========================
' Diagnostics helpers
' =========================
Private Sub DumpScopeProperties(scopeName As String, mgr As customPropertyManager)
    On Error Resume Next
    If mgr Is Nothing Then
        DebugPrintE "DumpScopeProperties got null mgr for " & scopeName
        Exit Sub
    End If
    
    Dim names As Variant
    names = mgr.GetNames
    If IsEmpty(names) Then
        DebugPrint scopeName & " names list is empty"
        Exit Sub
    End If
    
    Dim i As Long
    DebugPrintH scopeName & " names count " & (UBound(names) - LBound(names) + 1)
    For i = LBound(names) To UBound(names)
        Dim nm As String
        nm = CStr(names(i))
        DebugPrint "  [" & nm & "] Len=" & Len(nm)
        If DEBUG_TRACE_ASCII Then
            PrintAscii "    ASCII " & scopeName & " [" & nm & "]", nm
        End If
        
        If DEBUG_TRACE_VALUES Then
            Dim v As String, r As String, wr As Boolean, l As Boolean
            Dim ret As Long
            ret = mgr.Get6(nm, False, v, r, wr, l)
            DebugPrint "    Get6 ret=" & ret & " wasResolved=" & wr & " val=[" & v & "] resolved=[" & r & "]"
        End If
    Next i
End Sub

Private Sub VerifyOneNameAcrossScopes(swModel As ModelDoc2, propName As String)
    Dim mgr As customPropertyManager
    Set mgr = swModel.Extension.customPropertyManager("")
    VerifyOneNameInMgr "Document", mgr, propName
    
    Dim vCfgs As Variant, i As Long
    vCfgs = swModel.GetConfigurationNames
    If Not IsEmpty(vCfgs) Then
        For i = LBound(vCfgs) To UBound(vCfgs)
            Set mgr = swModel.Extension.customPropertyManager(CStr(vCfgs(i)))
            VerifyOneNameInMgr "Config " & CStr(vCfgs(i)), mgr, propName
        Next i
    End If
    
    If swModel.GetType = swDocPART Then
        Dim f As Feature
        Set f = swModel.FirstFeature
        Do While Not f Is Nothing
            If f.GetTypeName2 = "CutListFolder" Then
                VerifyOneNameInMgr "CutList " & f.Name, f.customPropertyManager, propName
            End If
            Set f = f.GetNextFeature
        Loop
    End If
End Sub

Private Sub VerifyOneNameInMgr(scopeLabel As String, mgr As customPropertyManager, propName As String)
    On Error Resume Next
    If mgr Is Nothing Then
        DebugPrint scopeLabel & " verify skip null mgr"
        Exit Sub
    End If
    Dim v As String, r As String, wr As Boolean, l As Boolean
    Dim ret As Long
    ret = mgr.Get6(propName, False, v, r, wr, l)
    DebugPrint scopeLabel & " verify [" & propName & "] ret=" & ret & " wasResolved=" & wr & " val=[" & v & "] resolved=[" & r & "]"
End Sub

Private Sub PrintAscii(labelLine As String, s As String)
    Dim k As Long
    DebugPrint labelLine & " length " & Len(s)
    For k = 1 To Len(s)
        DebugPrint "      pos " & k & " code " & Asc(Mid$(s, k, 1))
    Next k
End Sub

' =========================
' Debug print helpers
' =========================
Private Sub DebugPrintH(s As String)
    If Not DEBUG_MODE Then Exit Sub
    Debug.Print s
End Sub
Private Sub DebugPrint(s As String)
    If Not DEBUG_MODE Then Exit Sub
    Debug.Print s
End Sub
Private Sub DebugPrintE(s As String)
    If Not DEBUG_MODE Then Exit Sub
    Debug.Print "ERROR " & s
End Sub
