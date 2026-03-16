Attribute VB_Name = "MoveToPDM1"
Option Explicit

' ===== CONFIGURATION =====
Const SKIP_ENABLED As Boolean = True  ' Set to False to process all parts, True to skip standard/fastener parts
Const PROCESS_SUBFOLDERS As Boolean = True ' Set to True to process files in subfolders
Const RUN_FROM_ACTIVE_MODEL As Boolean = True ' Set to True to run on active model only, False for folder selection
Const DRAWINGS_ONLY As Boolean = False ' Set to True to process only drawing files (.slddrw)

' SolidWorks API objects
Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2

' Dictionary to store property mappings (old name to new name)
Dim propertyMappings As Object ' Use Object instead of Dictionary for compatibility

Sub Main()
    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    
    ' Initialize property mappings
    InitializePropertyMappings
    
    If RUN_FROM_ACTIVE_MODEL Then
        ' Get the active document
        Set swModel = swApp.ActiveDoc
        
        If swModel Is Nothing Then
            MsgBox "No active document found. Please open a document first.", vbExclamation
            Exit Sub
        End If
        
        ' Check if we should only process drawings
        If DRAWINGS_ONLY And swModel.GetType <> swDocDRAWING Then
            MsgBox "DRAWINGS_ONLY is enabled, but the active document is not a drawing.", vbExclamation
            Exit Sub
        End If
        
        ' Process the active model
        ProcessModels swModel
        
        ' Save the active model
        swModel.Save3 swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0
        
        MsgBox "Active model processing completed.", vbInformation
    Else
        ' Original folder processing code
        ' Configure SolidWorks for background processing
        swApp.Visible = False
        swApp.UserControlBackground = True
        
        ' Prompt user to select directory
        Dim folderPath As String
        folderPath = BrowseForFolder()
        
        If folderPath = "" Then
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
        
        ' Show configuration summary and get confirmation
        If Not ShowConfigurationSummary(folderPath) Then
            MsgBox "Operation cancelled by user.", vbInformation
            Exit Sub
        End If
        
        ' Process all files in the directory
        ProcessDirectory folderPath
        
        MsgBox "Directory processing completed.", vbInformation
    End If
End Sub

Function ShowConfigurationSummary(folderPath As String) As Boolean
    Dim summary As String
    Dim propName As Variant
    Dim action As String
    Dim propertyType As String
    
    summary = "Please review the following settings and changes:" & vbNewLine & vbNewLine
    
    ' Add folder processing settings
    summary = summary & "=== PROCESSING SETTINGS ===" & vbNewLine
    If RUN_FROM_ACTIVE_MODEL Then
        summary = summary & "Mode: Processing Active Model Only" & vbNewLine
    Else
        summary = summary & "Selected Path: " & folderPath & vbNewLine
        summary = summary & "Process Subfolders: " & IIf(PROCESS_SUBFOLDERS, "Yes", "No") & vbNewLine
    End If
    summary = summary & "Drawings Only: " & IIf(DRAWINGS_ONLY, "Yes", "No") & vbNewLine
    summary = summary & "Skip Standard Parts: " & IIf(SKIP_ENABLED, "Yes", "No") & vbNewLine & vbNewLine
    
    If SKIP_ENABLED And Not DRAWINGS_ONLY Then
        summary = summary & "Skipped Part Types: MI, FW, MB, HUC, NUT, PUR, RIV, WSR, HYD, PUR, MIS" & vbNewLine
        summary = summary & "Parts with IsFastener property will also be skipped" & vbNewLine & vbNewLine
    End If
    
    If DRAWINGS_ONLY Then
        summary = summary & "Only drawing files (.slddrw) will be processed" & vbNewLine & vbNewLine
    End If
    
    ' Add property changes in order
    summary = summary & "=== PROPERTY CHANGES (In Order of Execution) ===" & vbNewLine
    
    ' 1. Rename Properties
    summary = summary & "1. Rename Properties:" & vbNewLine
    For Each propName In propertyMappings
        action = propertyMappings(propName)
        If action <> "DELETE" And Left(action, 4) <> "NEW=" And action <> "" Then
            summary = summary & "   - " & propName & " ? " & action & vbNewLine
        End If
    Next propName
    
    ' 2. Remove Properties
    summary = summary & vbNewLine & "2. Remove Properties:" & vbNewLine
    For Each propName In propertyMappings
        action = propertyMappings(propName)
        If action = "DELETE" Then
            summary = summary & "   - " & propName & vbNewLine
        End If
    Next propName
    
    ' 3. Add New Properties
    summary = summary & vbNewLine & "3. Add New Properties:" & vbNewLine
    For Each propName In propertyMappings
        action = propertyMappings(propName)
        If Left(action, 4) = "NEW=" Then
            summary = summary & "   - " & propName & " = " & Right(action, Len(action) - 4) & vbNewLine
        End If
    Next propName
    
    summary = summary & vbNewLine & "Warning: This operation will modify "
    If RUN_FROM_ACTIVE_MODEL Then
        summary = summary & "the active model"
    Else
        summary = summary & "SolidWorks files in the selected folder"
        If PROCESS_SUBFOLDERS Then
            summary = summary & " and all subfolders"
        End If
    End If
    summary = summary & "." & vbNewLine & "Do you want to continue?"
    
    ' Show confirmation dialog
    ShowConfigurationSummary = (MsgBox(summary, vbYesNo + vbQuestion, "Confirm Settings") = vbYes)
End Function


Sub InitializePropertyMappings()
    Set propertyMappings = CreateObject("Scripting.Dictionary")
    
    ' ===== PROPERTY MANIPULATIONS =====
    ' Format:
    ' To rename a property:  propertyMappings.Add "OldName", "NewName"
    ' To remove a property:  propertyMappings.Add "PropertyName", "DELETE"
    ' To clear a property value: propertyMappings.Add "PropertyName", ""
    ' To add a new property: propertyMappings.Add "NewPropertyName", "NEW=Value"
    
    ' IMPORTANT: Operations are performed in the order they are added here
    
    ' Rename properties - Use case-insensitive keys but preserve case for new names
    propertyMappings.Add "Type", "Reference Category"
    propertyMappings.Add "Purchased Assembly", "Type"
    
    ' Remove properties
    propertyMappings.Add "Ref", "DELETE"
    propertyMappings.Add "Length", "DELETE"
    
    ' Add new properties
    propertyMappings.Add "Product Group", "NEW=MISC"
    propertyMappings.Add "Class", "NEW=MECH"
    propertyMappings.Add "Revision", "NEW=1"
End Sub

Sub GetAllModelProperties(model As ModelDoc2, ByRef allProps As Collection)
    Set allProps = New Collection
    
    Dim swCustPropMgr As CustomPropertyManager
    Set swCustPropMgr = model.extension.CustomPropertyManager("")
    
    Dim propNames As Variant
    propNames = swCustPropMgr.GetNames()
    
    If Not IsEmpty(propNames) Then
        Dim i As Long
        For i = 0 To UBound(propNames)
            allProps.Add CStr(propNames(i))
        Next i
    End If
End Sub

Function FindMatchingProperty(propertyName As String, allProps As Collection) As String
    Dim prop As Variant
    For Each prop In allProps
        If StrComp(propertyName, prop, vbTextCompare) = 0 Then
            FindMatchingProperty = prop
            Exit Function
        End If
    Next prop
    FindMatchingProperty = ""
End Function

Sub ManipulateProperties(model As ModelDoc2)
    Dim swCustPropMgr As CustomPropertyManager
    Set swCustPropMgr = model.extension.CustomPropertyManager("")
    
    ' Get all existing properties in the model
    Dim allProps As Collection
    GetAllModelProperties model, allProps
    
    Dim propName As Variant
    For Each propName In propertyMappings.Keys
        Dim action As String
        action = propertyMappings(propName)
        
        ' Find the actual property name with correct case
        Dim actualPropName As String
        actualPropName = FindMatchingProperty(CStr(propName), allProps)
        
        Dim propValue As String, resolvedValue As String
        Dim wasResolved As Boolean, linkToProp As Boolean
        
        ' Get the current property value using case-sensitive name if found
        If actualPropName <> "" Then
            swCustPropMgr.Get6 actualPropName, False, propValue, resolvedValue, wasResolved, linkToProp
        End If
        
        Select Case True
            Case action = "DELETE"
                ' Remove the property if it exists
                If actualPropName <> "" Then
                    swCustPropMgr.Delete2 actualPropName
                    Debug.Print "Removed property in " & model.GetTitle & ": " & actualPropName
                   ' WaitSeconds 0.2
                End If
            
            Case Left(action, 4) = "NEW="
                ' Add new property
                Dim newValue As String
                newValue = Right(action, Len(action) - 4)
                swCustPropMgr.Add3 CStr(propName), swCustomInfoText, newValue, swCustomPropertyReplaceValue
                Debug.Print "Added new property to " & model.GetTitle & ": " & propName & " = " & newValue
               ' WaitSeconds 0.5
            
            Case action = ""
                ' Clear property value if it exists
                If actualPropName <> "" And resolvedValue <> "" Then
                    swCustPropMgr.Set2 actualPropName, ""
                    Debug.Print "Cleared value of property in " & model.GetTitle & ": " & actualPropName
                   ' WaitSeconds 0.2
                End If
            
            Case Else
                ' Rename property
                If actualPropName <> "" And resolvedValue <> "" Then
                    swCustPropMgr.Add3 action, swCustomInfoText, resolvedValue, swCustomPropertyReplaceValue
                    'WaitSeconds 0.1
                    swCustPropMgr.Delete2 actualPropName
                    Debug.Print "Renamed property in " & model.GetTitle & ": " & actualPropName & " to " & action & " (Value: " & resolvedValue & ")"
                    'WaitSeconds 0.2
                End If
        End Select
    Next propName
End Sub

Sub ProcessModels(rootModel As ModelDoc2)
    ' For drawings, we only process the drawing itself, not referenced models
    If DRAWINGS_ONLY Or rootModel.GetType = swDocDRAWING Then
        Debug.Print "Processing drawing: " & rootModel.GetTitle
        ManipulateProperties rootModel
        Exit Sub
    End If
    
    Dim processedModels As Object
    Dim modelsToProcess As Object
    Set processedModels = CreateObject("Scripting.Dictionary")
    Set modelsToProcess = CreateObject("Scripting.Dictionary")
    
    ' Start with the root model
    modelsToProcess.Add rootModel.GetPathName(), rootModel
    
    Do While modelsToProcess.Count > 0
        Dim currentModel As ModelDoc2
        Set currentModel = modelsToProcess.Items()(0)
        modelsToProcess.Remove currentModel.GetPathName()
        
        If Not processedModels.Exists(currentModel.GetPathName()) Then
            processedModels.Add currentModel.GetPathName(), currentModel
            
            ' Check if the model should be skipped (only if skip is enabled)
            If Not SKIP_ENABLED Or Not IsSkippedType(currentModel) Then
                ' Process the model here
                Debug.Print "Processing model: " & currentModel.GetTitle
                
                ' Manipulate properties
                ManipulateProperties currentModel
                
                ' Add subcomponents and referenced models to the processing queue
                If currentModel.GetType = swDocASSEMBLY Then
                    AddSubcomponents currentModel, modelsToProcess, processedModels
                End If
                AddReferencedModels currentModel, modelsToProcess, processedModels
            Else
                Debug.Print "Skipping model: " & currentModel.GetTitle & " (Skip enabled)"
            End If
        End If
    Loop
End Sub

Function IsSkippedType(model As ModelDoc2) As Boolean
    ' If skip is disabled or we're only processing drawings, always return False
    If Not SKIP_ENABLED Or DRAWINGS_ONLY Then
        IsSkippedType = False
        Exit Function
    End If
    
    Dim swCustPropMgr As CustomPropertyManager
    Set swCustPropMgr = model.extension.CustomPropertyManager("")
    
    ' Check "Type" property
    Dim typeValue As String
    swCustPropMgr.Get2 UCase("Type"), "", typeValue
    
    If typeValue <> "" Then
        Dim skippedTypes As Variant
        skippedTypes = Array("MI", "FW", "MB", "HUC", "NUT", "PUR", "RIV", "WSR", "HYD", "PUR", "MIS")
        
        If IsInArray(UCase(typeValue), skippedTypes) Then
            IsSkippedType = True
            Exit Function
        End If
    End If
    
    ' Check "IsFastener" property
    Dim isFastenerValue As String
    swCustPropMgr.Get2 UCase("IsFastener"), "", isFastenerValue
    
    If isFastenerValue <> "" Then
        IsSkippedType = True
        Exit Function
    End If
    
    IsSkippedType = False
End Function

Sub WaitSeconds(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

Sub AddSubcomponents(model As ModelDoc2, ByRef modelsToProcess As Object, ByRef processedModels As Object)
    Dim swAssy As AssemblyDoc
    Set swAssy = model
    
    Dim vComponents As Variant
    vComponents = swAssy.GetComponents(False)
    
    Dim i As Long
    For i = 0 To UBound(vComponents)
        Dim swComp As Component2
        Set swComp = vComponents(i)
        
        Dim compModel As ModelDoc2
        Set compModel = swComp.GetModelDoc2
        
        If Not compModel Is Nothing Then
            If Not processedModels.Exists(compModel.GetPathName()) And Not modelsToProcess.Exists(compModel.GetPathName()) Then
                modelsToProcess.Add compModel.GetPathName(), compModel
            End If
        End If
    Next i
End Sub

Sub AddReferencedModels(model As ModelDoc2, ByRef modelsToProcess As Object, ByRef processedModels As Object)
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "RefFeature" Then
            Dim refModel As ModelDoc2
            Set refModel = GetReferencedDocument(swFeat)
            
            If Not refModel Is Nothing Then
                If Not processedModels.Exists(refModel.GetPathName()) And Not modelsToProcess.Exists(refModel.GetPathName()) Then
                    modelsToProcess.Add refModel.GetPathName(), refModel
                End If
            End If
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Wend
End Sub

Function GetReferencedDocument(feat As Feature) As ModelDoc2
    On Error Resume Next
    Dim refDoc As ModelDoc2
    Set refDoc = feat.GetSpecificFeature2.GetReferencedDocument
    On Error GoTo 0
    Set GetReferencedDocument = refDoc
End Function

Function IsInArray(value As Variant, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

Function BrowseForFolder() As String
    ' First try manual path input
    Dim manualPath As String
    manualPath = InputBox("Enter folder path directly, or click Cancel to browse:", "Select Folder")
    
    ' If user entered a path, validate it
    If manualPath <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FolderExists(manualPath) Then
            BrowseForFolder = manualPath
            Exit Function
        Else
            MsgBox "Invalid path. Switching to folder browser.", vbInformation
        End If
    End If
    
    ' If manual path was empty or invalid, show folder browser
    Dim shellApp As Object
    Dim folder As Object
    
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, "Select folder containing SolidWorks files", 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

Sub ProcessDirectory(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    ' Process all files in current directory
    Dim file As Object
    For Each file In folder.Files
        If IsSolidWorksFile(file.Name) Then
            ProcessSolidWorksFile file.Path
        End If
    Next file
    
    ' Process subfolders if enabled
    If PROCESS_SUBFOLDERS Then
        Dim subFolder As Object
        For Each subFolder In folder.SubFolders
            ProcessDirectory subFolder.Path
        Next subFolder
    End If
End Sub

Function IsSolidWorksFile(fileName As String) As Boolean
    Dim extension As String
    extension = LCase(Right(fileName, 6))
    
    ' If DRAWINGS_ONLY is enabled, only accept .slddrw files
    If DRAWINGS_ONLY Then
        IsSolidWorksFile = (extension = "slddrw")
    Else
        IsSolidWorksFile = (extension = "sldprt" Or _
                            extension = "sldasm" Or _
                            extension = "slddrw")
    End If
End Function

Function GetDocumentType(fileName As String) As swDocumentTypes_e
    Dim extension As String
    extension = LCase(Right(fileName, 6))
    
    Select Case extension
        Case "sldprt"
            GetDocumentType = swDocPART
        Case "sldasm"
            GetDocumentType = swDocASSEMBLY
        Case "slddrw"
            GetDocumentType = swDocDRAWING
        Case Else
            GetDocumentType = swDocPART ' Default fallback
    End Select
End Function

Sub ProcessSolidWorksFile(filePath As String)
    On Error Resume Next
    
    ' Get the appropriate document type
    Dim docType As swDocumentTypes_e
    docType = GetDocumentType(filePath)
    
    ' Open the document with background processing options
    Dim openDoc As ModelDoc2
    Set openDoc = swApp.OpenDoc6(filePath, docType, _
                                swOpenDocOptions_e.swOpenDocOptions_Silent + _
                                swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", 0, 0)
    
    If Err.Number <> 0 Then
        Debug.Print "Error opening file: " & filePath
        Err.Clear
        Exit Sub
    End If
    
    If Not openDoc Is Nothing Then
        ' Process the model and its references
        ProcessModels openDoc
        
        ' Save and close the document
        openDoc.Save3 swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0
        swApp.CloseDoc openDoc.GetTitle
    End If
    
    On Error GoTo 0
End Sub



