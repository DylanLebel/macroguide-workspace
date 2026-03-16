Attribute VB_Name = "PDFProcessor"
Option Explicit
Option Private Module

' ====================================================================
' COMPLETE PDF MODULE
' ====================================================================

' Additional global variables needed for PDF functionality
' (Add these to your existing global variables section)
Private pdmVault As EdmVault5 ' PDM Vault object for PDF operations
Private filePath As String
Private fileName As String
Private RevValue As String
Private ConfigRevisions As collection ' To store revisions by configuration

' Assembly handling constants (add if not already present)
Const swUserPreferenceIntegerValue_ComponentsToLoad As Long = 8
Const swComponentsToLoadLightweight As Long = 2
Const swDetailingNoOptionSpecified As Long = 0
Const swUserPreferenceToggle_SaveAssemblyAsPartSavesAssemOnly As Long = 106
Const swDontLoadComponents As Long = 128
Const swSaveAsCurrentVersion As Long = 0
Const swSaveAsOptions_Silent As Long = 1

' SolidWorks preference constants (add if not already present)
Const swUserPrefToggle_DisableCompLightweight As Long = 190
Const swUserPrefToggle_DynamicUpdateRebuild As Long = 72
Const swUserPrefToggle_RebuildOnSave As Long = 59
Const swOpenDocOptions_SupressRebuild As Long = 128

' Global settings for batch mode
Dim bIncludeSubfolders As Boolean
Dim bManualMode As Boolean
Dim manualOutputPath As String
Public CancelOperation As Boolean

' State management variables
Dim swWasRunning As Boolean
Dim drawingWasOpen As Boolean
Dim originalDrawingPath As String

' Polling constants
Const POLLING_TIMEOUT_MS As Long = 2000 ' Max time (2 seconds)
Const POLLING_INTERVAL_MS As Long = 250  ' Time between checks (0.25 seconds)


Dim swModel As SldWorks.ModelDoc2
Dim swDraw As SldWorks.drawingDoc
 ' PDM Vault object for PDF operations






#If VBA7 Then
    Dim startTime As LongLong
    Dim currentTime As LongLong
#Else
    Dim startTime As Long
    Dim currentTime As Long
#End If

' PDF Processing function that contains the original PDF main logic
Sub ProcessPDFFromDrawing(activeDoc As ModelDoc2)
    LogToFile "=== Starting PDF Processing ==="
    
    ' Set the global variables that the PDF macro expects
    Set swModel = activeDoc
    Set swDraw = activeDoc
    
    ' Connect to PDM (use existing vault or create new pdmVault)
    If vault Is Nothing Then
        Set vault = New EdmVault5
        If Not vault.IsLoggedIn Then
            vault.LoginAuto "NMT_PDM", 0
        End If
    End If
    Set pdmVault = vault ' Set the PDF-specific vault reference
    
    ' Initialize ConfigRevisions collection
    Set ConfigRevisions = New collection
    
    ' Store original user preferences
    Dim origUserPreference_LoadLightweight As Boolean
    Dim origUserPreference_DynamicUpdate As Boolean
    Dim origUserPreference_RebuildOnSave As Boolean
    Dim warningList As collection
    Set warningList = New collection

    On Error Resume Next
    origUserPreference_LoadLightweight = swApp.GetUserPreferenceToggle(swUserPrefToggle_DisableCompLightweight)
    origUserPreference_DynamicUpdate = swApp.GetUserPreferenceToggle(swUserPrefToggle_DynamicUpdateRebuild)
    origUserPreference_RebuildOnSave = swApp.GetUserPreferenceToggle(swUserPrefToggle_RebuildOnSave)
    On Error GoTo 0

    swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, True
    swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, False
    swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, False
    swApp.CommandInProgress = True

' Check if drawing view is ready
    Dim sheetView As view
    Dim attempts As Integer
    For attempts = 1 To 20
        DoEvents
        ' Simple delay loop instead of Sleep 200
        Dim i As Long
        For i = 1 To 50000: Next i
        On Error Resume Next
        Set sheetView = swDraw.GetFirstView
        On Error GoTo 0
        If Not sheetView Is Nothing Then Exit For
        LogToFile "Attempt " & attempts & " - waiting for view..."
    Next attempts

    If sheetView Is Nothing Then
        LogToFile "ERROR: Drawing view failed to load"
        MsgBox "Drawing view failed to load. Try again.", vbCritical
        GoTo PDFCleanup
    End If

    ' Get drawing file information
    Dim fullPath As String
    fullPath = swDraw.GetPathName
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    Dim fileDir As String
    fileDir = Left(fullPath, InStrRev(fullPath, "\"))

    If Not ConnectToPDMVaultPDF() Then
        LogToFile "ERROR: Failed to connect to PDM vault"
        MsgBox "Failed to connect to PDM vault.", vbCritical
        GoTo PDFCleanup
    End If

    Dim pdmDrawingFile As Object
    Set pdmDrawingFile = pdmVault.GetFileFromPath(fullPath)
    If pdmDrawingFile Is Nothing Then
        LogToFile "ERROR: Drawing not found in vault: " & fullPath
        MsgBox "Drawing not found in vault: " & fullPath, vbCritical
        GoTo PDFCleanup
    End If

    'If Not pdmDrawingFile.IsLocked And swDraw.GetSaveFlag Then
    '    LogToFile "ERROR: Drawing is checked in but has unsaved changes"
    '    MsgBox "Drawing is checked in but has unsaved changes. Please check it out first.", vbCritical
    '    GoTo PDFCleanup
    'End If

    GetAllConfigRevisionsPDF

    If ConfigRevisions.Count = 0 Then
        Dim cusPropMgr As SldWorks.customPropertyManager
        Set cusPropMgr = swModel.Extension.customPropertyManager("")
        cusPropMgr.Get3 "Revision", False, "", RevValue

        If Trim(RevValue) <> "" Then
            ConfigRevisions.Add RevValue, ""
        End If
    End If

    If ConfigRevisions.Count = 0 Then
        LogToFile "ERROR: No revision found in drawing"
        MsgBox "No revision found in any configuration.", vbCritical
        GoTo PDFCleanup
    End If

    Dim isLibraryPath As Boolean
    isLibraryPath = IsLibraryPathCheckPDF(fullPath)

    Dim drawingsFolder As String
    Dim obsoleteFolder As String

    If isLibraryPath Then
        drawingsFolder = fileDir & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        EnsureFolderExistsPDF fileDir, "Drawings", drawingsFolder
        EnsureFolderExistsPDF drawingsFolder, "Obsolete", obsoleteFolder
    Else
        Dim designFolder As String
        Dim pos As Long
        pos = InStrRev(fileDir, "3 - Design\")
        If pos > 0 Then
            designFolder = Left(fileDir, pos + Len("3 - Design") - 1) & "\"
        Else
            LogToFile "ERROR: Could not locate '3 - Design' in drawing path"
            MsgBox "Could not locate '3 - Design' in the drawing path.", vbCritical
            GoTo PDFCleanup
        End If

        drawingsFolder = designFolder & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        EnsureFolderExistsPDF designFolder, "Drawings", drawingsFolder
        EnsureFolderExistsPDF drawingsFolder, "Obsolete", obsoleteFolder
    End If

    RevValue = ConfigRevisions.item(1)
    Dim pdfPathHistory As String
    pdfPathHistory = drawingsFolder & fileName & "_Rev" & RevValue & ".pdf"

    Dim errorList As collection
    Set errorList = New collection

    Dim modelUpdated As Boolean
    modelUpdated = SyncDrawingWithModelPDF(fullPath, drawingsFolder, obsoleteFolder, errorList, warningList, False)

    If modelUpdated Then
        HandlePDFCheckInPDF pdfPathHistory, RevValue, drawingsFolder, obsoleteFolder, False
        LogToFile "PDF processing completed successfully"
        MsgBox "PDF '" & fileName & "_Rev" & RevValue & ".pdf" & "' created and processed successfully.", vbInformation
    End If

    ' Display any errors or warnings
    If errorList.Count > 0 Then
        Dim errorMsg As String
        Dim item As Variant
        errorMsg = "Errors encountered:" & vbCrLf
        For Each item In errorList
            errorMsg = errorMsg & "- " & item & vbCrLf
        Next item
        LogToFile "Errors: " & errorMsg
        MsgBox errorMsg, vbExclamation
    End If

PDFCleanup:
    ' Restore original settings
    On Error Resume Next
    If Not swApp Is Nothing Then
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
    End If
    On Error GoTo 0
    
    LogToFile "=== PDF Processing Complete ==="
End Sub

' Initialize settings
Sub InitializeSettingsPDF()
    bIncludeSubfolders = False
    bManualMode = False
    manualOutputPath = PDF_OUTPUT_PATH
End Sub

Private Function SyncDrawingWithModelPDF(drawingPath As String, drawingsFolder As String, _
    obsoleteFolder As String, ByRef errorList As collection, ByRef warningList As collection, _
    Optional batchMode As Boolean = False) As Boolean
    
    Dim swRefDoc As SldWorks.ModelDoc2
    Dim refDocPath As String
    Dim pdmFile As Object
    Dim pdmFolder As Object
    Dim lockedByUser As Object
    Dim errs As Long
    Dim warns As Long
    Dim modelUpdated As Boolean
    Dim isLibraryPath As Boolean
    Dim drawingRevision As String
    Dim allRevisionsMatch As Boolean
    Dim swModelConfigMgr As SldWorks.ConfigurationManager
    Dim configName As Variant
    Dim configPropMgr As SldWorks.customPropertyManager
    Dim configRevision As String
    Dim currentModelDefaultRevision As String
    Dim modelDefaultPropMgr As SldWorks.customPropertyManager
    Dim configCollection As collection
    Dim configNames As Variant
    Dim i As Long
    Dim origUserPreference_LoadLightweight As Boolean
    Dim origUserPreference_DynamicUpdate As Boolean
    Dim origUserPreference_RebuildOnSave As Boolean
    Dim userAlreadyHadCheckout As Boolean
    Dim drawingState As String
    
    modelUpdated = False
    userAlreadyHadCheckout = False
    drawingState = ""
    
    If ConfigRevisions.Count > 0 Then
        drawingRevision = ConfigRevisions.item(1)
        LogToFile "Drawing Revision: " & drawingRevision
    Else
        LogToFile "No revision found in drawing."
        errorList.Add "No revision found in drawing '" & drawingPath & "'."
        Exit Function
    End If
    
    isLibraryPath = IsLibraryPathCheckPDF(drawingPath)
    LogToFile "Is Library Path: " & isLibraryPath
    
    If isLibraryPath Then
        Dim drawingFolderPath As String
        drawingFolderPath = Left(drawingPath, InStrRev(drawingPath, "\"))
        drawingsFolder = drawingFolderPath & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        EnsureFolderExistsPDF drawingFolderPath, "Drawings", drawingsFolder
        EnsureFolderExistsPDF drawingsFolder, "Obsolete", obsoleteFolder
        LogToFile "Library path detected. Using drawing folder for PDFs: " & drawingsFolder
        LogToFile "Library obsolete folder: " & obsoleteFolder
    End If
    
    Dim sheetView As SldWorks.view
    Set sheetView = swDraw.GetFirstView
    If Not sheetView Is Nothing Then
        Dim firstModelView As SldWorks.view
        Set firstModelView = sheetView.GetNextView
        If Not firstModelView Is Nothing Then
            Set swRefDoc = firstModelView.ReferencedDocument
            If Not swRefDoc Is Nothing Then
                refDocPath = swRefDoc.GetPathName
                LogToFile "Referenced document found: " & refDocPath
            Else
                LogToFile "No referenced document found in drawing: " & drawingPath
                errorList.Add "No referenced document found in drawing: " & drawingPath
                Exit Function
            End If
        Else
            LogToFile "No model views found in drawing: " & drawingPath
            errorList.Add "No model views found in drawing: " & drawingPath
            Exit Function
        End If
    Else
        LogToFile "Failed to get first view in drawing: " & drawingPath
        errorList.Add "Failed to get first view in drawing: " & drawingPath
        Exit Function
    End If
    
    If refDocPath = "" Then
        LogToFile "No referenced model path found for drawing: " & drawingPath
        errorList.Add "No referenced model path found for drawing: " & drawingPath
        Exit Function
    End If
    LogToFile "Referenced model path: " & refDocPath
    
    ' Filename comparison code
    Dim drawingFileName As String
    Dim modelFileName As String
    
    drawingFileName = Mid(drawingPath, InStrRev(drawingPath, "\") + 1)
    drawingFileName = Left(drawingFileName, InStrRev(drawingFileName, ".") - 1)
    
    modelFileName = Mid(refDocPath, InStrRev(refDocPath, "\") + 1)
    modelFileName = Left(modelFileName, InStrRev(modelFileName, ".") - 1)
    
    LogToFile "Drawing filename: " & drawingFileName
    LogToFile "Model filename: " & modelFileName
    
    If StrComp(drawingFileName, modelFileName, vbTextCompare) <> 0 Then
        LogToFile "WARNING: Drawing filename does not match model filename"
        
        If batchMode Then
            LogToFile "Batch mode: Assuming NO SYNC due to filename mismatch"
            warningList.Add "Skipped sync: Drawing '" & drawingPath & "' references model with different filename '" & refDocPath & "'. PDF will be created from drawing only."
            SyncDrawingWithModelPDF = True
            Exit Function
        End If
        
        Dim msgResult As VbMsgBoxResult
        msgResult = MsgBox("The drawing filename (" & drawingFileName & ") " & _
                          "does not match the referenced model filename (" & modelFileName & ")." & vbCrLf & vbCrLf & _
                          "Do you want to continue syncing the revision with this model?", _
                          vbQuestion + vbYesNo, "Filename Mismatch")
                          
        If msgResult = vbNo Then
            LogToFile "User chose not to sync with mismatched filename"
            errorList.Add "User canceled sync: Drawing '" & drawingPath & "' references model with different filename '" & refDocPath & "'. PDF will be created from drawing only."
            SyncDrawingWithModelPDF = True
            Exit Function
        Else
            LogToFile "User chose to continue syncing despite filename mismatch"
        End If
    End If
    
    ' Close the drawing before opening the model
    Dim drawingWasOpen As Boolean
    Dim reopenReadOnly As Boolean
    drawingWasOpen = False
    reopenReadOnly = True
    If Not swDraw Is Nothing Then
        drawingWasOpen = True
        Dim pdmDrawingFile As Object
        Set pdmDrawingFile = pdmVault.GetFileFromPath(drawingPath)
        
        On Error Resume Next
        If Not pdmDrawingFile Is Nothing Then
            drawingState = pdmDrawingFile.GetEnumeratorVariable("State")
            LogToFile "Drawing state from PDM: " & drawingState
        End If
        On Error GoTo 0
        
        If Not pdmDrawingFile Is Nothing Then
            If pdmDrawingFile.IsLocked Then
                On Error Resume Next
                Set lockedByUser = pdmDrawingFile.lockedByUser
                If Err.Number = 0 And Not lockedByUser Is Nothing And lockedByUser.Name = pdmVault.CurrentUser.Name Then
                    If swDraw.GetSaveFlag And drawingState <> "Released" Then
                        swDraw.Save
                        LogToFile "Saved drawing before closing: " & drawingPath
                    ElseIf swDraw.GetSaveFlag And drawingState = "Released" Then
                        LogToFile "Drawing shows changes but is in Released state - not saving: " & drawingPath
                    End If
                    reopenReadOnly = False
                End If
                On Error GoTo 0
          
               Else
                ' TEMPORARILY DISABLED - Drawing unsaved changes check
                'If swDraw.GetSaveFlag And drawingState <> "Released" Then
                '    If batchMode Then
                '        LogToFile "Warning: Drawing is checked in but has unsaved changes: " & drawingPath
                '        warningList.Add "Warning: Drawing '" & drawingPath & "' is checked in but has unsaved changes. Creating PDF from current state without saving.", "CHECKED_IN_CHANGES_" & Mid(drawingPath, InStrRev(drawingPath, "\") + 1)
                '    Else
                '        MsgBox "Drawing is checked in but has unsaved changes. Please check out the drawing to save changes before running the macro.", vbCritical
                '        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
                '        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
                '        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
                '        swApp.CommandInProgress = False
                '        Exit Function
                '    End If
                'ElseIf swDraw.GetSaveFlag And drawingState = "Released" Then
                '    LogToFile "Drawing shows changes but is in Released state - continuing without error: " & drawingPath
                'End If
            End If
        End If
        swApp.CloseDoc drawingPath
        LogToFile "Closed drawing to allow model operations: " & drawingPath
    End If
    
    ' Determine document type
    Dim docType As Long
    If UCase(Right(refDocPath, 7)) = ".SLDPRT" Then
        docType = swDocPART
        LogToFile "Document type is PART"
    ElseIf UCase(Right(refDocPath, 7)) = ".SLDASM" Then
        docType = swDocASSEMBLY
        LogToFile "Document type is ASSEMBLY"
    Else
        LogToFile "Referenced file is not a part or assembly: " & refDocPath
        errorList.Add "Referenced file '" & refDocPath & "' is not a part or assembly."
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If
    
    ' Check if the model file is available in PDM
    On Error Resume Next
    Set pdmFile = pdmVault.GetFileFromPath(refDocPath)
    If Err.Number <> 0 Or pdmFile Is Nothing Then
        LogToFile "Error: Model file not found in PDM: " & refDocPath
        errorList.Add "Model file not found in PDM: " & refDocPath
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If
    On Error GoTo 0
    
    ' Check if model is locked by another user BEFORE trying to open it
    On Error Resume Next
    If pdmFile.IsLocked Then
        Set lockedByUser = pdmFile.lockedByUser
        If Err.Number = 0 And Not lockedByUser Is Nothing Then
            If lockedByUser.Name <> pdmVault.CurrentUser.Name Then
                LogToFile "Error: Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                errorList.Add "Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                If drawingWasOpen Then
                    Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
                End If
                Exit Function
            Else
                LogToFile "Model '" & refDocPath & "' is currently checked out by you. Continuing with macro."
            End If
        End If
    End If
    On Error GoTo 0
    
    ' Store original user preferences
    On Error Resume Next
    origUserPreference_LoadLightweight = swApp.GetUserPreferenceToggle(swUserPrefToggle_DisableCompLightweight)
    origUserPreference_DynamicUpdate = swApp.GetUserPreferenceToggle(swUserPrefToggle_DynamicUpdateRebuild)
    origUserPreference_RebuildOnSave = swApp.GetUserPreferenceToggle(swUserPrefToggle_RebuildOnSave)
    On Error GoTo 0
    
    ' Modify SolidWorks settings
    swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, True
    swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, False
    swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, False
    swApp.CommandInProgress = True
    
    ' Open model in read-only mode to check revisions
    LogToFile "Attempting to open model in read-only mode: " & refDocPath
    If docType = swDocASSEMBLY Then
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_ComponentsToLoad, swComponentsToLoadLightweight
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly Or swOpenDocOptions_SupressRebuild Or swDontLoadComponents, "", errs, warns)
    Else
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly Or swOpenDocOptions_SupressRebuild, "", errs, warns)
    End If
    
    If swRefDoc Is Nothing Then
        LogToFile "Failed to open model in read-only mode: " & refDocPath & " Errors: " & errs & " Warns: " & warns
        errorList.Add "Failed to open model '" & refDocPath & "' in read-only mode."
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If
    LogToFile "Successfully opened model in read-only mode"
    
    ' Check for broken features or rebuild errors
    If HasFeatureRebuildErrorsPDF(swRefDoc) Then
        LogToFile "Model has broken features or rebuild errors: " & refDocPath
        errorList.Add "Model '" & refDocPath & "' has broken or unresolved features. Please fix these issues before running the macro again."
        swApp.CloseDoc refDocPath
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        SyncDrawingWithModelPDF = False
        Exit Function
    End If
    
    ' Check default revision
    Set modelDefaultPropMgr = swRefDoc.Extension.customPropertyManager("")
    If Not modelDefaultPropMgr Is Nothing Then
        modelDefaultPropMgr.Get3 "Revision", False, "", currentModelDefaultRevision
        LogToFile "Model default revision: " & currentModelDefaultRevision
        
        If Trim(currentModelDefaultRevision) <> "" Then
            If currentModelDefaultRevision <> drawingRevision Then
                allRevisionsMatch = False
                LogToFile "Default revision mismatch: " & currentModelDefaultRevision & " vs " & drawingRevision
            Else
                allRevisionsMatch = True
                LogToFile "Default revision matches drawing revision: " & drawingRevision
            End If
        Else
            LogToFile "Default configuration has no revision property - will create it"
            allRevisionsMatch = False
        End If
    Else
        LogToFile "Unable to access model's default custom properties."
        allRevisionsMatch = False
    End If
    
    ' Collection to store all found configuration names
    Set configCollection = New collection
    
    ' Try multiple methods to get configurations
    LogToFile "Attempting to get configurations using multiple methods"
    
    ' METHOD 1: Try using GetConfigurationNames
    On Error Resume Next
    LogToFile "METHOD 1: Using GetConfigurationNames"
    configNames = swRefDoc.GetConfigurationNames()
    
    If Err.Number = 0 And Not IsEmpty(configNames) Then
        LogToFile "GetConfigurationNames successful, found " & UBound(configNames) + 1 & " configurations"
        For i = 0 To UBound(configNames)
            On Error Resume Next
            configCollection.Add configNames(i), configNames(i)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        Next i
    Else
        LogToFile "GetConfigurationNames failed: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' METHOD 2: Try using ConfigurationManager
    On Error Resume Next
    LogToFile "METHOD 2: Using ConfigurationManager"
    Set swModelConfigMgr = swRefDoc.ConfigurationManager
    
    If Err.Number = 0 And Not swModelConfigMgr Is Nothing Then
        Dim activeConfig As SldWorks.Configuration
        Set activeConfig = swModelConfigMgr.ActiveConfiguration
        
        If Not activeConfig Is Nothing Then
            configName = activeConfig.Name
            LogToFile "Found active configuration: " & configName
            
            On Error Resume Next
            configCollection.Add configName, CStr(configName)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Else
        LogToFile "ConfigurationManager approach failed: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' METHOD 3: Try direct enumeration through view configurations
    On Error Resume Next
    LogToFile "METHOD 3: Checking for referenced configurations in views"
    If docType = swDocASSEMBLY Then
        Dim swView As SldWorks.view
        Set swView = swRefDoc.GetFirstView
        
        If Not swView Is Nothing Then
            Dim swComp As SldWorks.Component2
            Set swComp = swView.GetFirstComponent
            
            While Not swComp Is Nothing
                configName = swComp.ReferencedConfiguration
                If configName <> "" Then
                    LogToFile "Found referenced configuration: " & configName
                    On Error Resume Next
                    configCollection.Add configName, CStr(configName)
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo 0
                End If
                Set swComp = swComp.GetNext
            Wend
        End If
    End If
    On Error GoTo 0
    
    LogToFile "Found " & configCollection.Count & " unique configurations to check"
    
    If configCollection.Count > 0 Then
        Dim allMatch As Boolean
        allMatch = True
        For Each configName In configCollection
            LogToFile "Checking configuration: " & configName
            Set configPropMgr = swRefDoc.Extension.customPropertyManager(configName)
            
            If Not configPropMgr Is Nothing Then
                configRevision = ""
                configPropMgr.Get3 "Revision", False, "", configRevision
                LogToFile "Config '" & configName & "' revision: " & configRevision
                
                If Trim(configRevision) <> "" Then
                    If configRevision <> drawingRevision Then
                        allMatch = False
                        LogToFile "Config '" & configName & "' revision mismatch: " & configRevision & " vs " & drawingRevision
                    End If
                Else
                    LogToFile "Config '" & configName & "' has no revision property - will create it"
                    allMatch = False
                End If
            Else
                LogToFile "Unable to access custom properties for config: " & configName
            End If
        Next configName
        allRevisionsMatch = allRevisionsMatch And allMatch
    Else
        LogToFile "No configurations found by any method"
    End If
    
    ' If all revisions match, no update needed
    If allRevisionsMatch Then
        LogToFile "All model revisions (default and configurations) match drawing revision: " & drawingRevision
        swApp.CloseDoc refDocPath
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            If reopenReadOnly Then
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
            Else
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent, "", errs, warns)
            End If
        End If
        SyncDrawingWithModelPDF = True
        Exit Function
    End If
    
    ' Close model before editing
    LogToFile "Closing model before editing"
    swApp.CloseDoc refDocPath

    Dim modelFolderPath As String
    modelFolderPath = Left(refDocPath, InStrRev(refDocPath, "\") - 1)
    LogToFile "Getting model folder from vault: " & modelFolderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(modelFolderPath)
    If pdmFolder Is Nothing Then
        LogToFile "Failed to get model folder in vault: " & modelFolderPath
        errorList.Add "Failed to get model folder for '" & refDocPath & "'."
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If

    LogToFile "Getting model file from vault: " & refDocPath
    Set pdmFile = pdmVault.GetFileFromPath(refDocPath)
    If pdmFile Is Nothing Then
        LogToFile "Failed to get model file from vault: " & refDocPath
        errorList.Add "Failed to get model file '" & refDocPath & "' from vault."
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If

    LogToFile "Checking lock status of model file"
    If pdmFile.IsLocked Then
        On Error Resume Next
        Set lockedByUser = pdmFile.lockedByUser
        If Err.Number = 0 And Not lockedByUser Is Nothing Then
            If lockedByUser.Name = pdmVault.CurrentUser.Name Then
                LogToFile "Model already checked out by current user: " & refDocPath
                userAlreadyHadCheckout = True
            Else
                errorList.Add "Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                LogToFile "Error: Model checked out by another user."
                swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
                swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
                swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
                swApp.CommandInProgress = False
                If drawingWasOpen Then
                    Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
                End If
                Exit Function
            End If
        Else
            errorList.Add "Model '" & refDocPath & "' is checked out and cannot be modified (lock user unavailable)."
            LogToFile "Error retrieving lock user."
            swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
            swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
            swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
            swApp.CommandInProgress = False
            If drawingWasOpen Then
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
            End If
            Exit Function
        End If
        On Error GoTo 0
    Else
        userAlreadyHadCheckout = False
        LogToFile "Attempting to check out model file"
        On Error Resume Next
        pdmFile.LockFile pdmFolder.ID, 0
        If Err.Number <> 0 Then
            LogToFile "Failed to check out model: " & Err.description
            errorList.Add "Failed to check out model '" & refDocPath & "': " & Err.description
            swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
            swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
            swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
            swApp.CommandInProgress = False
            If drawingWasOpen Then
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
            End If
            Exit Function
        End If
        On Error GoTo 0
        LogToFile "Checked out model: " & refDocPath
    End If

    LogToFile "Opening model for editing: " & refDocPath
    If docType = swDocASSEMBLY Then
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_ComponentsToLoad, swComponentsToLoadLightweight
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_SupressRebuild Or swDontLoadComponents, "", errs, warns)
    Else
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_SupressRebuild, "", errs, warns)
    End If
    
    If swRefDoc Is Nothing Then
        LogToFile "Failed to open model for editing: " & refDocPath & " Errors: " & errs & " Warns: " & warns
        If pdmFile.IsLocked Then pdmFile.UnlockFile pdmFolder.ID, "Failed to open", 0
        errorList.Add "Failed to open model '" & refDocPath & "' for editing."
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        Exit Function
    End If
    LogToFile "Successfully opened model for editing"
    
    ' Check for broken features when opened for editing
    If HasFeatureRebuildErrorsPDF(swRefDoc) Then
        LogToFile "Model has broken features or rebuild errors when opened for editing: " & refDocPath
        errorList.Add "Model '" & refDocPath & "' has broken or unresolved features. Please fix these issues before running the macro again."
        swApp.CloseDoc refDocPath
        If pdmFile.IsLocked And Not userAlreadyHadCheckout Then
            pdmFile.UnlockFile pdmFolder.ID, "Failed due to rebuild errors", 1
        End If
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        SyncDrawingWithModelPDF = False
        Exit Function
    End If
    
    ' Configure assembly save behavior
    If docType = swDocASSEMBLY And Not swRefDoc Is Nothing Then
        On Error Resume Next
        swRefDoc.Extension.SetUserPreferenceToggle swUserPreferenceToggle_SaveAssemblyAsPartSavesAssemOnly, True, swDetailingNoOptionSpecified
        Dim saveFlags As Long
        saveFlags = swRefDoc.GetSaveAsOption
        saveFlags = saveFlags Or &H20
        swRefDoc.SetSaveAsOption saveFlags
        On Error GoTo 0
    End If
    
    ' UPDATE REVISION PROPERTIES
    Dim changesMade As Boolean
    changesMade = False
    
    LogToFile "Updating default revision if necessary"
    Set modelDefaultPropMgr = swRefDoc.Extension.customPropertyManager("")
    If Not modelDefaultPropMgr Is Nothing Then
        Dim currentDefaultRev As String
        currentDefaultRev = ""
        modelDefaultPropMgr.Get3 "Revision", False, "", currentDefaultRev
        
        If Trim(currentDefaultRev) <> "" Then
            If currentDefaultRev <> drawingRevision Then
                modelDefaultPropMgr.Set2 "Revision", drawingRevision
                changesMade = True
                LogToFile "Updated model default revision from '" & currentDefaultRev & "' to: " & drawingRevision
            Else
                LogToFile "Default revision already matches: " & drawingRevision
            End If
        Else
            modelDefaultPropMgr.Set2 "Revision", drawingRevision
            changesMade = True
            LogToFile "Created model default revision property with value: " & drawingRevision
        End If
    Else
        LogToFile "Failed to access model's default custom properties."
    End If
    
    LogToFile "Refreshing configuration collection for updates"
    Set configCollection = New collection
    
    On Error Resume Next
    configNames = swRefDoc.GetConfigurationNames()
    If Err.Number = 0 And Not IsEmpty(configNames) Then
        For i = 0 To UBound(configNames)
            On Error Resume Next
            configCollection.Add configNames(i), configNames(i)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        Next i
    Else
        Err.Clear
    End If
    Set swModelConfigMgr = swRefDoc.ConfigurationManager
    If Err.Number = 0 And Not swModelConfigMgr Is Nothing Then
        Set activeConfig = swModelConfigMgr.ActiveConfiguration
        If Not activeConfig Is Nothing Then
            configName = activeConfig.Name
            On Error Resume Next
            configCollection.Add configName, CStr(configName)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If
    On Error GoTo 0
    
    LogToFile "Updating " & configCollection.Count & " configurations"
    If configCollection.Count > 0 Then
        For Each configName In configCollection
            LogToFile "Checking configuration: " & configName
            Set configPropMgr = swRefDoc.Extension.customPropertyManager(configName)
            If Not configPropMgr Is Nothing Then
                Dim currentConfigRev As String
                currentConfigRev = ""
                configPropMgr.Get3 "Revision", False, "", currentConfigRev
                
                If Trim(currentConfigRev) <> "" Then
                    If currentConfigRev <> drawingRevision Then
                        configPropMgr.Set2 "Revision", drawingRevision
                        changesMade = True
                        LogToFile "Updated revision for config '" & configName & "' from '" & currentConfigRev & "' to: " & drawingRevision
                    Else
                        LogToFile "Config '" & configName & "' revision already matches: " & drawingRevision
                    End If
                Else
                    configPropMgr.Set2 "Revision", drawingRevision
                    changesMade = True
                    LogToFile "Created revision property for config '" & configName & "' with value: " & drawingRevision
                End If
            Else
                LogToFile "Unable to access custom properties for config: " & configName
            End If
        Next configName
    Else
        LogToFile "No configurations found to update"
    End If
    
    ' SAVE AND CLOSE MODEL ONLY IF CHANGES WERE MADE
    If changesMade Then
        LogToFile "Changes were made. Saving and closing model."
        If docType = swDocASSEMBLY Then
            On Error Resume Next
            swRefDoc.Extension.SetUserPreferenceToggle swUserPreferenceToggle_SaveAssemblyAsPartSavesAssemOnly, True, swDetailingNoOptionSpecified
            swRefDoc.SetSaveAsOption &H20
            swRefDoc.Save2 True
            On Error GoTo 0
        Else
            swRefDoc.Save
        End If
    Else
        LogToFile "No changes made to revision properties. Closing model without saving."
    End If
    swApp.CloseDoc refDocPath
    
    Set swRefDoc = Nothing
Set swRefDoc = Nothing
' Brief delay to allow SolidWorks to release file
Dim releaseDelay As Long
For releaseDelay = 1 To 75000: Next releaseDelay
    
    LogToFile "Checking in model"
    Dim checkInSucceeded As Boolean
    Dim checkInAttempted As Boolean
    checkInSucceeded = True
    checkInAttempted = False
    
    If pdmFile.IsLocked Then
        If Not userAlreadyHadCheckout Or (userAlreadyHadCheckout And changesMade) Then
            Dim checkInModel As Boolean
            checkInModel = True
            
            If userAlreadyHadCheckout And changesMade Then
                Dim checkInPrompt As VbMsgBoxResult
                checkInPrompt = MsgBox("The model '" & refDocPath & "' was checked out by you before running this macro." & vbCrLf & _
                                   "Changes have been made to the revision properties." & vbCrLf & vbCrLf & _
                                   "Do you want to check in the model now?", _
                                   vbQuestion + vbYesNo, "Check In Model")
                checkInModel = (checkInPrompt = vbYes)
            End If
            
            If checkInModel Then
                checkInAttempted = True
                Dim checkInAttempts As Integer
                Dim checkInSuccess As Boolean
                Dim maxAttempts As Integer
                
                maxAttempts = 3
                checkInAttempts = 0
                checkInSuccess = False
                
                Do
                    checkInAttempts = checkInAttempts + 1
                    On Error Resume Next
                    Dim unlockFlags As Long
                    If changesMade Then
                        unlockFlags = 0
                    Else
                        unlockFlags = 1
                    End If
                    
                    pdmFile.UnlockFile pdmFolder.ID, IIf(changesMade, "Updated revisions by macro", "No changes made"), unlockFlags
                    
                    If Err.Number = 0 Then
                        checkInSuccess = True
                        LogToFile "Successfully checked in model: " & refDocPath
                    Else
                        LogToFile "Check-in attempt " & checkInAttempts & " failed: " & Err.description
                        Err.Clear
If checkInAttempts < maxAttempts Then
    ' Progressive delay using simple loop
    Dim checkInDelay As Long
    For checkInDelay = 1 To (75000 * checkInAttempts): Next checkInDelay
End If
                    End If
                    On Error GoTo 0
                Loop Until checkInSuccess Or checkInAttempts >= maxAttempts
                
                If Not checkInSuccess Then
                    LogToFile "Failed to check in model after " & maxAttempts & " attempts"
                    errorList.Add "Failed to check in model '" & refDocPath & "' after multiple attempts. Model remains checked out."
                    checkInSucceeded = False
                    If Not batchMode Then
                        MsgBox "Failed to check in model: " & refDocPath & vbCrLf & "Please check the model in manually.", vbExclamation
                    End If
                End If
            Else
                LogToFile "User chose to keep model checked out: " & refDocPath
            End If
        Else
            LogToFile "Not checking in model as it was already checked out by user and no changes were made: " & refDocPath
        End If
    End If
    
    ' Verify check-in success
    If checkInModel And checkInAttempted And checkInSucceeded Then
        On Error Resume Next
        Dim verifyFile As Object
        Set verifyFile = pdmVault.GetFileFromPath(refDocPath)
        If Not verifyFile Is Nothing Then
            If verifyFile.IsLocked Then
                Dim verifyUser As Object
                Set verifyUser = verifyFile.lockedByUser
                If Not verifyUser Is Nothing Then
                    If verifyUser.Name = pdmVault.CurrentUser.Name Then
                        LogToFile "WARNING: Model file still checked out by current user after check-in attempt"
                        errorList.Add "Model '" & refDocPath & "' is still checked out after check-in attempts."
                        checkInSucceeded = False
                    End If
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
    swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
    swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
    swApp.CommandInProgress = False

    If drawingWasOpen Then
        LogToFile "Reopening drawing"
        If reopenReadOnly Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        Else
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent, "", errs, warns)
        End If
        
        If drawingState = "Released" And Not swDraw Is Nothing And swDraw.GetSaveFlag Then
            LogToFile "Reopened Released drawing shows changes - silently ignoring these changes"
        End If
        
        If swDraw Is Nothing Then
            LogToFile "Failed to reopen drawing: " & drawingPath
        Else
            LogToFile "Reopened drawing: " & drawingPath
        End If
    End If
    
    modelUpdated = True
    LogToFile "Model revision update complete"
    SyncDrawingWithModelPDF = modelUpdated
End Function

Private Sub HandlePDFCheckInPDF(pdfPath As String, RevValue As String, drawingsFolder As String, obsoleteFolder As String, Optional batchMode As Boolean = False)
    Dim pdmFile As Object
    Dim pdmFolder As Object
    Dim pdmObsoleteFolder As Object
    Dim pdmObsoleteFile As Object
    Dim folderPath As String
    Dim existingRev As String
    Dim newRevNum As Integer
    Dim existingRevNum As Integer
    Dim errs As Long
    Dim warns As Long
    Dim existingFile As String
    Dim isLibraryPath As Boolean
    Dim pdfCreationSuccess As Boolean
    Dim foldersExist As Boolean
    Dim targetPDFName As String
    Dim targetPDFPath As String
    
    Dim posFile As IEdmPos5
    Dim pdmFileIter As IEdmFile5
    Dim pattern As String
    Dim pdfPath_iter As String
    Dim pdfName As String
    Dim pdmFileToMove As IEdmFile5
    Dim newFileInObsolete As IEdmFile5
    Dim localFilePath As String
    Dim searchObsolete As IEdmSearch5
    Dim resultObsolete As IEdmSearchResult5
    Dim batchChanger As Object
    Dim fileToChangeState As IEdmFile5
    Dim retVal As Boolean
    Dim exists As IEdmFile5
    Dim fileCount As Long
    
    pdfCreationSuccess = False
    isLibraryPath = IsLibraryPathCheckPDF(pdfPath)
    LogToFile "Is Library Path: " & isLibraryPath

    ' Extract file name without extension and revision suffix
    fileName = Mid(pdfPath, InStrRev(pdfPath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, "_Rev") - 1)
    LogToFile "Base File Name: " & fileName
    
    ' Construct target PDF filename
    targetPDFName = fileName & "_Rev" & RevValue & ".pdf"
    targetPDFPath = drawingsFolder & targetPDFName
    LogToFile "Target PDF Path: " & targetPDFPath

    folderPath = Left(pdfPath, InStrRev(pdfPath, "\") - 1)
    
    Dim parentPath As String
    parentPath = Left(drawingsFolder, InStrRev(drawingsFolder, "\", Len(drawingsFolder) - 1) - 1)
    Dim folderName As String
    folderName = Mid(drawingsFolder, InStrRev(drawingsFolder, "\", Len(drawingsFolder) - 1) + 1)
    folderName = Left(folderName, Len(folderName) - 1)
    
    foldersExist = EnsureFolderExistsPDF(parentPath, folderName, drawingsFolder)
    If Not foldersExist Then
        If Not batchMode Then
            MsgBox "Could not create or access the drawings folder: " & drawingsFolder & vbCrLf & _
                   "PDF creation canceled.", vbCritical
        End If
        Exit Sub
    End If
    
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If pdmFolder Is Nothing Then
        LogToFile "Error: Could not find folder in PDM: " & folderPath
        If Not batchMode Then
            MsgBox "Could not find or create folder in PDM: " & folderPath, vbCritical
        End If
        Exit Sub
    End If

    ' Ensure obsolete folder exists
    On Error Resume Next
    Dim obsoleteFolderName As String
    Dim obsoleteParentPath As String
    obsoleteParentPath = Left(obsoleteFolder, InStrRev(obsoleteFolder, "\", Len(obsoleteFolder) - 1) - 1)
    obsoleteFolderName = Mid(obsoleteFolder, InStrRev(obsoleteFolder, "\", Len(obsoleteFolder) - 1) + 1)
    obsoleteFolderName = Left(obsoleteFolderName, Len(obsoleteFolderName) - 1)
    
    Dim obsoleteFolderExists As Boolean
    obsoleteFolderExists = EnsureFolderExistsPDF(obsoleteParentPath, obsoleteFolderName, obsoleteFolder)
    If Not obsoleteFolderExists Then
        LogToFile "Warning: Could not create obsolete folder - will continue without moving old PDFs"
    End If
    
    Set pdmObsoleteFolder = pdmVault.GetFolderFromPath(obsoleteFolder)
    If Err.Number <> 0 Or pdmObsoleteFolder Is Nothing Then
        LogToFile "Error getting obsolete folder: " & Err.description
        LogToFile "Continuing without obsolete folder functionality"
        Err.Clear
    End If
    On Error GoTo 0

    ' COMPREHENSIVE OBSOLETE HANDLING
    LogToFile "Starting to list all files in Drawings folder for obsolete processing."
    Set posFile = pdmFolder.GetFirstFilePosition
    If posFile Is Nothing Then
        LogToFile "No files found in Drawings folder."
    Else
        pattern = fileName & "_Rev*.pdf"
        LogToFile "Looking for PDFs matching pattern: " & pattern
        
        fileCount = 0
        Const MAX_FILES As Long = 1000
        
        On Error Resume Next
        While Not posFile.IsNull And fileCount < MAX_FILES
            fileCount = fileCount + 1
            LogToFile "Processing file #" & fileCount
            
            Set pdmFileIter = pdmFolder.GetNextFile(posFile)
            
            If Err.Number <> 0 Then
                LogToFile "Error getting next file: " & Err.description
                Err.Clear
                GoTo EndFileLoop
            End If
            
            If Not pdmFileIter Is Nothing Then
                LogToFile "Found file: " & pdmFileIter.Name
                If pdmFileIter.Name Like pattern Then
                    pdfPath_iter = pdmFileIter.GetLocalPath(pdmFolder.ID)
                    pdfName = pdmFileIter.Name
                    LogToFile "Found matching PDF: " & pdfPath_iter
                    LogToFile "Processing file: " & pdfName

                    existingRev = ExtractRevisionFromPDFfile(pdfName, fileName)
                    LogToFile "Extracted revision: " & existingRev & ", Current revision: " & RevValue
                    
                    If existingRev <> "" And IsNumeric(existingRev) And IsNumeric(RevValue) Then
                        If CInt(existingRev) < CInt(RevValue) Then
                            LogToFile "Moving lower revision PDF to Obsolete: Revision " & existingRev & " < " & RevValue & "; Path: " & pdfPath_iter
                            Set pdmFileToMove = pdmFileIter
                            
                            If Not pdmFileToMove Is Nothing Then
                                LogToFile "pdmFileToMove.Name: " & pdmFileToMove.Name
                                LogToFile "pdmFileToMove.ID: " & pdmFileToMove.ID
                                LogToFile "pdmFileToMove.IsLocked before lock: " & pdmFileToMove.IsLocked
                                
                                If Not pdmFileToMove.IsLocked Then
                                    On Error Resume Next
                                    pdmFileToMove.LockFile pdmFolder.ID, 0
                                    If Err.Number <> 0 Then
                                        LogToFile "Error locking file for move: " & Err.description
                                        Err.Clear
                                    Else
                                        LogToFile "Locked file for move: " & pdfPath_iter
                                    End If
                                    On Error GoTo 0
                                End If
                                
                                localFilePath = pdmFolder.localPath & "\" & pdmFileToMove.Name
                                LogToFile "Local file path for move: " & localFilePath
                                
                                If Dir(localFilePath) = "" Then
                                    LogToFile "Local file not found: " & localFilePath
                                Else
                                    Set searchObsolete = pdmVault.CreateSearch
                                    searchObsolete.StartFolderID = pdmObsoleteFolder.ID
                                    searchObsolete.fileName = pdmFileToMove.Name

                                    On Error Resume Next
                                    Set resultObsolete = searchObsolete.GetFirstResult
                                    If Err.Number <> 0 Then
                                        LogToFile "Error searching Obsolete: " & Err.description
                                        Err.Clear
                                    End If
                                    On Error GoTo 0

                                    If Not resultObsolete Is Nothing Then
                                        LogToFile "File already in Obsolete: " & resultObsolete.path
                                    Else
                                        LogToFile "File not in Obsolete; adding now."
                                        On Error Resume Next
                                        pdmObsoleteFolder.AddFile 0, localFilePath
                                        If Err.Number <> 0 Then
                                            LogToFile "Error adding to Obsolete: " & Err.description
                                            Err.Clear
                                        Else
                                            LogToFile "Added to Obsolete: " & localFilePath
                                            Set newFileInObsolete = pdmObsoleteFolder.GetFile(pdmFileToMove.Name)
                                            If Not newFileInObsolete Is Nothing Then
                                                newFileInObsolete.UnlockFile pdmObsoleteFolder.ID, "Checked in after move", 0
                                                LogToFile "Checked in to Obsolete: " & newFileInObsolete.Name
                                            Else
                                                LogToFile "Failed to get file in Obsolete."
                                            End If
                                        End If
                                        On Error GoTo 0
                                    End If
                                    
                                    If Not newFileInObsolete Is Nothing Then
                                        Set fileToChangeState = newFileInObsolete
                                    Else
                                        Set fileToChangeState = pdmFileToMove
                                    End If
                                    
                                    Set batchChanger = pdmVault.CreateUtility(EdmLib.EdmUtility.EdmUtil_BatchChangeState)

                                    If batchChanger Is Nothing Then
                                        LogToFile "Error: Failed to create batch changer utility."
                                    Else
                                        LogToFile "Batch changer utility created successfully."
                                        batchChanger.AddFile fileToChangeState.ID, pdmObsoleteFolder.ID
                                        retVal = batchChanger.CreateTree("Obsolete")
                                        LogToFile "CreateTree returned: " & retVal
                                        If retVal Then
                                           On Error Resume Next
                                           batchChanger.ChangeState2 CLng(0), vbNullString
                                           If Err.Number <> 0 Then
                                               LogToFile "ChangeState2 error: " & Err.description
                                               Err.Clear
                                               batchChanger.ChangeState2 0, ""
                                               If Err.Number <> 0 Then
                                                   LogToFile "Alternative ChangeState2 also failed: " & Err.description
                                                   Err.Clear
                                               Else
                                                   LogToFile "Alternative ChangeState2 succeeded"
                                               End If
                                           Else
                                               LogToFile "File state changed to Obsolete."
                                           End If
                                           On Error GoTo 0
                                           LogToFile "File state changed to Obsolete."
                                        Else
                                            LogToFile "Error: State transition tree creation failed."
                                        End If
                                    End If
                                    
                                    LogToFile "pdmFileToMove.IsLocked before check-in: " & pdmFileToMove.IsLocked
                                    If pdmFileToMove.IsLocked Then
                                        pdmFileToMove.UnlockFile pdmFolder.ID, "Checked in before removal", 0
                                        LogToFile "Checked in file in Drawings folder."
                                    Else
                                        LogToFile "File already checked in Drawings folder."
                                    End If
                                    
                                    ' Wait for vault update using simple delay
Dim vaultUpdateDelay As Long
For vaultUpdateDelay = 1 To 125000: Next vaultUpdateDelay
LogToFile "pdmFileToMove.IsLocked after check-in: " & pdmFileToMove.IsLocked

                                    pdmFileToMove.Refresh
                                    LogToFile "Refreshed pdmFileToMove."

                                    On Error Resume Next
                                    Set exists = pdmFolder.GetFile(pdmFileToMove.Name)
                                    If Err.Number <> 0 Then
                                        LogToFile "Error getting file after check-in: " & Err.description & " - Assuming file is not in Drawings folder."
                                        Set exists = Nothing
                                        Err.Clear
                                    End If
                                    On Error GoTo 0

                                    If exists Is Nothing Then
                                        LogToFile "File no longer in Drawings folder after check-in."
                                    Else
                                        LogToFile "File still in Drawings folder; attempting removal."
                                        On Error Resume Next
                                        pdmFolder.DeleteFile pdmFolder.ID, pdmFileToMove.ID, 0
                                        If Err.Number <> 0 Then
                                            LogToFile "Error with DeleteFile: " & Err.description
                                            Err.Clear
                                        Else
                                            LogToFile "Deleted from Drawings folder with DeleteFile: " & pdmFileToMove.Name
                                        End If
                                        On Error GoTo 0

                                        On Error Resume Next
                                        Set exists = pdmFolder.GetFile(pdmFileToMove.Name)
                                        If Err.Number <> 0 Then
                                            LogToFile "Error confirming removal: " & Err.description & " - Assuming file is not in Drawings folder."
                                            Set exists = Nothing
                                            Err.Clear
                                        End If
                                        On Error GoTo 0

                                        If exists Is Nothing Then
                                            LogToFile "Confirmed file removed from Drawings folder."
                                        Else
                                            LogToFile "File still in Drawings folder after DeleteFile!"
                                        End If
                                    End If
                                    
                                    If Not resultObsolete Is Nothing Then
                                        LogToFile "File already in Obsolete, processed."
                                    Else
                                        LogToFile "Moved " & pdfName & " to Obsolete."
                                    End If
                                End If
                            Else
                                LogToFile "Failed to get pdmFileToMove for: " & pdfPath_iter
                            End If
                        End If
                    End If
                Else
                    LogToFile "File does not match pattern: " & pdmFileIter.Name
                End If
            Else
                LogToFile "pdmFileIter is Nothing - end of files"
                GoTo EndFileLoop
            End If
        Wend
        
EndFileLoop:
        On Error GoTo 0
        LogToFile "Finished processing " & fileCount & " files in Drawings folder"
    End If

    ' Handle the target PDF
    Set pdmFile = pdmVault.GetFileFromPath(targetPDFPath)
    LogToFile "Checking if target PDF exists in vault: " & targetPDFPath
    If Not pdmFile Is Nothing Then
        LogToFile "Target PDF exists in vault: " & targetPDFPath
        LogToFile "pdmFile.IsLocked: " & pdmFile.IsLocked
        If Not pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.LockFile pdmFolder.ID, 0
            If Err.Number <> 0 Then
                LogToFile "Error locking existing PDF: " & Err.description
                Err.Clear
            Else
                LogToFile "Locked existing PDF: " & targetPDFPath
            End If
            On Error GoTo 0
        End If
    Else
        LogToFile "Target PDF does not exist in vault: " & targetPDFPath
    End If

    ' Create the new PDF
    LogToFile "Creating PDF: " & targetPDFPath
    
    On Error Resume Next
    swDraw.Extension.SaveAs targetPDFPath, 0, 1, Nothing, errs, warns
    If Err.Number <> 0 Then
        LogToFile "Error saving PDF: " & Err.description
        If Not batchMode Then
            MsgBox "Error creating PDF: " & targetPDFPath & vbCrLf & "Error: " & Err.description, vbExclamation
        End If
        Err.Clear
    ElseIf errs <> 0 Then
        LogToFile "SaveAs reported errors: " & errs
        If Not batchMode Then
            MsgBox "Warnings or errors occurred while creating PDF: " & targetPDFPath & vbCrLf & "Error code: " & errs, vbExclamation
        End If
    Else
        pdfCreationSuccess = True
        LogToFile "PDF created successfully: " & targetPDFPath
    End If
    On Error GoTo 0
    
    ' Check if the file was created locally
    If Dir(targetPDFPath) = "" Then
        If Not batchMode Then
            MsgBox "Failed to create PDF at " & targetPDFPath & ". File not found.", vbCritical
        End If
        LogToFile "Error: Failed to create PDF. Local file not found: " & targetPDFPath
        Exit Sub
    End If
    LogToFile "Local PDF file exists: " & targetPDFPath

    ' Check in or add the file
    If Not pdmFile Is Nothing Then
        LogToFile "pdmFile is not Nothing, proceeding with update."
        If pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.UnlockFile pdmFolder.ID, "Updated by macro", 0
            If Err.Number <> 0 Then
                LogToFile "Error checking in updated PDF: " & Err.description
                Err.Clear
            Else
                LogToFile "Checked in updated PDF: " & targetPDFPath
            End If
            On Error GoTo 0
        Else
            LogToFile "Existing PDF was not locked after creation, which is unexpected."
        End If
    Else
        LogToFile "Adding new PDF to vault."
        On Error Resume Next
        pdmFolder.AddFile 0, targetPDFPath
        If Err.Number <> 0 Then
            LogToFile "Error adding PDF to vault: " & Err.description
            If Not batchMode Then
                MsgBox "PDF created but could not be added to PDM: " & targetPDFPath, vbExclamation
            End If
            Err.Clear
        Else
            LogToFile "Added PDF to vault: " & targetPDFPath
            Set pdmFile = pdmVault.GetFileFromPath(targetPDFPath)
            If Not pdmFile Is Nothing Then
                LogToFile "pdmFile retrieved after adding: " & pdmFile.Name
                LogToFile "pdmFile.IsLocked: " & pdmFile.IsLocked
                If pdmFile.IsLocked Then
                    pdmFile.UnlockFile pdmFolder.ID, "Added by macro", 0
                    LogToFile "Checked in new PDF: " & targetPDFPath
                End If
            Else
                LogToFile "Failed to retrieve pdmFile after adding to vault: " & targetPDFPath
            End If
        End If
        On Error GoTo 0
    End If
    
    If pdfCreationSuccess Then
        Dim runningFromCommandLine As Boolean
        runningFromCommandLine = (Trim(Command) <> "")
        
        If Not batchMode And Not runningFromCommandLine Then
            MsgBox "PDF '" & targetPDFName & "' created and processed successfully.", vbInformation
        Else
            LogToFile "PDF created successfully: " & targetPDFName
        End If
    End If
    
    LogToFile "PDF macro completed successfully."
End Sub

Private Function ExtractRevisionFromPDFfile(fileNameStr As String, baseName As String) As String
    Dim prefix As String
    prefix = baseName & "_Rev"
    LogToFile "ExtractRevisionFromPDF: File=" & fileNameStr & ", Base=" & baseName & ", Prefix=" & prefix
    If Left(fileNameStr, Len(prefix)) = prefix Then
        Dim suffix As String
        suffix = Mid(fileNameStr, Len(prefix) + 1)
        Dim dotPos As Integer
        dotPos = InStr(suffix, ".pdf")
        If dotPos > 0 Then
            ExtractRevisionFromPDFfile = Left(suffix, dotPos - 1)
            LogToFile "Extracted revision: " & ExtractRevisionFromPDFfile
        Else
            ExtractRevisionFromPDFfile = ""
            LogToFile "No valid extension delimiting revision."
        End If
    Else
        ExtractRevisionFromPDFfile = ""
        LogToFile "Filename does not start with expected prefix."
    End If
End Function

Private Sub GetAllConfigRevisionsPDF()
    Set ConfigRevisions = New collection
    Dim docType As Long
    
    If swModel Is Nothing Then
        LogToFile "Error: No document loaded in swModel"
        Exit Sub
    End If
    
    On Error Resume Next
    docType = swModel.GetType
    If Err.Number <> 0 Then
        LogToFile "Error getting document type: " & Err.description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    If docType = swDocDRAWING Then
        LogToFile "Document is a drawing, checking for file-level Revision property"
        Dim drawingPropMgr As SldWorks.customPropertyManager
        Set drawingPropMgr = swModel.Extension.customPropertyManager("")
        Dim drawingRevision As String
        
        If Not drawingPropMgr Is Nothing Then
            drawingPropMgr.Get3 "Revision", False, "", drawingRevision
            If Trim(drawingRevision) <> "" Then
                On Error Resume Next
                ConfigRevisions.Add drawingRevision
                If Err.Number <> 0 Then
                    Err.Clear
                    LogToFile "Error adding revision value for drawing"
                End If
                On Error GoTo 0
                LogToFile "Found Revision in drawing: " & drawingRevision
            End If
        Else
            LogToFile "Warning: Unable to access drawing custom properties"
        End If
    End If
    
    If ConfigRevisions.Count = 0 Then
        Dim defaultPropMgr As SldWorks.customPropertyManager
        Set defaultPropMgr = swModel.Extension.customPropertyManager("")
        Dim defaultRevision As String
        
        If Not defaultPropMgr Is Nothing Then
            defaultPropMgr.Get3 "Revision", False, "", defaultRevision
            If Trim(defaultRevision) <> "" Then
                On Error Resume Next
                ConfigRevisions.Add defaultRevision
                If Err.Number <> 0 Then
                    Err.Clear
                    LogToFile "Error adding revision value for default config"
                End If
                On Error GoTo 0
                LogToFile "Found Revision in default (file) properties: " & defaultRevision
            End If
        Else
            LogToFile "Warning: Unable to access default custom properties"
        End If
    End If
End Sub

Private Function IsLibraryPathCheckPDF(pathToCheck As String) As Boolean
    Dim originalLibraryCheck As Boolean
    originalLibraryCheck = (InStr(1, pathToCheck, PDM_LIBRARIES_PATH, vbTextCompare) > 0)
    
    Dim customLibraryCheck As Boolean
    customLibraryCheck = (InStr(1, pathToCheck, "C:\NMT_PDM\Projects\Capital\8546 - TestDL - testing\3 - Design\Models\Library Pack and Go\New Folder", vbTextCompare) > 0)
    
    IsLibraryPathCheckPDF = originalLibraryCheck Or customLibraryCheck
End Function

Private Function ConnectToPDMVaultPDF() As Boolean
    On Error Resume Next
    If pdmVault Is Nothing Then Set pdmVault = New EdmVault5
    If Not pdmVault.IsLoggedIn Then
        pdmVault.LoginAuto "NMT_PDM", 0
    End If
    If pdmVault.IsLoggedIn Then
        ConnectToPDMVaultPDF = True
        LogToFile "Logged into NMT_PDM vault successfully."
    Else
        ConnectToPDMVaultPDF = False
        LogToFile "Failed to log into NMT_PDM vault."
    End If
    On Error GoTo 0
End Function

Private Function EnsureFolderExistsPDF(parentFolderPath As String, folderName As String, folderPath As String) As Boolean
    Dim pdmParentFolder As Object
    Dim pdmFolder As Object
    Dim success As Boolean
    
    success = False
    LogToFile "Checking if folder already exists: " & folderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        LogToFile folderName & " folder already exists in vault: " & pdmFolder.localPath
        success = True
        EnsureFolderExistsPDF = success
        Exit Function
    End If
    
    LogToFile "Checking parent folder path: " & parentFolderPath
    Set pdmParentFolder = pdmVault.GetFolderFromPath(parentFolderPath)
    If pdmParentFolder Is Nothing Then
        LogToFile "Parent folder not found in vault: " & parentFolderPath
        EnsureFolderExistsPDF = success
        Exit Function
    End If
    
    LogToFile "Target folder not found, attempting to create: " & folderName
    On Error Resume Next
    Dim rootFolder As Object
    Set rootFolder = pdmVault.rootFolder
    If Not rootFolder Is Nothing Then
        Dim vaultPath As String
        vaultPath = Replace(folderPath, PDM_VAULT_ROOT, "\")
        rootFolder.CreateFolderPath vaultPath, 0
        If Err.Number <> 0 Then
            LogToFile "CreateFolderPath from root failed: " & Err.description & " (Error #" & Err.Number & ")"
            Err.Clear
            EnsureFolderExistsPDF = False
            Exit Function
        End If
    Else
        LogToFile "Could not get root folder"
        EnsureFolderExistsPDF = False
        Exit Function
    End If
    On Error GoTo 0
    
    startTime = GetTickCount
    Set pdmFolder = Nothing

    Do
        ' Simple delay loop instead of Sleep
        Dim pollingDelay As Long
        For pollingDelay = 1 To 62500: Next pollingDelay ' Approximately 250ms equivalent
        
        On Error Resume Next
        Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
        On Error GoTo 0

        If Not pdmFolder Is Nothing Then
            LogToFile "EnsureFolderExists: Folder found in vault after polling: " & folderPath
            success = True
            Exit Do
        End If

        currentTime = GetTickCount
        If currentTime - startTime >= POLLING_TIMEOUT_MS Then
            LogToFile "EnsureFolderExists: Timeout (" & POLLING_TIMEOUT_MS & "ms) reached while waiting for folder: " & folderPath
            MsgBox "Timeout waiting for folder '" & folderPath & "' to appear in the vault after creation attempt.", vbExclamation
            success = False
            Exit Do
        End If
    Loop While pdmFolder Is Nothing
    
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        LogToFile "Created " & folderName & " folder in vault: " & folderPath
        success = True
    Else
        LogToFile "Failed to create " & folderName & " folder in vault: " & folderPath
        MsgBox "Failed to create folder in vault: " & folderPath, vbCritical
        EnsureFolderExistsPDF = success
        Exit Function
    End If
    
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            LogToFile "Error creating local folder: " & Err.description
            Err.Clear
            success = False
        Else
            LogToFile "Created " & folderName & " folder locally: " & folderPath
        End If
        On Error GoTo 0
    Else
        LogToFile folderName & " folder already exists locally: " & folderPath
    End If
    
    EnsureFolderExistsPDF = success
End Function

Private Function GetProductSubfolderPDF(fullPath As String) As String
    Dim result As String
    Dim firstSlashPos As Long
    Dim modelsPos As Long
    Dim productsPos As Long
    Dim pathAfterModels As String
    Dim pathAfterProducts As String
    
    result = ""
    modelsPos = InStr(1, fullPath, "\Models\", vbTextCompare)
    
    If modelsPos > 0 Then
        pathAfterModels = Mid(fullPath, modelsPos + Len("\Models\"))
        firstSlashPos = InStr(1, pathAfterModels, "\")
        If firstSlashPos > 0 Then
            result = Left(pathAfterModels, firstSlashPos - 1)
            If Right(result, 2) = " -" Then
                result = Left(result, Len(result) - 2)
            End If
        End If
    Else
        productsPos = InStr(1, fullPath, "\Products\", vbTextCompare)
        If productsPos > 0 Then
            pathAfterProducts = Mid(fullPath, productsPos + Len("\Products\"))
            firstSlashPos = InStr(1, pathAfterProducts, "\")
            If firstSlashPos > 0 Then
                result = Left(pathAfterProducts, firstSlashPos - 1)
                If Right(result, 2) = " -" Then
                    result = Left(result, Len(result) - 2)
                End If
            End If
        End If
    End If
    
    LogToFile "Product subfolder identified: " & result
    GetProductSubfolderPDF = result
End Function

Private Function HasFeatureRebuildErrorsPDF(swModel As SldWorks.ModelDoc2) As Boolean
    ' TEMPORARILY DISABLED - Function will always return False
    HasFeatureRebuildErrorsPDF = False
    Exit Function
End Function

