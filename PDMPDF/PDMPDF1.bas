Attribute VB_Name = "PDMPDF1"
Option Explicit

' ============================================================================
' WINDOWS API DECLARATIONS
' ============================================================================
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongLong
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' ============================================================================
' GLOBAL OBJECTS
' ============================================================================
Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Public swDraw As SldWorks.DrawingDoc
Public pdmVault As EdmVault5

' ============================================================================
' GLOBAL VARIABLES
' ============================================================================
Public filePath As String
Public fileName As String
Public RevValue As String
Public ConfigRevisions As Collection
Public CancelOperation As Boolean

Dim cachedFolders As Object
Dim bIncludeSubfolders As Boolean
Dim bManualMode As Boolean
Dim manualOutputPath As String
Dim bManualModeUseSubfolders As Boolean
Dim swWasRunning As Boolean
Dim drawingWasOpen As Boolean
Dim originalDrawingPath As String

#If VBA7 Then
    Dim startTime As LongLong
    Dim currentTime As LongLong
#Else
    Dim startTime As Long
    Dim currentTime As Long
#End If

' ============================================================================
' CONSTANTS - SolidWorks
' ============================================================================
Const swUserPreferenceIntegerValue_ComponentsToLoad As Long = 8
Const swComponentsToLoadLightweight As Long = 2
Const swDetailingNoOptionSpecified As Long = 0
Const swUserPreferenceToggle_SaveAssemblyAsPartSavesAssemOnly As Long = 106
Const swDontLoadComponents As Long = 128
Const swSaveAsCurrentVersion As Long = 0
Const swSaveAsOptions_Silent As Long = 1
Const swUserPrefToggle_DisableCompLightweight As Long = 190
Const swUserPrefToggle_DynamicUpdateRebuild As Long = 72
Const swUserPrefToggle_RebuildOnSave As Long = 59
Const swOpenDocOptions_SupressRebuild As Long = 128

' ============================================================================
' CONSTANTS - Polling
' ============================================================================
Const POLLING_TIMEOUT_MS As Long = 1000
Const POLLING_INTERVAL_MS As Long = 100

' ============================================================================
' INITIALIZATION
' ============================================================================
Sub InitializeSettings()
    bIncludeSubfolders = False
    bManualMode = False
    manualOutputPath = "C:\Output\"
    bManualModeUseSubfolders = True
End Sub

' ============================================================================
' MAIN ENTRY POINT
' ============================================================================
Sub main()
    CheckForEasterEgg
    CheckUserAndShowInstructions
    
    ' Local variables
    Dim runningFromCommandLine As Boolean
    Dim originalVisibility As Boolean
    Dim origUserPreference_LoadLightweight As Boolean
    Dim origUserPreference_DynamicUpdate As Boolean
    Dim origUserPreference_RebuildOnSave As Boolean
    Dim warningList As Collection
    Dim errorList As Collection
    Dim drawingsFolder As String
    Dim obsoleteFolder As String
    Dim pdfPathHistory As String
    Dim modelUpdated As Boolean
    Dim fullPath As String
    Dim fileDir As String
    Dim docType As Long
    Dim sheetView As View
    Dim attempts As Integer
    Dim pdmDrawingFile As Object
    Dim cusPropMgr As SldWorks.CustomPropertyManager
    Dim item As Variant
    Dim errorMsg As String
    
    originalDrawingPath = ""
    Set warningList = New Collection
    Set errorList = New Collection
    Set ConfigRevisions = New Collection
    
    ' Initialize SolidWorks
    If Not swApp Is Nothing Then
        swWasRunning = True
        Set swModel = swApp.ActiveDoc
        If Not swModel Is Nothing Then
            If swModel.GetType = swDocDRAWING Then
                drawingWasOpen = True
                originalDrawingPath = swModel.GetPathName
            End If
        End If
    Else
        swWasRunning = False
        drawingWasOpen = False
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            Set swApp = CreateObject("SldWorks.Application")
        End If
    End If
    
    runningFromCommandLine = (Trim(Command) <> "")
    originalVisibility = True
    
    If runningFromCommandLine Then
        originalVisibility = swApp.Visible
        swApp.Visible = False
        swApp.UserControl = False
        Debug.Print "Running from command line - SolidWorks set to invisible"
    End If
    
    ' Store and modify user preferences
    On Error Resume Next
    origUserPreference_LoadLightweight = swApp.GetUserPreferenceToggle(swUserPrefToggle_DisableCompLightweight)
    origUserPreference_DynamicUpdate = swApp.GetUserPreferenceToggle(swUserPrefToggle_DynamicUpdateRebuild)
    origUserPreference_RebuildOnSave = swApp.GetUserPreferenceToggle(swUserPrefToggle_RebuildOnSave)
    On Error GoTo 0
    
    swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, True
    swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, False
    swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, False
    swApp.CommandInProgress = True
    
    Set swModel = swApp.ActiveDoc
    
    ' Handle no active document
    If swModel Is Nothing Then
        filePath = Trim(Command)
        
        If filePath = "" Then
            If Not runningFromCommandLine Then
                Dim batchChoice As VbMsgBoxResult
                batchChoice = MsgBox("No drawing is currently open." & vbCrLf & vbCrLf & _
                                    "Would you like to run batch PDF creation?" & vbCrLf & vbCrLf & _
                                    "YES - Process ALL drawings in a folder" & vbCrLf & _
                                    "NO - Process only MY checked-out drawings" & vbCrLf & _
                                    "CANCEL - Exit", _
                                    vbQuestion + vbYesNoCancel, "Run Batch PDF Creation?")

                If batchChoice = vbYes Then
                    BatchPDFRun
                    GoTo Cleanup
                ElseIf batchChoice = vbNo Then
                    BatchPDFRunCheckedOutOnly
                    GoTo Cleanup
                Else
                    GoTo Cleanup
                End If
            Else
                Debug.Print "Error: No file path provided in command line"
                GoTo Cleanup
            End If
        End If
        
        Debug.Print "Opening file: " & filePath
        
        Dim wasVisible As Boolean
        wasVisible = swApp.Visible
        If Not wasVisible And Not runningFromCommandLine Then
            swApp.Visible = True
        End If
        
        Set swModel = swApp.OpenDoc6(filePath, swDocDRAWING, swOpenDocOptions_Silent, "", 0, 0)
        
        If Not runningFromCommandLine Then
            swApp.Visible = wasVisible
        End If
        
        If swModel Is Nothing Then
            If Not runningFromCommandLine Then
                MsgBox "Failed to open file: " & filePath
            End If
            GoTo Cleanup
        End If
        Debug.Print "Opened drawing: " & filePath
    End If
    
    ' Check document type
    docType = swModel.GetType
    
    If docType = swDocASSEMBLY Then
        If Not runningFromCommandLine Then
            Dim assemblyChoice As VbMsgBoxResult
            assemblyChoice = MsgBox("The active document is an assembly." & vbCrLf & vbCrLf & _
                                   "What would you like to do?" & vbCrLf & vbCrLf & _
                                   "YES - AUDIT PDFs (check what's missing/out-of-date)" & vbCrLf & _
                                   "NO - CREATE PDFs for all component drawings" & vbCrLf & _
                                   "CANCEL - Exit the macro", _
                                   vbQuestion + vbYesNoCancel + vbDefaultButton1, "Assembly PDF Options")

            If assemblyChoice = vbYes Then
                AuditAssemblyPDFs
            ElseIf assemblyChoice = vbNo Then
                ProcessAssemblyDrawings
            End If
        Else
            Debug.Print "Assembly detected in command line mode - not supported"
        End If
        GoTo Cleanup
        
    ElseIf docType <> swDocDRAWING Then
        If Not runningFromCommandLine Then
            MsgBox "The active document must be a drawing or assembly." & vbCrLf & vbCrLf & _
                   "Current document type is not supported.", vbInformation
        End If
        GoTo Cleanup
    End If
    
    Set swDraw = swModel
    
    ' Wait for view to load
    For attempts = 1 To 10
        DoEvents
        Sleep 50
        On Error Resume Next
        Set sheetView = swDraw.GetFirstView
        On Error GoTo 0
        If Not sheetView Is Nothing Then Exit For
        Debug.Print "Attempt " & attempts & " - waiting for view..."
    Next attempts
    
    If sheetView Is Nothing Then
        If Not runningFromCommandLine Then
            MsgBox "Drawing view failed to load. Try again.", vbCritical
        End If
        GoTo Cleanup
    End If
    
    ' Get file info
    fullPath = swDraw.GetPathName
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    fileDir = Left(fullPath, InStrRev(fullPath, "\"))
    
    ' Connect to PDM
    If Not ConnectToPDMVault() Then
        If Not runningFromCommandLine Then
            MsgBox "Failed to connect to PDM vault."
        End If
        GoTo Cleanup
    End If
    
    ' Verify drawing in vault
    Set pdmDrawingFile = pdmVault.GetFileFromPath(fullPath)
    If pdmDrawingFile Is Nothing Then
        If Not runningFromCommandLine Then
            MsgBox "Drawing not found in vault: " & fullPath
        End If
        GoTo Cleanup
    End If
    
    ' Get revisions
    GetAllConfigRevisions
    
    If ConfigRevisions.count = 0 Then
        Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
        cusPropMgr.Get3 "Revision", False, "", RevValue
        
        If Trim(RevValue) <> "" Then
            ConfigRevisions.Add RevValue, ""
        End If
    End If
    
    If ConfigRevisions.count = 0 Then
        If Not runningFromCommandLine Then
            MsgBox "No revision found in any configuration."
        End If
        GoTo Cleanup
    End If
    
    ' Check for unsaved changes
    If Not runningFromCommandLine Then
        If swDraw.GetSaveFlag Then
            Dim unsavedChangeMsg As String
            Dim unsavedResponse As VbMsgBoxResult
            Dim drawingIsCheckedOut As Boolean
            
            drawingIsCheckedOut = False
            
            If Not pdmDrawingFile Is Nothing Then
                If pdmDrawingFile.IsLocked Then
                    On Error Resume Next
                    Dim lockedByUser As Object
                    Set lockedByUser = pdmDrawingFile.lockedByUser
                    If Err.Number = 0 And Not lockedByUser Is Nothing Then
                        If lockedByUser.Name = pdmVault.CurrentUser.Name Then
                            drawingIsCheckedOut = True
                        End If
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            
            If drawingIsCheckedOut Then
                unsavedChangeMsg = "WARNING: This drawing has UNSAVED CHANGES" & vbCrLf & vbCrLf & _
                                  "The drawing is currently checked out by you." & vbCrLf & vbCrLf & _
                                  "What the macro does:" & vbCrLf & _
                                  "� Temporarily closes the drawing to sync the model" & vbCrLf & _
                                  "� Creates/updates the PDF file" & vbCrLf & _
                                  "� Reopens the drawing afterward" & vbCrLf & vbCrLf & _
                                  "RISK: If you continue WITHOUT saving:" & vbCrLf & _
                                  "� Any unsaved changes will be LOST when the drawing closes" & vbCrLf & _
                                  "� The reopened drawing will be the last saved version" & vbCrLf & vbCrLf & _
                                  "Do you want to:" & vbCrLf & _
                                  "� YES - SAVE the drawing now and continue with macro" & vbCrLf & _
                                  "� NO - Continue WITHOUT saving (changes will be lost)" & vbCrLf & _
                                  "� CANCEL - Stop the macro and manually save first"
            Else
                unsavedChangeMsg = "WARNING: This drawing has UNSAVED CHANGES" & vbCrLf & vbCrLf & _
                                  "The drawing is currently CHECKED IN (not checked out)." & vbCrLf & vbCrLf & _
                                  "What the macro does:" & vbCrLf & _
                                  "� Temporarily closes the drawing to sync the model" & vbCrLf & _
                                  "� Creates/updates the PDF file" & vbCrLf & _
                                  "� Reopens the drawing afterward" & vbCrLf & vbCrLf & _
                                  "RISK: If you continue:" & vbCrLf & _
                                  "� Your unsaved changes will be LOST when the drawing closes" & vbCrLf & _
                                  "� You cannot save changes because the file is checked in" & vbCrLf & vbCrLf & _
                                  "RECOMMENDATION: Check out the drawing, save your changes, then run the macro again." & vbCrLf & vbCrLf & _
                                  "Do you want to:" & vbCrLf & _
                                  "� YES - Continue anyway (changes will be lost)" & vbCrLf & _
                                  "� NO - Cancel and save/check out manually first"
            End If
            
            If drawingIsCheckedOut Then
                unsavedResponse = MsgBox(unsavedChangeMsg, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Unsaved Changes Detected")
                
                If unsavedResponse = vbYes Then
                    Debug.Print "User chose to save drawing before continuing"
                    On Error Resume Next
                    swDraw.Save
                    If Err.Number <> 0 Then
                        MsgBox "Failed to save the drawing: " & Err.Description & vbCrLf & vbCrLf & _
                               "Macro will not continue.", vbCritical
                        Err.Clear
                        GoTo Cleanup
                    End If
                    On Error GoTo 0
                    Debug.Print "Drawing saved successfully"
                    
                ElseIf unsavedResponse = vbNo Then
                    Debug.Print "User chose to continue without saving - changes will be lost"
                    
                Else
                    Debug.Print "User canceled to save manually"
                    MsgBox "Macro canceled. Please save your changes and run the macro again.", vbInformation
                    GoTo Cleanup
                End If
                
            Else
                unsavedResponse = MsgBox(unsavedChangeMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Unsaved Changes Detected")
                
                If unsavedResponse = vbNo Then
                    Debug.Print "User canceled - drawing is checked in with unsaved changes"
                    MsgBox "Macro canceled. Please check out the drawing, save your changes, then run the macro again.", vbInformation
                    GoTo Cleanup
                Else
                    Debug.Print "User chose to continue - checked in drawing with unsaved changes will be lost"
                End If
            End If
        End If
    End If
    
    ' ===================================================================
    ' DETERMINE PDF SAVE LOCATION using centralized function
    ' ===================================================================
    If Not DeterminePDFSaveLocation(fullPath, drawingsFolder, obsoleteFolder) Then
        Debug.Print "Failed to determine PDF save location"
        GoTo Cleanup
    End If
    
    RevValue = ConfigRevisions.item(1)
    pdfPathHistory = drawingsFolder & fileName & "_Rev" & RevValue & ".pdf"
    
    ' Sync model and create PDF
    modelUpdated = SyncDrawingWithModel(fullPath, drawingsFolder, obsoleteFolder, errorList, warningList, runningFromCommandLine)
    
    If modelUpdated Then
        HandlePDFCheckIn pdfPathHistory, RevValue, drawingsFolder, obsoleteFolder, runningFromCommandLine
    End If
    
    ' Show errors
    If errorList.count > 0 Then
        If Not runningFromCommandLine Then
            errorMsg = "Errors encountered:" & vbCrLf
            For Each item In errorList
                errorMsg = errorMsg & "- " & item & vbCrLf
            Next item
            MsgBox errorMsg, vbExclamation
        Else
            Debug.Print "Errors encountered during batch processing:"
            For Each item In errorList
                Debug.Print "- " & item
            Next item
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not swApp Is Nothing Then
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
    End If
    On Error GoTo 0
    
    ForceExitForCommandLine
    
    On Error Resume Next
    If Not swApp Is Nothing Then
        If originalVisibility <> swApp.Visible Then
            swApp.Visible = originalVisibility
            swApp.UserControl = True
        End If
    End If
    On Error GoTo 0
End Sub




' ============================================================================
' CENTRALIZED PATH DETERMINATION
' ============================================================================
Public Function DeterminePDFSaveLocation(drawingPath As String, ByRef drawingsFolder As String, ByRef obsoleteFolder As String) As Boolean
    Dim isLibraryPath As Boolean
    Dim designFolder As String
    Dim pos As Long
    Dim currentFilePathOnly As String
    Dim productSubfolder As String
    
    DeterminePDFSaveLocation = False
    
    ' MANUAL MODE - OVERRIDES EVERYTHING
    If bManualMode Then
        Debug.Print "MANUAL MODE ACTIVE - Using manual output path: " & manualOutputPath
        
        If manualOutputPath = "" Then
            Debug.Print "ERROR: Manual output path is empty"
            MsgBox "Manual output path is empty. Please edit the macro settings.", vbCritical
            Exit Function
        End If
        
        If Right(manualOutputPath, 1) <> "\" Then manualOutputPath = manualOutputPath & "\"
        
        drawingsFolder = manualOutputPath
        
        If bManualModeUseSubfolders Then
            Debug.Print "Manual mode with subfolders enabled - checking for Product/Model subfolder"
            productSubfolder = GetProductSubfolder(drawingPath)
            If productSubfolder <> "" Then
                drawingsFolder = manualOutputPath & productSubfolder & "\"
                Debug.Print "Using product subfolder in manual mode: " & drawingsFolder
            End If
        End If
        
        obsoleteFolder = drawingsFolder & "Obsolete\"
        
        If Dir(drawingsFolder, vbDirectory) = "" Then
            On Error Resume Next
            MkDir drawingsFolder
            If Err.Number <> 0 Then
                Debug.Print "ERROR: Failed to create manual mode drawings folder: " & drawingsFolder
                Err.Clear
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        If Not EnsureFolderExists(drawingsFolder, "Obsolete", obsoleteFolder) Then
            Debug.Print "ERROR: Failed to create manual mode obsolete folder: " & obsoleteFolder
            Exit Function
        End If
        
        Debug.Print "Manual mode folders determined:"
        Debug.Print "  Drawings: " & drawingsFolder
        Debug.Print "  Obsolete: " & obsoleteFolder
        DeterminePDFSaveLocation = True
        Exit Function
    End If
    
    ' STANDARD MODE - Path-based logic
    Debug.Print "STANDARD MODE - Determining path from drawing location"
    
    isLibraryPath = IsLibraryPathCheck(drawingPath)
    currentFilePathOnly = Left(drawingPath, InStrRev(drawingPath, "\"))
    
    ' CASE 1: LIBRARY PATHS
    If isLibraryPath Then
        Debug.Print "Library path detected"
        drawingsFolder = currentFilePathOnly & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        
        If Not EnsureFolderExists(currentFilePathOnly, "Drawings", drawingsFolder) Then
            Debug.Print "ERROR: Failed to create library drawings folder"
            Exit Function
        End If
        
        If Not EnsureFolderExists(drawingsFolder, "Obsolete", obsoleteFolder) Then
            Debug.Print "ERROR: Failed to create library obsolete folder"
            Exit Function
        End If
        
        Debug.Print "Library folders determined:"
        Debug.Print "  Drawings: " & drawingsFolder
        Debug.Print "  Obsolete: " & obsoleteFolder
        DeterminePDFSaveLocation = True
        Exit Function
    End If
    
    ' CASE 2: PROJECT PATHS (3 - Design)
    pos = InStrRev(currentFilePathOnly, "3 - Design\")
    If pos > 0 Then
        Debug.Print "Project path (3 - Design) detected"
        designFolder = Left(currentFilePathOnly, pos + Len("3 - Design") - 1) & "\"
        drawingsFolder = designFolder & "Drawings\"
        
        If Not EnsureFolderExists(designFolder, "Drawings", drawingsFolder) Then
            Debug.Print "ERROR: Failed to create project drawings folder"
            Exit Function
        End If
        
        If InStr(1, drawingPath, "\Products\", vbTextCompare) > 0 Or _
           InStr(1, drawingPath, "\Models\", vbTextCompare) > 0 Then
            productSubfolder = GetProductSubfolder(drawingPath)
            If productSubfolder <> "" Then
                Debug.Print "Product/Model subfolder detected: " & productSubfolder
                drawingsFolder = drawingsFolder & productSubfolder & "\"
                
                If Not EnsureFolderExists(designFolder & "Drawings\", productSubfolder, drawingsFolder) Then
                    Debug.Print "ERROR: Failed to create product subfolder"
                    Exit Function
                End If
            End If
        End If
        
        obsoleteFolder = drawingsFolder & "Obsolete\"
        
        If Not EnsureFolderExists(drawingsFolder, "Obsolete", obsoleteFolder) Then
            Debug.Print "ERROR: Failed to create project obsolete folder"
            Exit Function
        End If
        
        Debug.Print "Project folders determined:"
        Debug.Print "  Drawings: " & drawingsFolder
        Debug.Print "  Obsolete: " & obsoleteFolder
        DeterminePDFSaveLocation = True
        Exit Function
    End If
    
    ' CASE 3: UNKNOWN PATH STRUCTURE
    Debug.Print "ERROR: Could not determine path structure (not Library, not 3 - Design)"
    MsgBox "Could not determine save location for drawing path:" & vbCrLf & drawingPath & vbCrLf & vbCrLf & _
           "Path must contain either:" & vbCrLf & _
           "� '\Libraries\' for library components" & vbCrLf & _
           "� '3 - Design\' for project files", vbCritical
    DeterminePDFSaveLocation = False
End Function

' ============================================================================
' SYNC DRAWING WITH MODEL (COMPLETE FUNCTION)
' ============================================================================
Public Function SyncDrawingWithModel(drawingPath As String, drawingsFolder As String, _
    obsoleteFolder As String, ByRef errorList As Collection, ByRef warningList As Collection, _
    Optional batchMode As Boolean = False) As Boolean
    
    ' Variable declarations
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
    Dim configPropMgr As SldWorks.CustomPropertyManager
    Dim configRevision As String
    Dim currentModelDefaultRevision As String
    Dim modelDefaultPropMgr As SldWorks.CustomPropertyManager
    Dim configCollection As Collection
    Dim configNames As Variant
    Dim i As Long
    Dim origUserPreference_LoadLightweight As Boolean
    Dim origUserPreference_DynamicUpdate As Boolean
    Dim origUserPreference_RebuildOnSave As Boolean
    Dim userAlreadyHadCheckout As Boolean
    Dim drawingState As String
    Dim referencingDocs As Collection
    Dim closedDocs As Collection
    Dim drawingWasOpen As Boolean
    Dim reopenReadOnly As Boolean
    Dim sheetView As SldWorks.View
    Dim firstModelView As SldWorks.View
    Dim drawingFileName As String
    Dim modelFileName As String
    Dim msgResult As VbMsgBoxResult
    Dim drawingFolderPath As String
    Dim docType As Long
    Dim pdmDrawingFile As Object
    Dim modelFolderPath As String
    Dim activeConfig As SldWorks.Configuration
    Dim swView As SldWorks.View
    Dim swComp As SldWorks.Component2
    Dim allMatch As Boolean
    Dim changesMade As Boolean
    Dim currentDefaultRev As String
    Dim currentConfigRev As String
    Dim saveFlags As Long
    Dim checkInSucceeded As Boolean
    Dim checkInAttempted As Boolean
    Dim checkInModel As Boolean
    Dim checkInPrompt As VbMsgBoxResult
    Dim checkInAttempts As Integer
    Dim checkInSuccess As Boolean
    Dim maxAttempts As Integer
    Dim unlockFlags As Long
    Dim verifyFile As Object
    Dim verifyUser As Object
    
    modelUpdated = False
    userAlreadyHadCheckout = False
    drawingState = ""
    
    If ConfigRevisions.count > 0 Then
        drawingRevision = ConfigRevisions.item(1)
        Debug.Print "Drawing Revision: " & drawingRevision
    Else
        Debug.Print "No revision found in drawing."
        errorList.Add "No revision found in drawing '" & drawingPath & "'."
        Exit Function
    End If
    
    isLibraryPath = IsLibraryPathCheck(drawingPath)
    Debug.Print "Is Library Path: " & isLibraryPath
    
    If isLibraryPath Then
        drawingFolderPath = Left(drawingPath, InStrRev(drawingPath, "\"))
        drawingsFolder = drawingFolderPath & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        EnsureFolderExists drawingFolderPath, "Drawings", drawingsFolder
        EnsureFolderExists drawingsFolder, "Obsolete", obsoleteFolder
        Debug.Print "Library path detected. Using drawing folder for PDFs: " & drawingsFolder
        Debug.Print "Library obsolete folder: " & obsoleteFolder
    End If
    
    Set sheetView = swDraw.GetFirstView
    If Not sheetView Is Nothing Then
        Set firstModelView = sheetView.GetNextView
        If Not firstModelView Is Nothing Then
            Set swRefDoc = firstModelView.ReferencedDocument
            If Not swRefDoc Is Nothing Then
                refDocPath = swRefDoc.GetPathName
                Debug.Print "Referenced document found: " & refDocPath
            Else
                Debug.Print "No referenced document found in drawing: " & drawingPath
                errorList.Add "No referenced document found in drawing: " & drawingPath
                Exit Function
            End If
        Else
            Debug.Print "No model views found in drawing: " & drawingPath
            errorList.Add "No model views found in drawing: " & drawingPath
            Exit Function
        End If
    Else
        Debug.Print "Failed to get first view in drawing: " & drawingPath
        errorList.Add "Failed to get first view in drawing: " & drawingPath
        Exit Function
    End If
    
    If refDocPath = "" Then
        Debug.Print "No referenced model path found for drawing: " & drawingPath
        errorList.Add "No referenced model path found for drawing: " & drawingPath
        Exit Function
    End If
    Debug.Print "Referenced model path: " & refDocPath
    
    drawingFileName = Mid(drawingPath, InStrRev(drawingPath, "\") + 1)
    drawingFileName = Left(drawingFileName, InStrRev(drawingFileName, ".") - 1)
    
    modelFileName = Mid(refDocPath, InStrRev(refDocPath, "\") + 1)
    modelFileName = Left(modelFileName, InStrRev(modelFileName, ".") - 1)
    
    Debug.Print "Drawing filename: " & drawingFileName
    Debug.Print "Model filename: " & modelFileName
    
    If StrComp(drawingFileName, modelFileName, vbTextCompare) <> 0 Then
        Debug.Print "WARNING: Drawing filename does not match model filename"
        
        If batchMode Then
            Debug.Print "Batch mode: Assuming NO SYNC due to filename mismatch"
            warningList.Add "Skipped sync: Drawing '" & drawingPath & "' references model with different filename '" & refDocPath & "'. PDF will be created from drawing only."
            SyncDrawingWithModel = True
            Exit Function
        End If
        
        msgResult = MsgBox("The drawing filename (" & drawingFileName & ") " & _
                          "does not match the referenced model filename (" & modelFileName & ")." & vbCrLf & vbCrLf & _
                          "Do you want to continue syncing the revision with this model?", _
                          vbQuestion + vbYesNo, "Filename Mismatch")
                          
        If msgResult = vbNo Then
            Debug.Print "User chose not to sync with mismatched filename"
            errorList.Add "User canceled sync: Drawing '" & drawingPath & "' references model with different filename '" & refDocPath & "'. PDF will be created from drawing only."
            SyncDrawingWithModel = True
            Exit Function
        Else
            Debug.Print "User chose to continue syncing despite filename mismatch"
        End If
    End If
    
    Set referencingDocs = GetDocumentsReferencingFile(refDocPath, drawingPath)
    
    If referencingDocs.count > 0 Then
        Debug.Print "Model is referenced by " & referencingDocs.count & " open document(s)"
        Set closedDocs = CloseReferencingDocuments(referencingDocs, refDocPath, batchMode)
        
        If closedDocs Is Nothing Then
            Debug.Print "User canceled operation - documents not closed, stopping macro"
            SyncDrawingWithModel = False
            Exit Function
        End If
        
        Debug.Print "Closed " & closedDocs.count & " referencing document(s)"
    Else
        Set closedDocs = New Collection
    End If
    
    drawingWasOpen = False
    reopenReadOnly = True
    If Not swDraw Is Nothing Then
        drawingWasOpen = True
        Set pdmDrawingFile = pdmVault.GetFileFromPath(drawingPath)
        
        On Error Resume Next
        If Not pdmDrawingFile Is Nothing Then
            drawingState = pdmDrawingFile.GetEnumeratorVariable("State")
            Debug.Print "Drawing state from PDM: " & drawingState
        End If
        On Error GoTo 0
        
        If Not pdmDrawingFile Is Nothing Then
            If pdmDrawingFile.IsLocked Then
                On Error Resume Next
                Set lockedByUser = pdmDrawingFile.lockedByUser
                If Err.Number = 0 And Not lockedByUser Is Nothing And lockedByUser.Name = pdmVault.CurrentUser.Name Then
                    If swDraw.GetSaveFlag And drawingState <> "Released" Then
                        swDraw.Save
                        Debug.Print "Saved drawing before closing: " & drawingPath
                    ElseIf swDraw.GetSaveFlag And drawingState = "Released" Then
                        Debug.Print "Drawing shows changes but is in Released state - not saving: " & drawingPath
                    End If
                    reopenReadOnly = False
                End If
                On Error GoTo 0
            End If
        End If
        
        swApp.CloseDoc drawingPath
        Debug.Print "Closed drawing to allow model operations: " & drawingPath
    End If
    
    If UCase(Right(refDocPath, 7)) = ".SLDPRT" Then
        docType = swDocPART
        Debug.Print "Document type is PART"
    ElseIf UCase(Right(refDocPath, 7)) = ".SLDASM" Then
        docType = swDocASSEMBLY
        Debug.Print "Document type is ASSEMBLY"
    Else
        Debug.Print "Referenced file is not a part or assembly: " & refDocPath
        errorList.Add "Referenced file '" & refDocPath & "' is not a part or assembly."
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If
    
    On Error Resume Next
    Set pdmFile = pdmVault.GetFileFromPath(refDocPath)
    If Err.Number <> 0 Or pdmFile Is Nothing Then
        Debug.Print "Error: Model file not found in PDM: " & refDocPath
        errorList.Add "Model file not found in PDM: " & refDocPath
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If
    On Error GoTo 0
    
    On Error Resume Next
    If pdmFile.IsLocked Then
        Set lockedByUser = pdmFile.lockedByUser
        If Err.Number = 0 And Not lockedByUser Is Nothing Then
            If lockedByUser.Name <> pdmVault.CurrentUser.Name Then
                Debug.Print "Error: Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                errorList.Add "Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                If drawingWasOpen Then
                    Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
                End If
                
                If Not closedDocs Is Nothing Then
                    If closedDocs.count > 0 Then
                        ReopenDocuments closedDocs, batchMode
                    End If
                End If
                Exit Function
            Else
                Debug.Print "Model '" & refDocPath & "' is currently checked out by you. Continuing with macro."
            End If
        End If
    End If
    On Error GoTo 0
    
    On Error Resume Next
    origUserPreference_LoadLightweight = swApp.GetUserPreferenceToggle(swUserPrefToggle_DisableCompLightweight)
    origUserPreference_DynamicUpdate = swApp.GetUserPreferenceToggle(swUserPrefToggle_DynamicUpdateRebuild)
    origUserPreference_RebuildOnSave = swApp.GetUserPreferenceToggle(swUserPrefToggle_RebuildOnSave)
    On Error GoTo 0
    
    swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, True
    swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, False
    swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, False
    swApp.CommandInProgress = True
    
    Debug.Print "Attempting to open model in read-only mode: " & refDocPath
    If docType = swDocASSEMBLY Then
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_ComponentsToLoad, swComponentsToLoadLightweight
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly Or swOpenDocOptions_SupressRebuild Or swDontLoadComponents, "", errs, warns)
    Else
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly Or swOpenDocOptions_SupressRebuild, "", errs, warns)
    End If
    
    If swRefDoc Is Nothing Then
        Debug.Print "Failed to open model in read-only mode: " & refDocPath & " Errors: " & errs & " Warns: " & warns
        errorList.Add "Failed to open model '" & refDocPath & "' in read-only mode."
        
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If
    Debug.Print "Successfully opened model in read-only mode"
    
    If HasFeatureRebuildErrors(swRefDoc) Then
        Debug.Print "Model has broken features or rebuild errors: " & refDocPath
        errorList.Add "Model '" & refDocPath & "' has broken or unresolved features. Please fix these issues before running the macro again."
        
        swApp.CloseDoc refDocPath
        
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        
        SyncDrawingWithModel = False
        Exit Function
    End If
    
    Set modelDefaultPropMgr = swRefDoc.Extension.CustomPropertyManager("")
    If Not modelDefaultPropMgr Is Nothing Then
        modelDefaultPropMgr.Get3 "Revision", False, "", currentModelDefaultRevision
        Debug.Print "Model default revision: " & currentModelDefaultRevision
        
        If Trim(currentModelDefaultRevision) <> "" Then
            If currentModelDefaultRevision <> drawingRevision Then
                allRevisionsMatch = False
                Debug.Print "Default revision mismatch: " & currentModelDefaultRevision & " vs " & drawingRevision
            Else
                allRevisionsMatch = True
                Debug.Print "Default revision matches drawing revision: " & drawingRevision
            End If
        Else
            Debug.Print "Default configuration has no revision property - will create it"
            allRevisionsMatch = False
        End If
    Else
        Debug.Print "Unable to access model's default custom properties."
        allRevisionsMatch = False
    End If
    
    Set configCollection = New Collection
    
    Debug.Print "Attempting to get configurations using multiple methods"
    
    On Error Resume Next
    Debug.Print "METHOD 1: Using GetConfigurationNames"
    configNames = swRefDoc.GetConfigurationNames()
    
    If Err.Number = 0 And Not IsEmpty(configNames) Then
        Debug.Print "GetConfigurationNames successful, found " & UBound(configNames) + 1 & " configurations"
        
        For i = 0 To UBound(configNames)
            On Error Resume Next
            configCollection.Add configNames(i), configNames(i)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        Next i
    Else
        Debug.Print "GetConfigurationNames failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Debug.Print "METHOD 2: Using ConfigurationManager"
    Set swModelConfigMgr = swRefDoc.ConfigurationManager
    
    If Err.Number = 0 And Not swModelConfigMgr Is Nothing Then
        Set activeConfig = swModelConfigMgr.ActiveConfiguration
        
        If Not activeConfig Is Nothing Then
            configName = activeConfig.Name
            Debug.Print "Found active configuration: " & configName
            
            On Error Resume Next
            configCollection.Add configName, CStr(configName)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Else
        Debug.Print "ConfigurationManager approach failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Debug.Print "METHOD 3: Checking for referenced configurations in views"
    If docType = swDocASSEMBLY Then
        Set swView = swRefDoc.GetFirstView
        
        If Not swView Is Nothing Then
            Set swComp = swView.GetFirstComponent
            
            While Not swComp Is Nothing
                configName = swComp.ReferencedConfiguration
                If configName <> "" Then
                    Debug.Print "Found referenced configuration: " & configName
                    
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
    
    Debug.Print "Found " & configCollection.count & " unique configurations to check"
    
    If configCollection.count > 0 Then
        allMatch = True
        For Each configName In configCollection
            Debug.Print "Checking configuration: " & configName
            Set configPropMgr = swRefDoc.Extension.CustomPropertyManager(configName)
            
            If Not configPropMgr Is Nothing Then
                configRevision = ""
                configPropMgr.Get3 "Revision", False, "", configRevision
                Debug.Print "Config '" & configName & "' revision: " & configRevision
                
                If Trim(configRevision) <> "" Then
                    If configRevision <> drawingRevision Then
                        allMatch = False
                        Debug.Print "Config '" & configName & "' revision mismatch: " & configRevision & " vs " & drawingRevision
                    End If
                Else
                    Debug.Print "Config '" & configName & "' has no revision property - will create it"
                    allMatch = False
                End If
            Else
                Debug.Print "Unable to access custom properties for config: " & configName
            End If
        Next configName
        allRevisionsMatch = allRevisionsMatch And allMatch
    Else
        Debug.Print "No configurations found by any method"
    End If
    
    If allRevisionsMatch Then
        Debug.Print "All model revisions (default and configurations) match drawing revision: " & drawingRevision
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
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        
        SyncDrawingWithModel = True
        Exit Function
    End If
    
    Debug.Print "Closing model before editing"
    swApp.CloseDoc refDocPath

    modelFolderPath = Left(refDocPath, InStrRev(refDocPath, "\") - 1)
    Debug.Print "Getting model folder from vault: " & modelFolderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(modelFolderPath)
    If pdmFolder Is Nothing Then
        Debug.Print "Failed to get model folder in vault: " & modelFolderPath
        errorList.Add "Failed to get model folder for '" & refDocPath & "'."
        
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If

    Debug.Print "Getting model file from vault: " & refDocPath
    Set pdmFile = pdmVault.GetFileFromPath(refDocPath)
    If pdmFile Is Nothing Then
        Debug.Print "Failed to get model file from vault: " & refDocPath
        errorList.Add "Failed to get model file '" & refDocPath & "' from vault."
        
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If

    Debug.Print "Checking lock status of model file"
    If pdmFile.IsLocked Then
        On Error Resume Next
        Set lockedByUser = pdmFile.lockedByUser
        If Err.Number = 0 And Not lockedByUser Is Nothing Then
            If lockedByUser.Name = pdmVault.CurrentUser.Name Then
                Debug.Print "Model already checked out by current user: " & refDocPath
                userAlreadyHadCheckout = True
            Else
                errorList.Add "Model '" & refDocPath & "' is checked out by another user and cannot be modified."
                Debug.Print "Error: Model checked out by another user."
                
                swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
                swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
                swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
                swApp.CommandInProgress = False
                
                If drawingWasOpen Then
                    Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
                End If
                
                If Not closedDocs Is Nothing Then
                    If closedDocs.count > 0 Then
                        ReopenDocuments closedDocs, batchMode
                    End If
                End If
                Exit Function
            End If
        Else
            errorList.Add "Model '" & refDocPath & "' is checked out and cannot be modified (lock user unavailable)."
            Debug.Print "Error retrieving lock user."
            
            swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
            swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
            swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
            swApp.CommandInProgress = False
            
            If drawingWasOpen Then
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
            End If
            
            If Not closedDocs Is Nothing Then
                If closedDocs.count > 0 Then
                    ReopenDocuments closedDocs, batchMode
                End If
            End If
            Exit Function
        End If
        On Error GoTo 0
    Else
        userAlreadyHadCheckout = False
    
        Debug.Print "Attempting to check out model file"
        On Error Resume Next
        pdmFile.LockFile pdmFolder.ID, 0
        If Err.Number <> 0 Then
            Debug.Print "Failed to check out model: " & Err.Description
            errorList.Add "Failed to check out model '" & refDocPath & "': " & Err.Description
            
            swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
            swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
            swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
            swApp.CommandInProgress = False
            
            If drawingWasOpen Then
                Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
            End If
            
            If Not closedDocs Is Nothing Then
                If closedDocs.count > 0 Then
                    ReopenDocuments closedDocs, batchMode
                End If
            End If
            Exit Function
        End If
        On Error GoTo 0
        Debug.Print "Checked out model: " & refDocPath
    End If

    Debug.Print "Opening model for editing: " & refDocPath
    If docType = swDocASSEMBLY Then
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_ComponentsToLoad, swComponentsToLoadLightweight
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_SupressRebuild Or swDontLoadComponents, "", errs, warns)
    Else
        Set swRefDoc = swApp.OpenDoc6(refDocPath, docType, swOpenDocOptions_Silent Or swOpenDocOptions_SupressRebuild, "", errs, warns)
    End If
    
    If swRefDoc Is Nothing Then
        Debug.Print "Failed to open model for editing: " & refDocPath & " Errors: " & errs & " Warns: " & warns
        If pdmFile.IsLocked Then pdmFile.UnlockFile pdmFolder.ID, "Failed to open", 0
        errorList.Add "Failed to open model '" & refDocPath & "' for editing."
        
        swApp.SetUserPreferenceToggle swUserPrefToggle_DisableCompLightweight, origUserPreference_LoadLightweight
        swApp.SetUserPreferenceToggle swUserPrefToggle_DynamicUpdateRebuild, origUserPreference_DynamicUpdate
        swApp.SetUserPreferenceToggle swUserPrefToggle_RebuildOnSave, origUserPreference_RebuildOnSave
        swApp.CommandInProgress = False
        
        If drawingWasOpen Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        End If
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        Exit Function
    End If
    Debug.Print "Successfully opened model for editing"
    
    If HasFeatureRebuildErrors(swRefDoc) Then
        Debug.Print "Model has broken features or rebuild errors when opened for editing: " & refDocPath
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
        
        If Not closedDocs Is Nothing Then
            If closedDocs.count > 0 Then
                ReopenDocuments closedDocs, batchMode
            End If
        End If
        
        SyncDrawingWithModel = False
        Exit Function
    End If
    
    If docType = swDocASSEMBLY And Not swRefDoc Is Nothing Then
        On Error Resume Next
        swRefDoc.Extension.SetUserPreferenceToggle swUserPreferenceToggle_SaveAssemblyAsPartSavesAssemOnly, True, swDetailingNoOptionSpecified
        saveFlags = swRefDoc.GetSaveAsOption
        saveFlags = saveFlags Or &H20
        swRefDoc.SetSaveAsOption saveFlags
        On Error GoTo 0
    End If
    
    changesMade = False
    
    Debug.Print "Updating default revision if necessary"
    Set modelDefaultPropMgr = swRefDoc.Extension.CustomPropertyManager("")
    If Not modelDefaultPropMgr Is Nothing Then
        currentDefaultRev = ""
        modelDefaultPropMgr.Get3 "Revision", False, "", currentDefaultRev
        
        If Trim(currentDefaultRev) <> "" Then
            If currentDefaultRev <> drawingRevision Then
                modelDefaultPropMgr.Set2 "Revision", drawingRevision
                changesMade = True
                Debug.Print "Updated model default revision from '" & currentDefaultRev & "' to: " & drawingRevision
            Else
                Debug.Print "Default revision already matches: " & drawingRevision
            End If
        Else
            modelDefaultPropMgr.Set2 "Revision", drawingRevision
            changesMade = True
            Debug.Print "Created model default revision property with value: " & drawingRevision
        End If
    Else
        Debug.Print "Failed to access model's default custom properties."
    End If
    
    Debug.Print "Refreshing configuration collection for updates"
    Set configCollection = New Collection
    
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
    
    Debug.Print "Updating " & configCollection.count & " configurations"
    If configCollection.count > 0 Then
        For Each configName In configCollection
            Debug.Print "Checking configuration: " & configName
            Set configPropMgr = swRefDoc.Extension.CustomPropertyManager(configName)
            If Not configPropMgr Is Nothing Then
                currentConfigRev = ""
                configPropMgr.Get3 "Revision", False, "", currentConfigRev
                
                If Trim(currentConfigRev) <> "" Then
                    If currentConfigRev <> drawingRevision Then
                        configPropMgr.Set2 "Revision", drawingRevision
                        changesMade = True
                        Debug.Print "Updated revision for config '" & configName & "' from '" & currentConfigRev & "' to: " & drawingRevision
                    Else
                        Debug.Print "Config '" & configName & "' revision already matches: " & drawingRevision
                    End If
                Else
                    configPropMgr.Set2 "Revision", drawingRevision
                    changesMade = True
                    Debug.Print "Created revision property for config '" & configName & "' with value: " & drawingRevision
                End If
            Else
                Debug.Print "Unable to access custom properties for config: " & configName
            End If
        Next configName
    Else
        Debug.Print "No configurations found to update"
    End If
    
    If changesMade Then
        Debug.Print "Changes were made. Saving and closing model."
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
        Debug.Print "No changes made to revision properties. Closing model without saving."
    End If
    swApp.CloseDoc refDocPath
    
    Set swRefDoc = Nothing
    Sleep 100
    
    Debug.Print "Checking in model"
    checkInSucceeded = True
    checkInAttempted = False
    
    If pdmFile.IsLocked Then
        If Not userAlreadyHadCheckout Or (userAlreadyHadCheckout And changesMade) Then
            checkInModel = True
            
            If userAlreadyHadCheckout And changesMade Then
                checkInPrompt = MsgBox("The model '" & refDocPath & "' was checked out by you before running this macro." & vbCrLf & _
                                   "Changes have been made to the revision properties." & vbCrLf & vbCrLf & _
                                   "Do you want to check in the model now?", _
                                   vbQuestion + vbYesNo, "Check In Model")
                checkInModel = (checkInPrompt = vbYes)
            End If
            
            If checkInModel Then
                checkInAttempted = True
                maxAttempts = 3
                checkInAttempts = 0
                checkInSuccess = False
                
                Do
                    checkInAttempts = checkInAttempts + 1
                    On Error Resume Next
                    If changesMade Then
                        unlockFlags = 0
                    Else
                        unlockFlags = 1
                    End If
                    
                    pdmFile.UnlockFile pdmFolder.ID, IIf(changesMade, "Updated revisions by macro", "No changes made"), unlockFlags
                    
                    If Err.Number = 0 Then
                        checkInSuccess = True
                        Debug.Print "Successfully checked in model: " & refDocPath
                    Else
                        Debug.Print "Check-in attempt " & checkInAttempts & " failed: " & Err.Description
                        Err.Clear
                        If checkInAttempts < maxAttempts Then
                            Sleep 100 * checkInAttempts
                        End If
                    End If
                    On Error GoTo 0
                Loop Until checkInSuccess Or checkInAttempts >= maxAttempts
                
                If Not checkInSuccess Then
                    Debug.Print "Failed to check in model after " & maxAttempts & " attempts"
                    errorList.Add "Failed to check in model '" & refDocPath & "' after multiple attempts. Model remains checked out."
                    checkInSucceeded = False
                    
                    If Not batchMode Then
                        MsgBox "Failed to check in model: " & refDocPath & vbCrLf & "Please check the model in manually.", vbExclamation
                    End If
                End If
            Else
                Debug.Print "User chose to keep model checked out: " & refDocPath
            End If
        Else
            Debug.Print "Not checking in model as it was already checked out by user and no changes were made: " & refDocPath
        End If
    End If
    
    If checkInModel And checkInAttempted And checkInSucceeded Then
        On Error Resume Next
        Set verifyFile = pdmVault.GetFileFromPath(refDocPath)
        If Not verifyFile Is Nothing Then
            If verifyFile.IsLocked Then
                Set verifyUser = verifyFile.lockedByUser
                If Not verifyUser Is Nothing Then
                    If verifyUser.Name = pdmVault.CurrentUser.Name Then
                        Debug.Print "WARNING: Model file still checked out by current user after check-in attempt"
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
        Debug.Print "Reopening drawing"
        If reopenReadOnly Then
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
        Else
            Set swDraw = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent, "", errs, warns)
        End If
        
        If drawingState = "Released" And Not swDraw Is Nothing And swDraw.GetSaveFlag Then
            Debug.Print "Reopened Released drawing shows changes - silently ignoring these changes"
        End If
        
        If swDraw Is Nothing Then
            Debug.Print "Failed to reopen drawing: " & drawingPath
        Else
            Debug.Print "Reopened drawing: " & drawingPath
        End If
    End If
    
    If Not closedDocs Is Nothing Then
        If closedDocs.count > 0 Then
            ReopenDocuments closedDocs, batchMode
        End If
    End If
    
    modelUpdated = True
    Debug.Print "Model revision update complete"
    SyncDrawingWithModel = modelUpdated
End Function





' ============================================================================
' HANDLE PDF CHECK-IN (COMPLETE FUNCTION)
' ============================================================================
Public Function HandlePDFCheckIn(pdfPath As String, RevValue As String, drawingsFolder As String, _
    obsoleteFolder As String, Optional batchMode As Boolean = False) As Boolean
    
    Dim pdmFile As Object
    Dim pdmFolder As Object
    Dim pdmObsoleteFolder As Object
    Dim folderPath As String
    Dim existingRev As String
    Dim errs As Long
    Dim warns As Long
    Dim isLibraryPath As Boolean
    Dim pdfCreationSuccess As Boolean
    Dim foldersExist As Boolean
    Dim fileName As String
    Dim targetPDFName As String
    Dim targetPDFPath As String
    Dim searchResults As IEdmSearch5
    Dim searchResult As IEdmSearchResult5
    Dim pattern As String
    Dim pdfPath_iter As String
    Dim pdfName As String
    Dim pdmFileToMove As IEdmFile5
    Dim newFileInObsolete As IEdmFile5
    Dim localFilePath As String
    Dim searchObsolete As IEdmSearch5
    Dim resultObsolete As IEdmSearchResult5
    Dim exists As IEdmFile5
    Dim parentPath As String
    Dim folderName As String
    Dim obsoleteFolderName As String
    Dim obsoleteParentPath As String
    Dim obsoleteFolderExists As Boolean
    Dim runningFromCommandLine As Boolean
    
    HandlePDFCheckIn = False
    pdfCreationSuccess = False
    
    isLibraryPath = IsLibraryPathCheck(pdfPath)
    Debug.Print "Is Library Path: " & isLibraryPath

    fileName = Mid(pdfPath, InStrRev(pdfPath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, "_Rev") - 1)
    Debug.Print "Base File Name: " & fileName
    
    targetPDFName = fileName & "_Rev" & RevValue & ".pdf"
    targetPDFPath = drawingsFolder & targetPDFName
    Debug.Print "Target PDF Path: " & targetPDFPath

    folderPath = Left(pdfPath, InStrRev(pdfPath, "\") - 1)
    
    parentPath = Left(drawingsFolder, InStrRev(drawingsFolder, "\", Len(drawingsFolder) - 1) - 1)
    folderName = Mid(drawingsFolder, InStrRev(drawingsFolder, "\", Len(drawingsFolder) - 1) + 1)
    folderName = Left(folderName, Len(folderName) - 1)
    
    foldersExist = EnsureFolderExists(parentPath, folderName, drawingsFolder)
    If Not foldersExist Then
        Debug.Print "CRITICAL: Could not create or access the drawings folder: " & drawingsFolder
        If Not batchMode Then
            MsgBox "Could not create or access the drawings folder: " & drawingsFolder & vbCrLf & _
                   "PDF creation canceled.", vbCritical
        End If
        HandlePDFCheckIn = False
        Exit Function
    End If
    
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If pdmFolder Is Nothing Then
        Debug.Print "CRITICAL: Could not find folder in PDM: " & folderPath
        If Not batchMode Then
            MsgBox "Could not find or create folder in PDM: " & folderPath, vbCritical
        End If
        HandlePDFCheckIn = False
        Exit Function
    End If

    On Error Resume Next
    obsoleteParentPath = Left(obsoleteFolder, InStrRev(obsoleteFolder, "\", Len(obsoleteFolder) - 1) - 1)
    obsoleteFolderName = Mid(obsoleteFolder, InStrRev(obsoleteFolder, "\", Len(obsoleteFolder) - 1) + 1)
    obsoleteFolderName = Left(obsoleteFolderName, Len(obsoleteFolderName) - 1)
    
    obsoleteFolderExists = EnsureFolderExists(obsoleteParentPath, obsoleteFolderName, obsoleteFolder)
    If Not obsoleteFolderExists Then
        Debug.Print "Warning: Could not create obsolete folder - will continue without moving old PDFs"
    End If
    
    Set pdmObsoleteFolder = pdmVault.GetFolderFromPath(obsoleteFolder)
    If Err.Number <> 0 Or pdmObsoleteFolder Is Nothing Then
        Debug.Print "Error getting obsolete folder: " & Err.Description
        Debug.Print "Continuing without obsolete folder functionality"
        Err.Clear
    End If
    On Error GoTo 0

    Debug.Print "Starting PDM search for obsolete PDFs"
    pattern = fileName & "_Rev*.pdf"
    Debug.Print "Searching for PDFs matching pattern: " & pattern
    
    On Error Resume Next
    Set searchResults = pdmVault.CreateSearch
    
    If Not searchResults Is Nothing Then
        searchResults.StartFolderID = pdmFolder.ID
        searchResults.fileName = pattern
        Set searchResult = searchResults.GetFirstResult()
        
        Do While Not searchResult Is Nothing
            pdfName = searchResult.Name
            pdfPath_iter = searchResult.Path
            Debug.Print "Found matching PDF: " & pdfName
            existingRev = ExtractRevisionFromPDF(pdfName, fileName)
            Debug.Print "Extracted revision: " & existingRev & ", Current revision: " & RevValue
            
            If existingRev <> "" And IsNumeric(existingRev) And IsNumeric(RevValue) Then
                If CInt(existingRev) < CInt(RevValue) Then
                    Debug.Print "Moving obsolete PDF (Rev " & existingRev & " < " & RevValue & "): " & pdfName
                    Set pdmFileToMove = pdmVault.GetFileFromPath(pdfPath_iter)
                    
                    If Not pdmFileToMove Is Nothing Then
                        If Not pdmFileToMove.IsLocked Then
                            pdmFileToMove.LockFile pdmFolder.ID, 0
                            If Err.Number <> 0 Then
                                Debug.Print "Error locking file: " & Err.Description
                                Err.Clear
                                GoTo NextObsoletePDF
                            End If
                        End If
                        
                        localFilePath = pdfPath_iter
                        
                        If Dir(localFilePath) <> "" Then
                            Set searchObsolete = pdmVault.CreateSearch
                            searchObsolete.StartFolderID = pdmObsoleteFolder.ID
                            searchObsolete.fileName = pdmFileToMove.Name
                            Set resultObsolete = searchObsolete.GetFirstResult
                            
                            If Err.Number <> 0 Then
                                Debug.Print "Error searching Obsolete folder: " & Err.Description
                                Err.Clear
                            End If
                            
                            If resultObsolete Is Nothing Then
                                Debug.Print "Adding to Obsolete folder: " & localFilePath
                                pdmObsoleteFolder.AddFile 0, localFilePath
                                
                                If Err.Number = 0 Then
                                    Debug.Print "Successfully added to Obsolete"
                                    Set newFileInObsolete = pdmObsoleteFolder.GetFile(pdmFileToMove.Name)
                                    If Not newFileInObsolete Is Nothing Then
                                        newFileInObsolete.UnlockFile pdmObsoleteFolder.ID, "Moved to obsolete by macro", 0
                                        Debug.Print "Checked in to Obsolete: " & newFileInObsolete.Name
                                    End If
                                Else
                                    Debug.Print "Error adding to Obsolete: " & Err.Description
                                    Err.Clear
                                    GoTo NextObsoletePDF
                                End If
                            Else
                                Debug.Print "File already exists in Obsolete: " & resultObsolete.Path
                            End If
                            
                            If pdmFileToMove.IsLocked Then
                                pdmFileToMove.UnlockFile pdmFolder.ID, "Checked in before removal", 0
                                Debug.Print "Checked in file in Drawings folder"
                            End If
                            
                            Sleep 100
                            pdmFileToMove.Refresh
                            
                            Set exists = pdmFolder.GetFile(pdmFileToMove.Name)
                            If Err.Number <> 0 Then
                                Set exists = Nothing
                                Err.Clear
                            End If
                            
                            If Not exists Is Nothing Then
                                pdmFolder.DeleteFile pdmFolder.ID, pdmFileToMove.ID, 0
                                If Err.Number = 0 Then
                                    Debug.Print "Deleted from Drawings folder: " & pdmFileToMove.Name
                                Else
                                    Debug.Print "Error deleting: " & Err.Description
                                    Err.Clear
                                End If
                            End If
                            
                            Debug.Print "Successfully moved " & pdfName & " to Obsolete"
                        Else
                            Debug.Print "Local file not found: " & localFilePath
                        End If
                    End If
                End If
            End If
            
NextObsoletePDF:
            Set searchResult = searchResults.GetNextResult()
        Loop
        
        Debug.Print "Finished processing obsolete PDFs using search"
    Else
        Debug.Print "Failed to create search object - skipping obsolete processing"
    End If
    
    On Error GoTo 0

    Set pdmFile = pdmVault.GetFileFromPath(targetPDFPath)
    Debug.Print "Checking if target PDF exists in vault: " & targetPDFPath
    If Not pdmFile Is Nothing Then
        Debug.Print "Target PDF exists in vault: " & targetPDFPath
        Debug.Print "pdmFile.IsLocked: " & pdmFile.IsLocked
        If Not pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.LockFile pdmFolder.ID, 0
            If Err.Number <> 0 Then
                Debug.Print "Error locking existing PDF: " & Err.Description
                Err.Clear
            Else
                Debug.Print "Locked existing PDF: " & targetPDFPath
            End If
            On Error GoTo 0
        End If
    Else
        Debug.Print "Target PDF does not exist in vault: " & targetPDFPath
    End If

    Debug.Print "Creating PDF: " & targetPDFPath
    
    On Error Resume Next
    swDraw.Extension.SaveAs targetPDFPath, 0, 1, Nothing, errs, warns
    If Err.Number <> 0 Then
        Debug.Print "CRITICAL ERROR saving PDF: " & Err.Description
        If Not batchMode Then
            MsgBox "Error creating PDF: " & targetPDFPath & vbCrLf & "Error: " & Err.Description, vbExclamation
        End If
        Err.Clear
        HandlePDFCheckIn = False
        Exit Function
    ElseIf errs <> 0 Then
        Debug.Print "CRITICAL: SaveAs reported errors: " & errs
        If Not batchMode Then
            MsgBox "Warnings or errors occurred while creating PDF: " & targetPDFPath & vbCrLf & "Error code: " & errs, vbExclamation
        End If
        HandlePDFCheckIn = False
        Exit Function
    Else
        pdfCreationSuccess = True
        Debug.Print "PDF created successfully: " & targetPDFPath
    End If
    On Error GoTo 0
    
    If Dir(targetPDFPath) = "" Then
        Debug.Print "CRITICAL ERROR: Failed to create PDF. Local file not found: " & targetPDFPath
        If Not batchMode Then
            MsgBox "Failed to create PDF at " & targetPDFPath & ". File not found.", vbCritical
        End If
        HandlePDFCheckIn = False
        Exit Function
    End If
    Debug.Print "Local PDF file exists: " & targetPDFPath

    If Not pdmFile Is Nothing Then
        Debug.Print "pdmFile is not Nothing, proceeding with update."
        If pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.UnlockFile pdmFolder.ID, "Updated by macro", 0
            If Err.Number <> 0 Then
                Debug.Print "Error checking in updated PDF: " & Err.Description
                Err.Clear
            Else
                Debug.Print "Checked in updated PDF: " & targetPDFPath
            End If
            On Error GoTo 0
        Else
            Debug.Print "Existing PDF was not locked after creation, which is unexpected."
        End If
    Else
        Debug.Print "Adding new PDF to vault."
        On Error Resume Next
        pdmFolder.AddFile 0, targetPDFPath
        If Err.Number <> 0 Then
            Debug.Print "Error adding PDF to vault: " & Err.Description
            If Not batchMode Then
                MsgBox "PDF created but could not be added to PDM: " & targetPDFPath, vbExclamation
            End If
            Err.Clear
        Else
            Debug.Print "Added PDF to vault: " & targetPDFPath
            Set pdmFile = pdmVault.GetFileFromPath(targetPDFPath)
            If Not pdmFile Is Nothing Then
                Debug.Print "pdmFile retrieved after adding: " & pdmFile.Name
                Debug.Print "pdmFile.IsLocked: " & pdmFile.IsLocked
                If pdmFile.IsLocked Then
                    pdmFile.UnlockFile pdmFolder.ID, "Added by macro", 0
                    Debug.Print "Checked in new PDF: " & targetPDFPath
                End If
            Else
                Debug.Print "Failed to retrieve pdmFile after adding to vault: " & targetPDFPath
            End If
        End If
        On Error GoTo 0
    End If
    
    If pdfCreationSuccess Then
        runningFromCommandLine = (Trim(Command) <> "")
        
        If Not batchMode And Not runningFromCommandLine Then
            MsgBox "PDF '" & targetPDFName & "' created and processed successfully.", vbInformation
        Else
            Debug.Print "PDF created successfully: " & targetPDFName
        End If
        
        HandlePDFCheckIn = True
    Else
        Debug.Print "CRITICAL: pdfCreationSuccess is False - something went wrong"
        HandlePDFCheckIn = False
    End If
    
    Debug.Print "HandlePDFCheckIn returning: " & HandlePDFCheckIn
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================
Public Function ExtractRevisionFromPDF(fileNameStr As String, baseName As String) As String
    Dim prefix As String
    prefix = baseName & "_Rev"
    Debug.Print "ExtractRevisionFromPDF: File=" & fileNameStr & ", Base=" & baseName & ", Prefix=" & prefix
    If Left(fileNameStr, Len(prefix)) = prefix Then
        Dim suffix As String
        suffix = Mid(fileNameStr, Len(prefix) + 1)
        Dim dotPos As Integer
        dotPos = InStr(suffix, ".pdf")
        If dotPos > 0 Then
            ExtractRevisionFromPDF = Left(suffix, dotPos - 1)
            Debug.Print "Extracted revision: " & ExtractRevisionFromPDF
        Else
            ExtractRevisionFromPDF = ""
            Debug.Print "No valid extension delimiting revision."
        End If
    Else
        ExtractRevisionFromPDF = ""
        Debug.Print "Filename does not start with expected prefix."
    End If
End Function

Public Function TryCheckOutFile(pdfPath As String) As Boolean
    Dim pdmFile As Object
    Dim pdmFolder As Object
    Dim success As Boolean
    
    success = False
    On Error Resume Next
    Set pdmFile = pdmVault.GetFileFromPath(pdfPath, pdmFolder)
    If Err.Number <> 0 Then
        MsgBox "We're sorry, but the macro was unable to access the file:" & vbCrLf & _
               pdfPath & vbCrLf & vbCrLf & _
               "This could be due to a vault error or incorrect file path." & vbCrLf & _
               "Please check the file path and try again.", vbExclamation, "File Access Failed"
        Err.Clear
        TryCheckOutFile = success
        Exit Function
    End If
    On Error GoTo 0
    
    If pdmFile Is Nothing Then
        MsgBox "The file could not be found in the PDM vault:" & vbCrLf & _
               pdfPath & vbCrLf & vbCrLf & _
               "Please ensure the file exists in the vault and try again.", vbExclamation, "File Not Found"
        TryCheckOutFile = success
        Exit Function
    End If
    
    On Error Resume Next
    pdmFile.LockFile pdmFolder.ID, 0
    If Err.Number <> 0 Then
        MsgBox "We're sorry, but the macro was unable to check out the PDF file:" & vbCrLf & _
               pdfPath & vbCrLf & vbCrLf & _
               "This could be because:" & vbCrLf & _
               "- The file is checked out by another user (it's locked)." & vbCrLf & _
               "- The file is open in another program, like Adobe Acrobat or a web browser." & vbCrLf & _
               "- The file is being viewed in the PDM vault's preview pane." & vbCrLf & vbCrLf & _
               "To fix this, please:" & vbCrLf & _
               "- Make sure the file is checked in to the PDM vault." & vbCrLf & _
               "- Close the file if it's open in any program." & vbCrLf & _
               "- Close the PDM vault's preview pane if the file is showing there." & vbCrLf & vbCrLf & _
               "After that, try running the macro again.", vbExclamation, "PDF Checkout Issue"
        Err.Clear
    Else
        success = True
    End If
    On Error GoTo 0
    TryCheckOutFile = success
End Function

Public Sub GetAllConfigRevisions()
    Set ConfigRevisions = New Collection
    Dim docType As Long
    
    If swModel Is Nothing Then Exit Sub
    
    On Error Resume Next
    docType = swModel.GetType
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    If docType = swDocDRAWING Then
        Dim drawingPropMgr As SldWorks.CustomPropertyManager
        Set drawingPropMgr = swModel.Extension.CustomPropertyManager("")
        Dim drawingRevision As String
        
        If Not drawingPropMgr Is Nothing Then
            drawingPropMgr.Get3 "Revision", False, "", drawingRevision
            If Trim(drawingRevision) <> "" Then
                On Error Resume Next
                ConfigRevisions.Add drawingRevision
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Else
        Dim configNames As Variant
        Dim i As Long
        
        On Error Resume Next
        configNames = swModel.GetConfigurationNames()
        
        If Err.Number = 0 And Not IsEmpty(configNames) Then
            Dim swConfigMgr As SldWorks.ConfigurationManager
            Set swConfigMgr = swModel.ConfigurationManager
            
            If Not swConfigMgr Is Nothing And Not swConfigMgr.ActiveConfiguration Is Nothing Then
                Dim activeConfigName As String
                activeConfigName = swConfigMgr.ActiveConfiguration.Name
                
                ProcessConfigRevision activeConfigName
                
                For i = 0 To UBound(configNames)
                    If configNames(i) <> activeConfigName Then
                        ProcessConfigRevision configNames(i)
                    End If
                Next i
            End If
        End If
        Err.Clear
        On Error GoTo 0
        
        If ConfigRevisions.count = 0 Then
            Dim defaultPropMgr As SldWorks.CustomPropertyManager
            Set defaultPropMgr = swModel.Extension.CustomPropertyManager("")
            Dim defaultRevision As String
            
            If Not defaultPropMgr Is Nothing Then
                defaultPropMgr.Get3 "Revision", False, "", defaultRevision
                If Trim(defaultRevision) <> "" Then
                    ConfigRevisions.Add defaultRevision
                End If
            End If
        End If
    End If
End Sub

Private Sub ProcessConfigRevision(ByVal configName As String)
    Dim cusPropMgr As SldWorks.CustomPropertyManager
    Dim configRevision As String
    
    Set cusPropMgr = swModel.Extension.CustomPropertyManager(configName)
    
    If Not cusPropMgr Is Nothing Then
        cusPropMgr.Get3 "Revision", False, "", configRevision
        If Trim(configRevision) <> "" Then
            On Error Resume Next
            ConfigRevisions.Add configRevision
            Err.Clear
            On Error GoTo 0
        End If
    End If
End Sub

Public Function GetProductSubfolder(fullPath As String) As String
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
    
    Debug.Print "Product subfolder identified: " & result
    GetProductSubfolder = result
End Function

Public Function ConnectToPDMVault() As Boolean
    On Error Resume Next
    Set pdmVault = New EdmVault5
    pdmVault.LoginAuto "NMT_PDM", 0
    If pdmVault.IsLoggedIn Then
        ConnectToPDMVault = True
        Debug.Print "Logged into NMT_PDM vault successfully."
        
        Set cachedFolders = CreateObject("Scripting.Dictionary")
        cachedFolders.CompareMode = vbTextCompare
    Else
        ConnectToPDMVault = False
        Debug.Print "Failed to log into NMT_PDM vault."
    End If
    On Error GoTo 0
End Function

Public Function EnsureFolderExists(parentFolderPath As String, folderName As String, folderPath As String) As Boolean
    Dim pdmParentFolder As Object
    Dim pdmFolder As Object
    Dim success As Boolean
    
    If Not cachedFolders Is Nothing Then
        If cachedFolders.exists(folderPath) Then
            Debug.Print "Folder found in cache: " & folderPath
            EnsureFolderExists = True
            Exit Function
        End If
    End If
    
    success = False
    Debug.Print "Checking if folder already exists: " & folderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        Debug.Print folderName & " folder already exists in vault: " & pdmFolder.LocalPath
        success = True
        
        If Not cachedFolders Is Nothing Then
            On Error Resume Next
            cachedFolders.Add folderPath, True
            Err.Clear
            On Error GoTo 0
        End If
        
        EnsureFolderExists = success
        Exit Function
    End If
    
    Debug.Print "Checking parent folder path: " & parentFolderPath
    Set pdmParentFolder = pdmVault.GetFolderFromPath(parentFolderPath)
    If pdmParentFolder Is Nothing Then
        Debug.Print "Parent folder not found in vault: " & parentFolderPath
        EnsureFolderExists = success
        Exit Function
    End If
    
    Debug.Print "Target folder not found, attempting to create: " & folderName
    On Error Resume Next
    Dim rootFolder As Object
    Set rootFolder = pdmVault.rootFolder
    If Not rootFolder Is Nothing Then
        Dim vaultPath As String
        vaultPath = Replace(folderPath, "C:\NMT_PDM\", "\")
        rootFolder.CreateFolderPath vaultPath, 0
        If Err.Number <> 0 Then
            Debug.Print "CreateFolderPath from root failed: " & Err.Description & " (Error #" & Err.Number & ")"
            Err.Clear
            EnsureFolderExists = False
            Exit Function
        End If
    Else
        Debug.Print "Could not get root folder"
        EnsureFolderExists = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' ===================================================================
    ' OPTIMIZATION: Removed the slow polling Do...Loop
    ' The CreateFolderPath command is synchronous, so we can
    ' immediately try to get the folder. This removes the
    ' 1-second delay for every new folder creation.
    ' ===================================================================
    On Error Resume Next
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    On Error GoTo 0

    If Not pdmFolder Is Nothing Then
        Debug.Print "Created " & folderName & " folder in vault: " & folderPath
        success = True
    Else
        Debug.Print "Failed to create " & folderName & " folder in vault: " & folderPath
        MsgBox "Failed to create folder in vault: " & folderPath, vbCritical
        EnsureFolderExists = False
        Exit Function
    End If
    
    ' --- This local folder creation logic is fine ---
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            Debug.Print "Error creating local folder: " & Err.Description
            Err.Clear
            success = False ' Local creation can fail, but vault might be ok
        Else
            Debug.Print "Created " & folderName & " folder locally: " & folderPath
        End If
        On Error GoTo 0
    Else
        Debug.Print folderName & " folder already exists locally: " & folderPath
    End If
    
    If success And Not cachedFolders Is Nothing Then
        On Error Resume Next
        cachedFolders.Add folderPath, True
        Err.Clear
        On Error GoTo 0
        Debug.Print "Added folder to cache: " & folderPath
    End If
    
    EnsureFolderExists = success
End Function

Public Function IsLibraryPathCheck(pathToCheck As String) As Boolean
    Dim originalLibraryCheck As Boolean
    originalLibraryCheck = (InStr(1, pathToCheck, "C:\NMT_PDM\Libraries", vbTextCompare) > 0)
    
    Dim customLibraryCheck As Boolean
    customLibraryCheck = (InStr(1, pathToCheck, "C:\NMT_PDM\Projects\Capital\8546 - TestDL - testing\3 - Design\Models\Library Pack and Go\New Folder", vbTextCompare) > 0)
    
    IsLibraryPathCheck = originalLibraryCheck Or customLibraryCheck
End Function

Public Function HasFeatureRebuildErrors(swModel As SldWorks.ModelDoc2) As Boolean
    HasFeatureRebuildErrors = False
    Exit Function
End Function




' ============================================================================
' BATCH PDF RUN (COMPLETE FUNCTION)
' ============================================================================
Public Sub BatchPDFRun()
    Dim batchFolder As String
    Dim designFolder As String
    Dim drawingsFolder As String
    Dim obsoleteFolder As String
    Dim filesCollection As Collection
    Dim localFilePath As Variant
    Dim fullPath As String
    Dim pdfPath As String
    Dim swDoc As SldWorks.ModelDoc2
    Dim errs As Long
    Dim warns As Long
    Dim errorList As Collection
    Dim successList As Collection
    Dim warningList As Collection
    Dim isLibraryPath As Boolean
    Dim fileCount As Long
    Dim currentFile As String
    Dim pattern As Variant
    Dim processedFiles As Object
    Dim includeSubfolders As Boolean
    Dim currentFileNum As Long
    Dim localFileName As String
    Dim checkedFolders As Object
    Dim parentPath As String
    Dim folderName As String
    Dim folderExists As Boolean
    Dim originalVisibility As Boolean
    Dim subfolderPrompt As VbMsgBoxResult
    Dim currentDrawingsFolder As String
    Dim currentObsoleteFolder As String
    Dim currentFilePathOnly As String
    Dim productSubfolder As String
    Dim pos As Long
    Dim subfolderPath As String
    Dim modelUpdated As Boolean
    Dim errorKey As String
    Dim item As Variant
    
    Const TARGET_UPDATES As Long = 10
    Dim updateInterval As Long
    
    CancelOperation = False
    InitializeSettings

    ' Initialize debug logging
    LogInit "BatchPDFRun"
    LogSectionStart "Batch PDF Run - All Files Mode"

    ' Note: Output mode (default vs custom folder) is set via PromptForCustomOutputFolder later

    Set processedFiles = CreateObject("Scripting.Dictionary")
    processedFiles.CompareMode = vbTextCompare
    
    Set errorList = New Collection
    Set successList = New Collection
    Set warningList = New Collection
    
    fileCount = 0
    
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            MsgBox "Failed to initialize SolidWorks application. Please ensure SolidWorks is running.", vbCritical
            Exit Sub
        End If
    End If
    
    Debug.Print "=== SHOWING SUBFOLDER PROMPT ==="
    subfolderPrompt = MsgBox("Do you want to include subfolders when searching for drawing files?" & vbCrLf & vbCrLf & _
                            "� YES - Search all subdirectories recursively" & vbCrLf & _
                            "� NO - Search only the top-level folder you select", _
                            vbQuestion + vbYesNo + vbDefaultButton1, _
                            "Include Subfolders in Search?")

    If subfolderPrompt = vbYes Then
        includeSubfolders = True
        Debug.Print "User chose: YES - Include subfolders"
        LogMessage "Include subfolders: Yes"
    Else
        includeSubfolders = False
        Debug.Print "User chose: NO - Top folder only"
        LogMessage "Include subfolders: No"
    End If

    ' Prompt for custom output folder (optional)
    PromptForCustomOutputFolder

    batchFolder = SelectFolder("Select the TOP folder containing .slddrw files for batch processing")
    If batchFolder = "" Then
        MsgBox "No folder selected. Exiting batch mode.", vbInformation
        LogWarning "No folder selected - exiting"
        LogClose
        Exit Sub
    End If

    ' Remove any surrounding quotes from the path
    If Left(batchFolder, 1) = """" Then batchFolder = Mid(batchFolder, 2)
    If Right(batchFolder, 1) = """" Then batchFolder = Left(batchFolder, Len(batchFolder) - 1)

    ' Ensure trailing backslash
    If Right(batchFolder, 1) <> "\" Then batchFolder = batchFolder & "\"

    ' Validate the folder exists
    On Error Resume Next
    Dim folderAttrCheck As Long
    folderAttrCheck = GetAttr(Left(batchFolder, Len(batchFolder) - 1))
    If Err.Number <> 0 Then
        MsgBox "Invalid folder path: " & batchFolder & vbCrLf & vbCrLf & _
               "Please ensure the path exists and is accessible.", vbExclamation
        LogError "Invalid folder path: " & batchFolder
        LogClose
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "Batch folder selected: " & batchFolder
    Debug.Print "Include subfolders: " & includeSubfolders
    LogMessage "Batch folder: " & batchFolder
    Debug.Print "Manual output mode: " & bManualMode
    If bManualMode Then Debug.Print "Manual output path: " & manualOutputPath

    If Not ConnectToPDMVault() Then
        MsgBox "Failed to connect to NMT_PDM vault.", vbCritical
        LogError "Failed to connect to PDM vault"
        LogClose
        Exit Sub
    End If
    LogPDMConnection "NMT_PDM", True

    On Error Resume Next
    originalVisibility = swApp.Visible
    swApp.Visible = False
    swApp.UserControl = False
    Debug.Print "SolidWorks application hidden for batch processing."
    On Error GoTo 0
    
    Set filesCollection = New Collection
    Set checkedFolders = CreateObject("Scripting.Dictionary")
    checkedFolders.CompareMode = vbTextCompare

    Debug.Print "=== COLLECTING DRAWING FILES ==="
    LogSectionStart "Collecting Drawing Files"
    If includeSubfolders Then
        Debug.Print "Scanning directories recursively..."
        CollectFilesWithSubfolders batchFolder, "*.slddrw", filesCollection, processedFiles
    Else
        Dim extensionPatterns As Variant
        extensionPatterns = Array("*.slddrw", "*.SLDDRW", "*.SldDrw")
        For Each pattern In extensionPatterns
            Debug.Print "Searching with pattern: " & pattern
            currentFile = Dir(batchFolder & pattern)
            While currentFile <> ""
                Dim fullFilePathCheck As String
                fullFilePathCheck = LCase(batchFolder & currentFile)
                If Not processedFiles.exists(fullFilePathCheck) Then
                    filesCollection.Add batchFolder & currentFile
                    processedFiles.Add fullFilePathCheck, currentFile
                End If
                currentFile = Dir
            Wend
        Next pattern
    End If
    fileCount = filesCollection.count
    Debug.Print "Total unique drawing files found: " & fileCount
    LogFolderCollection batchFolder, fileCount, includeSubfolders
    LogSectionEnd fileCount & " files found"

    If fileCount = 0 Then
        MsgBox "No drawing files (.slddrw) found...", vbExclamation
        LogWarning "No drawing files found"
        GoTo Cleanup
    End If
    
    updateInterval = fileCount \ TARGET_UPDATES
    If updateInterval < 1 Then updateInterval = 1  ' Update every file if less than 10 files
    Debug.Print "Progress Bar Update Interval: Every " & updateInterval & " files"

    ProgressForm.Show vbModeless
    ProgressForm.Caption = "PDF Creation Progress"
    ProgressForm.UpdateProgress 0, fileCount, "Starting PDF creation..."
    DoEvents

    LogSeparator
    LogSectionStart "Processing Files"
    currentFileNum = 0
    For Each localFilePath In filesCollection
        If CancelOperation Then
            MsgBox "Operation canceled by user...", vbInformation
            LogWarning "Operation canceled by user"
            GoTo Cleanup
        End If

        currentFileNum = currentFileNum + 1
        fullPath = CStr(localFilePath)
        localFileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
        localFileName = Left(localFileName, InStrRev(localFileName, ".") - 1)

        Debug.Print "Processing drawing (" & currentFileNum & "/" & fileCount & "): " & fullPath
        LogFileProcess currentFileNum, fileCount, fullPath, "Starting"

        ' Always update progress for every file (removed interval check for responsiveness)
        ProgressForm.UpdateProgress currentFileNum, fileCount, Mid(fullPath, InStrRev(fullPath, "\") + 1)
        DoEvents

        On Error Resume Next
        Set swDoc = swApp.OpenDoc6(fullPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly Or swOpenDocOptions_SupressRebuild, "", errs, warns)
        If Err.Number <> 0 Or swDoc Is Nothing Then
            errorList.Add "Failed to open drawing '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "': Err " & errs & ", Warns " & warns & ", VBA " & Err.Description, "OPEN_FAIL_" & localFileName
            LogError "Failed to open drawing", errs, Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextDrawing
        End If
        On Error GoTo 0
        
        Set swModel = swDoc
        Set swDraw = swDoc
        
        Set ConfigRevisions = New Collection
        GetAllConfigRevisions
        RevValue = ""
        If ConfigRevisions.count > 0 Then RevValue = ConfigRevisions.item(1)
        If RevValue = "" Then
            Dim cusPropMgrBatch As SldWorks.CustomPropertyManager
            Set cusPropMgrBatch = swDoc.Extension.CustomPropertyManager("")
            Dim tempRevValueBatch As String
            If Not cusPropMgrBatch Is Nothing Then cusPropMgrBatch.Get3 "Revision", False, "", tempRevValueBatch
            RevValue = Trim(tempRevValueBatch)
            Set cusPropMgrBatch = Nothing
        End If
        If RevValue = "" Then
            errorList.Add "No revision found for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "NO_REV_" & localFileName
            GoTo CloseDrawing
        End If
        
        currentFilePathOnly = Left(fullPath, InStrRev(fullPath, "\"))
        
        If bManualMode Then
            currentDrawingsFolder = manualOutputPath
            currentObsoleteFolder = manualOutputPath & "Obsolete\"
            If Not checkedFolders.exists(currentObsoleteFolder) Then
                parentPath = manualOutputPath
                folderName = "Obsolete"
                folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                If folderExists Then
                    checkedFolders.Add currentObsoleteFolder, True
                Else
                    errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                    GoTo CloseDrawing
                End If
            End If
        Else
            isLibraryPath = IsLibraryPathCheck(fullPath)
            If isLibraryPath Then
                currentDrawingsFolder = currentFilePathOnly & "Drawings\"
                currentObsoleteFolder = currentDrawingsFolder & "Obsolete\"
                
                If Not checkedFolders.exists(currentDrawingsFolder) Then
                    parentPath = currentFilePathOnly
                    folderName = "Drawings"
                    folderExists = EnsureFolderExists(parentPath, folderName, currentDrawingsFolder)
                    If folderExists Then
                        checkedFolders.Add currentDrawingsFolder, True
                    Else
                        errorList.Add "Failed ensure folder: " & currentDrawingsFolder, "FOLDER_FAIL_" & folderName
                        GoTo CloseDrawing
                    End If
                End If
                
                If Not checkedFolders.exists(currentObsoleteFolder) Then
                    parentPath = currentDrawingsFolder
                    folderName = "Obsolete"
                    folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                    If folderExists Then
                        checkedFolders.Add currentObsoleteFolder, True
                    Else
                        errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                        GoTo CloseDrawing
                    End If
                End If
            Else
                pos = InStrRev(currentFilePathOnly, "3 - Design\")
                If pos > 0 Then
                    designFolder = Left(currentFilePathOnly, pos + Len("3 - Design") - 1) & "\"
                    currentDrawingsFolder = designFolder & "Drawings\"
                    currentObsoleteFolder = currentDrawingsFolder & "Obsolete\"
                    If Not checkedFolders.exists(currentDrawingsFolder) Then
                        parentPath = designFolder
                        folderName = "Drawings"
                        folderExists = EnsureFolderExists(parentPath, folderName, currentDrawingsFolder)
                        If folderExists Then
                            checkedFolders.Add currentDrawingsFolder, True
                        Else
                            errorList.Add "Failed ensure folder: " & currentDrawingsFolder, "FOLDER_FAIL_" & folderName
                            GoTo CloseDrawing
                        End If
                    End If
                    If Not checkedFolders.exists(currentObsoleteFolder) Then
                        parentPath = currentDrawingsFolder
                        folderName = "Obsolete"
                        folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                        If folderExists Then
                            checkedFolders.Add currentObsoleteFolder, True
                        Else
                            errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                            GoTo CloseDrawing
                        End If
                    End If
                    
                    If InStr(1, fullPath, "\Products\", vbTextCompare) > 0 Or InStr(1, fullPath, "\Models\", vbTextCompare) > 0 Then
                        productSubfolder = GetProductSubfolder(fullPath)
                        If productSubfolder <> "" Then
                            subfolderPath = currentDrawingsFolder & productSubfolder & "\"
                            If Not checkedFolders.exists(subfolderPath) Then
                                parentPath = currentDrawingsFolder
                                folderName = productSubfolder
                                folderExists = EnsureFolderExists(parentPath, folderName, subfolderPath)
                                If folderExists Then
                                    checkedFolders.Add subfolderPath, True
                                Else
                                    errorList.Add "Failed ensure folder: " & subfolderPath, "FOLDER_FAIL_" & folderName
                                    GoTo CloseDrawing
                                End If
                            End If
                            currentObsoleteFolder = subfolderPath & "Obsolete\"
                            If Not checkedFolders.exists(currentObsoleteFolder) Then
                                parentPath = subfolderPath
                                folderName = "Obsolete"
                                folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                                If folderExists Then
                                    checkedFolders.Add currentObsoleteFolder, True
                                Else
                                    errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                                    GoTo CloseDrawing
                                End If
                            End If
                            currentDrawingsFolder = subfolderPath
                        End If
                    End If
                Else
                    errorList.Add "Could not find '3 - Design' in path for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "PATH_FAIL_" & localFileName
                    GoTo CloseDrawing
                End If
            End If
        End If
        
        pdfPath = currentDrawingsFolder & localFileName & "_Rev" & RevValue & ".pdf"
        Debug.Print "PDF Target Path: " & pdfPath
        
        modelUpdated = SyncDrawingWithModel(fullPath, currentDrawingsFolder, currentObsoleteFolder, errorList, warningList, True)
        
        If modelUpdated Then
            HandlePDFCheckIn pdfPath, RevValue, currentDrawingsFolder, currentObsoleteFolder, True
            
            If Dir(pdfPath) <> "" Then
                successList.Add pdfPath
                Debug.Print "Added to success list: " & pdfPath
            Else
                errorKey = "PDF_FAIL_" & localFileName
                On Error Resume Next
                errorList.Add "PDF creation failed or file not found after attempt for: '" & localFileName & "' at path " & pdfPath, errorKey
                If Err.Number = 0 Then Debug.Print "Error logged: PDF creation failed or file not found after attempt for: " & localFileName
                Err.Clear
                On Error GoTo 0
            End If
        Else
            errorList.Add "Skipped PDF creation due to model sync issue for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "SYNC_FAIL_" & localFileName
            Debug.Print "Error logged: Skipped PDF creation due to model sync issue for: " & localFileName
        End If

CloseDrawing:
        If Not swDoc Is Nothing Then
            On Error Resume Next
            swApp.CloseDoc fullPath
            On Error GoTo 0
            Set swDoc = Nothing
            Set swDraw = Nothing
            Set swModel = Nothing
        End If

NextDrawing:
    Next localFilePath
    
    If Not CancelOperation Then
        On Error Resume Next
        If Not ProgressForm Is Nothing Then
            ProgressForm.UpdateProgress fileCount, fileCount, "Batch Complete!"
            DoEvents
        End If
        On Error GoTo 0
        Sleep 300
    End If
    
    On Error Resume Next
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    On Error GoTo 0
    
    ShowBatchResults successList, errorList, warningList, fileCount

    Debug.Print "Batch PDF run completed."
    LogSectionEnd "Processing complete"
    LogBatchSummary fileCount, successList.count, errorList.count, warningList.count
    LogClose

Cleanup:
    On Error Resume Next
    If Not swApp Is Nothing Then
        swApp.Visible = originalVisibility
        swApp.UserControl = True
        Debug.Print "Restored SolidWorks visibility."
    Else
        Debug.Print "swApp object was not available to restore visibility."
    End If
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    On Error GoTo 0

    Set checkedFolders = Nothing
    Debug.Print "Cleanup complete."
End Sub

' ============================================================================
' BATCH PDF RUN - CHECKED OUT FILES ONLY
' ============================================================================
Public Sub BatchPDFRunCheckedOutOnly()
    Dim batchFolder As String
    Dim designFolder As String
    Dim drawingsFolder As String
    Dim obsoleteFolder As String
    Dim filesCollection As Collection
    Dim filteredFiles As Collection
    Dim localFilePath As Variant
    Dim fullPath As String
    Dim pdfPath As String
    Dim swDoc As SldWorks.ModelDoc2
    Dim errs As Long
    Dim warns As Long
    Dim errorList As Collection
    Dim successList As Collection
    Dim warningList As Collection
    Dim isLibraryPath As Boolean
    Dim fileCount As Long
    Dim currentFile As String
    Dim pattern As Variant
    Dim processedFiles As Object
    Dim includeSubfolders As Boolean
    Dim currentFileNum As Long
    Dim localFileName As String
    Dim checkedFolders As Object
    Dim parentPath As String
    Dim folderName As String
    Dim folderExists As Boolean
    Dim originalVisibility As Boolean
    Dim subfolderPrompt As VbMsgBoxResult
    Dim currentDrawingsFolder As String
    Dim currentObsoleteFolder As String
    Dim currentFilePathOnly As String
    Dim productSubfolder As String
    Dim pos As Long
    Dim subfolderPath As String
    Dim modelUpdated As Boolean
    Dim errorKey As String
    Dim item As Variant
    Dim pdmFileCheck As Object
    Dim lockedByUserCheck As Object
    Dim totalFound As Long
    Dim checkedOutCount As Long

    Const TARGET_UPDATES As Long = 10
    Dim updateInterval As Long

    CancelOperation = False
    InitializeSettings

    ' Initialize debug logging
    LogInit "BatchPDFRun_CheckedOutOnly"
    LogSectionStart "Batch PDF Run - Checked Out Files Only"

    ' Note: Output mode (default vs custom folder) is set via PromptForCustomOutputFolder later

    Set processedFiles = CreateObject("Scripting.Dictionary")
    processedFiles.CompareMode = vbTextCompare

    Set errorList = New Collection
    Set successList = New Collection
    Set warningList = New Collection

    fileCount = 0

    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            MsgBox "Failed to initialize SolidWorks application. Please ensure SolidWorks is running.", vbCritical
            LogError "Failed to initialize SolidWorks"
            LogClose
            Exit Sub
        End If
    End If

    ' Connect to PDM vault first (needed for checkout check)
    If Not ConnectToPDMVault() Then
        MsgBox "Failed to connect to NMT_PDM vault.", vbCritical
        LogError "Failed to connect to PDM vault"
        LogClose
        Exit Sub
    End If
    LogPDMConnection "NMT_PDM", True

    Debug.Print "=== SHOWING SUBFOLDER PROMPT ==="
    subfolderPrompt = MsgBox("Do you want to include subfolders when searching for drawing files?" & vbCrLf & vbCrLf & _
                            "YES - Search all subdirectories recursively" & vbCrLf & _
                            "NO - Search only the top-level folder you select", _
                            vbQuestion + vbYesNo + vbDefaultButton1, _
                            "Include Subfolders in Search?")

    If subfolderPrompt = vbYes Then
        includeSubfolders = True
        Debug.Print "User chose: YES - Include subfolders"
        LogMessage "Include subfolders: Yes"
    Else
        includeSubfolders = False
        Debug.Print "User chose: NO - Top folder only"
        LogMessage "Include subfolders: No"
    End If

    ' Prompt for custom output folder (optional)
    PromptForCustomOutputFolder

    batchFolder = SelectFolder("Select the TOP folder containing .slddrw files for batch processing")
    If batchFolder = "" Then
        MsgBox "No folder selected. Exiting batch mode.", vbInformation
        LogWarning "No folder selected - exiting"
        LogClose
        Exit Sub
    End If

    ' Remove any surrounding quotes from the path
    If Left(batchFolder, 1) = """" Then batchFolder = Mid(batchFolder, 2)
    If Right(batchFolder, 1) = """" Then batchFolder = Left(batchFolder, Len(batchFolder) - 1)

    ' Ensure trailing backslash
    If Right(batchFolder, 1) <> "\" Then batchFolder = batchFolder & "\"

    ' Validate the folder exists
    On Error Resume Next
    Dim folderAttr As Long
    folderAttr = GetAttr(Left(batchFolder, Len(batchFolder) - 1))
    If Err.Number <> 0 Then
        MsgBox "Invalid folder path: " & batchFolder & vbCrLf & vbCrLf & _
               "Please ensure the path exists and is accessible.", vbExclamation
        LogError "Invalid folder path: " & batchFolder
        LogClose
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "Batch folder selected: " & batchFolder
    Debug.Print "Include subfolders: " & includeSubfolders
    Debug.Print "Mode: CHECKED-OUT FILES ONLY"
    LogMessage "Batch folder: " & batchFolder
    LogMessage "Mode: Checked-out files only"

    On Error Resume Next
    originalVisibility = swApp.Visible
    swApp.Visible = False
    swApp.UserControl = False
    Debug.Print "SolidWorks application hidden for batch processing."
    On Error GoTo 0

    Set filesCollection = New Collection
    Set filteredFiles = New Collection
    Set checkedFolders = CreateObject("Scripting.Dictionary")
    checkedFolders.CompareMode = vbTextCompare

    Debug.Print "=== COLLECTING DRAWING FILES ==="
    LogSectionStart "Collecting Drawing Files"
    If includeSubfolders Then
        Debug.Print "Scanning directories recursively..."
        CollectFilesWithSubfolders batchFolder, "*.slddrw", filesCollection, processedFiles
    Else
        Dim extensionPatterns As Variant
        extensionPatterns = Array("*.slddrw", "*.SLDDRW", "*.SldDrw")
        For Each pattern In extensionPatterns
            Debug.Print "Searching with pattern: " & pattern
            currentFile = Dir(batchFolder & pattern)
            While currentFile <> ""
                Dim fullFilePathCheck As String
                fullFilePathCheck = LCase(batchFolder & currentFile)
                If Not processedFiles.exists(fullFilePathCheck) Then
                    filesCollection.Add batchFolder & currentFile
                    processedFiles.Add fullFilePathCheck, currentFile
                End If
                currentFile = Dir
            Wend
        Next pattern
    End If
    totalFound = filesCollection.count
    Debug.Print "Total unique drawing files found: " & totalFound
    LogFolderCollection batchFolder, totalFound, includeSubfolders
    LogSectionEnd totalFound & " files found"

    If totalFound = 0 Then
        MsgBox "No drawing files (.slddrw) found in the selected folder.", vbExclamation
        LogWarning "No drawing files found"
        GoTo Cleanup
    End If

    ' Filter to only files checked out by current user
    Debug.Print "=== FILTERING TO CHECKED-OUT FILES ==="
    LogSectionStart "Filtering to Checked-Out Files"
    checkedOutCount = 0
    For Each localFilePath In filesCollection
        On Error Resume Next
        Set pdmFileCheck = pdmVault.GetFileFromPath(CStr(localFilePath))
        If Err.Number = 0 And Not pdmFileCheck Is Nothing Then
            If pdmFileCheck.IsLocked Then
                Set lockedByUserCheck = pdmFileCheck.lockedByUser
                If Err.Number = 0 And Not lockedByUserCheck Is Nothing Then
                    If lockedByUserCheck.Name = pdmVault.CurrentUser.Name Then
                        filteredFiles.Add localFilePath
                        checkedOutCount = checkedOutCount + 1
                        Debug.Print "  INCLUDED (checked out by me): " & localFilePath
                    Else
                        Debug.Print "  Skipped (checked out by " & lockedByUserCheck.Name & "): " & localFilePath
                    End If
                End If
            Else
                Debug.Print "  Skipped (not checked out): " & localFilePath
            End If
        Else
            Debug.Print "  Skipped (not in vault): " & localFilePath
        End If
        Err.Clear
        On Error GoTo 0
    Next localFilePath

    fileCount = filteredFiles.count
    Debug.Print "Files checked out by me: " & fileCount & " of " & totalFound
    LogFilterResults totalFound, fileCount, "Checked out by current user"
    LogSectionEnd fileCount & " files to process"

    If fileCount = 0 Then
        MsgBox "No drawing files are currently checked out by you in the selected folder." & vbCrLf & vbCrLf & _
               "Total drawings found: " & totalFound, vbInformation, "No Checked-Out Files"
        LogWarning "No checked-out files found"
        GoTo Cleanup
    End If

    ' Confirm with user
    Dim confirmMsg As VbMsgBoxResult
    confirmMsg = MsgBox("Found " & fileCount & " drawing(s) checked out by you." & vbCrLf & _
                        "(Total drawings in folder: " & totalFound & ")" & vbCrLf & vbCrLf & _
                        "Do you want to proceed with PDF creation?", _
                        vbQuestion + vbYesNo, "Confirm Batch Processing")
    If confirmMsg <> vbYes Then
        LogWarning "User declined to proceed"
        GoTo Cleanup
    End If

    LogSeparator
    LogSectionStart "Processing Files"
    updateInterval = fileCount \ TARGET_UPDATES
    If updateInterval < 1 Then updateInterval = 1
    Debug.Print "Progress Bar Update Interval: Every " & updateInterval & " files"

    ProgressForm.Show vbModeless
    ProgressForm.Caption = "PDF Creation Progress (Checked-Out Files)"
    ProgressForm.UpdateProgress 0, fileCount, "Starting PDF creation..."
    DoEvents

    currentFileNum = 0
    For Each localFilePath In filteredFiles
        If CancelOperation Then
            MsgBox "Operation canceled by user...", vbInformation
            LogWarning "Operation canceled by user"
            GoTo Cleanup
        End If

        currentFileNum = currentFileNum + 1
        fullPath = CStr(localFilePath)
        localFileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
        localFileName = Left(localFileName, InStrRev(localFileName, ".") - 1)
        LogFileProcess currentFileNum, fileCount, fullPath, "Starting"

        Debug.Print "Processing drawing (" & currentFileNum & "/" & fileCount & "): " & fullPath

        ' Always update progress for every file (removed interval check for responsiveness)
        ProgressForm.UpdateProgress currentFileNum, fileCount, Mid(fullPath, InStrRev(fullPath, "\") + 1)
        DoEvents

        On Error Resume Next
        ' Open without ReadOnly flag since these files are checked out by current user
        Set swDoc = swApp.OpenDoc6(fullPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_SupressRebuild, "", errs, warns)
        If Err.Number <> 0 Or swDoc Is Nothing Then
            errorList.Add "Failed to open drawing '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "': Err " & errs & ", Warns " & warns & ", VBA " & Err.Description, "OPEN_FAIL_" & localFileName
            LogError "Failed to open drawing", errs, Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextDrawingCO
        End If
        On Error GoTo 0

        Set swModel = swDoc
        Set swDraw = swDoc

        Set ConfigRevisions = New Collection
        GetAllConfigRevisions
        RevValue = ""
        If ConfigRevisions.count > 0 Then RevValue = ConfigRevisions.item(1)
        If RevValue = "" Then
            Dim cusPropMgrBatch As SldWorks.CustomPropertyManager
            Set cusPropMgrBatch = swDoc.Extension.CustomPropertyManager("")
            Dim tempRevValueBatch As String
            If Not cusPropMgrBatch Is Nothing Then cusPropMgrBatch.Get3 "Revision", False, "", tempRevValueBatch
            RevValue = Trim(tempRevValueBatch)
            Set cusPropMgrBatch = Nothing
        End If
        If RevValue = "" Then
            errorList.Add "No revision found for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "NO_REV_" & localFileName
            GoTo CloseDrawingCO
        End If

        currentFilePathOnly = Left(fullPath, InStrRev(fullPath, "\"))

        If bManualMode Then
            currentDrawingsFolder = manualOutputPath
            currentObsoleteFolder = manualOutputPath & "Obsolete\"
            If Not checkedFolders.exists(currentObsoleteFolder) Then
                parentPath = manualOutputPath
                folderName = "Obsolete"
                folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                If folderExists Then
                    checkedFolders.Add currentObsoleteFolder, True
                Else
                    errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                    GoTo CloseDrawingCO
                End If
            End If
        Else
            isLibraryPath = IsLibraryPathCheck(fullPath)
            If isLibraryPath Then
                currentDrawingsFolder = currentFilePathOnly & "Drawings\"
                currentObsoleteFolder = currentDrawingsFolder & "Obsolete\"

                If Not checkedFolders.exists(currentDrawingsFolder) Then
                    parentPath = currentFilePathOnly
                    folderName = "Drawings"
                    folderExists = EnsureFolderExists(parentPath, folderName, currentDrawingsFolder)
                    If folderExists Then
                        checkedFolders.Add currentDrawingsFolder, True
                    Else
                        errorList.Add "Failed ensure folder: " & currentDrawingsFolder, "FOLDER_FAIL_" & folderName
                        GoTo CloseDrawingCO
                    End If
                End If

                If Not checkedFolders.exists(currentObsoleteFolder) Then
                    parentPath = currentDrawingsFolder
                    folderName = "Obsolete"
                    folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                    If folderExists Then
                        checkedFolders.Add currentObsoleteFolder, True
                    Else
                        errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                        GoTo CloseDrawingCO
                    End If
                End If
            Else
                pos = InStrRev(currentFilePathOnly, "3 - Design\")
                If pos > 0 Then
                    designFolder = Left(currentFilePathOnly, pos + Len("3 - Design") - 1) & "\"
                    currentDrawingsFolder = designFolder & "Drawings\"
                    currentObsoleteFolder = currentDrawingsFolder & "Obsolete\"
                    If Not checkedFolders.exists(currentDrawingsFolder) Then
                        parentPath = designFolder
                        folderName = "Drawings"
                        folderExists = EnsureFolderExists(parentPath, folderName, currentDrawingsFolder)
                        If folderExists Then
                            checkedFolders.Add currentDrawingsFolder, True
                        Else
                            errorList.Add "Failed ensure folder: " & currentDrawingsFolder, "FOLDER_FAIL_" & folderName
                            GoTo CloseDrawingCO
                        End If
                    End If
                    If Not checkedFolders.exists(currentObsoleteFolder) Then
                        parentPath = currentDrawingsFolder
                        folderName = "Obsolete"
                        folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                        If folderExists Then
                            checkedFolders.Add currentObsoleteFolder, True
                        Else
                            errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                            GoTo CloseDrawingCO
                        End If
                    End If

                    If InStr(1, fullPath, "\Products\", vbTextCompare) > 0 Or InStr(1, fullPath, "\Models\", vbTextCompare) > 0 Then
                        productSubfolder = GetProductSubfolder(fullPath)
                        If productSubfolder <> "" Then
                            subfolderPath = currentDrawingsFolder & productSubfolder & "\"
                            If Not checkedFolders.exists(subfolderPath) Then
                                parentPath = currentDrawingsFolder
                                folderName = productSubfolder
                                folderExists = EnsureFolderExists(parentPath, folderName, subfolderPath)
                                If folderExists Then
                                    checkedFolders.Add subfolderPath, True
                                Else
                                    errorList.Add "Failed ensure folder: " & subfolderPath, "FOLDER_FAIL_" & folderName
                                    GoTo CloseDrawingCO
                                End If
                            End If
                            currentObsoleteFolder = subfolderPath & "Obsolete\"
                            If Not checkedFolders.exists(currentObsoleteFolder) Then
                                parentPath = subfolderPath
                                folderName = "Obsolete"
                                folderExists = EnsureFolderExists(parentPath, folderName, currentObsoleteFolder)
                                If folderExists Then
                                    checkedFolders.Add currentObsoleteFolder, True
                                Else
                                    errorList.Add "Failed ensure folder: " & currentObsoleteFolder, "FOLDER_FAIL_" & folderName
                                    GoTo CloseDrawingCO
                                End If
                            End If
                            currentDrawingsFolder = subfolderPath
                        End If
                    End If
                Else
                    errorList.Add "Could not find '3 - Design' in path for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "PATH_FAIL_" & localFileName
                    GoTo CloseDrawingCO
                End If
            End If
        End If

        pdfPath = currentDrawingsFolder & localFileName & "_Rev" & RevValue & ".pdf"
        Debug.Print "PDF Target Path: " & pdfPath

        modelUpdated = SyncDrawingWithModel(fullPath, currentDrawingsFolder, currentObsoleteFolder, errorList, warningList, True)

        If modelUpdated Then
            HandlePDFCheckIn pdfPath, RevValue, currentDrawingsFolder, currentObsoleteFolder, True

            If Dir(pdfPath) <> "" Then
                successList.Add pdfPath
                Debug.Print "Added to success list: " & pdfPath
            Else
                errorKey = "PDF_FAIL_" & localFileName
                On Error Resume Next
                errorList.Add "PDF creation failed or file not found after attempt for: '" & localFileName & "' at path " & pdfPath, errorKey
                If Err.Number = 0 Then Debug.Print "Error logged: PDF creation failed or file not found after attempt for: " & localFileName
                Err.Clear
                On Error GoTo 0
            End If
        Else
            errorList.Add "Skipped PDF creation due to model sync issue for: '" & Mid(fullPath, InStrRev(fullPath, "\") + 1) & "'", "SYNC_FAIL_" & localFileName
            Debug.Print "Error logged: Skipped PDF creation due to model sync issue for: " & localFileName
        End If

CloseDrawingCO:
        If Not swDoc Is Nothing Then
            On Error Resume Next
            swApp.CloseDoc fullPath
            On Error GoTo 0
            Set swDoc = Nothing
            Set swDraw = Nothing
            Set swModel = Nothing
        End If

NextDrawingCO:
    Next localFilePath

    If Not CancelOperation Then
        On Error Resume Next
        If Not ProgressForm Is Nothing Then
            ProgressForm.UpdateProgress fileCount, fileCount, "Batch Complete!"
            DoEvents
        End If
        On Error GoTo 0
        Sleep 300
    End If

    On Error Resume Next
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    On Error GoTo 0

    ShowBatchResults successList, errorList, warningList, fileCount

    Debug.Print "Batch PDF run (checked-out only) completed."
    LogSectionEnd "Processing complete"
    LogBatchSummary fileCount, successList.count, errorList.count, warningList.count
    LogClose

Cleanup:
    On Error Resume Next
    If Not swApp Is Nothing Then
        swApp.Visible = originalVisibility
        swApp.UserControl = True
        Debug.Print "Restored SolidWorks visibility."
    Else
        Debug.Print "swApp object was not available to restore visibility."
    End If
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    On Error GoTo 0

    Set checkedFolders = Nothing
    Set filteredFiles = Nothing
    Debug.Print "Cleanup complete."
End Sub

Sub ShowBatchResults(successList As Collection, errorList As Collection, warningList As Collection, totalFiles As Long)
    Dim frm As New UserForm1
    frm.PopulateResults successList, errorList, warningList, totalFiles
    frm.Show vbModeless
End Sub

Function SelectFolder(Optional promptText As String = "Select a folder", Optional defaultPath As String = "") As String
    Dim selectedPath As String
    selectedPath = InputBox(promptText, "Folder Selection", defaultPath)
    If selectedPath = "" Then
        MsgBox "No folder selected. Operation may be canceled.", vbInformation
        SelectFolder = ""
        Exit Function
    End If

    ' Remove surrounding quotes if user pasted path with quotes
    selectedPath = Trim(selectedPath)
    If Left(selectedPath, 1) = """" Then selectedPath = Mid(selectedPath, 2)
    If Right(selectedPath, 1) = """" Then selectedPath = Left(selectedPath, Len(selectedPath) - 1)

    ' Ensure trailing backslash
    If selectedPath <> "" And Right(selectedPath, 1) <> "\" Then
        selectedPath = selectedPath & "\"
    End If
    SelectFolder = selectedPath
End Function

Function SelectFolder_WithDialog(Optional defaultPath As String = "") As String
    Dim objShell As Object
    Dim objFolder As Object
    Dim selectedPath As String
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Select or Enter Folder Path:", 1, defaultPath)
    If Not objFolder Is Nothing Then
        selectedPath = objFolder.Self.Path
    Else
        selectedPath = ""
    End If
    selectedPath = InputBox("Enter or modify the folder path:", "Folder Selection", selectedPath)
    If selectedPath = "" Then
        MsgBox "No folder selected. Exiting.", vbInformation
        SelectFolder_WithDialog = ""
        Exit Function
    End If

    ' Remove surrounding quotes if user pasted path with quotes
    selectedPath = Trim(selectedPath)
    If Left(selectedPath, 1) = """" Then selectedPath = Mid(selectedPath, 2)
    If Right(selectedPath, 1) = """" Then selectedPath = Left(selectedPath, Len(selectedPath) - 1)

    ' Ensure trailing backslash
    If selectedPath <> "" And Right(selectedPath, 1) <> "\" Then
        selectedPath = selectedPath & "\"
    End If
    SelectFolder_WithDialog = selectedPath
End Function

Public Sub CollectFilesWithSubfolders(folderPath As String, filePattern As String, ByRef filesCollection As Collection, ByRef processedFiles As Object)
    Dim fileName As String
    Dim subFolder As String
    Dim fullFilePath As String
    Dim subFolders As Collection
    
    On Error Resume Next
    ProgressForm.lblCurrentFile.Caption = "Scanning: " & folderPath
    DoEvents
    On Error GoTo 0
    
    Set subFolders = New Collection
    
    subFolder = Dir(folderPath, vbDirectory)
    Do While subFolder <> ""
        If subFolder <> "." And subFolder <> ".." Then
            If StrComp(subFolder, "Obsolete", vbTextCompare) <> 0 And _
               StrComp(subFolder, "History", vbTextCompare) <> 0 Then
                If (GetAttr(folderPath & subFolder) And vbDirectory) = vbDirectory Then
                    On Error Resume Next
                    subFolders.Add folderPath & subFolder & "\", subFolder
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            Else
                Debug.Print "Skipping folder: " & folderPath & subFolder & " (Obsolete or History folder)"
            End If
        End If
        subFolder = Dir()
    Loop
    
    fileName = Dir(folderPath & filePattern, vbNormal)
    Do While fileName <> ""
        fullFilePath = LCase(folderPath & fileName)
        If Not processedFiles.exists(fullFilePath) Then
            filesCollection.Add folderPath & fileName
            processedFiles.Add fullFilePath, fileName
            Debug.Print "Added file: " & folderPath & fileName
        End If
        fileName = Dir()
    Loop
    
    If CancelOperation Then Exit Sub
    
    Dim subfolder_path As Variant
    For Each subfolder_path In subFolders
        Debug.Print "Scanning subfolder: " & subfolder_path
        CollectFilesWithSubfolders CStr(subfolder_path), filePattern, filesCollection, processedFiles
        
        If CancelOperation Then Exit Sub
    Next subfolder_path
End Sub

' ============================================================================
' DOCUMENT MANAGEMENT FUNCTIONS
' ============================================================================
Public Function GetDocumentsReferencingFile(targetFilePath As String, Optional excludePath As String = "") As Collection
    Dim referencingDocs As Collection
    Set referencingDocs = New Collection
    
    Dim uniqueDocs As Object
    Set uniqueDocs = CreateObject("Scripting.Dictionary")
    uniqueDocs.CompareMode = vbTextCompare
    
    Dim docCount As Long
    Dim i As Long
    Dim doc As SldWorks.ModelDoc2
    Dim docPath As String
    Dim docType As Long
    Dim docDeps As Variant
    Dim j As Long
    
    If swApp Is Nothing Then
        Set GetDocumentsReferencingFile = referencingDocs
        Exit Function
    End If
    
    docCount = swApp.GetDocumentCount
    
    If docCount <= 1 Then
        Debug.Print "Only 0-1 documents open, no references possible"
        Set GetDocumentsReferencingFile = referencingDocs
        Exit Function
    End If
    
    Debug.Print "Checking " & docCount & " open documents for references to: " & targetFilePath
    If excludePath <> "" Then Debug.Print "Excluding from results: " & excludePath
    
    For i = 0 To docCount - 1
        On Error Resume Next
        Set doc = swApp.GetDocuments(i)
        On Error GoTo 0
        
        If Not doc Is Nothing Then
            On Error Resume Next
            docPath = doc.GetPathName
            docType = doc.GetType
            On Error GoTo 0
            
            If docPath = "" Then GoTo NextDoc
            
            If Not IsTopLevelDocument(doc) Then
                Debug.Print "Skipping component (not top-level): " & docPath
                GoTo NextDoc
            End If
            
            If StrComp(docPath, targetFilePath, vbTextCompare) <> 0 And _
               StrComp(docPath, excludePath, vbTextCompare) <> 0 Then
                
                On Error Resume Next
                docDeps = doc.GetDependencies2(True, False, False)
                
                If Err.Number = 0 And Not IsEmpty(docDeps) Then
                    For j = LBound(docDeps) To UBound(docDeps)
                        If StrComp(CStr(docDeps(j)), targetFilePath, vbTextCompare) = 0 Then
                            If Not uniqueDocs.exists(docPath) Then
                                referencingDocs.Add docPath
                                uniqueDocs.Add docPath, True
                                Debug.Print "Found referencing document: " & docPath
                            End If
                            Exit For
                        End If
                    Next j
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
        
NextDoc:
    Next i
    
    Debug.Print "Total referencing documents found: " & referencingDocs.count
    Set GetDocumentsReferencingFile = referencingDocs
End Function

Private Function IsTopLevelDocument(doc As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    
    Dim docPath As String
    Dim docTitle As String
    Dim modelView As SldWorks.modelView
    
    docPath = doc.GetPathName
    docTitle = doc.GetTitle
    
    If docPath = "" Or docTitle = "" Then
        IsTopLevelDocument = False
        Exit Function
    End If
    
    Set modelView = doc.GetFirstModelView
    
    If Not modelView Is Nothing Then
        IsTopLevelDocument = True
    Else
        IsTopLevelDocument = False
    End If
    
    On Error GoTo 0
End Function

Public Function CloseReferencingDocuments(referencingDocs As Collection, targetFile As String, Optional batchMode As Boolean = False) As Collection
    Dim closedDocs As Collection
    Set closedDocs = New Collection
    
    Dim docPath As Variant
    Dim docName As String
    Dim userResponse As VbMsgBoxResult
    Dim messageText As String
    Dim i As Long
    
    If referencingDocs.count = 0 Then
        Set CloseReferencingDocuments = closedDocs
        Exit Function
    End If
    
    messageText = "CANNOT UPDATE MODEL REVISION" & vbCrLf & vbCrLf
    messageText = messageText & "The following " & referencingDocs.count & " document(s) are currently using this model:" & vbCrLf & vbCrLf
    
    i = 1
    For Each docPath In referencingDocs
        docName = Mid(CStr(docPath), InStrRev(CStr(docPath), "\") + 1)
        messageText = messageText & i & ". " & docName & vbCrLf
        i = i + 1
    Next docPath
    
    messageText = messageText & vbCrLf & "These documents must be closed to update the model's revision." & vbCrLf & vbCrLf
    messageText = messageText & "WHAT HAPPENS IF YOU CLOSE THEM:" & vbCrLf
    messageText = messageText & "� Model revision will be synced with drawing" & vbCrLf
    messageText = messageText & "� PDF will be created" & vbCrLf
    messageText = messageText & "� Documents will be automatically reopened" & vbCrLf & vbCrLf
    messageText = messageText & "WHAT HAPPENS IF YOU DON'T CLOSE THEM:" & vbCrLf
    messageText = messageText & "� Macro will STOP - no PDF will be created" & vbCrLf
    messageText = messageText & "� You'll need to close these documents manually and run again" & vbCrLf & vbCrLf
    messageText = messageText & "Do you want to close these documents now?"
    
    If Not batchMode Then
        userResponse = MsgBox(messageText, vbQuestion + vbYesNo + vbDefaultButton1, "Close Documents to Update Model?")
        
        If userResponse = vbNo Then
            Debug.Print "User declined to close referencing documents"
            Set CloseReferencingDocuments = Nothing
            Exit Function
        End If
    Else
        Debug.Print "Batch mode: Auto-closing referencing documents"
    End If
    
    For Each docPath In referencingDocs
        On Error Resume Next
        swApp.CloseDoc CStr(docPath)
        If Err.Number = 0 Then
            closedDocs.Add CStr(docPath)
            Debug.Print "Closed document: " & CStr(docPath)
        Else
            Debug.Print "Failed to close document: " & CStr(docPath) & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next docPath
    
    Sleep 300
    
    Set CloseReferencingDocuments = closedDocs
End Function

Public Sub ReopenDocuments(closedDocs As Collection, Optional batchMode As Boolean = False)
    Dim docPath As Variant
    Dim errs As Long
    Dim warns As Long
    Dim doc As SldWorks.ModelDoc2
    Dim docType As Long
    Dim reopenedCount As Long
    Dim failedCount As Long
    
    If closedDocs Is Nothing Or closedDocs.count = 0 Then Exit Sub
    
    Debug.Print "Reopening " & closedDocs.count & " previously closed document(s)"
    reopenedCount = 0
    failedCount = 0
    
    For Each docPath In closedDocs
        If UCase(Right(CStr(docPath), 7)) = ".SLDPRT" Then
            docType = swDocPART
        ElseIf UCase(Right(CStr(docPath), 7)) = ".SLDASM" Then
            docType = swDocASSEMBLY
        ElseIf UCase(Right(CStr(docPath), 7)) = ".SLDDRW" Then
            docType = swDocDRAWING
        Else
            Debug.Print "Unknown document type for: " & CStr(docPath)
            failedCount = failedCount + 1
            GoTo NextDoc
        End If
        
        On Error Resume Next
        Set doc = swApp.OpenDoc6(CStr(docPath), docType, swOpenDocOptions_Silent, "", errs, warns)
        
        If Err.Number = 0 And Not doc Is Nothing Then
            reopenedCount = reopenedCount + 1
            Debug.Print "Reopened document: " & CStr(docPath)
        Else
            failedCount = failedCount + 1
            Debug.Print "Failed to reopen document: " & CStr(docPath) & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        
NextDoc:
    Next docPath
    
    If Not batchMode And (reopenedCount > 0 Or failedCount > 0) Then
        Dim resultMsg As String
        resultMsg = "Reopened " & reopenedCount & " document(s) successfully."
        If failedCount > 0 Then
            resultMsg = resultMsg & vbCrLf & failedCount & " document(s) failed to reopen."
        End If
        MsgBox resultMsg, vbInformation, "Documents Reopened"
    End If
End Sub

' ============================================================================
' CLEANUP AND EXIT FUNCTIONS
' ============================================================================
Public Sub CleanupBasedOnInitialState()
    If Not swWasRunning Then
        Debug.Print "SolidWorks wasn't running initially - will be closed by batch script"
        Exit Sub
    End If
    
    If swWasRunning And Not drawingWasOpen Then
        Debug.Print "SolidWorks was running but no drawing was open initially"
        If Not swModel Is Nothing Then
            If swModel.GetType = swDocDRAWING Then
                Dim currentPath As String
                currentPath = swModel.GetPathName
                If currentPath <> "" And currentPath <> originalDrawingPath Then
                    swApp.CloseDoc currentPath
                    Debug.Print "Closed drawing that was opened by macro: " & currentPath
                End If
            End If
        End If
        Exit Sub
    End If
    
    If drawingWasOpen And originalDrawingPath <> "" Then
        Debug.Print "A drawing was open initially: " & originalDrawingPath
        Dim targetDrawingOpen As Boolean
        targetDrawingOpen = False
        
        If Not swModel Is Nothing Then
            If swModel.GetPathName = originalDrawingPath Then
                targetDrawingOpen = True
            End If
        End If
        
        If Not targetDrawingOpen Then
            Dim docCount As Long
            Dim i As Long
            docCount = swApp.GetDocumentCount
            
            For i = 0 To docCount - 1
                Dim doc As SldWorks.ModelDoc2
                Set doc = swApp.GetDocument(i)
                If Not doc Is Nothing Then
                    If doc.GetPathName = originalDrawingPath Then
                        swApp.ActivateDoc doc.GetPathName
                        targetDrawingOpen = True
                        Debug.Print "Reactivated original drawing: " & originalDrawingPath
                        Exit For
                    End If
                End If
            Next i
        End If
        
        If Not targetDrawingOpen Then
            On Error Resume Next
            Set swModel = swApp.OpenDoc6(originalDrawingPath, swDocDRAWING, swOpenDocOptions_Silent, "", 0, 0)
            If Err.Number = 0 And Not swModel Is Nothing Then
                Debug.Print "Reopened original drawing: " & originalDrawingPath
            Else
                Debug.Print "Failed to reopen original drawing: " & originalDrawingPath
            End If
            Err.Clear
            On Error GoTo 0
        End If
    End If
End Sub

' ============================================================================
' CUSTOM OUTPUT FOLDER PROMPT
' ============================================================================
' Prompts user to optionally use a custom output folder for PDFs
' Returns True if user selected custom folder, False for default behavior
Private Function PromptForCustomOutputFolder() As Boolean
    Dim response As VbMsgBoxResult
    Dim customPath As String

    ' Ask if user wants to use custom output folder
    response = MsgBox("Do you want to save PDFs to a CUSTOM folder?" & vbCrLf & vbCrLf & _
                      "YES - Choose a folder where ALL PDFs will be saved" & vbCrLf & _
                      "NO - Save PDFs to their default locations (job folders)", _
                      vbQuestion + vbYesNo + vbDefaultButton2, _
                      "Custom PDF Output Folder?")

    If response = vbNo Then
        ' Use default behavior
        bManualMode = False
        manualOutputPath = ""
        Debug.Print "User chose default output locations"
        LogMessage "Output mode: Default (job folders)"
        PromptForCustomOutputFolder = False
        Exit Function
    End If

    ' User wants custom folder - prompt for path
    customPath = SelectFolder_WithDialog("Select the folder where ALL PDFs will be saved")

    If customPath = "" Then
        ' User cancelled - fall back to default
        MsgBox "No folder selected. Using default output locations.", vbInformation
        bManualMode = False
        manualOutputPath = ""
        Debug.Print "Custom folder cancelled - using default"
        LogMessage "Output mode: Default (custom folder selection cancelled)"
        PromptForCustomOutputFolder = False
        Exit Function
    End If

    ' Validate the path exists
    On Error Resume Next
    If GetAttr(customPath) = -1 Then
        On Error GoTo 0
        MsgBox "Invalid folder path: " & customPath & vbCrLf & vbCrLf & _
               "Using default output locations instead.", vbExclamation
        bManualMode = False
        manualOutputPath = ""
        PromptForCustomOutputFolder = False
        Exit Function
    End If
    On Error GoTo 0

    ' Ensure trailing backslash
    If Right(customPath, 1) <> "\" Then customPath = customPath & "\"

    ' Set manual mode
    bManualMode = True
    manualOutputPath = customPath

    Debug.Print "Custom output folder selected: " & manualOutputPath
    LogMessage "Output mode: Custom folder - " & manualOutputPath

    MsgBox "PDFs will be saved to:" & vbCrLf & vbCrLf & manualOutputPath, vbInformation, "Custom Output Folder Set"

    PromptForCustomOutputFolder = True
End Function

Public Sub ForceExitForCommandLine()
    Dim runningFromCommandLine As Boolean
    runningFromCommandLine = (Trim(Command) <> "")
    
    If runningFromCommandLine Then
        Debug.Print "ForceExitForCommandLine: Command line execution detected - forcing exit"
        
        On Error Resume Next
        If Not swModel Is Nothing Then
            Dim docPath As String
            docPath = swModel.GetPathName
            If docPath <> "" Then
                swApp.CloseDoc docPath
                Debug.Print "ForceExitForCommandLine: Closed document: " & docPath
            End If
        End If
        On Error GoTo 0
        
        On Error Resume Next
        swApp.ExitApp
        Debug.Print "ForceExitForCommandLine: Called swApp.ExitApp"
        On Error GoTo 0
        
        Set swModel = Nothing
        Set swDraw = Nothing
        Set pdmVault = Nothing
        Set swApp = Nothing
        
        Debug.Print "ForceExitForCommandLine: Calling End to terminate VBA"
        End
    End If
End Sub


