Attribute VB_Name = "Main"
' ====================================================================================
' Main.bas - SolidWorks Property Automation Macro
' ====================================================================================
'
' PURPOSE:
'   Automatically populates and validates custom properties on SolidWorks parts
'   and assemblies for CADLink/PDM integration.
'
' ENTRY POINTS:
'   - main()           : Primary entry point. Processes active document or drawing.
'   - ProcessModels()  : Core loop that iterates through models and subcomponents.
'
' EXECUTION FLOW:
'   1. User runs macro on a part, assembly, or drawing
'   2. If drawing: extracts referenced model(s) to process
'   3. If face selected: delegates to DXF_SelectedFace for export
'   4. For each model:
'      a. Check PDM checkout status (read-only if not checked out)
'      b. Collect existing properties into propertiesToSet dictionary
'      c. Calculate physics/geometry (dimensions, weight, material)
'      d. Apply business rules via AddPropertiesBasedOnRules (sorting only)
'      e. Validate properties against regex rules
'      f. Single write via ApplyCollectedProperties (Main.bas)
'      g. Save model if checked out
'   5. Display validation results in UnifiedResultsForm
'
' KEY MODULES:
'   - AddProperties.bas     : Sorts properties per business rules (NO file writes)
'   - PropertyRules.bas     : Defines required properties and validation regex
'   - DimensionsModule.bas  : Calculates bounding box dimensions
'   - WeightModule.bas      : Calculates part weight
'   - MaterialModule.bas    : Extracts material information
'   - Locations.bas         : Validates file location paths
'
' IMPORTANT:
'   Property writes happen ONLY in ApplyCollectedProperties (this file).
'   All other modules should only modify the propertiesToSet dictionary.
'
' CONFIGURATION:
'   PROCESS_PUR_PARTS - Set to True to process PUR (Purchase) parts normally (default).
'                       Set to False to skip PUR parts entirely.
'
' ====================================================================================

Option Explicit

' Global variables

    Global Const MacroVersion As Long = 100 ' Update this version number manually on each check-in

    ' ====================================================================================
    ' CONFIGURATION: PUR Part Processing
    ' ====================================================================================
    ' Set to True to process PUR (Purchase) parts normally
    ' Set to False to skip processing PUR parts entirely
    Global Const PROCESS_PUR_PARTS As Boolean = True
    ' ====================================================================================

    Global swApp As SldWorks.SldWorks
    Global vault As EdmVault5

    ' Global variable to hold the log file path
    ' LOG_FILE_PATH moved to ConfigConstants.bas as DEBUG_LOG_PATH




' Updated logging constants

' SUMMARY_LOG_PATH moved to ConfigConstants.bas

' Global counter for models processed
Private modelProcessCounter As Long



' PDM Lock Status Constants
Const EdmLockStatus_NotLocked = 0
Const EdmLockStatus_LockedByUser = 1
Const EdmLockStatus_LockedByOther = 2



Global checkoutUserInfo As Scripting.Dictionary

' ====================================================================================
' OPTIMIZATION: Global SolidWorks App Accessor
' ====================================================================================
' Use this function instead of creating local swApp variables in each module.
' This eliminates redundant Application.SldWorks calls across the codebase.
' ====================================================================================
Public Function GetSwApp() As SldWorks.SldWorks
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
    End If
    Set GetSwApp = swApp
End Function

' Initialize both log files
Private Sub InitLogFile()
    ' Initialize debug log via Logger
    Logger.StartSession "Model Processing Session"

    Dim fileNum As Integer
    
    ' Initialize summary log (new)
    fileNum = FreeFile
    On Error Resume Next
    ' Use Append instead of Output to keep history for the day (since filename has date)
    Open SUMMARY_LOG_PATH For Append As #fileNum
    Print #fileNum, ""
    Print #fileNum, "=================================================="
    Print #fileNum, "    MODEL PROCESSING SUMMARY REPORT"
    Print #fileNum, "    Started: " & Now()
    Print #fileNum, "    User: " & Environ("USERNAME")
    Print #fileNum, "=================================================="
    Print #fileNum, ""
    Close #fileNum
    On Error GoTo 0
    
    modelProcessCounter = 0
End Sub

    
 
 
 
 Private Sub LogModelResult(modelName As String, modelPath As String, status As String, details As String)
    modelProcessCounter = modelProcessCounter + 1
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open SUMMARY_LOG_PATH For Append As #fileNum
    Print #fileNum, "Model #" & modelProcessCounter & ": " & modelName
    Print #fileNum, "  Path: " & modelPath
    Print #fileNum, "  Status: " & status
    Print #fileNum, "  Details: " & details
    Print #fileNum, "  Time: " & Format(Now(), "hh:mm:ss")
    Print #fileNum, ""
    Close #fileNum
    On Error GoTo 0
    
    ' Also log to debug file
    LogToFile "=== MODEL RESULT #" & modelProcessCounter & " ==="
    LogToFile "Model: " & modelName & " | Status: " & status & " | Details: " & details
End Sub


Private Sub LogFinalSummary()
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open SUMMARY_LOG_PATH For Append As #fileNum
    Print #fileNum, "=================================================="
    Print #fileNum, "    PROCESSING COMPLETE"
    Print #fileNum, "    Total Models Processed: " & modelProcessCounter
    Print #fileNum, "    Finished: " & Now()
    Print #fileNum, "=================================================="
    Close #fileNum
    On Error GoTo 0
    
    ' Finish logger session and sync to shared folder
    Logger.EndSession modelProcessCounter
    
    ' Show completion message
  '  MsgBox "Processing complete! " & modelProcessCounter & " models processed." & vbCrLf & vbCrLf & _
   '        "Summary saved to:" & vbCrLf & SUMMARY_LOG_PATH, vbInformation
End Sub
 
 Sub main()
    Dim activeDoc As ModelDoc2
    Dim modelToProcess As ModelDoc2
    Dim issuesDetected As Boolean
    Dim issuesReport As String
    Dim aggregatedFailedProperties As Scripting.Dictionary
    Dim aggregatedFailedLocations As Scripting.Dictionary
    Dim aggregatedFailedPieceParts As Scripting.Dictionary
    Dim startTime As Double
    Dim endTime As Double
    Dim totalStartTime As Double
    Dim totalEndTime As Double
    Dim processSubcomponents As Boolean
    Dim isDrawing As Boolean
    
    On Error GoTo ErrorHandler

    ' Initialize log file
    InitLogFile
    LogToFile "=== Starting Main procedure ==="

    ' Track version usage
    VersionTracker.LogVersionUsage

    ' Reset flat bar skip flag for this run
    FlatBarModule.ResetFlatBarSkipAll

    ' Initialize the global SolidWorks application object
    Set swApp = Application.SldWorks

    ' Initialize the checkout user info dictionary
    If checkoutUserInfo Is Nothing Then
        Set checkoutUserInfo = New Scripting.Dictionary
    End If

    ' Initialize the PDM vault object
    If vault Is Nothing Then
        Set vault = New EdmVault5
        If Not vault.IsLoggedIn Then
            vault.LoginAuto "NMT_PDM", 0
        End If
    End If

    ' Check for and get the latest version of the macro
    Dim updateResult As Boolean
    updateResult = GetLatestMacroVersion()

    If Not updateResult Then
        ' Exit immediately if the user chose "No" or if any issue occurred
        Exit Sub
    End If

    totalStartTime = Timer

    Set activeDoc = swApp.activeDoc

    If activeDoc Is Nothing Then
        MsgBox "No active document. Please open a model or drawing to run the macro."
        Exit Sub
    End If

    ' ============================
    ' FACE SELECTION ROUTING GATE
    ' If exactly one face is selected, run the DXF exporter and exit.
    ' ============================
    On Error Resume Next
    Dim selMgr As SelectionMgr
    Set selMgr = activeDoc.SelectionManager
    If Not selMgr Is Nothing Then
        Dim selCount As Long
        selCount = selMgr.GetSelectedObjectCount2(-1)
        If selCount = 1 Then
            Dim selType As Long
            selType = selMgr.GetSelectedObjectType3(1, -1)
            If selType = 2 Then ' 2 = swSelFACES
                LogToFile "Detected single face selection. Delegating to DXF_SelectedFace.main"
                On Error GoTo 0
                DXF_SelectedFace.main
                Exit Sub
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    ' ============================
    ' END FACE SELECTION ROUTING
    ' ============================

    ' ============================
    ' SMART DOCUMENT PROCESSING LOGIC
    ' Handle drawings vs models appropriately - only require save for what we modify
    ' ============================
    Dim actualModelToProcess As ModelDoc2
    Dim originalDocument As ModelDoc2
    Dim wasOriginallyDrawing As Boolean
    
    Set originalDocument = activeDoc
    wasOriginallyDrawing = (activeDoc.GetType = swDocDRAWING)
    
    If wasOriginallyDrawing Then
        LogToFile "=== PROCESSING STARTED FROM DRAWING ==="
        LogToFile "Drawing: " & activeDoc.GetTitle
        
        ' For drawings, we don't need them to be saved since we're not modifying them
        Dim drawingPath As String
        drawingPath = activeDoc.GetPathName()
        If drawingPath = "" Then
            LogToFile "Drawing is not saved, but that's OK - we're not modifying it"
            LogToFile "Drawing name: " & activeDoc.GetTitle
        Else
            LogToFile "Drawing path: " & drawingPath
        End If
        
        ' Get the referenced model from the drawing
        Set actualModelToProcess = GetReferencedModelFromDrawing(activeDoc)
        
        If actualModelToProcess Is Nothing Then
            MsgBox "Could not find a referenced model in the drawing '" & activeDoc.GetTitle & "'." & vbCrLf & _
                   "Please ensure the drawing has at least one view that references a part or assembly.", _
                   vbExclamation, "No Referenced Model Found"
            Exit Sub
        End If
        
        ' IMPORTANT: The referenced MODEL must be saved since we're modifying it
        Dim refModelPath As String
        refModelPath = actualModelToProcess.GetPathName()
        If refModelPath = "" Then
            MsgBox "The referenced model '" & actualModelToProcess.GetTitle & "' is not saved." & vbCrLf & vbCrLf & _
                   "The macro needs to modify the model's properties, so the model must be saved." & vbCrLf & _
                   "Please save the referenced model before running this macro.", _
                   vbExclamation, "Referenced Model Not Saved"
            Exit Sub
        End If
        
        LogToFile "Found referenced model: " & actualModelToProcess.GetTitle
        LogToFile "Referenced model path: " & refModelPath
        
        ' ============================
        ' NEW APPROACH: DO NOT ACTIVATE THE MODEL AT ALL
        ' We'll work with the model reference we already have
        ' ============================
        LogToFile "=== WORKING WITH REFERENCED MODEL WITHOUT ACTIVATION ==="
        LogToFile "Current active document remains: " & swApp.activeDoc.GetTitle
        LogToFile "Referenced model to process: " & actualModelToProcess.GetTitle
        
        ' We already have the model reference from GetReferencedModelFromDrawing
        ' No need to activate it - just use it directly
        LogToFile "Model reference obtained without activation"
        LogToFile "=== READY TO PROCESS MODEL ==="
        ' ============================
        ' END NEW APPROACH
        ' ============================
        
    Else
        LogToFile "=== PROCESSING STARTED FROM MODEL ==="
        LogToFile "Model: " & activeDoc.GetTitle
        
        ' For models, they must be saved since we're modifying them
        Dim modelPath As String
        modelPath = activeDoc.GetPathName()
        If modelPath = "" Then
            MsgBox "Please save the model '" & activeDoc.GetTitle & "' before running this macro." & vbCrLf & vbCrLf & _
                   "The macro modifies the model's properties, so it must be saved first.", _
                   vbExclamation, "Model Not Saved"
            Exit Sub
        End If
        
        LogToFile "Model path: " & modelPath
        Set actualModelToProcess = activeDoc
    End If
    
    ' Set the model to process
    Set modelToProcess = actualModelToProcess
    LogToFile "Final model to process: " & modelToProcess.GetTitle & " (" & modelToProcess.GetPathName() & ")"
    ' ============================
    ' END SMART DOCUMENT PROCESSING
    ' ============================

    ' Initialize the dictionaries for tracking issues
    Set aggregatedFailedProperties = New Scripting.Dictionary
    Set aggregatedFailedLocations = New Scripting.Dictionary
    Set aggregatedFailedPieceParts = New Scripting.Dictionary
    issuesDetected = False
    issuesReport = ""

    ' Check PDM checkout status for the model
    modelPath = modelToProcess.GetPathName
    Dim modelName As String
    modelName = modelToProcess.GetTitle
    
    LogToFile "Checking PDM checkout status for model: " & modelName & " (" & modelPath & ")"
    
    If Not IsModelCheckedOut(modelPath) Then
        LogToFile "Model is not checked out by us - will skip but not treat as an error"

        ' Check if we have username info - just for logging
        ' Safety check: ensure dictionary is initialized before accessing
        If Not checkoutUserInfo Is Nothing Then
            If checkoutUserInfo.exists(modelPath) Then
                Dim checkoutUser As String
                checkoutUser = checkoutUserInfo(modelPath)
                LogToFile "Model checkout info: MODEL CHECKED OUT BY: " & checkoutUser
            Else
                LogToFile "Model checkout info: MODEL NOT CHECKED OUT"
            End If
        Else
            LogToFile "Model checkout info: Dictionary not initialized"
        End If

        LogToFile "Model will be skipped for editing but not reported as an error"
    Else
        LogToFile "Model is checked out by us - continuing"
    End If

    ' Prompt user if the model is an assembly
    If modelToProcess.GetType = swDocASSEMBLY Then
        Dim userChoice As VbMsgBoxResult
        Dim promptText As String
        
        If wasOriginallyDrawing Then
            promptText = "The drawing references an assembly. Do you want to process all parts in the assembly (including subcomponents and references)?" & vbCrLf & _
                        "Click Yes to process all parts, or No to process only the assembly itself."
        Else
            promptText = "Do you want to process all parts in the assembly (including subcomponents and references)?" & vbCrLf & _
                        "Click Yes to process all parts, or No to process only the assembly itself."
        End If
        
        userChoice = MsgBox(promptText, vbQuestion + vbYesNo, "Process Subcomponents?")
        If userChoice = vbYes Then
            processSubcomponents = True
        Else
            processSubcomponents = False
        End If
    Else
        processSubcomponents = True
    End If

    ' Resolve all lightweight components of the active model
    startTime = Timer
    Call ResolveLightweightComponents(modelToProcess)
    endTime = Timer

    ' Process models with the chosen option for processing subcomponents
    startTime = Timer
    Call ProcessModels(modelToProcess, issuesReport, aggregatedFailedProperties, aggregatedFailedLocations, aggregatedFailedPieceParts, issuesDetected, processSubcomponents)
    endTime = Timer

    ' ============================
    ' ENSURE DRAWING IS ACTIVE (if we started from drawing)
    ' Since we never switched away, this should be minimal
    ' ============================
    If wasOriginallyDrawing Then
        LogToFile "=== ENSURING DRAWING IS ACTIVE ==="
        
        ' The drawing should still be active since we didn't activate the model
        If Not (swApp.activeDoc Is originalDocument) Then
            LogToFile "WARNING: Active document changed unexpectedly"
            LogToFile "Current active: " & swApp.activeDoc.GetTitle
            LogToFile "Expected: " & originalDocument.GetTitle
            
            ' Force activate the drawing
            On Error Resume Next
            swApp.ActivateDoc2 originalDocument.GetTitle, True, 0
            If Err.Number <> 0 Then
                LogToFile "ERROR restoring drawing: " & Err.description
                Err.Clear
            Else
                LogToFile "Drawing restored as active document"
            End If
            On Error GoTo ErrorHandler
        Else
            LogToFile "Drawing is still active (as expected)"
        End If
        
        LogToFile "=== DRAWING CHECK COMPLETE ==="
    End If
    ' ============================
    ' END ENSURE DRAWING IS ACTIVE
    ' ============================

    ' ============================
    ' POST-PROCESSING FOR DRAWINGS
    ' If we started from a drawing, rebuild it to reflect any changes
    ' ============================
    If wasOriginallyDrawing Then
        LogToFile "=== POST-PROCESSING: REBUILDING ORIGINAL DRAWING ==="
        LogToFile "Rebuilding drawing: " & originalDocument.GetTitle
        
        ' Force rebuild of the drawing to update any changes from the referenced model
        On Error Resume Next
        originalDocument.ForceRebuild3 True
        If Err.Number <> 0 Then
            LogToFile "WARNING: Error during drawing rebuild: " & Err.description
            Err.Clear
        Else
            LogToFile "Drawing rebuilt successfully"
        End If
        
        ' Also refresh the drawing views
        Call RebuildAllDrawingViews(originalDocument)
        
        On Error GoTo ErrorHandler
        LogToFile "=== DRAWING POST-PROCESSING COMPLETE ==="
    End If
    ' ============================
    ' END DRAWING POST-PROCESSING
    ' ============================

    totalEndTime = Timer

    ' Log final counts before showing form
    LogToFile "=== FINAL COUNTS BEFORE SHOWING FORM ==="
    LogToFile "issuesDetected = " & issuesDetected
    LogToFile "aggregatedFailedProperties.Count = " & aggregatedFailedProperties.Count
    LogToFile "aggregatedFailedLocations.Count = " & aggregatedFailedLocations.Count
    LogToFile "aggregatedFailedPieceParts.Count = " & aggregatedFailedPieceParts.Count
    
    ' Log location failures
    If aggregatedFailedLocations.Count > 0 Then
        LogToFile "Location failures:"
        Dim locKey As Variant
        For Each locKey In aggregatedFailedLocations.keys
            LogToFile "  - Path: " & locKey
            LogToFile "    Message: " & aggregatedFailedLocations(locKey)
        Next locKey
    End If

    ' Display results using the UnifiedResultsForm if issues are detected
    If issuesDetected Then
        LogToFile "Calling ShowUnifiedResultsForm with " & aggregatedFailedLocations.Count & " location failures"

        ' Calculate unique problematic models
        Dim uniqueModels As Scripting.Dictionary
        Set uniqueModels = New Scripting.Dictionary
        Dim key As Variant
      '  Dim modelName As String

        For Each key In aggregatedFailedProperties.keys
            Dim pathParts As Variant
            pathParts = Split(CStr(key), " - ")
            If UBound(pathParts) >= 1 Then
                modelName = ExtractFilenameFromPath(CStr(pathParts(0)))
            Else
                modelName = CStr(key)
            End If
            If InStr(modelName, ".") > 0 Then
                modelName = Left(modelName, InStrRev(modelName, ".") - 1)
            End If
            uniqueModels(modelName) = True
        Next key

        For Each key In aggregatedFailedLocations.keys
            modelName = ExtractFilenameFromPath(CStr(key))
            If InStr(modelName, ".") > 0 Then
                modelName = Left(modelName, InStrRev(modelName, ".") - 1)
            End If
            uniqueModels(modelName) = True
        Next key

        For Each key In aggregatedFailedPieceParts.keys
            If InStr(CStr(key), "\") > 0 Then
                modelName = ExtractFilenameFromPath(CStr(key))
            Else
                modelName = CStr(key)
            End If
            If InStr(modelName, ".") > 0 Then
                modelName = Left(modelName, InStrRev(modelName, ".") - 1)
            End If
            uniqueModels(modelName) = True
        Next key

        ' Use the dynamic form builder
        Call UnifiedResultsFormBuilder.ShowUnifiedResultsForm(aggregatedFailedProperties, aggregatedFailedLocations, aggregatedFailedPieceParts, uniqueModels.Count)
Else
        LogToFile "No issues detected, showing success message"
        Dim successMessage As String
If wasOriginallyDrawing Then
    successMessage = "All models have been processed. Properties have been checked and updated where necessary."
Else
    successMessage = "All models have been processed. Properties have been checked and updated where necessary."
End If
MsgBox successMessage
    End If

    ' ============================
    ' CLEANUP ALL OBJECT REFERENCES
    ' ============================
    LogToFile "=== CLEANING UP OBJECT REFERENCES ==="
    Set actualModelToProcess = Nothing
    Set modelToProcess = Nothing
    Set originalDocument = Nothing
    Set swApp = Nothing
    LogToFile "=== CLEANUP COMPLETE ==="
    ' ============================
LogFinalSummary
    Exit Sub

ErrorHandler:
    ' Cleanup references even if error occurs
    On Error Resume Next
    Set actualModelToProcess = Nothing
    Set modelToProcess = Nothing
    Set originalDocument = Nothing
    Set swApp = Nothing
    On Error GoTo 0
    
    LogToFile "ERROR in Main: " & Err.description & " (Line: " & Erl & ")"
    MsgBox "An error occurred: " & Err.description, vbCritical
    
     LogFinalSummary
    Exit Sub
End Sub




Private Function IsModelCheckedOut(modelPath As String) As Boolean
    LogToFile "======================================================"
    LogToFile "=== DEBUGGING IsModelCheckedOut for: " & modelPath
    LogToFile "======================================================"
    
    ' Initialize the Scripting.Dictionary if it doesn't exist
    If checkoutUserInfo Is Nothing Then
        Set checkoutUserInfo = New Scripting.Dictionary
        LogToFile "Created new checkoutUserInfo Scripting.Dictionary"
    End If
    
    ' Get Windows username early for comparisons and debugging
    Dim windowsUser As String
    windowsUser = Environ("USERNAME")
    LogToFile "Current Windows User: " & windowsUser
    
    On Error GoTo ErrorHandler
    
    ' Ensure vault is initialized and logged in
    If vault Is Nothing Then
        LogToFile "Vault object is Nothing, initializing..."
        Set vault = New EdmVault5
        LogToFile "New vault object created: " & (Not vault Is Nothing)
    End If
    
    ' Check login status and log in if needed
    If Not vault.IsLoggedIn Then
        LogToFile "Vault not logged in. Attempting login..."
        Dim vaultName As String
        
        ' Get vault name from path
        On Error Resume Next
        vaultName = vault.GetVaultNameFromPath(modelPath)
        LogToFile "GetVaultNameFromPath result: " & vaultName & " (Error: " & Err.Number & ")"
        If Err.Number <> 0 Or vaultName = "" Then
            vaultName = "NMT_PDM" ' Fallback
            LogToFile "Using fallback vault name: " & vaultName
        End If
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Attempt login
        LogToFile "Attempting LoginAuto with vault: " & vaultName
        On Error Resume Next
        vault.LoginAuto vaultName, 0
        If Err.Number <> 0 Then
            LogToFile "LoginAuto ERROR: " & Err.description & " (ErrNo: " & Err.Number & ")"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        
        If vault.IsLoggedIn Then
            LogToFile "LoginAuto successful. Vault: " & vault.Name
        Else
            LogToFile "LoginAuto FAILED. Will assume file is editable to avoid blocking."
            MsgBox "Failed to log in to PDM. Please check your PDM connection and credentials.", vbCritical
            IsModelCheckedOut = True ' Fallback to allow editing
            Exit Function
        End If
    Else
        LogToFile "Vault already logged in. Vault: " & vault.Name
    End If
    
    ' Get the file from PDM
    Dim pdmFolder As IEdmFolder5
    Dim file As IEdmFile5
    LogToFile "Attempting to get file from path: " & modelPath
    On Error Resume Next
    Set file = vault.GetFileFromPath(modelPath, pdmFolder)
    If Err.Number <> 0 Then
        LogToFile "GetFileFromPath ERROR: " & Err.description & " (ErrNo: " & Err.Number & ")"
        Err.Clear
        IsModelCheckedOut = True ' Not in vault (or PDM error) = Assume OK to edit locally
        LogToFile "RESULT: File not in vault or error. Returning IsModelCheckedOut = TRUE"
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    If file Is Nothing Then
        LogToFile "File is Nothing (not found in vault)"
        IsModelCheckedOut = True
        LogToFile "RESULT: File not in vault. Returning IsModelCheckedOut = TRUE"
        Exit Function
    End If
    
    LogToFile "File found in vault: " & file.Name & " (ID: " & file.ID & ")"
    
    ' Use the direct LockedByUser method to determine checkout status
    Dim pdmUser As IEdmUser5
    LogToFile "Using direct file.LockedByUser method..."
    
    On Error Resume Next
    Set pdmUser = file.lockedByUser
    If Err.Number <> 0 Then
        LogToFile "LockedByUser ERROR: " & Err.description & " (ErrNo: " & Err.Number & ")"
        pdmUser = Nothing
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Default values
    Dim checkedOut As Boolean
    Dim checkedOutBy As String
    Dim weHaveIt As Boolean
    
    ' Determine checkout status based on pdmUser result
    If Not pdmUser Is Nothing Then
        LogToFile "File is checked out by user: " & pdmUser.Name
        checkedOut = True
        checkedOutBy = pdmUser.Name
        
        ' Compare usernames to determine if we have it checked out
        LogToFile "Comparing user names: '" & checkedOutBy & "' vs '" & windowsUser & "'"
        
        If LCase(checkedOutBy) = LCase(windowsUser) Then
            LogToFile "MATCH: Exact name match (case-insensitive)"
            weHaveIt = True
        ElseIf InStr(1, LCase(checkedOutBy), LCase(windowsUser), vbTextCompare) > 0 Then
            LogToFile "MATCH: Username contained in checkout name"
            weHaveIt = True
        ElseIf InStr(1, LCase(windowsUser), LCase(checkedOutBy), vbTextCompare) > 0 Then
            LogToFile "MATCH: Checkout name contained in username"
            weHaveIt = True
        Else
            LogToFile "NO MATCH: Different users"
            weHaveIt = False
        End If
    Else
        ' No user has it checked out
        LogToFile "File is not checked out."
        checkedOut = False
        checkedOutBy = "Not Checked Out"
        weHaveIt = False
    End If
    
    ' Store the checkout info and return result
    LogToFile "FINAL DECISION:"
    LogToFile "Checked Out: " & checkedOut
    LogToFile "Checked Out By: " & checkedOutBy
    LogToFile "We Have It: " & weHaveIt
    
    ' Save user info to Scripting.Dictionary for reporting
    checkoutUserInfo(modelPath) = checkedOutBy
    
    ' Set the function result
    IsModelCheckedOut = weHaveIt
    
    LogToFile "RETURNING IsModelCheckedOut = " & IsModelCheckedOut
    LogToFile "===== END IsModelCheckedOut DEBUG ====="
    Exit Function
    
ErrorHandler:
    LogToFile "CRITICAL ERROR in IsModelCheckedOut: " & Err.description & " (Line: " & Erl & ", ErrNo: " & Err.Number & ")"
    
    ' Show error in debug message
    MsgBox "CRITICAL ERROR in PDM Check:" & vbCrLf & _
           "Error: " & Err.description & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "File Path: " & modelPath & vbCrLf & vbCrLf & _
           "Will assume file is editable to avoid blocking.", _
           vbExclamation, "PDM Error Debug"
           
    IsModelCheckedOut = True ' On error, assume OK to prevent blocking user
    LogToFile "ERROR STATE: Returning IsModelCheckedOut = TRUE (fallback)"
    Err.Clear
End Function







    
 Private Function GetLatestMacroVersion() As Boolean
    On Error GoTo ErrorHandler
    
    DebugLog "--- Starting Enhanced GetLatestMacroVersion ---"
    
    Dim macroPath As String
    macroPath = PDM_MACRO_PATH
    
    DebugLog "Checking macro at path: " & macroPath
    
    If Dir(macroPath) = "" Then
        DebugLog "ERROR: Macro file not found at " & macroPath
        MsgBox "Error: The macro file could not be found at " & macroPath, vbCritical
        GetLatestMacroVersion = False
        Exit Function
    End If
    
    DebugLog "Macro file found. Attempting to get file object from vault."
    
    ' Ensure the vault is initialized and logged in
    If vault Is Nothing Then
        DebugLog "Vault object is not initialized. Creating a new one."
        Set vault = New EdmVault5
        If Not vault.IsLoggedIn Then
            vault.LoginAuto "NMT_PDM", 0
        End If
    End If

    DebugLog "Vault initialized and logged in."

    ' Get the folder object from the vault
    Dim folder As IEdmFolder5
    Set folder = vault.GetFolderFromPath(Left(macroPath, InStrRev(macroPath, "\")))

    If folder Is Nothing Then
        DebugLog "ERROR: Folder object is Nothing after GetFolderFromPath"
        MsgBox "Error: Could not find the macro folder in PDM vault", vbCritical
        GetLatestMacroVersion = False
        Exit Function
    End If
    
    DebugLog "Folder object retrieved successfully. Folder ID: " & folder.ID

    ' Get the file object from the vault
    Dim file As IEdmFile5
    Set file = vault.GetFileFromPath(macroPath)
    
    If file Is Nothing Then
        DebugLog "ERROR: File object is Nothing after GetFileFromPath"
        MsgBox "Error: Could not find the macro file in PDM vault", vbCritical
        GetLatestMacroVersion = False
        Exit Function
    End If
    
    DebugLog "File object retrieved successfully. File ID: " & file.ID
    
    ' Retrieve version numbers
    Dim latestVersion As Long
    latestVersion = file.CurrentVersion
    
    DebugLog "Current Macro Version (from global var): " & MacroVersion
    DebugLog "Latest Version in PDM: " & latestVersion
    
    ' MODIFIED: Don't allow users to continue with outdated versions
    If MacroVersion < latestVersion Then
        DebugLog "WARNING: Using outdated macro version"
        
        ' Give user clear instructions to get the latest version
       MsgBox "This macro requires updating." & vbCrLf & vbCrLf & _
       "Your version: v" & MacroVersion & vbCrLf & _
       "Latest version: v" & latestVersion & vbCrLf & vbCrLf & _
       "To update, please access the latest version from:" & vbCrLf & _
       PDM_LIBRARIES_PATH & "Macro\" & vbCrLf & vbCrLf & _
       "The current operation has been cancelled.", _
       vbInformation, "Macro Update Required"
        
        DebugLog "Execution stopped due to outdated version"
        GetLatestMacroVersion = False
    Else
        DebugLog "Using current version"
        GetLatestMacroVersion = True
    End If
    
    DebugLog "--- Ending Enhanced GetLatestMacroVersion ---"
    Exit Function
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Error in GetLatestMacroVersion: " & Err.description & " (Error number: " & Err.Number & ")"
    DebugLog "ERROR: " & errorMsg
    MsgBox errorMsg, vbCritical
    GetLatestMacroVersion = False
End Function
    


Private Sub ProcessModels(model As ModelDoc2, _
                  ByRef issuesReport As String, _
                  ByRef aggregatedFailedProperties As Scripting.Dictionary, _
                  ByRef aggregatedFailedLocations As Scripting.Dictionary, _
                  ByRef aggregatedFailedPieceParts As Scripting.Dictionary, _
                  ByRef issuesDetected As Boolean, _
                  processSubcomponents As Boolean)

    On Error GoTo ErrorHandler

    LogToFile "=== Starting ProcessModels ==="
    LogToFile "Initial model: " & model.GetTitle
    LogToFile "Initial aggregatedFailedProperties count: " & aggregatedFailedProperties.Count
    LogToFile "Initial aggregatedFailedLocations count: " & aggregatedFailedLocations.Count
    LogToFile "Process subcomponents flag: " & processSubcomponents

    ' Initialize property dictionaries
    Dim propertiesToSet As Scripting.Dictionary
    Set propertiesToSet = New Scripting.Dictionary
    Dim initialPropertiesToSet As Scripting.Dictionary
    Dim skipForPUR As Boolean
    Dim canSave As Boolean
    Dim modelType As Long
    Dim shouldClose As Boolean

    ' Initialize main dictionaries
    Dim processedModels As Scripting.Dictionary
    Set processedModels = New Scripting.Dictionary
    Dim modelsToProcess As Scripting.Dictionary
    Set modelsToProcess = New Scripting.Dictionary

    ' Track which models were already open before we started (so we don't close them)
    Dim modelsAlreadyOpen As Scripting.Dictionary
    Set modelsAlreadyOpen = New Scripting.Dictionary

    ' Capture all currently open documents
    LogToFile "=== CAPTURING INITIALLY OPEN DOCUMENTS ==="
    Dim openDoc As ModelDoc2
    On Error Resume Next
    Set openDoc = swApp.GetFirstDocument
    If Err.Number <> 0 Then
        LogToFile "Error getting first document: " & Err.description
        Err.Clear
        Set openDoc = Nothing
    End If
    On Error GoTo 0

    Do While Not openDoc Is Nothing
        Dim openPath As String
        openPath = openDoc.GetPathName()
        If openPath <> "" Then
            If Not modelsAlreadyOpen.exists(openPath) Then
                modelsAlreadyOpen.Add openPath, True
                LogToFile "Already open: " & openPath
            End If
        End If
        Set openDoc = openDoc.GetNext
    Loop
    On Error GoTo ErrorHandler
    LogToFile "Total documents already open: " & modelsAlreadyOpen.Count

    If model Is Nothing Then
        LogToFile "ERROR: Initial model is Nothing"
        Exit Sub
    End If

    ' Add initial model to the processing queue
    Dim initialPath As String
    initialPath = model.GetPathName()
    
    If initialPath <> "" Then
        modelsToProcess.Add initialPath, model
        LogToFile "Added initial model to queue: " & initialPath
    Else
        LogToFile "UNEXPECTED ERROR: Model path is empty after validation in main()"
        MsgBox "Unexpected error: Model path validation failed. Please contact support.", vbCritical
        Exit Sub
    End If

    ' Add collection for save errors
    Dim saveErrorsList As collection
    Set saveErrorsList = New collection

    ' DEBUG: Log initial queue state
    LogToFile "=== MAIN PROCESSING LOOP START ==="
    LogToFile "Models in queue: " & modelsToProcess.Count
    Dim debugKey As Variant
    For Each debugKey In modelsToProcess.keys
        LogToFile "Queue item: " & debugKey
    Next debugKey

    Do While modelsToProcess.Count > 0
        Dim currentModel As ModelDoc2
        Dim modelKey As Variant
        modelKey = modelsToProcess.keys()(0) ' Get the key (path)
        Set currentModel = modelsToProcess.items()(0) ' Get the model object
        Dim modelPath As String
        modelPath = CStr(modelKey)

        If currentModel Is Nothing Then
            LogToFile "WARNING: Null model object found in queue for key '" & modelKey & "', removing..."
            modelsToProcess.Remove modelKey
            GoTo ContinueLoop
        End If

        If modelPath = "" Then
             LogToFile "WARNING: Model path is empty in queue for '" & currentModel.GetTitle & "', removing..."
             modelsToProcess.Remove modelKey
             Set currentModel = Nothing
             GoTo ContinueLoop
        End If

        LogToFile vbNewLine & "=== Processing queue item ==="
        LogToFile "Model: " & currentModel.GetTitle & " (" & modelPath & ")"

        ' Check if already processed
        If processedModels.exists(modelPath) Then
             LogToFile "Model already processed, skipping."
             modelsToProcess.Remove modelPath
             Set currentModel = Nothing
             GoTo ContinueLoop
        End If

        ' --- Handle Drawings by Getting Referenced Models ---
        Dim originalDrawing As ModelDoc2
        Dim isOriginallyDrawing As Boolean
        Set originalDrawing = Nothing
        isOriginallyDrawing = False
        
        If currentModel.GetType = swDocDRAWING Then
            LogToFile "WARNING: Drawing found in processing queue - this should have been handled in main()"
            LogToFile "Drawing detected, getting referenced model..."
            isOriginallyDrawing = True
            Set originalDrawing = currentModel
            
            Dim referencedModel As ModelDoc2
            Set referencedModel = GetReferencedModelFromDrawing(currentModel)
            
            If Not referencedModel Is Nothing Then
                LogToFile "Found referenced model: " & referencedModel.GetTitle
                Set currentModel = referencedModel
                modelPath = referencedModel.GetPathName()
                LogToFile "Switched to referenced model: " & modelPath
                
                If modelPath = "" Then
                    LogToFile "ERROR: Referenced model has no path, skipping"
                    LogModelResult currentModel.GetTitle, "No Path", "ERROR", "Referenced model has no valid path"
                    GoTo RemoveAndContinue
                End If
            Else
                LogToFile "No referenced model found in drawing, skipping"
                LogModelResult "Unknown Drawing", modelPath, "ERROR", "No referenced model found in drawing"
                GoTo RemoveAndContinue
            End If
        End If
        
        ' ========================================================================
        ' CRITICAL FIX: ADD SUBCOMPONENTS FIRST (BEFORE CHECKING SKIP CONDITIONS)
        ' This ensures that even if an assembly is skipped, its parts are still processed
        ' ========================================================================
        Dim currentModelTypeForSubCheck As Long
        On Error Resume Next
        currentModelTypeForSubCheck = currentModel.GetType
        If Err.Number <> 0 Then
            currentModelTypeForSubCheck = 0
            Err.Clear
        End If
        On Error GoTo ErrorHandler

        If processSubcomponents And currentModelTypeForSubCheck = swDocASSEMBLY Then
            LogToFile "=== ADDING ASSEMBLY SUBCOMPONENTS TO QUEUE ==="
            LogToFile "Assembly: " & currentModel.GetTitle
            LogToFile "This happens BEFORE checking if assembly itself should be skipped"
            
            On Error Resume Next
            Call AddSubcomponents(currentModel, modelsToProcess, processedModels)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error in AddSubcomponents: " & Err.description
                Err.Clear
            End If
            
            Call AddReferencedModels(currentModel, modelsToProcess, processedModels)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error in AddReferencedModels: " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            LogToFile "Subcomponents added. Queue now has " & modelsToProcess.Count & " items"
            LogToFile "=== SUBCOMPONENT ADDITION COMPLETE ==="
        End If
        ' ========================================================================
        ' END CRITICAL FIX
        ' ========================================================================

        ' NOW check if THIS model should be skipped
        Dim isSkipped As Boolean
        isSkipped = IsSkippedType(currentModel)
        LogToFile "IsSkippedType returned: " & isSkipped

        If isSkipped Then
            LogToFile "Model skipped due to type (but subcomponents were already added to queue if applicable)"
            LogModelResult currentModel.GetTitle, modelPath, "SKIPPED - TYPE", "Fastener, secret, or other skipped type"
            GoTo RemoveAndContinue
        End If

        ' Initialize critical variables
        modelType = currentModel.GetType
        canSave = IsModelCheckedOut(modelPath)
        LogToFile "Model type: " & modelType
        LogToFile "Checkout status - canSave: " & canSave

        ' Collect all properties for this model
        Set propertiesToSet = New Scripting.Dictionary
        LogToFile "Collecting all properties for model..."
        On Error Resume Next
        Call CollectAllProperties(currentModel, propertiesToSet)
        If Err.Number <> 0 Then
            LogToFile "ERROR in CollectAllProperties: " & Err.description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        LogToFile "Properties collected: " & propertiesToSet.Count

        ' Check if Reference Category = PUR and skip if PROCESS_PUR_PARTS is False
        skipForPUR = False
        If Not PROCESS_PUR_PARTS Then
            If propertiesToSet.exists("Reference Category") Then
                LogToFile "Found Reference Category: " & propertiesToSet("Reference Category")
                If UCase(Trim(propertiesToSet("Reference Category"))) = "PUR" Then
                    LogToFile "*** SKIPPING model modifications because Reference Category is PUR and PROCESS_PUR_PARTS = False ***"
                    skipForPUR = True
                End If
            End If
        Else
            LogToFile "PROCESS_PUR_PARTS = True - PUR parts will be processed normally"
        End If

        ' FIXED: Process the model and always log the result
        If skipForPUR Then
            LogToFile "Model skipped due to PUR category (subcomponents were already added to queue if applicable)"
            LogModelResult currentModel.GetTitle, modelPath, "SKIPPED - PUR", "Reference Category is PUR"
            
        ElseIf isSkipped Then
            LogToFile "Model skipped due to type"
            LogModelResult currentModel.GetTitle, modelPath, "SKIPPED - TYPE", "Fastener, secret, or other skipped type"
            
        Else
            LogToFile "Model is not a skipped type - proceeding with full processing"

            ' Create deep copy with error handling
            On Error Resume Next
            Set initialPropertiesToSet = DeepCopyDictionary(propertiesToSet)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error in DeepCopyDictionary: " & Err.description
                Err.Clear
                Set initialPropertiesToSet = propertiesToSet
            End If
            On Error GoTo ErrorHandler

            ' Apply property logic/rules with error handling
            On Error Resume Next
            AddPropertiesBasedOnRules currentModel, propertiesToSet
            Call AddCreatedDateToProperties(currentModel, propertiesToSet)
            AutofillMaterial currentModel, propertiesToSet
            If Err.Number <> 0 Then LogToFile "ERROR in AutofillMaterial: " & Err.description: Err.Clear
            
            AutofillWeight currentModel, propertiesToSet
          ' Stop
            If Err.Number <> 0 Then LogToFile "ERROR in AutofillWeight: " & Err.description: Err.Clear
            PopulateAuthorCustomProperty currentModel, propertiesToSet
            If Err.Number <> 0 Then LogToFile "ERROR in PopulateAuthorCustomProperty: " & Err.description: Err.Clear
            AutofillUnits currentModel, propertiesToSet
            If Err.Number <> 0 Then LogToFile "ERROR in AutofillUnits: " & Err.description: Err.Clear
            On Error GoTo ErrorHandler

            LogToFile "Properties after autofill count: " & propertiesToSet.Count

            ' Sort properties
            Dim sortedProperties As Scripting.Dictionary
            Set sortedProperties = New Scripting.Dictionary
            On Error Resume Next
            If modelType = swDocPART Then
                Set sortedProperties = SortProperties(propertiesToSet, GetRequiredPartPropertyRules())
            ElseIf modelType = swDocASSEMBLY Then
                Set sortedProperties = SortProperties(propertiesToSet, GetRequiredAssemblyPropertyRules())
            End If

            If Err.Number <> 0 Or sortedProperties Is Nothing Or sortedProperties.Count = 0 Then
                Err.Clear
                Set sortedProperties = propertiesToSet
            End If
            On Error GoTo ErrorHandler

            ' Check properties
            Dim failingProperties As Scripting.Dictionary
            Set failingProperties = New Scripting.Dictionary

            On Error Resume Next
            Set failingProperties = CheckProperties(currentModel, sortedProperties)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error checking properties: " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler

            If failingProperties Is Nothing Then Set failingProperties = New Scripting.Dictionary

            ' Process failing properties
            If failingProperties.Count > 0 Then
                LogToFile "Found " & failingProperties.Count & " failing properties for " & currentModel.GetTitle
                issuesDetected = True
                
                On Error Resume Next
                Dim prop As Variant
                For Each prop In failingProperties.keys
                    If Not IsEmpty(prop) And Not IsNull(prop) Then
                        Dim keyName As String
                        keyName = currentModel.GetTitle() & " - " & CStr(prop)
                        If Not aggregatedFailedProperties.exists(keyName) Then
                            aggregatedFailedProperties.Add keyName, True
                            LogToFile "Added failing property: " & keyName
                        End If
                    End If
                Next prop
                If Err.Number <> 0 Then
                    LogToFile "ERROR adding failing properties: " & Err.description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
            
            ' Check location
            On Error Resume Next
            Dim locationIsValid As Boolean
            locationIsValid = CheckLocation(currentModel)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error checking location: " & Err.description
                Err.Clear
                locationIsValid = True
            End If
            On Error GoTo ErrorHandler
            
            If Not locationIsValid Then
                LogToFile "LOCATION CHECK FAILED: " & currentModel.GetPathName
                issuesDetected = True
                
                On Error Resume Next
                If Not aggregatedFailedLocations.exists(currentModel.GetPathName) Then
                    aggregatedFailedLocations.Add currentModel.GetPathName, "Invalid file location: must be in approved PDM folder"
                    LogToFile "Added failed location: " & currentModel.GetPathName
                End If
                If Err.Number <> 0 Then
                    LogToFile "ERROR adding failed location: " & Err.description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            Else
                LogToFile "LOCATION CHECK PASSED: " & currentModel.GetPathName
            End If
            
            ' Check piece parts
            On Error Resume Next
            Dim piecePartsFailed As Boolean
            piecePartsFailed = CheckPieceParts(currentModel)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error checking piece parts: " & Err.description
                Err.Clear
                piecePartsFailed = False
            End If
            On Error GoTo ErrorHandler
            
            If piecePartsFailed Then
                LogToFile "PIECE PARTS CHECK FAILED: " & currentModel.GetTitle
                issuesDetected = True
                
                On Error Resume Next
                If Not aggregatedFailedPieceParts.exists(currentModel.GetTitle) Then
                    aggregatedFailedPieceParts.Add currentModel.GetTitle, True
                    LogToFile "Added failed piece part: " & currentModel.GetTitle
                End If
                If Err.Number <> 0 Then
                    LogToFile "ERROR adding failed piece part: " & Err.description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            Else
                LogToFile "PIECE PARTS CHECK PASSED: " & currentModel.GetTitle
            End If

            ' Apply properties
            LogToFile "Applying properties to model..."
            Dim propertiesChanged As Boolean
            propertiesChanged = False
            
            On Error Resume Next
            propertiesChanged = ApplyCollectedProperties(currentModel, sortedProperties, initialPropertiesToSet)
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error applying properties: " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler

            LogToFile "Properties changed: " & propertiesChanged

            ' Save the model only if changes were made
            If propertiesChanged Then
                LogToFile "=== SAVE OPERATION DEBUG ==="
                LogToFile "Model: " & currentModel.GetTitle
                LogToFile "Path: " & modelPath
                LogToFile "canSave flag: " & canSave
                
                If Not canSave Then
                    LogToFile "WARNING: Model not checked out - attempting save anyway"
                Else
                    LogToFile "Model is checked out - save should succeed"
                End If
                
                ' Force rebuild before saving
                LogToFile "Forcing model rebuild before save..."
                On Error Resume Next
                currentModel.ForceRebuild3 True
                If Err.Number <> 0 Then
                    LogToFile "WARNING: Error during ForceRebuild3: " & Err.description
                    Err.Clear
                Else
                    LogToFile "Model rebuild completed successfully"
                End If
                On Error GoTo ErrorHandler
                
                ' Save the model
                LogToFile "Calling Save3..."
                On Error Resume Next
                Dim saveErrors As Long, saveWarnings As Long
                saveErrors = 0
                saveWarnings = 0
                
                Dim saveResult As Boolean
                saveResult = currentModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, saveErrors, saveWarnings)
                
                LogToFile "Save3 returned: " & saveResult
                LogToFile "Save errors: " & saveErrors
                LogToFile "Save warnings: " & saveWarnings
                
                If Err.Number <> 0 Then
                    LogToFile "VBA ERROR during save: " & Err.description & " (Error #: " & Err.Number & ")"
                    Err.Clear
                ElseIf saveErrors <> 0 Then
                    LogToFile "SOLIDWORKS SAVE ERRORS: " & saveErrors
                ElseIf Not saveResult Then
                    LogToFile "SAVE FAILED: Save3 returned False"
                Else
                    LogToFile "MODEL SAVED SUCCESSFULLY"
                End If
                
                LogToFile "=== END SAVING MODEL CHANGES ==="
            Else
                LogToFile "No properties changed - SKIPPING SAVE and REBUILD"
            End If

            ' FIXED: Determine final status and log result
            Dim finalStatus As String
            Dim statusDetails As String
            
            If Not canSave Then
                finalStatus = "READ-ONLY"
                statusDetails = "Not checked out - properties updated in memory only"
            ElseIf failingProperties.Count > 0 Or Not locationIsValid Or piecePartsFailed Then
                finalStatus = "COMPLETED WITH ISSUES"
                statusDetails = "Prop failures: " & failingProperties.Count & ", Location OK: " & locationIsValid & ", PieceParts OK: " & Not piecePartsFailed
            Else
                If propertiesChanged Then
                    finalStatus = "SUCCESS"
                    statusDetails = "All validations passed, properties updated and saved"
                Else
                    finalStatus = "SUCCESS (NO CHANGES)"
                    statusDetails = "Properties already correct - no save needed"
                End If
            End If
            
            ' Log the result for this model
            LogModelResult currentModel.GetTitle, modelPath, finalStatus, statusDetails
        End If

        ' --- Rebuild Drawing if Originally a Drawing ---
        If isOriginallyDrawing And Not originalDrawing Is Nothing Then
            LogToFile "=== REBUILDING ORIGINAL DRAWING ==="
            LogToFile "Drawing: " & originalDrawing.GetTitle
            
            On Error Resume Next
            originalDrawing.ForceRebuild3 True
            If Err.Number <> 0 Then
                LogToFile "WARNING: Error during drawing rebuild: " & Err.description
                Err.Clear
            Else
                LogToFile "Drawing rebuilt successfully"
            End If
            
            If originalDrawing.GetType = swDocDRAWING Then
                Dim swDraw As drawingDoc
                Set swDraw = originalDrawing
                Dim drawingView As view
                Set drawingView = swDraw.GetFirstView
                
                While Not drawingView Is Nothing
                    drawingView.Rebuild
                    LogToFile "Rebuilt drawing view: " & drawingView.Name
                    Set drawingView = drawingView.GetNextView
                Wend
            End If
            
            On Error GoTo ErrorHandler
            LogToFile "=== DRAWING REBUILD COMPLETE ==="
        End If

RemoveAndContinue:
        ' FIXED: Add to processed list only at the very end
        If Not processedModels.exists(modelPath) Then
            processedModels.Add modelPath, currentModel
            LogToFile "Added '" & modelPath & "' to processed list"
        End If

        If modelsToProcess.exists(modelPath) Then
            modelsToProcess.Remove modelPath
            LogToFile "Removed '" & modelPath & "' from processing queue."
        End If

        ' OPTIMIZATION: Close subcomponent models after processing (to prevent memory bloat)
        ' Only close if this is NOT the initial model (assembly/part we started with)
        ' Assembly subcomponents that were loaded as part of opening the assembly should be closed
        If modelPath <> initialPath Then
            ' Check if this model is a subcomponent (part) - assemblies should stay open
            shouldClose = False

            On Error Resume Next
            If Not currentModel Is Nothing Then
                If currentModel.GetType = swDocPART Then
                    shouldClose = True
                    LogToFile "Model is a PART - will close after processing"
                Else
                    LogToFile "Model is an ASSEMBLY - keeping open"
                End If
            Else
                ' currentModel was set to Nothing, use the path to decide
                ' If it ends in .SLDPRT, close it
                If UCase(Right(modelPath, 7)) = ".SLDPRT" Then
                    shouldClose = True
                    LogToFile "Model path indicates PART - will close after processing"
                End If
            End If
            On Error GoTo ErrorHandler

            If shouldClose Then
                LogToFile "Closing subcomponent after processing: " & modelPath
                On Error Resume Next
                swApp.CloseDoc modelPath
                If Err.Number <> 0 Then
                    LogToFile "WARNING: Could not close model: " & Err.description
                    Err.Clear
                Else
                    LogToFile "Successfully closed: " & modelPath
                End If
                On Error GoTo ErrorHandler
            End If
        Else
            LogToFile "Keeping initial model open: " & modelPath
        End If

        Set currentModel = Nothing

ContinueLoop:
    Loop

    ' Final Statistics
    LogToFile vbNewLine & "=== Final Statistics ==="
    LogToFile "Total failed properties recorded: " & aggregatedFailedProperties.Count
    LogToFile "Total failed locations: " & aggregatedFailedLocations.Count
    LogToFile "Total failed piece parts: " & aggregatedFailedPieceParts.Count
    LogToFile "Total models processed: " & processedModels.Count
    LogToFile "Final issuesDetected flag value: " & issuesDetected
    
    LogToFile "=== Finished ProcessModels ==="

    Exit Sub

ErrorHandler:
    LogToFile "%%% CRITICAL ERROR in ProcessModels: " & Err.description & " (Line: " & Erl & ") %%%"
    LogToFile "Error Number: " & Err.Number
    Dim errorModelPath As String
    If Not currentModel Is Nothing Then
         On Error Resume Next
         errorModelPath = currentModel.GetPathName
         If Err.Number = 0 And errorModelPath <> "" Then
              If modelsToProcess.exists(errorModelPath) Then
                   modelsToProcess.Remove errorModelPath
              End If
         End If
         Err.Clear
         On Error GoTo 0
    End If
    MsgBox "A critical error occurred during model processing: " & Err.description & vbCrLf & "Attempting to continue...", vbCritical
    Resume ContinueLoop
End Sub


    
Private Function IsMultiBodyPart(model As ModelDoc2) As Boolean
    On Error GoTo ErrorHandler
    
    If model.GetType <> swDocPART Then
        IsMultiBodyPart = False
        Exit Function
    End If
    
    Dim swPart As PartDoc
    Set swPart = model
    
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, False)
    
    ' Skip model if we can't properly evaluate bodies
    If IsEmpty(vBodies) Then
        DebugLog "Model skipped (No valid bodies): " & model.GetPathName
        IsMultiBodyPart = True ' Return True to skip processing
        Exit Function
    End If
    
    IsMultiBodyPart = UBound(vBodies) > 0
    Exit Function
    
ErrorHandler:
    DebugLog "Model skipped (Body evaluation error): " & model.GetPathName
    IsMultiBodyPart = True ' Return True to skip processing
End Function
    
    ' Create a deep copy of a Scripting.Dictionary
Private Function DeepCopyDictionary(orig As Object) As Scripting.Dictionary
    Dim copy As New Scripting.Dictionary
    Dim key As Variant
    For Each key In orig.keys
        copy.Add key, orig(key)
    Next key
    Set DeepCopyDictionary = copy
End Function
    
    Private Sub AddSubcomponents(model As ModelDoc2, ByRef modelsToProcess As Scripting.Dictionary, ByRef processedModels As Scripting.Dictionary)
    LogToFile "=== ENTERING AddSubcomponents ==="
    LogToFile "Model: " & model.GetTitle & " (" & model.GetPathName() & ")"
    
    ' Check if model is Nothing
    If model Is Nothing Then
        LogToFile "ERROR: Model is Nothing in AddSubcomponents"
        Exit Sub
    End If
    
    ' Check model type
    LogToFile "Model type: " & model.GetType
    If model.GetType <> swDocASSEMBLY Then
        LogToFile "Model is not an assembly, exiting AddSubcomponents"
        Exit Sub
    End If
    
    ' Cast the model to an assembly document
    Dim swAssy As AssemblyDoc
    Set swAssy = model
    
    If swAssy Is Nothing Then
        LogToFile "ERROR: Could not cast model to AssemblyDoc"
        Exit Sub
    End If
    
    LogToFile "Successfully cast to AssemblyDoc"
    
    ' Try multiple approaches to get components
    Dim components As Variant
    Dim componentsObtained As Boolean
    componentsObtained = False
    
    ' ATTEMPT 1: Get all components (including lightweight)
    LogToFile "ATTEMPT 1: Getting components with GetComponents(True)..."
    On Error Resume Next
    components = swAssy.GetComponents(True)
    If Err.Number = 0 And Not IsEmpty(components) Then
        LogToFile "SUCCESS: GetComponents(True) returned " & (UBound(components) + 1) & " components"
        componentsObtained = True
    Else
        LogToFile "FAILED: GetComponents(True) error: " & Err.description & " (Error #: " & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0
    
    ' ATTEMPT 2: Get only resolved components if first attempt failed
    If Not componentsObtained Then
        LogToFile "ATTEMPT 2: Getting only resolved components with GetComponents(False)..."
        On Error Resume Next
        components = swAssy.GetComponents(False)
        If Err.Number = 0 And Not IsEmpty(components) Then
            LogToFile "SUCCESS: GetComponents(False) returned " & (UBound(components) + 1) & " components"
            componentsObtained = True
        Else
            LogToFile "FAILED: GetComponents(False) error: " & Err.description & " (Error #: " & Err.Number & ")"
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' ATTEMPT 3: Try to resolve lightweight components first, then get components
    If Not componentsObtained Then
        LogToFile "ATTEMPT 3: Resolving lightweight components first..."
        On Error Resume Next
        swAssy.ResolveAllLightWeightComponents False
        If Err.Number = 0 Then
            LogToFile "Lightweight components resolved, trying GetComponents again..."
            components = swAssy.GetComponents(True)
            If Err.Number = 0 And Not IsEmpty(components) Then
                LogToFile "SUCCESS: After resolving lightweight, got " & (UBound(components) + 1) & " components"
                componentsObtained = True
            Else
                LogToFile "FAILED: Even after resolving lightweight: " & Err.description & " (Error #: " & Err.Number & ")"
            End If
        Else
            LogToFile "FAILED: Could not resolve lightweight components: " & Err.description
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    ' If all attempts failed, log and exit
    If Not componentsObtained Then
        LogToFile "CRITICAL: All attempts to get components failed for assembly: " & model.GetTitle
        LogToFile "This assembly will not have its subcomponents processed"
        Exit Sub
    End If
    
    ' Check if components array is valid
    If IsEmpty(components) Then
        LogToFile "Components array is empty (no components in assembly)"
        Exit Sub
    End If
    
    LogToFile "Processing " & (UBound(components) + 1) & " components..."
    
    Dim componentCount As Integer
    componentCount = 0
    Dim successCount As Integer
    successCount = 0
    Dim errorCount As Integer
    errorCount = 0
    
    Dim component As Variant
    For Each component In components
        componentCount = componentCount + 1
        LogToFile "--- Processing component " & componentCount & " ---"
        
        If component Is Nothing Then
            LogToFile "WARNING: Component " & componentCount & " is Nothing"
            errorCount = errorCount + 1
            GoTo NextComponent
        End If
        
        LogToFile "Component name: " & component.Name2
        
        Dim compModel As ModelDoc2
        On Error Resume Next
        Set compModel = component.GetModelDoc2
        If Err.Number <> 0 Then
            LogToFile "ERROR: Could not get ModelDoc2 for component " & component.Name2 & ": " & Err.description
            Err.Clear
            errorCount = errorCount + 1
            GoTo NextComponent
        End If
        On Error GoTo 0
        
        If compModel Is Nothing Then
            LogToFile "WARNING: ModelDoc2 is Nothing for component " & component.Name2
            errorCount = errorCount + 1
            GoTo NextComponent
        End If
        
        LogToFile "Component model title: " & compModel.GetTitle
        
        Dim compPath As String
        On Error Resume Next
        compPath = compModel.GetPathName()
        If Err.Number <> 0 Or compPath = "" Then
            LogToFile "ERROR: Could not get path for component " & compModel.GetTitle & ": " & Err.description
            Err.Clear
            errorCount = errorCount + 1
            GoTo NextComponent
        End If
        On Error GoTo 0
        
        LogToFile "Component path: " & compPath
        
        ' Check if already processed
        If processedModels.exists(compPath) Then
            LogToFile "Component already processed, skipping"
            GoTo NextComponent
        End If
        
        ' Check if already in processing queue
        If modelsToProcess.exists(compPath) Then
            LogToFile "Component already in processing queue"
            GoTo NextComponent
        End If
        
        ' Add to processing queue
        On Error Resume Next
        modelsToProcess.Add compPath, compModel
        If Err.Number <> 0 Then
            LogToFile "ERROR: Could not add component to processing queue: " & Err.description
            Err.Clear
            errorCount = errorCount + 1
            GoTo NextComponent
        End If
        On Error GoTo 0
        
        LogToFile "SUCCESS: Added to processing queue: " & compModel.GetTitle
        successCount = successCount + 1
        
        ' Recursively process subcomponents if it's an assembly
        If compModel.GetType = swDocASSEMBLY Then
            LogToFile "Component is assembly, recursively processing: " & compModel.GetTitle
            AddSubcomponents compModel, modelsToProcess, processedModels
        End If
        
NextComponent:
    Next component
    
    LogToFile "=== ADDSUBCOMPONENTS SUMMARY ==="
    LogToFile "Total components found: " & componentCount
    LogToFile "Successfully added to queue: " & successCount
    LogToFile "Errors encountered: " & errorCount
    LogToFile "=== EXITING AddSubcomponents ==="
End Sub
    
    
    Private Sub AddReferencedModels(model As ModelDoc2, ByRef modelsToProcess As Scripting.Dictionary, ByRef processedModels As Scripting.Dictionary)
        '''''debug.Print "Entering AddReferencedModels function"
        '''''debug.Print "Model path: " & model.GetPathName()
        
        If model Is Nothing Then
            '''''debug.Print "Error: Model is Nothing"
            Exit Sub
        End If
    
        Dim swFeat As Feature
        Set swFeat = model.FirstFeature
    
        '''''debug.Print "Starting to iterate through features"
        While Not swFeat Is Nothing
            '''''debug.Print "Processing feature: " & swFeat.Name & " (Type: " & swFeat.GetTypeName2 & ")"
            
            If swFeat.GetTypeName2 = "RefFeature" Then
                Dim swRefFeat As Object
                Set swRefFeat = swFeat.GetSpecificFeature2
    
                If swRefFeat Is Nothing Then
                    '''''debug.Print "Warning: RefFeature is Nothing"
                Else
                    Dim refModel As ModelDoc2
                    On Error Resume Next
                    Set refModel = swRefFeat.GetModelDoc
                    If Err.Number <> 0 Then
                        '''''debug.Print "Error getting referenced document: " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
    
                    If Not refModel Is Nothing Then
                        '''''debug.Print "Referenced model found: " & refModel.GetPathName()
                        If Not processedModels.exists(refModel.GetPathName()) Then
                            If Not modelsToProcess.exists(refModel.GetPathName()) Then
                                modelsToProcess.Add refModel.GetPathName(), refModel
                                '''''debug.Print "Added referenced model to process: " & refModel.GetPathName()
                            End If
                        End If
                    Else
                        '''''debug.Print "Warning: Referenced model is Nothing"
                    End If
                End If
            End If
    
            Set swFeat = swFeat.GetNextFeature
        Wend
    
        If model.GetType = swDocASSEMBLY Then
            Dim swAssy As AssemblyDoc
            Set swAssy = model
    
            '''''debug.Print "Processing assembly components"
            On Error Resume Next
            Dim components As Variant
            components = swAssy.GetComponents(True)
            If Err.Number <> 0 Then
                '''''debug.Print "Error getting components: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
    
            If Not IsEmpty(components) Then
                '''''debug.Print "Number of components: " & UBound(components) + 1
                Dim component As Variant
                For Each component In components
                    On Error Resume Next
                    Dim compModel As ModelDoc2
                    Set compModel = component.GetModelDoc2
                    If Err.Number <> 0 Then
                        '''''debug.Print "Error getting component model: " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
    
                    If Not compModel Is Nothing Then
                        '''''debug.Print "Processing component: " & component.Name2
                        If Not processedModels.exists(compModel.GetPathName()) Then
                            If Not modelsToProcess.exists(compModel.GetPathName()) Then
                                modelsToProcess.Add compModel.GetPathName(), compModel
                                '''''debug.Print "Added component model to process: " & compModel.GetPathName()
                            End If
                        End If
                    Else
                        '''''debug.Print "Warning: Component model is Nothing for " & component.Name2
                    End If
                Next component
            Else
                '''''debug.Print "No components found in the assembly"
            End If
        End If
    
        '''''debug.Print "Exiting AddReferencedModels function"
    End Sub
    
   ' Collect all properties, including existing ones
Private Sub CollectAllProperties(model As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    On Error GoTo ErrorHandler
    
    DebugLog "Starting CollectAllProperties..."
    DebugLog "Checking model..."
    
    If model Is Nothing Then
        DebugLog "ERROR: Model is Nothing in CollectAllProperties"
        Exit Sub
    End If
    
    DebugLog "Getting CustomPropertyManager..."
    Dim custPropMgr As customPropertyManager
    On Error Resume Next
    Set custPropMgr = model.Extension.customPropertyManager("")
    If Err.Number <> 0 Then
        DebugLog "ERROR getting customPropertyManager: " & Err.description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If custPropMgr Is Nothing Then
        DebugLog "ERROR: Failed to get customPropertyManager"
        Exit Sub
    End If
    
    DebugLog "CustomPropertyManager obtained successfully"
    Dim propNames As Variant
    Dim propName As Variant
    Dim propValue As String
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    Dim linkToProperty As Boolean
    
    ' Get all existing custom property names with error handling
    On Error Resume Next
    DebugLog "Getting property names..."
    propNames = custPropMgr.GetNames
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error getting property names: " & Err.description
        Err.Clear
        propNames = Array()
    End If
    On Error GoTo ErrorHandler
    
    DebugLog "Got property names"
    
    ' Collect existing properties with error handling
    If Not IsEmpty(propNames) Then
        DebugLog "Processing " & UBound(propNames) + 1 & " properties..."
        For Each propName In propNames
            On Error Resume Next
            DebugLog "Processing property: " & propName
            custPropMgr.Get6 CStr(propName), False, propValue, resolvedValue, wasResolved, linkToProperty
            If Err.Number = 0 Then
                propertiesToSet(CStr(propName)) = Trim(resolvedValue)
                DebugLog "Added property: " & propName & " = " & Trim(resolvedValue)
            Else
                DebugLog "WARNING: Error getting property " & propName & ": " & Err.description
                propertiesToSet(CStr(propName)) = ""
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Next
    Else
        DebugLog "No existing properties found"
    End If
    
    ' Add missing required properties with error handling
    DebugLog "Adding required properties..."
    Dim requiredProps As Object
    On Error Resume Next
    If model.GetType = swDocPART Then
        Set requiredProps = GetRequiredPartPropertyRules()
        DebugLog "Getting part property rules"
    ElseIf model.GetType = swDocASSEMBLY Then
        Set requiredProps = GetRequiredAssemblyPropertyRules()
        DebugLog "Getting assembly property rules"
    End If
    
    If Err.Number <> 0 Then
        DebugLog "ERROR getting required properties: " & Err.description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If Not requiredProps Is Nothing Then
        DebugLog "Processing required properties..."
        On Error Resume Next
        For Each propName In requiredProps.keys
            If Not propertiesToSet.exists(propName) Then
                propertiesToSet(propName) = ""
                DebugLog "Added missing required property: " & propName
            End If
        Next
        If Err.Number <> 0 Then
            DebugLog "ERROR processing required properties: " & Err.description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Else
        DebugLog "No required properties rules found"
    End If
    
    ' Additional processing functions with error handling
    DebugLog "Running additional property processing..."
    
    On Error Resume Next
    DebugLog "Running AddPropertiesBasedOnRules..."
    Call AddPropertiesBasedOnRules(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in AddPropertiesBasedOnRules: " & Err.description
        Err.Clear
    End If
    
    DebugLog "Running AutofillUnits..."
    Call AutofillUnits(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in AutofillUnits: " & Err.description
        Err.Clear
    End If
    
    DebugLog "Running AddDimensionsProperties..."
    Call AddDimensionsProperties(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in AddDimensionsProperties: " & Err.description
        Err.Clear
    End If

    ' *** IMPORTANT: AutofillMaterial MUST run BEFORE ExtractStructuralMemberType ***
    ' ExtractStructuralMemberType calls ProcessPLorCP which generates Mtl Part Number
    ' using the Material property. If Material isn't updated first, Mtl Part Number
    ' will use the OLD material value from previously saved custom properties.
    DebugLog "Running AutofillMaterial (before ExtractStructuralMemberType)..."
    Call AutofillMaterial(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in AutofillMaterial: " & Err.description
        Err.Clear
    End If

    DebugLog "Running ExtractStructuralMemberType..."
    Call ExtractStructuralMemberType(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in ExtractStructuralMemberType: " & Err.description
        Err.Clear
    End If
'    Stop
    DebugLog "Running RemoveProperties..."
    Call RemoveProperties(model, propertiesToSet)
    
    DebugLog "Running SetNonDestructiveTesting..."
    Call SetNonDestructiveTesting(propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in SetNonDestructiveTesting: " & Err.description
        Err.Clear
    End If
    
    DebugLog "Running PopulateAuthorCustomProperty..."
    Call PopulateAuthorCustomProperty(model, propertiesToSet)
    If Err.Number <> 0 Then
        DebugLog "WARNING: Error in PopulateAuthorCustomProperty: " & Err.description
        Err.Clear
    End If
    
    ' ============================================================================
    ' CLEANUP PROPERTY NAMES - Run this AFTER all properties are collected/added
    ' This standardizes capitalization to match PropertyRules
    ' ============================================================================
    DebugLog ""
    DebugLog "=========================================="
    DebugLog "=== ABOUT TO CLEAN UP PROPERTY NAMES ==="
    DebugLog "=========================================="
    DebugLog "Properties before cleanup: " & propertiesToSet.Count
    
    ' Print all properties BEFORE cleanup
    Dim debugKey As Variant
    DebugLog "--- ALL PROPERTIES BEFORE CLEANUP ---"
    For Each debugKey In propertiesToSet.keys
        DebugLog "  '" & debugKey & "' = '" & propertiesToSet(debugKey) & "'"
    Next debugKey
    
    DebugLog ""
    DebugLog "Calling CleanupPropertyNames now..."
    Call CleanupPropertyNames(model, propertiesToSet)
    DebugLog "CleanupPropertyNames returned"
    
 '   Stop  ' *** STOP HERE TO INSPECT RESULTS ***
    
    DebugLog ""
    DebugLog "Properties after cleanup: " & propertiesToSet.Count
    
    ' Print all properties AFTER cleanup
    DebugLog "--- ALL PROPERTIES AFTER CLEANUP ---"
    For Each debugKey In propertiesToSet.keys
        DebugLog "  '" & debugKey & "' = '" & propertiesToSet(debugKey) & "'"
    Next debugKey
    
    DebugLog "=========================================="
    DebugLog "=== CLEANUP COMPLETE ==="
    DebugLog "=========================================="
    ' ============================================================================
    
    On Error GoTo ErrorHandler
    DebugLog "CollectAllProperties completed successfully with " & propertiesToSet.Count & " properties"
    Exit Sub

ErrorHandler:
    DebugLog "ERROR in CollectAllProperties at step: " & Err.Source & vbCrLf & _
                "Description: " & Err.description & vbCrLf & _
                "Line: " & Erl & vbCrLf & _
                "Error Number: " & Err.Number
    Resume Next
End Sub
    
    
    ' Join keys and values of a Scripting.Dictionary into a string
    Private Function JoinKeysAndValues(properties As Scripting.Dictionary) As String
        Dim result As String
        Dim key As Variant
        For Each key In properties
            result = result & key & ": " & properties(key) & "; "
        Next key
        JoinKeysAndValues = result
    End Function
    
    ' Set a property value in a model
    Private Sub SetProperty(model As ModelDoc2, propertyName As Variant, propertyValue As String, ByRef initialPropertiesToSet As Scripting.Dictionary)
        Dim customPropertyManager As customPropertyManager
        Set customPropertyManager = model.Extension.customPropertyManager("")
    
        ' Determine if the property needs to be added or updated
        Dim propExists As Boolean
        propExists = customPropertyManager.Get(propertyName) <> ""
    
        If Not propExists Then
            ' Property doesn't exist, add it
            customPropertyManager.Add2 propertyName, swCustomInfoText, propertyValue
        Else
            ' Property exists, update its value
            customPropertyManager.Set2 propertyName, propertyValue
        End If
    End Sub
    
Private Function IsSkippedType(model As ModelDoc2) As Boolean
    On Error GoTo ErrorHandler
    
    ' ===================================================================
    ' FIRST CHECK: Library path - skip everything in Libraries folder
    ' ===================================================================
    Dim modelPath As String
    modelPath = model.GetPathName()
    
    ' Skip all files in the Libraries folder
    If InStr(1, LCase(modelPath), "c:\nmt_pdm\libraries", vbTextCompare) > 0 Then
        DebugLog "Model skipped (Library path): " & model.GetTitle & " (" & modelPath & ")"
        IsSkippedType = True
        Exit Function
    End If
    
    ' ===================================================================
    ' SECOND CHECK: Secret property
    ' ===================================================================
    Dim swCustPropMgr As customPropertyManager
    Dim results As Object
    Set swCustPropMgr = model.Extension.customPropertyManager("")
    Set results = CreateObject("Scripting.Dictionary")
    
    ' Properties to check
    Dim properties As Variant
    properties = Array("IsFastener", "Secret")
    
    Dim ret As Boolean
    Dim valOut As String
    Dim resolvedValOut As String
    Dim wasResolved As Boolean
    Dim linkToProperty As Boolean
    
    ' Retrieve and check properties
    Dim prop As Variant
    For Each prop In properties
        ret = swCustPropMgr.Get6(prop, False, valOut, resolvedValOut, wasResolved, linkToProperty)
        If wasResolved Then
            results(prop) = Trim(resolvedValOut)
            DebugLog prop & " Value: '" & results(prop) & "' for " & model.GetTitle
        Else
            results(prop) = ""
        End If
    Next
    
    ' Check Secret level - BEFORE other checks
    If results.exists("Secret") Then
        Dim secretValue As String
        secretValue = UCase(Trim(results("Secret")))
        
        ' Skip if Secret is "1", "LEVEL1", or "TRUE"
        If secretValue = "1" Or secretValue = "LEVEL1" Or secretValue = "TRUE" Then
            DebugLog "Model skipped (Secret Level 1): " & model.GetTitle
            IsSkippedType = True
            Exit Function
        End If
    End If
    
    ' ===================================================================
    ' THIRD CHECK: IsFastener property
    ' ===================================================================
    If results.exists("IsFastener") And Len(results("IsFastener")) > 0 Then
        DebugLog "Model skipped (IsFastener): " & model.GetTitle
        IsSkippedType = True
        Exit Function
    End If
    
    ' ===================================================================
    ' FOURTH CHECK: Reference Category
    ' ===================================================================
    Dim swDocExt As ModelDocExtension
    Set swDocExt = model.Extension
    
    If Not swDocExt Is Nothing Then
        Dim refCategory As String
        refCategory = ""
        
        ' Try to get the reference category
        On Error Resume Next
        refCategory = swDocExt.GetStringProperty("REFCATEGORY")
        On Error GoTo ErrorHandler
        
        If Len(refCategory) > 0 Then
            DebugLog "Reference Category: " & refCategory & " for model: " & model.GetTitle
            
            ' Using the original type codes to check against reference category
            Dim skippedTypes As Variant
            skippedTypes = Array("MI", "FW", "MB", "HUC", "NUT", "RIV", "WSR", "HYD", "MIS")
            
            If Contains(skippedTypes, UCase(refCategory)) Then
                DebugLog "Model skipped (Reference Category): " & refCategory & " - " & model.GetTitle
                IsSkippedType = True
                Exit Function
            End If
        End If
    End If
    
    ' If we made it here, don't skip
    IsSkippedType = False
    Exit Function
    
ErrorHandler:
    DebugLog "Error in IsSkippedType: " & Err.description
    IsSkippedType = False
    Resume Next

End Function

    ' Check if an array contains a specific value
    Private Function Contains(arr As Variant, value As String) As Boolean
        Dim element As Variant
        For Each element In arr
            If element = value Then
                Contains = True
                Exit Function
            End If
        Next element
        Contains = False
    End Function
    
    ' Check if a model's location is valid
    Function CheckLocation(model As ModelDoc2) As Boolean
        Dim Locations As collection
        Set Locations = GetLocations()
    
        Dim evaluator As Object
        Set evaluator = CreateObject("VBScript.RegExp")
        evaluator.Global = True
        evaluator.IgnoreCase = True
    
        ' Check each location pattern
        Dim locationPattern As Variant
        For Each locationPattern In Locations
            evaluator.pattern = locationPattern
            If evaluator.Test(model.GetPathName()) Then
                CheckLocation = True
                Exit Function
            End If
        Next locationPattern
    
        CheckLocation = False
    End Function
    
    ' Check for failed piece parts in a model
    Function CheckPieceParts(compModel As ModelDoc2) As Boolean
        On Error GoTo ErrorHandler  ' Enable error handling to capture unexpected errors
    
        Dim swPart As PartDoc
        Dim vBodies As Variant
        Dim nBodyCount As Long
    
        If compModel Is Nothing Then
            CheckPieceParts = False
            Exit Function
        End If
    
        ' Check if the model is a part document
        If compModel.GetType = swDocPART Then
            Set swPart = compModel
            ' Safely attempt to get bodies
            vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, False)
            
            ' Check if vBodies is an array and has elements
            If IsArray(vBodies) Then
                On Error Resume Next
                nBodyCount = UBound(vBodies) - LBound(vBodies) + 1
                On Error GoTo ErrorHandler
                
                If nBodyCount > 1 Then
                    CheckPieceParts = True
                Else
                    CheckPieceParts = False
                End If
            Else
                ' vBodies is not an array, which means there are no bodies
                CheckPieceParts = False
            End If
        Else
            CheckPieceParts = False
        End If
    
        Exit Function
    
ErrorHandler:
        '''''debug.Print "Error in CheckPieceParts: " & Err.Description
        CheckPieceParts = False
        Resume Next
    End Function
    
   Function CheckProperty(model As ModelDoc2, propertyName As String, propertyRule As String, Optional propValue As String = "") As Boolean
    DebugLog "-----------------------------"
    DebugLog "CheckProperty Details:"
    DebugLog "Property Name: " & propertyName
    DebugLog "Property Rule: " & propertyRule
    
    Dim evaluator As Object
    Set evaluator = CreateObject("VBScript.RegExp")
    evaluator.pattern = propertyRule
    evaluator.Global = False
    evaluator.IgnoreCase = False ' Make case sensitive
    
    If propValue = "" Then
        ' Property value not provided, retrieve it from the model
        Dim swCustPropMgr As customPropertyManager
        Set swCustPropMgr = model.Extension.customPropertyManager("")
    
        Dim resolvedValue As String
        swCustPropMgr.Get6 propertyName, False, propValue, resolvedValue, False, False
        propValue = Trim(resolvedValue) ' Trim whitespace
    End If
    
    DebugLog "Property Value: '" & propValue & "'"
    DebugLog "Matches Rule: " & evaluator.Test(propValue)
    DebugLog "-----------------------------"
    
    CheckProperty = evaluator.Test(propValue)
End Function
    
    ' Check model properties against rules
    Sub CheckModelProperties(model As ModelDoc2, PropertyRules As Object, ByRef failedProperties As collection, ByRef propertiesToSet As Object)
        For Each propName In PropertyRules.keys
            If Not CheckProperty(model, CStr(propName), CStr(PropertyRules(propName)), propertiesToSet) Then
                failedProperties.Add propName
            End If
        Next
    End Sub
    
Function CheckProperties(model As ModelDoc2, propertiesToSet As Scripting.Dictionary) As Scripting.Dictionary
    Dim failingProperties As New Scripting.Dictionary
    Dim PropertyRules As Object
    Dim skipPurchaseValidation As Boolean
    Dim purchaseSkipProps As Object
    
    DebugLog vbNewLine & "=== CheckProperties for " & model.GetTitle & " ==="
    
    ' Ensure the model object is valid
    If model Is Nothing Then
        MsgBox "Invalid model object passed to CheckProperties"
        Set CheckProperties = failingProperties
        Exit Function
    End If
    
    ' Setup property rules based on model type
    If model.GetType = swDocPART Then
        Set PropertyRules = GetRequiredPartPropertyRules()
        DebugLog "Using Part property rules"

        ' Purchased parts intentionally clear a subset of properties; skip validation for those.
        Dim refCategory As String
        refCategory = ""

        If propertiesToSet.exists("Reference Category") Then
            refCategory = UCase(Trim(CStr(propertiesToSet("Reference Category"))))
        End If
        skipPurchaseValidation = (refCategory = "PUR")
        If skipPurchaseValidation Then
            Set purchaseSkipProps = CreateObject("Scripting.Dictionary")
            purchaseSkipProps.Add "Material", True
            purchaseSkipProps.Add "Length", True
            purchaseSkipProps.Add "LengthA", True
            purchaseSkipProps.Add "Height", True
            purchaseSkipProps.Add "Width", True
            purchaseSkipProps.Add "MtlUOM", True
            DebugLog "Reference Category PUR detected - skipping validation for Material/Length/LengthA/Height/Width/MtlUOM"
        End If
    ElseIf model.GetType = swDocASSEMBLY Then
        Set PropertyRules = GetRequiredAssemblyPropertyRules()
        DebugLog "Using Assembly property rules"
    End If
    
    ' Check each property against the rules
    Dim propName As Variant
    For Each propName In PropertyRules.keys
        If skipPurchaseValidation Then
            If purchaseSkipProps.exists(CStr(propName)) Then
                DebugLog vbNewLine & "Skipping purchased-part property check: " & CStr(propName)
                GoTo ContinuePropertyLoop
            End If
        End If

        DebugLog vbNewLine & "Checking property: " & propName
        
        Dim propValue As String
        If propertiesToSet.exists(propName) Then
            propValue = propertiesToSet(propName)
            DebugLog "Found value: '" & propValue & "'"
        Else
            propValue = ""
            DebugLog "Property not found in Scripting.Dictionary"
        End If
        
        Dim rulePattern As String
        rulePattern = PropertyRules(propName)
        DebugLog "Using rule pattern: " & rulePattern
        
        If Not CheckProperty(model, CStr(propName), rulePattern, propValue) Then
            DebugLog "*** FAILED: Property '" & propName & "' with value '" & propValue & "' failed rule '" & rulePattern & "'"
            failingProperties.Add propName, True
        Else
            DebugLog "PASSED: Property validation succeeded"
        End If
ContinuePropertyLoop:
    Next
    
    DebugLog vbNewLine & "Total failing properties: " & failingProperties.Count
    If failingProperties.Count > 0 Then
        DebugLog "Failed properties list:"
        Dim failedProp As Variant
        For Each failedProp In failingProperties.keys
            DebugLog "- " & failedProp
        Next failedProp
    End If
    DebugLog "=== Finished CheckProperties ===" & vbNewLine
    
    Set CheckProperties = failingProperties
End Function
    
   Sub CheckModel(model As ModelDoc2, ByRef issuesReport As String, ByRef propertiesToSet As Scripting.Dictionary, ByRef failingProperties As Scripting.Dictionary, ByRef failedLocations As Scripting.Dictionary, ByRef failedPieceParts As Scripting.Dictionary, ByRef issuesDetected As Boolean)
    
           Dim locationIsValid As Boolean
        Dim piecePartsFailed As Boolean
        
        ' Check properties
      '  Set failingProperties = CheckProperties(model, propertiesToSet)
        
        ' Check locations
        locationIsValid = CheckLocation(model)
        If Not locationIsValid Then
            issuesDetected = True
            failedLocations.Add model.GetPathName, model.GetPathName  ' Store full path
        End If
        
        ' Check piece parts
        piecePartsFailed = CheckPieceParts(model)
        If piecePartsFailed Then
            issuesDetected = True
            ' Check if the key already exists in the Scripting.Dictionary before adding
            If Not failedPieceParts.exists(model.GetTitle) Then
                failedPieceParts.Add model.GetTitle, True
            End If
        End If
        
        ' Update issues report
        If failingProperties.Count > 0 Or Not locationIsValid Or piecePartsFailed Then
            issuesReport = issuesReport & vbNewLine & "Issues found in model: " & model.GetTitle
            If failingProperties.Count > 0 Then
                issuesReport = issuesReport & vbNewLine & "Failed properties:"
                For Each prop In failingProperties.keys
                    issuesReport = issuesReport & vbNewLine & "- " & prop
                Next prop
            End If
            If Not locationIsValid Then
                issuesReport = issuesReport & vbNewLine & "- Invalid location: " & model.GetPathName
            End If
            If piecePartsFailed Then
                issuesReport = issuesReport & vbNewLine & "- Failed piece parts check"
            End If
        End If
    End Sub
    
    
    ' Create a Scripting.Dictionary from arrays
Private Function CreateDictFromArrays(keys As Variant, values As Variant) As Scripting.Dictionary
Dim dict As New Scripting.Dictionary
    Dim i As Long
        For i = LBound(keys) To UBound(keys)
            dict.Add keys(i), values(i)
        Next i
    
        Set CreateDictFromArrays = dict
    End Function
    
    ' Get the index of a property in an array
    Private Function GetPropertyIndex(propName As String, propNamesArr() As String) As Long
        Dim i As Long
        For i = LBound(propNamesArr) To UBound(propNamesArr)
            If propNamesArr(i) = propName Then
                GetPropertyIndex = i
                Exit Function
            End If
        Next i
        GetPropertyIndex = -1
    End Function
    
    ' Get the value of a property from arrays
    Private Function GetPropertyValueFromArrays(propName As String, propNames As Variant, propValues As Variant) As String
        Dim i As Integer
        For i = 0 To UBound(propNames)
            If propNames(i) = propName Then
                GetPropertyValueFromArrays = propValues(i)
                Exit Function
            End If
        Next i
        GetPropertyValueFromArrays = ""
    End Function
    
' Apply collected properties, updating only what's necessary
' NON-DESTRUCTIVE: Preserves user-added custom properties that aren't managed by this macro
Function ApplyCollectedProperties(model As ModelDoc2, propertiesToSet As Scripting.Dictionary, ByRef initialPropertiesToSet As Object) As Boolean
    ' SPECIAL CASE: Respect path override for plate parts
    If propertiesToSet.exists("Reference Category") Then
        If propertiesToSet("Reference Category") = "PL" Or propertiesToSet("Reference Category") = "CP" Then
            Dim pathOverride As String
            pathOverride = GetPathBasedUnitOverride(model)

            If pathOverride = "" Then
                ' No path override - apply standard inch forcing
                DebugLog "PLATE PART DETECTED: Forcing UOM to inches (no path override)"
                propertiesToSet("Dimensional UOM") = "in"
            Else
                DebugLog "PLATE PART DETECTED: Path override active - keeping UOM: " & pathOverride
            End If
        End If
    End If

    Dim swCustPropMgr As customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")
    Dim existingProps As Variant
    Dim changesDetected As Boolean
    Dim isPurchaseType As Boolean
    changesDetected = False
    isPurchaseType = False

    ' Get existing properties
    existingProps = swCustPropMgr.GetNames

    ' Get the appropriate property rules based on model type
    Dim PropertyRules As Object
    If model.GetType = swDocPART Then
        Set PropertyRules = GetRequiredPartPropertyRules()
    ElseIf model.GetType = swDocASSEMBLY Then
        Set PropertyRules = GetRequiredAssemblyPropertyRules()
    End If

    If propertiesToSet.exists("Type") Then
        isPurchaseType = (UCase(Trim(CStr(propertiesToSet("Type")))) = "P")
    End If

    ' Build dictionary of existing values for comparison
    Dim existingValues As New Scripting.Dictionary
    Dim i As Long
    If Not IsEmpty(existingProps) Then
        Dim propVal As String
        Dim resolvedVal As String
        Dim wasResolved As Boolean
        Dim linkToProp As Boolean

        For i = 0 To UBound(existingProps)
            swCustPropMgr.Get6 CStr(existingProps(i)), False, propVal, resolvedVal, wasResolved, linkToProp
            existingValues.Add CStr(existingProps(i)), resolvedVal
        Next i
    End If

    ' Check for changes in properties we're managing
    Dim propName As Variant
    For Each propName In propertiesToSet.keys
        Dim newValue As String
        newValue = CStr(propertiesToSet(propName))

        If existingValues.exists(CStr(propName)) Then
            If existingValues(CStr(propName)) <> newValue Then
                changesDetected = True
                Exit For
            End If
        Else
            ' Property doesn't exist yet - need to add it
            changesDetected = True
            Exit For
        End If
    Next propName

    ' FORCE CHANGE DETECTION FOR PLATE PARTS
    If propertiesToSet.exists("Reference Category") Then
        If propertiesToSet("Reference Category") = "PL" Or propertiesToSet("Reference Category") = "CP" Then
            changesDetected = True
        End If
    End If

    ' FORCE CHANGE DETECTION FOR TYPE P CLEANUP
    If isPurchaseType And Not IsEmpty(existingProps) Then
        For i = 0 To UBound(existingProps)
            Dim existingPropUpper As String
            existingPropUpper = UCase(Trim(CStr(existingProps(i))))
            If existingPropUpper = "MTL PART NUMBER" Or existingPropUpper = "MTL UNIT QTY" Then
                changesDetected = True
                Exit For
            End If
        Next i
    End If

    ' If changes detected, update properties
    If changesDetected Then
        ' Identify which existing properties are USER properties (not managed by us)
        Dim userProps As New Scripting.Dictionary
        If Not IsEmpty(existingProps) Then
            For i = 0 To UBound(existingProps)
                Dim existingPropName As String
                existingPropName = CStr(existingProps(i))
                Dim existingPropNameUpper As String
                existingPropNameUpper = UCase(Trim(existingPropName))

                ' If this property is NOT in propertiesToSet and NOT in propertyRules,
                ' it's a user-added property that we should preserve
                If isPurchaseType And (existingPropNameUpper = "MTL PART NUMBER" Or existingPropNameUpper = "MTL UNIT QTY") Then
                    DebugLog "Type P cleanup: not preserving " & existingPropName
                ElseIf Not propertiesToSet.exists(existingPropName) Then
                    If Not PropertyRules.exists(existingPropName) Then
                        userProps.Add existingPropName, existingValues(existingPropName)
                        DebugLog "Preserving user property: " & existingPropName
                    End If
                End If
            Next i
        End If

        ' Build ordered list of properties to set
        Dim orderedProps As New collection
        Dim unorderedProps As New collection

        ' Separate properties into ordered (in rules) and unordered
        For Each propName In propertiesToSet.keys
            If PropertyRules.exists(propName) Then
                orderedProps.Add Array(propName, propertiesToSet(propName))
            Else
                unorderedProps.Add Array(propName, propertiesToSet(propName))
            End If
        Next propName

        ' Delete only the properties we're managing (not user properties)
        If Not IsEmpty(existingProps) Then
            For i = 0 To UBound(existingProps)
                Dim propToCheck As String
                propToCheck = CStr(existingProps(i))
                Dim propToCheckUpper As String
                propToCheckUpper = UCase(Trim(propToCheck))
                Dim forceDeleteForTypeP As Boolean
                forceDeleteForTypeP = (isPurchaseType And (propToCheckUpper = "MTL PART NUMBER" Or propToCheckUpper = "MTL UNIT QTY"))

                ' Only delete if it's a property we're managing OR it's in the rules
                If propertiesToSet.exists(propToCheck) Or PropertyRules.exists(propToCheck) Or forceDeleteForTypeP Then
                    swCustPropMgr.Delete2 propToCheck
                    If forceDeleteForTypeP Then
                        DebugLog "Type P cleanup: deleted " & propToCheck
                    End If
                End If
            Next i
        End If

        ' Add ordered properties first (following the rules order)
        For Each propName In PropertyRules.keys
            For i = 1 To orderedProps.Count
                If orderedProps.item(i)(0) = propName Then
                    swCustPropMgr.Add3 CStr(propName), swCustomInfoText, CStr(orderedProps.item(i)(1)), swCustomPropertyDeleteAndAdd
                    Exit For
                End If
            Next i
        Next propName

        ' Add remaining unordered properties from propertiesToSet
        For i = 1 To unorderedProps.Count
            Dim prop As Variant
            prop = unorderedProps.item(i)
            swCustPropMgr.Add3 CStr(prop(0)), swCustomInfoText, CStr(prop(1)), swCustomPropertyDeleteAndAdd
        Next i

        ' Re-add preserved user properties at the end
        Dim userPropName As Variant
        For Each userPropName In userProps.keys
            swCustPropMgr.Add3 CStr(userPropName), swCustomInfoText, CStr(userProps(userPropName)), swCustomPropertyDeleteAndAdd
            DebugLog "Restored user property: " & userPropName
        Next userPropName

        ' ADDITIONAL CHECK: Ensure Dimensional UOM respects path override for plates
        If propertiesToSet.exists("Reference Category") Then
            If propertiesToSet("Reference Category") = "PL" Or propertiesToSet("Reference Category") = "CP" Then
                Dim pathOverrideFinal As String
                pathOverrideFinal = GetPathBasedUnitOverride(model)

                If pathOverrideFinal = "" Then
                    swCustPropMgr.Set2 "Dimensional UOM", "in"
                    DebugLog "FINAL VERIFICATION: Direct API set of Dimensional UOM to 'in'"
                End If
            End If
        End If
    End If
    
    ApplyCollectedProperties = changesDetected
End Function
    
    ' Validate a property against a rule
    Function ValidateProperty(propertiesToSet As Object, propertyName As String, propertyRule As String) As Boolean
        Dim evaluator As Object
        Set evaluator = CreateObject("VBScript.RegExp")
        evaluator.pattern = propertyRule
    
        Dim propValue As String
        propValue = propertiesToSet(propertyName)
    
        ' Check if the property value matches the rule
        ValidateProperty = evaluator.Test(propValue)
    End Function
    
    ' Resolve lightweight components in an assembly
    Sub ResolveLightweightComponents(model As ModelDoc2)
        If model.GetType = swDocASSEMBLY Then
            Dim assembly As SldWorks.AssemblyDoc
            Set assembly = model
            assembly.ResolveAllLightWeightComponents False
        End If
    End Sub
    
    
Sub ShowUnifiedResultsForm(failedProperties As Scripting.Dictionary, failedLocations As Scripting.Dictionary, failedPieceParts As Scripting.Dictionary)
    LogToFile "=== Starting ShowUnifiedResultsForm ==="
    
    ' Log input parameters
    LogToFile "Input parameters:"
    LogToFile "- failedProperties.Count: " & failedProperties.Count
    LogToFile "- failedLocations.Count: " & failedLocations.Count
    LogToFile "- failedPieceParts.Count: " & failedPieceParts.Count
    
    ' Create the form
    Dim frm As UnifiedResultsForm
    Set frm = New UnifiedResultsForm
    
    ' Calculate unique models
    Dim uniqueModels As Scripting.Dictionary
    Set uniqueModels = New Scripting.Dictionary
    
    Dim key As Variant
    Dim modelName As String
    
    ' Process all dictionaries to build uniqueModels
    For Each key In failedProperties.keys
        Dim pathParts As Variant
        pathParts = Split(CStr(key), " - ")
        
        If UBound(pathParts) >= 1 Then
            modelName = ExtractFilenameFromPath(CStr(pathParts(0)))
        Else
            modelName = CStr(key)
        End If
        
        If InStr(modelName, ".") > 0 Then
            modelName = Left(modelName, InStrRev(modelName, ".") - 1)
        End If
        
        uniqueModels(modelName) = True
    Next key
    
    For Each key In failedLocations.keys
        modelName = ExtractFilenameFromPath(CStr(key))
        If InStr(modelName, ".") > 0 Then
            modelName = Left(modelName, InStrRev(modelName, ".") - 1)
        End If
        uniqueModels(modelName) = True
    Next key
    
    For Each key In failedPieceParts.keys
        If InStr(CStr(key), "\") > 0 Then
            modelName = ExtractFilenameFromPath(CStr(key))
        Else
            modelName = CStr(key)
        End If
        
        If InStr(modelName, ".") > 0 Then
            modelName = Left(modelName, InStrRev(modelName, ".") - 1)
        End If
        
        uniqueModels(modelName) = True
    Next key
    
    ' Set summary text and populate form
    frm.SetSummaryText uniqueModels.Count
    
    ' IMPORTANT: Call PopulateData with only 3 parameters
    frm.PopulateData failedProperties, failedLocations, failedPieceParts

    ' Show the form - MUST be vbModal to prevent it from disappearing
    ' vbModal blocks until the user closes the form, keeping it visible
    frm.Show vbModal

    LogToFile "=== ShowUnifiedResultsForm completed ==="
End Sub
    
    
    
    
 Function ExtractFilenameFromPath(path As String) As String
    On Error Resume Next
    LogToFile "Extracting filename from: " & path
    
    ' Check if this is a checkout error message
    If InStr(path, "MODEL NOT CHECKED OUT") > 0 Or InStr(path, "MODEL CHECKED OUT BY") > 0 Then
        ' Extract just the model name from the checkout error message
        LogToFile "This appears to be a checkout message"
        Dim parts As Variant
        parts = Split(path, ":")
        If UBound(parts) >= 1 Then
            ' Get the part after the colon and trim it
            ExtractFilenameFromPath = Trim(parts(UBound(parts)))
            LogToFile "Extracted from checkout message: " & ExtractFilenameFromPath
        Else
            ExtractFilenameFromPath = path
            LogToFile "Could not parse checkout message, returning full path"
        End If
    Else
        ' Normal path processing
        Dim fileParts As Variant
        fileParts = Split(path, "\")
        If UBound(fileParts) >= 0 Then
            ExtractFilenameFromPath = fileParts(UBound(fileParts))
            LogToFile "Extracted from path: " & ExtractFilenameFromPath
        Else
            ExtractFilenameFromPath = path
            LogToFile "Could not split path, returning full path"
        End If
    End If
    On Error GoTo 0
End Function
    
    
    
    
    Function GetModelFromDrawing(drawingDoc As drawingDoc) As ModelDoc2
        Dim swSheet As sheet
        Dim vViews As Variant
        Dim swView As view
        Dim i As Integer
        
        Set swSheet = drawingDoc.GetCurrentSheet
        vViews = swSheet.GetViews
        
        If Not IsEmpty(vViews) Then
            For i = 0 To UBound(vViews)
                Set swView = vViews(i)
                If Not swView Is Nothing Then
                    Set GetModelFromDrawing = swView.ReferencedDocument
                    If Not GetModelFromDrawing Is Nothing Then
                        Exit Function
                    End If
                End If
            Next i
        End If
        
        Set GetModelFromDrawing = Nothing
    End Function



Function GetReferencedModelFromDrawing(drawingDoc As ModelDoc2) As ModelDoc2
    On Error GoTo ErrorHandler
    
    LogToFile "=== Getting Referenced Model from Drawing ==="
    
    If drawingDoc Is Nothing Then
        LogToFile "Error: Drawing document is Nothing"
        Set GetReferencedModelFromDrawing = Nothing
        Exit Function
    End If
    
    If drawingDoc.GetType <> swDocDRAWING Then
        LogToFile "Error: Document is not a drawing"
        Set GetReferencedModelFromDrawing = Nothing
        Exit Function
    End If
    
    Dim swDraw As drawingDoc
    Set swDraw = drawingDoc
    
    ' Try multiple approaches to get referenced model
    
    ' Approach 1: Get first model view
    Dim sheetView As view
    Set sheetView = swDraw.GetFirstView ' This is the sheet view
    
    If Not sheetView Is Nothing Then
        Dim modelView As view
        Set modelView = sheetView.GetNextView ' This should be the first model view
        
        While Not modelView Is Nothing
            Dim refDoc As ModelDoc2
            Set refDoc = modelView.ReferencedDocument
            
            If Not refDoc Is Nothing Then
                LogToFile "Found referenced model: " & refDoc.GetTitle
                LogToFile "Model type: " & refDoc.GetType
                Set GetReferencedModelFromDrawing = refDoc
                Exit Function
            End If
            
            Set modelView = modelView.GetNextView
        Wend
    End If
    
    ' Approach 2: Try getting from drawing documents collection
    Dim docCount As Long
    docCount = swDraw.GetDocumentCount
    LogToFile "Document count in drawing: " & docCount
    
    If docCount > 0 Then
        Dim i As Long
        For i = 0 To docCount - 1
            Dim doc As ModelDoc2
            Set doc = swDraw.GetDocument(i)
            If Not doc Is Nothing Then
                If doc.GetType = swDocPART Or doc.GetType = swDocASSEMBLY Then
                    LogToFile "Found referenced document via GetDocument: " & doc.GetTitle
                    Set GetReferencedModelFromDrawing = doc
                    Exit Function
                End If
            End If
        Next i
    End If
    
    LogToFile "No referenced model found in drawing"
    Set GetReferencedModelFromDrawing = Nothing
    Exit Function
    
ErrorHandler:
    LogToFile "Error in GetReferencedModelFromDrawing: " & Err.description
    Set GetReferencedModelFromDrawing = Nothing
End Function

Sub RebuildAllDrawingViews(drawingDoc As ModelDoc2)
    On Error GoTo ErrorHandler
    
    LogToFile "=== REBUILDING ALL DRAWING VIEWS ==="
    
    If drawingDoc.GetType <> swDocDRAWING Then
        LogToFile "Document is not a drawing, skipping view rebuild"
        Exit Sub
    End If
    
    Dim swDraw As drawingDoc
    Set swDraw = drawingDoc
    
    ' Get all sheets and rebuild views on each sheet
    Dim sheetNames As Variant
    sheetNames = swDraw.GetSheetNames
    
    If Not IsEmpty(sheetNames) Then
        Dim currentSheetName As String
        currentSheetName = swDraw.GetCurrentSheet.GetName
        
        Dim sheetName As Variant
        For Each sheetName In sheetNames
            LogToFile "Processing sheet: " & sheetName
            
            ' Activate the sheet
            Dim bRet As Boolean
            bRet = swDraw.ActivateSheet(CStr(sheetName))
            
            If bRet Then
                ' Get the sheet
                Dim sheet As sheet
                Set sheet = swDraw.GetCurrentSheet
                
                ' Get all views on the sheet
                Dim views As Variant
                views = sheet.GetViews
                
                If Not IsEmpty(views) Then
                    Dim view As view
                    Dim i As Long
                    For i = 0 To UBound(views)
                        Set view = views(i)
                        If Not view Is Nothing Then
                            LogToFile "Rebuilding view: " & view.Name
                            view.Rebuild
                        End If
                    Next i
                End If
            Else
                LogToFile "Failed to activate sheet: " & sheetName
            End If
        Next sheetName
        
        ' Restore original sheet
        swDraw.ActivateSheet currentSheetName
        LogToFile "Restored original sheet: " & currentSheetName
    End If
    
    LogToFile "=== ALL DRAWING VIEWS REBUILT ==="
    Exit Sub
    
ErrorHandler:
    LogToFile "Error in RebuildAllDrawingViews: " & Err.description
End Sub




