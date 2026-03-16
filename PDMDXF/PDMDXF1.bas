Attribute VB_Name = "PDMDXF1"
Option Explicit

' Declare Windows API Sleep function with PtrSafe for 64-bit compatibility
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Global variables
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim pdmVault As EdmVault5 ' PDM Vault object
Dim filePath As String
Dim fileName As String
Dim RevValue As String
Dim bManualMode As Boolean
Dim manualOutputPath As String

' =============================
' USER CONFIGURATION (TOP-LEVEL)
' =============================
' Set to True only if you want the Yes/No custom-folder prompt to appear.
Private Const ENABLE_CUSTOM_DXF_PROMPT As Boolean = False

' Set to True to force DXF output to manual path below (no prompt required).
Private Const DEFAULT_MANUAL_MODE As Boolean = False

' Used only when DEFAULT_MANUAL_MODE = True.
Private Const DEFAULT_MANUAL_OUTPUT_PATH As String = "C:\DXF Exports\"

' Set True to allow running from assembly selections and multi-body parts.
' Assembly behavior: select one face on a component; macro exports that component part.
Private Const ENABLE_ASSEMBLY_OR_MULTIBODY_MODE As Boolean = True

' If assembly face transfer fails, wait this long for manual face selection in opened part.
Private Const MANUAL_FACE_SELECTION_TIMEOUT_SEC As Double = 30#

' Bump this when troubleshooting to verify SolidWorks is running the expected module.
Private Const DXF_MACRO_BUILD As String = "2026-02-09-06"

' Define surface type constants from SolidWorks API
Const swSurfType_PLANE As Long = 4001
Const swSurfType_CYLINDER As Long = 4002
Const swSurfType_CONE As Long = 4003
Const swSurfType_SPHERE As Long = 4004
Const swSurfType_TORUS As Long = 4005
Const swSurfType_BSPLINE As Long = 4006

' Selection type constants
Const swSelFACES As Long = 2 ' Correct value for faces in SolidWorks API

' Main subroutine to execute the macro
Sub main()
    Debug.Print "Starting DXF macro execution."
    Debug.Print "DXF Macro Build: " & DXF_MACRO_BUILD
    
    ' Initialize settings
    bManualMode = DEFAULT_MANUAL_MODE
    manualOutputPath = DEFAULT_MANUAL_OUTPUT_PATH
    
    If bManualMode Then
        If Trim$(manualOutputPath) = "" Then
            MsgBox "Manual mode is enabled, but DEFAULT_MANUAL_OUTPUT_PATH is blank.", vbCritical
            Exit Sub
        End If
        
        ' Normalize path for path concatenation logic.
        If Right$(manualOutputPath, 1) <> "\" Then manualOutputPath = manualOutputPath & "\"
    Else
        manualOutputPath = ""
    End If
    
    ' Prompt only when explicitly enabled in top-level config.
    If ENABLE_CUSTOM_DXF_PROMPT Then PromptForCustomDXFOutputFolder

    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Debug.Print "Active document: " & IIf(swModel Is Nothing, "None", swModel.GetPathName)

    ' Check if no document is open
    If swModel Is Nothing Then
        MsgBox "No document open.", vbCritical
        Debug.Print "Error: No document open."
        Exit Sub
    End If

    Dim activeDocType As Long
    activeDocType = swModel.GetType

    If activeDocType <> swDocPART And activeDocType <> swDocASSEMBLY Then
        MsgBox "Active document must be a PART or ASSEMBLY.", vbCritical
        Debug.Print "Error: Unsupported active document type: " & activeDocType
        Exit Sub
    End If
    Debug.Print "Active document type: " & activeDocType

    ' Check if the model is saved
    If swModel.GetPathName = "" Then
        MsgBox "Please save the model first.", vbCritical
        Debug.Print "Error: Model not saved. Path: " & swModel.GetPathName
        Exit Sub
    End If
    Debug.Print "Model is saved."

    ' Get selected face
    Dim selMgr As SldWorks.SelectionMgr
    Set selMgr = swModel.SelectionManager
    Debug.Print "Selection count: " & selMgr.GetSelectedObjectCount2(-1)
    If selMgr.GetSelectedObjectCount2(-1) <> 1 Then
        MsgBox "Please select exactly one face.", vbCritical
        Debug.Print "Error: Selection count is not 1. Count: " & selMgr.GetSelectedObjectCount2(-1)
        Exit Sub
    End If
    Dim selType As Long
    selType = selMgr.GetSelectedObjectType3(1, -1)
    Debug.Print "Selected object type ID: " & selType & " (swSelFACES = 2)"
    If selType <> swSelFACES Then
        MsgBox "Selected object is not a face.", vbCritical
        Debug.Print "Error: Selected object is not a face. Type: " & selType
        Exit Sub
    End If
    Dim swFace As SldWorks.Face2
    Set swFace = selMgr.GetSelectedObject6(1, -1)
    Debug.Print "Face object retrieved: " & IIf(swFace Is Nothing, "No", "Yes")
    If swFace Is Nothing Then
        MsgBox "Failed to retrieve face object.", vbCritical
        Debug.Print "Error: Face object is Nothing."
        Exit Sub
    End If

    If Not IsPlanarFace(swFace) Then
        MsgBox "Selected face must be a planar face for DXF export.", vbCritical
        Debug.Print "Error: Selected face is not planar."
        Exit Sub
    End If

    Dim selectedFaceRef As Variant
    selectedFaceRef = Empty
    On Error Resume Next
    selectedFaceRef = swModel.Extension.GetPersistReference3(swFace)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Dim selectedFaceId As Long
    selectedFaceId = GetFaceIdSafe(swFace)
    Debug.Print "Selected face ID: " & selectedFaceId

    ' Resolve export source model.
    If activeDocType = swDocASSEMBLY Then
        If Not ENABLE_ASSEMBLY_OR_MULTIBODY_MODE Then
            MsgBox "Active document is an assembly and assembly mode is disabled in config." & vbCrLf & vbCrLf & _
                   "Set ENABLE_ASSEMBLY_OR_MULTIBODY_MODE = True to allow this.", vbCritical
            Debug.Print "Error: Assembly mode disabled by config."
            Exit Sub
        End If

        Dim swComp As Object
        Set swComp = GetSelectedComponentFromFaceSelection(selMgr)
        If swComp Is Nothing Then
            MsgBox "Could not resolve the selected face's component in the assembly.", vbCritical
            Debug.Print "Error: Failed to resolve selected component from assembly face."
            Exit Sub
        End If

        Dim sourcePart As SldWorks.ModelDoc2
        Set sourcePart = swComp.GetModelDoc2
        If sourcePart Is Nothing Then
            MsgBox "Could not load the selected component model from the assembly.", vbCritical
            Debug.Print "Error: Selected component model is Nothing."
            Exit Sub
        End If

        If sourcePart.GetType <> swDocPART Then
            MsgBox "Selected face must belong to a PART component.", vbCritical
            Debug.Print "Error: Selected component is not a part. Type: " & sourcePart.GetType
            Exit Sub
        End If

        Dim sourcePartPath As String
        sourcePartPath = sourcePart.GetPathName
        If sourcePartPath = "" Then
            MsgBox "Selected component part is not saved. Save it and try again.", vbCritical
            Debug.Print "Error: Selected component part path is blank."
            Exit Sub
        End If

        Dim openErrors As Long
        Dim openWarnings As Long
        Set swModel = swApp.OpenDoc6(sourcePartPath, swDocPART, swOpenDocOptions_Silent, "", openErrors, openWarnings)
        If swModel Is Nothing Then
            MsgBox "Failed to open selected component part for export:" & vbCrLf & sourcePartPath, vbCritical
            Debug.Print "Error: OpenDoc6 failed for source part. Errors=" & openErrors & ", Warnings=" & openWarnings
            Exit Sub
        End If

        Dim activateErrors As Long
        Set swModel = swApp.ActivateDoc3(swModel.GetTitle, True, 0, activateErrors)
        If swModel Is Nothing Then
            MsgBox "Could not activate selected component part for export.", vbCritical
            Debug.Print "Error: ActivateDoc3 failed for source part. Errors=" & activateErrors
            Exit Sub
        End If

        swModel.ClearSelection2 True
        Dim faceToSelect As Object
        Set faceToSelect = Nothing

        If Not IsEmpty(selectedFaceRef) Then
            Dim persistResolveError As Long
            On Error Resume Next
            Set faceToSelect = swModel.Extension.GetObjectByPersistReference3(selectedFaceRef, persistResolveError)
            On Error GoTo 0
            Debug.Print "Persist ref resolve error: " & persistResolveError & ", face resolved: " & IIf(faceToSelect Is Nothing, "No", "Yes")
        End If

        If faceToSelect Is Nothing Then
            On Error Resume Next
            Set faceToSelect = CallByName(swComp, "GetCorrespondingEntity", VbMethod, swFace)
            If faceToSelect Is Nothing Then Set faceToSelect = CallByName(swComp, "GetCorresponding", VbMethod, swFace)
            On Error GoTo 0
            Debug.Print "Component correspondence resolved face: " & IIf(faceToSelect Is Nothing, "No", "Yes")
        End If

        If faceToSelect Is Nothing And selectedFaceId <> 0 Then
            Set faceToSelect = FindFaceById(swModel, selectedFaceId)
            Debug.Print "FaceId lookup resolved face: " & IIf(faceToSelect Is Nothing, "No", "Yes")
        End If

        If faceToSelect Is Nothing Then Set faceToSelect = swFace

        Dim faceSelectionTransferred As Boolean
        faceSelectionTransferred = False
        Debug.Print "Face candidate type for selection: " & TypeName(faceToSelect)
        On Error Resume Next
        faceSelectionTransferred = CallByName(faceToSelect, "Select4", VbMethod, False, Nothing)
        If Not faceSelectionTransferred Then
            faceSelectionTransferred = CallByName(faceToSelect, "Select2", VbMethod, False, -1)
        End If
        On Error GoTo 0

        Dim partSelMgr As SldWorks.SelectionMgr
        Set partSelMgr = swModel.SelectionManager
        Debug.Print "Part selection count after transfer: " & partSelMgr.GetSelectedObjectCount2(-1)

        If Not faceSelectionTransferred Or partSelMgr.GetSelectedObjectCount2(-1) <> 1 Then
            Debug.Print "Auto-transfer failed. Attempting to find matching planar face automatically."
            swModel.ClearSelection2 True
            Set swFace = Nothing

            ' Try to automatically find and select a planar face in the opened part.
            ' This avoids requiring user interaction (MsgBox blocks model access).
            Dim autoFace As SldWorks.Face2
            Set autoFace = FindFirstPlanarFace(swModel)

            If Not autoFace Is Nothing Then
                Dim autoSelectOk As Boolean
                On Error Resume Next
                autoSelectOk = autoFace.Select4(False, Nothing)
                If Not autoSelectOk Then autoSelectOk = autoFace.Select2(False, -1)
                On Error GoTo 0

                If autoSelectOk Then
                    Debug.Print "Auto-selected a planar face in the opened part."
                Else
                    Debug.Print "Error: Found a planar face but could not select it."
                    MsgBox "Found a planar face in the part but failed to select it." & vbCrLf & _
                           "Open the part directly, select a face, and run the macro again.", vbCritical
                    Exit Sub
                End If
            Else
                Debug.Print "Error: No planar faces found in the part."
                MsgBox "No planar faces found in the opened component part." & vbCrLf & _
                       "Open the part directly, select a face, and run the macro again.", vbCritical
                Exit Sub
            End If
        End If

        Debug.Print "Assembly mode: exporting selected component part: " & sourcePartPath
    End If

    If Not ENABLE_ASSEMBLY_OR_MULTIBODY_MODE Then
        Dim solidBodyCount As Long
        solidBodyCount = GetSolidBodyCount(swModel)
        If solidBodyCount > 1 Then
            MsgBox "Part has multiple solid bodies and multibody mode is disabled in config." & vbCrLf & vbCrLf & _
                   "Set ENABLE_ASSEMBLY_OR_MULTIBODY_MODE = True to allow this.", vbCritical
            Debug.Print "Error: Multi-body part blocked by config. Body count: " & solidBodyCount
            Exit Sub
        End If
    End If

    ' Get revision from custom properties
    Dim cusPropMgr As SldWorks.CustomPropertyManager
    Set cusPropMgr = swModel.Extension.CustomPropertyManager("")
    Dim wasResolved As Boolean
    cusPropMgr.Get5 "Revision", False, "", RevValue, wasResolved
    Debug.Print "Revision retrieved: " & RevValue & ", Resolved: " & wasResolved
    If Not wasResolved Or Trim(RevValue) = "" Then
        MsgBox "Revision is not set or blank.", vbCritical
        Debug.Print "Error: Revision is not set or blank."
        Exit Sub
    End If

   ' Determine file paths
Dim fullPath As String
fullPath = swModel.GetPathName
fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
fileName = Left(fileName, InStrRev(fileName, ".") - 1)
filePath = Left(fullPath, InStrRev(fullPath, "\"))
Debug.Print "File Name: " & fileName
Debug.Print "File Path: " & filePath

Dim burnProfilesFolder As String
Dim obsoleteFolder As String
Dim isLibraryPath As Boolean
isLibraryPath = (InStr(1, fullPath, "C:\NMT_PDM\Libraries", vbTextCompare) > 0)

If bManualMode Then
    burnProfilesFolder = manualOutputPath
    obsoleteFolder = burnProfilesFolder & "Obsolete\"
    Debug.Print "Manual mode active. Using folder: " & burnProfilesFolder
Else
    ' Check if in Libraries path
    Debug.Print "Is Library Path: " & isLibraryPath

    If isLibraryPath Then
        ' Create a Drawings\Burn Profiles subfolder in the library path
        burnProfilesFolder = filePath & "Drawings\Burn Profiles\"
        obsoleteFolder = burnProfilesFolder & "Obsolete\"
    Else
        ' Original logic for project files
        Dim designFolder As String
        Dim pos As Long
        pos = InStrRev(filePath, "3 - Design\")
        If pos > 0 Then
            designFolder = Left(filePath, pos + Len("3 - Design") - 1) & "\"
        Else
            MsgBox "Could not find '3 - Design' in the model path.", vbCritical
            Debug.Print "Error: Could not find '3 - Design' in model path: " & filePath
            Exit Sub
        End If
        Debug.Print "Design Folder: " & designFolder

        burnProfilesFolder = designFolder & "Drawings\Burn Profiles\"
        obsoleteFolder = burnProfilesFolder & "Obsolete\"
    End If
End If

' Connect to PDM vault
If Not ConnectToPDMVault() Then
    MsgBox "Failed to connect to NMT_PDM vault.", vbCritical
    Debug.Print "Error: Failed to connect to NMT_PDM vault."
    Exit Sub
End If
Debug.Print "Connected to PDM vault successfully."

' Ensure folders exist
If bManualMode Then
    ' Create local folders if they don't exist
    If Dir(burnProfilesFolder, vbDirectory) = "" Then MkDir burnProfilesFolder
    If Dir(obsoleteFolder, vbDirectory) = "" Then MkDir obsoleteFolder
Else
    If isLibraryPath Then
        EnsureFolderExists filePath, "Drawings", filePath & "Drawings\"
        EnsureFolderExists filePath & "Drawings\", "Burn Profiles", burnProfilesFolder
    Else
        EnsureFolderExists designFolder & "Drawings\", "Burn Profiles", burnProfilesFolder
    End If
    EnsureFolderExists burnProfilesFolder, "Obsolete", obsoleteFolder
End If

' PDM specific logic - only if path is inside vault
Dim isInsideVault As Boolean
isInsideVault = (InStr(1, burnProfilesFolder, pdmVault.RootFolderPath, vbTextCompare) > 0)

If isInsideVault Then
    ' Get PDM folder objects
    Dim pdmBurnFolder As Object ' IEdmFolder5
    Set pdmBurnFolder = pdmVault.GetFolderFromPath(burnProfilesFolder)
    If pdmBurnFolder Is Nothing Then
        MsgBox "Could not find Burn Profiles folder in vault.", vbCritical
        Debug.Print "Error: Could not find Burn Profiles folder in vault: " & burnProfilesFolder
        Exit Sub
    End If

    Dim pdmObsoleteFolder As Object ' IEdmFolder5
    Set pdmObsoleteFolder = pdmVault.GetFolderFromPath(obsoleteFolder)
    If pdmObsoleteFolder Is Nothing Then
        MsgBox "Could not find Obsolete folder in vault.", vbCritical
        Debug.Print "Error: Could not find Obsolete folder in vault: " & obsoleteFolder
        Exit Sub
    End If
End If

    ' Construct target DXF filename without version suffix
    Dim targetDXFName As String
    targetDXFName = fileName & "_Rev" & RevValue & ".dxf"
    Dim targetDXFPath As String
    targetDXFPath = burnProfilesFolder & targetDXFName
    Debug.Print "Target DXF Path: " & targetDXFPath

    ' Step 1: Obsolete old revisions
    If isInsideVault Then
        ' Existing PDM iteration logic...
        Debug.Print "Starting to list all files in Burn Profiles folder (PDM)."
        Dim posFile As IEdmPos5
        Set posFile = pdmBurnFolder.GetFirstFilePosition
        
        If Not posFile Is Nothing Then
            While Not posFile.IsNull
                Dim pdmFileIter As IEdmFile5
                Set pdmFileIter = pdmBurnFolder.GetNextFile(posFile)
                If Not pdmFileIter Is Nothing Then
                    If pdmFileIter.Name Like fileName & "_Rev*.dxf" Then
                        Dim existingRev As String
                        existingRev = ExtractRevision(pdmFileIter.Name, fileName)
                        If existingRev <> "" And IsNumeric(existingRev) And IsNumeric(RevValue) Then
                            If CInt(existingRev) < CInt(RevValue) Then
                                ' Move to Obsolete logic here
                                MoveFileToObsoletePDM pdmFileIter, pdmBurnFolder, pdmObsoleteFolder
                            End If
                        End If
                    End If
                End If
            Wend
        End If
    Else
        ' Manual mode / Outside vault - simple file move
        Debug.Print "Manual mode: Checking for older revisions locally."
        Dim localFile As String
        localFile = Dir(burnProfilesFolder & fileName & "_Rev*.dxf")
        Do While localFile <> ""
            Dim localRev As String
            localRev = ExtractRevision(localFile, fileName)
            If localRev <> "" And IsNumeric(localRev) And IsNumeric(RevValue) Then
                If CInt(localRev) < CInt(RevValue) Then
                    Debug.Print "Moving local file to obsolete: " & localFile
                    On Error Resume Next
                    Name burnProfilesFolder & localFile As obsoleteFolder & localFile
                    On Error GoTo 0
                End If
            End If
            localFile = Dir()
        Loop
    End If

    ' Step 2: Handle the target DXF (PDM Locking)
    Dim pdmFile As Object ' IEdmFile5
    If isInsideVault Then
        Set pdmFile = pdmVault.GetFileFromPath(targetDXFPath)
        If Not pdmFile Is Nothing Then
            If Not pdmFile.IsLocked Then
                pdmFile.LockFile pdmBurnFolder.ID, 0
            End If
        End If
    End If

    ' Export the new DXF
    Debug.Print "Exporting DXF to: " & targetDXFPath
    Dim dataAlignment(11) As Double
    dataAlignment(0) = 0: dataAlignment(1) = 0: dataAlignment(2) = 0
    dataAlignment(3) = 1: dataAlignment(4) = 0: dataAlignment(5) = 0
    dataAlignment(6) = 0: dataAlignment(7) = 1: dataAlignment(8) = 0
    dataAlignment(9) = 0: dataAlignment(10) = 0: dataAlignment(11) = 0
    Dim varAlignment As Variant
    varAlignment = dataAlignment
    
    Dim fileExistedBefore As Boolean
    fileExistedBefore = (Dir(targetDXFPath) <> "")

    On Error Resume Next
    Dim swPart As SldWorks.PartDoc
    Dim exportOk As Boolean
    Set swPart = swModel
    exportOk = swPart.ExportToDWG2(targetDXFPath, swModel.GetPathName, swExportToDWG_ExportSelectedFacesOrLoops, True, varAlignment, False, False, 0, Null)
    If Err.Number <> 0 Then
        Debug.Print "Error exporting DXF: " & Err.Description
        MsgBox "Failed to export DXF to " & targetDXFPath & ": " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Debug.Print "ExportToDWG2 returned: " & exportOk
    If Dir(targetDXFPath) = "" Then
        MsgBox "DXF export reported success, but output file was not found." & vbCrLf & vbCrLf & _
               "Expected: " & targetDXFPath, vbCritical
        Debug.Print "Export verification failed: file not found after export."
        Exit Sub
    End If

    If Not exportOk Then
        Debug.Print "Warning: ExportToDWG2 returned False but output file exists. Continuing."
    End If

    If fileExistedBefore Then
        Debug.Print "Output file existed before export and still exists: " & targetDXFPath
    Else
        Debug.Print "Output file created: " & targetDXFPath
    End If

    ' Check in or add the file (PDM)
    If isInsideVault Then
        If Not pdmFile Is Nothing Then
            If pdmFile.IsLocked Then pdmFile.UnlockFile pdmBurnFolder.ID, "Updated by macro", 0
        Else
            pdmBurnFolder.AddFile 0, targetDXFPath
            Set pdmFile = pdmVault.GetFileFromPath(targetDXFPath)
            If Not pdmFile Is Nothing Then
                If pdmFile.IsLocked Then
                    pdmFile.UnlockFile pdmBurnFolder.ID, "Added by macro", 0
                End If
            End If
        End If
    End If

    MsgBox "DXF '" & targetDXFName & "' created and processed successfully.", vbInformation
    Debug.Print "DXF macro completed successfully."
End Sub

' Get the selected component for the selected face (assembly context).
' Uses CallByName for compatibility with different SolidWorks API versions.
Private Function GetSelectedComponentFromFaceSelection(selMgr As SldWorks.SelectionMgr) As Object
    Dim swComp As Object

    On Error Resume Next
    Set swComp = CallByName(selMgr, "GetSelectedObjectsComponent4", VbMethod, 1, -1)
    If swComp Is Nothing Then Set swComp = CallByName(selMgr, "GetSelectedObjectsComponent3", VbMethod, 1, -1)
    If swComp Is Nothing Then Set swComp = CallByName(selMgr, "GetSelectedObjectsComponent2", VbMethod, 1)
    On Error GoTo 0

    Set GetSelectedComponentFromFaceSelection = swComp
End Function

' Validate selected face for DXF face export (must be planar).
Private Function IsPlanarFace(faceObj As SldWorks.Face2) As Boolean
    Dim swSurface As SldWorks.Surface

    On Error Resume Next
    Set swSurface = faceObj.GetSurface
    If swSurface Is Nothing Then
        IsPlanarFace = False
    Else
        IsPlanarFace = (swSurface.Identity = swSurfType_PLANE)
    End If
    On Error GoTo 0
End Function

' Get face ID safely for cross-context lookup.
Private Function GetFaceIdSafe(faceObj As SldWorks.Face2) As Long
    On Error Resume Next
    GetFaceIdSafe = faceObj.GetFaceId
    If Err.Number <> 0 Then
        Err.Clear
        GetFaceIdSafe = 0
    End If
    On Error GoTo 0
End Function

' Find a face in the target part by FaceId.
Private Function FindFaceById(modelDoc As SldWorks.ModelDoc2, targetFaceId As Long) As SldWorks.Face2
    Dim swPart As SldWorks.PartDoc
    Dim vBodies As Variant
    Dim i As Long
    Dim swBody As SldWorks.Body2
    Dim swFace As SldWorks.Face2

    If targetFaceId = 0 Then Exit Function

    Set swPart = modelDoc
    vBodies = swPart.GetBodies2(0, True) ' 0 = solid bodies

    If IsEmpty(vBodies) Then Exit Function

    For i = LBound(vBodies) To UBound(vBodies)
        Set swBody = vBodies(i)
        If Not swBody Is Nothing Then
            Set swFace = swBody.GetFirstFace
            Do While Not swFace Is Nothing
                If swFace.GetFaceId = targetFaceId Then
                    Set FindFaceById = swFace
                    Exit Function
                End If
                Set swFace = swFace.GetNextFace
            Loop
        End If
    Next i
End Function

' Find the first planar face in the part model.
' Used as fallback when assembly-to-part face transfer fails.
Private Function FindFirstPlanarFace(modelDoc As SldWorks.ModelDoc2) As SldWorks.Face2
    Dim swPart As SldWorks.PartDoc
    Dim vBodies As Variant
    Dim i As Long
    Dim swBody As SldWorks.Body2
    Dim swFace As SldWorks.Face2

    Set swPart = modelDoc
    vBodies = swPart.GetBodies2(0, True) ' 0 = solid bodies

    If IsEmpty(vBodies) Then Exit Function

    For i = LBound(vBodies) To UBound(vBodies)
        Set swBody = vBodies(i)
        If Not swBody Is Nothing Then
            Set swFace = swBody.GetFirstFace
            Do While Not swFace Is Nothing
                If IsPlanarFace(swFace) Then
                    Set FindFirstPlanarFace = swFace
                    Exit Function
                End If
                Set swFace = swFace.GetNextFace
            Loop
        End If
    Next i
End Function

' Non-blocking wait for a valid manual face selection in active part.
Private Function WaitForManualPlanarFaceSelection(modelDoc As SldWorks.ModelDoc2, timeoutSeconds As Double, ByRef selectedFace As SldWorks.Face2) As Boolean
    Dim selMgr As SldWorks.SelectionMgr
    Dim startTime As Single
    Dim elapsedSeconds As Double
    Dim currentTime As Single

    startTime = Timer
    WaitForManualPlanarFaceSelection = False
    Set selectedFace = Nothing

    Do
        DoEvents
        Sleep 100

        Set selMgr = modelDoc.SelectionManager
        If Not selMgr Is Nothing Then
            If selMgr.GetSelectedObjectCount2(-1) = 1 Then
                If selMgr.GetSelectedObjectType3(1, -1) = swSelFACES Then
                    Set selectedFace = selMgr.GetSelectedObject6(1, -1)
                    If Not selectedFace Is Nothing Then
                        If IsPlanarFace(selectedFace) Then
                            WaitForManualPlanarFaceSelection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

        currentTime = Timer
        If currentTime >= startTime Then
            elapsedSeconds = currentTime - startTime
        Else
            elapsedSeconds = (86400# - startTime) + currentTime
        End If
    Loop While elapsedSeconds < timeoutSeconds
End Function

' Count solid bodies in the provided part model.
Private Function GetSolidBodyCount(modelDoc As SldWorks.ModelDoc2) As Long
    Dim swPart As SldWorks.PartDoc
    Dim vBodies As Variant

    Set swPart = modelDoc
    vBodies = swPart.GetBodies2(0, True) ' 0 = solid bodies

    If IsEmpty(vBodies) Then
        GetSolidBodyCount = 0
        Exit Function
    End If

    On Error Resume Next
    GetSolidBodyCount = UBound(vBodies) - LBound(vBodies) + 1
    If Err.Number <> 0 Then
        Err.Clear
        GetSolidBodyCount = 1
    End If
    On Error GoTo 0
End Function

' Helper to move file to obsolete in PDM (Extracted from your original main for cleaner logic)
Private Sub MoveFileToObsoletePDM(pdmFile As Object, sourceFolder As Object, targetFolder As Object)
    On Error Resume Next
    If Not pdmFile.IsLocked Then pdmFile.LockFile sourceFolder.ID, 0
    
    Dim localPath As String
    localPath = sourceFolder.localPath & "\" & pdmFile.Name
    
    If Dir(localPath) <> "" Then
        targetFolder.AddFile 0, localPath
        Dim newFile As Object
        Set newFile = targetFolder.GetFile(pdmFile.Name)
        If Not newFile Is Nothing Then
            newFile.UnlockFile targetFolder.ID, "Moved to Obsolete", 0
            ' Change state logic if needed...
        End If
        
        If pdmFile.IsLocked Then pdmFile.UnlockFile sourceFolder.ID, "Checked in before removal", 0
        sourceFolder.DeleteFile sourceFolder.ID, pdmFile.ID, 0
    End If
    On Error GoTo 0
End Sub


' Function to connect to NMT_PDM vault
Private Function ConnectToPDMVault() As Boolean
    On Error Resume Next
    Set pdmVault = New EdmVault5
    pdmVault.LoginAuto "NMT_PDM", 0 ' Auto-login with current Windows credentials
    If pdmVault.IsLoggedIn Then
        ConnectToPDMVault = True
        Debug.Print "Logged into NMT_PDM vault successfully."
    Else
        ConnectToPDMVault = False
        Debug.Print "Failed to log into NMT_PDM vault."
    End If
    On Error GoTo 0
End Function

' Simplified function to ensure a folder exists in vault and locally
Private Function EnsureFolderExists(parentFolderPath As String, folderName As String, folderPath As String) As Boolean
    Dim pdmParentFolder As Object ' IEdmFolder5
    Dim pdmFolder As Object ' IEdmFolder5
    Dim success As Boolean
    
    success = False ' Initialize return value

    ' First check if the folder already exists in the vault
    Debug.Print "Checking if folder already exists: " & folderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        Debug.Print folderName & " folder already exists in vault: " & pdmFolder.localPath
        success = True
        EnsureFolderExists = success
        Exit Function
    End If

    ' If not, check parent folder
    Debug.Print "Checking parent folder path: " & parentFolderPath
    Set pdmParentFolder = pdmVault.GetFolderFromPath(parentFolderPath)
    
    If pdmParentFolder Is Nothing Then
        Debug.Print "Parent folder not found in vault: " & parentFolderPath
        EnsureFolderExists = success ' Return false
        Exit Function
    End If

    ' Folder doesn't exist, try to create it using rootFolder approach
    Debug.Print "Target folder not found, attempting to create: " & folderName
    
    On Error Resume Next
    ' Try using rootFolder and CreateFolderPath for complete path
    Dim rootFolder As Object
    Set rootFolder = pdmVault.rootFolder
    If Not rootFolder Is Nothing Then
        ' Create full relative path from vault root
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
    
    ' Wait a moment for the vault to update
    Sleep 1000
    
    ' Check if folder now exists
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        Debug.Print "Created " & folderName & " folder in vault: " & folderPath
        success = True
    Else
        Debug.Print "Failed to create " & folderName & " folder in vault: " & folderPath
        MsgBox "Failed to create folder in vault: " & folderPath, vbCritical
        EnsureFolderExists = success ' Return false
        Exit Function
    End If

    ' Create the local folder if it doesn't exist
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            Debug.Print "Error creating local folder: " & Err.Description
            Err.Clear
            success = False
        Else
            Debug.Print "Created " & folderName & " folder locally: " & folderPath
        End If
        On Error GoTo 0
    Else
        Debug.Print folderName & " folder already exists locally: " & folderPath
    End If
    
    EnsureFolderExists = success ' Return success status
End Function

' Function to extract revision from filename
Private Function ExtractRevision(fileNameStr As String, baseName As String) As String
    Dim prefix As String
    prefix = baseName & "_Rev"
    Debug.Print "ExtractRevision: File=" & fileNameStr & ", Base=" & baseName & ", Prefix=" & prefix
    If Left(fileNameStr, Len(prefix)) = prefix Then
        Dim suffix As String
        suffix = Mid(fileNameStr, Len(prefix) + 1)
        Dim dotPos As Integer
        dotPos = InStr(suffix, ".dxf")
        If dotPos > 0 Then
            ExtractRevision = Left(suffix, dotPos - 1)
            Debug.Print "Extracted revision: " & ExtractRevision
        Else
            ExtractRevision = ""
            Debug.Print "No valid extension delimiting revision."
        End If
    Else
        ExtractRevision = ""
        Debug.Print "Filename does not start with expected prefix."
    End If
End Function

' Function to prompt for custom DXF output folder
Function PromptForCustomDXFOutputFolder() As Boolean
    Dim result As VbMsgBoxResult
    Dim customPath As String

    result = MsgBox("Do you want to save the DXF to a custom folder outside of PDM?" & vbCrLf & vbCrLf & _
                   "Yes - Select a custom folder" & vbCrLf & _
                   "No - Use default PDM folder locations", _
                   vbYesNo + vbQuestion, "Output Location Selection")

    If result = vbNo Then
        bManualMode = False
        manualOutputPath = ""
        PromptForCustomDXFOutputFolder = False
        Exit Function
    End If

    ' User wants custom folder - prompt for path
    customPath = SelectFolder_WithDialog("C:\")

    If customPath = "" Then
        ' User cancelled - fall back to default
        MsgBox "No folder selected. Using default PDM locations.", vbInformation
        bManualMode = False
        manualOutputPath = ""
        PromptForCustomDXFOutputFolder = False
        Exit Function
    End If

    ' Validate the path exists
    On Error Resume Next
    If GetAttr(customPath) = -1 Then
        On Error GoTo 0
        MsgBox "Invalid folder path: " & customPath & vbCrLf & vbCrLf & _
               "Using default PDM locations instead.", vbExclamation
        bManualMode = False
        manualOutputPath = ""
        PromptForCustomDXFOutputFolder = False
        Exit Function
    End If
    On Error GoTo 0

    ' Ensure trailing backslash
    If Right(customPath, 1) <> "\" Then customPath = customPath & "\"

    ' Set manual mode
    bManualMode = True
    manualOutputPath = customPath

    MsgBox "DXFs will be saved to:" & vbCrLf & vbCrLf & manualOutputPath, vbInformation, "Custom Output Folder Set"

    PromptForCustomDXFOutputFolder = True
End Function

' Function to show folder browser dialog
Function SelectFolder_WithDialog(Optional defaultPath As String = "") As String
    Dim objShell As Object
    Dim objFolder As Object
    Dim selectedPath As String
    
    On Error Resume Next
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Select or Enter Folder Path for DXF export:", 1, defaultPath)
    
    If Not objFolder Is Nothing Then
        selectedPath = objFolder.Self.Path
    Else
        selectedPath = ""
    End If
    
    selectedPath = InputBox("Enter or modify the folder path:", "DXF Export Folder", selectedPath)
    
    If selectedPath = "" Then
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

