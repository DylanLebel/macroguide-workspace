Attribute VB_Name = "DXF_SelectedFace"
' ===== NEW MODULE: DXF_SelectedFace.bas =====
Option Explicit
Option Private Module

' Declare Windows API Sleep function with PtrSafe for 64-bit compatibility
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Global variables
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim pdmVault As EdmVault5 ' PDM Vault object
Dim filePath As String
Dim fileName As String
Dim RevValue As String

' Define surface type constants from SolidWorks API
Const swSurfType_PLANE As Long = 4001
Const swSurfType_CYLINDER As Long = 4002
Const swSurfType_CONE As Long = 4003
Const swSurfType_SPHERE As Long = 4004
Const swSurfType_TORUS As Long = 4005
Const swSurfType_BSPLINE As Long = 4006


Dim movedFiles As collection ' Track files moved to Obsolete

' Selection type constants
Const swSelFACES As Long = 2 ' Correct value for faces in SolidWorks API

' Main subroutine to execute the macro
Sub main()
Set movedFiles = New collection
    DebugLog "Starting DXF macro execution."
    ' Initialize SolidWorks application
    Set swApp = Application.SldWorks
    Set swModel = swApp.activeDoc
    DebugLog "Active document: " & IIf(swModel Is Nothing, "None", swModel.GetPathName)

    ' Check if no document is open
    If swModel Is Nothing Then
        MsgBox "No document open.", vbCritical
        DebugLog "Error: No document open."
        Exit Sub
    End If

    ' Check if the active document is a part or assembly
    If swModel.GetType <> swDocPART And swModel.GetType <> swDocASSEMBLY Then
        MsgBox "Active document is not a part or assembly.", vbCritical
        DebugLog "Error: Active document is not a part or assembly. Type: " & swModel.GetType
        Exit Sub
    End If
    DebugLog "Document is a part or assembly."

    ' Check if the model is saved
    If swModel.GetPathName = "" Then
        MsgBox "Please save the model first.", vbCritical
        DebugLog "Error: Model not saved. Path: " & swModel.GetPathName
        Exit Sub
    End If
    DebugLog "Model is saved."

    ' Get selected face
    Dim selMgr As SldWorks.SelectionMgr
    Set selMgr = swModel.SelectionManager
    DebugLog "Selection count: " & selMgr.GetSelectedObjectCount2(-1)
    If selMgr.GetSelectedObjectCount2(-1) <> 1 Then
        MsgBox "Please select exactly one face.", vbCritical
        DebugLog "Error: Selection count is not 1. Count: " & selMgr.GetSelectedObjectCount2(-1)
        Exit Sub
    End If
    Dim selType As Long
    selType = selMgr.GetSelectedObjectType3(1, -1)
    DebugLog "Selected object type ID: " & selType & " (swSelFACES = 2)"
    If selType <> swSelFACES Then
        MsgBox "Selected object is not a face.", vbCritical
        DebugLog "Error: Selected object is not a face. Type: " & selType
        Exit Sub
    End If
    Dim swFace As SldWorks.Face2
    Set swFace = selMgr.GetSelectedObject6(1, -1)
    DebugLog "Face object retrieved: " & IIf(swFace Is Nothing, "No", "Yes")
    If swFace Is Nothing Then
        MsgBox "Failed to retrieve face object.", vbCritical
        DebugLog "Error: Face object is Nothing."
        Exit Sub
    End If

    ' Get revision from custom properties
    Dim cusPropMgr As SldWorks.customPropertyManager
    Set cusPropMgr = swModel.Extension.customPropertyManager("")
    Dim wasResolved As Boolean
    cusPropMgr.Get5 "Revision", False, "", RevValue, wasResolved
    DebugLog "Revision retrieved: " & RevValue & ", Resolved: " & wasResolved
    If Not wasResolved Or Trim(RevValue) = "" Then
        MsgBox "Revision is not set or blank.", vbCritical
        DebugLog "Error: Revision is not set or blank."
        Exit Sub
    End If

   ' Determine file paths
Dim fullPath As String
fullPath = swModel.GetPathName
fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
fileName = Left(fileName, InStrRev(fileName, ".") - 1)
filePath = Left(fullPath, InStrRev(fullPath, "\"))
DebugLog "File Name: " & fileName
DebugLog "File Path: " & filePath

' Check if in Libraries path
Dim isLibraryPath As Boolean
isLibraryPath = (InStr(1, fullPath, PDM_LIBRARIES_PATH, vbTextCompare) > 0)
DebugLog "Is Library Path: " & isLibraryPath

Dim burnProfilesFolder As String
Dim obsoleteFolder As String

If isLibraryPath Then
    ' Create a Drawings\Burn Profiles subfolder in the library path
    burnProfilesFolder = filePath & "Drawings\Burn Profiles\"
    obsoleteFolder = burnProfilesFolder & "Obsolete\"
    
    ' Ensure Drawings folder exists first
    EnsureFolderExists filePath, "Drawings", filePath & "Drawings\"
    
    ' Then ensure Burn Profiles folder exists
    EnsureFolderExists filePath & "Drawings\", "Burn Profiles", burnProfilesFolder
    
    DebugLog "Library path detected. Using burn profiles folder: " & burnProfilesFolder
    DebugLog "Library obsolete folder: " & obsoleteFolder
Else
    ' Original logic for project files
    Dim designFolder As String
    Dim pos As Long
    pos = InStrRev(filePath, "3 - Design\")
    If pos > 0 Then
        designFolder = Left(filePath, pos + Len("3 - Design") - 1) & "\"
    Else
        MsgBox "Could not find '3 - Design' in the model path.", vbCritical
        DebugLog "Error: Could not find '3 - Design' in model path: " & filePath
        Exit Sub
    End If
    DebugLog "Design Folder: " & designFolder

    burnProfilesFolder = designFolder & "Drawings\Burn Profiles\"
    obsoleteFolder = burnProfilesFolder & "Obsolete\"
    DebugLog "Burn Profiles Folder: " & burnProfilesFolder
    DebugLog "Obsolete Folder: " & obsoleteFolder
End If

' Connect to PDM vault
If Not ConnectToPDMVault() Then
    MsgBox "Failed to connect to NMT_PDM vault.", vbCritical
    DebugLog "Error: Failed to connect to NMT_PDM vault."
    Exit Sub
End If
DebugLog "Connected to PDM vault successfully."

' Ensure folders exist
If isLibraryPath Then
    ' We already created the necessary folders above
    DebugLog "Library path: Folders already created."
Else
    DebugLog "Ensuring 'Burn Profiles' folder exists."
    EnsureFolderExists designFolder & "Drawings\", "Burn Profiles", burnProfilesFolder
End If

DebugLog "Ensuring 'Obsolete' folder exists."
EnsureFolderExists burnProfilesFolder, "Obsolete", obsoleteFolder

    ' Get PDM folder objects
    Dim pdmBurnFolder As Object ' IEdmFolder5
    Set pdmBurnFolder = pdmVault.GetFolderFromPath(burnProfilesFolder)
    If pdmBurnFolder Is Nothing Then
        MsgBox "Could not find Burn Profiles folder in vault.", vbCritical
        DebugLog "Error: Could not find Burn Profiles folder in vault: " & burnProfilesFolder
        Exit Sub
    End If
    DebugLog "Burn Profiles folder ID: " & pdmBurnFolder.ID
    DebugLog "Burn Profiles folder path: " & pdmBurnFolder.localPath

    Dim pdmObsoleteFolder As Object ' IEdmFolder5
    Set pdmObsoleteFolder = pdmVault.GetFolderFromPath(obsoleteFolder)
    If pdmObsoleteFolder Is Nothing Then
        MsgBox "Could not find Obsolete folder in vault.", vbCritical
        DebugLog "Error: Could not find Obsolete folder in vault: " & obsoleteFolder
        Exit Sub
    End If
    DebugLog "Obsolete folder ID: " & pdmObsoleteFolder.ID
    DebugLog "Obsolete folder path: " & pdmObsoleteFolder.localPath

    ' Construct target DXF filename without version suffix
    Dim targetDXFName As String
    targetDXFName = fileName & "_Rev" & RevValue & ".dxf"
    Dim targetDXFPath As String
    targetDXFPath = burnProfilesFolder & targetDXFName
    DebugLog "Target DXF Path: " & targetDXFPath

    ' Step 1: Iterate through all files in Burn Profiles folder and filter manually (adopted from ListFilesInBurnProfiles)
    DebugLog "Starting to list all files in Burn Profiles folder."
    Dim posFile As IEdmPos5
    Set posFile = pdmBurnFolder.GetFirstFilePosition
    If posFile Is Nothing Then
        DebugLog "No files found in Burn Profiles folder."
    Else
        Dim pdmFileIter As IEdmFile5
        Dim pattern As String
        pattern = fileName & "_Rev*.dxf"
        DebugLog "Looking for files matching pattern: " & pattern
        While Not posFile.IsNull
            Set pdmFileIter = pdmBurnFolder.GetNextFile(posFile)
            If Not pdmFileIter Is Nothing Then
                If pdmFileIter.Name Like pattern Then
                    Dim dxfPath As String
                    dxfPath = pdmFileIter.GetLocalPath(pdmBurnFolder.ID)
                    Dim dxfName As String
                    dxfName = pdmFileIter.Name
                    DebugLog "Found DXF: " & dxfPath
                    DebugLog "Processing file: " & dxfName

                    ' Extract revision from filename
                    Dim existingRev As String
                    existingRev = ExtractRevision(dxfName, fileName)
                    DebugLog "Extracted revision: " & existingRev & ", Current revision: " & RevValue
                    If existingRev <> "" And IsNumeric(existingRev) And IsNumeric(RevValue) Then
                        If CInt(existingRev) < CInt(RevValue) Then
                            DebugLog "Moving lower revision DXF to Obsolete: Revision " & existingRev & " < " & RevValue & "; Path: " & dxfPath
                            Dim pdmFileToMove As IEdmFile5
                            Set pdmFileToMove = pdmFileIter
                            Dim newFileInObsolete As IEdmFile5
                            ' Process only if file object retrieved successfully
                            If Not pdmFileToMove Is Nothing Then
                                DebugLog "pdmFileToMove.Name: " & pdmFileToMove.Name
                                DebugLog "pdmFileToMove.ID: " & pdmFileToMove.ID
                                DebugLog "pdmFileToMove.IsLocked before lock: " & pdmFileToMove.IsLocked
                                If Not pdmFileToMove.IsLocked Then
                                    On Error Resume Next
                                    pdmFileToMove.LockFile pdmBurnFolder.ID, 0
                                    If Err.Number <> 0 Then
                                        DebugLog "Error locking file for move: " & Err.description
                                        Err.Clear
                                    Else
                                        DebugLog "Locked file for move: " & dxfPath
                                    End If
                                    On Error GoTo 0
                                End If
                                Dim localFilePath As String
                                localFilePath = pdmBurnFolder.localPath & "\" & pdmFileToMove.Name
                                DebugLog "Local file path for move: " & localFilePath
                                If Dir(localFilePath) = "" Then
                                    DebugLog "Local file not found: " & localFilePath
                                Else
                                    ' Check if file already exists in Obsolete
                                    Dim searchObsolete As IEdmSearch5
                                    Set searchObsolete = pdmVault.CreateSearch
                                    searchObsolete.StartFolderID = pdmObsoleteFolder.ID
                                    searchObsolete.fileName = pdmFileToMove.Name

                                    Dim resultObsolete As IEdmSearchResult5
                                    On Error Resume Next
                                    Set resultObsolete = searchObsolete.GetFirstResult
                                    If Err.Number <> 0 Then
                                        DebugLog "Error searching Obsolete: " & Err.description
                                        Err.Clear
                                    End If
                                    On Error GoTo 0

                                    If Not resultObsolete Is Nothing Then
                                        DebugLog "File already in Obsolete: " & resultObsolete.path
                                    Else
                                        DebugLog "File not in Obsolete; adding now."
                                        On Error Resume Next
                                        pdmObsoleteFolder.AddFile 0, localFilePath
                                        If Err.Number <> 0 Then
                                            DebugLog "Error adding to Obsolete: " & Err.description
                                            Err.Clear
                                        Else
                                            DebugLog "Added to Obsolete: " & localFilePath
                                            Set newFileInObsolete = pdmObsoleteFolder.GetFile(pdmFileToMove.Name)
                                            If Not newFileInObsolete Is Nothing Then
                                                newFileInObsolete.UnlockFile pdmObsoleteFolder.ID, "Checked in after move", 0
                                                DebugLog "Checked in to Obsolete: " & newFileInObsolete.Name
                                            Else
                                                DebugLog "Failed to get file in Obsolete."
                                            End If
                                        End If
                                        On Error GoTo 0
                                    End If
                                    
                                    ' Change file state to Obsolete using batch changer utility
                                    Dim fileToChangeState As IEdmFile5
                                    If Not newFileInObsolete Is Nothing Then
                                        Set fileToChangeState = newFileInObsolete
                                    Else
                                        Set fileToChangeState = pdmFileToMove
                                    End If
                                    
                                    Dim batchChanger As EdmLib.IEdmBatchChangeState4
                                    Set batchChanger = pdmVault.CreateUtility(EdmLib.EdmUtility.EdmUtil_BatchChangeState)

                                    If batchChanger Is Nothing Then
                                        DebugLog "Error: Failed to create batch changer utility."
                                    Else
                                        DebugLog "Batch changer utility created successfully."
                                        batchChanger.AddFile fileToChangeState.ID, pdmObsoleteFolder.ID
                                        Dim retVal As Boolean
                                        retVal = batchChanger.CreateTree("Obsolete")
                                        DebugLog "CreateTree returned: " & retVal
                                        If retVal Then
                                            batchChanger.ChangeState2 0, ""
                                            DebugLog "File state changed to Obsolete."
                                        Else
                                            DebugLog "Error: State transition tree creation failed."
                                        End If
                                    End If
                                    
                                    ' Check-in in Burn Profiles before removal
                                    DebugLog "pdmFileToMove.IsLocked before check-in: " & pdmFileToMove.IsLocked
                                    If pdmFileToMove.IsLocked Then
                                        pdmFileToMove.UnlockFile pdmBurnFolder.ID, "Checked in before removal", 0
                                        DebugLog "Checked in file in Burn Profiles."
                                    Else
                                        DebugLog "File already checked in Burn Profiles."
                                    End If
                                  '  Sleep 500 ' Wait for vault update
                                    DebugLog "pdmFileToMove.IsLocked after check-in: " & pdmFileToMove.IsLocked

                                    ' Refresh file object
                                    pdmFileToMove.Refresh
                                    DebugLog "Refreshed pdmFileToMove."

                                    ' Verify file still in Burn Profiles
                                    Dim exists As IEdmFile5
                                    On Error Resume Next
                                    Set exists = pdmBurnFolder.GetFile(pdmFileToMove.Name)
                                    If Err.Number <> 0 Then
                                        DebugLog "Error getting file after check-in: " & Err.description & " - Assuming file is not in Burn Profiles."
                                        Set exists = Nothing
                                        Err.Clear
                                    End If
                                    On Error GoTo 0

                                    If exists Is Nothing Then
                                        DebugLog "File no longer in Burn Profiles after check-in."
                                    Else
                                        DebugLog "File still in Burn Profiles; attempting removal."
                                        On Error Resume Next
                                        pdmBurnFolder.DeleteFile pdmBurnFolder.ID, pdmFileToMove.ID, 0
                                        If Err.Number <> 0 Then
                                            DebugLog "Error with DeleteFile: " & Err.description
                                            Err.Clear
                                        Else
                                            DebugLog "Deleted from Burn Profiles with DeleteFile: " & pdmFileToMove.Name
                                        End If
                                        On Error GoTo 0

                                        ' Confirm removal
                                        On Error Resume Next
                                        Set exists = pdmBurnFolder.GetFile(pdmFileToMove.Name)
                                        If Err.Number <> 0 Then
                                            DebugLog "Error confirming removal: " & Err.description & " - Assuming file is not in Burn Profiles."
                                            Set exists = Nothing
                                            Err.Clear
                                        End If
                                        On Error GoTo 0

                                        If exists Is Nothing Then
                                            DebugLog "Confirmed file removed from Burn Profiles."
                                        Else
                                            DebugLog "File still in Burn Profiles after DeleteFile!"
                                        End If
                                    End If
                                    If Not resultObsolete Is Nothing Then
                                        DebugLog "File already in Obsolete, processed."
                                    Else
' Track this move for potential undo
movedFiles.Add dxfName
                                        DebugLog "Moved " & dxfName & " to Obsolete."
                                    End If
                                End If
                            Else
                                DebugLog "Failed to get pdmFileToMove for: " & dxfPath
                            End If
                        End If
                    End If
                End If
            End If
        Wend
    End If

    ' Step 2: Handle the target DXF
    Dim pdmFile As Object ' IEdmFile5
    Set pdmFile = pdmVault.GetFileFromPath(targetDXFPath)
    DebugLog "Checking if target DXF exists in vault: " & targetDXFPath
    If Not pdmFile Is Nothing Then
        DebugLog "Target DXF exists in vault: " & targetDXFPath
        DebugLog "pdmFile.IsLocked: " & pdmFile.IsLocked
        If Not pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.LockFile pdmBurnFolder.ID, 0
            If Err.Number <> 0 Then
                DebugLog "Error locking existing DXF: " & Err.description
                Err.Clear
            Else
                DebugLog "Locked existing DXF: " & targetDXFPath
            End If
            On Error GoTo 0
        End If
    Else
        DebugLog "Target DXF does not exist in vault: " & targetDXFPath
    End If

    ' Export the new DXF
    DebugLog "Exporting DXF to: " & targetDXFPath
    Dim dataAlignment(11) As Double
    ' Initialize alignment array (ensure this is properly defined as per your SolidWorks setup)
    dataAlignment(0) = 0: dataAlignment(1) = 0: dataAlignment(2) = 0
    dataAlignment(3) = 1: dataAlignment(4) = 0: dataAlignment(5) = 0
    dataAlignment(6) = 0: dataAlignment(7) = 1: dataAlignment(8) = 0
    dataAlignment(9) = 0: dataAlignment(10) = 0: dataAlignment(11) = 0
    Dim varAlignment As Variant
    varAlignment = dataAlignment
    On Error Resume Next
    swModel.ExportToDWG2 targetDXFPath, swModel.GetPathName, swExportToDWG_ExportSelectedFacesOrLoops, True, varAlignment, False, False, 0, Null
    If Err.Number <> 0 Then
        DebugLog "Error exporting DXF: " & Err.description
        MsgBox "Failed to export DXF to " & targetDXFPath & ": " & Err.description, vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    DebugLog "DXF export completed."

    ' Check if the file was created locally
    If Dir(targetDXFPath) = "" Then
        MsgBox "Failed to export DXF to " & targetDXFPath & ". File not found.", vbCritical
        DebugLog "Error: Failed to export DXF. Local file not found: " & targetDXFPath
        Exit Sub
    End If
    DebugLog "Local DXF file exists: " & targetDXFPath

    ' Check in or add the file
    If Not pdmFile Is Nothing Then
        DebugLog "pdmFile is not Nothing, proceeding with update."
        If pdmFile.IsLocked Then
            On Error Resume Next
            pdmFile.UnlockFile pdmBurnFolder.ID, "Updated by macro", 0
            If Err.Number <> 0 Then
                DebugLog "Error checking in updated DXF: " & Err.description
                Err.Clear
            Else
                DebugLog "Checked in updated DXF: " & targetDXFPath
            End If
            On Error GoTo 0
        Else
            DebugLog "Existing DXF was not locked after export, which is unexpected."
        End If
    Else
        DebugLog "Adding new DXF to vault."
        On Error Resume Next
        pdmBurnFolder.AddFile 0, targetDXFPath
        If Err.Number <> 0 Then
            DebugLog "Error adding DXF to vault: " & Err.description
            Err.Clear
        Else
            DebugLog "Added DXF to vault: " & targetDXFPath
            Set pdmFile = pdmVault.GetFileFromPath(targetDXFPath)
            If Not pdmFile Is Nothing Then
                DebugLog "pdmFile retrieved after adding: " & pdmFile.Name
                DebugLog "pdmFile.IsLocked: " & pdmFile.IsLocked
                If pdmFile.IsLocked Then
                    pdmFile.UnlockFile pdmBurnFolder.ID, "Added by macro", 0
                    DebugLog "Checked in new DXF: " & targetDXFPath
                End If
            Else
                DebugLog "Failed to retrieve pdmFile after adding to vault: " & targetDXFPath
            End If
        End If
        On Error GoTo 0
    End If

    Dim undoResponse As VbMsgBoxResult
undoResponse = MsgBox("DXF '" & targetDXFName & "' created and processed successfully." & vbCrLf & vbCrLf & _
                      "Click YES to UNDO this operation, or NO to keep the changes.", _
                      vbYesNo + vbQuestion, "DXF Complete")

If undoResponse = vbYes Then
    Call UndoOperation(pdmFile, pdmBurnFolder, pdmObsoleteFolder, targetDXFPath, targetDXFName)
End If


    DebugLog "DXF macro completed successfully."
End Sub

' Function to connect to NMT_PDM vault
Private Function ConnectToPDMVault() As Boolean
    On Error Resume Next
    Set pdmVault = New EdmVault5
    pdmVault.LoginAuto "NMT_PDM", 0 ' Auto-login with current Windows credentials
    If pdmVault.IsLoggedIn Then
        ConnectToPDMVault = True
        DebugLog "Logged into NMT_PDM vault successfully."
    Else
        ConnectToPDMVault = False
        DebugLog "Failed to log into NMT_PDM vault."
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
    DebugLog "Checking if folder already exists: " & folderPath
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        DebugLog folderName & " folder already exists in vault: " & pdmFolder.localPath
        success = True
        EnsureFolderExists = success
        Exit Function
    End If

    ' If not, check parent folder
    DebugLog "Checking parent folder path: " & parentFolderPath
    Set pdmParentFolder = pdmVault.GetFolderFromPath(parentFolderPath)
    
    If pdmParentFolder Is Nothing Then
        DebugLog "Parent folder not found in vault: " & parentFolderPath
        EnsureFolderExists = success ' Return false
        Exit Function
    End If

    ' Folder doesn't exist, try to create it using rootFolder approach
    DebugLog "Target folder not found, attempting to create: " & folderName
    
    On Error Resume Next
    ' Try using rootFolder and CreateFolderPath for complete path
    Dim rootFolder As Object
    Set rootFolder = pdmVault.rootFolder
    If Not rootFolder Is Nothing Then
        ' Create full relative path from vault root
        Dim vaultPath As String
        vaultPath = Replace(folderPath, PDM_VAULT_ROOT, "\")
        rootFolder.CreateFolderPath vaultPath, 0
        
        If Err.Number <> 0 Then
            DebugLog "CreateFolderPath from root failed: " & Err.description & " (Error #" & Err.Number & ")"
            Err.Clear
            EnsureFolderExists = False
            Exit Function
        End If
    Else
        DebugLog "Could not get root folder"
        EnsureFolderExists = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' Wait a moment for the vault to update
    Sleep 1000
    
    ' Check if folder now exists
    Set pdmFolder = pdmVault.GetFolderFromPath(folderPath)
    If Not pdmFolder Is Nothing Then
        DebugLog "Created " & folderName & " folder in vault: " & folderPath
        success = True
    Else
        DebugLog "Failed to create " & folderName & " folder in vault: " & folderPath
        MsgBox "Failed to create folder in vault: " & folderPath, vbCritical
        EnsureFolderExists = success ' Return false
        Exit Function
    End If

    ' Create the local folder if it doesn't exist
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            DebugLog "Error creating local folder: " & Err.description
            Err.Clear
            success = False
        Else
            DebugLog "Created " & folderName & " folder locally: " & folderPath
        End If
        On Error GoTo 0
    Else
        DebugLog folderName & " folder already exists locally: " & folderPath
    End If
    
    EnsureFolderExists = success ' Return success status
End Function

' Function to extract revision from filename
Private Function ExtractRevision(fileNameStr As String, baseName As String) As String
    Dim prefix As String
    prefix = baseName & "_Rev"
    DebugLog "ExtractRevision: File=" & fileNameStr & ", Base=" & baseName & ", Prefix=" & prefix
    If Left(fileNameStr, Len(prefix)) = prefix Then
        Dim suffix As String
        suffix = Mid(fileNameStr, Len(prefix) + 1)
        Dim dotPos As Integer
        dotPos = InStr(suffix, ".dxf")
        If dotPos > 0 Then
            ExtractRevision = Left(suffix, dotPos - 1)
            DebugLog "Extracted revision: " & ExtractRevision
        Else
            ExtractRevision = ""
            DebugLog "No valid extension delimiting revision."
        End If
    Else
        ExtractRevision = ""
        DebugLog "Filename does not start with expected prefix."
    End If
End Function
' ===== END NEW MODULE =====


' Undo the DXF operation
Private Sub UndoOperation(pdmFile As Object, pdmBurnFolder As Object, pdmObsoleteFolder As Object, dxfPath As String, dxfName As String)
    DebugLog "Starting undo operation..."
    
    ' Delete the newly created DXF
    If Not pdmFile Is Nothing Then
        DebugLog "Deleting newly created DXF: " & dxfPath
        On Error Resume Next
        
        ' Lock if not already locked
        If Not pdmFile.IsLocked Then
            pdmFile.LockFile pdmBurnFolder.ID, 0
        End If
        
        ' Delete from vault
        pdmBurnFolder.DeleteFile pdmBurnFolder.ID, pdmFile.ID, 0
        If Err.Number <> 0 Then
            DebugLog "Error deleting DXF from vault: " & Err.description
            Err.Clear
        Else
            DebugLog "Deleted DXF from vault successfully."
        End If
        On Error GoTo 0
    End If
    
    ' Delete local file if it exists
    If Dir(dxfPath) <> "" Then
        On Error Resume Next
        Kill dxfPath
        If Err.Number <> 0 Then
            DebugLog "Error deleting local DXF file: " & Err.description
            Err.Clear
        Else
            DebugLog "Deleted local DXF file."
        End If
        On Error GoTo 0
    End If
    
    ' Move files back from Obsolete to Burn Profiles
    Dim fileName As Variant
    For Each fileName In movedFiles
        DebugLog "Moving back from Obsolete: " & fileName
        
        Dim pdmObsoleteFile As IEdmFile5
        Set pdmObsoleteFile = pdmObsoleteFolder.GetFile(CStr(fileName))
        
        If Not pdmObsoleteFile Is Nothing Then
            Dim obsoleteFilePath As String
            obsoleteFilePath = pdmObsoleteFile.GetLocalPath(pdmObsoleteFolder.ID)
            
            On Error Resume Next
            
            ' Lock the file in Obsolete
            If Not pdmObsoleteFile.IsLocked Then
                pdmObsoleteFile.LockFile pdmObsoleteFolder.ID, 0
                DebugLog "Locked file in Obsolete: " & fileName
            End If
            
            ' Check if file already exists in Burn Profiles
            Dim existingInBurn As IEdmFile5
            Set existingInBurn = pdmBurnFolder.GetFile(CStr(fileName))
            
            If existingInBurn Is Nothing Then
                ' Add back to Burn Profiles
                pdmBurnFolder.AddFile 0, obsoleteFilePath
                If Err.Number <> 0 Then
                    DebugLog "Error adding file back to Burn Profiles: " & Err.description
                    Err.Clear
                Else
                    DebugLog "Added back to Burn Profiles: " & fileName
                    
                    ' Check in the file in Burn Profiles
                    Dim restoredFile As IEdmFile5
                    Set restoredFile = pdmBurnFolder.GetFile(CStr(fileName))
                    If Not restoredFile Is Nothing And restoredFile.IsLocked Then
                        restoredFile.UnlockFile pdmBurnFolder.ID, "Restored by undo", 0
                        DebugLog "Checked in restored file."
                    End If
                End If
            Else
                DebugLog "File already exists in Burn Profiles: " & fileName
            End If
            
            ' Check in and delete from Obsolete
            If pdmObsoleteFile.IsLocked Then
                pdmObsoleteFile.UnlockFile pdmObsoleteFolder.ID, "Removing from Obsolete", 0
            End If
            
            pdmObsoleteFolder.DeleteFile pdmObsoleteFolder.ID, pdmObsoleteFile.ID, 0
            If Err.Number <> 0 Then
                DebugLog "Error deleting from Obsolete: " & Err.description
                Err.Clear
            Else
                DebugLog "Deleted from Obsolete: " & fileName
            End If
            
            On Error GoTo 0
        Else
            DebugLog "Could not find file in Obsolete: " & fileName
        End If
    Next fileName
    
    MsgBox "Undo completed. All changes have been reversed.", vbInformation
    DebugLog "Undo operation completed."
End Sub
