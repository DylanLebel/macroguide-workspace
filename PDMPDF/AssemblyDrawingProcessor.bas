Attribute VB_Name = "AssemblyDrawingProcessor"
' ============================================================================
' MODULE: AssemblyDrawingProcessor (FIXED VERSION)
' Description: Process PDFs for all drawings related to an assembly
' FIX: Closes the assembly before processing component drawings
' ============================================================================

Option Explicit

' Windows API declaration for Sleep function
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Main entry point for assembly drawing processing
Sub ProcessAssemblyDrawings()
    Dim swAssy As SldWorks.ModelDoc2
    Dim componentList As Collection
    Dim drawingList As Collection
    Dim totalDrawings As Long
    Dim currentNum As Long
    Dim drawingPath As Variant
    Dim userResponse As VbMsgBoxResult
    Dim messageText As String
    Dim originalVisibility As Boolean
    
    ' Verify we have an active assembly
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            MsgBox "Failed to connect to SolidWorks.", vbCritical
            Exit Sub
        End If
    End If
    
    Set swAssy = swApp.ActiveDoc
    If swAssy Is Nothing Then
        MsgBox "No document is currently open.", vbExclamation
        Exit Sub
    End If
    
    If swAssy.GetType <> swDocASSEMBLY Then
        MsgBox "The active document is not an assembly." & vbCrLf & vbCrLf & _
               "Please open an assembly and try again.", vbExclamation, "Not an Assembly"
        Exit Sub
    End If
    
    ' Get all unique components from the assembly
    Debug.Print "=== COLLECTING ASSEMBLY COMPONENTS ==="
    Set componentList = GetUniqueComponents(swAssy)
    Debug.Print "Found " & componentList.count & " unique components"
    
    If componentList.count = 0 Then
        MsgBox "No components found in this assembly.", vbInformation
        Exit Sub
    End If
    
    ' Find drawings for each component (excluding library items)
    Debug.Print "=== FINDING DRAWINGS FOR COMPONENTS ==="
    Set drawingList = FindDrawingsForComponents(componentList)
    Debug.Print "Found " & drawingList.count & " drawings to process"
    
    If drawingList.count = 0 Then
        MsgBox "No drawings found for any components in this assembly." & vbCrLf & vbCrLf & _
               "Note: Library items are automatically excluded.", vbInformation
        Exit Sub
    End If
    
    ' Build confirmation message
    messageText = "Found " & drawingList.count & " drawing(s) for components in this assembly." & vbCrLf & vbCrLf
    messageText = messageText & "This will:" & vbCrLf
    messageText = messageText & "� Process each component drawing" & vbCrLf
    messageText = messageText & "� Temporarily close any assemblies using each component" & vbCrLf
    messageText = messageText & "� Sync model revisions with drawing revisions" & vbCrLf
    messageText = messageText & "� Create/update PDF files" & vbCrLf
    messageText = messageText & "� Move obsolete PDFs to Obsolete folders" & vbCrLf
    messageText = messageText & "� Reopen closed assemblies automatically" & vbCrLf
    messageText = messageText & "� Skip all library items" & vbCrLf & vbCrLf
    messageText = messageText & "Do you want to continue?"
    
    userResponse = MsgBox(messageText, vbQuestion + vbYesNo + vbDefaultButton1, "Process Assembly Drawings?")
    
    If userResponse = vbNo Then
        Debug.Print "User canceled assembly drawing processing"
        Exit Sub
    End If
    
    ' Note: We do NOT close assemblies here. Instead, the SyncDrawingWithModel
    ' function will detect and close any assemblies that reference each component
    ' when processing its drawing. This reuses the existing, working logic.
    
    ' Hide SolidWorks during processing
    On Error Resume Next
    originalVisibility = swApp.Visible
    swApp.Visible = False
    swApp.UserControl = False
    Debug.Print "SolidWorks hidden for batch processing"
    On Error GoTo 0
    
    ' Show progress form
    ProgressForm.Show vbModeless
    ProgressForm.Caption = "Assembly PDF Creation Progress"
    totalDrawings = drawingList.count
    
    ' Process each drawing using the EXISTING main module logic
    currentNum = 0
    For Each drawingPath In drawingList
        currentNum = currentNum + 1
        
        ' Update progress
        ProgressForm.UpdateProgress currentNum, totalDrawings, Mid(CStr(drawingPath), InStrRev(CStr(drawingPath), "\") + 1)
        DoEvents
        
        ' Process this drawing using the main module's logic
        ProcessSingleDrawing CStr(drawingPath)
    Next drawingPath
    
    ' Cleanup and restore
    On Error Resume Next
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    swApp.Visible = originalVisibility
    swApp.UserControl = True
    On Error GoTo 0
    
    MsgBox "Processed " & totalDrawings & " drawing(s) successfully!", vbInformation, "Assembly Processing Complete"
    
    Debug.Print "Assembly drawing processing complete"
End Sub

' Get all unique components from an assembly (recursive)
Public Function GetUniqueComponents(swAssy As SldWorks.ModelDoc2) As Collection
    Dim uniquePaths As Object ' Scripting.Dictionary
    Dim componentList As Collection
    
    Set uniquePaths = CreateObject("Scripting.Dictionary")
    uniquePaths.CompareMode = vbTextCompare
    
    Set componentList = New Collection
    
    ' Get the active configuration
    Dim swConfMgr As SldWorks.ConfigurationManager
    Dim swConf As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    
    Set swConfMgr = swAssy.ConfigurationManager
    If swConfMgr Is Nothing Then
        Set GetUniqueComponents = componentList
        Exit Function
    End If
    
    Set swConf = swConfMgr.ActiveConfiguration
    If swConf Is Nothing Then
        Set GetUniqueComponents = componentList
        Exit Function
    End If
    
    Set swRootComp = swConf.GetRootComponent3(True)
    If swRootComp Is Nothing Then
        Set GetUniqueComponents = componentList
        Exit Function
    End If
    
    ' Recursively traverse the assembly
    TraverseComponent swRootComp, uniquePaths
    
    ' Convert dictionary to collection
    Dim componentPath As Variant
    For Each componentPath In uniquePaths.Keys
        componentList.Add CStr(componentPath)
        Debug.Print "Component: " & componentPath
    Next componentPath
    
    Set GetUniqueComponents = componentList
End Function

' Recursively traverse assembly components
Private Sub TraverseComponent(swComp As SldWorks.Component2, ByRef uniquePaths As Object)
    If swComp Is Nothing Then Exit Sub
    
    ' Get the component's path directly (works with lightweight components)
    Dim compPath As String
    compPath = swComp.GetPathName
    
    If compPath <> "" Then
        ' Add to dictionary if not already present
        If Not uniquePaths.exists(compPath) Then
            uniquePaths.Add compPath, True
            ' Debug info shows suppression state but we don't skip
            Debug.Print "Added component: " & compPath & " (Suppressed in active config: " & swComp.IsSuppressed & ")"
        End If
    End If
    
    ' Traverse children
    Dim swChildComp As SldWorks.Component2
    Dim childrenArray As Variant
    Dim i As Long
    
    childrenArray = swComp.GetChildren
    If Not IsEmpty(childrenArray) Then
        For i = LBound(childrenArray) To UBound(childrenArray)
            Set swChildComp = childrenArray(i)
            TraverseComponent swChildComp, uniquePaths
        Next i
    End If
End Sub

' Find drawings for a list of component paths
Public Function FindDrawingsForComponents(componentList As Collection) As Collection
    Dim drawingList As Collection
    Dim componentPath As Variant
    Dim drawingPath As String
    Dim isLibrary As Boolean
    Dim foundDrawings As Object ' Scripting.Dictionary to avoid duplicates
    
    Set foundDrawings = CreateObject("Scripting.Dictionary")
    foundDrawings.CompareMode = vbTextCompare
    
    Set drawingList = New Collection
    
    For Each componentPath In componentList
        ' Check if this is a library item - skip if it is
        isLibrary = IsLibraryPathCheck(CStr(componentPath))
        
        If isLibrary Then
            Debug.Print "Skipping library item: " & componentPath
        Else
            ' Look for a drawing file
            drawingPath = FindDrawingForModel(CStr(componentPath))
            
            If drawingPath <> "" Then
                ' Add to collection if not already present
                If Not foundDrawings.exists(drawingPath) Then
                    drawingList.Add drawingPath
                    foundDrawings.Add drawingPath, True
                    Debug.Print "Found drawing: " & drawingPath
                End If
            Else
                Debug.Print "No drawing found for: " & componentPath
            End If
        End If
    Next componentPath
    
    Set FindDrawingsForComponents = drawingList
End Function

' Find a drawing file for a given model path
Private Function FindDrawingForModel(modelPath As String) As String
    Dim baseName As String
    Dim modelDir As String
    Dim drawingPath As String
    Dim possibleExtensions As Variant
    Dim ext As Variant
    
    FindDrawingForModel = "" ' Default to not found
    
    If modelPath = "" Then Exit Function
    
    ' Extract base name and directory
    baseName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
    baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    modelDir = Left(modelPath, InStrRev(modelPath, "\"))
    
    ' Try different case variations of .slddrw extension
    possibleExtensions = Array(".slddrw", ".SLDDRW", ".SldDrw")
    
    For Each ext In possibleExtensions
        drawingPath = modelDir & baseName & ext
        
        ' Check if file exists locally
        If Dir(drawingPath) <> "" Then
            FindDrawingForModel = drawingPath
            Exit Function
        End If
    Next ext
    
    ' Drawing not found
    Debug.Print "No drawing file found for model: " & baseName
End Function

' Process a single drawing - leverages the existing main module logic
Private Sub ProcessSingleDrawing(drawingPath As String)
    Dim swDoc As SldWorks.ModelDoc2
    Dim errs As Long
    Dim warns As Long
    
    Debug.Print "Processing: " & drawingPath
    
    ' Open the drawing
    On Error Resume Next
    Set swDoc = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errs, warns)
    If Err.Number <> 0 Or swDoc Is Nothing Then
        Debug.Print "Failed to open: " & drawingPath
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Set the global variables that the main module functions expect
    Set swModel = swDoc
    Set swDraw = swDoc
    
    ' Now run the main module's existing logic
    ' This will handle PDM connection, revision sync, PDF creation, etc.
    Dim fullPath As String
    Dim fileName As String
    Dim fileDir As String
    Dim drawingsFolder As String
    Dim obsoleteFolder As String
    Dim isLibraryPath As Boolean
    Dim pos As Long
    Dim designFolder As String
    Dim errorList As Collection
    Dim warningList As Collection
    
    Set errorList = New Collection
    Set warningList = New Collection
    
    fullPath = swDraw.GetPathName
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    fileDir = Left(fullPath, InStrRev(fullPath, "\"))
    
    ' Connect to PDM if not already connected
    If pdmVault Is Nothing Then
        If Not ConnectToPDMVault() Then
            Debug.Print "Failed to connect to PDM vault"
            GoTo CloseDrawing
        End If
    End If
    
    ' Get revisions
    Set ConfigRevisions = New Collection
    GetAllConfigRevisions
    
    If ConfigRevisions.count = 0 Then
        Debug.Print "No revision found for: " & fileName
        GoTo CloseDrawing
    End If
    
    RevValue = ConfigRevisions.item(1)
    Debug.Print "Found revision: " & RevValue & " for " & fileName
    
    ' Determine folders (use existing logic from main module)
    isLibraryPath = IsLibraryPathCheck(fullPath)
    
    If isLibraryPath Then
        drawingsFolder = fileDir & "Drawings\"
        obsoleteFolder = drawingsFolder & "Obsolete\"
        EnsureFolderExists fileDir, "Drawings", drawingsFolder
        EnsureFolderExists drawingsFolder, "Obsolete", obsoleteFolder
    Else
        pos = InStrRev(fileDir, "3 - Design\")
        If pos > 0 Then
            designFolder = Left(fileDir, pos + Len("3 - Design") - 1) & "\"
            drawingsFolder = designFolder & "Drawings\"
            obsoleteFolder = drawingsFolder & "Obsolete\"
            EnsureFolderExists designFolder, "Drawings", drawingsFolder
            EnsureFolderExists drawingsFolder, "Obsolete", obsoleteFolder
            
            ' Handle product/model subfolders
            If InStr(1, fullPath, "\Products\", vbTextCompare) > 0 Or _
               InStr(1, fullPath, "\Models\", vbTextCompare) > 0 Then
                Dim productSubfolder As String
                productSubfolder = GetProductSubfolder(fullPath)
                If productSubfolder <> "" Then
                    Dim subfolderPath As String
                    subfolderPath = drawingsFolder & productSubfolder & "\"
                    EnsureFolderExists drawingsFolder, productSubfolder, subfolderPath
                    obsoleteFolder = subfolderPath & "Obsolete\"
                    EnsureFolderExists subfolderPath, "Obsolete", obsoleteFolder
                    drawingsFolder = subfolderPath
                End If
            End If
        Else
            Debug.Print "Could not find '3 - Design' in path: " & fullPath
            GoTo CloseDrawing
        End If
    End If
    
    ' Define PDF path
    Dim pdfPath As String
    pdfPath = drawingsFolder & fileName & "_Rev" & RevValue & ".pdf"
    Debug.Print "PDF Target: " & pdfPath
    
    ' Use the existing main module functions to sync and create PDF
    Dim modelUpdated As Boolean
    modelUpdated = SyncDrawingWithModel(fullPath, drawingsFolder, obsoleteFolder, errorList, warningList, True)
    
    If modelUpdated Then
        HandlePDFCheckIn pdfPath, RevValue, drawingsFolder, obsoleteFolder, True
        Debug.Print "Successfully processed: " & fileName
    Else
        Debug.Print "Model sync failed for: " & fileName
    End If
    
CloseDrawing:
    ' Close the drawing
    If Not swDoc Is Nothing Then
        On Error Resume Next
        swApp.CloseDoc fullPath
        On Error GoTo 0
        Set swDoc = Nothing
        Set swDraw = Nothing
        Set swModel = Nothing
    End If
End Sub


