Attribute VB_Name = "DrawingTemplateUpdate1"
Option Explicit

' Object declarations
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDraw As SldWorks.DrawingDoc
Dim swSheet As SldWorks.Sheet

' Constants
Const METERS_TO_INCHES As Double = 39.3701
Const REMOVE_MODIFIED_NOTES As Boolean = True

' Collection to store failed files
Dim failedFiles As Collection

Sub BatchChangeSheetFormat()
    Dim folderPath As String
    Dim fileName As String
    Dim newFormatB As String, newFormatD As String
    Dim processedCount As Integer
    Dim failedCount As Integer
    Dim resultMessage As String
    Dim failedFilesList As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize collection for failed files
    Set failedFiles = New Collection
    
    ' Select folder containing drawings
    folderPath = BrowseForFolder("Select folder containing the drawings")
    If folderPath = "" Then
        MsgBox "No folder selected. Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    Debug.Print "Selected folder: " & folderPath
    
    ' Define correct sheet formats with error checking
    newFormatB = "C:\NMT_PDM\Libraries\Template\Sheet Formats\ANSI B - Landscape - Metric.slddrt"
    newFormatD = "C:\NMT_PDM\Libraries\Template\Sheet Formats\ANSI D - Landscape - Metric.slddrt"
    
    ' Verify format files exist
    If Not FileExists(newFormatB) Then
        MsgBox "ERROR: B-size format not found at: " & newFormatB, vbCritical
        Exit Sub
    End If
    
    If Not FileExists(newFormatD) Then
        MsgBox "ERROR: D-size format not found at: " & newFormatD, vbCritical
        Exit Sub
    End If
    
    ' Initialize SolidWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "ERROR: Failed to connect to SolidWorks.", vbCritical
        Exit Sub
    End If
    
    ' Process all drawings
    fileName = Dir(folderPath & "\*.slddrw")
    
    If fileName = "" Then
        MsgBox "ERROR: No drawing files found in selected folder.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize counters
    processedCount = 0
    failedCount = 0
    
    Do While fileName <> ""
        Debug.Print String(50, "-")
        Debug.Print "Processing file: " & fileName
        
        If ProcessDrawing(folderPath, fileName, newFormatB, newFormatD) Then
            processedCount = processedCount + 1
        Else
            failedCount = failedCount + 1
        End If
        fileName = Dir()
    Loop
    
    ' Ensure all documents are closed
    CloseAllDocuments
    
    Debug.Print String(50, "=")
    Debug.Print "Batch process completed!"
    Debug.Print "Successfully processed: " & processedCount & " files"
    Debug.Print "Failed to process: " & failedCount & " files"
    Debug.Print String(50, "=")
    
    ' Prepare result message
    resultMessage = "Batch process completed!" & vbCrLf & _
                   "Successfully processed: " & processedCount & " files" & vbCrLf & _
                   "Failed to process: " & failedCount & " files"
    
    ' Add list of failed files if any
    If failedFiles.Count > 0 Then
        failedFilesList = vbCrLf & vbCrLf & "Files that failed (not checked out or other errors):" & vbCrLf
        Dim item As Variant
        For Each item In failedFiles
            failedFilesList = failedFilesList & "- " & item & vbCrLf
        Next
        resultMessage = resultMessage & failedFilesList
    End If
    
    ' Display final result message
    MsgBox resultMessage, vbInformation, "Sheet Format Update Results"
    
    CleanupObjects
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in main process: " & Err.Description
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    CloseAllDocuments
    CleanupObjects
End Sub

Private Function ProcessDrawing(folderPath As String, fileName As String, newFormatB As String, newFormatD As String) As Boolean
    Dim vSheetNames As Variant
    Dim i As Integer
    Dim sheetsProcessed As Integer
    Dim sheetsFailed As Integer
    Dim activeSheet As String
    Dim errorMessage As String
    Dim docToClose As Object
    
    On Error GoTo ProcessError
    
    ' Open drawing
    Set swModel = swApp.OpenDoc6(folderPath & "\" & fileName, swDocDRAWING, swOpenDocOptions_Silent, "", 0, 0)
    If swModel Is Nothing Then
        Debug.Print "  Failed to open: " & fileName
        failedFiles.Add fileName & " (Failed to open)"
        ProcessDrawing = False
        Exit Function
    End If
    
    ' Keep a direct reference to the document
    Set docToClose = swModel
    
    Set swDraw = swModel
    vSheetNames = swDraw.GetSheetNames
    
    ' Store active sheet
    activeSheet = swDraw.GetCurrentSheet().GetName
    
    ' Initialize counters
    sheetsProcessed = 0
    sheetsFailed = 0
    
    ' Process each sheet
    For i = 0 To UBound(vSheetNames)
        If ProcessSheet(CStr(vSheetNames(i)), newFormatB, newFormatD) Then
            sheetsProcessed = sheetsProcessed + 1
        Else
            sheetsFailed = sheetsFailed + 1
        End If
    Next i
    
    ' Restore active sheet
    swDraw.ActivateSheet activeSheet
    
    Debug.Print "  Sheets processed successfully: " & sheetsProcessed
    Debug.Print "  Sheets failed: " & sheetsFailed
    
    ' Save and close
    swModel.ForceRebuild3 True
    
    Dim saveResult As Long
    saveResult = swModel.Save3(swSaveAsOptions_Silent, 0, 0)
    
    If saveResult <> 0 Then ' 0 means success
        Debug.Print "  ERROR: Failed to save " & fileName & " (Return code: " & saveResult & ")"
        errorMessage = "Failed to save (likely not checked out)"
        failedFiles.Add fileName & " (" & errorMessage & ")"
        ProcessDrawing = False
    Else
        ProcessDrawing = (sheetsFailed = 0)
    End If
    
    ' Now let's try to forcefully close the document
    Debug.Print "  Attempting to close " & fileName
    
    ' Close the document using different methods
    On Error Resume Next
    
    ' First, try directly closing the specific document object
    swApp.QuitDoc fileName
    DoEvents
    
    ' Second approach - try after making it active
    swApp.ActivateDoc fileName
    DoEvents
    swApp.CloseDoc fileName
    DoEvents
    
    ' Third approach - try closing all SolidWorks windows
    Dim windowCount As Long
    windowCount = swApp.GetDocumentCount
    If windowCount > 0 Then
        Debug.Print "  Attempting alternative close method..."
        
        ' Try closing all documents silently as a fallback
        swApp.CloseAllDocuments True
        DoEvents
    End If
    
    ' If we still have a valid object reference, try closing it directly
    If Not docToClose Is Nothing Then
        Set docToClose = Nothing
    End If
    
    ' Reset variables to avoid issues in next iteration
    Set swModel = Nothing
    Set swDraw = Nothing
    
    On Error GoTo 0
    
    Exit Function
    
ProcessError:
    Debug.Print "  ERROR processing " & fileName & ": " & Err.Description
    failedFiles.Add fileName & " (" & Err.Description & ")"
    ProcessDrawing = False
    
    On Error Resume Next
    ' Try all closing methods in error handler too
    swApp.QuitDoc fileName
    swApp.CloseDoc fileName
    Set docToClose = Nothing
    Set swModel = Nothing
    Set swDraw = Nothing
End Function

Private Function ProcessSheet(sheetName As String, newFormatB As String, newFormatD As String) As Boolean
    Dim sheetWidth As Double, sheetHeight As Double
    Dim sheetWidthInches As Double, sheetHeightInches As Double
    Dim currentFormat As String
    Dim newFormat As String
    Dim vProps As Variant
    
    On Error GoTo SheetError
    
    Debug.Print "    Processing sheet: " & sheetName
    
    Set swSheet = swDraw.Sheet(sheetName)
    If swSheet Is Nothing Then
        Debug.Print "    Failed to access sheet: " & sheetName
        Exit Function
    End If
    
    ' Get sheet size
    swSheet.GetSize sheetWidth, sheetHeight
    sheetWidthInches = Round(sheetWidth * METERS_TO_INCHES, 1)
    sheetHeightInches = Round(sheetHeight * METERS_TO_INCHES, 1)
    Debug.Print "    Sheet size: " & sheetWidthInches & " x " & sheetHeightInches & " inches"
    
    ' Activate sheet
    swDraw.ActivateSheet sheetName
    swModel.ViewZoomtofit2
    
    currentFormat = swSheet.GetTemplateName
    Debug.Print "    Current format: " & currentFormat
    
    ' Apply appropriate format
    If IsB_Size(sheetWidthInches, sheetHeightInches) Then
        newFormat = newFormatB
    ElseIf IsD_Size(sheetWidthInches, sheetHeightInches) Then
        newFormat = newFormatD
    Else
        Debug.Print "    ERROR: Unknown sheet size"
        ProcessSheet = False
        Exit Function
    End If
    
    ' Only change if format is different
    If currentFormat <> newFormat Then
        Debug.Print "    Changing format to: " & newFormat
        
        ' Get sheet properties
        vProps = swSheet.GetProperties()
        
        ' Apply new format using SetupSheet5
        ProcessSheet = swDraw.SetupSheet5(sheetName, CInt(vProps(0)), CInt(vProps(1)), CDbl(vProps(2)), CDbl(vProps(3)), CBool(vProps(4)), newFormat, CDbl(vProps(5)), CDbl(vProps(6)), swSheet.CustomPropertyView, REMOVE_MODIFIED_NOTES)
        
        If ProcessSheet Then
            If swSheet.ReloadTemplate(Not REMOVE_MODIFIED_NOTES) = swReloadTemplateResult_e.swReloadTemplate_Success Then
                swModel.EditRebuild3
                Debug.Print "    Successfully updated sheet format"
            Else
                Debug.Print "    ERROR: Failed to reload template"
                ProcessSheet = False
            End If
        Else
            Debug.Print "    ERROR: Failed to set template"
        End If
    Else
        Debug.Print "    Sheet already has correct format"
        ProcessSheet = True
    End If
    
    Exit Function
    
SheetError:
    Debug.Print "    ERROR: " & Err.Description
    ProcessSheet = False
End Function

Private Function BrowseForFolder(Optional prompt As String = "Select a folder") As String
    Dim shellApp As Object
    Dim folder As Object
    
    'Create a file browser window at the default folder
    Set shellApp = CreateObject("Shell.Application"). _
        BrowseForFolder(0, prompt, 0, 0)
    
    'Set the folder to the selected folder
    If Not shellApp Is Nothing Then
        Set folder = shellApp.Self
        'Return the path
        BrowseForFolder = folder.Path
    Else
        'Return empty string if cancelled
        BrowseForFolder = ""
    End If
    
    'Clean up
    Set folder = Nothing
    Set shellApp = Nothing
End Function

Private Function IsB_Size(width As Double, height As Double) As Boolean
    IsB_Size = (width = 11 And height = 8.5) Or (width = 17 And height = 11)
End Function

Private Function IsD_Size(width As Double, height As Double) As Boolean
    IsD_Size = (width = 22 And height = 17) Or (width = 34 And height = 22)
End Function

Private Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Private Function TerminateSolidWorksAndRestart() As Boolean
    On Error Resume Next
    
    Dim startTime As Double
    Debug.Print "WARNING: Unable to close documents normally. Attempting emergency termination..."
    
    ' First try to close all documents with force option
    swApp.CloseAllDocuments True
    DoEvents
    
    ' Check if that worked
    If swApp.GetDocumentCount > 0 Then
        ' If we still have documents, try the nuclear option
        Debug.Print "Using extreme measures to terminate SolidWorks..."
        
        ' Store the swApp application path before terminating
        Dim swPath As String
        swPath = swApp.GetExecutablePath
        
        ' Release all references
        Set swSheet = Nothing
        Set swDraw = Nothing
        Set swModel = Nothing
        
        ' Kill and restart SolidWorks
        swApp.ExitApp
        Set swApp = Nothing
        
        ' Wait for SolidWorks to fully exit
        startTime = Timer
        Do While Timer < startTime + 3
            DoEvents
        Loop
        
        ' Restart SolidWorks
        Set swApp = CreateObject("SldWorks.Application")
        swApp.Visible = True
        
        ' Wait for SolidWorks to initialize
        startTime = Timer
        Do While Timer < startTime + 5
            DoEvents
        Loop
        
        Debug.Print "SolidWorks has been forcefully restarted."
        TerminateSolidWorksAndRestart = True
    Else
        Debug.Print "All documents closed successfully."
        TerminateSolidWorksAndRestart = True
    End If
End Function

Private Sub CloseAllDocuments()
    On Error Resume Next
    
    Debug.Print "Performing final check for any open documents..."
    
    Dim remainingCount As Long
    remainingCount = swApp.GetDocumentCount
    
    If remainingCount > 0 Then
        Debug.Print "Found " & remainingCount & " documents still open at end of process"
        
        ' First try: Use QuitDoc on each open document
        Dim doc As Object
        For Each doc In swApp.GetDocuments
            Debug.Print "  Attempting to force quit: " & doc.GetTitle
            swApp.QuitDoc doc.GetTitle
            DoEvents
        Next
        
        ' Second try: CloseAllDocuments with force flag
        If swApp.GetDocumentCount > 0 Then
            Debug.Print "  Using CloseAllDocuments with force flag"
            swApp.CloseAllDocuments True
            DoEvents
        End If
        
        ' Third try: Extreme measure - terminate and restart SolidWorks
        remainingCount = swApp.GetDocumentCount
        If remainingCount > 0 Then
            If TerminateSolidWorksAndRestart() Then
                Debug.Print "SolidWorks has been restarted after forced termination"
            End If
        End If
        
        ' Final count
        remainingCount = swApp.GetDocumentCount
        If remainingCount > 0 Then
            Debug.Print "WARNING: " & remainingCount & " documents still could not be closed"
            
            ' Add a message to the failed files list
            If Not failedFiles Is Nothing Then
                failedFiles.Add "WARNING: " & remainingCount & " documents remained open after processing"
            End If
        Else
            Debug.Print "All documents successfully closed."
        End If
    Else
        Debug.Print "No documents remain open - all were closed successfully."
    End If
End Sub

Private Sub CleanupObjects()
    Set swSheet = Nothing
    Set swDraw = Nothing
    Set swModel = Nothing
    Set swApp = Nothing
    Set failedFiles = Nothing
End Sub



