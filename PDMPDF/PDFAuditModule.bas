Attribute VB_Name = "PDFAuditModule"
' ============================================================================
' MODULE: PDFAuditModule
' Description: Audit PDFs for all drawings in an assembly
' Reports missing PDFs and their approval states
' ============================================================================

Option Explicit

' Debug log file handle
Private debugLogFile As Integer
Private debugLogPath As String

' Main entry point for PDF audit
Public Sub AuditAssemblyPDFs()
    Dim swAssy As SldWorks.ModelDoc2
    Dim componentList As Collection
    Dim drawingList As Collection
    Dim auditRecords As Collection
    Dim i As Long

    ' Initialize debug logging
    InitializeDebugLog

    ' Clean up any leftover audit forms
    On Error Resume Next
    CleanupAuditForms
    On Error GoTo 0

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

    ' Connect to PDM
    If Not ConnectToPDMVault() Then
        MsgBox "Failed to connect to PDM vault. Cannot perform audit.", vbCritical
        Exit Sub
    End If

    ' Get all unique components from the assembly
    Debug.Print "=== PDF AUDIT: COLLECTING COMPONENTS ==="
    Set componentList = GetUniqueComponents(swAssy)
    Debug.Print "Found " & componentList.count & " unique components"

    If componentList.count = 0 Then
        MsgBox "No components found in this assembly.", vbInformation
        Exit Sub
    End If

    ' Find drawings for each component (excluding library items)
    Debug.Print "=== PDF AUDIT: FINDING DRAWINGS ==="
    Set drawingList = FindDrawingsForComponents(componentList)
    Debug.Print "Found " & drawingList.count & " drawings to audit"

    If drawingList.count = 0 Then
        MsgBox "No drawings found for any components in this assembly." & vbCrLf & vbCrLf & _
               "Note: Library items are automatically excluded.", vbInformation
        Exit Sub
    End If

    ' Perform audit on each drawing
    Debug.Print "=== PDF AUDIT: AUDITING DRAWINGS ==="
    Set auditRecords = New Collection

    Dim drawingPath As Variant
    For Each drawingPath In drawingList
        AuditSingleDrawing CStr(drawingPath), auditRecords
    Next drawingPath

    Debug.Print "=== PDF AUDIT: COMPLETE - " & auditRecords.count & " records ==="
    LogToFile "=== PDF AUDIT: COMPLETE - " & auditRecords.count & " records ==="

    ' Show results
    ShowAuditResults auditRecords

    ' Close debug log
    CloseDebugLog
End Sub

' Audit a single drawing for PDF status
' Each record is stored as a Dictionary object in the collection
Private Sub AuditSingleDrawing(drawingPath As String, auditRecords As Collection)
    Dim record As Object ' Dictionary
    Dim pdmDrawingFile As Object
    Dim fileName As String
    Dim drawingsFolder As String
    Dim obsoleteFolder As String
    Dim expectedPDFPath As String

    ' Create record as Dictionary
    Set record = CreateObject("Scripting.Dictionary")

    ' Initialize record
    record("DrawingPath") = drawingPath
    record("DrawingFileName") = Mid(drawingPath, InStrRev(drawingPath, "\") + 1)
    record("DrawingRevision") = ""
    record("DrawingState") = "Unknown"
    record("ModelPath") = ""
    record("ModelRevision") = ""
    record("ModelInSync") = True
    record("PDFPath") = ""
    record("PDFExists") = False
    record("PDFRevision") = ""
    record("PDFUpToDate") = False
    record("ErrorMessage") = ""

    Debug.Print "Auditing: " & record("DrawingFileName")
    LogToFile ""
    LogToFile "=========================================="
    LogToFile "Auditing: " & record("DrawingFileName")
    LogToFile "Full Path: " & drawingPath

    ' Get drawing state from PDM
    On Error Resume Next
    Set pdmDrawingFile = pdmVault.GetFileFromPath(drawingPath)
    If Err.Number <> 0 Or pdmDrawingFile Is Nothing Then
        record("ErrorMessage") = "Not found in PDM vault"
        Debug.Print "  ERROR: " & record("ErrorMessage")
        Err.Clear
        On Error GoTo 0
        auditRecords.Add record
        Exit Sub
    End If
    On Error GoTo 0

    ' Try to read state using CurrentState property
    On Error Resume Next
    Dim stateObj As Object
    Set stateObj = pdmDrawingFile.CurrentState
    If Err.Number = 0 And Not stateObj Is Nothing Then
        record("DrawingState") = stateObj.Name
        LogToFile "  State: " & record("DrawingState")
    Else
        LogToFile "  Could not read State (Error: " & Err.Number & ")"
        record("DrawingState") = "Unknown"
        LogToFile "  State: Unknown"
    End If
    Err.Clear

    ' Try to read revision from PDM
    record("DrawingRevision") = GetPDMVariable(pdmDrawingFile, drawingPath)
    If record("DrawingRevision") = "" Then
        LogToFile "  WARNING: No revision found in drawing data card"
    Else
        LogToFile "  Revision: " & record("DrawingRevision")
    End If
    On Error GoTo 0

    Debug.Print "  State: " & record("DrawingState")

    ' Get referenced model and check revision sync
    LogToFile "  Checking referenced model..."
    record("ModelPath") = GetReferencedModelPath(drawingPath)

    If record("ModelPath") <> "" Then
        LogToFile "  Model Path: " & record("ModelPath")
        record("ModelRevision") = GetModelRevisionFromPDM(record("ModelPath"))

        If record("ModelRevision") <> "" Then
            LogToFile "  Model Revision: " & record("ModelRevision")

            ' Compare drawing and model revisions
            If Trim(record("DrawingRevision")) <> "" And Trim(record("ModelRevision")) <> "" Then
                If UCase(Trim(record("DrawingRevision"))) = UCase(Trim(record("ModelRevision"))) Then
                    record("ModelInSync") = True
                    Debug.Print "  Model/Drawing Sync: MATCH (Drawing Rev " & record("DrawingRevision") & " = Model Rev " & record("ModelRevision") & ")"
                    LogToFile "  Model/Drawing Sync: MATCH (Drawing Rev " & record("DrawingRevision") & " = Model Rev " & record("ModelRevision") & ")"
                Else
                    record("ModelInSync") = False
                    Debug.Print "  Model/Drawing Sync: MISMATCH (Drawing Rev " & record("DrawingRevision") & " <> Model Rev " & record("ModelRevision") & ")"
                    LogToFile "  Model/Drawing Sync: MISMATCH (Drawing Rev " & record("DrawingRevision") & " <> Model Rev " & record("ModelRevision") & ")"
                End If
            Else
                ' Can't compare - assume in sync if at least one revision is missing
                record("ModelInSync") = True
                LogToFile "  Cannot compare model/drawing revisions (Drawing Rev: '" & record("DrawingRevision") & "', Model Rev: '" & record("ModelRevision") & "')"
            End If
        Else
            LogToFile "  WARNING: Could not read model revision"
            record("ModelInSync") = True ' Assume in sync if we can't read model revision
        End If
    Else
        LogToFile "  WARNING: No model reference found for drawing"
        record("ModelInSync") = True ' Assume in sync if no model reference
    End If

    ' Determine PDF save location
    fileName = Left(record("DrawingFileName"), InStrRev(record("DrawingFileName"), ".") - 1)

    If Not DeterminePDFSaveLocation(drawingPath, drawingsFolder, obsoleteFolder) Then
        record("ErrorMessage") = "Could not determine PDF save location"
        Debug.Print "  ERROR: " & record("ErrorMessage")
        LogToFile "  ERROR: " & record("ErrorMessage")
        auditRecords.Add record
        Exit Sub
    End If

    LogToFile "  PDF Save Location: " & drawingsFolder

    ' SIMPLIFIED: Just search for ANY PDF matching the base filename
    ' Don't worry about revision matching for now
    LogToFile "  Searching for ANY PDF with base name: " & fileName

    Dim foundPDF As String
    foundPDF = FindAnyPDF(drawingsFolder, fileName)

    If foundPDF <> "" Then
        record("PDFExists") = True
        record("PDFPath") = foundPDF
        record("PDFRevision") = ExtractRevisionFromPDF(Mid(foundPDF, InStrRev(foundPDF, "\") + 1), fileName)
        record("ErrorMessage") = "" ' Clear any error message since we found a PDF

        Debug.Print "  PDF: FOUND - " & foundPDF
        LogToFile "  PDF: FOUND - " & foundPDF
        LogToFile "  PDF Revision: " & record("PDFRevision")

        ' Compare revisions if we have both
        If Trim(record("DrawingRevision")) <> "" And Trim(record("PDFRevision")) <> "" Then
            If UCase(Trim(record("DrawingRevision"))) = UCase(Trim(record("PDFRevision"))) Then
                record("PDFUpToDate") = True
                Debug.Print "  Revision Match: Drawing Rev " & record("DrawingRevision") & " = PDF Rev " & record("PDFRevision")
                LogToFile "  Revision Match: Drawing Rev " & record("DrawingRevision") & " = PDF Rev " & record("PDFRevision")
            Else
                record("PDFUpToDate") = False
                Debug.Print "  Revision Mismatch: Drawing Rev " & record("DrawingRevision") & " <> PDF Rev " & record("PDFRevision")
                LogToFile "  Revision Mismatch: Drawing Rev " & record("DrawingRevision") & " <> PDF Rev " & record("PDFRevision")
            End If
        Else
            ' Can't compare revisions - mark as up to date since PDF exists
            record("PDFUpToDate") = True
            LogToFile "  Cannot compare revisions (Drawing Rev: '" & record("DrawingRevision") & "', PDF Rev: '" & record("PDFRevision") & "')"
        End If
    Else
        record("PDFExists") = False
        record("PDFPath") = ""
        record("PDFRevision") = ""
        record("PDFUpToDate") = False
        record("ErrorMessage") = "No PDF found"
        Debug.Print "  PDF: NOT FOUND (searched in " & drawingsFolder & ")"
        LogToFile "  PDF: NOT FOUND (searched in " & drawingsFolder & ")"
    End If

    ' Add record to collection
    auditRecords.Add record
End Sub

' Get revision variable from PDM file using GetEnumerator
Private Function GetPDMVariable(pdmFile As Object, filePath As String) As String
    Dim pdmFolder As Object
    Dim enumVars As Object
    Dim varNames As Variant
    Dim varName As Variant
    Dim revisionValue As Variant

    GetPDMVariable = ""

    On Error Resume Next

    ' Get the folder object
    Set pdmFolder = pdmVault.GetFolderFromPath(Left(filePath, InStrRev(filePath, "\") - 1))

    If pdmFolder Is Nothing Then
        LogToFile "    GetPDMVariable: Could not get folder from path"
        Exit Function
    End If

    LogToFile "    GetPDMVariable: File ID = " & pdmFile.ID & ", Folder ID = " & pdmFolder.ID
    LogToFile "    GetPDMVariable: TypeName = " & TypeName(pdmFile)

    ' Get enumerator for the file - GetEnumerator takes a position as Long parameter
    ' Pass 0 to get the first configuration's variables
    Set enumVars = pdmFile.GetEnumerator(0&)

    If Err.Number <> 0 Or enumVars Is Nothing Then
        LogToFile "    GetPDMVariable: GetEnumerator(0) failed (Err=" & Err.Number & "), trying with folder ID..."
        Err.Clear

        ' Try with folder ID as parameter
        Set enumVars = pdmFile.GetEnumerator(pdmFolder.ID)

        If Err.Number <> 0 Or enumVars Is Nothing Then
            LogToFile "    GetPDMVariable: GetEnumerator(FolderID) also failed (Err=" & Err.Number & ")"
            Err.Clear
            Exit Function
        End If
    End If

    LogToFile "    GetPDMVariable: Got enumerator successfully, TypeName = " & TypeName(enumVars)

    ' Try different revision variable names
    varNames = Array("Revision", "Rev", "REV", "Drawing Revision", "DWG_REV")

    For Each varName In varNames
        Err.Clear

        ' Use GetVar method on the enumerator
        revisionValue = enumVars.GetVar(CStr(varName))

        If Err.Number = 0 Then
            If Not IsEmpty(revisionValue) And Not IsNull(revisionValue) Then
                If Trim(CStr(revisionValue)) <> "" Then
                    GetPDMVariable = CStr(revisionValue)
                    LogToFile "    GetPDMVariable: SUCCESS - Found '" & varName & "' = '" & revisionValue & "'"
                    Exit Function
                Else
                    LogToFile "    GetPDMVariable: Variable '" & varName & "' exists but is empty"
                End If
            Else
                LogToFile "    GetPDMVariable: Variable '" & varName & "' is NULL or Empty"
            End If
        Else
            LogToFile "    GetPDMVariable: GetVar('" & varName & "') failed (Err=" & Err.Number & ")"
        End If
    Next varName

    LogToFile "    GetPDMVariable: FAILED - Could not find revision in any variable"

    On Error GoTo 0
End Function

' Find ANY PDF matching the base filename (regardless of revision)
Private Function FindAnyPDF(drawingsFolder As String, baseName As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileName As String
    Dim baseLower As String

    FindAnyPDF = ""
    baseLower = LCase(baseName)

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If Not fso.folderExists(drawingsFolder) Then
        LogToFile "    FindAnyPDF: Folder does not exist: " & drawingsFolder
        Exit Function
    End If

    Set folder = fso.GetFolder(drawingsFolder)
    If Err.Number <> 0 Then
        LogToFile "    FindAnyPDF: Cannot access folder: " & drawingsFolder
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    LogToFile "    FindAnyPDF: Scanning folder for PDFs matching '" & baseName & "'"

    ' Loop through all files in the folder
    For Each file In folder.Files
        fileName = file.Name

        ' Check if this is a PDF matching our base name
        If LCase(Right(fileName, 4)) = ".pdf" Then
            ' Check if the filename starts with our base name
            If InStr(1, LCase(fileName), baseLower, vbTextCompare) = 1 Then
                FindAnyPDF = file.Path
                LogToFile "    FindAnyPDF: MATCH FOUND - " & fileName
                Exit Function
            End If
        End If
    Next file

    LogToFile "    FindAnyPDF: No PDF found matching base name '" & baseName & "'"
End Function

' Find an older revision PDF in the drawings folder
Private Function FindOlderPDF(drawingsFolder As String, baseName As String, currentRev As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileName As String
    Dim maxRev As Long
    Dim maxRevFile As String
    Dim fileRev As String
    Dim baseLower As String

    FindOlderPDF = ""
    maxRev = -1
    baseLower = LCase(baseName)

    ' Use FileSystemObject to search folder
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If Not fso.folderExists(drawingsFolder) Then
        Debug.Print "  FindOlderPDF: Folder does not exist: " & drawingsFolder
        Exit Function
    End If

    Set folder = fso.GetFolder(drawingsFolder)
    If Err.Number <> 0 Then
        Debug.Print "  FindOlderPDF: Cannot access folder: " & drawingsFolder
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Loop through all files in the folder
    For Each file In folder.Files
        fileName = file.Name

        ' Check if this is a PDF matching our base name pattern
        If LCase(Right(fileName, 4)) = ".pdf" Then
            If InStr(1, LCase(fileName), baseLower & "_rev", vbTextCompare) > 0 Then
                ' Extract revision from filename
                fileRev = ExtractRevisionFromPDF(fileName, baseName)

                If fileRev <> "" And IsNumeric(fileRev) Then
                    If CLng(fileRev) > maxRev Then
                        maxRev = CLng(fileRev)
                        maxRevFile = file.Path
                    End If
                End If
            End If
        End If
    Next file

    ' Return the highest revision found (even if it's older than current)
    If maxRev >= 0 Then
        FindOlderPDF = maxRevFile
        Debug.Print "  FindOlderPDF: Found older PDF Rev " & maxRev & ": " & maxRevFile
    Else
        Debug.Print "  FindOlderPDF: No PDFs found matching pattern: " & baseName & "_Rev*.pdf"
    End If
End Function

' Show audit results using dynamic form
Private Sub ShowAuditResults(auditRecords As Collection)
    ' Create and show dynamic form (with text file fallback on error)
    AuditResultsFormBuilder.ShowAuditResults auditRecords
End Sub

' OLD CODE - Fallback text export (kept for reference)
Private Sub ShowAuditResults_OLD(auditRecords As Collection)
    Dim filePath As String
    Dim fileNum As Integer
    Dim record As Object
    Dim missingCount As Long, outOfDateCount As Long, upToDateCount As Long, errorCount As Long
    Dim stateGroups As Object
    Dim stateKey As Variant
    Dim stateRec As Object

    ' Create file on desktop
    filePath = Environ("USERPROFILE") & "\Desktop\PDF_Audit_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    fileNum = FreeFile

    ' Count records
    For Each record In auditRecords
        If record("ErrorMessage") <> "" Then
            errorCount = errorCount + 1
        ElseIf Not record("PDFExists") Then
            missingCount = missingCount + 1
        ElseIf Not record("PDFUpToDate") Then
            outOfDateCount = outOfDateCount + 1
        Else
            upToDateCount = upToDateCount + 1
        End If
    Next record

    ' Write to file
    Open filePath For Output As #fileNum

    Print #fileNum, "PDF AUDIT RESULTS - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNum, String(100, "=")
    Print #fileNum, ""
    Print #fileNum, "SUMMARY:"
    Print #fileNum, "Total Drawings: " & auditRecords.count
    Print #fileNum, "  Up-to-Date PDFs: " & upToDateCount
    Print #fileNum, "  Out-of-Date PDFs: " & outOfDateCount
    Print #fileNum, "  Missing PDFs: " & missingCount
    Print #fileNum, "  Errors: " & errorCount
    Print #fileNum, ""
    Print #fileNum, String(100, "=")
    Print #fileNum, ""

    ' Missing PDFs grouped by state
    If missingCount > 0 Then
        Print #fileNum, "MISSING PDFs (by State):"
        Print #fileNum, String(100, "-")
        Print #fileNum, "Drawing" & vbTab & "Revision" & vbTab & "State"

        Set stateGroups = CreateObject("Scripting.Dictionary")

        For Each record In auditRecords
            If record("ErrorMessage") = "" And Not record("PDFExists") Then
                If Not stateGroups.exists(record("DrawingState")) Then
                    stateGroups.Add record("DrawingState"), New Collection
                End If
                stateGroups(record("DrawingState")).Add record
            End If
        Next record

        For Each stateKey In stateGroups.Keys
            Print #fileNum, ""
            Print #fileNum, "[" & stateKey & "]:"
            For Each stateRec In stateGroups(stateKey)
                Print #fileNum, "  " & stateRec("DrawingFileName") & vbTab & stateRec("DrawingRevision") & vbTab & stateRec("DrawingState")
            Next stateRec
        Next stateKey
        Print #fileNum, ""
    End If

    ' Out-of-date PDFs
    If outOfDateCount > 0 Then
        Print #fileNum, String(100, "=")
        Print #fileNum, "OUT-OF-DATE PDFs:"
        Print #fileNum, String(100, "-")
        Print #fileNum, "Drawing" & vbTab & "Drawing Rev" & vbTab & "PDF Rev" & vbTab & "State"

        For Each record In auditRecords
            If record("ErrorMessage") = "" And record("PDFExists") And Not record("PDFUpToDate") Then
                Print #fileNum, record("DrawingFileName") & vbTab & record("DrawingRevision") & vbTab & record("PDFRevision") & vbTab & record("DrawingState")
            End If
        Next record
        Print #fileNum, ""
    End If

    ' Errors
    If errorCount > 0 Then
        Print #fileNum, String(100, "=")
        Print #fileNum, "ERRORS:"
        Print #fileNum, String(100, "-")

        For Each record In auditRecords
            If record("ErrorMessage") <> "" Then
                Print #fileNum, record("DrawingFileName") & vbTab & record("ErrorMessage")
            End If
        Next record
        Print #fileNum, ""
    End If

    ' Up-to-date (if not too many)
    If upToDateCount > 0 And upToDateCount < 100 Then
        Print #fileNum, String(100, "=")
        Print #fileNum, "UP-TO-DATE PDFs: (" & upToDateCount & " total)"
        Print #fileNum, String(100, "-")
        Print #fileNum, "Drawing" & vbTab & "Revision" & vbTab & "State"

        For Each record In auditRecords
            If record("ErrorMessage") = "" And record("PDFExists") And record("PDFUpToDate") Then
                Print #fileNum, record("DrawingFileName") & vbTab & record("DrawingRevision") & vbTab & record("DrawingState")
            End If
        Next record
    End If

    Close #fileNum

    ' Open the file in Notepad
    Shell "notepad.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus

    MsgBox "Audit complete!" & vbCrLf & vbCrLf & _
           "Total: " & auditRecords.count & vbCrLf & _
           "Missing: " & missingCount & vbCrLf & _
           "Out-of-Date: " & outOfDateCount & vbCrLf & _
           "Up-to-Date: " & upToDateCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & vbCrLf & _
           "Results saved to:" & vbCrLf & filePath, vbInformation, "PDF Audit Complete"
End Sub

' Clean up any leftover AuditForm UserForms
Private Sub CleanupAuditForms()
    On Error Resume Next
    Dim vbProj As Object
    Dim comp As Object
    Dim i As Long

    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then Exit Sub

    For i = vbProj.VBComponents.count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        If Not comp Is Nothing Then
            If Left(comp.Name, 9) = "AuditForm" Then
                vbProj.VBComponents.Remove comp
                DoEvents
            End If
        End If
    Next i
End Sub

' ============================================================================
' DEBUG LOGGING FUNCTIONS
' ============================================================================

' Initialize debug log file
Private Sub InitializeDebugLog()
    Dim fso As Object
    Dim macroFolder As String
    Dim logsFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Use the macro's directory (hardcoded for SolidWorks VBA)
    logsFolder = "C:\Users\dlebel\OneDrive - Nordic Minesteel Technologies Inc\Documents\new pdm pdf\DebugLogs\"

    ' Create DebugLogs subfolder if it doesn't exist
    If Not fso.folderExists(logsFolder) Then
        fso.CreateFolder logsFolder
    End If

    ' Create log file with timestamp
    debugLogPath = logsFolder & "PDFAudit_" & Format(Now, "yyyymmdd_hhnnss") & ".log"
    debugLogFile = FreeFile

    Open debugLogPath For Output As #debugLogFile

    Print #debugLogFile, "PDF AUDIT DEBUG LOG - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #debugLogFile, String(80, "=")
    Print #debugLogFile, ""

    Debug.Print "Debug log created: " & debugLogPath
End Sub

' Write to debug log
Private Sub LogToFile(msg As String)
    On Error Resume Next
    If debugLogFile > 0 Then
        Print #debugLogFile, msg
    End If
    On Error GoTo 0
End Sub

' Get the referenced model path from a drawing file via PDM
' Returns the path to the model, or empty string if not found
Private Function GetReferencedModelPath(drawingPath As String) As String
    Dim pdmFile As Object
    Dim references As Object
    Dim refFile As Object
    Dim i As Long
    Dim refPath As String

    GetReferencedModelPath = ""

    On Error Resume Next

    ' Get the drawing file from PDM (IEdmFile5)
    Set pdmFile = pdmVault.GetFileFromPath(drawingPath)
    If Err.Number <> 0 Or pdmFile Is Nothing Then
        LogToFile "    GetReferencedModelPath: Could not get drawing file from PDM: " & Err.Description
        Err.Clear
        Exit Function
    End If

    LogToFile "    GetReferencedModelPath: Got drawing file from PDM"

    ' Get first file reference position using IEdmFile10 interface
    Dim pdmPos As Object
    Set pdmPos = pdmFile.GetFirstFileReference(, , , True)
    If Err.Number <> 0 Then
        LogToFile "    GetReferencedModelPath: Error calling GetFirstFileReference: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If pdmPos Is Nothing Then
        LogToFile "    GetReferencedModelPath: GetFirstFileReference returned Nothing (no references)"
        Exit Function
    End If

    LogToFile "    GetReferencedModelPath: Got first reference position, looping through references..."

    ' Loop through references to find a model file
    Dim pdmRef As Object
    Dim refConfig As String
    Dim refExt As String
    Dim refCount As Long

    refCount = 0
    Do While Not pdmPos Is Nothing
        refPath = ""
        refConfig = ""
        Set pdmRef = pdmFile.GetNextFileReference(pdmPos, refPath, refConfig)

        If Err.Number <> 0 Then
            LogToFile "    GetReferencedModelPath: Error in GetNextFileReference: " & Err.Description
            Err.Clear
            Exit Do
        End If

        If Not pdmRef Is Nothing And refPath <> "" Then
            refCount = refCount + 1
            refExt = LCase(Right(refPath, 7))
            LogToFile "    GetReferencedModelPath: Reference #" & refCount & ": " & refPath

            ' Check if this is a model file (not another drawing)
            If refExt = ".sldprt" Or refExt = ".sldasm" Then
                GetReferencedModelPath = refPath
                LogToFile "    GetReferencedModelPath: SUCCESS - Found model reference: " & refPath
                Exit Function
            End If
        End If
    Loop

    If refCount = 0 Then
        LogToFile "    GetReferencedModelPath: No references found at all"
    Else
        LogToFile "    GetReferencedModelPath: Checked " & refCount & " references, no model file found"
    End If
    On Error GoTo 0
End Function

' Get the revision of a model file from PDM
Private Function GetModelRevisionFromPDM(modelPath As String) As String
    Dim pdmFile As Object

    GetModelRevisionFromPDM = ""

    If modelPath = "" Then Exit Function

    On Error Resume Next
    Set pdmFile = pdmVault.GetFileFromPath(modelPath)
    If Err.Number <> 0 Or pdmFile Is Nothing Then
        LogToFile "    GetModelRevisionFromPDM: Could not get model file from PDM: " & modelPath
        Err.Clear
        Exit Function
    End If

    ' Use the same GetPDMVariable function we already have
    GetModelRevisionFromPDM = GetPDMVariable(pdmFile, modelPath)

    If GetModelRevisionFromPDM <> "" Then
        LogToFile "    GetModelRevisionFromPDM: Model revision = " & GetModelRevisionFromPDM
    Else
        LogToFile "    GetModelRevisionFromPDM: No revision found for model"
    End If

    On Error GoTo 0
End Function

' Close debug log
Private Sub CloseDebugLog()
    On Error Resume Next
    If debugLogFile > 0 Then
        Print #debugLogFile, ""
        Print #debugLogFile, String(80, "=")
        Print #debugLogFile, "LOG END - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Close #debugLogFile
        debugLogFile = 0

        ' Log file saved (don't auto-open)
        Debug.Print "Log file saved: " & debugLogPath
    End If
    On Error GoTo 0
End Sub
