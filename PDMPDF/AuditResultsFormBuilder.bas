Attribute VB_Name = "AuditResultsFormBuilder"
Option Explicit

' ============================================
' AuditResultsFormBuilder - Creates audit results dialog at runtime
' No .frm/.frx files needed - form built entirely in code
' ============================================

' Module-level variables
Private m_FormCounter As Long
Private m_AuditRecords As Collection
Public currentAuditRecords As Collection ' Accessible by button click handlers

' Main function to show the audit results dialog
Public Sub ShowAuditResults(auditRecords As Collection)
    Set m_AuditRecords = auditRecords
    Set currentAuditRecords = auditRecords ' Make accessible to button handlers

    Dim vbProj As Object
    Dim vbComp As Object
    Dim formName As String
    Dim frm As Object
    Dim errMsg As String

    Debug.Print "AuditResultsFormBuilder: Starting..."
    On Error GoTo FormError

    ' Get VBE from Application object (works in SolidWorks)
    Debug.Print "AuditResultsFormBuilder: Getting VBE..."
    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then
        errMsg = "Could not get ActiveVBProject"
        GoTo FormError
    End If
    Debug.Print "AuditResultsFormBuilder: Got VBProject"

    ' Clean up any old audit forms
    CleanupOldForms vbProj

    ' Use unique form name
    m_FormCounter = m_FormCounter + 1
    formName = "AuditForm" & m_FormCounter

    ' Remove existing component if it exists
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents(formName)
    If Err.Number = 0 And Not existingComp Is Nothing Then
        vbProj.VBComponents.Remove existingComp
        DoEvents
    End If
    Err.Clear
    On Error GoTo FormError

    ' Create new form (3 = vbext_ct_MSForm)
    Debug.Print "AuditResultsFormBuilder: Creating form component..."
    On Error Resume Next
    Err.Clear
    Set vbComp = vbProj.VBComponents.Add(3)

    If Err.Number <> 0 Then
        errMsg = "VBComponents.Add failed: " & Err.Description
        Debug.Print "AuditResultsFormBuilder: ERROR - " & errMsg
        On Error GoTo FormError
        GoTo FormError
    End If
    On Error GoTo FormError

    If vbComp Is Nothing Then
        errMsg = "VBComponents.Add returned Nothing"
        Debug.Print "AuditResultsFormBuilder: ERROR - " & errMsg
        GoTo FormError
    End If
    Debug.Print "AuditResultsFormBuilder: Form component created"

    ' Rename form (allow some retries)
    DoEvents
    On Error Resume Next
    vbComp.Name = formName
    If Err.Number <> 0 Then
        DoEvents
        Err.Clear
        vbComp.Name = formName
    End If
    If Err.Number <> 0 Then
        formName = vbComp.Name ' Use default name if rename fails
        Err.Clear
    End If
    On Error GoTo FormError

    ' Set form properties
    With vbComp
        .Properties("Caption") = "PDF Audit Results"
        .Properties("Width") = 600
        .Properties("Height") = 400
    End With

    ' Get the form designer object
    Set frm = vbComp.designer
    If frm Is Nothing Then
        errMsg = "Could not get form Designer"
        GoTo FormError
    End If

    ' Build the form layout
    Debug.Print "AuditResultsFormBuilder: Building form layout..."
    BuildFormLayout frm, auditRecords
    Debug.Print "AuditResultsFormBuilder: Layout built"

    ' Add event handlers
    Debug.Print "AuditResultsFormBuilder: Adding event code..."
    AddFormCode vbComp, formName
    Debug.Print "AuditResultsFormBuilder: Event code added"

    DoEvents

    ' Show the form
    Debug.Print "AuditResultsFormBuilder: Showing form..."
    On Error Resume Next
    Set frm = VBA.UserForms.Add(formName)
    If Err.Number <> 0 Or frm Is Nothing Then
        errMsg = "Failed to add form to UserForms: " & Err.Description
        Debug.Print "AuditResultsFormBuilder: ERROR - " & errMsg
        On Error GoTo FormError
        GoTo FormError
    End If

    frm.Show 1  ' 1 = modal (blocks until closed)
    If Err.Number <> 0 Then
        errMsg = "Failed to show form: " & Err.Description
        Debug.Print "AuditResultsFormBuilder: ERROR - " & errMsg
        On Error GoTo FormError
        GoTo FormError
    End If
    On Error GoTo 0

    Debug.Print "AuditResultsFormBuilder: Form shown successfully!"

    ' Don't remove the form component yet - let user close it
    ' Clean up will happen on next run

    Exit Sub

FormError:
    If errMsg = "" Then errMsg = Err.Description
    Debug.Print "AuditResultsFormBuilder: FORM ERROR - " & errMsg
    MsgBox "Error creating audit form: " & errMsg & vbCrLf & vbCrLf & _
           "Falling back to text file export...", vbExclamation

    ' Fallback to text file export
    ExportToTextFile auditRecords
End Sub

' Build the form layout and populate with data
Private Sub BuildFormLayout(designer As Object, auditRecords As Collection)
    On Error GoTo LayoutError

    ' Create TextBox for results
    Dim txtResults As Object
    Set txtResults = designer.Controls.Add("Forms.TextBox.1", "txtResults")
    With txtResults
        .Left = 10
        .Top = 10
        .Width = 580
        .Height = 320
        .MultiLine = True
        .ScrollBars = 3 ' Both
        .Locked = True
        .text = GenerateAuditReport(auditRecords)
    End With

    ' Create Export button
    Dim btnExport As Object
    Set btnExport = designer.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Left = 10
        .Top = 340
        .Width = 100
        .Height = 30
        .Caption = "Export to Desktop"
    End With

    ' Create Copy button
    Dim btnCopy As Object
    Set btnCopy = designer.Controls.Add("Forms.CommandButton.1", "btnCopy")
    With btnCopy
        .Left = 120
        .Top = 340
        .Width = 80
        .Height = 30
        .Caption = "Copy All"
    End With

    ' Create "Create Missing PDFs" button (only if there are missing PDFs)
    Dim btnCreatePDFs As Object
    Debug.Print "BuildFormLayout: Checking if we should create the button..."
    If HasMissingPDFs(auditRecords) Then
        Debug.Print "BuildFormLayout: Creating 'Create Missing PDFs' button"
        Set btnCreatePDFs = designer.Controls.Add("Forms.CommandButton.1", "btnCreatePDFs")
        With btnCreatePDFs
            .Left = 10
            .Top = 378
            .Width = 120
            .Height = 30
            .Caption = "Create Missing PDFs"
        End With
        Debug.Print "BuildFormLayout: Button created successfully"
    Else
        Debug.Print "BuildFormLayout: No missing PDFs found, skipping button"
    End If

    ' Create Close button
    Dim btnClose As Object
    Set btnClose = designer.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Left = 210
        .Top = 340
        .Width = 80
        .Height = 30
        .Caption = "Close"
    End With

    Exit Sub

LayoutError:
    Err.Raise Err.Number, "BuildFormLayout", "Error building form: " & Err.Description
End Sub

' Add VBA code to the form for button handlers
Private Sub AddFormCode(vbComp As Object, formName As String)
    On Error GoTo CodeError

    Dim code As String

    code = "Option Explicit" & vbCrLf & vbCrLf

    ' Export button handler
    code = code & "Private Sub btnExport_Click()" & vbCrLf
    code = code & "    Dim filePath As String" & vbCrLf
    code = code & "    Dim fileNum As Integer" & vbCrLf
    code = code & "    filePath = Environ(""USERPROFILE"") & ""\Desktop\PDF_Audit_"" & Format(Now, ""yyyymmdd_hhnnss"") & "".txt""" & vbCrLf
    code = code & "    fileNum = FreeFile" & vbCrLf
    code = code & "    Open filePath For Output As #fileNum" & vbCrLf
    code = code & "    Print #fileNum, Me.txtResults.Text" & vbCrLf
    code = code & "    Close #fileNum" & vbCrLf
    code = code & "    Shell ""notepad.exe "" & Chr(34) & filePath & Chr(34), vbNormalFocus" & vbCrLf
    code = code & "    MsgBox ""Exported to: "" & vbCrLf & filePath, vbInformation" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Copy button handler
    code = code & "Private Sub btnCopy_Click()" & vbCrLf
    code = code & "    Dim dataObj As Object" & vbCrLf
    code = code & "    Set dataObj = CreateObject(""new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"")" & vbCrLf
    code = code & "    dataObj.SetText Me.txtResults.Text" & vbCrLf
    code = code & "    dataObj.PutInClipboard" & vbCrLf
    code = code & "    MsgBox ""Text copied to clipboard!"", vbInformation" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Close button handler
    code = code & "Private Sub btnClose_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' Create Missing PDFs button handler
    code = code & "Private Sub btnCreatePDFs_Click()" & vbCrLf
    code = code & "    AuditResultsFormBuilder.btnCreatePDFs_Click" & vbCrLf
    code = code & "End Sub" & vbCrLf

    ' Add code to form module
    vbComp.CodeModule.AddFromString code

    Exit Sub

CodeError:
    Err.Raise Err.Number, "AddFormCode", "Error adding form code: " & Err.Description
End Sub

' Generate the audit report text
Private Function GenerateAuditReport(auditRecords As Collection) As String
    Dim output As String
    Dim record As Object
    Dim missingCount As Long, outOfDateCount As Long, upToDateCount As Long, errorCount As Long
    Dim modelMismatchCount As Long
    Dim stateGroups As Object
    Dim stateKey As Variant
    Dim stateRec As Object

    ' Count records
    For Each record In auditRecords
        ' Check model sync status first
        If Not record("ModelInSync") Then
            modelMismatchCount = modelMismatchCount + 1
        End If

        ' Check if PDF exists first (most important)
        If Not record("PDFExists") Then
            missingCount = missingCount + 1
        ElseIf Not record("PDFUpToDate") Then
            outOfDateCount = outOfDateCount + 1
        Else
            upToDateCount = upToDateCount + 1
        End If

        ' Separately track real errors (not just "No PDF found")
        If record("ErrorMessage") <> "" And record("ErrorMessage") <> "No PDF found" Then
            errorCount = errorCount + 1
        End If
    Next record

    ' Build report
    output = "PDF AUDIT RESULTS - " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    output = output & String(100, "=") & vbCrLf & vbCrLf

    output = output & "SUMMARY:" & vbCrLf
    output = output & "Total Drawings: " & auditRecords.count & vbCrLf
    output = output & "  Model/Drawing Mismatches: " & modelMismatchCount & vbCrLf
    output = output & "  Up-to-Date PDFs: " & upToDateCount & vbCrLf
    output = output & "  Out-of-Date PDFs: " & outOfDateCount & vbCrLf
    output = output & "  Missing PDFs: " & missingCount & vbCrLf
    output = output & "  Errors: " & errorCount & vbCrLf
    output = output & vbCrLf & String(100, "=") & vbCrLf & vbCrLf

    ' Model/Drawing Mismatches
    If modelMismatchCount > 0 Then
        output = output & "MODEL/DRAWING REVISION MISMATCHES:" & vbCrLf
        output = output & String(100, "-") & vbCrLf
        output = output & "Drawing" & vbTab & "Drawing Rev" & vbTab & "Model Rev" & vbTab & "State" & vbCrLf

        For Each record In auditRecords
            If Not record("ModelInSync") Then
                output = output & record("DrawingFileName") & vbTab & _
                                record("DrawingRevision") & vbTab & _
                                record("ModelRevision") & vbTab & _
                                record("DrawingState") & vbCrLf
            End If
        Next record
        output = output & vbCrLf
        output = output & "ACTION REQUIRED: These drawings need model revision sync before PDF generation." & vbCrLf
        output = output & vbCrLf & String(100, "=") & vbCrLf & vbCrLf
    End If

    ' Missing PDFs grouped by state
    If missingCount > 0 Then
        output = output & "MISSING PDFs (by State):" & vbCrLf
        output = output & String(100, "-") & vbCrLf
        output = output & "Drawing" & vbTab & "Revision" & vbTab & "State" & vbCrLf

        Set stateGroups = CreateObject("Scripting.Dictionary")

        For Each record In auditRecords
            ' Include all missing PDFs (ignore the "No PDF found" error message)
            If Not record("PDFExists") Then
                If Not stateGroups.exists(record("DrawingState")) Then
                    stateGroups.Add record("DrawingState"), New Collection
                End If
                stateGroups(record("DrawingState")).Add record
            End If
        Next record

        For Each stateKey In stateGroups.Keys
            output = output & vbCrLf & "[" & stateKey & "]:" & vbCrLf
            For Each stateRec In stateGroups(stateKey)
                output = output & "  " & stateRec("DrawingFileName") & vbTab & _
                                stateRec("DrawingRevision") & vbTab & _
                                stateRec("DrawingState") & vbCrLf
            Next stateRec
        Next stateKey
        output = output & vbCrLf
    End If

    ' Out-of-date PDFs
    If outOfDateCount > 0 Then
        output = output & String(100, "=") & vbCrLf
        output = output & "OUT-OF-DATE PDFs:" & vbCrLf
        output = output & String(100, "-") & vbCrLf
        output = output & "Drawing" & vbTab & "Drawing Rev" & vbTab & "PDF Rev" & vbTab & "State" & vbCrLf

        For Each record In auditRecords
            If record("ErrorMessage") = "" And record("PDFExists") And Not record("PDFUpToDate") Then
                output = output & record("DrawingFileName") & vbTab & _
                                record("DrawingRevision") & vbTab & _
                                record("PDFRevision") & vbTab & _
                                record("DrawingState") & vbCrLf
            End If
        Next record
        output = output & vbCrLf
    End If

    ' Errors
    If errorCount > 0 Then
        output = output & String(100, "=") & vbCrLf
        output = output & "ERRORS:" & vbCrLf
        output = output & String(100, "-") & vbCrLf

        For Each record In auditRecords
            If record("ErrorMessage") <> "" Then
                output = output & record("DrawingFileName") & vbTab & record("ErrorMessage") & vbCrLf
            End If
        Next record
        output = output & vbCrLf
    End If

    ' Up-to-date (if not too many)
    If upToDateCount > 0 And upToDateCount < 100 Then
        output = output & String(100, "=") & vbCrLf
        output = output & "UP-TO-DATE PDFs: (" & upToDateCount & " total)" & vbCrLf
        output = output & String(100, "-") & vbCrLf
        output = output & "Drawing" & vbTab & "Revision" & vbTab & "State" & vbCrLf

        For Each record In auditRecords
            If record("ErrorMessage") = "" And record("PDFExists") And record("PDFUpToDate") Then
                output = output & record("DrawingFileName") & vbTab & _
                                record("DrawingRevision") & vbTab & _
                                record("DrawingState") & vbCrLf
            End If
        Next record
    End If

    GenerateAuditReport = output
End Function

' Clean up old audit forms
Private Sub CleanupOldForms(vbProj As Object)
    On Error Resume Next
    Dim comp As Object
    Dim i As Long

    For i = vbProj.VBComponents.count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        If Left(comp.Name, 9) = "AuditForm" Then
            vbProj.VBComponents.Remove comp
            DoEvents
        End If
    Next i
End Sub

' Fallback: Export to text file if form creation fails
Public Sub ExportToTextFile(auditRecords As Collection)
    Dim filePath As String
    Dim fileNum As Integer
    Dim record As Object
    Dim missingCount As Long, outOfDateCount As Long, upToDateCount As Long, errorCount As Long

    ' Count records
    For Each record In auditRecords
        ' Check if PDF exists first (most important)
        If Not record("PDFExists") Then
            missingCount = missingCount + 1
        ElseIf Not record("PDFUpToDate") Then
            outOfDateCount = outOfDateCount + 1
        Else
            upToDateCount = upToDateCount + 1
        End If

        ' Separately track real errors (not just "No PDF found")
        If record("ErrorMessage") <> "" And record("ErrorMessage") <> "No PDF found" Then
            errorCount = errorCount + 1
        End If
    Next record

    ' Create file on desktop
    filePath = Environ("USERPROFILE") & "\Desktop\PDF_Audit_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    fileNum = FreeFile

    Open filePath For Output As #fileNum
    Print #fileNum, GenerateAuditReport(auditRecords)
    Close #fileNum

    ' Open in notepad
    Shell "notepad.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus

    MsgBox "Audit complete!" & vbCrLf & vbCrLf & _
           "Total: " & auditRecords.count & vbCrLf & _
           "Missing: " & missingCount & vbCrLf & _
           "Out-of-Date: " & outOfDateCount & vbCrLf & _
           "Up-to-Date: " & upToDateCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & vbCrLf & _
           "Results saved to:" & vbCrLf & filePath, vbInformation, "PDF Audit Complete"
End Sub

' ============================================================================
' HELPER FUNCTIONS FOR CREATE MISSING PDFs BUTTON
' ============================================================================

' Check if there are any missing PDFs in the audit results
Private Function HasMissingPDFs(auditRecords As Collection) As Boolean
    Dim record As Object
    Dim missingCount As Long
    HasMissingPDFs = False

    Debug.Print "HasMissingPDFs: Checking audit records..."

    If auditRecords Is Nothing Then
        Debug.Print "HasMissingPDFs: auditRecords is Nothing!"
        Exit Function
    End If

    Debug.Print "HasMissingPDFs: Found " & auditRecords.count & " records"

    For Each record In auditRecords
        If Not record("PDFExists") Then
            missingCount = missingCount + 1
        End If
    Next record

    Debug.Print "HasMissingPDFs: Found " & missingCount & " missing PDFs"

    If missingCount > 0 Then
        HasMissingPDFs = True
    End If
End Function

' Handler for "Create Missing PDFs" button click
Public Sub btnCreatePDFs_Click()
    Dim missingCount As Long
    Dim record As Object
    Dim response As VbMsgBoxResult

    ' Count missing PDFs
    For Each record In currentAuditRecords
        If Not record("PDFExists") Then
            missingCount = missingCount + 1
        End If
    Next record

    ' Confirm with user
    response = MsgBox("Create PDFs for " & missingCount & " missing drawings?" & vbCrLf & vbCrLf & _
                      "This will open each drawing and generate a PDF." & vbCrLf & _
                      "This may take several minutes.", _
                      vbYesNo + vbQuestion, "Create Missing PDFs")

    If response = vbYes Then
        CreateMissingPDFs currentAuditRecords
    End If
End Sub

' Create PDFs for all missing drawings
Private Sub CreateMissingPDFs(auditRecords As Collection)
    Dim record As Object
    Dim successCount As Long
    Dim failCount As Long
    Dim drawingPath As String

    For Each record In auditRecords
        If Not record("PDFExists") Then
            drawingPath = record("DrawingPath")

            ' Call the main PDF creation function
            If CreatePDFForDrawing(drawingPath) Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        End If
    Next record

    ' Show results
    MsgBox "PDF Creation Complete!" & vbCrLf & vbCrLf & _
           "Successfully created: " & successCount & vbCrLf & _
           "Failed: " & failCount, vbInformation, "Create PDFs Complete"
End Sub

' Create PDF for a single drawing (wrapper around existing function)
Private Function CreatePDFForDrawing(drawingPath As String) As Boolean
    On Error Resume Next

    ' This will call the existing PDF creation logic from PDMPDF1
    ' You'll need to expose or call the appropriate function
    CreatePDFForDrawing = ProcessSingleDrawing(drawingPath)

    If Err.Number <> 0 Then
        Debug.Print "Error creating PDF for " & drawingPath & ": " & Err.Description
        CreatePDFForDrawing = False
        Err.Clear
    End If

    On Error GoTo 0
End Function
