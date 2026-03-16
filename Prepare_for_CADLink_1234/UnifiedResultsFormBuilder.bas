Attribute VB_Name = "UnifiedResultsFormBuilder"
Option Explicit

' ============================================
' UnifiedResultsFormBuilder - Creates results dialog at runtime
' No .frm/.frx files needed - form built entirely in code
' ============================================

' Module-level variables to store the form reference
Private m_FormInstance As Object
Private m_FormComponent As Object

' Counter for unique form names
Private m_FormCounter As Long

' Tab count tracking
Private m_TabCounts(0 To 3) As Long

' Main function to show the unified results form
Public Sub ShowUnifiedResultsForm(failedProperties As Dictionary, _
                                   failedLocations As Dictionary, _
                                   failedPieceParts As Dictionary, _
                                   totalProblematicModels As Long)

    Dim vbProj As Object
    Dim vbComp As Object
    Dim formName As String
    Dim frm As Object
    Dim errMsg As String

    On Error GoTo FormError

    DebugLog "UnifiedResultsFormBuilder: Starting dynamic form creation"

    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then
        errMsg = "Could not get ActiveVBProject"
        GoTo FormError
    End If

    ' Clean up any old forms first
    CleanupOldForms vbProj

    ' Use unique form name
    m_FormCounter = m_FormCounter + 1
    formName = "URForm" & m_FormCounter

    DebugLog "UnifiedResultsFormBuilder: Creating form " & formName

    ' Create new form (3 = vbext_ct_MSForm)
    On Error Resume Next
    Err.Clear
    Set vbComp = vbProj.VBComponents.Add(3)

    If Err.Number <> 0 Then
        errMsg = "VBComponents.Add failed: " & Err.description
        On Error GoTo FormError
        GoTo FormError
    End If
    On Error GoTo FormError

    If vbComp Is Nothing Then
        errMsg = "VBComponents.Add returned Nothing"
        GoTo FormError
    End If

    DoEvents

    ' Rename the form
    On Error Resume Next
    Dim actualName As String
    actualName = vbComp.Name
    Err.Clear
    vbComp.Name = formName

    If Err.Number <> 0 Then
        DebugLog "UnifiedResultsFormBuilder: Using default name " & actualName
        formName = actualName
        Err.Clear
    End If
    On Error GoTo FormError

    DebugLog "UnifiedResultsFormBuilder: Form component created as " & formName

    ' Get form designer
    Dim designer As Object
    Set designer = vbComp.designer
    If designer Is Nothing Then
        errMsg = "Could not get form Designer"
        GoTo FormError
    End If

    ' Set form properties
    On Error Resume Next
    With vbComp
        .properties("Caption") = GenerateFormTitle(totalProblematicModels)
        .properties("Width") = 650
        .properties("Height") = 550
        .properties("BackColor") = RGB(240, 240, 240)
        .properties("ShowModal") = True
        .properties("KeyPreview") = True  ' Enable keyboard events
    End With
    On Error GoTo FormError

    DebugLog "UnifiedResultsFormBuilder: Building controls"

    ' Check if there are checkout errors in the failedLocations
    Dim hasCheckoutErrors As Boolean
    hasCheckoutErrors = False
    If Not failedLocations Is Nothing Then
        Dim locItem As Variant
        For Each locItem In failedLocations.keys
            Dim displayValue As String
            displayValue = CStr(failedLocations(locItem))
            If InStr(displayValue, "MODEL NOT CHECKED OUT") > 0 Or _
               InStr(displayValue, "MODEL CHECKED OUT BY") > 0 Then
                hasCheckoutErrors = True
                Exit For
            End If
        Next locItem
    End If

    ' Build the form controls
    BuildFormControls designer, totalProblematicModels, hasCheckoutErrors

    DebugLog "UnifiedResultsFormBuilder: Adding event code"

    ' Add event code to the form
    AddEventCode vbComp

    DebugLog "UnifiedResultsFormBuilder: Populating data"

    ' Create form instance and populate data
    Set frm = VBA.UserForms.Add(formName)
    If frm Is Nothing Then
        errMsg = "Failed to add form to UserForms collection"
        GoTo FormError
    End If

    ' Store reference for event handlers
    Set m_FormInstance = frm
    Set m_FormComponent = vbComp

    ' Populate the form with data
    PopulateFormData frm, failedProperties, failedLocations, failedPieceParts

    DebugLog "UnifiedResultsFormBuilder: Showing form"

    ' Show modal
    frm.Show 1  ' 1 = vbModal

    DoEvents

    DebugLog "UnifiedResultsFormBuilder: Form closed, cleaning up"

    ' Cleanup
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    Set m_FormInstance = Nothing
    DoEvents

    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    Set m_FormComponent = Nothing
    On Error GoTo 0

    DebugLog "UnifiedResultsFormBuilder: Cleanup complete"
    Exit Sub

FormError:
    If errMsg = "" Then errMsg = Err.description
    DebugLog "UnifiedResultsFormBuilder ERROR: " & errMsg

    ' Cleanup on error
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    Set m_FormInstance = Nothing
    Set m_FormComponent = Nothing
    On Error GoTo 0

    MsgBox "Error creating results form: " & errMsg & vbCrLf & vbCrLf & _
           "See debug log for details.", vbExclamation
End Sub

' Generate form title based on problem count
Private Function GenerateFormTitle(totalModels As Long) As String
    If totalModels = 1 Then
        GenerateFormTitle = "Problems - 1 Model Needs Attention"
    ElseIf totalModels > 1 Then
        GenerateFormTitle = "Problems - " & totalModels & " Models Need Attention"
    Else
        GenerateFormTitle = "Problems Found"
    End If
End Function

' Clean up old dynamic forms
Private Sub CleanupOldForms(ByVal vbProj As Object)
    On Error Resume Next

    Dim comp As Object
    Dim compName As String
    Dim i As Long

    DebugLog "UnifiedResultsFormBuilder: Cleaning up old forms..."

    ' Unload any forms in UserForms collection
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        If Left(VBA.UserForms(i).Name, 6) = "URForm" Then
            DebugLog "UnifiedResultsFormBuilder: Unloading UserForm " & VBA.UserForms(i).Name
            Unload VBA.UserForms(i)
        End If
    Next i

    DoEvents

    ' Remove VBComponents
    For Each comp In vbProj.VBComponents
        compName = comp.Name
        If Left(compName, 6) = "URForm" Then
            DebugLog "UnifiedResultsFormBuilder: Removing component " & compName
            vbProj.VBComponents.Remove comp
            DoEvents
        End If
    Next comp

    On Error GoTo 0
End Sub

' Build all form controls
Private Sub BuildFormControls(ByVal designer As Object, totalProblematicModels As Long, hasCheckoutErrors As Boolean)
    Dim yPos As Single
    yPos = 10

    ' Info label at top
    Dim lblInfo As Object
    Set lblInfo = designer.Controls.Add("Forms.Label.1", "lblInfo")
    With lblInfo
        .Left = 10
        .Top = yPos
        .width = 620
        .height = 90
        .BackColor = RGB(217, 237, 247)  ' Light blue
        .ForeColor = RGB(49, 112, 143)   ' Dark blue
        .BorderStyle = 1
        .WordWrap = True
        .Font.size = 9
        .Caption = BuildInfoText(totalProblematicModels, hasCheckoutErrors)
    End With
    yPos = yPos + 95

    ' MultiPage control for tabs
    Dim multiPage As Object
    Set multiPage = designer.Controls.Add("Forms.MultiPage.1", "MultiPage1")
    With multiPage
        .Left = 10
        .Top = yPos
        .width = 620
        .height = 380
        .BackColor = RGB(255, 255, 255)
    End With

    ' Add extra pages to MultiPage (starts with 2, we need 4)
    On Error Resume Next
    Do While multiPage.Pages.Count < 4
        multiPage.Pages.Add
    Loop
    On Error GoTo 0

    ' Set page captions
    multiPage.Pages(0).Caption = "Properties"
    multiPage.Pages(1).Caption = "Locations"
    multiPage.Pages(2).Caption = "Piece Parts"
    multiPage.Pages(3).Caption = "Errors"

    ' Add ListBox to each page
    Dim i As Integer
    For i = 0 To 3
        Dim lb As Object
        Set lb = multiPage.Pages(i).Controls.Add("Forms.ListBox.1", "ListBox" & (i + 1))
        With lb
            .Left = 5
            .Top = 5
            .width = multiPage.width - 20
            .height = multiPage.height - 40
            .BackColor = RGB(255, 255, 255)
        End With
    Next i

    yPos = yPos + 385

    ' Close button
    Dim btnClose As Object
    Set btnClose = designer.Controls.Add("Forms.CommandButton.1", "cmdClose")
    With btnClose
        .Left = 540
        .Top = yPos
        .width = 90
        .height = 30
        .Caption = "Close"
        .BackColor = RGB(0, 120, 215)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' Export button
    Dim btnExport As Object
    Set btnExport = designer.Controls.Add("Forms.CommandButton.1", "cmdExport")
    With btnExport
        .Left = 430
        .Top = yPos
        .width = 100
        .height = 30
        .Caption = "Export"
        '.Caption = "Export (Ctrl+E)"
        .BackColor = RGB(240, 240, 240)
        .Font.size = 8
    End With
End Sub

' Build info text
Private Function BuildInfoText(totalModels As Long, hasCheckoutErrors As Boolean) As String
    Dim text As String
    text = "Please fix the following errors. There are " & totalModels & " problematic model(s)." & vbCrLf & vbCrLf

    text = text & "USAGE: "
    text = text & "Double-click any item to copy it to clipboard" & vbCrLf
    text = text & "Double-click Locations to open file in Explorer | "
    text = text & "Use Export button to save all errors to file"

    ' Only show checkout note if there are actually checkout errors
    If hasCheckoutErrors Then
        text = text & vbCrLf & vbCrLf
        text = text & "NOTE: Models checked out by another user could not be modified."
    End If

    BuildInfoText = text
End Function

' Add event handler code to the form
Private Sub AddEventCode(ByVal vbComp As Object)
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule

    ' Build code in chunks to avoid "too many line continuations" error
    Dim code1 As String
    Dim code2 As String
    Dim code3 As String
    Dim code4 As String
    Dim eventCode As String

    ' Part 1: Option and Initialize
    code1 = "Option Explicit" & vbCrLf & vbCrLf
    code1 = code1 & "Private Sub UserForm_Initialize()" & vbCrLf
    code1 = code1 & "    Me.StartUpPosition = 1" & vbCrLf
    code1 = code1 & "End Sub" & vbCrLf & vbCrLf

    ' Part 2: Empty (no longer needed)
    code2 = ""

    ' Part 3: Button handlers
    code3 = "Private Sub cmdClose_Click()" & vbCrLf
    code3 = code3 & "    Unload Me" & vbCrLf
    code3 = code3 & "End Sub" & vbCrLf & vbCrLf
    code3 = code3 & "Private Sub cmdExport_Click()" & vbCrLf
    code3 = code3 & "    UnifiedResultsFormBuilder.HandleExport Me" & vbCrLf
    code3 = code3 & "End Sub" & vbCrLf & vbCrLf

    ' Part 4: Add Copy button functionality and ListBox double-click events
    code4 = "Private Sub UserForm_Activate()" & vbCrLf
    code4 = code4 & "    On Error Resume Next" & vbCrLf
    code4 = code4 & "    Me.MultiPage1.Pages(Me.MultiPage1.Value).Controls(0).SetFocus" & vbCrLf
    code4 = code4 & "End Sub" & vbCrLf & vbCrLf
    code4 = code4 & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    code4 = code4 & "End Sub" & vbCrLf & vbCrLf
    
    ' ListBox double-click handlers
    code4 = code4 & "Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    code4 = code4 & "    UnifiedResultsFormBuilder.CopySelected Me" & vbCrLf
    code4 = code4 & "End Sub" & vbCrLf & vbCrLf
    
    code4 = code4 & "Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    code4 = code4 & "    UnifiedResultsFormBuilder.HandleLocationDoubleClick Me" & vbCrLf
    code4 = code4 & "End Sub" & vbCrLf & vbCrLf
    
    code4 = code4 & "Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    code4 = code4 & "    UnifiedResultsFormBuilder.CopySelected Me" & vbCrLf
    code4 = code4 & "End Sub" & vbCrLf & vbCrLf
    
    code4 = code4 & "Private Sub ListBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    code4 = code4 & "    UnifiedResultsFormBuilder.CopySelected Me" & vbCrLf
    code4 = code4 & "End Sub"

    ' Combine all parts
    eventCode = code1 & code2 & code3 & code4

    codeMod.InsertLines codeMod.CountOfLines + 1, eventCode
End Sub

' Populate form with data
Private Sub PopulateFormData(ByVal frm As Object, _
                              failedProperties As Dictionary, _
                              failedLocations As Dictionary, _
                              failedPieceParts As Dictionary)
    On Error Resume Next

    ' Separate checkout errors from regular locations
    Dim regularLocations As New Dictionary
    Dim checkoutErrors As New Dictionary

    If failedLocations.Count > 0 Then
        Dim locItem As Variant
        For Each locItem In failedLocations.keys
            Dim displayValue As String
            displayValue = CStr(failedLocations(locItem))

            If InStr(displayValue, "MODEL NOT CHECKED OUT") > 0 Or _
               InStr(displayValue, "MODEL CHECKED OUT BY") > 0 Then
                checkoutErrors.Add locItem, displayValue
            Else
                regularLocations.Add locItem, displayValue
            End If
        Next locItem
    End If

    ' Get the listboxes from each page
    Dim lb1 As Object  ' Properties
    Dim lb2 As Object  ' Locations
    Dim lb3 As Object  ' Piece Parts
    Dim lb4 As Object  ' Errors

    Set lb1 = frm.MultiPage1.Pages(0).Controls("ListBox1")
    Set lb2 = frm.MultiPage1.Pages(1).Controls("ListBox2")
    Set lb3 = frm.MultiPage1.Pages(2).Controls("ListBox3")
    Set lb4 = frm.MultiPage1.Pages(3).Controls("ListBox4")

    ' Populate Properties
    If Not failedProperties Is Nothing Then
        Dim propItem As Variant
        For Each propItem In failedProperties.keys
            If Not IsNull(propItem) And Not IsEmpty(propItem) Then
                lb1.AddItem CStr(propItem)
            End If
        Next propItem
    End If

    ' Populate Locations
    If regularLocations.Count > 0 Then
        lb2.ColumnCount = 2
        lb2.ColumnWidths = "200;400"

        For Each locItem In regularLocations.keys
            Dim modelPath As String
            Dim modelName As String
            modelPath = CStr(locItem)

            Dim pathParts As Variant
            pathParts = Split(modelPath, "\")
            If UBound(pathParts) >= 0 Then
                modelName = pathParts(UBound(pathParts))
            Else
                modelName = modelPath
            End If

            lb2.AddItem
            If lb2.ListCount > 0 Then
                lb2.List(lb2.ListCount - 1, 0) = modelName
                lb2.List(lb2.ListCount - 1, 1) = modelPath
            End If
        Next locItem
    End If

    ' Populate Piece Parts
    If Not failedPieceParts Is Nothing Then
        Dim partItem As Variant
        For Each partItem In failedPieceParts.keys
            If Not IsNull(partItem) And Not IsEmpty(partItem) Then
                lb3.AddItem CStr(partItem)
            End If
        Next partItem
    End If

    ' Populate Errors (checkout errors)
    If checkoutErrors.Count > 0 Then
        For Each locItem In checkoutErrors.keys
            lb4.AddItem CStr(checkoutErrors(locItem))
        Next locItem
    End If

    ' Update tab counts
    m_TabCounts(0) = lb1.ListCount
    m_TabCounts(1) = lb2.ListCount
    m_TabCounts(2) = lb3.ListCount
    m_TabCounts(3) = lb4.ListCount

    ' Update tab captions with counts
    If m_TabCounts(0) > 0 Then
        frm.MultiPage1.Pages(0).Caption = "Properties (" & m_TabCounts(0) & ")"
    End If
    If m_TabCounts(1) > 0 Then
        frm.MultiPage1.Pages(1).Caption = "Locations (" & m_TabCounts(1) & ")"
    End If
    If m_TabCounts(2) > 0 Then
        frm.MultiPage1.Pages(2).Caption = "Piece Parts (" & m_TabCounts(2) & ")"
    End If
    If m_TabCounts(3) > 0 Then
        frm.MultiPage1.Pages(3).Caption = "Errors (" & m_TabCounts(3) & ")"
    End If

    ' Select appropriate tab
    If m_TabCounts(3) > 0 Then
        frm.MultiPage1.value = 3
    ElseIf m_TabCounts(0) > 0 Then
        frm.MultiPage1.value = 0
    ElseIf m_TabCounts(1) > 0 Then
        frm.MultiPage1.value = 1
    ElseIf m_TabCounts(2) > 0 Then
        frm.MultiPage1.value = 2
    End If

    On Error GoTo 0
End Sub

' Event handlers called from dynamic form
Public Sub CopySelected(ByVal frm As Object)
    On Error GoTo ClipboardError

    Dim selectedText As String
    Dim currentTab As Integer
    currentTab = frm.MultiPage1.value

    Dim lb As Object
    Set lb = frm.MultiPage1.Pages(currentTab).Controls(0)

    If Not lb Is Nothing And lb.ListIndex >= 0 Then
        If lb.ColumnCount > 1 Then
            selectedText = lb.List(lb.ListIndex, 0) & " - " & lb.List(lb.ListIndex, 1)
        Else
            selectedText = lb.List(lb.ListIndex)
        End If

        If selectedText <> "" Then
            ' Method 1: Try MSForms.DataObject
            On Error Resume Next
            Dim clipboard As Object
            Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            If clipboard Is Nothing Then
                Set clipboard = CreateObject("MSForms.DataObject")
            End If

            If Not clipboard Is Nothing Then
                clipboard.Clear
                clipboard.SetText selectedText
                clipboard.PutInClipboard
            End If

            If Err.Number <> 0 Then
                ' Method 2: Fall back to VBA.Interaction method
                Err.Clear
                On Error GoTo ClipboardError

                ' Use a hidden textbox approach
                Dim objIE As Object
                Set objIE = CreateObject("InternetExplorer.Application")
                objIE.Navigate "about:blank"
                Do While objIE.Busy: DoEvents: Loop
                objIE.document.parentWindow.clipboardData.SetData "text", selectedText
                objIE.Quit
                Set objIE = Nothing
            End If
            On Error GoTo ClipboardError

            ' Show confirmation
            MsgBox "Copied to clipboard:" & vbCrLf & vbCrLf & Left(selectedText, 200), vbInformation, "Copied"
        End If
    Else
        MsgBox "No item selected. Please select an item first.", vbExclamation, "Nothing Selected"
    End If

    Exit Sub

ClipboardError:
    MsgBox "Failed to copy to clipboard: " & Err.description & vbCrLf & vbCrLf & _
           "Text was: " & selectedText, vbExclamation, "Clipboard Error"
End Sub

Public Sub CopyAllFromTab(ByVal frm As Object)
    On Error GoTo ClipboardError

    Dim allText As String
    Dim currentTab As Integer
    Dim i As Long

    currentTab = frm.MultiPage1.value
    allText = ""

    Dim lb As Object
    Set lb = frm.MultiPage1.Pages(currentTab).Controls(0)

    If Not lb Is Nothing And lb.ListCount > 0 Then
        For i = 0 To lb.ListCount - 1
            If lb.ColumnCount > 1 Then
                allText = allText & lb.List(i, 0) & " - " & lb.List(i, 1) & vbCrLf
            Else
                allText = allText & lb.List(i) & vbCrLf
            End If
        Next i

        ' Use improved clipboard method
        On Error Resume Next
        Dim clipboard As Object
        Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        If clipboard Is Nothing Then
            Set clipboard = CreateObject("MSForms.DataObject")
        End If

        If Not clipboard Is Nothing Then
            clipboard.Clear
            clipboard.SetText allText
            clipboard.PutInClipboard
        End If
        On Error GoTo ClipboardError

        frm.Caption = "Problems - All items copied!"
    Else
        frm.Caption = "Problems - No items to copy"
    End If

    Exit Sub

ClipboardError:
    MsgBox "Failed to copy all items: " & Err.description, vbExclamation, "Clipboard Error"
End Sub

Public Sub HandleLocationDoubleClick(ByVal frm As Object)
    On Error Resume Next

    ' Get the Locations listbox (page 1, index 0)
    Dim lb As Object
    Set lb = frm.MultiPage1.Pages(1).Controls(0)

    If Not lb Is Nothing And lb.ListIndex >= 0 Then
        Dim filePath As String

        ' Get the full path from column 1 if it's a 2-column listbox
        If lb.ColumnCount > 1 Then
            filePath = lb.List(lb.ListIndex, 1)
        Else
            filePath = lb.List(lb.ListIndex)
        End If

        If filePath <> "" Then
            ' Open in Windows Explorer and select the file
            Dim shell As Object
            Set shell = CreateObject("Shell.Application")
            shell.Open "explorer.exe /select," & Chr(34) & filePath & Chr(34)
            Set shell = Nothing
        End If
    End If

    On Error GoTo 0
End Sub

Public Sub HandleExport(ByVal frm As Object)
    On Error GoTo ErrorHandler

    Dim fileName As String
    Dim fileNum As Integer
    Dim i As Long
    Dim lb As Object

    ' Use simple file path - save to desktop or current directory
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    fileName = desktopPath & "\CADLink_Errors_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"

    ' Ask user if they want to export
    If MsgBox("Export errors to:" & vbCrLf & fileName & "?", vbYesNo + vbQuestion, "Export Errors") = vbNo Then
        Exit Sub
    End If

    fileNum = FreeFile
    Open fileName For Output As #fileNum

    Print #fileNum, "CADLink Preparation Errors Report"
    Print #fileNum, "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, String(60, "=")
    Print #fileNum, ""

    ' Export each tab
    Dim tabNames As Variant
    tabNames = Array("PROPERTIES", "LOCATIONS", "PIECE PARTS", "CHECKOUT ERRORS")

    Dim tabIdx As Integer
    For tabIdx = 0 To 3
        Set lb = frm.MultiPage1.Pages(tabIdx).Controls(0)
        If Not lb Is Nothing And lb.ListCount > 0 Then
            Print #fileNum, tabNames(tabIdx) & " (" & lb.ListCount & "):"
            Print #fileNum, String(60, "-")
            For i = 0 To lb.ListCount - 1
                If lb.ColumnCount > 1 Then
                    Print #fileNum, lb.List(i, 0) & " | " & lb.List(i, 1)
                Else
                    Print #fileNum, lb.List(i)
                End If
            Next i
            Print #fileNum, ""
        End If
    Next tabIdx

    Close #fileNum

    MsgBox "Errors exported successfully to:" & vbCrLf & fileName, vbInformation, "Export Complete"
    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error exporting file: " & Err.description, vbExclamation, "Export Error"
End Sub
