Attribute VB_Name = "FlatBarFormBuilder"
Option Explicit

' ============================================
' FlatBarFormBuilder - Creates flat bar dialog at runtime
' No .frm/.frx files needed - form built entirely in code
' ============================================

' Result constants
Public Const FB_CONVERT As Integer = 1
Public Const FB_SKIP As Integer = 2
Public Const FB_CANCEL As Integer = 3

' Module-level variables to store results
Private m_Result As Integer
Private m_NeverAskAgain As Boolean
Private m_SkipAllThisRun As Boolean

' Counter for unique form names
Private m_FormCounter As Long

' Public properties to read results
Public Property Get result() As Integer
    result = m_Result
End Property

Public Property Get NeverAskAgain() As Boolean
    NeverAskAgain = m_NeverAskAgain
End Property

Public Property Get SkipAllThisRun() As Boolean
    SkipAllThisRun = m_SkipAllThisRun
End Property

' Main function to show the flat bar dialog
Public Function ShowFlatBarDialog(ByVal partName As String, _
                                   ByVal description As String, _
                                   ByVal material As String, _
                                   ByVal dimensions As String, _
                                   ByVal flatBarSize As String) As Integer

    ' Reset results
    m_Result = FB_CANCEL
    m_NeverAskAgain = False
    m_SkipAllThisRun = False

    Dim vbProj As Object
    Dim vbComp As Object
    Dim formName As String
    Dim frm As Object
    Dim errMsg As String

    ' Try to create and show the form
    On Error GoTo FormError

    DebugLog "FlatBarFormBuilder: Starting dynamic form creation"

    Set vbProj = Application.VBE.ActiveVBProject
    If vbProj Is Nothing Then
        errMsg = "Could not get ActiveVBProject"
        GoTo FormError
    End If

    DebugLog "FlatBarFormBuilder: Got VBProject"

    ' Clean up any old forms first (from previous runs)
    CleanupOldForms vbProj

    ' Use unique form name each time to avoid conflicts
    m_FormCounter = m_FormCounter + 1
    formName = "FBForm" & m_FormCounter

    DebugLog "FlatBarFormBuilder: Creating form " & formName

    ' Check if a component with this name already exists and remove it
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents(formName)
    If Err.Number = 0 And Not existingComp Is Nothing Then
        DebugLog "FlatBarFormBuilder: Removing existing component " & formName
        vbProj.VBComponents.Remove existingComp
        DoEvents
    End If
    Err.Clear
    On Error GoTo FormError

    ' Create new form (3 = vbext_ct_MSForm)
    On Error Resume Next
    Err.Clear
    Set vbComp = vbProj.VBComponents.Add(3)

    If Err.Number <> 0 Then
        errMsg = "VBComponents.Add failed: " & Err.description & " (Err " & Err.Number & ")"
        On Error GoTo FormError
        GoTo FormError
    End If
    On Error GoTo FormError

    If vbComp Is Nothing Then
        errMsg = "VBComponents.Add returned Nothing"
        GoTo FormError
    End If

    DebugLog "FlatBarFormBuilder: Component created, getting default name..."

    ' Allow VBA to fully initialize the component
    DoEvents
    
    ' Get the actual name that was assigned (usually UserForm1, UserForm2, etc.)
    On Error Resume Next
    Err.Clear
    Dim actualName As String
    actualName = vbComp.Name
    
    If Err.Number <> 0 Then
        errMsg = "Could not get component name: " & Err.description & " (Err " & Err.Number & ")"
        On Error GoTo FormError
        GoTo FormError
    End If
    
    DebugLog "FlatBarFormBuilder: Component default name is " & actualName
    
    ' Now try to rename it
    Err.Clear
    vbComp.Name = formName

    ' If it fails, wait a bit and retry (sometimes VBA needs time to register the component)
    If Err.Number <> 0 Then
        DebugLog "FlatBarFormBuilder: First rename attempt failed, retrying..."
        DoEvents
        Err.Clear
        vbComp.Name = formName
    End If

    ' If still fails, try getting component from collection and renaming
    If Err.Number <> 0 Then
        DebugLog "FlatBarFormBuilder: Second rename attempt failed, trying collection access..."
        Err.Clear
        Set vbComp = vbProj.VBComponents(actualName)
        If Err.Number = 0 And Not vbComp Is Nothing Then
            Err.Clear
            vbComp.Name = formName
        End If
    End If

    ' If rename still fails, use the default name
    If Err.Number <> 0 Then
        DebugLog "FlatBarFormBuilder: Rename failed, using default name " & actualName
        formName = actualName
        Err.Clear
    End If
    On Error GoTo FormError

    DebugLog "FlatBarFormBuilder: Form component created as " & formName

    ' Get form designer
    Dim designer As Object
    Set designer = vbComp.designer
    If designer Is Nothing Then
        errMsg = "Could not get form Designer"
        GoTo FormError
    End If

    ' Set form properties
    With vbComp
        .properties("Caption") = "Flat Bar Conversion"
        .properties("Width") = 350
        .properties("Height") = 290
        .properties("BackColor") = RGB(240, 240, 240)
        .properties("ShowModal") = True
    End With

    DebugLog "FlatBarFormBuilder: Building controls"

    ' Build the form controls
    BuildFormControls designer, partName, description, material, dimensions, flatBarSize

    DebugLog "FlatBarFormBuilder: Adding event code"

    ' Add event code to the form
    AddEventCode vbComp

    DebugLog "FlatBarFormBuilder: Showing form"

    ' Show the form modally
    ' IMPORTANT: Use late binding and vbModal constant (1) to ensure the form stays visible
    Set frm = VBA.UserForms.Add(formName)
    If frm Is Nothing Then
        errMsg = "Failed to add form to UserForms collection"
        GoTo FormError
    End If

    ' Show modal - this should block until user closes the form
    frm.Show 1  ' 1 = vbModal

    ' Give VBA time to process the Hide event
    DoEvents

    DebugLog "FlatBarFormBuilder: Form closed, cleaning up"

    ' Cleanup - IMPORTANT: Unload the form first, then remove component
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    DoEvents  ' Let VBA process the unload

    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    On Error GoTo 0

    DebugLog "FlatBarFormBuilder: Cleanup complete, result = " & m_Result

    ShowFlatBarDialog = m_Result
    Exit Function

FormError:
    ' If dynamic form creation fails, fall back to MsgBox
    If errMsg = "" Then errMsg = Err.description
    DebugLog "FlatBarFormBuilder ERROR: " & errMsg & " (Err " & Err.Number & "). Using MsgBox fallback."

    ' Try to clean up on error
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    On Error GoTo 0

    ShowFlatBarDialog = ShowFlatBarMsgBox(partName, description, material, dimensions, flatBarSize)
End Function

' Clean up old dynamic forms from previous runs
Private Sub CleanupOldForms(ByVal vbProj As Object)
    On Error Resume Next

    Dim comp As Object
    Dim compName As String
    Dim i As Long

    DebugLog "FlatBarFormBuilder: Cleaning up old forms..."

    ' Unload any forms that might be in the UserForms collection
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        If Left(VBA.UserForms(i).Name, 6) = "FBForm" Then
            DebugLog "FlatBarFormBuilder: Unloading UserForm " & VBA.UserForms(i).Name
            Unload VBA.UserForms(i)
        End If
    Next i

    DoEvents

    ' Now remove the VBComponents
    For Each comp In vbProj.VBComponents
        compName = comp.Name
        If Left(compName, 6) = "FBForm" Then
            DebugLog "FlatBarFormBuilder: Removing component " & compName
            vbProj.VBComponents.Remove comp
            DoEvents
        End If
    Next comp

    DebugLog "FlatBarFormBuilder: Cleanup complete"

    On Error GoTo 0
End Sub

' Build all form controls
Private Sub BuildFormControls(ByVal designer As Object, _
                               ByVal partName As String, _
                               ByVal description As String, _
                               ByVal material As String, _
                               ByVal dimensions As String, _
                               ByVal flatBarSize As String)
    Dim yPos As Single
    yPos = 10

    ' Part info label
    Dim lblInfo As Object
    Set lblInfo = designer.Controls.Add("Forms.Label.1", "lblInfo")
    With lblInfo
        .Left = 10
        .Top = yPos
        .width = 320
        .height = 58
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .WordWrap = True
        .Caption = "Part: " & partName & vbCrLf & _
                   IIf(description <> "", "Desc: " & description & vbCrLf, "") & _
                   IIf(material <> "", "Material: " & material & vbCrLf, "") & _
                   "Dimensions: " & dimensions
    End With
    yPos = yPos + 66

    ' Match background
    Dim lblMatchBg As Object
    Set lblMatchBg = designer.Controls.Add("Forms.Label.1", "lblMatchBg")
    With lblMatchBg
        .Left = 10
        .Top = yPos
        .width = 320
        .height = 48
        .BackColor = RGB(220, 245, 220)
        .BorderStyle = 1
        .Caption = ""
    End With

    ' Match title
    Dim lblMatchTitle As Object
    Set lblMatchTitle = designer.Controls.Add("Forms.Label.1", "lblMatchTitle")
    With lblMatchTitle
        .Left = 16
        .Top = yPos + 6
        .width = 300
        .height = 14
        .BackColor = RGB(220, 245, 220)
        .ForeColor = RGB(40, 120, 40)
        .Caption = "Matches Flat Bar:"
    End With

    ' Match size (large)
    Dim lblMatchSize As Object
    Set lblMatchSize = designer.Controls.Add("Forms.Label.1", "lblMatchSize")
    With lblMatchSize
        .Left = 16
        .Top = yPos + 22
        .width = 300
        .height = 20
        .BackColor = RGB(220, 245, 220)
        .ForeColor = RGB(40, 120, 40)
        .Font.size = 14
        .Font.Bold = True
        .Caption = flatBarSize
    End With
    yPos = yPos + 56

    ' Checkbox: Never ask again
    Dim chkNeverAsk As Object
    Set chkNeverAsk = designer.Controls.Add("Forms.CheckBox.1", "chkNeverAsk")
    With chkNeverAsk
        .Left = 10
        .Top = yPos
        .width = 320
        .height = 18
        .BackColor = RGB(240, 240, 240)
        .Caption = "Don't ask again for this part"
    End With
    yPos = yPos + 22

    ' Checkbox: Skip all
    Dim chkSkipAll As Object
    Set chkSkipAll = designer.Controls.Add("Forms.CheckBox.1", "chkSkipAll")
    With chkSkipAll
        .Left = 10
        .Top = yPos
        .width = 320
        .height = 18
        .BackColor = RGB(240, 240, 240)
        .Caption = "Skip all flat bar prompts this run"
    End With
    yPos = yPos + 30

    ' Convert button
    Dim btnConvert As Object
    Set btnConvert = designer.Controls.Add("Forms.CommandButton.1", "btnConvert")
    With btnConvert
        .Left = 10
        .Top = yPos
        .width = 95
        .height = 26
        .Caption = "Convert to FB"
        .BackColor = RGB(0, 120, 215)
    End With

    ' Skip button
    Dim btnSkip As Object
    Set btnSkip = designer.Controls.Add("Forms.CommandButton.1", "btnSkip")
    With btnSkip
        .Left = 115
        .Top = yPos
        .width = 70
        .height = 26
        .Caption = "Skip"
    End With

    ' Cancel button
    Dim btnCancel As Object
    Set btnCancel = designer.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 240
        .Top = yPos
        .width = 90
        .height = 26
        .Caption = "Cancel"
    End With
End Sub

' Add event handler code to the form
Private Sub AddEventCode(ByVal vbComp As Object)
    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule

    Dim eventCode As String
    eventCode = _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    ' Ensure form is visible and stays on top" & vbCrLf & _
        "    Me.StartUpPosition = 1" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnConvert_Click()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    FlatBarFormBuilder.HandleConvert" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnSkip_Click()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    FlatBarFormBuilder.HandleSkip Me.chkNeverAsk.Value, Me.chkSkipAll.Value" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnCancel_Click()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    FlatBarFormBuilder.HandleCancel" & vbCrLf & _
        "    Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    If CloseMode = 0 Then FlatBarFormBuilder.HandleCancel" & vbCrLf & _
        "End Sub"

    codeMod.InsertLines codeMod.CountOfLines + 1, eventCode
End Sub

' Event handlers called from dynamic form
Public Sub HandleConvert()
    m_Result = FB_CONVERT
    m_NeverAskAgain = False
    m_SkipAllThisRun = False
End Sub

Public Sub HandleSkip(ByVal neverAsk As Boolean, ByVal skipAll As Boolean)
    m_Result = FB_SKIP
    m_NeverAskAgain = neverAsk
    m_SkipAllThisRun = skipAll
End Sub

Public Sub HandleCancel()
    m_Result = FB_CANCEL
    m_NeverAskAgain = False
    m_SkipAllThisRun = False
End Sub

' Fallback MsgBox implementation
Private Function ShowFlatBarMsgBox(ByVal partName As String, _
                                    ByVal description As String, _
                                    ByVal material As String, _
                                    ByVal dimensions As String, _
                                    ByVal flatBarSize As String) As Integer

    Dim partInfo As String
    partInfo = "PART: " & partName & vbCrLf
    If description <> "" Then partInfo = partInfo & "Description: " & description & vbCrLf
    If material <> "" Then partInfo = partInfo & "Material: " & material & vbCrLf
    partInfo = partInfo & "Dimensions: " & dimensions

    Dim mainPrompt As String
    mainPrompt = partInfo & vbCrLf & vbCrLf & _
                 String(40, "-") & vbCrLf & _
                 "  MATCHES FLAT BAR: " & flatBarSize & vbCrLf & _
                 String(40, "-") & vbCrLf & vbCrLf & _
                 "Convert this part to Flat Bar (FB)?"

    Dim response As VbMsgBoxResult
    response = MsgBox(mainPrompt, vbYesNo + vbQuestion, "Flat Bar Match Found")

    If response = vbYes Then
        m_Result = FB_CONVERT
        ShowFlatBarMsgBox = FB_CONVERT
        Exit Function
    End If

    ' Skip all?
    response = MsgBox("Skip ALL flat bar prompts for the rest of this run?", _
                     vbYesNo + vbQuestion, "Skip All?")

    If response = vbYes Then
        m_Result = FB_SKIP
        m_SkipAllThisRun = True
        ShowFlatBarMsgBox = FB_SKIP
        Exit Function
    End If

    ' Never ask again?
    response = MsgBox("Never ask about this specific part again?", _
                     vbYesNo + vbQuestion, "Remember Choice?")

    If response = vbYes Then
        m_Result = FB_SKIP
        m_NeverAskAgain = True
    Else
        m_Result = FB_SKIP
    End If

    ShowFlatBarMsgBox = FB_SKIP
End Function
