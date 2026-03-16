VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Create a UserForm named "frmFolderBrowser" with the following controls:
' - txtPath: TextBox for manual path entry
' - btnBrowse: CommandButton to launch folder browser
' - btnOK: CommandButton to accept the path
' - btnCancel: CommandButton to cancel
' - lblPrompt: Label for instructions

Option Explicit

Public selectedPath As String
Public UserCancelled As Boolean

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    ' Center the form on screen
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ' Set default state
    selectedPath = ""
    UserCancelled = True
    
    ' Adjust size
    Me.Width = 400
    txtPath.Width = 300
    
    ' Set initial focus
    txtPath.SetFocus
End Sub

Private Sub btnBrowse_Click()
    Dim shellApp As Object
    Dim folder As Object
    
    On Error Resume Next
    
    ' Create Shell Application object
    Set shellApp = CreateObject("Shell.Application")
    
    ' Show folder browser dialog
    Set folder = shellApp.BrowseForFolder(0, "Select a folder:", 0)
    
    ' Process result
    If Not folder Is Nothing Then
        txtPath.text = folder.Self.Path
    End If
    
    ' Clean up
    Set folder = Nothing
    Set shellApp = Nothing
    
    ' Set focus back to the text box for any manual adjustments
    txtPath.SetFocus
    
    On Error GoTo 0
End Sub

Private Sub btnOK_Click()
    ' Validate the path
    Dim folderPath As String
    folderPath = Trim(txtPath.text)
    
    ' Check if path exists
    If folderPath = "" Then
        MsgBox "Please enter or select a folder path.", vbExclamation
        txtPath.SetFocus
        Exit Sub
    End If
    
    ' Optional: Add trailing backslash if missing
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Check if directory exists
    On Error Resume Next
    Dim dirExists As Boolean
    dirExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
    
    If Not dirExists Then
        Dim result As VbMsgBoxResult
        result = MsgBox("The folder path you entered doesn't exist or is not accessible:" & vbCrLf & _
                       folderPath & vbCrLf & vbCrLf & _
                       "Do you want to use this path anyway?", _
                       vbQuestion + vbYesNo, "Invalid Path")
                       
        If result = vbNo Then
            txtPath.SetFocus
            Exit Sub
        End If
    End If
    
    ' Return the path
    selectedPath = folderPath
    UserCancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    ' User cancelled
    selectedPath = ""
    UserCancelled = True
    Me.Hide
End Sub

Private Sub txtPath_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow Enter key to act as OK button
    If KeyAscii = 13 Then
        KeyAscii = 0 ' Suppress the beep
        btnOK_Click
    End If
End Sub

' Show audit results in a message box
Public Sub ShowAuditResults(resultsText As String)
    MsgBox resultsText, vbInformation, "PDF Audit Results"
End Sub

