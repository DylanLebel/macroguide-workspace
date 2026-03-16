VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "PDF Audit Results"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExport_Click()
    Dim filePath As String
    Dim fileNum As Integer
    filePath = Environ("USERPROFILE") & "\Desktop\PDF_Audit_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, Me.txtResults.text
    Close #fileNum
    Shell "notepad.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus
    MsgBox "Exported to: " & vbCrLf & filePath, vbInformation
End Sub

Private Sub btnCopy_Click()
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText Me.txtResults.text
    dataObj.PutInClipboard
    MsgBox "Text copied to clipboard!", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnCreatePDFs_Click()
    AuditResultsFormBuilder.btnCreatePDFs_Click
End Sub

