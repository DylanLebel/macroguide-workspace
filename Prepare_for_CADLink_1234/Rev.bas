Attribute VB_Name = "Rev"
Option Explicit
Option Private Module

Sub UpdateRevisionForModel(model As ModelDoc2, ByRef propertiesToSet As Object)
    If model Is Nothing Then
        Exit Sub
    End If
    
    Dim filePath As String
    filePath = model.GetPathName
    
    If filePath = "" Then
        Exit Sub
    End If
    
    ' Update the revision from the PDM based on the active model's file path.
    UpdateRevisionFromPDM filePath, propertiesToSet
End Sub

Sub UpdateRevisionFromPDM(filePath As String, ByRef propertiesToSet As Object)
    Dim currentRevision As Long
    currentRevision = GetCurrentRevisionFromPDM(filePath)
    
    If currentRevision > 0 Then
        ' Update the custom property 'Revision' in the propertiesToSet dictionary
        propertiesToSet("Revision") = CStr(currentRevision)
    End If
End Sub

Function GetCurrentRevisionFromPDM(filePath As String) As Long
    Dim file As IEdmFile5
    Dim folder As IEdmFolder5
    Set file = vault.GetFileFromPath(filePath, folder)
    
    If Not file Is Nothing Then
        Dim currentRevision As Long
        currentRevision = file.CurrentVersion
        Dim revision As Long
        
        If currentRevision > 0 Then
            ' Get the local version of the file
            Dim localVersion As Long
            localVersion = file.GetLocalVersionNo(folder.ID)
            If localVersion = -1 Then
                ' If the file is not checked out, increment the revision by 1
                revision = currentRevision + 1
            Else
                ' If the file is checked out, use the current revision number
                revision = currentRevision
            End If
            ' Return the current revision number
            GetCurrentRevisionFromPDM = revision
        End If
    End If
End Function
