Attribute VB_Name = "AddDate"
Option Explicit
Option Private Module

Sub AddCreatedDateToProperties(model As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    ' Ensure the model has been saved and has a valid path
    If model.GetPathName = "" Then
        ' If the model hasn't been saved, skip setting the date
        Exit Sub
    End If
    
    ' Check if the "Date" custom property already has a value
    If Not propertiesToSet.exists("Date") Or propertiesToSet("Date") = "" Then
        ' Retrieve the created date of the model
        Dim createdDate As String
        createdDate = GetCreatedDate(model)
        
        ' Add the created date to the dictionary only if it's not empty
        If createdDate <> "" Then
            propertiesToSet("Date") = createdDate
        End If
    End If
End Sub

Function GetCreatedDate(model As ModelDoc2) As String
    Dim filePath As String
    Dim fso As Object
    Dim file As Object

    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the path of the current model
    filePath = model.GetPathName
    
    If fso.FileExists(filePath) Then
        ' Get the file object
        Set file = fso.GetFile(filePath)
        ' Format the creation date as "yyyy-MM-dd"
        GetCreatedDate = Format(file.DateCreated, "yyyy-MM-dd")
    Else
        ' Return an empty string if the file does not exist
        GetCreatedDate = ""
    End If

    ' Clean up
    Set file = Nothing
    Set fso = Nothing
End Function

