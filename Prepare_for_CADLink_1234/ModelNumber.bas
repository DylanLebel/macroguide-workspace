Attribute VB_Name = "ModelNumber"
Option Explicit
Option Private Module

Sub PopulateModelNumber(model As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    Dim modelName As String
    Dim dotPosition As Long
    
    ' Get the full model title
    modelName = model.GetTitle()
    
    ' Find the position of the last dot (if any)
    dotPosition = InStrRev(modelName, ".")
    
    ' If a dot is found, remove the extension. Otherwise, use the full name.
    If dotPosition > 0 Then
        modelName = Left(modelName, dotPosition - 1)
    End If
    
    ' Add "Model Number" to the propertiesToSet dictionary
    propertiesToSet("Model Number") = modelName
End Sub
