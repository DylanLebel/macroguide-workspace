Attribute VB_Name = "MaterialModule"
Option Explicit
Option Private Module

Public Sub AutofillMaterial(model As ModelDoc2, ByRef propertiesToSet As Object)
    ' *** Get the actual applied material first to compare ***
    Dim actualMaterial As String
    If model.GetType <> swDocASSEMBLY Then
        Dim swPart As PartDoc
        Set swPart = model
        actualMaterial = swPart.GetMaterialPropertyName2("", "")
    End If
    
    ' *** Don't overwrite Material if it already matches the actual applied material ***
    If propertiesToSet.exists("Material") Then
        If propertiesToSet("Material") <> "" And propertiesToSet("Material") <> "Material <not specified>" Then
            ' Check if current property matches actual material
            If propertiesToSet("Material") = actualMaterial And actualMaterial <> "" Then
                DebugLog "AutofillMaterial: Material already matches actual material '" & actualMaterial & "', skipping autofill"
                Exit Sub
            End If
            ' Check if it's not a SW-Material format that needs updating
            If Not (InStr(propertiesToSet("Material"), "SW-Material@") > 0) And actualMaterial = "" Then
                DebugLog "AutofillMaterial: Material already set to '" & propertiesToSet("Material") & "' and no actual material applied, skipping autofill"
                Exit Sub
            End If
        End If
    End If
    
    If model Is Nothing Then
        MsgBox "Error: No model provided.", vbCritical
        Exit Sub
    End If

    ' OPTIMIZATION: Use global swApp accessor (not used in this function but kept for consistency)
    ' Dim swAppLocal As SldWorks.SldWorks
    ' Set swAppLocal = GetSwApp()

    ' Get the document name
    Dim docName As String
    docName = model.GetPathName
    
    If docName = "" Then
        MsgBox "Error: Document name is empty.", vbCritical
        Exit Sub
    End If
    
    ' Extract the file name with extension from the full path
    Dim fileName As String
    fileName = Right(docName, Len(docName) - InStrRev(docName, "\"))
    
    Dim materialValue As String
    
    ' Check if the model is an assembly
    If model.GetType = swDocASSEMBLY Then
        materialValue = "ASMW"
    Else
        ' Use the actual material we already retrieved
        If actualMaterial <> "" And actualMaterial <> "Material <not specified>" Then
            materialValue = actualMaterial
            DebugLog "AutofillMaterial: Using actual applied material: " & actualMaterial
        Else
            ' Fall back to UNKNOWN if no material is applied
            materialValue = "UNKNOWN"
            DebugLog "AutofillMaterial: No material applied, using UNKNOWN"
        End If
    End If
    
    ' Check if the custom property already exists in the propertiesToSet dictionary
    If Not propertiesToSet.exists("Material") Or propertiesToSet("Material") <> materialValue Then
        ' Add or update the material property in the propertiesToSet dictionary
        propertiesToSet("Material") = materialValue
        DebugLog "AutofillMaterial: Set Material to " & materialValue
    End If
End Sub
