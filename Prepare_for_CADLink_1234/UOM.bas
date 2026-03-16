Attribute VB_Name = "UOM"
Option Explicit
Option Private Module

'Sub AutofillUnits(model As ModelDoc2, ByRef propertiesToSet As Object)
'    If model Is Nothing Then
'        Exit Sub
'    End If
'
'    ' Retrieve the units for mass and length
'    Dim massUnit As String
'    Dim lengthUnit As String
'    massUnit = GetMassUnit(model)
'    lengthUnit = GetLengthUnit(model)
    
     ' If "WeightUOM" exists, remove it and replace with "Weight UOM"
'    If propertiesToSet.Exists("WeightUOM") Then
'        propertiesToSet.Remove "WeightUOM"
'    End If
    
    ' Set the custom properties in the propertiesToSet dictionary
'    propertiesToSet("Weight UOM") = massUnit
'    propertiesToSet("Dimensional UOM") = lengthUnit
'End Sub

Sub AutofillUnits(model As ModelDoc2, ByRef propertiesToSet As Object)
    If model Is Nothing Then
        Exit Sub
    End If
    
    ' Retrieve the units for mass and length
    Dim massUnit As String
    Dim lengthUnit As String
    massUnit = GetMassUnit(model)
    lengthUnit = GetLengthUnit(model)
    
     ' If "WeightUOM" exists, remove it and replace with "Weight UOM"
    If propertiesToSet.exists("WeightUOM") Then
        propertiesToSet.Remove "WeightUOM"
    End If
    
    ' Set the weight UOM
    propertiesToSet("Weight UOM") = massUnit
    
    ' Only set Dimensional UOM if it hasn't been converted to inches already
    If propertiesToSet.exists("Dimensional UOM") Then
        Dim currentUOM As String
        currentUOM = propertiesToSet("Dimensional UOM")
        If currentUOM = "in" Then
            DebugLog "AutofillUnits: Dimensional UOM already converted to inches, keeping 'in'"
        Else
            DebugLog "AutofillUnits: Updating Dimensional UOM from '" & currentUOM & "' to '" & lengthUnit & "'"
            propertiesToSet("Dimensional UOM") = lengthUnit
        End If
    Else
        DebugLog "AutofillUnits: Setting initial Dimensional UOM to '" & lengthUnit & "'"
        propertiesToSet("Dimensional UOM") = lengthUnit
    End If
End Sub



' Wrapper for centralized unit function in ConversionUtilitiesModule
Function GetLengthUnit(model As ModelDoc2) As String
    GetLengthUnit = GetModelUnitOfMeasure(model)
End Function

Function GetMassUnit(model As ModelDoc2) As String
    Dim unitType As Integer
    unitType = model.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swUnitsMassPropMass)
    Select Case unitType
        Case swUnitsMassPropMass_Kilograms
            GetMassUnit = "kg"
        Case swUnitsMassPropMass_Grams
            GetMassUnit = "g"
        'Case swUnitsMassPropMass_Tonnes
        '    GetMassUnit = "tonne"
        Case swUnitsMassPropMass_Pounds
            GetMassUnit = "lb"
        'Case swUnitsMassPropMass_Ounces
        '    GetMassUnit = "oz"
        Case Else
            GetMassUnit = "Unknown"
    End Select
End Function
