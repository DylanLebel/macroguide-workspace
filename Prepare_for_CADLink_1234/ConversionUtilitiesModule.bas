Attribute VB_Name = "ConversionUtilitiesModule"
Option Explicit
Option Private Module

' =============================================================================
' CONVERSION UTILITIES MODULE
' =============================================================================
' This module contains all unit conversion functions used across the application
' =============================================================================

' Convert from various units to inches
Function ConvertModelUnitsToInches(value As Double, fromUnit As String) As Double
    ' Convert from various units to inches
    Select Case LCase(fromUnit)
        Case "m"  ' Meters to inches
            ConvertModelUnitsToInches = value * 39.3701
        Case "cm" ' Centimeters to inches
            ConvertModelUnitsToInches = value * 0.393701
        Case "mm" ' Millimeters to inches
            ConvertModelUnitsToInches = value * 0.0393701
        Case "um" ' Micrometers to inches
            ConvertModelUnitsToInches = value * 0.0000393701
        Case "nm" ' Nanometers to inches
            ConvertModelUnitsToInches = value * 0.0000000393701
        Case "in" ' Already in inches
            ConvertModelUnitsToInches = value
        Case "ft" ' Feet to inches
            ConvertModelUnitsToInches = value * 12
        Case "ft-in" ' Feet-inches (already in inches)
            ConvertModelUnitsToInches = value
        Case Else
            ' Unknown unit, return as-is with warning
            DebugLog "WARNING: Unknown unit in ConvertModelUnitsToInches: " & fromUnit
            ConvertModelUnitsToInches = value
    End Select
End Function

' Convert from inches to various model units
Function ConvertInchesToModelUnits(inches As Double, modelUOM As String) As Double
    Dim result As Double
    
    Select Case LCase(modelUOM)
        Case "m"  ' Inches to meters
            result = inches * 25.4 / 1000
        Case "cm" ' Inches to centimeters
            result = inches * 2.54
        Case "mm" ' Inches to millimeters
            result = inches * 25.4
        Case "in" ' Already inches
            result = inches
        Case "ft" ' Inches to feet
            result = inches / 12
        Case "ft-in" ' Feet-inches (keep as inches)
            result = inches
        Case Else
            ' Unknown unit, return as-is with warning
            DebugLog "WARNING: Unknown unit in ConvertInchesToModelUnits: " & modelUOM
            result = inches
    End Select
    
    ' Round the result for metric units to nearest whole number
    If IsMetricUnit(modelUOM) Then
        result = Round(result, 0)
    End If
    
    ConvertInchesToModelUnits = result
End Function

' Helper function to determine if a unit is metric
Function IsMetricUnit(unitStr As String) As Boolean
    Select Case LCase(unitStr)
        Case "m", "cm", "mm", "um", "nm"
            IsMetricUnit = True
        Case Else
            IsMetricUnit = False
    End Select
End Function

' Get the model's unit of measure from SolidWorks
Function GetModelUnitOfMeasure(model As ModelDoc2) As String
    Dim unitType As Integer
    unitType = model.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swUnitsLinear)
    
    Select Case unitType
        Case 0 ' swMM
            GetModelUnitOfMeasure = "mm"
        Case 1 ' swCM
            GetModelUnitOfMeasure = "cm"
        Case 2 ' swMETER
            GetModelUnitOfMeasure = "m"
        Case 3 ' swINCHES
            GetModelUnitOfMeasure = "in"
        Case 4 ' swFEET
            GetModelUnitOfMeasure = "ft"
        Case 5 ' swFEETINCHES
            GetModelUnitOfMeasure = "ft-in"
        Case Else
            GetModelUnitOfMeasure = "unknown"
    End Select
End Function

' Helper function specifically for structural member length conversion
Function ConvertStructuralMemberLengthToInches(model As ModelDoc2, lengthValue As Double) As Double
    Dim modelUOM As String
    
    ' Get the model's unit of measure
    modelUOM = GetModelUnitOfMeasure(model)
    
    ' Convert to inches
    ConvertStructuralMemberLengthToInches = ConvertModelUnitsToInches(lengthValue, modelUOM)
    
    ' Debug output
    DebugLog "Converted " & lengthValue & " " & modelUOM & " to " & ConvertStructuralMemberLengthToInches & " inches"
End Function

' Convert a cutlist length value to inches, handling string cleanup
Function ConvertCutlistLengthToInches(cutlistLength As String, modelUOM As String) As String
    ' Clean the string by removing known unit suffixes to isolate the number
    Dim lengthStr As String
    lengthStr = cutlistLength
    lengthStr = Replace(lengthStr, "mm", "", 1, -1, vbTextCompare)
    lengthStr = Replace(lengthStr, "cm", "", 1, -1, vbTextCompare)
    lengthStr = Replace(lengthStr, "m", "", 1, -1, vbTextCompare)
    lengthStr = Replace(lengthStr, "in", "", 1, -1, vbTextCompare)
    lengthStr = Replace(lengthStr, "ft", "", 1, -1, vbTextCompare)
    lengthStr = Trim(lengthStr)

    ' Check if the cleaned string is a valid number
    If IsNumeric(lengthStr) Then
        Dim lengthValue As Double
        lengthValue = CDbl(lengthStr)

        ' If the model is metric, convert the length value to inches
        If IsMetricUnit(modelUOM) Then
            lengthValue = ConvertModelUnitsToInches(lengthValue, modelUOM)
        End If

        ' Format the final inch value as a clean decimal string
        ConvertCutlistLengthToInches = RemoveTrailingZeros(Format(lengthValue, "0.######"))
    Else
        ' If the string is still not numeric, return the original value
        ConvertCutlistLengthToInches = cutlistLength
    End If
End Function

' Remove trailing zeros from a number string
Function RemoveTrailingZeros(ByVal strNumber As String) As String
    If InStr(strNumber, ".") > 0 Then
        Do While Right(strNumber, 1) = "0"
            strNumber = Left(strNumber, Len(strNumber) - 1)
        Loop
        If Right(strNumber, 1) = "." Then
            strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    RemoveTrailingZeros = strNumber
End Function
