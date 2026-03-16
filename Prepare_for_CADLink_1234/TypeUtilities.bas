Attribute VB_Name = "TypeUtilities"
Option Explicit
Option Private Module

' ====================================================================================
' TypeUtilities.bas - Utility Functions for Type Processing
' ====================================================================================
'
' PURPOSE:
'   Contains helper functions for dimension conversion, fraction handling,
'   number formatting, and cutlist operations used by TypeModule.bas
'   and shape processors.
'
' EXTRACTED FROM:
'   TypeModule.bas (lines 4746-5318, 4936-5198)
'
' ====================================================================================

' ===========================
' PROPERTY VALUE HELPERS
' ===========================

' Get property value safely from dictionary
Function GetPropertyValue(ByRef propertiesToSet As Object, propName As String) As String
    If propertiesToSet.exists(propName) Then
        GetPropertyValue = CStr(propertiesToSet(propName))
    Else
        GetPropertyValue = ""
    End If
End Function

' Get custom property value from SolidWorks Custom Property Manager
Function GetCustomPropertyValue(custPropMgr As SldWorks.customPropertyManager, propName As String) As String
    Dim valOut As String
    Dim valResolved As String
    Dim wasResolved As Boolean

    If custPropMgr Is Nothing Then
        GetCustomPropertyValue = ""
        Exit Function
    End If

    custPropMgr.Get6 propName, False, valOut, valResolved, wasResolved, False

    If wasResolved Then
        GetCustomPropertyValue = valResolved
    Else
        GetCustomPropertyValue = ""
    End If
End Function

' Check if a type is valid in the given array
Function IsValidType(ByVal typeValue As String, ByRef validTypes As Variant) As Boolean
    Dim i As Integer
    For i = 0 To UBound(validTypes)
        If typeValue = validTypes(i) Then
            IsValidType = True
            Exit Function
        End If
    Next i
    IsValidType = False
End Function

' ===========================
' FRACTION/DECIMAL CONVERSION
' ===========================

' Extract numeric and symbol parts from a string like "1 1/2 in"
Function ExtractNumericAndSymbol(ByVal inputString As String, ByRef numericPart As Double, ByRef symbolPart As String) As Boolean
    Dim wholePart As Double
    Dim fractionPart As Double
    Dim currentNum As String
    Dim i As Integer

    wholePart = 0
    fractionPart = 0
    currentNum = ""
    i = 1

    ' Skip any leading spaces
    Do While i <= Len(inputString) And Mid(inputString, i, 1) = " "
        i = i + 1
    Loop

    ' Process whole number part
    Do While i <= Len(inputString)
        Dim currentChar As String
        currentChar = Mid(inputString, i, 1)

        Select Case currentChar
            Case "0" To "9", "."
                currentNum = currentNum & currentChar
            Case " "
                ' Space between whole number and fraction
                If currentNum <> "" Then
                    wholePart = CDbl(currentNum)
                    currentNum = ""
                End If
            Case "/"
                ' Start of fraction
                If currentNum <> "" Then
                    Dim numerator As Double
                    Dim denominator As Double
                    numerator = CDbl(currentNum)
                    currentNum = ""
                    i = i + 1

                    ' Get denominator
                    Do While i <= Len(inputString) And IsNumeric(Mid(inputString, i, 1))
                        currentNum = currentNum & Mid(inputString, i, 1)
                        i = i + 1
                    Loop

                    If currentNum <> "" Then
                        denominator = CDbl(currentNum)
                        fractionPart = numerator / denominator
                    End If
                End If
                Exit Do
            Case Else
                ' Non-numeric character that's not a fraction - end of number
                If currentNum <> "" Then
                    wholePart = CDbl(currentNum)
                End If
                Exit Do
        End Select
        i = i + 1
    Loop

    ' If we have a remaining number and haven't processed a fraction
    If currentNum <> "" And fractionPart = 0 Then
        wholePart = CDbl(currentNum)
    End If

    ' Get the total numeric value
    numericPart = wholePart + fractionPart

    ' Everything after the number is the symbol part
    symbolPart = Trim(Mid(inputString, i))

    ExtractNumericAndSymbol = True
End Function

' Convert decimal to fraction string (e.g., 1.5 -> "1 1/2")
Function ConvertDecimalToFraction(decimalNumber As Double) As String
    Dim fraction As String
    Dim intPart As Long
    Dim fracPart As Double
    Dim gcdValue As Long
    Dim numerator As Long
    Dim denominator As Long
    Const tol As Double = 0.00001

    ' First check if value is very close to a standard fraction
    If IsCloseToStandardFraction(decimalNumber) Then
        ' Continue with fraction conversion logic
    ElseIf Abs(decimalNumber) <= 0.04 Then
        ' Only use decimal format for small values that are NOT close to standard fractions
        ConvertDecimalToFraction = RemoveTrailingZeros(FormatNumber(decimalNumber, 4))
        Exit Function
    End If

    ' Extract the integer part using Fix() for consistent handling of negative numbers
    intPart = Fix(decimalNumber)
    fracPart = decimalNumber - intPart

    ' If the fractional part is very small, return just the integer part
    If Abs(fracPart) < tol Then
        ConvertDecimalToFraction = CStr(intPart)
        Exit Function
    End If

    ' Set denominator to 16 to round to the nearest 1/16
    denominator = 16
    numerator = Round(fracPart * denominator)

    ' If numerator rounds to 0, return just the integer part
    If numerator = 0 Then
        ConvertDecimalToFraction = CStr(intPart)
        Exit Function
    End If

    ' If numerator rounds to denominator (i.e., 16/16), increment whole part
    If numerator = denominator Then
        ConvertDecimalToFraction = CStr(intPart + 1)
        Exit Function
    End If

    ' Simplify the fraction using GCD
    gcdValue = gcd(Abs(numerator), denominator)
    numerator = numerator / gcdValue
    denominator = denominator / gcdValue

    ' Construct the fraction string
    If intPart <> 0 Then
        fraction = CStr(intPart) & " " & numerator & "/" & denominator
    Else
        fraction = numerator & "/" & denominator
    End If

    ConvertDecimalToFraction = fraction
End Function

' Convert fraction string to decimal (e.g., "1 1/2" -> 1.5)
Function ConvertFractionToDecimal(fractionStr As String) As Double
    Dim parts() As String
    Dim wholePart As Double
    Dim numerator As Double
    Dim denominator As Double
    Dim fracParts() As String

    parts = Split(fractionStr, " ")

    If UBound(parts) = 0 Then
        ' Only fraction part or whole number
        If InStr(parts(0), "/") > 0 Then
            fracParts = Split(parts(0), "/")
            numerator = CDbl(fracParts(0))
            denominator = CDbl(fracParts(1))
            ConvertFractionToDecimal = numerator / denominator
        Else
            ' Whole number
            ConvertFractionToDecimal = CDbl(parts(0))
        End If
    ElseIf UBound(parts) = 1 Then
        ' Whole and fraction part
        wholePart = CDbl(parts(0))
        fracParts = Split(parts(1), "/")
        numerator = CDbl(fracParts(0))
        denominator = CDbl(fracParts(1))
        ConvertFractionToDecimal = wholePart + (numerator / denominator)
    Else
        ' Invalid format
        ConvertFractionToDecimal = 0
    End If
End Function

' Greatest common divisor
Function gcd(ByVal a As Long, ByVal b As Long) As Long
    If b = 0 Then
        gcd = a
    Else
        gcd = gcd(b, a Mod b)
    End If
End Function

' Check if a decimal value is close to a standard fraction
Function IsCloseToStandardFraction(decimalNumber As Double) As Boolean
    ' Define standard fractions that we want to convert (as decimals)
    Dim standardFractions As Variant
    standardFractions = Array( _
        1 / 32, 1 / 16, 3 / 32, 1 / 8, 5 / 32, 3 / 16, 7 / 32, 1 / 4, 9 / 32, 5 / 16, 11 / 32, 3 / 8, _
        13 / 32, 7 / 16, 15 / 32, 1 / 2, 17 / 32, 9 / 16, 19 / 32, 5 / 8, 21 / 32, 11 / 16, 23 / 32, 3 / 4, _
        25 / 32, 13 / 16, 27 / 32, 7 / 8, 29 / 32, 15 / 16, 31 / 32 _
    )

    ' Tolerance for comparison (about 1/128th of an inch)
    Const tolerance As Double = 0.008

    ' Check if the decimal is close to any standard fraction
    Dim i As Integer
    For i = LBound(standardFractions) To UBound(standardFractions)
        If Abs(decimalNumber - standardFractions(i)) < tolerance Then
            IsCloseToStandardFraction = True
            Exit Function
        End If
    Next i

    ' Also check for whole numbers plus fractions
    Dim wholePart As Double
    Dim fracPart As Double
    wholePart = Fix(decimalNumber)
    fracPart = decimalNumber - wholePart

    ' Check if the fractional part is close to a standard fraction
    For i = LBound(standardFractions) To UBound(standardFractions)
        If Abs(fracPart - standardFractions(i)) < tolerance Then
            IsCloseToStandardFraction = True
            Exit Function
        End If
    Next i

    ' Not close to any standard fraction
    IsCloseToStandardFraction = False
End Function

' ===========================
' NUMBER FORMATTING
' ===========================

' Format clean number (remove unnecessary decimals)
Function FormatCleanNumber(value As Double) As String
    If value = Int(value) Then
        FormatCleanNumber = CStr(Int(value))
    Else
        FormatCleanNumber = Format(value, "0.####")
    End If
End Function

' Clean floating point errors by rounding to nearest 1/64
Function CleanFloatingPointErrors(value As Double) As Double
    Dim rounded As Double
    rounded = Round(value * 64) / 64

    ' If the difference is tiny (floating point error), use the rounded value
    If Abs(value - rounded) < 0.0000001 Then
        CleanFloatingPointErrors = rounded
    Else
        CleanFloatingPointErrors = value
    End If
End Function

' Format diameter to show properly
Function FormatDiameter(ByVal diameter As Double) As String
    If diameter = Fix(diameter) Then
        FormatDiameter = Format(diameter, "0")
    Else
        FormatDiameter = Format(diameter, "0.00")
    End If
End Function

' Check if a dimension exists in a string array of parts
Function IsDimensionInStockSize(ByVal dimension As Double, ByVal parts As Variant) As Boolean
    Dim i As Integer
    Dim numericPart As Double
    Dim symbolPart As String
    Dim result As Boolean

    For i = 0 To UBound(parts)
        result = ExtractNumericAndSymbol(parts(i), numericPart, symbolPart)
        If result Then
            If FormatCleanNumber(numericPart) = FormatCleanNumber(dimension) Then
                IsDimensionInStockSize = True
                Exit Function
            End If
        End If
    Next i
    IsDimensionInStockSize = False
End Function

' Sort array of fraction strings
Sub BubbleSortFractions(arr() As String)
    Dim i As Integer, j As Integer
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If ConvertFractionToDecimal(arr(i)) > ConvertFractionToDecimal(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' ===========================
' THICKNESS ROUNDING
' ===========================

' Round up thickness to standard plate size
Function RoundUpThickness(ByVal thickness As Double) As Double
    Dim allowedThicknesses As Variant
    allowedThicknesses = Array(3 / 16, 1 / 4, 5 / 16, 3 / 8, 7 / 16, 1 / 2, 9 / 16, 5 / 8, 3 / 4, 7 / 8, 15 / 16, 1, 1.125, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 3.25, 3.5, 3.75, 4)

    ' If thickness is less than minimum standard (3/16), return original thickness
    If thickness < 3 / 16 Then
        RoundUpThickness = thickness
        Exit Function
    End If

    Dim i As Integer
    For i = LBound(allowedThicknesses) To UBound(allowedThicknesses)
        If allowedThicknesses(i) >= thickness Then
            RoundUpThickness = allowedThicknesses(i)
            Exit Function
        End If
    Next i

    ' If thickness is larger than all standard sizes, return the original thickness
    RoundUpThickness = thickness
End Function

' Round up round bar diameter to standard size
Function RoundUpThicknessRB(ByVal diameter As Double) As Double
    Dim allowedDiameters As Variant
    allowedDiameters = Array( _
        0.125, 0.1875, 0.25, 0.3125, 0.375, 0.4375, 0.5, 0.5625, 0.625, 0.75, _
        0.875, 1, 1.125, 1.25, 1.375, 1.5, 1.625, 1.75, 1.875, 2, 2.125, _
        2.25, 2.375, 2.5, 2.625, 2.75, 2.875, 3, 3.125, 3.25, 3.375, 3.5, _
        3.625, 3.75, 3.875, 4, 4.125, 4.25, 4.375, 4.5, 4.625, 4.75, 4.875, _
        5, 5.125, 5.25, 5.375, 5.5, 5.625, 5.75, 5.875, 6, 6.125, 6.25, _
        6.375, 6.5, 6.625, 6.75, 6.875, 7, 7.125, 7.25, 7.375, 7.5, 7.625, _
        7.75, 7.875, 8)

    ' First check if the diameter is already in the array
    Dim i As Integer
    For i = LBound(allowedDiameters) To UBound(allowedDiameters)
        If Abs(allowedDiameters(i) - diameter) < 0.0001 Then
            RoundUpThicknessRB = allowedDiameters(i)
            Exit Function
        End If
    Next i

    ' If not an exact match, find the next larger diameter
    For i = LBound(allowedDiameters) To UBound(allowedDiameters)
        If allowedDiameters(i) >= diameter Then
            RoundUpThicknessRB = allowedDiameters(i)
            Exit Function
        End If
    Next i

    RoundUpThicknessRB = diameter
End Function

' ===========================
' CUTLIST OPERATIONS
' ===========================

' Get the original Mtl Part Number directly from cutlist
Function GetOriginalMtlPartNumberFromCutlist(model As ModelDoc2) As String
    GetOriginalMtlPartNumberFromCutlist = ""

    If model Is Nothing Then
        Exit Function
    End If

    Dim swFeature As SldWorks.Feature
    Set swFeature = model.FirstFeature

    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As SldWorks.customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager

            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean

                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And resolvedValueOut = "S" Then
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, resolvedValueOut)
                    If bRet And resolvedValueOut <> "" Then
                        GetOriginalMtlPartNumberFromCutlist = resolvedValueOut
                        Exit Function
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
End Function

' Set Mtl Unit Qty from LengthA for non PL/CP parts
Sub SetMtlUnitQtyFromLengthA(ByRef propertiesToSet As Object)
    ' Skip if Type is P
    If propertiesToSet.exists("Type") Then
        If UCase(Trim(propertiesToSet("Type"))) = "P" Then
            Exit Sub
        End If
    End If

    If propertiesToSet.exists("LengthA") Then
        If propertiesToSet.exists("Reference Category") Then
            If propertiesToSet("Reference Category") <> "PL" And propertiesToSet("Reference Category") <> "CP" Then
                propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
            End If
        Else
            propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
        End If
    End If
End Sub

' Set MtlUOM based on Reference Category and Dimensional UOM
Sub SetMtlUOMBasedOnType(ByRef propertiesToSet As Object)
    If propertiesToSet.exists("Reference Category") Then
        Dim refCategory As String
        refCategory = propertiesToSet("Reference Category")

        Dim dimensionalUOM As String
        If propertiesToSet.exists("Dimensional UOM") Then
            dimensionalUOM = propertiesToSet("Dimensional UOM")
        Else
            dimensionalUOM = "in"
        End If

        If refCategory = "PL" Or refCategory = "CP" Then
            If IsMetricUnit(dimensionalUOM) Then
                propertiesToSet("MtlUOM") = "SQMM"
            Else
                propertiesToSet("MtlUOM") = "SQIN"
            End If
        Else
            If IsMetricUnit(dimensionalUOM) Then
                propertiesToSet("MtlUOM") = "MM"
            Else
                propertiesToSet("MtlUOM") = "IN"
            End If
        End If
    Else
        propertiesToSet("MtlUOM") = "IN"
    End If
End Sub

' Exclude a cutlist folder from the BOM
Sub ExcludeFromCutlist(model As ModelDoc2, cutListFolderName As String)
    ' OPTIMIZATION: Removed unused swApp variable - not needed for this operation
    Dim swFeat As SldWorks.Feature
    Dim swBodyFolder As SldWorks.BodyFolder

    Set swFeat = model.FirstFeature

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" And swFeat.Name = cutListFolderName Then
            Set swBodyFolder = swFeat.GetSpecificFeature2

            If Not swBodyFolder Is Nothing Then
                Dim vBodies As Variant
                vBodies = swBodyFolder.GetBodies

                If Not IsEmpty(vBodies) Then
                    swFeat.ExcludeFromCutlist = True
                End If
            End If

            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

' Exclude ALL cutlist folders from BOM
Sub ExcludeAllCutlistFolders(model As ModelDoc2)
    ' OPTIMIZATION: Removed unused swApp variable - not needed for this operation
    Dim swFeat As SldWorks.Feature
    Dim swBodyFolder As SldWorks.BodyFolder

    Set swFeat = model.FirstFeature

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Set swBodyFolder = swFeat.GetSpecificFeature2

            If Not swBodyFolder Is Nothing Then
                Dim vBodies As Variant
                vBodies = swBodyFolder.GetBodies

                If Not IsEmpty(vBodies) Then
                    swFeat.ExcludeFromCutlist = True
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

' Process structural member dimensions from stock size string
Sub ProcessStructuralMemberDimensions(ByVal stockSize As String, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    Dim stockSizeParts() As String
    stockSizeParts = Split(stockSize, "x")

    Dim i As Integer
    For i = 0 To UBound(stockSizeParts)
        Dim numericPart As Double
        Dim symbolPart As String

        If ExtractNumericAndSymbol(Trim(stockSizeParts(i)), numericPart, symbolPart) Then
            Dim formattedValue As String
            If numericPart = Int(numericPart) Then
                formattedValue = CStr(Int(numericPart))
            Else
                formattedValue = RemoveTrailingZeros(Format(numericPart, "0.#####"))
            End If

            stockSizeParts(i) = formattedValue & symbolPart
        End If
    Next i

    stockSize = Join(stockSizeParts, " x ")
    propertiesToSet("Stock Size") = stockSize
End Sub

' Log cutlist properties for debugging (minimal version)
Sub LogCutlistProperties(swCustPropMgr As customPropertyManager, featureName As String)
    If swCustPropMgr Is Nothing Then
        Exit Sub
    End If
    ' Logging removed - use Logger.LogDebug if needed
End Sub
