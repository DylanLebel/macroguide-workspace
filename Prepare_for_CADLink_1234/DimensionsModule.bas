Attribute VB_Name = "DimensionsModule"
Option Explicit
Option Private Module

' Global variables
'Public Const DECIMAL_PLACES As Integer = 0  ' Used when USE_DECIMALS is True
Public USE_DECIMALS As Boolean  ' Flag to enable/disable use of decimals
Public DECIMAL_PLACES As Integer



Public FRACTION_DENOMINATOR As Long



' Initialization sub to be called when the module is loaded
'Sub InitializeModule()
'    USE_DECIMALS = False  ' Set the default value
'End Sub
Sub AddDimensionsProperties(model As ModelDoc2, ByRef propertiesToSet As Object)
    ' === ENTRY POINT DEBUGGING ===
    DebugLog "##################################################"
    DebugLog "### ENTERING AddDimensionsProperties FUNCTION ###"
    DebugLog "##################################################"
    DebugLog "Timestamp: " & Now()
    
    ' Check if model parameter is valid
    DebugLog "=== PARAMETER VALIDATION ==="
    DebugLog "model Is Nothing: " & (model Is Nothing)
    DebugLog "propertiesToSet Is Nothing: " & (propertiesToSet Is Nothing)
    
    If model Is Nothing Then
        DebugLog "ERROR: Model is Nothing - EXITING FUNCTION"
        Exit Sub
    End If
    
    ' Get model information
    DebugLog "=== MODEL INFORMATION ==="
    DebugLog "Model Type: " & model.GetType
    DebugLog "Model Type Names: 1=Part, 2=Assembly, 3=Drawing"
    DebugLog "Model Title: " & model.GetTitle
    DebugLog "Model Path: " & model.GetPathName
    
    ' Test GetDocumentUnits call
    DebugLog "=== CALLING GetDocumentUnits ==="
    On Error GoTo GetDocumentUnitsError
    Call GetDocumentUnits
    DebugLog "GetDocumentUnits completed successfully"
    GoTo ContinueAfterGetDocumentUnits
    
GetDocumentUnitsError:
    DebugLog "ERROR in GetDocumentUnits: " & Err.description
    DebugLog "Err.Number: " & Err.Number
    Resume ContinueAfterGetDocumentUnits
    
ContinueAfterGetDocumentUnits:
    On Error GoTo 0  ' Reset error handling
    
    ' OPTIMIZATION: Use global swApp accessor instead of creating local instance
    DebugLog "=== USING GLOBAL SOLIDWORKS APP ==="
    Dim swAppLocal As SldWorks.SldWorks
    Set swAppLocal = GetSwApp()
    DebugLog "SolidWorks App initialized: " & Not (swAppLocal Is Nothing)
    
    ' Log document units
    DebugLog "=== CALLING LogDocumentPrecision ==="
    On Error GoTo LogDocumentPrecisionError
    LogDocumentPrecision model
    DebugLog "LogDocumentPrecision completed successfully"
    GoTo ContinueAfterLogDocumentPrecision
    
LogDocumentPrecisionError:
    DebugLog "ERROR in LogDocumentPrecision: " & Err.description
    DebugLog "Err.Number: " & Err.Number
    Resume ContinueAfterLogDocumentPrecision
    
ContinueAfterLogDocumentPrecision:
    On Error GoTo 0  ' Reset error handling
    
    ' Get the document name
    DebugLog "=== GETTING DOCUMENT NAME ==="
    Dim docName As String
    docName = model.GetTitle
    DebugLog "Document Name: '" & docName & "'"
    DebugLog "Document Name Length: " & Len(docName)
    DebugLog "Document Name is empty: " & (docName = "")
    
    If docName = "" Then
        DebugLog "ERROR: Document name is empty - EXITING FUNCTION"
        Exit Sub
    End If
    
    ' Model type check and processing
    DebugLog "=== MODEL TYPE PROCESSING ==="
    DebugLog "About to enter Select Case with model.GetType = " & model.GetType
    DebugLog "swDocPART constant = " & swDocPART
    DebugLog "swDocASSEMBLY constant = " & swDocASSEMBLY
    DebugLog "Comparison: model.GetType = swDocPART is " & (model.GetType = swDocPART)
    
    Select Case model.GetType
        Case swDocPART
            DebugLog "*** ENTERED swDocPART CASE ***"
            DebugLog "About to call GetDimensions function"
            
            ' Process parts - call GetDimensions
            Dim swDimensions As Variant
            DebugLog "Declared swDimensions variable"
            
            ' Add error handling for GetDimensions call
            On Error GoTo GetDimensionsError
            DebugLog "Calling GetDimensions now..."
            swDimensions = GetDimensions(model, propertiesToSet)
            DebugLog "GetDimensions call completed"
            GoTo ContinueAfterGetDimensions
            
GetDimensionsError:
            DebugLog "ERROR in GetDimensions: " & Err.description
            DebugLog "Err.Number: " & Err.Number
            DebugLog "GetDimensions call failed - continuing without dimensions"
            swDimensions = Empty
            Resume ContinueAfterGetDimensions
            
ContinueAfterGetDimensions:
            On Error GoTo 0  ' Reset error handling
            
            DebugLog "=== CHECKING GetDimensions RESULTS ==="
            DebugLog "IsEmpty(swDimensions): " & IsEmpty(swDimensions)
            
            If Not IsEmpty(swDimensions) Then
                DebugLog "swDimensions is NOT empty"
                DebugLog "=== Dimension Processing Start ==="
                DebugLog "Dimensions Received:"
                DebugLog "swDimensions(0) (Length): " & swDimensions(0)
                DebugLog "swDimensions(1) (Width): " & swDimensions(1)
                DebugLog "swDimensions(2) (Height): " & swDimensions(2)
                
                ' *** KEY FIX: Check what units the dimensions are CURRENTLY in ***
                Dim originalModelUOM As String
                Dim currentDimensionUOM As String
                
                originalModelUOM = GetModelUnitOfMeasure(model)  ' Original model units (mm, cm, etc.)
                
                ' Check if GetDimensions converted the dimensions to inches
                If propertiesToSet.exists("Dimensional UOM") Then
                    currentDimensionUOM = propertiesToSet("Dimensional UOM")
                Else
                    currentDimensionUOM = originalModelUOM
                End If
                
                DebugLog "=== UNIT ANALYSIS ==="
                DebugLog "Original Model UOM: " & originalModelUOM
                DebugLog "Current Dimension UOM: " & currentDimensionUOM
                
                
                
                ' *** CHECK PATH-BASED UNIT OVERRIDE ***
Dim pathOverride As String
pathOverride = GetPathBasedUnitOverride(model)

If pathOverride <> "" Then
    DebugLog "=== PATH-BASED OVERRIDE DETECTED IN AddDimensionsProperties ==="
    DebugLog "Override Unit: " & pathOverride
    DebugLog "Current Dimension UOM: " & currentDimensionUOM
    
    ' Use temporary variables for conversion
    Dim tempLength As Double, tempWidth As Double, tempHeight As Double
    tempLength = CDbl(swDimensions(0))
    tempWidth = CDbl(swDimensions(1))
    tempHeight = CDbl(swDimensions(2))
    
    ' Apply conversion if needed
    If pathOverride = "mm" And currentDimensionUOM <> "mm" Then
        DebugLog "*** CONVERTING TO MM PER PATH OVERRIDE ***"
        
        If currentDimensionUOM = "in" Then
            tempLength = tempLength * 25.4
            tempWidth = tempWidth * 25.4
            tempHeight = tempHeight * 25.4
        ElseIf currentDimensionUOM = "cm" Then
            tempLength = tempLength * 10
            tempWidth = tempWidth * 10
            tempHeight = tempHeight * 10
        ElseIf currentDimensionUOM = "m" Then
            tempLength = tempLength * 1000
            tempWidth = tempWidth * 1000
            tempHeight = tempHeight * 1000
        End If
        
        currentDimensionUOM = "mm"
        propertiesToSet("Dimensional UOM") = "mm"
        DebugLog "Converted dimensions to MM"
        
    ElseIf pathOverride = "in" And currentDimensionUOM <> "in" Then
        DebugLog "*** CONVERTING TO INCHES PER PATH OVERRIDE ***"
        
        tempLength = ConvertModelUnitsToInches(tempLength, currentDimensionUOM)
        tempWidth = ConvertModelUnitsToInches(tempWidth, currentDimensionUOM)
        tempHeight = ConvertModelUnitsToInches(tempHeight, currentDimensionUOM)
        
        currentDimensionUOM = "in"
        propertiesToSet("Dimensional UOM") = "in"
        DebugLog "Converted dimensions to INCHES"
    End If
    
    ' Assign converted values back to the array
    swDimensions(0) = tempLength
    swDimensions(1) = tempWidth
    swDimensions(2) = tempHeight
    
    DebugLog "Final dimensions after override: L=" & swDimensions(0) & ", W=" & swDimensions(1) & ", H=" & swDimensions(2) & " (" & currentDimensionUOM & ")"
End If
                
                DebugLog "Were dimensions converted to inches: " & (currentDimensionUOM = "in" And originalModelUOM <> "in")
                DebugLog "Processing dimensions as: " & currentDimensionUOM
                
                
                DebugLog "About to enter dimension formatting - currentDimensionUOM = '" & currentDimensionUOM & "'"
                
                
                ' *** CRITICAL FIX: Process based on current dimension units, not original model units ***
                If currentDimensionUOM = "in" Then
                    DebugLog "--- IMPERIAL PROCESSING (dimensions are in inches) ---"
                    ' Detailed debugging for each conversion
                    DebugLog "Converting Length: " & CDbl(swDimensions(0))
                    Dim lengthFraction As String
                    lengthFraction = ConvertToFraction(CDbl(swDimensions(0)), propertiesToSet, False)
                    propertiesToSet("Length") = lengthFraction
                    DebugLog "Imperial Length Fraction: " & lengthFraction
                    
                    DebugLog "Converting Width: " & CDbl(swDimensions(1))
                    Dim widthFraction As String
                    widthFraction = ConvertToFraction(CDbl(swDimensions(1)), propertiesToSet, False)
                    propertiesToSet("Width") = widthFraction
                    DebugLog "Imperial Width Fraction: " & widthFraction
                    
                    DebugLog "Converting Height: " & CDbl(swDimensions(2))
                    Dim heightFraction As String
                    heightFraction = ConvertToFraction(CDbl(swDimensions(2)), propertiesToSet, True)  ' Treat Height as thickness
                    propertiesToSet("Height") = heightFraction
                    DebugLog "Imperial Height Fraction: " & heightFraction
                    
                    ' Set LengthA (always in decimal form for imperial)
                    Dim lengthADecimal As String
                    lengthADecimal = RemoveTrailingZeros(FormatNumber(CDbl(swDimensions(0)), 4))
                    propertiesToSet("LengthA") = lengthADecimal
                    DebugLog "Imperial LengthA Decimal: " & lengthADecimal
                    
                    DebugLog "--- END IMPERIAL PROCESSING ---"
                    
ElseIf IsMetricUnit(currentDimensionUOM) Then
    ' Check if this is a structural member that needs imperial conversion
    Dim isStructuralMember As Boolean
    Dim needsImperialConversion As Boolean
    
    isStructuralMember = CheckIfStructuralMember(model)
    needsImperialConversion = isStructuralMember And Not IsASNZSStructuralMember(model)
    
    If needsImperialConversion Then
        DebugLog "--- METRIC STRUCTURAL MEMBER WITH IMPERIAL ORDERING - CONVERTING TO INCHES ---"
        
        ' Convert dimensions to inches
        Dim lengthInches As Double, widthInches As Double, heightInches As Double
        lengthInches = ConvertModelUnitsToInches(CDbl(swDimensions(0)), currentDimensionUOM)
        widthInches = ConvertModelUnitsToInches(CDbl(swDimensions(1)), currentDimensionUOM)
        heightInches = ConvertModelUnitsToInches(CDbl(swDimensions(2)), currentDimensionUOM)
        
        ' Set properties in imperial format
        propertiesToSet("Length") = ConvertToFraction(lengthInches, propertiesToSet, False)
        propertiesToSet("LengthA") = RemoveTrailingZeros(FormatNumber(lengthInches, 4))
        propertiesToSet("Width") = ConvertToFraction(widthInches, propertiesToSet, False)
        propertiesToSet("Height") = ConvertToFraction(heightInches, propertiesToSet, True)
        
        ' Update Dimensional UOM
        propertiesToSet("Dimensional UOM") = "in"
        
        DebugLog "Converted Length: " & propertiesToSet("Length")
        DebugLog "Converted Width: " & propertiesToSet("Width")
        DebugLog "Converted Height: " & propertiesToSet("Height")
        DebugLog "Updated Dimensional UOM to: in"
        
    Else
        DebugLog "--- METRIC PROCESSING (dimensions are in " & currentDimensionUOM & ") ---"
        ' Standard metric processing
        Dim formattedLength As String
        formattedLength = FormatMetricDimension(CDbl(swDimensions(0)))
        propertiesToSet("Length") = formattedLength
        DebugLog "Metric Length: " & formattedLength
        
        propertiesToSet("LengthA") = formattedLength
        DebugLog "Metric LengthA: " & formattedLength
        
        propertiesToSet("Width") = FormatMetricDimension(CDbl(swDimensions(1)))
        DebugLog "Metric Width: " & propertiesToSet("Width")
        
        propertiesToSet("Height") = FormatMetricDimension(CDbl(swDimensions(2)))
        DebugLog "Metric Height: " & propertiesToSet("Height")
        
        DebugLog "--- END METRIC PROCESSING ---"
    End If
                Else
                    DebugLog "--- UNKNOWN UNIT SYSTEM - DEFAULTING TO IMPERIAL ---"
                    ' Default to imperial processing if units are unclear
                    lengthFraction = ConvertToFraction(CDbl(swDimensions(0)), propertiesToSet, False)
                    propertiesToSet("Length") = lengthFraction
                    DebugLog "Default Imperial Length Fraction: " & lengthFraction
                    
                    widthFraction = ConvertToFraction(CDbl(swDimensions(1)), propertiesToSet, False)
                    propertiesToSet("Width") = widthFraction
                    DebugLog "Default Imperial Width Fraction: " & widthFraction
                    
                    heightFraction = ConvertToFraction(CDbl(swDimensions(2)), propertiesToSet, True)
                    propertiesToSet("Height") = heightFraction
                    DebugLog "Default Imperial Height Fraction: " & heightFraction
                    
                    lengthADecimal = RemoveTrailingZeros(FormatNumber(CDbl(swDimensions(0)), 4))
                    propertiesToSet("LengthA") = lengthADecimal
                    DebugLog "Default Imperial LengthA Decimal: " & lengthADecimal
                End If
                
                DebugLog "=== Dimension Processing End ==="
            Else
                DebugLog "ERROR: swDimensions is EMPTY - no dimensions to process"
            End If
            
        Case swDocASSEMBLY
            DebugLog "*** ENTERED swDocASSEMBLY CASE ***"
            ' For all assembly types, set Stock Size to "SEE BOM"
            ' propertiesToSet("Stock Size") = "SEE BOM"
            
        Case Else
            DebugLog "*** ENTERED Case Else - Unknown document type ***"
            DebugLog "Document type: " & model.GetType
            ' Do nothing for other document types
    End Select
    
    DebugLog "=== REACHED END OF FUNCTION ==="
    DebugLog "About to execute final Stop statement"
    'Stop  ' DEBUGGING: Stop here to check dimension values in Immediate Window
    DebugLog "=== EXITING AddDimensionsProperties FUNCTION ==="
    
End Sub


Function FormatMetricDimension(value As Double) As String
    ' Round to the nearest whole number
    Dim roundedValue As Double
    roundedValue = Round(value, 0)
    
    ' Format the rounded value according to DECIMAL_PLACES
    If DECIMAL_PLACES = 0 Then
        FormatMetricDimension = CStr(CLng(roundedValue))
    Else
        FormatMetricDimension = Format(roundedValue, "0." & String(DECIMAL_PLACES, "0"))
    End If
End Function



Function RoundUp(value As Double) As Long
    RoundUp = CLng(-Int(-value))
End Function




Function GetDimensionalUOM(propertiesToSet As Scripting.Dictionary) As String
    If propertiesToSet.exists("Dimensional UOM") Then
        GetDimensionalUOM = propertiesToSet("Dimensional UOM")
    Else
        GetDimensionalUOM = "in" ' Default to inches if not specified
    End If
End Function

Function ConvertToFraction(decimalValue As Double, propertiesToSet As Object, Optional isThickness As Boolean = False) As String
    DebugLog "--- ConvertToFraction Debug Start ---"
    DebugLog "Input Decimal Value: " & decimalValue

    Const tol As Double = 0.00001
    Const DECIMAL_THRESHOLD As Double = 0.015  ' Increased threshold to catch 0.0100000054
    
    If Abs(decimalValue) < tol Then
        DebugLog "Value is effectively zero - returning empty string"
        ConvertToFraction = ""
        Exit Function
    End If
    
    ' Check if value is below threshold - if so, return decimal format
    If Abs(decimalValue) <= DECIMAL_THRESHOLD Then
        DebugLog "Value at or below threshold (" & DECIMAL_THRESHOLD & ") - returning decimal format"
        
        ' Use a more specific format for small decimals
        Dim decimalResult As String
        If decimalValue < 0.001 Then
            decimalResult = Format(decimalValue, "0.0000")
        ElseIf decimalValue < 0.01 Then
            decimalResult = Format(decimalValue, "0.000")
        Else
            decimalResult = Format(decimalValue, "0.00")
        End If
        
        ' Remove trailing zeros but ensure we don't lose significant digits
        ConvertToFraction = RemoveTrailingZerosCarefully(decimalResult)
        
        DebugLog "Formatted as: " & decimalResult
        DebugLog "After careful zero removal: " & ConvertToFraction
        DebugLog "--- ConvertToFraction Debug End ---"
        Exit Function
    End If

    Dim wholePart As Long
    Dim fractionPart As Double
    wholePart = Int(decimalValue)
    fractionPart = decimalValue - wholePart

    DebugLog "Whole Part: " & wholePart
    DebugLog "Fraction Part: " & fractionPart

    If Abs(fractionPart) < tol Then
        DebugLog "Fraction part too small - returning whole part"
        ConvertToFraction = CStr(wholePart)
        Exit Function
    End If

    ' Use the global FRACTION_DENOMINATOR variable
    Dim denominator As Long
    If FRACTION_DENOMINATOR > 0 Then
        denominator = FRACTION_DENOMINATOR
    Else
        denominator = 16 ' Default to 16 if not set
    End If
    
    DebugLog "Using denominator: " & denominator

    ' Round directly to nearest fraction using document's denominator
    Dim rawNumerator As Double
    Dim numerator As Long
    rawNumerator = fractionPart * denominator
    numerator = Round(rawNumerator)

    DebugLog "Raw numerator calculation: " & fractionPart & " � " & denominator & " = " & rawNumerator
    DebugLog "Rounded Numerator (to nearest 1/" & denominator & "): " & numerator

    ' FIXED: If numerator rounds to 0, just return the whole number
    If numerator = 0 Then
        DebugLog "Numerator is 0 - returning whole number: " & wholePart
        ConvertToFraction = CStr(wholePart)
        Exit Function
    End If

    If numerator = denominator Then
        DebugLog "Numerator equals denominator - incrementing whole part"
        ConvertToFraction = CStr(wholePart + 1)
        Exit Function
    End If

    ' Simplify
    Dim simplifiedNumerator As Long, simplifiedDenominator As Long
    Call SimplifyFraction(numerator, denominator, simplifiedNumerator, simplifiedDenominator)

    Dim result As String
    If simplifiedNumerator = 0 Then
        result = CStr(wholePart)
    Else
        result = IIf(wholePart > 0, CStr(wholePart) & " ", "") & simplifiedNumerator & "/" & simplifiedDenominator
    End If

    DebugLog "Final Result: " & result
    DebugLog "--- ConvertToFraction Debug End ---"

    ConvertToFraction = result
End Function

' Helper function for careful zero removal
Function RemoveTrailingZerosCarefully(value As String) As String
    Dim result As String
    result = value
    
    ' Only remove trailing zeros if there's a decimal point
    If InStr(result, ".") > 0 Then
        ' Remove trailing zeros after decimal point, but preserve significant digits
        Do While Right(result, 1) = "0" And Len(result) > InStr(result, ".") + 1
            result = Left(result, Len(result) - 1)
        Loop
        
        ' Remove decimal point if it's the last character
        If Right(result, 1) = "." Then
            result = Left(result, Len(result) - 1)
        End If
    End If
    
    ' Ensure we don't return an empty string
    If result = "" Or result = "." Then
        result = "0"
    End If
    
    RemoveTrailingZerosCarefully = result
End Function





Sub SimplifyFraction(ByVal numerator As Long, ByVal denominator As Long, ByRef simplifiedNumerator As Long, ByRef simplifiedDenominator As Long)
    ' Handle division by zero error for input denominator
    If denominator = 0 Then
        simplifiedNumerator = numerator ' Or perhaps handle as an error?
        simplifiedDenominator = 1       ' Default to 1 to avoid division by zero later
        DebugLog "SimplifyFraction Warning: Input denominator was 0."
        Exit Sub
    End If

    ' Handle zero numerator
    If numerator = 0 Then
        simplifiedNumerator = 0
        simplifiedDenominator = 1 ' Denominator doesn't matter if numerator is 0, use 1.
        Exit Sub
    End If

    Dim gcd As Long
    gcd = CalculateGCD(Abs(numerator), Abs(denominator))

    ' Defensive check for GCD result (shouldn't be 0 if inputs are valid)
    If gcd = 0 Then
        gcd = 1 ' Avoid division by zero if GCD somehow returned 0
        DebugLog "SimplifyFraction Warning: GCD calculation resulted in 0."
    End If

    ' Perform the simplification ONCE
    simplifiedNumerator = numerator \ gcd
    simplifiedDenominator = denominator \ gcd

    ' Ensure the final denominator is positive
    If simplifiedDenominator < 0 Then
        simplifiedNumerator = -simplifiedNumerator
        simplifiedDenominator = -simplifiedDenominator
    End If

    ' Optional: Add the debug print here before exiting
    DebugLog "SimplifyFraction returning: Num=" & simplifiedNumerator & ", Den=" & simplifiedDenominator

End Sub ' *** NO MORE CODE AFTER THIS LINE ***

Function CalculateGCD(a As Long, b As Long) As Long
    a = Abs(a)
    b = Abs(b)
    
    If a = 0 Then
        CalculateGCD = b
        Exit Function
    ElseIf b = 0 Then
        CalculateGCD = a
        Exit Function
    End If
    
    Dim temp As Long
    
    Do While b <> 0
        temp = b
        b = a Mod b
        a = temp
    Loop
    
    CalculateGCD = a
End Function




Function CheckForCheckerPattern(swModel As ModelDoc2) As Boolean
    Dim swModelExt As ModelDocExtension
    Dim swFeature As Feature
    Dim featureName As String
    Dim isCheckerPatternFound As Boolean: isCheckerPatternFound = False

    Set swModelExt = swModel.Extension
    Set swFeature = swModel.FirstFeature

    Do While Not swFeature Is Nothing
        featureName = swFeature.Name
        If InStr(1, featureName, "checker", vbTextCompare) > 0 Then
            isCheckerPatternFound = True
            Exit Do
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop

    CheckForCheckerPattern = isCheckerPatternFound
End Function

Function isPlate(swBody As Body2) As Boolean
    On Error GoTo ErrorHandler
    
    '''''debug.Print "isPlate function called"
    
    Dim swFace As Face2
    Dim vNormal As Variant
    Dim LargestArea1 As Double, LargestArea2 As Double
    Dim SmallestDimension As Double
    Dim MediumDimension As Double
    Dim LargestDimension As Double
    Dim faceArea As Double
    Dim isPlateCriteriaMet As Boolean
    Dim vBox As Variant
    
    isPlateCriteriaMet = False
    
    ' Get the bounding box of the body
    vBox = swBody.GetBodyBox()
    
    ' Calculate dimensions
    SmallestDimension = Abs(vBox(3) - vBox(0))
    MediumDimension = Abs(vBox(5) - vBox(2))
    LargestDimension = Abs(vBox(4) - vBox(1))
    
    ' Sort dimensions
    If SmallestDimension > MediumDimension Then SwapValues SmallestDimension, MediumDimension
    If MediumDimension > LargestDimension Then SwapValues MediumDimension, LargestDimension
    If SmallestDimension > MediumDimension Then SwapValues SmallestDimension, MediumDimension
    
    '''''debug.Print "Dimensions: " & SmallestDimension & " x " & MediumDimension & " x " & LargestDimension
    
    ' Criteria for a plate:
    ' 1. The smallest dimension (thickness) should be significantly smaller than the other two
    ' 2. The two larger dimensions should be significantly different from each other
    If (SmallestDimension * 5 < MediumDimension) And (SmallestDimension * 5 < LargestDimension) Then
        If (MediumDimension * 1.2 < LargestDimension) Or (MediumDimension > LargestDimension * 0.8) Then
            isPlateCriteriaMet = True
            '''''debug.Print "Is a plate: Criteria met"
        Else
            '''''debug.Print "Not a plate: Medium and largest dimensions too similar"
        End If
    Else
        '''''debug.Print "Not a plate: Smallest dimension not significantly smaller than others"
    End If
    
    isPlate = isPlateCriteriaMet
    Exit Function

ErrorHandler:
    '''''debug.Print "Error in isPlate function: " & Err.Description
    isPlate = False
End Function

Sub SwapValues(ByRef a As Double, ByRef b As Double)
    Dim temp As Double
    temp = a
    a = b
    b = temp
End Sub

' Simplified GetDimensions function that uses the model's native units
Function GetDimensions(model As ModelDoc2, ByRef propertiesToSet As Object) As Variant
    If model Is Nothing Then Exit Function
    
    DebugLog "=== ENTERING GetDimensions ==="
    
    ' Declare variables
    Dim swDimensions(2) As Double
    Dim vBox As Variant
    Dim swFeature As Feature
    Dim flatPatternFeature As Feature
    Dim wasFlatPatternSuppressed As Boolean
    Dim swPart As PartDoc
    Dim length As Double, width As Double, height As Double
    Dim dimensionalUOM As String
    Dim swBody As Body2
    Dim sortedDimensions(2) As Double
    Dim totalLength As Double
    Dim foundTotalLength As Boolean
    Dim isStructuralMember As Boolean
    Dim configs As Variant
    Dim i As Long
    Dim currentConfig As String
    Dim flatPatternConfig As String
    Dim usingConfig As Boolean
    Dim valueOut As String, resolvedValueOut As String
    Dim swCustPropMgr As customPropertyManager
    Dim foundValidCutlist As Boolean

    foundTotalLength = False
    isStructuralMember = CheckIfStructuralMember(model)
    foundValidCutlist = False
    
    ' Get model's native units
    dimensionalUOM = GetModelUnitOfMeasure(model)
    DebugLog "Model's native units: " & dimensionalUOM
    
    If model.GetType = swDocPART Then
        Set swPart = model
        currentConfig = swPart.GetActiveConfiguration.Name
        
        ' DEBUG: Check model unit system
        Dim unitSystem As Long
        unitSystem = model.GetUnits(0)
        DebugLog "=== MODEL UNIT SYSTEM ==="
        DebugLog "Unit System Code: " & unitSystem
        DebugLog "Unit System Name: " & GetUnitSystemName(unitSystem)
        DebugLog "Model UOM: " & dimensionalUOM
        DebugLog "========================="

        ' ===== SHEET METAL CUTLIST PROCESSING =====
        Dim hasFlatPattern As Boolean
        hasFlatPattern = False
        Set swFeature = swPart.FirstFeature
        Do While Not swFeature Is Nothing
            If swFeature.GetTypeName2 = "FlatPattern" Then
                hasFlatPattern = True
                Exit Do
            End If
            Set swFeature = swFeature.GetNextFeature
        Loop
        
        If hasFlatPattern Then
            DebugLog "Flat pattern exists. Forcing rebuild and cutlist update..."
            
            ' Force rebuild first
            model.ForceRebuild3 True
            
            ' Force update all cutlists through BodyFolder interface
            Dim swFeat As Feature
            Dim swBodyFolder As BodyFolder
            Set swFeat = model.FirstFeature
            Do While Not swFeat Is Nothing
                If swFeat.GetTypeName2 = "CutListFolder" Then
                    Set swBodyFolder = swFeat.GetSpecificFeature2
                    If Not swBodyFolder Is Nothing Then
                        swBodyFolder.UpdateCutList
                        DebugLog "Updated cutlist: " & swFeat.Name
                    End If
                End If
                Set swFeat = swFeat.GetNextFeature
            Loop
            
            ' Now read the cutlist properties - LOOP TO FIND CUTLISTFOLDER WITH VALID VALUES
            DebugLog "Checking cutlist properties..."
            Set swFeature = model.FirstFeature
            Do While Not swFeature Is Nothing
                If swFeature.GetTypeName2 = "CutListFolder" Then
                    Set swCustPropMgr = swFeature.customPropertyManager
                    
                    Dim boxLength As String, boxWidth As String, metalThickness As String
                    Dim boxLengthRaw As String, boxWidthRaw As String, metalThicknessRaw As String
                    
                    swCustPropMgr.Get4 "Bounding Box Length", False, boxLengthRaw, boxLength
                    swCustPropMgr.Get4 "Bounding Box Width", False, boxWidthRaw, boxWidth
                    swCustPropMgr.Get4 "Sheet Metal Thickness", False, metalThicknessRaw, metalThickness
                    
                    DebugLog "Checking cutlist: " & swFeature.Name
                    DebugLog "  boxLength: " & boxLength & ", boxWidth: " & boxWidth & ", metalThickness: " & metalThickness
                    
                    If boxLength <> "" And boxWidth <> "" And metalThickness <> "" Then
                        ' Check if we got formula strings instead of resolved values
                        If InStr(boxLength, "@") > 0 Or InStr(boxWidth, "@") > 0 Or InStr(metalThickness, "@") > 0 Then
                            DebugLog "  Skipping - got formula strings instead of resolved values"
                        Else
                            DebugLog "Found valid sheet metal cutlist properties."
                            foundValidCutlist = True
                            
                            ' Enhanced parsing for complex dimension formats like "1'-4 7/16""
                            boxLength = Trim(Replace(boxLength, """", ""))
                            boxWidth = Trim(Replace(boxWidth, """", ""))
                            metalThickness = Trim(Replace(metalThickness, """", ""))

                            ' Handle feet-inches format for length (like "1'-4 7/16")
                            If InStr(boxLength, "'") > 0 Then
                                Dim lengthParts() As String
                                lengthParts = Split(boxLength, "'")
                                Dim feet As Double
                                feet = val(lengthParts(0))
                                Dim inchPart As String
                                If UBound(lengthParts) > 0 Then
                                    inchPart = Trim(lengthParts(1))
                                    If Left(inchPart, 1) = "-" Then inchPart = Trim(Mid(inchPart, 2))
                                    ' Parse mixed number inches like "4 7/16"
                                    If InStr(inchPart, " ") > 0 Then
                                        Dim inchComponents() As String
                                        inchComponents = Split(inchPart, " ")
                                        Dim wholeInches As Double
                                        wholeInches = val(inchComponents(0))
                                        Dim fracInches As Double
                                        If UBound(inchComponents) > 0 And InStr(inchComponents(1), "/") > 0 Then
                                            Dim fracParts() As String
                                            fracParts = Split(inchComponents(1), "/")
                                            fracInches = val(fracParts(0)) / val(fracParts(1))
                                        End If
                                        length = feet * 12 + wholeInches + fracInches
                                    Else
                                        length = feet * 12 + FractionToDecimal(inchPart)
                                    End If
                                Else
                                    length = feet * 12
                                End If
                            Else
                                length = FractionToDecimal(boxLength)
                            End If

                            ' Parse width and thickness normally
                            width = FractionToDecimal(boxWidth)
                            height = FractionToDecimal(metalThickness)
                            
                            DebugLog "Parsed sheet metal dimensions: Length=" & length & ", Width=" & width & ", Height=" & height
                            DebugLog "Cutlist dimensions are already in model units (" & dimensionalUOM & ") - no conversion needed"
                            
                            ' Set the dimensional UOM to match the model
                            propertiesToSet("Dimensional UOM") = dimensionalUOM
                            
                            ' Sort and assign dimensions
                            sortedDimensions(0) = length
                            sortedDimensions(1) = width
                            sortedDimensions(2) = height
                            Call BubbleSort(sortedDimensions)
                            
                            swDimensions(1) = sortedDimensions(0)
                            swDimensions(2) = sortedDimensions(1)
                            swDimensions(0) = sortedDimensions(2)
                            
                            DebugLog "=== FINAL SHEET METAL DIMENSIONS ==="
                            DebugLog "swDimensions(0) [Length]: " & swDimensions(0) & " " & dimensionalUOM
                            DebugLog "swDimensions(1) [Width]: " & swDimensions(1) & " " & dimensionalUOM
                            DebugLog "swDimensions(2) [Height]: " & swDimensions(2) & " " & dimensionalUOM
                            
                            GetDimensions = swDimensions
                            Exit Function
                        End If
                    End If
                End If
                Set swFeature = swFeature.GetNextFeature
            Loop
            
            If Not foundValidCutlist Then
                DebugLog "No valid cutlist properties found - falling back to bounding box method"
            End If
        End If
        
        ' ===== STANDARD BOUNDING BOX PROCESSING =====
        DebugLog "Using standard bounding box processing..."

        ' OPTIMIZATION: Only rebuild if we haven't already rebuilt for sheet metal
        ' The ForceRebuild above for hasFlatPattern already rebuilt the model
        If Not hasFlatPattern Then
            swPart.ForceRebuild3 True
            DebugLog "Performed initial rebuild (no flat pattern path)"
        Else
            DebugLog "Skipping rebuild - already performed for flat pattern processing"
        End If

        ' Get solid bodies for fallback
        Dim vSolidBodies As Variant
        vSolidBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)
        
        ' ===== IMPROVED DYNAMIC FEATURE SUPPRESSION WITH STATE PRESERVATION =====
        DebugLog "=== ATTEMPTING DYNAMIC FEATURE SUPPRESSION METHOD ==="
        
        ' Declare the variable
        Dim gotCleanBoundingBox As Boolean
        gotCleanBoundingBox = False
        
        ' *** IMPROVED: Capture original suppression states BEFORE making changes ***
        Dim featureStateMap As Object
        Set featureStateMap = CreateObject("Scripting.Dictionary")
        
        ' First pass: Capture ALL feature states
        DebugLog "=== CAPTURING ORIGINAL FEATURE STATES ==="
        Set swFeature = swPart.FirstFeature
        Do While Not swFeature Is Nothing
            featureStateMap.Add swFeature.Name, swFeature.IsSuppressed()
            DebugLog "Captured state: " & swFeature.Name & " = " & IIf(swFeature.IsSuppressed(), "Suppressed", "Unsuppressed")
            Set swFeature = swFeature.GetNextFeature
        Loop
        
        ' Dynamically find features to suppress by type (not hardcoded names)
        Dim featuresToSuppress As collection
        Set featuresToSuppress = New collection
        
        ' Second pass: Identify features to suppress
        Set swFeature = swPart.FirstFeature
        Do While Not swFeature Is Nothing
            Dim featureType As String
            featureType = swFeature.GetTypeName2
            
            ' Check if this is a feature type that typically adds/removes material from base geometry
            Select Case featureType
                Case "Chamfer", "Fillet", "ICE", "RevCut", "CutSweep", "Mirror", "CirPattern", "LPattern", "SweepThread"
                    ' Only consider unsuppressed features for suppression
                    If Not swFeature.IsSuppressed() Then
                        featuresToSuppress.Add swFeature
                        DebugLog "Will suppress feature: " & swFeature.Name & " (Type: " & featureType & ")"
                    End If
            End Select
            
            Set swFeature = swFeature.GetNextFeature
        Loop
        
        ' Suppress the identified features if any were found
        If featuresToSuppress.Count > 0 Then
            DebugLog "Suppressing " & featuresToSuppress.Count & " features for clean bounding box"

            Dim featureToSuppress As Feature
            For Each featureToSuppress In featuresToSuppress
                featureToSuppress.SetSuppression2 swSuppressFeature, swThisConfiguration, Nothing
                DebugLog "Suppressed: " & featureToSuppress.Name
            Next featureToSuppress

            ' OPTIMIZATION: Single rebuild after all suppressions (not per-feature)
            DebugLog "Rebuilding after all suppressions"
            swPart.ForceRebuild3 True

            ' Get clean bounding box
            vSolidBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)

            If Not IsEmpty(vSolidBodies) Then
                Set swBody = vSolidBodies(0)
                vBox = swBody.GetBodyBox()
                gotCleanBoundingBox = True
                DebugLog "Successfully got clean bounding box with suppressed features"
            End If

            ' *** RESTORE ALL FEATURES TO THEIR ORIGINAL STATES ***
            ' OPTIMIZATION: Batch all restorations, then single rebuild at end
            DebugLog "=== RESTORING ALL FEATURES TO ORIGINAL STATES ==="
            Dim needsRebuildAfterRestore As Boolean
            needsRebuildAfterRestore = False

            Set swFeature = swPart.FirstFeature
            Do While Not swFeature Is Nothing
                Dim originalState As Boolean
                originalState = featureStateMap(swFeature.Name)
                Dim currentState As Boolean
                currentState = swFeature.IsSuppressed()

                ' Only change suppression if it differs from original state
                If originalState <> currentState Then
                    If originalState Then
                        ' Was originally suppressed, restore to suppressed
                        swFeature.SetSuppression2 swSuppressFeature, swThisConfiguration, Nothing
                        DebugLog "Restored to SUPPRESSED: " & swFeature.Name
                    Else
                        ' Was originally unsuppressed, restore to unsuppressed
                        swFeature.SetSuppression2 swUnSuppressFeature, swThisConfiguration, Nothing
                        DebugLog "Restored to UNSUPPRESSED: " & swFeature.Name
                    End If
                    needsRebuildAfterRestore = True
                Else
                    DebugLog "No change needed for: " & swFeature.Name & " (already in correct state)"
                End If

                Set swFeature = swFeature.GetNextFeature
            Loop

            ' OPTIMIZATION: Only rebuild if we actually changed something
            If needsRebuildAfterRestore Then
                DebugLog "Final rebuild after restoring features"
                swPart.ForceRebuild3 True
            Else
                DebugLog "No rebuild needed - no features were changed"
            End If
        Else
            DebugLog "No features identified for suppression"
        End If
        
        ' Fall back to solid body method if suppression method didn't work
        If Not gotCleanBoundingBox Then
            DebugLog "Falling back to standard solid body bounding box method"
            If Not IsEmpty(vSolidBodies) Then
                DebugLog "Using solid body bounding box"
                Set swBody = vSolidBodies(0)
                vBox = swBody.GetBodyBox()
            Else
                DebugLog "Using part bounding box"
                vBox = swPart.GetPartBox(False)
            End If
        End If
        
        ' Calculate dimensions from bounding box
        If Not IsEmpty(vBox) Then
            length = Abs(vBox(3) - vBox(0))
            width = Abs(vBox(5) - vBox(2))
            height = Abs(vBox(4) - vBox(1))
            
            DebugLog "=== RAW BOUNDING BOX DIMENSIONS (in meters) ==="
            DebugLog "Length (X): " & length
            DebugLog "Width (Z): " & width
            DebugLog "Height (Y): " & height
            
            ' SIMPLIFIED: Convert from meters to model's native units
            DebugLog "Converting from meters to model units (" & dimensionalUOM & ")..."
            
            Select Case LCase(dimensionalUOM)
                Case "mm"
                    length = length * 1000
                    width = width * 1000
                    height = height * 1000
                    DebugLog "Converted to mm: Length=" & length & ", Width=" & width & ", Height=" & height
                    
                Case "cm"
                    length = length * 100
                    width = width * 100
                    height = height * 100
                    DebugLog "Converted to cm: Length=" & length & ", Width=" & width & ", Height=" & height
                    
                Case "m"
                    ' Already in meters, no conversion needed
                    DebugLog "Already in meters: Length=" & length & ", Width=" & width & ", Height=" & height
                    
                Case "in"
                    length = length * 39.3701
                    width = width * 39.3701
                    height = height * 39.3701
                    DebugLog "Converted to inches: Length=" & length & ", Width=" & width & ", Height=" & height
                    
                Case "ft"
                    length = length * 3.28084
                    width = width * 3.28084
                    height = height * 3.28084
                    DebugLog "Converted to feet: Length=" & length & ", Width=" & width & ", Height=" & height
                    
                Case Else
                    DebugLog "Unknown unit '" & dimensionalUOM & "' - defaulting to inches"
                    length = length * 39.3701
                    width = width * 39.3701
                    height = height * 39.3701
                    dimensionalUOM = "in"
            End Select
            
            ' Set the dimensional UOM to match the model
            propertiesToSet("Dimensional UOM") = dimensionalUOM
            
            ' Sort and assign dimensions
            sortedDimensions(0) = length
            sortedDimensions(1) = width
            sortedDimensions(2) = height
            Call BubbleSort(sortedDimensions)
            
            swDimensions(1) = sortedDimensions(0)  ' Smallest (Width)
            swDimensions(2) = sortedDimensions(1)  ' Middle (Height)
            swDimensions(0) = sortedDimensions(2)  ' Largest (Length)
            
            DebugLog "=== FINAL DIMENSIONS ==="
            DebugLog "swDimensions(0) [Length]: " & swDimensions(0) & " " & dimensionalUOM
            DebugLog "swDimensions(1) [Width]: " & swDimensions(1) & " " & dimensionalUOM
            DebugLog "swDimensions(2) [Height]: " & swDimensions(2) & " " & dimensionalUOM
            DebugLog "Dimensional UOM: " & dimensionalUOM
        End If
    End If

    GetDimensions = swDimensions
End Function




Function CheckIfStructuralMember(model As ModelDoc2) As Boolean
    Dim swFeature As SldWorks.Feature
    Dim swCustPropMgr As SldWorks.customPropertyManager
    Dim valueOut As String
    Dim resolvedValueOut As String
    
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            ' Check for structural member type
            swCustPropMgr.Get4 "Type", False, valueOut, resolvedValueOut
            If resolvedValueOut <> "" Then
                CheckIfStructuralMember = True
                Exit Function
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    CheckIfStructuralMember = False
End Function

' Helper function to sort dimensions
Sub BubbleSort(arr() As Double)
    Dim i As Long, j As Long
    Dim temp As Double
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then ' Changed from < to > for ascending order
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub



' Function to toggle the use of decimals
Sub ToggleDecimals(useDecimals As Boolean)
    USE_DECIMALS = useDecimals
End Sub

Function FractionToDecimal(fractionStr As String) As Double
    Dim parts() As String
    Dim wholePart As Double
    Dim fractionPart As Double
    Dim fracParts() As String
    Dim cleanedFraction As String

    ' Remove unwanted characters: quotes (") and extra spaces
    cleanedFraction = Trim(Replace(fractionStr, """", "")) ' Removes inch symbol

    ' Split whole number and fraction (e.g., "23 5/8" ? ["23", "5/8"])
    parts = Split(cleanedFraction, " ")

    ' Initialize parts
    wholePart = 0
    fractionPart = 0

    If UBound(parts) = 0 Then
        ' Only fraction part or whole number
        If InStr(parts(0), "/") > 0 Then
            ' It's a fraction (e.g., "5/8")
            fracParts = Split(parts(0), "/")
            If UBound(fracParts) = 1 Then
                fractionPart = CDbl(Trim(fracParts(0))) / CDbl(Trim(fracParts(1)))
            End If
        Else
            ' It's a whole number (e.g., "23")
            wholePart = CDbl(Trim(parts(0)))
        End If
    ElseIf UBound(parts) = 1 Then
        ' Whole number + fraction case (e.g., "23 5/8")
        wholePart = CDbl(Trim(parts(0)))
        fracParts = Split(Trim(parts(1)), "/")
        If UBound(fracParts) = 1 Then
            fractionPart = CDbl(Trim(fracParts(0))) / CDbl(Trim(fracParts(1)))
        End If
    End If

    ' Return final decimal value
    FractionToDecimal = wholePart + fractionPart
End Function









Sub LogDocumentPrecision(model As ModelDoc2)
    ' Ensure model is valid
    If model Is Nothing Then
        DebugLog "Error: No active model found."
        Exit Sub
    End If

    ' Get document units
    Dim lengthUnit As UserUnit
    Set lengthUnit = model.GetUserUnit(swUserUnitsType_e.swLengthUnit)
    
    ' Get document precision
    DebugLog "----------------------------------------"
 '   DebugLog "Document Precision: " & lengthUnit.Decimals
    DebugLog "----------------------------------------"
End Sub

Function GetUnitSystemName(unitSystem As Long) As String
    Select Case unitSystem
        Case swUnitSystem_e.swUnitSystem_IPS
            GetUnitSystemName = "IPS (Inch, Pound, Second)"
        Case swUnitSystem_e.swUnitSystem_MMGS
            GetUnitSystemName = "MMGS (Millimeter, Gram, Second)"
        Case swUnitSystem_e.swUnitSystem_CGS
            GetUnitSystemName = "CGS (Centimeter, Gram, Second)"
        Case swUnitSystem_e.swUnitSystem_MKS
            GetUnitSystemName = "MKS (Meter, Kilogram, Second)"
        Case swUnitSystem_e.swUnitSystem_Custom
            GetUnitSystemName = "Custom"
        Case Else
            GetUnitSystemName = "Unknown (" & unitSystem & ")"
    End Select
End Function

Public Function IsASNZSStructuralMember(ByVal model As ModelDoc2) As Boolean
    IsASNZSStructuralMember = False
    
    If model Is Nothing Then Exit Function
    
    DebugLog "=== Checking for AS NZS structural member ==="
    
    ' Look for CutListFolder with AS NZS in the profile path or weldment member name
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                ' Check various properties that might contain AS NZS reference
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check Profile property
                bRet = swCustPropMgr.Get4("Profile", False, valueOut, resolvedValueOut)
                If bRet And InStr(UCase(resolvedValueOut), "AS NZS") > 0 Then
                    DebugLog "Found AS NZS in Profile: " & resolvedValueOut
                    IsASNZSStructuralMember = True
                    Exit Function
                End If
                
                ' Check Profile Description
                bRet = swCustPropMgr.Get4("Profile Description", False, valueOut, resolvedValueOut)
                If bRet And InStr(UCase(resolvedValueOut), "AS NZS") > 0 Then
                    DebugLog "Found AS NZS in Profile Description: " & resolvedValueOut
                    IsASNZSStructuralMember = True
                    Exit Function
                End If
                
                ' Check StockSize for AS NZS reference
                bRet = swCustPropMgr.Get4("StockSize", False, valueOut, resolvedValueOut)
                If bRet And InStr(UCase(resolvedValueOut), "AS NZS") > 0 Then
                    DebugLog "Found AS NZS in StockSize: " & resolvedValueOut
                    IsASNZSStructuralMember = True
                    Exit Function
                End If
                
                ' Check Type property - sometimes AS NZS is indicated there
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And InStr(UCase(resolvedValueOut), "AS NZS") > 0 Then
                    DebugLog "Found AS NZS in Type: " & resolvedValueOut
                    IsASNZSStructuralMember = True
                    Exit Function
                End If
            End If
        End If
        
        ' Also check WeldMemberFeat features
        If swFeature.GetTypeName2 = "WeldMemberFeat" Then
            Dim featureName As String
            featureName = swFeature.Name
            If InStr(UCase(featureName), "AS NZS") > 0 Then
                DebugLog "Found AS NZS in WeldMemberFeat name: " & featureName
                IsASNZSStructuralMember = True
                Exit Function
            End If
        End If
        
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    DebugLog "No AS NZS reference found in structural member"
End Function

