Attribute VB_Name = "Plate"
#Const TESTING = 1

Option Explicit
Option Private Module

' Constants for maximum dimensions
Private Const MAX_SHORT_SIDE_FT As Double = 6   ' 6 feet maximum for shorter side (Height)
Private Const MAX_LONG_SIDE_FT As Double = 20   ' 20 feet maximum for longer side (Length)

' Material-specific thickness limits (in inches)
Private Const MAX_STEEL_THICKNESS_IN As Double = 1.5      ' Steel thickness limit
Private Const MAX_ALUMINUM_THICKNESS_IN As Double = 1.5  ' Aluminum thickness limit

' Dictionary for processed models
Private processedModelNumbers As Object
Private lastRunTime As Date

' Material type enumerations
Private Enum MaterialType
    Steel = 1
    Aluminum = 2
    UNKNOWN = 0
End Enum

' Material lists
Private Type MaterialLists
    CarbonSteels() As String
    stainlessSteels() As String
    aluminums() As String
End Type

Private materials As MaterialLists

' Ensures that the dictionary is initialized before use and checks for new run
Private Sub EnsureDictionaryInitialized()
    Static firstRun As Boolean
    
    ' Initialize dictionary if needed
    If processedModelNumbers Is Nothing Then
        Set processedModelNumbers = CreateObject("Scripting.Dictionary")
        firstRun = True
    End If
    
    ' Check if this is a new run (more than 1 second since last run)
    If DateDiff("s", lastRunTime, Now) > 1 Or firstRun = True Then
        firstRun = False
        Set processedModelNumbers = CreateObject("Scripting.Dictionary")
        lastRunTime = Now
        DebugLog "New macro run detected - Reset tracking at " & Now
    End If
End Sub

' Initialize material lists
Private Sub InitializeMaterialLists()
   
    Dim i As Long
   
    
    ' Carbon steels
    Dim tempCarbonSteels As Variant
    tempCarbonSteels = Array("44W", "50W", "300W", "350W", "400W", "QT100", "44W-PERF", "STRENX100XF", "A514-Q")
    ReDim materials.CarbonSteels(LBound(tempCarbonSteels) To UBound(tempCarbonSteels))
    For i = LBound(tempCarbonSteels) To UBound(tempCarbonSteels)
        materials.CarbonSteels(i) = tempCarbonSteels(i)
    Next i
    
    ' Stainless Steels
    Dim tempStainlessSteels As Variant
    tempStainlessSteels = Array("316", "304", "308", "309")
    ReDim materials.stainlessSteels(LBound(tempStainlessSteels) To UBound(tempStainlessSteels))
    For i = LBound(tempStainlessSteels) To UBound(tempStainlessSteels)
        materials.stainlessSteels(i) = tempStainlessSteels(i)
    Next i
    
    ' Aluminums
    Dim tempAluminums As Variant
    tempAluminums = Array("3003-H12", "3003-H14", "5052-H19", "5052-H32", _
                     "5083-H32", "5083-H34", "5083-H116", "5083-H321", _
                     "6061-T6", "6063-O")
    ReDim materials.aluminums(LBound(tempAluminums) To UBound(tempAluminums))
    For i = LBound(tempAluminums) To UBound(tempAluminums)
        materials.aluminums(i) = tempAluminums(i)
    Next i
End Sub

' Function to determine material type
Private Function GetMaterialType(ByVal materialName As String) As MaterialType
    Dim i As Long
    
    ' Check Carbon Steels
    For i = LBound(materials.CarbonSteels) To UBound(materials.CarbonSteels)
        If materialName = materials.CarbonSteels(i) Then
            GetMaterialType = MaterialType.Steel
            Exit Function
        End If
    Next i
    

    
    ' Check Stainless Steels
    For i = LBound(materials.stainlessSteels) To UBound(materials.stainlessSteels)
        If materialName = materials.stainlessSteels(i) Then
            GetMaterialType = MaterialType.Steel
            Exit Function
        End If
    Next i
    
    ' Check Aluminums
    For i = LBound(materials.aluminums) To UBound(materials.aluminums)
        If materialName = materials.aluminums(i) Then
            GetMaterialType = MaterialType.Aluminum
            Exit Function
        End If
    Next i
    
    GetMaterialType = MaterialType.UNKNOWN
End Function

' Get maximum thickness based on material type
Private Function GetMaxThickness(ByVal matType As MaterialType) As Double
    Select Case matType
        Case MaterialType.Steel
            GetMaxThickness = MAX_STEEL_THICKNESS_IN
        Case MaterialType.Aluminum
            GetMaxThickness = MAX_ALUMINUM_THICKNESS_IN
        Case Else
            GetMaxThickness = MAX_STEEL_THICKNESS_IN  ' Default to steel thickness
    End Select
End Function

' Function to convert units to feet
Private Function ConvertToFeet(ByVal value As Double, ByVal unitSystem As String) As Double
    Select Case LCase(unitSystem)
        Case "in": ConvertToFeet = value / 12        ' Inches to feet
        Case "mm": ConvertToFeet = value / 304.8     ' Millimeters to feet
        Case "cm": ConvertToFeet = value / 30.48     ' Centimeters to feet
        Case "m": ConvertToFeet = value * 3.28084    ' Meters to feet
        Case "ft": ConvertToFeet = value             ' Already in feet
        Case Else: ConvertToFeet = value             ' Default (assumes feet)
    End Select
End Function

Public Function CheckPlateDimensions(ByVal model As ModelDoc2, ByVal width As Double, ByVal height As Double, ByVal length As Double, ByVal unitSystem As String, ByRef propertiesToSet As Object) As Boolean
    ' CHECKPOINT 1: Entry parameters
    DebugLog "=== CHECKPOINT 1: ENTRY ==="
    DebugLog "Input: width=" & width & ", height=" & height & ", length=" & length & ", unit=" & unitSystem
    '  HERE TO CHECK INITIAL VALUES
    
    Dim modelNum As String
    Dim widthInFeet As Double, heightInFeet As Double, lengthInFeet As Double
    Dim exceedsSizeLimits As Boolean, exceedsThicknessLimit As Boolean
    Dim dimensionStr As String
    Dim currentType As String
    Dim fileName As String
    
    fileName = model.GetTitle()
    
    ' REMOVED: Problematic unit assumption workaround
    ' The function now trusts the input unit system and converts properly
    ' If there are unit issues, they should be fixed at the source where
    ' dimensions are extracted, not by guessing based on magnitude

    InitializeMaterialLists
    EnsureDictionaryInitialized

    ' CHECKPOINT 2: Material extraction
    DebugLog "=== CHECKPOINT 2: MATERIAL EXTRACTION ==="
    Dim actualMaterial As String
    actualMaterial = ""
    
    If propertiesToSet.exists("Material") Then
        Dim rawMaterial As String
        rawMaterial = propertiesToSet("Material")
        DebugLog "Raw Material: " & rawMaterial
        
        ' Clean up material name
        rawMaterial = Replace(rawMaterial, """", "")
        If InStr(rawMaterial, "SW-Material@") > 0 Then
            rawMaterial = Mid(rawMaterial, InStr(rawMaterial, "@") + 1)
        End If
        If InStr(rawMaterial, ".SLDPRT") > 0 Then
            rawMaterial = Left(rawMaterial, InStr(rawMaterial, ".SLDPRT") - 1)
        End If
        actualMaterial = rawMaterial
    End If
    
    DebugLog "Cleaned Material: " & actualMaterial
    '  HERE TO CHECK MATERIAL EXTRACTION
    
    ' CHECKPOINT 3: Material type determination
    DebugLog "=== CHECKPOINT 3: MATERIAL TYPE ==="
    Dim matType As MaterialType
    matType = GetMaterialType(actualMaterial)
    DebugLog "Material Type: " & matType & " (0=Unknown, 1=Steel, 2=Aluminum)"
    
    If matType = MaterialType.UNKNOWN Then
        DebugLog "CHECKPOINT 3A: Unknown material - will set to P"
        '  HERE FOR UNKNOWN MATERIAL ISSUE
        propertiesToSet("Type") = "P"
        CheckPlateDimensions = False
        Exit Function
    End If
    
    ' CHECKPOINT 4: Unit conversion
    DebugLog "=== CHECKPOINT 4: UNIT CONVERSION ==="
    widthInFeet = ConvertToFeet(width, unitSystem)
    heightInFeet = ConvertToFeet(height, unitSystem)
    lengthInFeet = ConvertToFeet(length, unitSystem)
    
    DebugLog "Before conversion: " & width & "�" & height & "�" & length & " " & unitSystem
    DebugLog "After conversion: " & widthInFeet & "�" & heightInFeet & "�" & lengthInFeet & " ft"
    DebugLog "Thickness in inches: " & (widthInFeet * 12)
    '  HERE TO VERIFY UNIT CONVERSION
    
    ' CHECKPOINT 5: Limit checking
    DebugLog "=== CHECKPOINT 5: LIMIT CHECKING ==="
    Dim maxThickness As Double
    Dim minThickness As Double
    maxThickness = GetMaxThickness(matType)
    minThickness = 0.0001 ' Minimum thickness in inches

    exceedsSizeLimits = (heightInFeet > MAX_SHORT_SIDE_FT) Or (lengthInFeet > MAX_LONG_SIDE_FT)
    exceedsThicknessLimit = (widthInFeet * 12) > maxThickness
    Dim belowMinThickness As Boolean
    belowMinThickness = (widthInFeet * 12) < minThickness

    DebugLog "Max thickness allowed: " & maxThickness & " inches"
    DebugLog "Min thickness required: " & minThickness & " inches"
    DebugLog "Actual thickness: " & (widthInFeet * 12) & " inches"
    DebugLog "Height limit check: " & heightInFeet & " > " & MAX_SHORT_SIDE_FT & " = " & (heightInFeet > MAX_SHORT_SIDE_FT)
    DebugLog "Length limit check: " & lengthInFeet & " > " & MAX_LONG_SIDE_FT & " = " & (lengthInFeet > MAX_LONG_SIDE_FT)
    DebugLog "Size exceeds limits: " & exceedsSizeLimits
    DebugLog "Thickness exceeds limit: " & exceedsThicknessLimit
    DebugLog "Below minimum thickness: " & belowMinThickness
    '  HERE TO CHECK LIMIT CALCULATIONS

    ' Get current type
    If propertiesToSet.exists("Type") Then
        currentType = propertiesToSet("Type")
    Else
        currentType = ""
    End If

    ' CHECKPOINT 6: Decision logic
    DebugLog "=== CHECKPOINT 6: DECISION LOGIC ==="
    DebugLog "Current Type: " & currentType
    DebugLog "Exceeds ANY limits: " & (exceedsSizeLimits Or exceedsThicknessLimit Or belowMinThickness)

    If exceedsSizeLimits Or exceedsThicknessLimit Or belowMinThickness Then
        If belowMinThickness Then
            DebugLog "BRANCH: Setting to P (below minimum thickness)"
        Else
            DebugLog "BRANCH: Setting to P (exceeds limits)"
        End If
        '  HERE FOR OVERSIZE OR UNDERSIZE PARTS
        propertiesToSet("Type") = "P"
        CheckPlateDimensions = False
        
    ElseIf currentType = "P" Then
        DebugLog "BRANCH: Currently P, within limits - changing to M"
        '  HERE FOR P->M CONVERSION
        propertiesToSet("Type") = "M"
        CheckPlateDimensions = True
        
    Else
        DebugLog "BRANCH: Within limits, setting to M"
        '  HERE FOR SUCCESSFUL M CLASSIFICATION
        propertiesToSet("Type") = "M"
        CheckPlateDimensions = True
    End If
        
    ' CHECKPOINT 7: Final result
    DebugLog "=== CHECKPOINT 7: FINAL RESULT ==="
    DebugLog "Final Type: " & propertiesToSet("Type")
    DebugLog "Function returns: " & CheckPlateDimensions
    '  HERE TO SEE FINAL RESULT

End Function

' Additional helper function to check material lists
Public Sub DebugMaterialLists()
    DebugLog "=== MATERIAL LISTS DEBUG ==="
    
    InitializeMaterialLists
    
    DebugLog "Carbon Steels:"
    Dim i As Long
    For i = LBound(materials.CarbonSteels) To UBound(materials.CarbonSteels)
        DebugLog "  " & materials.CarbonSteels(i)
    Next i
    
    DebugLog "Stainless Steels:"
    For i = LBound(materials.stainlessSteels) To UBound(materials.stainlessSteels)
        DebugLog "  " & materials.stainlessSteels(i)
    Next i
    
    DebugLog "Aluminums:"
    For i = LBound(materials.aluminums) To UBound(materials.aluminums)
        DebugLog "  " & materials.aluminums(i)
    Next i
End Sub



' Function to reset the processed models list
Public Sub ResetProcessedModels()
    Set processedModelNumbers = CreateObject("Scripting.Dictionary")
    lastRunTime = Now
End Sub



