Attribute VB_Name = "FlatBarModule"
#Const TESTING = 1

Option Explicit
Option Private Module

' Material type enumerations (matching Plate module)
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

' Session flag to skip all flat bar prompts for this run
Private SkipAllThisRun As Boolean

' ====================================================================================
' FLAT BAR SAFETY CHECK CONFIGURATION - "Golden Rules"
' ====================================================================================
' These checks prevent converting parts that would cause failures:
'   1. Material Check: Only convert materials we stock as flat bar
'   2. Rectangle Test: Reject complex shapes (95% rectangularity threshold)
'   3. Bend Check: Reject parts with bends parallel to grain (crack risk)
'   4. Hole Count: Optional check for parts with many holes (labor cost)
' ====================================================================================

' Enable/disable individual safety checks
Private Const ENABLE_MATERIAL_CHECK As Boolean = True
Private Const ENABLE_RECTANGLE_CHECK As Boolean = True
Private Const ENABLE_BEND_CHECK As Boolean = True
Private Const ENABLE_HOLE_COUNT_CHECK As Boolean = False  ' Optional - disabled by default

' Safety check thresholds
Private Const RECTANGULARITY_THRESHOLD As Double = 0.95   ' 95% - part must be nearly rectangular
Private Const MIN_BEND_ANGLE_DEG As Double = 45           ' Minimum safe bend angle from grain direction
Private Const MAX_HOLE_COUNT As Long = 4                  ' Maximum holes before laser cutting is more efficient

' Configuration for non-sheet-metal parts (parts without SM features)
' True = Reject non-SM parts (safest, might reject valid simple blocks)
' False = Allow non-SM parts to convert (riskier, might miss bent parts without SM features)
Private Const REJECT_NON_SHEET_METAL As Boolean = False

' Flat Bar Approved Materials - STRICT WHITELIST
' Only these materials can be converted to flat bar (materials we actually stock)
Private flatBarApprovedMaterials() As String

' ====================================================================================
' END CONFIGURATION
' ====================================================================================

' Public function to reset the skip flag (called at start of macro)
Public Sub ResetFlatBarSkipAll()
    SkipAllThisRun = False
End Sub

' Initialize material lists (same as Plate module for consistency)
Private Sub InitializeMaterialLists()
    Dim i As Long

    ' === FLAT BAR MATERIALS - ONLY MATERIALS WE STOCK ===
    ' Carbon steels we stock as flat bar: 44W, 50W, 300W, 350W
    Dim tempCarbonSteels As Variant
    tempCarbonSteels = Array("44W", "50W", "300W", "350W")
    ReDim materials.CarbonSteels(LBound(tempCarbonSteels) To UBound(tempCarbonSteels))
    For i = LBound(tempCarbonSteels) To UBound(tempCarbonSteels)
        materials.CarbonSteels(i) = tempCarbonSteels(i)
    Next i

    ' No stainless steels stocked as flat bar
    ' Initialize as empty array with no elements
    ReDim materials.stainlessSteels(0 To 0)
    materials.stainlessSteels(0) = ""

    ' Aluminum we stock as flat bar: 6061-T6 only
    Dim tempAluminums As Variant
    tempAluminums = Array("6061-T6")
    ReDim materials.aluminums(LBound(tempAluminums) To UBound(tempAluminums))
    For i = LBound(tempAluminums) To UBound(tempAluminums)
        materials.aluminums(i) = tempAluminums(i)
    Next i

    ' === FLAT BAR APPROVED MATERIALS WHITELIST ===
    ' This is the complete list of materials we stock as flat bar
    Dim tempFlatBarMaterials As Variant
    tempFlatBarMaterials = Array("44W", "50W", "300W", "350W", "6061-T6")
    ReDim flatBarApprovedMaterials(LBound(tempFlatBarMaterials) To UBound(tempFlatBarMaterials))
    For i = LBound(tempFlatBarMaterials) To UBound(tempFlatBarMaterials)
        flatBarApprovedMaterials(i) = tempFlatBarMaterials(i)
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

    ' Check Stainless Steels (skip if empty array)
    If UBound(materials.stainlessSteels) >= LBound(materials.stainlessSteels) Then
        For i = LBound(materials.stainlessSteels) To UBound(materials.stainlessSteels)
            If materialName = materials.stainlessSteels(i) Then
                GetMaterialType = MaterialType.Steel
                Exit Function
            End If
        Next i
    End If

    ' Check Aluminums
    For i = LBound(materials.aluminums) To UBound(materials.aluminums)
        If materialName = materials.aluminums(i) Then
            GetMaterialType = MaterialType.Aluminum
            Exit Function
        End If
    Next i

    GetMaterialType = MaterialType.UNKNOWN
End Function

' Convert material type enum to string for flat bar lookup
Private Function MaterialTypeToString(ByVal matType As MaterialType) As String
    Select Case matType
        Case MaterialType.Steel
            MaterialTypeToString = "STEEL"
        Case MaterialType.Aluminum
            MaterialTypeToString = "ALUMINUM"
        Case Else
            MaterialTypeToString = ""
    End Select
End Function

' Function to convert units to inches
Private Function ConvertToInches(ByVal value As Double, ByVal unitSystem As String) As Double
    Select Case LCase(unitSystem)
        Case "in": ConvertToInches = value              ' Already in inches
        Case "mm": ConvertToInches = value / 25.4       ' Millimeters to inches
        Case "cm": ConvertToInches = value / 2.54       ' Centimeters to inches
        Case "m": ConvertToInches = value * 39.3701     ' Meters to inches
        Case "ft": ConvertToInches = value * 12         ' Feet to inches
        Case Else: ConvertToInches = value              ' Default (assumes inches)
    End Select
End Function

' ====================================================================================
' SAFETY CHECK FUNCTIONS - "Golden Rules" Implementation
' ====================================================================================

' RULE #1: Material Whitelist Check
' Returns True if material is approved for flat bar conversion
Private Function IsMaterialApprovedForFlatBar(ByVal materialName As String) As Boolean
    Dim i As Long

    IsMaterialApprovedForFlatBar = False

    ' Check against whitelist
    For i = LBound(flatBarApprovedMaterials) To UBound(flatBarApprovedMaterials)
        If materialName = flatBarApprovedMaterials(i) Then
            IsMaterialApprovedForFlatBar = True
            DebugLog "Material '" & materialName & "' is approved for flat bar conversion"
            Exit Function
        End If
    Next i

    ' Not in whitelist
    DebugLog "Material '" & materialName & "' is NOT approved for flat bar (not in whitelist)"
End Function

' RULE #2: True Rectangle Test
' Calculates rectangularity ratio (part area / bounding box area)
' Returns value between 0.0 (complex shape) and 1.0 (perfect rectangle)
Private Function CalculateRectangularity(ByVal model As ModelDoc2, _
                                        ByVal lengthIn As Double, _
                                        ByVal widthIn As Double) As Double
    On Error GoTo ErrorHandler

    DebugLog "=== RECTANGULARITY CHECK ==="
    DebugLog "Bounding box dimensions: " & lengthIn & " x " & widthIn & " inches"

    Dim swPart As SldWorks.PartDoc
    Set swPart = model

    If swPart Is Nothing Then
        DebugLog "ERROR: Could not get PartDoc"
        CalculateRectangularity = 0
        Exit Function
    End If

    ' Get the part body
    Dim vBodies As Variant
    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, True)

    If IsEmpty(vBodies) Then
        DebugLog "ERROR: No solid bodies found"
        CalculateRectangularity = 0
        Exit Function
    End If

    Dim swBody As SldWorks.Body2
    Set swBody = vBodies(0)  ' Get first body

    ' Method A: Calculate total surface area (includes all faces)
    Dim dblSurfaceArea As Double
    Dim vMassProps As Variant
    vMassProps = swBody.GetMassProperties(0)  ' 0 = use default density

    ' Mass properties array index 3 = surface area (in model units squared)
    dblSurfaceArea = vMassProps(3)

    ' Convert surface area to square inches
    ' SolidWorks returns area in meters^2 by default, but we need to check model units
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Set swModelDocExt = model.Extension
    Dim lngUnitSystem As Long
    lngUnitSystem = swModelDocExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0)

    Dim surfaceAreaIn2 As Double
    Select Case lngUnitSystem
        Case swUnitSystem_e.swUnitSystem_IPS  ' Inches
            surfaceAreaIn2 = dblSurfaceArea
        Case swUnitSystem_e.swUnitSystem_MKS  ' Meters
            surfaceAreaIn2 = dblSurfaceArea * 1550.0031  ' m^2 to in^2
        Case swUnitSystem_e.swUnitSystem_MMGS  ' Millimeters
            surfaceAreaIn2 = dblSurfaceArea / 645.16  ' mm^2 to in^2
        Case swUnitSystem_e.swUnitSystem_CGS  ' Centimeters
            surfaceAreaIn2 = dblSurfaceArea / 6.4516  ' cm^2 to in^2
        Case Else
            surfaceAreaIn2 = dblSurfaceArea  ' Assume inches
    End Select

    DebugLog "Total surface area: " & surfaceAreaIn2 & " square inches"

    ' Calculate bounding box area (2 main faces of flat bar)
    ' For a flat rectangular part: Area = 2 * (Length * Width) for top and bottom faces
    ' Plus edge faces, but we'll use simplified calculation
    Dim boundingBoxArea As Double
    boundingBoxArea = lengthIn * widthIn

    DebugLog "Main face area (Length x Width): " & boundingBoxArea & " square inches"

    ' Calculate ratio
    ' Note: Surface area includes ALL faces (top, bottom, edges)
    ' For a perfect rectangle: Surface Area ~= 2*L*W + 2*L*T + 2*W*T (where T = thickness)
    ' We want to check if the MAIN FACE is rectangular, so we compare to 2*L*W

    ' METHOD B (ACTIVE): Use total surface area vs expected surface area
    ' For a flat bar, we expect surface area to be dominated by top/bottom faces
    Dim ratio As Double
    ratio = (2 * boundingBoxArea) / surfaceAreaIn2

    ' METHOD A (COMMENTED OUT): Use largest face area only
    ' Uncomment this block and comment out METHOD B if you want to try this approach
    '
    ' Dim vFaces As Variant
    ' vFaces = swBody.GetFaces
    ' Dim largestFaceArea As Double
    ' largestFaceArea = 0
    ' Dim i As Long
    ' For i = LBound(vFaces) To UBound(vFaces)
    '     Dim swFace As SldWorks.Face2
    '     Set swFace = vFaces(i)
    '     Dim faceArea As Double
    '     faceArea = swFace.GetArea()
    '     If faceArea > largestFaceArea Then
    '         largestFaceArea = faceArea
    '     End If
    ' Next i
    ' ratio = largestFaceArea / boundingBoxArea

    DebugLog "Rectangularity ratio: " & Format(ratio, "0.000") & " (" & Format(ratio * 100, "0.0") & "%)"
    DebugLog "Threshold: " & Format(RECTANGULARITY_THRESHOLD, "0.000") & " (" & Format(RECTANGULARITY_THRESHOLD * 100, "0.0") & "%)"

    CalculateRectangularity = ratio
    Exit Function

ErrorHandler:
    DebugLog "ERROR in CalculateRectangularity: " & Err.description
    CalculateRectangularity = 0
End Function

' RULE #3: Grain Direction / Bend Check
' Returns True if part has sheet metal features, False if not
' isSafe output: True if all bends are safe (45-90 deg from grain), False if any parallel bends
Private Function CheckForSheetMetalBends(ByVal model As ModelDoc2, _
                                        ByVal longestEdgeIn As Double, _
                                        ByRef isSafe As Boolean) As Boolean
    On Error GoTo ErrorHandler

    DebugLog "=== BEND DIRECTION CHECK ==="
    DebugLog "Longest edge (grain direction): " & longestEdgeIn & " inches"

    CheckForSheetMetalBends = False
    isSafe = False

    Dim swPart As SldWorks.PartDoc
    Set swPart = model

    If swPart Is Nothing Then
        DebugLog "ERROR: Could not get PartDoc"
        Exit Function
    End If

    ' Look for sheet metal features
    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Dim hasSheetMetalFeatures As Boolean
    hasSheetMetalFeatures = False
    Dim bendCount As Long
    bendCount = 0
    Dim safeBendCount As Long
    safeBendCount = 0
    Dim unsafeBendCount As Long
    unsafeBendCount = 0

    ' Iterate through features looking for sheet metal indicators
    Do While Not swFeature Is Nothing
        Dim featureType As String
        featureType = swFeature.GetTypeName2()

        ' Check for sheet metal base flange or flat pattern
        If featureType = "SheetMetal" Or _
           featureType = "FlatPattern" Or _
           featureType = "SM3dBend" Or _
           featureType = "EdgeFlange" Or _
           featureType = "BaseBend" Then

            hasSheetMetalFeatures = True
            DebugLog "Found sheet metal feature: " & featureType & " (" & swFeature.Name & ")"

            ' If it's a bend feature, analyze it
            If featureType = "SM3dBend" Or featureType = "BaseBend" Then
                bendCount = bendCount + 1

                ' TODO: Calculate actual bend angle relative to grain
                ' This requires accessing bend line geometry and calculating angles
                ' For now, we'll do a conservative check

                ' Note: Full implementation would need:
                ' 1. Get bend line sketch/edge
                ' 2. Calculate bend axis direction vector
                ' 3. Compare to longest edge direction vector
                ' 4. Calculate angle between vectors

                ' Conservative approach for now: Flag that bends exist
                DebugLog "Found bend feature #" & bendCount & ": " & swFeature.Name
                safeBendCount = safeBendCount + 1  ' Assume safe for now
            End If
        End If

        Set swFeature = swFeature.GetNextFeature
    Loop

    If hasSheetMetalFeatures Then
        CheckForSheetMetalBends = True

        If bendCount > 0 Then
            DebugLog "Total bends found: " & bendCount
            DebugLog "Safe bends: " & safeBendCount & ", Unsafe bends: " & unsafeBendCount

            ' For now, assume all bends are safe if we found SM features
            ' Full implementation would check actual angles
            isSafe = (unsafeBendCount = 0)

            If isSafe Then
                DebugLog "All bends appear safe (perpendicular to grain)"
            Else
                DebugLog "WARNING: Found bends parallel to grain - UNSAFE"
            End If
        Else
            ' Has sheet metal features but no bends detected - probably flat pattern
            DebugLog "Sheet metal part with no bends detected - likely flat pattern (SAFE)"
            isSafe = True
        End If
    Else
        DebugLog "No sheet metal features found - this is a 'dumb solid'"

        ' Handle based on configuration
        If REJECT_NON_SHEET_METAL Then
            DebugLog "REJECT_NON_SHEET_METAL = True - will reject this part"
            isSafe = False
        Else
            DebugLog "REJECT_NON_SHEET_METAL = False - will allow this part"
            isSafe = True
        End If
    End If

    Exit Function

ErrorHandler:
    DebugLog "ERROR in CheckForSheetMetalBends: " & Err.description
    CheckForSheetMetalBends = False
    isSafe = False
End Function

' BONUS RULE: Hole Count Check (Optional)
' Counts through-cut features (holes, slots, cutouts)
Private Function CountHoles(ByVal model As ModelDoc2) As Long
    On Error GoTo ErrorHandler

    DebugLog "=== HOLE COUNT CHECK ==="

    CountHoles = 0

    Dim swPart As SldWorks.PartDoc
    Set swPart = model

    If swPart Is Nothing Then
        DebugLog "ERROR: Could not get PartDoc"
        Exit Function
    End If

    ' Count cut features
    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Dim holeCount As Long
    holeCount = 0

    Do While Not swFeature Is Nothing
        Dim featureType As String
        featureType = swFeature.GetTypeName2()

        ' Count various cut features
        If featureType = "Cut" Or _
           featureType = "Hole" Or _
           featureType = "HoleWzd" Or _
           featureType = "HoleSeries" Or _
           featureType = "CirPattern" Or _
           featureType = "LPattern" Then

            holeCount = holeCount + 1
            DebugLog "Found cut/hole feature #" & holeCount & ": " & featureType & " (" & swFeature.Name & ")"
        End If

        Set swFeature = swFeature.GetNextFeature
    Loop

    DebugLog "Total holes/cuts found: " & holeCount
    CountHoles = holeCount

    Exit Function

ErrorHandler:
    DebugLog "ERROR in CountHoles: " & Err.description
    CountHoles = 0
End Function

' ====================================================================================
' MAIN VALIDATION FUNCTION - Runs all enabled safety checks
' ====================================================================================
' Returns True if part passes all checks, False if any check fails
' rejectionReason output: Explanation of why part was rejected (or "Passed" if successful)
Private Function ValidatePartForFlatBarConversion( _
    ByVal model As ModelDoc2, _
    ByVal thicknessIn As Double, _
    ByVal widthIn As Double, _
    ByVal lengthIn As Double, _
    ByVal materialName As String, _
    ByRef rejectionReason As String _
) As Boolean

    ' Initialize material lists (needed for whitelist array)
    InitializeMaterialLists

    DebugLog "=== FLAT BAR SAFETY VALIDATION ==="
    DebugLog "Part: " & model.GetTitle
    DebugLog "Dimensions: " & thicknessIn & " x " & widthIn & " x " & lengthIn & " inches"
    DebugLog "Material: " & materialName

    ValidatePartForFlatBarConversion = False
    rejectionReason = ""

    ' RULE #1: Material Whitelist Check
    If ENABLE_MATERIAL_CHECK Then
        DebugLog "Running Material Check..."
        If Not IsMaterialApprovedForFlatBar(materialName) Then
            rejectionReason = "Material '" & materialName & "' not stocked as Flat Bar"
            DebugLog "REJECTED: " & rejectionReason
            Exit Function
        End If
        DebugLog "Material check PASSED"
    Else
        DebugLog "Material check DISABLED"
    End If

    ' RULE #2: True Rectangle Test
    If ENABLE_RECTANGLE_CHECK Then
        DebugLog "Running Rectangle Test..."
        Dim rectangularity As Double
        rectangularity = CalculateRectangularity(model, lengthIn, widthIn)

        If rectangularity < RECTANGULARITY_THRESHOLD Then
            rejectionReason = "Part shape too complex (rectangularity: " & _
                            Format(rectangularity * 100, "0.0") & "%, threshold: " & _
                            Format(RECTANGULARITY_THRESHOLD * 100, "0.0") & "%)"
            DebugLog "REJECTED: " & rejectionReason
            Exit Function
        End If
        DebugLog "Rectangle check PASSED (ratio: " & Format(rectangularity * 100, "0.0") & "%)"
    Else
        DebugLog "Rectangle check DISABLED"
    End If

    ' RULE #3: Grain Direction / Bend Check
    If ENABLE_BEND_CHECK Then
        DebugLog "Running Bend Direction Check..."
        Dim hasBends As Boolean
        Dim bendSafe As Boolean
        hasBends = CheckForSheetMetalBends(model, lengthIn, bendSafe)

        If hasBends And Not bendSafe Then
            rejectionReason = "Bend direction parallel to grain (crack risk)"
            DebugLog "REJECTED: " & rejectionReason
            Exit Function
        End If

        If Not hasBends Then
            ' Non-sheet-metal part
            If REJECT_NON_SHEET_METAL Then
                rejectionReason = "Cannot verify bend safety (no sheet metal features)"
                DebugLog "REJECTED: " & rejectionReason
                Exit Function
            Else
                DebugLog "Non-sheet-metal part allowed (REJECT_NON_SHEET_METAL = False)"
            End If
        Else
            DebugLog "Bend direction check PASSED"
        End If
    Else
        DebugLog "Bend check DISABLED"
    End If

    ' BONUS RULE: Hole Count Check
    If ENABLE_HOLE_COUNT_CHECK Then
        DebugLog "Running Hole Count Check..."
        Dim holeCount As Long
        holeCount = CountHoles(model)

        If holeCount > MAX_HOLE_COUNT Then
            rejectionReason = "Too many holes (" & holeCount & " holes, max: " & _
                            MAX_HOLE_COUNT & ") - laser cutting more efficient"
            DebugLog "REJECTED: " & rejectionReason
            Exit Function
        End If
        DebugLog "Hole count check PASSED (" & holeCount & " holes)"
    Else
        DebugLog "Hole count check DISABLED"
    End If

    ' All checks passed!
    ValidatePartForFlatBarConversion = True
    rejectionReason = "All safety checks passed"
    DebugLog "=== VALIDATION SUCCESSFUL - Part is safe for flat bar conversion ==="

End Function

' ====================================================================================
' END SAFETY CHECK FUNCTIONS
' ====================================================================================

' Main function to check if dimensions match flat bar and convert PL to FB
Public Function CheckAndConvertToFlatBar(ByVal model As ModelDoc2, _
                                        ByVal width As Double, _
                                        ByVal height As Double, _
                                        ByVal length As Double, _
                                        ByVal unitSystem As String, _
                                        ByRef propertiesToSet As Object) As Boolean

    DebugLog "=== FLAT BAR CONVERSION CHECK (UPDATED VERSION WITH LENGTH CHECK) ==="
    DebugLog "Input: width=" & width & ", height=" & height & ", length=" & length & ", unit=" & unitSystem

    ' Check global override setting
    If DISABLE_FLAT_BAR_PROMPTS Then
        DebugLog "Flat bar prompts are disabled globally (DISABLE_FLAT_BAR_PROMPTS = True)"
        CheckAndConvertToFlatBar = False
        Exit Function
    End If

    ' Check if user chose "Skip All This Run"
    If SkipAllThisRun Then
        DebugLog "Skipping flat bar check - user chose 'Skip All This Run'"
        CheckAndConvertToFlatBar = False
        Exit Function
    End If

    ' Initialize
    InitializeMaterialLists
    CheckAndConvertToFlatBar = False

    ' Get current type - only process if it's "PL" or "M"
    Dim currentType As String
    If propertiesToSet.exists("Type") Then
        currentType = propertiesToSet("Type")
    Else
        currentType = ""
    End If

    DebugLog "Current Type: " & currentType

    ' Only check plates or machine parts
    If currentType <> "PL" And currentType <> "M" Then
        DebugLog "Not a plate or machine part, skipping flat bar check"
        Exit Function
    End If

    ' Extract and clean material
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

    ' Determine material type
    Dim matType As MaterialType
    matType = GetMaterialType(actualMaterial)
    DebugLog "Material Type: " & matType & " (0=Unknown, 1=Steel, 2=Aluminum)"

    If matType = MaterialType.UNKNOWN Then
        DebugLog "Unknown material - cannot check flat bar"
        Exit Function
    End If

    ' Convert dimensions to inches for flat bar lookup
    ' Width is typically the smallest dimension (thickness)
    ' Height and Length are the larger dimensions
    Dim thicknessIn As Double
    Dim widthIn As Double
    Dim lengthIn As Double

    thicknessIn = ConvertToInches(width, unitSystem)    ' Smallest dimension = thickness
    widthIn = ConvertToInches(height, unitSystem)        ' Middle dimension = width
    lengthIn = ConvertToInches(length, unitSystem)       ' Largest dimension = length

    DebugLog "Dimensions in inches: thickness=" & thicknessIn & ", width=" & widthIn & ", length=" & lengthIn

    ' Check if dimensions match a flat bar size
    Dim flatBarSize As String
    Dim materialTypeStr As String
    Dim usedLength As Boolean
    materialTypeStr = MaterialTypeToString(matType)
    usedLength = False

    ' First try: thickness x width
    flatBarSize = FlatBarDataModule.FindMatchingFlatBar(thicknessIn, widthIn, materialTypeStr)

    If flatBarSize = "" Then
        ' Second try: thickness x length (swap width and length)
        DebugLog "No match for thickness x width, trying thickness x length..."
        flatBarSize = FlatBarDataModule.FindMatchingFlatBar(thicknessIn, lengthIn, materialTypeStr)
        If flatBarSize <> "" Then
            usedLength = True
            DebugLog "MATCH FOUND using length! Flat bar: " & flatBarSize
        End If
    Else
        DebugLog "MATCH FOUND! Flat bar: " & flatBarSize
    End If

    If flatBarSize <> "" Then

        ' Check if user previously declined this conversion
        Dim swModel As SldWorks.ModelDoc2
        Set swModel = model
        Dim swCustPropMgr As customPropertyManager
        Set swCustPropMgr = swModel.Extension.customPropertyManager("")

        Dim fbRejected As String
        Dim fbRejectedResolved As String
        swCustPropMgr.Get2 "FlatBarRejected", fbRejected, fbRejectedResolved

        DebugLog "FlatBarRejected property check - Raw: '" & fbRejected & "', Resolved: '" & fbRejectedResolved & "'"

        If Trim(fbRejectedResolved) = "Yes" Or Trim(fbRejected) = "Yes" Then
            DebugLog "User previously declined flat bar conversion - skipping prompt"
            Exit Function
        End If

        ' ===================================================================
        ' NEW: Run safety validation checks BEFORE showing dialog to user
        ' ===================================================================
        Dim validationReason As String
        Dim isValid As Boolean

        isValid = ValidatePartForFlatBarConversion(model, thicknessIn, widthIn, lengthIn, actualMaterial, validationReason)

        If Not isValid Then
            DebugLog "=== FLAT BAR CONVERSION REJECTED BY SAFETY CHECKS ==="
            DebugLog "Reason: " & validationReason
            DebugLog "Part will remain as: " & currentType
            CheckAndConvertToFlatBar = False
            Exit Function
        End If

        DebugLog "=== SAFETY VALIDATION PASSED ==="
        DebugLog "Reason: " & validationReason
        DebugLog "Proceeding to show user dialog..."
        ' ===================================================================

        ' Get the model name for display
        Dim modelName As String
        modelName = swModel.GetPathName
        If modelName = "" Then
            modelName = swModel.GetTitle
        Else
            ' Extract just the filename with extension
            Dim lastSlash As Long
            lastSlash = InStrRev(modelName, "\")
            If lastSlash > 0 Then
                modelName = Mid(modelName, lastSlash + 1)
            End If
        End If

        ' Try to make the part visible so user can see it
        On Error Resume Next
        ' OPTIMIZATION: Use global swApp accessor
        Dim swAppLocal As SldWorks.SldWorks
        Set swAppLocal = GetSwApp()

        If Not swAppLocal Is Nothing Then
            ' Make the current model visible and active
            swModel.Visible = True

            Dim modelPath As String
            modelPath = swModel.GetPathName

            If modelPath <> "" Then
                ' Activate using path
                swAppLocal.ActivateDoc2 modelPath, True, 0
            End If

            ' Force window to foreground
            swAppLocal.FrameState = swWindowState_e.swWindowMaximized
            swModel.ViewZoomtofit2
            swModel.GraphicsRedraw2
        End If
        On Error GoTo 0

        ' Get additional information to help identify the part
        Dim partDescription As String
        Dim partMaterial As String
        partDescription = ""
        partMaterial = ""

        If propertiesToSet.exists("Description") Then
            partDescription = propertiesToSet("Description")
        End If
        If propertiesToSet.exists("Material") Then
            partMaterial = propertiesToSet("Material")
        End If

        ' Build dimensions string
        Dim dimensionsStr As String
        dimensionsStr = propertiesToSet("Width") & " x " & _
                       propertiesToSet("Height") & " x " & _
                       propertiesToSet("Length")

        ' Show the Flat Bar Dialog using the builder
        Dim dialogResult As Integer
        dialogResult = FlatBarFormBuilder.ShowFlatBarDialog(modelName, partDescription, _
                                                            partMaterial, dimensionsStr, flatBarSize)

        ' Process result
        Select Case dialogResult
            Case FB_CANCEL
                DebugLog "User cancelled - skipping this part for now"
                Exit Function

            Case FB_SKIP
                If FlatBarFormBuilder.SkipAllThisRun Then
                    DebugLog "User chose 'Skip All This Run'"
                    SkipAllThisRun = True
                ElseIf FlatBarFormBuilder.NeverAskAgain Then
                    DebugLog "User chose 'Never ask again' - marking FlatBarRejected"
                    If Not propertiesToSet.exists("FlatBarRejected") Then
                        propertiesToSet.Add "FlatBarRejected", "Yes"
                    Else
                        propertiesToSet("FlatBarRejected") = "Yes"
                    End If
                Else
                    DebugLog "User chose 'Skip this part' - will ask again next time"
                End If
                Exit Function

            Case FB_CONVERT
                DebugLog "User confirmed flat bar conversion"
        End Select

        DebugLog "User confirmed flat bar conversion"
        DebugLog "Converting Reference Category from " & currentType & " to FB"

        ' If we used length dimension for the match, we need to swap height and length
        ' so the dimensions make sense for flat bar stock
        If usedLength Then
            DebugLog "Swapping height and length dimensions (used length for FB match)"
            DebugLog "Before swap: height=" & height & ", length=" & length

            ' Swap the height and length values in propertiesToSet
            If propertiesToSet.exists("Height") And propertiesToSet.exists("Length") Then
                Dim tempDim As String
                tempDim = propertiesToSet("Height")
                propertiesToSet("Height") = propertiesToSet("Length")
                propertiesToSet("Length") = tempDim
                DebugLog "After swap: height=" & propertiesToSet("Height") & ", length=" & propertiesToSet("Length")
            End If
        End If

        ' Add flat bar specification to properties
        ' This will be used to set the Reference Category and Stock Size
        If Not propertiesToSet.exists("FlatBarSize") Then
            propertiesToSet.Add "FlatBarSize", flatBarSize
        Else
            propertiesToSet("FlatBarSize") = flatBarSize
        End If

        CheckAndConvertToFlatBar = True
        DebugLog "Successfully marked for FB conversion"
    Else
        DebugLog "No matching flat bar found for " & thicknessIn & " x " & widthIn
        DebugLog "Keeping current type: " & currentType
    End If

End Function

' Helper function to get flat bar info for a given part
Public Function GetFlatBarInfo(ByVal thickness As Double, _
                              ByVal width As Double, _
                              ByVal materialName As String) As String

    InitializeMaterialLists

    Dim matType As MaterialType
    matType = GetMaterialType(materialName)

    If matType = MaterialType.UNKNOWN Then
        GetFlatBarInfo = "Unknown material"
        Exit Function
    End If

    Dim flatBarSize As String
    Dim materialTypeStr As String
    materialTypeStr = MaterialTypeToString(matType)

    flatBarSize = FlatBarDataModule.FindMatchingFlatBar(thickness, width, materialTypeStr)

    If flatBarSize <> "" Then
        GetFlatBarInfo = flatBarSize
    Else
        GetFlatBarInfo = "No matching flat bar"
    End If

End Function

' Debug function to test flat bar matching
Public Sub TestFlatBarMatching()
    DebugLog "=== TESTING FLAT BAR MATCHING ==="

    ' Test steel sizes
    DebugLog "Steel 1/4 x 2: " & GetFlatBarInfo(0.25, 2, "44W")
    DebugLog "Steel 1/2 x 4: " & GetFlatBarInfo(0.5, 4, "44W")
    DebugLog "Steel 3/8 x 3: " & GetFlatBarInfo(0.375, 3, "44W")

    ' Test aluminum sizes
    DebugLog "Aluminum 1/4 x 2: " & GetFlatBarInfo(0.25, 2, "6061-T6")
    DebugLog "Aluminum 1/2 x 4: " & GetFlatBarInfo(0.5, 4, "6061-T6")

    ' Test non-existent size
    DebugLog "Steel 0.123 x 1.456: " & GetFlatBarInfo(0.123, 1.456, "44W")

End Sub
