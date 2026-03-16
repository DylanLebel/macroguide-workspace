Attribute VB_Name = "NewPL1"
Option Explicit

Dim swApp As SldWorks.SldWorks

' ==================== MANUFACTURING CAPABILITY CHECKING MODULE ====================

' Constants for maximum dimensions
Private Const MAX_SHORT_SIDE_FT As Double = 6   ' 6 feet maximum for shorter side (Height)
Private Const MAX_LONG_SIDE_FT As Double = 20   ' 20 feet maximum for longer side (Length)

' Material-specific thickness limits (in inches)
Private Const MAX_STEEL_THICKNESS_IN As Double = 1.5      ' Steel thickness limit
Private Const MAX_ALUMINUM_THICKNESS_IN As Double = 1.5  ' Aluminum thickness limit

' Material type enumerations
Private Enum MaterialType
    Steel = 1
    Aluminum = 2
    Unknown = 0
End Enum

' Material lists structure
Private Type MaterialLists
    CarbonSteels() As String
    stainlessSteels() As String
    aluminums() As String
End Type

Private Materials As MaterialLists
Private materialsInitialized As Boolean

' Initialize material lists
Private Sub InitializeMaterialLists()
    If materialsInitialized Then Exit Sub
    
    Dim i As Long
    
    ' Carbon steels
    Dim tempCarbonSteels As Variant
    tempCarbonSteels = Array("44W", "50W", "300W", "350W", "400W", "QT100", "44W-PERF")
    ReDim Materials.CarbonSteels(LBound(tempCarbonSteels) To UBound(tempCarbonSteels))
    For i = LBound(tempCarbonSteels) To UBound(tempCarbonSteels)
        Materials.CarbonSteels(i) = tempCarbonSteels(i)
    Next i
    
    ' Stainless Steels
    Dim tempStainlessSteels As Variant
    tempStainlessSteels = Array("316", "304", "308", "309")
    ReDim Materials.stainlessSteels(LBound(tempStainlessSteels) To UBound(tempStainlessSteels))
    For i = LBound(tempStainlessSteels) To UBound(tempStainlessSteels)
        Materials.stainlessSteels(i) = tempStainlessSteels(i)
    Next i
    
    ' Aluminums
    Dim tempAluminums As Variant
    tempAluminums = Array("3003-H12", "3003-H14", "5052-H19", "5052-H32", _
                     "5083-H32", "5083-H34", "5083-H116", "5083-H321", _
                     "6061-T6", "6063-O")
    ReDim Materials.aluminums(LBound(tempAluminums) To UBound(tempAluminums))
    For i = LBound(tempAluminums) To UBound(tempAluminums)
        Materials.aluminums(i) = tempAluminums(i)
    Next i
    
    materialsInitialized = True
End Sub

' Function to determine material type
Private Function GetMaterialType(ByVal materialName As String) As MaterialType
    InitializeMaterialLists
    
    Dim i As Long
    
    ' Check Carbon Steels
    For i = LBound(Materials.CarbonSteels) To UBound(Materials.CarbonSteels)
        If UCase(materialName) = UCase(Materials.CarbonSteels(i)) Then
            GetMaterialType = MaterialType.Steel
            Exit Function
        End If
    Next i
    
    ' Check Stainless Steels
    For i = LBound(Materials.stainlessSteels) To UBound(Materials.stainlessSteels)
        If UCase(materialName) = UCase(Materials.stainlessSteels(i)) Then
            GetMaterialType = MaterialType.Steel
            Exit Function
        End If
    Next i
    
    ' Check Aluminums
    For i = LBound(Materials.aluminums) To UBound(Materials.aluminums)
        If UCase(materialName) = UCase(Materials.aluminums(i)) Then
            GetMaterialType = MaterialType.Aluminum
            Exit Function
        End If
    Next i
    
    GetMaterialType = MaterialType.Unknown
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

' Function to convert dimensions to feet
Private Function ConvertToFeet(ByVal value As Double, ByVal fromUnits As String) As Double
    Select Case UCase(fromUnits)
        Case "INCHES", "IN", "INCH"
            ConvertToFeet = value / 12        ' Inches to feet
        Case "MILLIMETERS", "MM", "MILLIMETER"
            ConvertToFeet = value / 304.8     ' Millimeters to feet
        Case "CENTIMETERS", "CM", "CENTIMETER"
            ConvertToFeet = value / 30.48     ' Centimeters to feet
        Case "METERS", "M", "METER"
            ConvertToFeet = value * 3.28084   ' Meters to feet
        Case "FEET", "FT", "FOOT"
            ConvertToFeet = value             ' Already in feet
        Case Else
            ConvertToFeet = value / 12        ' Default assume inches
    End Select
End Function

Function DetermineManufacturingType(ByVal allDimensions As Variant, ByVal materialGrade As String, ByVal itemName As String) As String
    ' Determine if plate should be Manufactured (M) or Purchased (P)
    ' allDimensions = [thickness, width, length] in document units
    
    If IsEmpty(allDimensions) Then
        DetermineManufacturingType = "P"
        Exit Function
    End If
    
    ' Extract dimensions: thickness (smallest), width (middle), length (largest)
    Dim thickness As Double, width As Double, length As Double
    thickness = allDimensions(0)  ' Smallest dimension
    width = allDimensions(1)      ' Middle dimension
    length = allDimensions(2)     ' Largest dimension
    
    ' Get document units for conversion
    Dim docUnits As String
    docUnits = GetDocumentUnits()
    
    ' Convert dimensions to appropriate units
    Dim thicknessInches As Double, widthFt As Double, lengthFt As Double
    thicknessInches = ConvertDimensionToInches(thickness, itemName & " thickness")
    widthFt = ConvertToFeet(width, docUnits)
    lengthFt = ConvertToFeet(length, docUnits)
    
    ' Get material type and thickness limit
    Dim matType As MaterialType
    matType = GetMaterialType(materialGrade)
    
    If matType = MaterialType.Unknown Then
        DetermineManufacturingType = "P"
        Exit Function
    End If
    
    Dim maxThickness As Double
    maxThickness = GetMaxThickness(matType)
    
    ' Check manufacturing limits
    Dim exceedsSizeLimit As Boolean
    Dim exceedsThicknessLimit As Boolean
    
    ' Size check: width (shorter side) = 6', length (longer side) = 20'
    exceedsSizeLimit = (widthFt > MAX_SHORT_SIDE_FT) Or (lengthFt > MAX_LONG_SIDE_FT)
    
    ' Thickness check
    exceedsThicknessLimit = (thicknessInches > maxThickness)
    
    ' Determine manufacturing type
    If exceedsSizeLimit Or exceedsThicknessLimit Then
        DetermineManufacturingType = "P"
    Else
        DetermineManufacturingType = "M"
    End If
End Function

' ==================== END MANUFACTURING CAPABILITY MODULE ====================

Function GetCurrentDecimalPlaces() As Long
    ' Get the current decimal places setting for linear dimensions
    
    ' Ensure SolidWorks application is connected
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            GetCurrentDecimalPlaces = -1
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        GetCurrentDecimalPlaces = -1
        Exit Function
    End If
    
    ' Try different preference keys for decimal places
    Dim decimalPlaces As Long
    Dim success As Boolean
    success = False
    
    On Error Resume Next
    
    ' Try swUnitsLinearDecimalPlaces constant first
    decimalPlaces = swModel.GetUserPreferenceIntegerValue(swUnitsLinearDecimalPlaces)
    If Err.Number = 0 Then
        success = True
    End If
    
    ' Try key 72 as fallback
    If Not success Then
        Err.Clear
        decimalPlaces = swModel.GetUserPreferenceIntegerValue(72)
        If Err.Number = 0 Then
            success = True
        End If
    End If
    
    ' Try alternative keys if others fail
    If Not success Then
        Err.Clear
        decimalPlaces = swModel.GetUserPreferenceIntegerValue(73)
        If Err.Number = 0 Then
            success = True
        End If
    End If
    
    On Error GoTo 0
    
    If success Then
        GetCurrentDecimalPlaces = decimalPlaces
    Else
        GetCurrentDecimalPlaces = 3 ' Default fallback
    End If
End Function

Function SetDecimalPlaces(ByVal newDecimalPlaces As Long) As Boolean
    ' Set the decimal places for linear dimensions
    
    ' Ensure SolidWorks application is connected
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            SetDecimalPlaces = False
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        SetDecimalPlaces = False
        Exit Function
    End If
    
    Dim success As Boolean
    success = False
    
    On Error Resume Next
    
    ' Use the SolidWorks constant that actually works
    swModel.SetUserPreferenceIntegerValue swUnitsLinearDecimalPlaces, newDecimalPlaces
    If Err.Number = 0 Then
        success = True
    Else
        ' Fallback to key 72 if constant fails
        Err.Clear
        swModel.SetUserPreferenceIntegerValue 72, newDecimalPlaces
        If Err.Number = 0 Then
            success = True
        End If
    End If
    
    On Error GoTo 0
    
    If success Then
        ' Force document rebuild to apply changes
        swModel.ForceRebuild3 False
    End If
    
    SetDecimalPlaces = success
End Function

Sub main()
    Debug.Print "Starting macro execution..."
    
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        Debug.Print "ERROR: Could not connect to SolidWorks application."
        Exit Sub
    End If
    Debug.Print "Connected to SolidWorks"

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        Debug.Print "ERROR: No active SolidWorks document found."
        Exit Sub
    End If
    Debug.Print "Found active document: " & swModel.GetTitle
    
    ' Check document type
    Dim docType As Long
    docType = swModel.GetType
    Debug.Print "Document type: " & docType & " (Part=1, Assembly=2, Drawing=3)"
    
    ' Process the cutlist items
    If swModel.GetType = swDocPART Then
        Debug.Print "Processing as Part document"
        ProcessWeldmentCutList swModel
    ElseIf swModel.GetType = swDocASSEMBLY Then
        Debug.Print "Processing as Assembly document"
        Dim swAssembly As SldWorks.AssemblyDoc
        Set swAssembly = swModel
        
        Dim swConf As SldWorks.Configuration
        Dim swRootComp As SldWorks.Component2
        
        Set swConf = swModel.GetActiveConfiguration
        If Not swConf Is Nothing Then
            Set swRootComp = swConf.GetRootComponent3(True)
            If Not swRootComp Is Nothing Then
                TraverseComponents swRootComp
            End If
        End If
    Else
        Debug.Print "ERROR: Unsupported document type: " & docType
    End If

    Debug.Print "Macro execution completed"
End Sub

Sub ProcessWeldmentCutList(ByVal swModelDoc As SldWorks.ModelDoc2)
    Debug.Print "Entering ProcessWeldmentCutList function"
    
    If swModelDoc Is Nothing Then
        Debug.Print "ERROR: swModelDoc is Nothing - exiting"
        Exit Sub
    End If

    Dim swFeat As SldWorks.Feature
    Set swFeat = swModelDoc.FirstFeature
    
    If swFeat Is Nothing Then
        Debug.Print "ERROR: No features found in document"
        Exit Sub
    End If
    Debug.Print "Found first feature: " & swFeat.Name
    
    Dim cutlistCount As Integer
    Dim plCount As Integer
    Dim totalFeatureCount As Integer
    cutlistCount = 0
    plCount = 0
    totalFeatureCount = 0
    
    ' First, try to process cutlist folders (existing logic)
    Do While Not swFeat Is Nothing
        totalFeatureCount = totalFeatureCount + 1
        Dim currentFeatName As String
        currentFeatName = swFeat.Name
        Dim currentFeatType As String
        currentFeatType = swFeat.GetTypeName2
        
        ' Show every 10th feature to avoid too much output
        If totalFeatureCount Mod 10 = 1 Then
            Debug.Print "Feature #" & totalFeatureCount & ": " & currentFeatName & " (Type: " & currentFeatType & ")"
        End If
        
        If currentFeatType = "CutListFolder" Then
            cutlistCount = cutlistCount + 1
            Debug.Print "Found CutListFolder #" & cutlistCount & ": " & currentFeatName
            
            ' Rest of cutlist processing...
            Dim swCustPropMgr As SldWorks.CustomPropertyManager
            Set swCustPropMgr = swFeat.CustomPropertyManager

            If Not swCustPropMgr Is Nothing Then
                Debug.Print "Got CustomPropertyManager for: " & currentFeatName
                
                ' Check for PL properties
                Dim vPropNames As Variant
                vPropNames = swCustPropMgr.GetNames
                Dim isPlate As Boolean: isPlate = False
                
                If Not IsNull(vPropNames) And IsArray(vPropNames) Then
                    Dim i As Long
                    For i = LBound(vPropNames) To UBound(vPropNames)
                        Dim propName As String: propName = vPropNames(i)
                        Dim propValue As String: propValue = GetPropertyValue(swCustPropMgr, propName)
                        
                        Debug.Print "  Property: " & propName & " = " & propValue
                        
                        ' Look for "PL" in any property value
                        If UCase(Trim(propValue)) = "PL" Then
                            isPlate = True
                            Debug.Print "  *** FOUND PL in property '" & propName & "'"
                        End If
                    Next i
                End If
                
                If isPlate Then
                    plCount = plCount + 1
                    Debug.Print "Processing PL item #" & plCount & ": " & currentFeatName
                    ProcessPLItem swCustPropMgr, currentFeatName, Nothing
                Else
                    Debug.Print "Not a PL item: " & currentFeatName
                End If
            Else
                Debug.Print "No CustomPropertyManager for: " & currentFeatName
            End If
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    Debug.Print "Feature scan complete. Total features: " & totalFeatureCount & ", CutListFolders: " & cutlistCount & ", PL items: " & plCount
    
    ' If no cutlist was found, try document-level properties
    If cutlistCount = 0 Then
        Debug.Print "No cutlist found - trying document-level processing"
        ProcessDocumentLevelPLProperties swModelDoc
    End If
End Sub

Sub ProcessDocumentLevelPLProperties(ByVal swModelDoc As SldWorks.ModelDoc2)
    Debug.Print "Entering ProcessDocumentLevelPLProperties"
    
    ' Get document-level custom properties
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModelDoc.Extension.CustomPropertyManager("")
    
    If swCustPropMgr Is Nothing Then
        Debug.Print "ERROR: Could not get document CustomPropertyManager"
        Exit Sub
    End If
    Debug.Print "Got document CustomPropertyManager"
    
    ' List all document properties
    Dim vPropNames As Variant
    vPropNames = swCustPropMgr.GetNames
    
    If IsNull(vPropNames) Or Not IsArray(vPropNames) Then
        Debug.Print "No document properties found"
    Else
        Debug.Print "Document properties found:"
        Dim i As Long
        For i = LBound(vPropNames) To UBound(vPropNames)
            Dim propName As String: propName = vPropNames(i)
            Dim propValue As String: propValue = GetPropertyValue(swCustPropMgr, propName)
            Debug.Print "  " & propName & " = " & propValue
        Next i
    End If
    
    ' Check if this is a plate part
    Dim isPlate As Boolean
    isPlate = CheckIfDocumentIsPlate(swCustPropMgr, swModelDoc)
    Debug.Print "Is plate check result: " & isPlate
    
    If isPlate Then
        Debug.Print "Processing as plate part"
        ' Get the main body of the part
        Dim swBody As SldWorks.Body2
        Set swBody = GetMainPartBody(swModelDoc)
        
        If swBody Is Nothing Then
            Debug.Print "ERROR: Could not get main part body"
        Else
            Debug.Print "Got main part body - processing PL item"
            ' Process as a PL item
            ProcessPLItem swCustPropMgr, "Document-Level PL", swBody
        End If
    Else
        Debug.Print "Not identified as a plate part"
    End If
End Sub

Function CheckIfDocumentIsPlate(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal swModelDoc As SldWorks.ModelDoc2) As Boolean
    Debug.Print "Checking if document is a plate..."
    
    ' Method 1: Check for existing PL property
    Dim vPropNames As Variant
    vPropNames = swCustPropMgr.GetNames
    
    If Not IsNull(vPropNames) And IsArray(vPropNames) Then
        Dim i As Long
        For i = LBound(vPropNames) To UBound(vPropNames)
            Dim propName As String: propName = vPropNames(i)
            Dim propValue As String: propValue = GetPropertyValue(swCustPropMgr, propName)
            
            ' Look for "PL" in any property value
            If UCase(Trim(propValue)) = "PL" Then
                Debug.Print "Found PL in document property: " & propName
                CheckIfDocumentIsPlate = True
                Exit Function
            End If
            
            ' Also check for other plate indicators
            If UCase(propName) = "TYPE" Or UCase(propName) = "PART_TYPE" Then
                If UCase(Trim(propValue)) = "PLATE" Or UCase(Trim(propValue)) = "PL" Then
                    Debug.Print "Found PLATE indicator in property: " & propName
                    CheckIfDocumentIsPlate = True
                    Exit Function
                End If
            End If
        Next i
    End If
    
    ' Method 2: Check part geometry - if it's thin and rectangular, might be a plate
    Dim swBody As SldWorks.Body2
    Set swBody = GetMainPartBody(swModelDoc)
    
    If Not swBody Is Nothing Then
        If IsGeometryPlatelike(swBody) Then
            Debug.Print "Geometry appears plate-like"
            CheckIfDocumentIsPlate = True
            Exit Function
        Else
            Debug.Print "Geometry does not appear plate-like"
        End If
    Else
        Debug.Print "Could not get body for geometry check"
    End If
    
    ' Method 3: Check material properties for plate materials
    Dim materialName As String
    materialName = GetPropertyValue(swCustPropMgr, "Material")
    If materialName = "" Then materialName = GetPropertyValue(swCustPropMgr, "MATERIAL")
    If materialName = "" Then materialName = GetPropertyValue(swCustPropMgr, "Grade")
    
    Debug.Print "Material name: " & materialName
    
    If materialName <> "" Then
        If IsPlateMaterial(materialName) Then
            Debug.Print "Material is a known plate material"
            CheckIfDocumentIsPlate = True
            Exit Function
        Else
            Debug.Print "Material is not a known plate material"
        End If
    End If
    
    Debug.Print "Document does not appear to be a plate"
    CheckIfDocumentIsPlate = False
End Function

Function GetMainPartBody(ByVal swModelDoc As SldWorks.ModelDoc2) As SldWorks.Body2
    Debug.Print "Getting main part body..."
    
    If swModelDoc.GetType <> swDocPART Then
        Debug.Print "ERROR: Document is not a part"
        Set GetMainPartBody = Nothing
        Exit Function
    End If
    
    Dim swPartDoc As SldWorks.PartDoc
    Set swPartDoc = swModelDoc
    
    Dim vBodies As Variant
    vBodies = swPartDoc.GetBodies2(swSolidBody, True)
    
    If Not IsEmpty(vBodies) And IsArray(vBodies) Then
        If UBound(vBodies) >= LBound(vBodies) Then
            Set GetMainPartBody = vBodies(LBound(vBodies)) ' Get first body
            Debug.Print "Found main part body"
            Exit Function
        End If
    End If
    
    Debug.Print "No solid bodies found"
    Set GetMainPartBody = Nothing
End Function

Function IsGeometryPlatelike(ByVal swBody As SldWorks.Body2) As Boolean
    Debug.Print "Checking if geometry is suitable for plate manufacturing..."
    
    If swBody Is Nothing Then
        Debug.Print "Body is Nothing"
        IsGeometryPlatelike = False
        Exit Function
    End If
    
    Dim vBodyBox As Variant
    vBodyBox = swBody.GetBodyBox ' Returns in METERS
    
    If IsEmpty(vBodyBox) Or IsNull(vBodyBox) Then
        Debug.Print "Could not get body bounding box"
        IsGeometryPlatelike = False
        Exit Function
    End If
    
    ' Calculate dimensions
    Dim dimX As Double, dimY As Double, dimZ As Double
    dimX = Abs(vBodyBox(3) - vBodyBox(0))
    dimY = Abs(vBodyBox(4) - vBodyBox(1))
    dimZ = Abs(vBodyBox(5) - vBodyBox(2))
    
    Debug.Print "Body dimensions (meters): " & dimX & " x " & dimY & " x " & dimZ
    
    ' Sort dimensions to find thickness (smallest)
    Dim dims(2) As Double
    dims(0) = dimX: dims(1) = dimY: dims(2) = dimZ
    
    ' Simple sort
    Dim i As Integer, j As Integer
    For i = 0 To 1
        For j = i + 1 To 2
            If dims(i) > dims(j) Then
                Dim temp As Double
                temp = dims(i): dims(i) = dims(j): dims(j) = temp
            End If
        Next j
    Next i
    
    ' Convert to manufacturing units
    Dim thickness_in As Double, width_ft As Double, length_ft As Double
    thickness_in = dims(0) * 39.3701  ' Convert to inches
    width_ft = dims(1) * 3.28084      ' Convert to feet
    length_ft = dims(2) * 3.28084     ' Convert to feet
    
    Debug.Print "Manufacturing dimensions: T=" & Format(thickness_in, "0.00") & " in, W=" & Format(width_ft, "0.00") & " ft, L=" & Format(length_ft, "0.00") & " ft"
    
    ' CHECK AGAINST YOUR ACTUAL MANUFACTURING LIMITS
    ' Thickness: 0.01" to 1.5" (reasonable plate thickness range)
    ' Width: up to 6 feet
    ' Length: up to 20 feet
    
    Dim validThickness As Boolean
    Dim validWidth As Boolean
    Dim validLength As Boolean
    
    validThickness = (thickness_in >= 0.01 And thickness_in <= 1.5)
    validWidth = (width_ft > 0 And width_ft <= 6)
    validLength = (length_ft > 0 And length_ft <= 20)
    
    Debug.Print "Thickness valid (0.01-1.5 in): " & validThickness & " (" & Format(thickness_in, "0.000") & ")"
    Debug.Print "Width valid (0-6 ft): " & validWidth & " (" & Format(width_ft, "0.00") & ")"
    Debug.Print "Length valid (0-20 ft): " & validLength & " (" & Format(length_ft, "0.00") & ")"
    
    ' Also check that it's reasonably rectangular (not a thin rod or wire)
    ' The smallest dimension should be the thickness for a plate
    Dim isRectangularish As Boolean
    isRectangularish = (dims(1) > dims(0) * 2) And (dims(2) > dims(0) * 2)
    Debug.Print "Rectangular check (width & length > 2x thickness): " & isRectangularish
    
    If validThickness And validWidth And validLength And isRectangularish Then
        Debug.Print "*** GEOMETRY PASSES: Part is suitable for plate manufacturing"
        IsGeometryPlatelike = True
    Else
        Debug.Print "*** GEOMETRY FAILS: Part exceeds manufacturing limits or wrong shape"
        IsGeometryPlatelike = False
    End If
End Function

Function IsPlateMaterial(ByVal materialName As String) As Boolean
    Debug.Print "Checking if material is plate material: " & materialName
    
    InitializeMaterialLists
    
    Dim i As Long
    
    ' Check against known plate materials
    For i = LBound(Materials.CarbonSteels) To UBound(Materials.CarbonSteels)
        If UCase(materialName) = UCase(Materials.CarbonSteels(i)) Then
            Debug.Print "Material matches carbon steel: " & Materials.CarbonSteels(i)
            IsPlateMaterial = True
            Exit Function
        End If
    Next i
    
    For i = LBound(Materials.stainlessSteels) To UBound(Materials.stainlessSteels)
        If UCase(materialName) = UCase(Materials.stainlessSteels(i)) Then
            Debug.Print "Material matches stainless steel: " & Materials.stainlessSteels(i)
            IsPlateMaterial = True
            Exit Function
        End If
    Next i
    
    For i = LBound(Materials.aluminums) To UBound(Materials.aluminums)
        If UCase(materialName) = UCase(Materials.aluminums(i)) Then
            Debug.Print "Material matches aluminum: " & Materials.aluminums(i)
            IsPlateMaterial = True
            Exit Function
        End If
    Next i
    
    Debug.Print "Material is not a known plate material"
    IsPlateMaterial = False
End Function



Function GetDocumentUnits() As String
    Debug.Print "*** Getting document units using proper API..."
    
    ' Ensure SolidWorks application is connected
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            Debug.Print "ERROR: Could not connect to SolidWorks app"
            GetDocumentUnits = "INCHES"
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        Debug.Print "ERROR: No active document"
        GetDocumentUnits = "INCHES"
        Exit Function
    End If
    
    ' Use the proper GetModelUnitOfMeasure function
    Dim modelUOM As String
    modelUOM = GetModelUnitOfMeasure(swModel)
    Debug.Print "GetModelUnitOfMeasure returned: " & modelUOM
    
    ' Convert to the format expected by the rest of the code
    Select Case LCase(modelUOM)
        Case "mm"
            GetDocumentUnits = "MILLIMETERS"
        Case "cm"
            GetDocumentUnits = "CENTIMETERS"
        Case "m"
            GetDocumentUnits = "METERS"
        Case "in"
            GetDocumentUnits = "INCHES"
        Case "ft"
            GetDocumentUnits = "FEET"
        Case "ft-in"
            GetDocumentUnits = "INCHES" ' Feet-inches uses inches as base unit
        Case Else
            Debug.Print "Unknown unit from API: " & modelUOM & ", defaulting to INCHES"
            GetDocumentUnits = "INCHES"
    End Select
    
    Debug.Print "*** Final document units: " & GetDocumentUnits
End Function

Function GetModelUnitOfMeasure(model As ModelDoc2) As String
    Dim unitType As Integer
    unitType = model.GetUserPreferenceIntegerValue(swUnitsLinear)
    
    Debug.Print "Raw unit type from SolidWorks API: " & unitType
    
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
            Debug.Print "WARNING: Unknown unit type: " & unitType
            GetModelUnitOfMeasure = "unknown"
    End Select
    
    Debug.Print "Interpreted as: " & GetModelUnitOfMeasure
End Function

Function DetectUnitsFromCutlistValues() As String
    ' Analyze actual cutlist thickness values to determine units
    
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            DetectUnitsFromCutlistValues = ""
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        DetectUnitsFromCutlistValues = ""
        Exit Function
    End If
    
    Dim thicknessSamples() As Double
    Dim sampleCount As Integer
    sampleCount = 0
    
    ' Collect thickness samples from cutlist StockSize values
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing And sampleCount < 10 ' Limit samples
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As SldWorks.CustomPropertyManager
            Set swCustPropMgr = swFeat.CustomPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim stockSize As String
                stockSize = GetPropertyValue(swCustPropMgr, "StockSize")
                If stockSize <> "" Then
                    Dim thickness As String
                    thickness = ParseThicknessFromStockSize(stockSize)
                    If thickness <> "" And IsNumeric(thickness) Then
                        ReDim Preserve thicknessSamples(sampleCount)
                        thicknessSamples(sampleCount) = CDbl(thickness)
                        sampleCount = sampleCount + 1
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    If sampleCount = 0 Then
        DetectUnitsFromCutlistValues = ""
        Exit Function
    End If
    
    ' Analyze the samples
    Dim avgThickness As Double
    Dim minThickness As Double
    Dim maxThickness As Double
    Dim i As Integer
    
    avgThickness = 0
    minThickness = thicknessSamples(0)
    maxThickness = thicknessSamples(0)
    
    For i = 0 To sampleCount - 1
        avgThickness = avgThickness + thicknessSamples(i)
        If thicknessSamples(i) < minThickness Then minThickness = thicknessSamples(i)
        If thicknessSamples(i) > maxThickness Then maxThickness = thicknessSamples(i)
    Next i
    
    avgThickness = avgThickness / sampleCount
    
    ' Count imperial-like values (common fractions)
    Dim imperialCount As Integer
    Dim metricCount As Integer
    imperialCount = 0
    metricCount = 0
    
    For i = 0 To sampleCount - 1
        Dim val As Double
        val = thicknessSamples(i)
        
        ' Check for common imperial fractions
        If IsLikelyImperialThickness(val) Then
            imperialCount = imperialCount + 1
        End If
        
        ' Check for common metric thicknesses
        If IsLikelyMetricThickness(val) Then
            metricCount = metricCount + 1
        End If
    Next i
    
    ' Decision logic
    If imperialCount > metricCount And imperialCount >= sampleCount / 2 Then
        DetectUnitsFromCutlistValues = "INCHES"
    ElseIf metricCount > imperialCount And metricCount >= sampleCount / 2 Then
        DetectUnitsFromCutlistValues = "MILLIMETERS"
    ElseIf avgThickness <= 6 And minThickness >= 0.01 Then
        ' Range suggests inches
        DetectUnitsFromCutlistValues = "INCHES"
    ElseIf avgThickness >= 3 And avgThickness <= 150 Then
        ' Range suggests millimeters
        DetectUnitsFromCutlistValues = "MILLIMETERS"
    Else
        DetectUnitsFromCutlistValues = ""
    End If
End Function

Function IsLikelyImperialThickness(ByVal value As Double) As Boolean
    ' Check if value matches common imperial plate thicknesses
    Dim commonImperial As Variant
    commonImperial = Array(0.0625, 0.125, 0.1875, 0.25, 0.3125, 0.375, 0.4375, 0.5, 0.5625, 0.625, 0.6875, 0.75, 0.875, 1, 1.25, 1.5, 2, 2.5, 3)
    
    Dim i As Integer
    For i = 0 To UBound(commonImperial)
        If Abs(value - commonImperial(i)) < 0.02 Then ' Small tolerance
            IsLikelyImperialThickness = True
            Exit Function
        End If
    Next i
    
    ' Also check if it's a clean decimal that could be imperial
    If value >= 0.01 And value <= 6 Then
        ' Check if it's close to a fraction (like 0.125, 0.25, 0.375, 0.5, etc.)
        Dim fractional As Double
        fractional = value * 8 ' Convert to eighths
        If Abs(fractional - Round(fractional)) < 0.1 Then
            IsLikelyImperialThickness = True
            Exit Function
        End If
    End If
    
    IsLikelyImperialThickness = False
End Function

Function IsLikelyMetricThickness(ByVal value As Double) As Boolean
    ' Check if value matches common metric plate thicknesses
    Dim commonMetric As Variant
    commonMetric = Array(1, 1.5, 2, 3, 4, 5, 6, 8, 10, 12, 15, 16, 20, 25, 30, 32, 40, 50, 60, 80, 100)
    
    Dim i As Integer
    For i = 0 To UBound(commonMetric)
        If Abs(value - commonMetric(i)) < 0.5 Then ' Tolerance for metric
            IsLikelyMetricThickness = True
            Exit Function
        End If
    Next i
    
    IsLikelyMetricThickness = False
End Function

Function ValidateUnitsWithSampleData(ByVal detectedUnits As String) As Boolean
    ' Validate detected units against actual thickness values
    Dim sampleThickness As Double
    sampleThickness = GetFirstThicknessSample()
    
    If sampleThickness <= 0 Then
        ValidateUnitsWithSampleData = True ' Can't validate, assume correct
        Exit Function
    End If
    
    Select Case UCase(detectedUnits)
        Case "MILLIMETERS"
            ' For mm, expect thickness roughly 1-100mm for plates
            ValidateUnitsWithSampleData = (sampleThickness >= 0.5 And sampleThickness <= 200)
        Case "INCHES"
            ' For inches, expect thickness roughly 0.01-6 inches for plates
            ValidateUnitsWithSampleData = (sampleThickness >= 0.005 And sampleThickness <= 10)
        Case Else
            ValidateUnitsWithSampleData = True
    End Select
End Function

Function GetFirstThicknessSample() As Double
    ' Get first available thickness value for validation
    
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            GetFirstThicknessSample = 0
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        GetFirstThicknessSample = 0
        Exit Function
    End If
    
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As SldWorks.CustomPropertyManager
            Set swCustPropMgr = swFeat.CustomPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim stockSize As String
                stockSize = GetPropertyValue(swCustPropMgr, "StockSize")
                If stockSize <> "" Then
                    Dim thickness As String
                    thickness = ParseThicknessFromStockSize(stockSize)
                    If thickness <> "" And IsNumeric(thickness) Then
                        GetFirstThicknessSample = CDbl(thickness)
                        Exit Function
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    GetFirstThicknessSample = 0
End Function

Function CheckDocumentProperties() As String
    ' Check document properties for unit clues
    
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
        If swApp Is Nothing Then
            CheckDocumentProperties = ""
            Exit Function
        End If
    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        CheckDocumentProperties = ""
        Exit Function
    End If
    
    ' Check file properties that might indicate units
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")
    
    If Not swCustPropMgr Is Nothing Then
        ' Look for unit indicators in custom properties
        Dim unitProp As String
        unitProp = GetPropertyValue(swCustPropMgr, "Units")
        If unitProp <> "" Then
            If InStr(UCase(unitProp), "INCH") > 0 Or InStr(UCase(unitProp), "IPS") > 0 Then
                CheckDocumentProperties = "INCHES"
                Exit Function
            ElseIf InStr(UCase(unitProp), "MM") > 0 Or InStr(UCase(unitProp), "METRIC") > 0 Then
                CheckDocumentProperties = "MILLIMETERS"
                Exit Function
            End If
        End If
        
        ' Check weight units as indicator
        Dim weightUOM As String
        weightUOM = GetPropertyValue(swCustPropMgr, "WeightUOM")
        If weightUOM <> "" Then
            If UCase(weightUOM) = "LBS" Or UCase(weightUOM) = "POUNDS" Then
                CheckDocumentProperties = "INCHES"
                Exit Function
            ElseIf UCase(weightUOM) = "KG" Or UCase(weightUOM) = "GRAMS" Then
                CheckDocumentProperties = "MILLIMETERS"
                Exit Function
            End If
        End If
    End If
    
    CheckDocumentProperties = ""
End Function


Sub ProcessPLItem(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal itemName As String, Optional ByVal swPlateBody As SldWorks.Body2 = Nothing)
    Debug.Print "*** Starting ProcessPLItem for: " & itemName
    
    ' Get material/grade
    Dim materialGrade As String
    materialGrade = GetPropertyValue(swCustPropMgr, "Grade")
    Debug.Print "Grade property: " & materialGrade
    
    If materialGrade = "" Then
        materialGrade = GetPropertyValue(swCustPropMgr, "MATERIAL")
        Debug.Print "MATERIAL property: " & materialGrade
    End If
    If materialGrade = "" Then
        materialGrade = "300W" ' Default
        Debug.Print "Using default material: " & materialGrade
    End If
    
    Debug.Print "Final material grade: " & materialGrade
    
    ' Get thickness from various possible sources
    Dim thicknessRaw As String
    Debug.Print "About to extract thickness"
    thicknessRaw = ExtractThickness(swCustPropMgr, itemName, swPlateBody)
    Debug.Print "Thickness extracted: " & thicknessRaw
    
    ' Get all three dimensions for Mtl Unit Qty calculation and manufacturing check
    Dim allDimensions As Variant
    Debug.Print "Extracting all dimensions..."
    allDimensions = ExtractAllDimensions(swCustPropMgr, itemName, swPlateBody)
    
    If IsEmpty(allDimensions) Then
        Debug.Print "ERROR: Failed to extract dimensions"
    Else
        Debug.Print "All dimensions: T=" & allDimensions(0) & ", W=" & allDimensions(1) & ", L=" & allDimensions(2)
    End If
    
    ' Convert thickness to inches and round to nearest 32nd
    Dim thicknessInches As String
    If thicknessRaw <> "" Then
        Debug.Print "Converting thickness to inches..."
        Dim rawInches As String
        rawInches = ConvertThicknessToInches(thicknessRaw, itemName)
        Debug.Print "Thickness in inches (before rounding): " & rawInches
        
        thicknessInches = RoundToNearest32nd(rawInches)
        Debug.Print "Thickness rounded to 32nd: " & thicknessInches
    Else
        thicknessInches = ""
        Debug.Print "No thickness found"
    End If
    
    ' Calculate Mtl Unit Qty (length × width, excluding thickness)
    Dim mtlUnitQty As String
    Debug.Print "Calculating Mtl Unit Qty..."
    mtlUnitQty = CalculateMtlUnitQty(allDimensions, itemName)
    Debug.Print "Mtl Unit Qty calculated: " & mtlUnitQty
    
    ' Determine manufacturing type (M or P) based on dimensions and material
    Dim manufacturingType As String
    Debug.Print "Determining manufacturing type..."
    manufacturingType = DetermineManufacturingType(allDimensions, materialGrade, itemName)
    Debug.Print "Manufacturing type determined: " & manufacturingType
    
    ' Create Mtl Part Number in format PLT-thickness-material
    Dim mtlPartNumber As String
    If thicknessInches <> "" Then
        mtlPartNumber = "PLT-" & thicknessInches & "-" & materialGrade
    Else
        mtlPartNumber = "PLT-UNK-" & materialGrade
    End If
    Debug.Print "Mtl Part Number created: " & mtlPartNumber
    
    ' Set properties
    Debug.Print "Setting custom properties..."
    SetCustomProperty swCustPropMgr, "Mtl Part Number", mtlPartNumber
    SetCustomProperty swCustPropMgr, "Purchased Assembly", manufacturingType
    
    If mtlUnitQty <> "" Then
        SetCustomProperty swCustPropMgr, "Mtl Unit Qty", mtlUnitQty
    End If
    
    Debug.Print "*** Finished ProcessPLItem for: " & itemName
End Sub

Function ExtractThickness(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal itemName As String, Optional ByVal swPlateBody As SldWorks.Body2 = Nothing) As String
    Debug.Print "*** Starting ExtractThickness for: " & itemName
    
    ' Attempt 1: Sheet Metal Thickness property (most reliable for sheet metal)
    Dim sheetMetalThickness As String
    sheetMetalThickness = GetPropertyValue(swCustPropMgr, "Sheet Metal Thickness")
    Debug.Print "Sheet Metal Thickness: " & sheetMetalThickness
    
    If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
        Debug.Print "Using Sheet Metal Thickness: " & sheetMetalThickness
        ExtractThickness = sheetMetalThickness
        Exit Function
    End If

    ' Attempt 2: Fallback to cutlist "StockSize" property
    Dim stockSize As String
    stockSize = GetPropertyValue(swCustPropMgr, "StockSize")
    Debug.Print "StockSize property: " & stockSize
    
    If stockSize <> "" Then
        Dim thicknessFromStockSize As String
        thicknessFromStockSize = ParseThicknessFromStockSize(stockSize)
        Debug.Print "Parsed thickness from StockSize: " & thicknessFromStockSize
        
        If thicknessFromStockSize <> "" Then
            ExtractThickness = thicknessFromStockSize
            Exit Function
        End If
    End If

    ' Attempt 3: Extract directly from body geometry (NEW - for non-cutlist parts)
    If Not swPlateBody Is Nothing Then
        Debug.Print "Attempting to extract thickness from body geometry..."
        Dim bodyDimensions As Variant
        bodyDimensions = ExtractAllDimensionsFromBody(swPlateBody)
        
        If Not IsEmpty(bodyDimensions) Then
            ' bodyDimensions(0) is the thickness (smallest dimension)
            Dim thicknessValue As Double
            thicknessValue = bodyDimensions(0)
            Debug.Print "Thickness from body geometry: " & thicknessValue
            
            ' Convert to string with reasonable precision
            ExtractThickness = Format(thicknessValue, "0.######")
            Debug.Print "Formatted thickness: " & ExtractThickness
            Exit Function
        Else
            Debug.Print "Failed to get dimensions from body"
        End If
    Else
        Debug.Print "No body provided for geometry extraction"
    End If

    ' Attempt 4: Fallback to scanning all model dimensions
    Debug.Print "Attempting GetSmallestModelDimension..."
    Dim thicknessFromModelDims As String
    thicknessFromModelDims = GetSmallestModelDimension()
    Debug.Print "Thickness from model dimensions: " & thicknessFromModelDims

    If thicknessFromModelDims <> "" Then
        ExtractThickness = thicknessFromModelDims
        Exit Function
    End If
    
    Debug.Print "*** ExtractThickness failed - returning empty string"
    ExtractThickness = ""
End Function
Function GetSmallestModelDimension() As String
    ' Scan all dimensions in the model and return the smallest reasonable thickness
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        GetSmallestModelDimension = ""
        Exit Function
    End If
    
    Dim allDimensions() As Double
    Dim dimCount As Long
    dimCount = 0
    
    ' Get all features and scan their dimensions
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing
        ' Get display dimensions from this feature
        Dim vDispDims As Variant
       
        On Error Resume Next
        vDispDims = swFeat.GetDisplayDimensions2(-1)
        On Error GoTo 0
        
        If Not IsNull(vDispDims) And IsArray(vDispDims) Then
            Dim i As Long
            For i = 0 To UBound(vDispDims)
                Dim swDispDim As SldWorks.DisplayDimension
                Set swDispDim = vDispDims(i)
                
                If Not swDispDim Is Nothing Then
                    Dim swDim As SldWorks.dimension
                    Set swDim = swDispDim.GetDimension2(0)
                    
                    If Not swDim Is Nothing Then
                        Dim dimValue As Double
                        dimValue = swDim.GetValue2(-1) ' Get actual value in model units
                        
                        ' Convert to mm based on document units
                        Dim dimValueMM As Double
                        Dim docUnits As String
                        docUnits = GetDocumentUnits()
                        
                        If UCase(docUnits) = "MILLIMETERS" Then
                            dimValueMM = dimValue
                        ElseIf UCase(docUnits) = "INCHES" Then
                            dimValueMM = dimValue * 25.4
                        ElseIf UCase(docUnits) = "METERS" Then
                            dimValueMM = dimValue * 1000
                        Else
                            dimValueMM = dimValue ' Assume mm
                        End If
                        
                        ' Filter to reasonable thickness range (1mm to 100mm)
                        If dimValueMM >= 1 And dimValueMM <= 100 Then
                            ' Expand array and add dimension
                            ReDim Preserve allDimensions(dimCount)
                            allDimensions(dimCount) = dimValueMM
                            dimCount = dimCount + 1
                        End If
                    End If
                End If
            Next i
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    ' Find the smallest dimension
    If dimCount > 0 Then
        Dim smallestDim As Double
        smallestDim = allDimensions(0)
        
        Dim j As Long
        For j = 1 To dimCount - 1
            If allDimensions(j) < smallestDim Then
                smallestDim = allDimensions(j)
            End If
        Next j
        
        ' Format with appropriate precision
        If smallestDim = Int(smallestDim) Then
            GetSmallestModelDimension = CStr(Int(smallestDim))
        Else
            GetSmallestModelDimension = Format(smallestDim, "0.0")
        End If
    Else
        GetSmallestModelDimension = ""
    End If
End Function

Function ParseThicknessFromStockSize(ByVal stockSize As String) As String
    ' Remove quotes and clean up
    stockSize = Replace(stockSize, """", "")
    stockSize = Trim(stockSize)
    
    
    ' Handle feet notation like "3/4' x 2 1/2'"
    If InStr(stockSize, "'") > 0 Then
        stockSize = Replace(stockSize, "'", "")
        ' Note: This removes the feet marker but doesn't convert units
        ' You may need additional logic to convert feet to inches
    End If
    
    
    ' Split by common delimiters
    Dim parts() As String
    Dim delimiter As String
    
    If InStr(stockSize, " X ") > 0 Then
        delimiter = " X "
    ElseIf InStr(stockSize, "x") > 0 Then
        delimiter = "x"
    ElseIf InStr(stockSize, " x ") > 0 Then
        delimiter = " x "
    Else
        ' No clear delimiter, try to extract first number
        Dim i As Integer
        For i = 1 To Len(stockSize)
            If Not IsNumeric(Mid(stockSize, i, 1)) And Mid(stockSize, i, 1) <> "." Then
                If i > 1 Then
                    ParseThicknessFromStockSize = Left(stockSize, i - 1)
                    Exit Function
                End If
            End If
        Next i
        ParseThicknessFromStockSize = ""
        Exit Function
    End If
    
    parts = Split(stockSize, delimiter)
    If UBound(parts) >= 0 Then
        Dim firstPart As String
        firstPart = Trim(parts(0))
        If IsNumeric(firstPart) Then
            ParseThicknessFromStockSize = firstPart
        Else
            ParseThicknessFromStockSize = ""
        End If
    Else
        ParseThicknessFromStockSize = ""
    End If
End Function

Function ConvertThicknessToInches(ByVal thicknessValue As String, ByVal itemName As String) As String
    If Not IsNumeric(thicknessValue) Then
        ConvertThicknessToInches = thicknessValue
        Exit Function
    End If
    
    Dim thickness As Double
    thickness = CDbl(thicknessValue)
    
    ' Get document units
    Dim docUnits As String
    docUnits = GetDocumentUnits()
    
    Dim convertedThickness As Double
    Dim convertedStr As String
    
    Select Case UCase(docUnits)
        Case "INCHES", "IN", "INCH"
            ' Already in inches, no conversion needed
            convertedThickness = thickness
            
        Case "MILLIMETERS", "MM", "MILLIMETER"
            ' Convert mm to inches (divide by 25.4)
            convertedThickness = thickness / 25.4
            
        Case "CENTIMETERS", "CM", "CENTIMETER"
            ' Convert cm to inches (divide by 2.54)
            convertedThickness = thickness / 2.54
            
        Case "METERS", "M", "METER"
            ' Convert m to inches (multiply by 39.3701)
            convertedThickness = thickness * 39.3701
            
        Case "MICROMETERS", "UM", "MICROMETER"
            ' Convert micrometers to inches (divide by 25400)
            convertedThickness = thickness / 25400
            
        Case "FEET", "FT", "FOOT"
            ' Convert feet to inches (multiply by 12)
            convertedThickness = thickness * 12
            
        Case Else
            ' Unknown units, assume inches but warn
            convertedThickness = thickness
    End Select
    
    ' Format the result to a reasonable number of decimal places
    If convertedThickness >= 1 Then
        convertedStr = Format(convertedThickness, "0.000")
    ElseIf convertedThickness >= 0.001 Then
        convertedStr = Format(convertedThickness, "0.0000")
    Else
        convertedStr = Format(convertedThickness, "0.000000")
    End If
    
    ' Remove trailing zeros
    Do While Right(convertedStr, 1) = "0" And InStr(convertedStr, ".") > 0
        convertedStr = Left(convertedStr, Len(convertedStr) - 1)
    Loop
    
    ' Remove trailing decimal point if all decimals were zeros
    If Right(convertedStr, 1) = "." Then
        convertedStr = Left(convertedStr, Len(convertedStr) - 1)
    End If
    
    ConvertThicknessToInches = convertedStr
End Function

Function RoundToNearest32nd(ByVal thicknessInches As String) As String
    ' Round thickness to nearest 32nd of an inch for standard plate sizes
    If Not IsNumeric(thicknessInches) Then
        RoundToNearest32nd = thicknessInches
        Exit Function
    End If
    
    Dim thickness As Double
    thickness = CDbl(thicknessInches)
    
    ' Convert to 32nds, round, then convert back
    Dim thirtySeconds As Long
    thirtySeconds = CLng(thickness * 32 + 0.5) ' Add 0.5 for proper rounding
    
    ' Convert back to decimal inches
    Dim roundedThickness As Double
    roundedThickness = thirtySeconds / 32
    
    ' Format result - remove unnecessary decimal places
    Dim result As String
    If roundedThickness = Int(roundedThickness) Then
        ' Whole number
        result = CStr(Int(roundedThickness))
    ElseIf roundedThickness * 2 = Int(roundedThickness * 2) Then
        ' Half inch increment
        result = Format(roundedThickness, "0.#")
    ElseIf roundedThickness * 4 = Int(roundedThickness * 4) Then
        ' Quarter inch increment
        result = Format(roundedThickness, "0.##")
    ElseIf roundedThickness * 8 = Int(roundedThickness * 8) Then
        ' Eighth inch increment
        result = Format(roundedThickness, "0.###")
    ElseIf roundedThickness * 16 = Int(roundedThickness * 16) Then
        ' Sixteenth inch increment
        result = Format(roundedThickness, "0.####")
    Else
        ' Thirty-second inch increment
        result = Format(roundedThickness, "0.#####")
    End If
    
    RoundToNearest32nd = result
End Function

Function GetPropertyValue(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal propName As String) As String
    Dim propValue As String
    Dim resolvedValue As String
    Dim retVal As Long
    
    ' Try Get6 first
    retVal = swCustPropMgr.Get6(propName, False, propValue, resolvedValue, False, False)
    If retVal = 0 Then
        GetPropertyValue = resolvedValue
        Exit Function
    End If
    
    ' Try Get4 as fallback
    retVal = swCustPropMgr.Get4(propName, False, propValue, resolvedValue)
    If retVal = 0 Then
        GetPropertyValue = resolvedValue
    Else
        GetPropertyValue = ""
    End If
End Function

Sub TraverseComponents(ByVal swComp As SldWorks.Component2)
    If swComp Is Nothing Then Exit Sub

    Dim compModel As SldWorks.ModelDoc2
    Set compModel = swComp.GetModelDoc2

    If Not compModel Is Nothing Then
        Dim isActiveDoc As Boolean
        isActiveDoc = False
        If Not swApp.ActiveDoc Is Nothing And Not compModel Is Nothing Then
            If swApp.ActiveDoc.GetPathName <> "" And compModel.GetPathName <> "" Then
                If LCase(compModel.GetPathName) = LCase(swApp.ActiveDoc.GetPathName) Then
                    isActiveDoc = True
                End If
            ElseIf compModel Is swApp.ActiveDoc Then
                isActiveDoc = True
            End If
        End If

        If Not isActiveDoc Then
            If swComp.GetSuppression <> swComponentSuppressed And swComp.GetSuppression <> swComponentHiddenThenSuppressed Then
                If compModel.GetType = swDocPART Then
                    ProcessWeldmentCutList compModel
                End If
            End If
        End If
    End If

    Dim vChildren As Variant
    vChildren = swComp.GetChildren

    If Not IsNull(vChildren) Then
        Dim i As Long
        For i = LBound(vChildren) To UBound(vChildren)
            Dim swChildComp As SldWorks.Component2
            Set swChildComp = vChildren(i)
            If Not swChildComp Is Nothing Then
                TraverseComponents swChildComp
            End If
        Next i
    End If
End Sub

Function ExtractAllDimensions(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal itemName As String, Optional ByVal swPlateBody As SldWorks.Body2 = Nothing) As Variant
    ' Extract all three dimensions (thickness, width, length) from various sources
    ' Returns array: [thickness, width, length] in document units
    
    Dim dimensions(2) As Double ' [0]=thickness, [1]=width, [2]=length
    Dim i As Integer
    For i = 0 To 2
        dimensions(i) = 0
    Next i
    
    ' For sheet metal parts, get thickness from sheet metal property first
    Dim sheetMetalThickness As String
    sheetMetalThickness = GetPropertyValue(swCustPropMgr, "Sheet Metal Thickness")
    
    ' Attempt 1: Get dimensions from bounding box properties (most reliable for formed sheet metal)
    Dim thickness As String, width As String, length As String
    
    thickness = GetPropertyValue(swCustPropMgr, "3D-Bounding Box Thickness")
    width = GetPropertyValue(swCustPropMgr, "3D-Bounding Box Width")
    length = GetPropertyValue(swCustPropMgr, "3D-Bounding Box Length")
    
    If thickness <> "" And width <> "" And length <> "" Then
        If IsNumeric(thickness) And IsNumeric(width) And IsNumeric(length) Then
            ' Use sheet metal thickness if available, otherwise use bounding box thickness
            If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
                dimensions(0) = CDbl(sheetMetalThickness)
            Else
                dimensions(0) = CDbl(thickness)
            End If
            dimensions(1) = CDbl(width)
            dimensions(2) = CDbl(length)
            
            ' Sort width and length (keep thickness as is)
            If dimensions(1) > dimensions(2) Then
                Dim temp As Double
                temp = dimensions(1): dimensions(1) = dimensions(2): dimensions(2) = temp
            End If
            
            ExtractAllDimensions = dimensions
            Exit Function
        End If
    End If
    
    ' Attempt 2: Get dimensions from bounding box of body (works for both cutlist and non-cutlist parts)
    If Not swPlateBody Is Nothing Then
        Dim bodyDimensions As Variant
        bodyDimensions = ExtractAllDimensionsFromBody(swPlateBody)
        
        If Not IsEmpty(bodyDimensions) Then
            ' Override thickness with sheet metal thickness if available
            If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
                dimensions(0) = CDbl(sheetMetalThickness)
                dimensions(1) = bodyDimensions(1)
                dimensions(2) = bodyDimensions(2)
            Else
                For i = 0 To 2
                    dimensions(i) = bodyDimensions(i)
                Next i
            End If
            
            ExtractAllDimensions = dimensions
            Exit Function
        End If
    End If
    
    ' Attempt 3: Extract from individual bounding box properties (fallback)
    If thickness = "" Then thickness = GetPropertyValue(swCustPropMgr, "Bounding Box Thickness")
    If width = "" Then width = GetPropertyValue(swCustPropMgr, "Bounding Box Width")
    If length = "" Then length = GetPropertyValue(swCustPropMgr, "Bounding Box Length")
    
    If thickness <> "" And width <> "" And length <> "" Then
        If IsNumeric(thickness) And IsNumeric(width) And IsNumeric(length) Then
            ' Use sheet metal thickness if available, otherwise use bounding box thickness
            If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
                dimensions(0) = CDbl(sheetMetalThickness)
            Else
                dimensions(0) = CDbl(thickness)
            End If
            dimensions(1) = CDbl(width)
            dimensions(2) = CDbl(length)
            ExtractAllDimensions = dimensions
            Exit Function
        End If
    End If
    
    ' Attempt 4: Parse from StockSize if it contains all dimensions
    Dim stockSize As String
    stockSize = GetPropertyValue(swCustPropMgr, "StockSize")
    If stockSize <> "" Then
        Dim parsedDims As Variant
        parsedDims = ParseAllDimensionsFromStockSize(stockSize)
        If Not IsEmpty(parsedDims) Then
            ' Use sheet metal thickness if available, otherwise use parsed thickness
            If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
                dimensions(0) = CDbl(sheetMetalThickness)
                dimensions(1) = parsedDims(1)
                dimensions(2) = parsedDims(2)
            Else
                For i = 0 To 2
                    dimensions(i) = parsedDims(i)
                Next i
            End If
            ExtractAllDimensions = dimensions
            Exit Function
        End If
    End If
    
    ' Attempt 5: Try individual property names
    Dim propWidth As String, propLength As String, propHeight As String
    
    propWidth = GetPropertyValue(swCustPropMgr, "Width")
    If propWidth = "" Then propWidth = GetPropertyValue(swCustPropMgr, "WIDTH")
    
    propLength = GetPropertyValue(swCustPropMgr, "Length")
    If propLength = "" Then propLength = GetPropertyValue(swCustPropMgr, "LENGTH")
    
    propHeight = GetPropertyValue(swCustPropMgr, "Height")
    If propHeight = "" Then propHeight = GetPropertyValue(swCustPropMgr, "HEIGHT")
    
    ' Try to get thickness - prioritize sheet metal thickness
    Dim thicknessVal As String
    If sheetMetalThickness <> "" And IsNumeric(sheetMetalThickness) Then
        thicknessVal = sheetMetalThickness
    Else
        thicknessVal = ExtractThickness(swCustPropMgr, itemName, swPlateBody)
    End If
    
    If thicknessVal <> "" And propWidth <> "" And propLength <> "" Then
        If IsNumeric(thicknessVal) And IsNumeric(propWidth) And IsNumeric(propLength) Then
            dimensions(0) = CDbl(thicknessVal)
            dimensions(1) = CDbl(propWidth)
            dimensions(2) = CDbl(propLength)
            ExtractAllDimensions = dimensions
            Exit Function
        End If
    End If
    
    ' Attempt 6: Final fallback - extract from body geometry directly (for non-cutlist parts)
    If Not swPlateBody Is Nothing Then
        Dim finalBodyDims As Variant
        finalBodyDims = ExtractAllDimensionsFromBody(swPlateBody)
        
        If Not IsEmpty(finalBodyDims) Then
            ExtractAllDimensions = finalBodyDims
            Exit Function
        End If
    End If
    
    ExtractAllDimensions = Empty
End Function

Sub SortDimensions(ByRef rawDims() As Double, ByRef sortedDims() As Double)
    ' Sort dimensions so that sortedDims(0) = thickness (smallest),
    ' sortedDims(1) = width (middle), sortedDims(2) = length (largest)
    
    ' Simple bubble sort for 3 elements
    Dim tempDims(2) As Double
    Dim i As Integer
    For i = 0 To 2
        tempDims(i) = rawDims(i)
    Next i
    
    ' Sort ascending
    If tempDims(0) > tempDims(1) Then
        Dim temp As Double
        temp = tempDims(0): tempDims(0) = tempDims(1): tempDims(1) = temp
    End If
    If tempDims(1) > tempDims(2) Then
        temp = tempDims(1): tempDims(1) = tempDims(2): tempDims(2) = temp
    End If
    If tempDims(0) > tempDims(1) Then
        temp = tempDims(0): tempDims(0) = tempDims(1): tempDims(1) = temp
    End If
    
    ' Assign to output array: smallest=thickness, middle=width, largest=length
    For i = 0 To 2
        sortedDims(i) = tempDims(i)
    Next i
End Sub

Function ParseAllDimensionsFromStockSize(ByVal stockSize As String) As Variant
    ' Parse StockSize to extract all dimensions if possible
    ' Returns array [thickness, width, length] or Empty if not enough dimensions
    
    stockSize = Replace(stockSize, """", "")
    stockSize = Trim(stockSize)
    
    Dim parts() As String
    Dim delimiter As String
    
    ' Find delimiter
    If InStr(stockSize, " X ") > 0 Then
        delimiter = " X "
    ElseIf InStr(stockSize, "x") > 0 Then
        delimiter = "x"
    ElseIf InStr(stockSize, " x ") > 0 Then
        delimiter = " x "
    ElseIf InStr(stockSize, "*") > 0 Then
        delimiter = "*"
    Else
        ParseAllDimensionsFromStockSize = Empty
        Exit Function
    End If
    
    parts = Split(stockSize, delimiter)
    
    ' Need at least 2 dimensions
    If UBound(parts) >= 1 Then
        Dim dimensions(2) As Double
        Dim numericParts As Integer
        numericParts = 0
        
        Dim i As Integer
        For i = 0 To UBound(parts)
            If i > 2 Then Exit For ' Only take first 3 dimensions
            Dim part As String
            part = Trim(parts(i))
            If IsNumeric(part) Then
                dimensions(numericParts) = CDbl(part)
                numericParts = numericParts + 1
            End If
        Next i
        
        If numericParts >= 2 Then
            ' Sort dimensions: smallest first (thickness)
            If numericParts = 2 Then
                ' Only have 2 dimensions, assume third is 1 or use width as length
                If dimensions(0) > dimensions(1) Then
                    Dim temp As Double
                    temp = dimensions(0): dimensions(0) = dimensions(1): dimensions(1) = temp
                End If
                dimensions(2) = dimensions(1) ' Use width as length for now
            ElseIf numericParts = 3 Then
                ' Have all 3 dimensions, sort them
                Dim rawDims(2) As Double
                rawDims(0) = dimensions(0)
                rawDims(1) = dimensions(1)
                rawDims(2) = dimensions(2)
                Call SortDimensions(rawDims, dimensions)
            End If
            
            ParseAllDimensionsFromStockSize = dimensions
        Else
            ParseAllDimensionsFromStockSize = Empty
        End If
    Else
        ParseAllDimensionsFromStockSize = Empty
    End If
End Function

Function CalculateMtlUnitQty(ByVal allDimensions As Variant, ByVal itemName As String) As String
    ' Calculate Mtl Unit Qty as length × height in SQUARE INCHES (excluding thickness)
    ' This matches the ProcessPLorCP approach: materialQuantity = length * height
    ' allDimensions should be [thickness, width, length] in document units
    
    If IsEmpty(allDimensions) Then
        CalculateMtlUnitQty = ""
        Exit Function
    End If
    
    Dim thickness As Double, width As Double, length As Double
    thickness = allDimensions(0)  ' Smallest dimension (thickness)
    width = allDimensions(1)      ' Middle dimension (width/height)
    length = allDimensions(2)     ' Largest dimension (length)
    
    If width <= 0 Or length <= 0 Then
        CalculateMtlUnitQty = ""
        Exit Function
    End If
    
    ' Convert dimensions to inches
    Dim widthInches As Double, lengthInches As Double
    widthInches = ConvertDimensionToInches(width, itemName & " width")
    lengthInches = ConvertDimensionToInches(length, itemName & " length")
    
    If widthInches <= 0 Or lengthInches <= 0 Then
        CalculateMtlUnitQty = ""
        Exit Function
    End If
    
    ' Calculate area as length × height (width is height in this context)
    ' Following ProcessPLorCP: materialQuantity = length * height
    Dim materialQuantity As Double
    materialQuantity = lengthInches * widthInches
    
    ' Format exactly like ProcessPLorCP: RemoveTrailingZeros(Format(materialQuantity, "0.0000"))
    CalculateMtlUnitQty = RemoveTrailingZeros(Format(materialQuantity, "0.0000"))
    
    Debug.Print "Material Quantity (length x height): " & lengthInches & " x " & widthInches & " = " & materialQuantity
End Function



' Add the RemoveTrailingZeros function if it doesn't exist in the second macro
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

Function ConvertDimensionToInches(ByVal dimensionValue As Double, ByVal debugLabel As String) As Double
    ' Convert a dimension from document units to inches
    
    If dimensionValue <= 0 Then
        ConvertDimensionToInches = 0
        Exit Function
    End If
    
    ' Get document units
    Dim docUnits As String
    docUnits = GetDocumentUnits()
    
    Dim convertedDimension As Double
    
    Select Case UCase(docUnits)
        Case "INCHES", "IN", "INCH"
            convertedDimension = dimensionValue
            
        Case "MILLIMETERS", "MM", "MILLIMETER"
            convertedDimension = dimensionValue / 25.4
            
        Case "CENTIMETERS", "CM", "CENTIMETER"
            convertedDimension = dimensionValue / 2.54
            
        Case "METERS", "M", "METER"
            convertedDimension = dimensionValue * 39.3701
            
        Case "MICROMETERS", "UM", "MICROMETER"
            convertedDimension = dimensionValue / 25400
            
        Case "FEET", "FT", "FOOT"
            convertedDimension = dimensionValue * 12
            
        Case Else
            convertedDimension = dimensionValue
    End Select
    
    ConvertDimensionToInches = convertedDimension
End Function

Sub SetCustomProperty(ByVal swCustPropMgr As SldWorks.CustomPropertyManager, ByVal propName As String, ByVal propValue As String)
    Debug.Print "*** SETTING PROPERTY: " & propName & " = " & propValue
    
    ' Set custom properties with proper error handling
    Dim propType As Long
    Dim addOption As Long
    Dim addResult As Long
    Dim result As Boolean
    
    propType = swCustomInfoText
    addOption = swCustomPropertyDeleteAndAdd
    
    Dim valOut As String, resolvedOut As String
    Dim wasResolved As Boolean, linkToProp As Boolean
    Dim getResult As Long
    
    getResult = swCustPropMgr.Get6(propName, False, valOut, resolvedOut, wasResolved, linkToProp)
    Debug.Print "Property " & propName & " current value: " & resolvedOut & " (get result: " & getResult & ")"
    
    If getResult <> 0 Then
        ' Property doesn't exist, add it
        Debug.Print "Property doesn't exist, adding..."
        addResult = swCustPropMgr.Add3(propName, propType, propValue, 0)
        Debug.Print "Add3 result: " & addResult
        If addResult <> 0 Then
            addResult = swCustPropMgr.Add3(propName, propType, propValue, addOption)
            Debug.Print "Add3 with delete option result: " & addResult
        End If
    Else
        ' Property exists, update it
        Debug.Print "Property exists, updating..."
        result = swCustPropMgr.Set2(propName, propValue)
        Debug.Print "Set2 result: " & result
        If Not result Then
            Debug.Print "Set2 failed, trying Add3 with delete..."
            addResult = swCustPropMgr.Add3(propName, propType, propValue, addOption)
            Debug.Print "Add3 with delete result: " & addResult
        End If
    End If
    
    ' Verify the property was set
    getResult = swCustPropMgr.Get6(propName, False, valOut, resolvedOut, wasResolved, linkToProp)
    Debug.Print "VERIFICATION - " & propName & " now equals: " & resolvedOut
End Sub









Function ExtractAllDimensionsFromBody(ByVal swBody As SldWorks.Body2) As Variant
    ' Extract dimensions directly from body geometry when no cutlist properties exist
    Debug.Print "Extracting dimensions from body geometry..."
    
    If swBody Is Nothing Then
        Debug.Print "ERROR: swBody is Nothing"
        ExtractAllDimensionsFromBody = Empty
        Exit Function
    End If
    
    Dim vBodyBox As Variant
    vBodyBox = swBody.GetBodyBox ' Returns in METERS
    
    If IsEmpty(vBodyBox) Or IsNull(vBodyBox) Then
        Debug.Print "ERROR: Could not get body bounding box"
        ExtractAllDimensionsFromBody = Empty
        Exit Function
    End If
    
    ' Calculate dimensions in meters
    Dim dimX_m As Double, dimY_m As Double, dimZ_m As Double
    dimX_m = Abs(vBodyBox(3) - vBodyBox(0))
    dimY_m = Abs(vBodyBox(4) - vBodyBox(1))
    dimZ_m = Abs(vBodyBox(5) - vBodyBox(2))
    
    Debug.Print "Raw body dimensions (meters): " & dimX_m & " x " & dimY_m & " x " & dimZ_m
    
    ' Convert to document units
    Dim docUnits As String
    docUnits = GetDocumentUnits()
    Debug.Print "Document units: " & docUnits
    
    Dim rawDims(2) As Double
    Select Case UCase(docUnits)
        Case "MILLIMETERS", "MM", "MILLIMETER"
            rawDims(0) = dimX_m * 1000
            rawDims(1) = dimY_m * 1000
            rawDims(2) = dimZ_m * 1000
        Case "INCHES", "IN", "INCH"
            rawDims(0) = dimX_m * 39.3701
            rawDims(1) = dimY_m * 39.3701
            rawDims(2) = dimZ_m * 39.3701
        Case Else
            rawDims(0) = dimX_m * 1000 ' Default to mm
            rawDims(1) = dimY_m * 1000
            rawDims(2) = dimZ_m * 1000
    End Select
    
    Debug.Print "Dimensions in document units: " & rawDims(0) & " x " & rawDims(1) & " x " & rawDims(2)
    
    ' Sort dimensions: thickness (smallest), width (middle), length (largest)
    Dim dimensions(2) As Double
    Call SortDimensions(rawDims, dimensions)
    
    Debug.Print "Sorted dimensions: T=" & dimensions(0) & " W=" & dimensions(1) & " L=" & dimensions(2)
    
    ExtractAllDimensionsFromBody = dimensions
End Function


