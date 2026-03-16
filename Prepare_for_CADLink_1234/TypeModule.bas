Attribute VB_Name = "TypeModule"
Option Explicit
Option Private Module

Private modelUOM As String
Private dimensionsChecked As Boolean

Public Sub ResetDimensionsChecked()
    dimensionsChecked = False
End Sub

Private Function GetMaterialOverridePreference(ByVal model As ModelDoc2) As String
    ' Check if the document has a material override preference set
    ' Returns: "USE_MODEL" or "USE_CUTLIST" or "" if not set
    On Error Resume Next
    Dim swCustPropMgr As SldWorks.customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")

    Dim valueOut As String
    Dim resolvedValueOut As String
    swCustPropMgr.Get4 "Material Override Preference", False, valueOut, resolvedValueOut

    GetMaterialOverridePreference = resolvedValueOut
    On Error GoTo 0
End Function

Private Sub SetMaterialOverridePreference(ByVal model As ModelDoc2, ByVal preference As String)
    ' Set the material override preference as a custom property
    ' preference should be "USE_MODEL" or "USE_CUTLIST"
    On Error Resume Next
    Dim swCustPropMgr As SldWorks.customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")

    swCustPropMgr.Add3 "Material Override Preference", swCustomInfoText, preference, swCustomPropertyOnlyIfNew
    swCustPropMgr.Set2 "Material Override Preference", preference
    On Error GoTo 0
End Sub






' Main function that extracts structural member type from a model
Sub ExtractStructuralMemberType(model As ModelDoc2, ByRef propertiesToSet As Object)
    
    
   DebugLog "=== ENTERING ExtractStructuralMemberType ==="
    DebugLog "Initial Reference Category check:"
    If propertiesToSet.exists("Reference Category") Then
        DebugLog "Reference Category exists: '" & propertiesToSet("Reference Category") & "'"
    Else
        DebugLog "Reference Category does NOT exist"
    End If
    
    
    
    ' Early exit if model is invalid
    If model Is Nothing Then
        Exit Sub
    End If
    
    ' Handle assembly models
    If model.GetType = swDocASSEMBLY Then
        propertiesToSet("Stock Size") = "ASMW"
        Exit Sub
    End If
    
    ' Get model unit of measure
    modelUOM = GetModelUnitOfMeasure(model)
    
    ' Initialize variables
    Dim swFeature As SldWorks.Feature
    Dim swCustPropMgr As SldWorks.customPropertyManager
    Dim valueOut As String
    Dim resolvedValueOut As String
    Dim typeFound As Boolean
    Dim bRet As Boolean
    Dim stockSize As String
    Dim refCategory As String
    Dim typeValue As String
    
    ' Add flag to track if Type="P"
    Dim isTypeP As Boolean
    isTypeP = False
    
    Dim swDimensions(2) As Double
    typeFound = False
    
    ' Get the valid reference categories from the PropertyRules module
    Dim validRefCategories As String
    validRefCategories = GetRequiredPartPropertyRules()("Reference Category")
    DebugLog "Valid ref categories pattern: " & validRefCategories
    
    'validRefCategories = Mid(validRefCategories, 2, Len(validRefCategories) - 2)  ' Remove parentheses
    
    DebugLog "After removing parentheses: " & validRefCategories
    
    
    
' Extract the content between ^( and )$ more reliably
Dim startPos As Integer, endPos As Integer
startPos = InStr(validRefCategories, "^(")
endPos = InStrRev(validRefCategories, ")$")

If startPos > 0 And endPos > startPos Then
    ' Extract just the part between the parentheses
    validRefCategories = Mid(validRefCategories, startPos + 2, endPos - startPos - 2)
    DebugLog "Extracted categories: " & validRefCategories
Else
    DebugLog "ERROR: Could not parse regex pattern"
    ' Fallback - assume it's already in the correct format
End If



    Dim validTypes As Variant
    validTypes = Split(validRefCategories, "|")
    DebugLog "Valid types array created with " & (UBound(validTypes) + 1) & " items"
    
    ' Check if Reference Category already exists
    'Dim refCategory As String
    refCategory = ""
    
    If propertiesToSet.exists("Reference Category") Then
        refCategory = propertiesToSet("Reference Category")
        DebugLog "Found existing Reference Category: '" & refCategory & "'"
        
        If refCategory <> "" Then
            ' Test the IsValidType function
            Dim isValid As Boolean
            isValid = IsValidType(refCategory, validTypes)
            DebugLog "IsValidType result for '" & refCategory & "': " & isValid
            
            If Not isValid Then
                DebugLog "*** CLEARING Reference Category because it's invalid ***"
                refCategory = ""
                propertiesToSet("Reference Category") = ""
            Else
                DebugLog "Reference Category is valid - keeping it"

                ' If Reference Category is PUR, process it and exit early
                If UCase(Trim(refCategory)) = "PUR" Then
                    DebugLog "*** EARLY EXIT PATH: Reference Category is PUR - calling ProcessPURShape then exiting ***"
                    Call ProcessPURShape(model, refCategory, isTypeP, modelUOM, propertiesToSet)
                    Exit Sub
                End If
            End If
        Else
            DebugLog "Reference Category is empty string"
        End If
    Else
        DebugLog "Reference Category property does not exist"
    End If
    
    DebugLog "Final refCategory value: '" & refCategory & "'"
    DebugLog "About to enter main processing logic..."
    
    ' Look for the CutListFolder for parts
    Set swFeature = model.FirstFeature
    
    Do While Not swFeature Is Nothing And Not typeFound
        If swFeature.GetTypeName2 = "CutListFolder" Then
            ' Get length from cutlist first
            Set swCustPropMgr = swFeature.customPropertyManager
            Dim cutlistLength As String
            swCustPropMgr.Get4 "Length", False, valueOut, cutlistLength
            
            ' Check if total length is 0 or empty
            Dim totalLength As String
            totalLength = GetPropertyValue(propertiesToSet, "LengthA")
            
            If (totalLength = "" Or totalLength = "0" Or CDbl(totalLength) = 0) And cutlistLength <> "" Then
                ' Use the new utility function for conversion
                Dim finalLengthStr As String
                finalLengthStr = ConvertCutlistLengthToInches(cutlistLength, modelUOM)
                
                propertiesToSet("LengthA") = finalLengthStr
                propertiesToSet("Length") = finalLengthStr

                ' Update Mtl Unit Qty based on the (potentially converted) reference category
                If propertiesToSet.exists("Reference Category") Then
                    If propertiesToSet("Reference Category") <> "PL" And propertiesToSet("Reference Category") <> "CP" Then
                        propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
                    End If
                Else
                    propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
                End If
            End If
            
            ' Validate that this is a proper cutlist folder with bodies
            Dim swBodyFolder As BodyFolder
            Set swBodyFolder = swFeature.GetSpecificFeature2
            
            If Not swBodyFolder Is Nothing Then
                Dim vBodies As Variant
                vBodies = swBodyFolder.GetBodies
                
                If Not IsEmpty(vBodies) Then
                    Set swCustPropMgr = swFeature.customPropertyManager
                    
                    ' Log all cutlist properties
                    LogCutlistProperties swCustPropMgr, swFeature.Name
                    
                    ' Check for structural member type
                    bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                    
                    If bRet And resolvedValueOut <> "" Then
                        ' Set Reference Category property
                        refCategory = resolvedValueOut
                        propertiesToSet("Reference Category") = refCategory
                        
                        ' Check if user has already set a Type value manually
                        If propertiesToSet.exists("Type") Then
                            If UCase(Trim(propertiesToSet("Type"))) = "P" Then
                                ' Keep the Type="P" setting the user has already set
                                isTypeP = True
                            Else
                                ' Otherwise default to Type="M"
                                propertiesToSet("Type") = "M"
                            End If
                        Else
                            ' No Type value exists yet, default to "M"
                            propertiesToSet("Type") = "M"
                        End If
                        
' Get StockSize for structural members
swCustPropMgr.Get4 "StockSize", False, valueOut, resolvedValueOut

If resolvedValueOut <> "" Then
    stockSize = resolvedValueOut
    ProcessStructuralMemberDimensions stockSize, modelUOM, propertiesToSet
Else
    stockSize = "UNKNOWN SIZE"
    propertiesToSet("Stock Size") = stockSize
End If

' *** NEW: Apply material based on Grade from cutlist ***
swCustPropMgr.Get4 "Grade", False, valueOut, resolvedValueOut
If resolvedValueOut <> "" Then
    DebugLog "Found Grade in cutlist: " & resolvedValueOut
    Call ApplyMaterialFromGrade(model, resolvedValueOut)
End If
                        
                        ' Handle Mtl Part Number for painted parts
                        'If isTypeP Then
                        '    If propertiesToSet.Exists("Mtl Part Number") Then
                        '        propertiesToSet.Remove "Mtl Part Number"
                        '    End If
                        'End If
                        
                        typeFound = True
                        ' Exclude the item from the cutlist
                        Call ExcludeFromCutlist(model, swFeature.Name)
                    End If
                End If
            End If
        End If
        
        If Not typeFound Then
            Set swFeature = swFeature.GetNextFeature
        End If
    Loop
    
    ' If not a structural member, process as a regular part
  '  If Not typeFound Then
  '      ProcessNonStructuralMember refCategory, propertiesToSet
  '  End If
    
  ' If not a structural member, process as a regular part
If Not typeFound Then
    ' Only process as non-structural member if we don't already have a valid reference category
    If refCategory = "" Or Not IsValidType(refCategory, validTypes) Then
        ProcessNonStructuralMember refCategory, propertiesToSet
    End If
End If
    
' Update modelUOM if dimensions were converted to inches in GetDimensions
If propertiesToSet.exists("Dimensional UOM") Then
    Dim updatedUOM As String
    updatedUOM = propertiesToSet("Dimensional UOM")
    If updatedUOM <> modelUOM Then
        DebugLog "=== UOM UPDATE ==="
        DebugLog "Original modelUOM: " & modelUOM
        DebugLog "Updated to: " & updatedUOM
        modelUOM = updatedUOM
        DebugLog "=================="
    End If
End If
'Stop
' Process the reference category (existing or newly set)
If refCategory <> "" Then
    Select Case refCategory
            Case "PL", "CP"
                ProcessPLorCP model, refCategory, isTypeP, modelUOM, propertiesToSet
                
            Case "FB"
                ProcessFB model, refCategory, isTypeP, typeFound, propertiesToSet
                
            Case "L"
                ProcessLShape model, stockSize, isTypeP, modelUOM, propertiesToSet
                
           Case "C", "Z"
    ProcessCShape model, refCategory, stockSize, isTypeP, modelUOM, propertiesToSet
            
            Case "W", "I"
    ProcessWShape model, refCategory, stockSize, isTypeP, propertiesToSet
                
            Case "HSS"
    ProcessHSSShape model, "HSS", stockSize, isTypeP, modelUOM, propertiesToSet, swDimensions

Case "MEC", "AS"
    ProcessMECASShape model, refCategory, stockSize, isTypeP, modelUOM, propertiesToSet, swDimensions
                
            Case "RB", "PIN"
                ProcessRBShape model, isTypeP, propertiesToSet

            Case "SQ"
                ProcessSQShape model, isTypeP, propertiesToSet

            Case "PUR"
                Call ProcessPURShape(model, refCategory, isTypeP, modelUOM, propertiesToSet)
            
            
           Case "EXP"
    ProcessEXP model, refCategory, modelUOM, propertiesToSet
            
            
            Case "WD"
                    ProcessWDShape refCategory, isTypeP, propertiesToSet
     Case "UB"
    ProcessUBShape model, refCategory, stockSize, isTypeP, modelUOM, propertiesToSet
            
           
            Case "SHIM"
                Call ProcessSHIMShape(model, refCategory, isTypeP, modelUOM, propertiesToSet)
          
            Case "S"
                ' Handle S shapes - get original Mtl Part Number directly from cutlist
                ProcessSShape model, stockSize, isTypeP, modelUOM, propertiesToSet
                
            Case "MC", "WT", "S", "T", "HP", "MT", "ST", "WWF"

                ProcessMiscShapes refCategory, stockSize, isTypeP, modelUOM, propertiesToSet
                
            Case "PI"
                ProcessPIShape model, stockSize, isTypeP, propertiesToSet
                
             Case "KS"
                ProcessKS model, "KS", isTypeP, modelUOM, propertiesToSet
                
                
         Case "HOSE"
    DebugLog "=== ENTERED HOSE CASE ==="
    ProcessHOSEShape model, refCategory, isTypeP, modelUOM, propertiesToSet
    DebugLog "=== ProcessHOSEShape COMPLETED - NOW CHECKING GEOMETRY ==="
    
    DebugLog "=== Checking if geometry should be created ==="
    'If HoseGeometryModule.ShouldCreateHoseGeometry() Then
        DebugLog "=== Calling HoseGeometryModule.CreateHoseGeometry ==="
        Call HoseGeometryModule.CreateHoseGeometry(model, propertiesToSet)
        DebugLog "=== CreateHoseGeometry COMPLETED ==="
      '  Stop
   ' Else
       ' DebugLog "=== ShouldCreateHoseGeometry returned False ==="
   ' End If
    Call ExcludeAllCutlistFolders(model)
    DebugLog "=== HOSE CASE COMPLETED ==="
                
            Case Else
                If IsValidType(typeValue, validTypes) Then
                    ProcessOtherValidTypes typeValue, stockSize, propertiesToSet
                Else
                    stockSize = "UNKNOWN TYPE"
                    propertiesToSet("Stock Size") = stockSize
                End If
        End Select
    Else
        stockSize = "UNKNOWN REFERENCE CATEGORY"
        propertiesToSet("Stock Size") = stockSize
    End If
    
    ' Set Mtl Unit Qty to LengthA globally EXCEPT for PL
    SetMtlUnitQtyFromLengthA propertiesToSet
    

    
    ' Set MtlUOM based on part type
    SetMtlUOMBasedOnType propertiesToSet

    ' *** FINAL CLEANUP: Remove Mtl Part Number and Mtl Unit Qty for all P types ***
    If propertiesToSet.exists("Type") Then
        If UCase(Trim(propertiesToSet("Type"))) = "P" Then
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "Cleanup: Removed Mtl Part Number (Type P)"
            End If
            If propertiesToSet.exists("Mtl Unit Qty") Then
                propertiesToSet.Remove "Mtl Unit Qty"
                DebugLog "Cleanup: Removed Mtl Unit Qty (Type P)"
            End If
        End If
    End If

     Call DeleteWeldment_OnlyIfNoMembers(model)
End Sub

' Process the PL or CP reference categories
Sub ProcessPLorCP(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "#########################"
    DebugLog "### ENTERING ProcessPLorCP FUNCTION ###"
    DebugLog "#########################"
    DebugLog "refCategory: " & refCategory
    DebugLog "isTypeP (initial): " & isTypeP
    DebugLog "modelUOM: " & modelUOM
    
    Dim widthStr As String, heightStr As String, lengthStr As String
    Dim width As Double, height As Double, length As Double
    Dim PLthickness As String
    Dim PLthicknessDecimal As Double
    Dim PLMaterial As String
   
    ' Get dimensions from properties
    DebugLog "=== GETTING DIMENSIONS FROM PROPERTIES ==="
    widthStr = GetPropertyValue(propertiesToSet, "Width")
    heightStr = GetPropertyValue(propertiesToSet, "Height")
    lengthStr = GetPropertyValue(propertiesToSet, "Length")
    
    DebugLog "Raw property values:"
    DebugLog "widthStr from properties: '" & widthStr & "'"
    DebugLog "heightStr from properties: '" & heightStr & "'"
    DebugLog "lengthStr from properties: '" & lengthStr & "'"
    DebugLog "modelUOM: " & modelUOM
    
    ' Get the UPDATED Dimensional UOM from properties (AddDimensionsProperties may have changed it)
    Dim actualUOM As String
    If propertiesToSet.exists("Dimensional UOM") Then
        actualUOM = propertiesToSet("Dimensional UOM")
        DebugLog "Using updated Dimensional UOM from properties: " & actualUOM
    Else
        actualUOM = modelUOM
        DebugLog "Using original model UOM: " & actualUOM
    End If
    
    ' Convert strings to numbers
    DebugLog "=== CONVERTING TO NUMBERS ==="
    width = FractionToDecimal(widthStr)
    height = FractionToDecimal(heightStr)
    length = FractionToDecimal(lengthStr)
    
    DebugLog "After FractionToDecimal conversion:"
    DebugLog "width: " & width
    DebugLog "height: " & height
    DebugLog "length: " & length
    
    ' *** CHECK FOR PATH-BASED UNIT OVERRIDE BEFORE CheckPlateDimensions ***
    DebugLog "=== CHECKING PATH-BASED UNIT OVERRIDE ==="
    Dim pathOverride As String
    pathOverride = GetPathBasedUnitOverride(model)
    
    If pathOverride <> "" Then
        DebugLog "=== PATH-BASED OVERRIDE ACTIVE ==="
        DebugLog "Override Unit: " & pathOverride
        DebugLog "Original actualUOM: " & actualUOM
        
        ' Apply path-based conversion if needed
        If pathOverride = "mm" And actualUOM <> "mm" Then
            DebugLog "*** CONVERTING TO MM PER PATH OVERRIDE ***"
            DebugLog "Before conversion: width=" & width & ", height=" & height & ", length=" & length & " (" & actualUOM & ")"
            
            ' Convert from current units to mm
            If actualUOM = "in" Then
                width = width * 25.4
                height = height * 25.4
                length = length * 25.4
            ElseIf actualUOM = "cm" Then
                width = width * 10
                height = height * 10
                length = length * 10
            ElseIf actualUOM = "m" Then
                width = width * 1000
                height = height * 1000
                length = length * 1000
            ElseIf actualUOM = "ft" Then
                width = width * 304.8
                height = height * 304.8
                length = length * 304.8
            End If
            
            actualUOM = "mm"
            propertiesToSet("Dimensional UOM") = "mm"
            
            ' Reformat strings as metric
            widthStr = FormatMetricDimension(width)
            heightStr = FormatMetricDimension(height)
            lengthStr = FormatMetricDimension(length)
            
            DebugLog "After MM conversion: width=" & width & ", height=" & height & ", length=" & length & " (mm)"
            
        ElseIf pathOverride = "in" And actualUOM <> "in" Then
            DebugLog "*** CONVERTING TO INCHES PER PATH OVERRIDE ***"
            DebugLog "Before conversion: width=" & width & ", height=" & height & ", length=" & length & " (" & actualUOM & ")"
            
            ' Convert from current units to inches
            width = ConvertModelUnitsToInches(width, actualUOM)
            height = ConvertModelUnitsToInches(height, actualUOM)
            length = ConvertModelUnitsToInches(length, actualUOM)
            
            actualUOM = "in"
            propertiesToSet("Dimensional UOM") = "in"
            
            ' Reformat strings as fractions
            widthStr = ConvertDecimalToFraction(width)
            heightStr = ConvertDecimalToFraction(height)
            lengthStr = ConvertDecimalToFraction(length)
            
            DebugLog "After INCH conversion: width=" & width & ", height=" & height & ", length=" & length & " (in)"
        Else
            DebugLog "Path override matches current units - no conversion needed"
        End If
        
        DebugLog "Final actualUOM after path override: " & actualUOM
        DebugLog "=== END PATH-BASED OVERRIDE ==="
    Else
        DebugLog "No path-based override - proceeding with standard logic"
    End If
    
    ' *** Run CheckPlateDimensions ONLY if no path override ***
    DebugLog "=== CALLING CheckPlateDimensions ==="
    Dim checkResult As Boolean
    If pathOverride = "" Then
        checkResult = CheckPlateDimensions(model, width, height, length, actualUOM, propertiesToSet)
        DebugLog "CheckPlateDimensions returned: " & checkResult
    Else
        DebugLog "SKIPPING CheckPlateDimensions - path override is controlling units"
        checkResult = True  ' Assume valid since path override is active
    End If

    ' *** Check if dimensions match flat bar specifications ***
    ' Check for flat bar if it's a PL part (regardless of whether structural member was found)
    If refCategory = "PL" Then
        DebugLog "=== CALLING CheckAndConvertToFlatBar ==="
        Dim flatBarResult As Boolean
        flatBarResult = FlatBarModule.CheckAndConvertToFlatBar(model, width, height, length, actualUOM, propertiesToSet)
        If flatBarResult Then
            DebugLog "Part converted to FB (Flat Bar)"
            ' Update Reference Category and Type to FB in both variable and dictionary
            refCategory = "FB"
            propertiesToSet("Reference Category") = "FB"
            propertiesToSet("Type") = "FB"
            DebugLog "Updated refCategory and Type to: FB"
        Else
            DebugLog "Part does not match flat bar specifications"
        End If
    Else
        DebugLog "SKIPPING flat bar check - not a PL part (refCategory = " & refCategory & ")"
    End If
'Stop
    ' *** Update isTypeP flag based on Type property ***
    Dim finalType As String
    If propertiesToSet.exists("Type") Then
        finalType = UCase(Trim(propertiesToSet("Type")))
        DebugLog "Type after checks: " & finalType
        
        If finalType = "P" Then
            isTypeP = True
            DebugLog "Updated isTypeP to True"
        Else
            isTypeP = False
            DebugLog "Updated isTypeP to False"
        End If
    End If
    
    ' *** Get current Type ***
    DebugLog "=== CHECKING FINAL TYPE PROPERTY ==="
    Dim currentType As String
    If propertiesToSet.exists("Type") Then
        currentType = UCase(Trim(propertiesToSet("Type")))
    Else
        currentType = ""
    End If
    DebugLog "currentType: '" & currentType & "'"
    
    ' If Type is P, handle as purchase part
    If currentType = "P" Or isTypeP Then
        DebugLog "=== TYPE IS P - PURCHASE PART PROCESSING ==="
        If propertiesToSet.exists("Mtl Part Number") Then
            propertiesToSet.Remove "Mtl Part Number"
            DebugLog "Removed Mtl Part Number"
        End If
        If propertiesToSet.exists("Mtl Unit Qty") Then
            propertiesToSet.Remove "Mtl Unit Qty"
            DebugLog "Removed Mtl Unit Qty"
        End If

        ' Ensure width is the smaller dimension for stock size
        DebugLog "Before dimension swap - width: " & width & ", height: " & height
        Dim temp As Double
        If width > height Then
            temp = width
            width = height
            height = temp
            DebugLog "Swapped dimensions - width: " & width & ", height: " & height
        End If

        ' Set just the Stock Size for Purchase parts
        If actualUOM = "in" Then
            widthStr = ConvertDecimalToFraction(width)
            heightStr = ConvertDecimalToFraction(height)
        Else
            widthStr = FormatMetricDimension(width)
            heightStr = FormatMetricDimension(height)
        End If

        Dim stockSize As String
        stockSize = refCategory & " " & widthStr & " x " & heightStr
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Set Stock Size for Purchase part: " & stockSize
        DebugLog "=== EXITING ProcessPLorCP - TYPE P ==="
        Exit Sub
    End If

    ' *** CONTINUE WITH MANUFACTURING LOGIC ONLY IF TYPE M ***
    DebugLog "=== TYPE IS M - MANUFACTURING PART PROCESSING ==="
    
    ' Calculate material quantity (area = length * height) BEFORE dimension swapping
    Dim materialQuantity As Double
    materialQuantity = length * height
    propertiesToSet("Mtl Unit Qty") = RemoveTrailingZeros(Format(materialQuantity, "0.0000"))
    DebugLog "Material Quantity (length x height): " & length & " x " & height & " = " & materialQuantity
    DebugLog "Set Mtl Unit Qty: " & propertiesToSet("Mtl Unit Qty")
    
    ' Ensure width is the smaller dimension for stock size
    DebugLog "=== DIMENSION SWAPPING FOR STOCK SIZE ==="
    DebugLog "Before swap - width: " & width & ", height: " & height
    If width > height Then
        temp = width
        width = height
        height = temp
        DebugLog "Swapped dimensions - width: " & width & ", height: " & height
    Else
        DebugLog "No swap needed - width is already smaller"
    End If
    
    ' --- Update Stock Size to Use the Rounded-Up Width ---
    Dim roundedWidth As Double
    roundedWidth = width
    DebugLog "roundedWidth: " & roundedWidth
    
    ' IMPORTANT FIX: Convert the dimensions to appropriate format based on actualUOM
    DebugLog "=== CONVERTING TO APPROPRIATE FORMAT ==="
    DebugLog "Current actualUOM: " & actualUOM
    
    If actualUOM = "in" Then
        ' Imperial - use fractions
        widthStr = ConvertDecimalToFraction(width)
        heightStr = ConvertDecimalToFraction(height)
        DebugLog "Using imperial fractions:"
    Else
        ' Metric - use FormatMetricDimension
        widthStr = FormatMetricDimension(width)
        heightStr = FormatMetricDimension(height)
        DebugLog "Using metric format:"
    End If
    
    DebugLog "widthStr: " & widthStr
    DebugLog "heightStr: " & heightStr

    ' Check if this is a flat bar
    If refCategory = "FB" Then
        ' Check if FlatBarSize was set (from flat bar matching)
        If propertiesToSet.exists("FlatBarSize") Then
            stockSize = "FB " & propertiesToSet("FlatBarSize")
            DebugLog "Using matched Flat Bar size: " & stockSize
        Else
            ' Use decimal format for consistency with ProcessFB: "FB 0.25 x 1.5"
            Dim fbThickness As String, fbWidth As String
            fbThickness = RemoveTrailingZeros(Format(width, "0.######"))
            fbWidth = RemoveTrailingZeros(Format(height, "0.######"))
            stockSize = "FB " & fbThickness & " x " & fbWidth
            DebugLog "Using Flat Bar decimal format: " & stockSize
        End If
    Else
        stockSize = refCategory & " " & widthStr & " x " & heightStr
        DebugLog "Initial Stock Size: " & stockSize
    End If
    
    ' --- Material Handling Section ---
    DebugLog "=== MATERIAL HANDLING ==="
    Dim materialFound As Boolean
    materialFound = False
    PLMaterial = ""
    
    ' First try getting from properties
    If propertiesToSet.exists("Material") Then
        PLMaterial = propertiesToSet("Material")
        DebugLog "Material from 'Material' property: " & PLMaterial
    End If
    
    ' If empty, try SW-Material
    If PLMaterial = "" Then
        If propertiesToSet.exists("SW-Material") Then
            PLMaterial = propertiesToSet("SW-Material")
            DebugLog "Material from 'SW-Material' property: " & PLMaterial
        End If
    End If
    
    ' If still empty, try getting from material feature
    If PLMaterial = "" Then
        DebugLog "Searching for MaterialFolder feature..."
        Dim swFeat As Feature
        Set swFeat = model.FirstFeature
        
        Do While Not swFeat Is Nothing
            If swFeat.GetTypeName2 = "MaterialFolder" Then
                PLMaterial = swFeat.Name
                materialFound = True
                DebugLog "Found MaterialFolder: " & PLMaterial
                Exit Do
            End If
            Set swFeat = swFeat.GetNextFeature
        Loop
    End If
    
    DebugLog "Final PLMaterial: " & PLMaterial
    
    Dim PLMtlNumb As String
    If PLMaterial <> "" Then
        DebugLog "=== MATERIAL PROCESSING ==="
        ' Remove quotes if they exist
        PLMaterial = Replace(PLMaterial, """", "")
        DebugLog "After removing quotes: " & PLMaterial
        
        ' Clean up the SW-Material@ prefix if it exists
        If InStr(PLMaterial, "SW-Material@") > 0 Then
            PLMaterial = Mid(PLMaterial, InStr(PLMaterial, "@") + 1)
            DebugLog "After removing SW-Material@ prefix: " & PLMaterial
        End If
        
        ' Remove .SLDPRT if it exists
        If InStr(PLMaterial, ".SLDPRT") > 0 Then
            PLMaterial = Left(PLMaterial, InStr(PLMaterial, ".SLDPRT") - 1)
            DebugLog "After removing .SLDPRT: " & PLMaterial
        End If
        
        Dim stockThickness As Double
        stockThickness = ConvertFractionToDecimal(widthStr)
        PLthicknessDecimal = stockThickness
        PLthickness = RemoveTrailingZeros(Format(PLthicknessDecimal, "0.0000"))
        
        DebugLog "stockThickness: " & stockThickness
        DebugLog "PLthickness: " & PLthickness

        If refCategory = "CP" Then
            PLMtlNumb = "CP-" & PLthickness & "-" & PLMaterial
        ElseIf refCategory = "FB" Then
            PLMtlNumb = "FB-" & PLthickness & "-" & PLMaterial
        Else
            PLMtlNumb = "PLT-" & PLthickness & "-" & PLMaterial
        End If
        DebugLog "Generated PLMtlNumb: " & PLMtlNumb
    Else
        DebugLog "=== NO MATERIAL FOUND - USING UNKNOWN ==="
        PLthickness = Format(width, "0.0000")
        PLMtlNumb = refCategory & "-" & PLthickness & "-UNKNOWN"
        DebugLog "Generated PLMtlNumb with UNKNOWN: " & PLMtlNumb
    End If
    
    ' Set the final properties for Manufacturing parts
    DebugLog "=== SETTING FINAL PROPERTIES ==="
    propertiesToSet("Mtl Part Number") = PLMtlNumb
    DebugLog "Set Mtl Part Number: " & PLMtlNumb
    
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size: " & stockSize

    ' Remove FlatBarSize property if it exists (internal use only, not a final property)
    If propertiesToSet.exists("FlatBarSize") Then
        propertiesToSet.Remove "FlatBarSize"
        DebugLog "Removed temporary FlatBarSize property"
    End If

    DebugLog "=== ProcessPLorCP COMPLETED ==="
    DebugLog "Final Results:"
    DebugLog "Stock Size: " & stockSize
    DebugLog "Mtl Part Number: " & PLMtlNumb
    DebugLog "Dimensional UOM: " & actualUOM
    DebugLog "#########################"
End Sub


Public Sub ProcessFB(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, ByVal typeFound As Boolean, ByRef propertiesToSet As Object)
    
    ' *** DEBUG: Show all properties in propertiesToSet ***
    DebugLog "=== ProcessFB - All Properties in propertiesToSet ==="
    Dim key As Variant
    For Each key In propertiesToSet.keys
        DebugLog "Property: " & key & " = " & propertiesToSet(key)
    Next key
    DebugLog "=== End Properties List ==="
    
    Dim stockSize As String

    ' Check if FlatBarSize was set from flat bar matching
    If propertiesToSet.exists("FlatBarSize") Then
        stockSize = refCategory & " " & propertiesToSet("FlatBarSize")
        propertiesToSet("Stock Size") = stockSize
        DebugLog "ProcessFB: Using matched Flat Bar size: " & stockSize
        ' Remove after using
        propertiesToSet.Remove "FlatBarSize"
    ElseIf Not typeFound Then
        ' --- PATH 1: Logic for non-structural members (use decimal format) ---
        Dim w As Double, h As Double, l As Double
        w = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
        h = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
        l = FractionToDecimal(GetPropertyValue(propertiesToSet, "Length"))

        Dim fbSortedDimensions(2) As Double
        fbSortedDimensions(0) = w
        fbSortedDimensions(1) = h
        fbSortedDimensions(2) = l
        Call BubbleSort(fbSortedDimensions)

        DebugLog "ProcessFB: Sorted dimensions: " & fbSortedDimensions(0) & ", " & fbSortedDimensions(1) & ", " & fbSortedDimensions(2)

        ' Try to match against flat bar inventory to determine correct stock size and cut length
        Dim materialName As String
        materialName = GetPropertyValue(propertiesToSet, "Material")

        Dim fbMatch As String
        Dim fbThickness As Double, fbWidth As Double, fbCutLength As Double

        ' Try smallest x middle
        fbMatch = FlatBarModule.GetFlatBarInfo(fbSortedDimensions(0), fbSortedDimensions(1), materialName)
        If fbMatch <> "Unknown material" And fbMatch <> "No matching flat bar" Then
            DebugLog "ProcessFB: Found FB match (thickness x middle): " & fbMatch
            fbThickness = fbSortedDimensions(0)
            fbWidth = fbSortedDimensions(1)
            fbCutLength = fbSortedDimensions(2)
            stockSize = refCategory & " " & fbMatch
        Else
            ' Try smallest x largest
            fbMatch = FlatBarModule.GetFlatBarInfo(fbSortedDimensions(0), fbSortedDimensions(2), materialName)
            If fbMatch <> "Unknown material" And fbMatch <> "No matching flat bar" Then
                DebugLog "ProcessFB: Found FB match (thickness x largest): " & fbMatch
                fbThickness = fbSortedDimensions(0)
                fbWidth = fbSortedDimensions(2)
                fbCutLength = fbSortedDimensions(1)
                stockSize = refCategory & " " & fbMatch
            Else
                ' No match - use decimal format with 2 smallest dimensions
                DebugLog "ProcessFB: No FB inventory match - using decimal format"
                fbThickness = fbSortedDimensions(0)
                fbWidth = fbSortedDimensions(1)
                fbCutLength = fbSortedDimensions(2)
                
                stockSize = refCategory & " " & RemoveTrailingZeros(Format(fbThickness, "0.######")) & " x " & RemoveTrailingZeros(Format(fbWidth, "0.######"))
            End If
        End If

        propertiesToSet("Stock Size") = stockSize
        propertiesToSet("Width") = ConvertToFraction(fbWidth, propertiesToSet)
        propertiesToSet("Height") = ConvertToFraction(fbThickness, propertiesToSet)
        propertiesToSet("Length") = ConvertToFraction(fbCutLength, propertiesToSet)

        DebugLog "ProcessFB: Final - Stock Size=" & stockSize & ", W=" & propertiesToSet("Width") & ", H=" & propertiesToSet("Height") & ", L=" & propertiesToSet("Length")

        ' Manual Mtl Part Number for non cutlist items
        If Not isTypeP Then
            Dim dim1 As String, dim2 As String
            dim1 = RemoveTrailingZeros(Format(fbThickness, "0.######"))
            dim2 = RemoveTrailingZeros(Format(fbWidth, "0.######"))

            Dim gradeVal As String
            gradeVal = Trim$(GetPropertyValue(propertiesToSet, "Grade"))
            If Len(gradeVal) = 0 Then gradeVal = Trim$(GetPropertyValue(propertiesToSet, "Material"))
            If Len(gradeVal) = 0 Then
                ' last resort pull model material
                On Error Resume Next
                Dim matIdName As String
                matIdName = model.GetMaterialPropertyName2("") ' returns "db|name" or just "name"
                On Error GoTo 0
                If Len(matIdName) > 0 Then
                    Dim barPos As Long
                    barPos = InStr(matIdName, "|")
                    If barPos > 0 Then
                        gradeVal = Mid$(matIdName, barPos + 1)
                    Else
                        gradeVal = matIdName
                    End If
                End If
            End If
            gradeVal = Replace(Trim$(gradeVal), " ", "")
            
            Dim pnNC As String
            pnNC = "FB-" & dim1 & "x" & dim2
            If Len(gradeVal) > 0 Then pnNC = pnNC & "-" & gradeVal
            propertiesToSet("Mtl Part Number") = pnNC
            DebugLog "ProcessFB: Created manual Mtl Part Number (non cutlist): " & pnNC
        Else
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "ProcessFB: Removed Mtl Part Number for Type P in non cutlist path"
            End If
        End If
        
    Else
        ' --- PATH 2: Logic for structural members from a cutlist ---
        Dim currentStockSize As String
        currentStockSize = GetPropertyValue(propertiesToSet, "Stock Size")
        
        If Len(currentStockSize) > 0 Then
            Dim originalModelUOM As String
            originalModelUOM = GetModelUnitOfMeasure(model)
            DebugLog "ProcessFB: Original model UOM: " & originalModelUOM
            DebugLog "ProcessFB: Current StockSize from cutlist: " & currentStockSize
            
            If IsMetricUnit(originalModelUOM) Then
                DebugLog "ProcessFB: Converting cutlist StockSize from metric (" & originalModelUOM & ") to inches"
                currentStockSize = Replace(currentStockSize, "FB ", "")
                currentStockSize = Trim$(currentStockSize)
                
                Dim dimensions() As String
                dimensions = Split(currentStockSize, "x")
                
                If UBound(dimensions) >= 1 Then
                    Dim i As Integer
                    For i = 0 To UBound(dimensions)
                        Dim tok As String
                        tok = Trim$(dimensions(i))
                        If IsNumeric(tok) Then
                            Dim dimValue As Double
                            dimValue = CDbl(tok)
                            dimValue = ConvertModelUnitsToInches(dimValue, originalModelUOM)
                            dimensions(i) = RemoveTrailingZeros(Format(dimValue, "0.#####"))
                        Else
                            dimensions(i) = tok
                        End If
                    Next i
                    
                    stockSize = "FB " & Join(dimensions, " x ")
                    propertiesToSet("Stock Size") = stockSize
                    DebugLog "ProcessFB: Converted StockSize: " & stockSize
                End If
            Else
                DebugLog "ProcessFB: Original model is imperial (" & originalModelUOM & "), no conversion needed"
                If Left$(currentStockSize, 3) <> "FB " Then
                    stockSize = "FB " & currentStockSize
                    propertiesToSet("Stock Size") = stockSize
                End If
            End If
        End If
        
        ' Try to read Mtl Part Number from cutlist first
        If Not isTypeP Then
            DebugLog "ProcessFB: Looking for Mtl Part Number in cutlist..."
            Dim mtlPartNumber As String
            mtlPartNumber = ""
            Dim foundInCutlist As Boolean
            foundInCutlist = False
            
            Dim swFeature As Feature
            Set swFeature = model.FirstFeature
            Do While Not swFeature Is Nothing
                If swFeature.GetTypeName2 = "CutListFolder" Then
                    Dim swCustPropMgr As customPropertyManager
                    Set swCustPropMgr = swFeature.customPropertyManager
                    If Not swCustPropMgr Is Nothing Then
                        Dim valueOut As String, resolvedValueOut As String
                        Dim bRet As Boolean
                        bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                        If bRet And resolvedValueOut = "FB" Then
                            bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, mtlPartNumber)
                            If bRet And Len(Trim$(mtlPartNumber)) > 0 Then
                                propertiesToSet("Mtl Part Number") = mtlPartNumber
                                foundInCutlist = True
                                DebugLog "ProcessFB: Added Mtl Part Number from cutlist: " & mtlPartNumber
                            End If
                            Exit Do
                        End If
                    End If
                End If
                Set swFeature = swFeature.GetNextFeature
            Loop
            
            ' Manual fallback if cutlist did not provide it
            If Not foundInCutlist Then
                DebugLog "ProcessFB: No cutlist Mtl Part Number found, creating manually..."
                
                Dim finalStockSize As String
                finalStockSize = GetPropertyValue(propertiesToSet, "Stock Size")
                
                If finalStockSize = "" Then
                    Dim wManual As Double, hManual As Double
                    wManual = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
                    hManual = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
                    
                    Dim manualSortedDims(1) As Double
                    manualSortedDims(0) = wManual
                    manualSortedDims(1) = hManual
                    Call BubbleSort(manualSortedDims)
                    
                    Dim manualDim1 As String, manualDim2 As String
                    manualDim1 = RemoveTrailingZeros(Format(manualSortedDims(0), "0.######"))
                    manualDim2 = RemoveTrailingZeros(Format(manualSortedDims(1), "0.######"))
                    
                    finalStockSize = "FB " & manualDim1 & " x " & manualDim2
                    propertiesToSet("Stock Size") = finalStockSize
                End If
                
                ' Parse back into thickness and width
                Dim cleanStock As String, dims() As String
                cleanStock = finalStockSize
                If Left$(cleanStock, 3) = "FB " Then cleanStock = Mid$(cleanStock, 4)
                dims = Split(cleanStock, "x")
                
                Dim t1 As String, t2 As String
                If UBound(dims) >= 1 Then
                    t1 = Replace(Trim$(dims(0)), " ", "")
                    t2 = Replace(Trim$(dims(1)), " ", "")
                Else
                    ' Fallback to manual dimensions if parsing fails
                    Dim wVal As Double, hVal As Double
                    wVal = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
                    hVal = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
                    
                    If wVal > 0 And hVal > 0 Then
                        Dim p2Sorted(1) As Double
                        p2Sorted(0) = wVal
                        p2Sorted(1) = hVal
                        Call BubbleSort(p2Sorted)
                        t1 = RemoveTrailingZeros(Format(p2Sorted(0), "0.######"))
                        t2 = RemoveTrailingZeros(Format(p2Sorted(1), "0.######"))
                    Else
                        ' Absolute fallback
                        t1 = "ERR"
                        t2 = "ERR"
                    End If
                End If
                
                Dim gradeCL As String
                gradeCL = Trim$(GetPropertyValue(propertiesToSet, "Grade"))
                If Len(gradeCL) = 0 Then gradeCL = Trim$(GetPropertyValue(propertiesToSet, "Material"))
                If Len(gradeCL) = 0 Then
                    On Error Resume Next
                    Dim matIdName2 As String
                    matIdName2 = model.GetMaterialPropertyName2("")
                    On Error GoTo 0
                    If Len(matIdName2) > 0 Then
                        Dim p As Long
                        p = InStr(matIdName2, "|")
                        If p > 0 Then
                            gradeCL = Mid$(matIdName2, p + 1)
                        Else
                            gradeCL = matIdName2
                        End If
                    End If
                End If
                gradeCL = Replace(Trim$(gradeCL), " ", "")
                
                Dim pnCL As String
                If t1 <> "ERR" Then
                    pnCL = "FB-" & t1 & "x" & t2
                    If Len(gradeCL) > 0 Then pnCL = pnCL & "-" & gradeCL
                    propertiesToSet("Mtl Part Number") = pnCL
                    DebugLog "ProcessFB: Created manual Mtl Part Number (structural fallback): " & pnCL
                End If
            End If
            
        ElseIf isTypeP Then
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "ProcessFB: Removed Mtl Part Number for Type P"
            End If
        End If
    End If
End Sub






' Process L shapes with conditional metric handling based on AS NZS
Sub ProcessLShape(ByVal model As ModelDoc2, ByVal stockSize As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object, Optional ByVal swDimensions As Variant = Empty)
    DebugLog "=== ProcessLShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "modelUOM: " & modelUOM
    
    ' *** CHECK IF THIS IS AN AS NZS STRUCTURAL MEMBER ***
    Dim isASNZS As Boolean
    isASNZS = IsASNZSStructuralMember(model)
    DebugLog "IsASNZSStructuralMember: " & isASNZS
    
    ' *** NEW: Extract dimensions from structural member name for imperial members ***
    Dim extractedFromMemberName As Boolean
    Dim memberDimensions As String
    extractedFromMemberName = False

    ' Look for the structural member feature to get clean dimensions
    Dim swModelDoc As SldWorks.ModelDoc2
    Set swModelDoc = model

    Dim swFeat As SldWorks.Feature
    Set swFeat = swModelDoc.FirstFeature

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "WeldMemberFeat" Then
            Dim memberName As String
            memberName = swFeat.Name
            DebugLog "Found WeldMemberFeat: " & memberName
            
            ' Look for pattern like "L angle 4X3X0.5"
            If InStr(UCase(memberName), "L ANGLE") > 0 Then
                Dim startPos As Integer
                startPos = InStr(UCase(memberName), "L ANGLE") + 7 ' Skip "L ANGLE"
                
                Dim endPos As Integer
                endPos = InStr(startPos, memberName, "(")
                If endPos = 0 Then endPos = Len(memberName) + 1
                
                Dim dimensionPart As String
                dimensionPart = Trim(Mid(memberName, startPos, endPos - startPos))
                
                ' Clean up and validate format (should be like "4X3X0.5")
                If InStr(UCase(dimensionPart), "X") > 0 Then
                    Dim testParts() As String
                    testParts = Split(Replace(UCase(dimensionPart), "X", "x"), "x")
                    
                    If UBound(testParts) = 2 Then
                        ' Validate all parts are numeric
                        Dim allNumeric As Boolean
                        allNumeric = True
                        Dim testI As Integer
                        For testI = 0 To 2
                            If Not IsNumeric(Trim(testParts(testI))) Then
                                allNumeric = False
                                Exit For
                            End If
                        Next testI
                        
                        If allNumeric Then
                            memberDimensions = dimensionPart
                            extractedFromMemberName = True
                            DebugLog "Extracted dimensions from member name: " & memberDimensions
                            Exit Do
                        End If
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    ' If we successfully extracted from member name, use those dimensions
    If extractedFromMemberName Then
        DebugLog "=== USING DIMENSIONS FROM MEMBER NAME ==="
        DebugLog "Member dimensions: " & memberDimensions
        
        ' Parse the extracted dimensions
        Dim memberParts() As String
        memberParts = Split(Replace(UCase(memberDimensions), "X", "x"), "x")
        
        If UBound(memberParts) = 2 Then
            Dim memberLDimensions(2) As Double
            Dim memberSuccess As Boolean
            memberSuccess = True
            
            ' Parse each dimension
            Dim memberI As Integer
            For memberI = 0 To 2
                memberParts(memberI) = Trim(memberParts(memberI))
                If IsNumeric(memberParts(memberI)) Then
                    memberLDimensions(memberI) = CDbl(memberParts(memberI))
                Else
                    memberSuccess = False
                    Exit For
                End If
            Next memberI
            
            If memberSuccess Then
                DebugLog "Imperial structural member dimensions (inches): " & memberLDimensions(0) & ", " & memberLDimensions(1) & ", " & memberLDimensions(2)
                
                ' These are already in inches, no conversion needed
                ' Sort dimensions for stock size (thickness x leg1 x leg2)
                BubbleSort memberLDimensions
                
                ' Format dimensions
                Dim memberFormattedDims(2) As String
                For memberI = 0 To 2
                    If memberLDimensions(memberI) = Int(memberLDimensions(memberI)) Then
                        memberFormattedDims(memberI) = CStr(Int(memberLDimensions(memberI)))
                    Else
                        memberFormattedDims(memberI) = RemoveTrailingZeros(Format(memberLDimensions(memberI), "0.#####"))
                    End If
                Next memberI
                
                stockSize = "L " & memberFormattedDims(0) & " x " & memberFormattedDims(1) & " x " & memberFormattedDims(2)
                propertiesToSet("Stock Size") = stockSize
                DebugLog "Stock Size from member name: " & stockSize

                ' Continue to check for total length even though we got dimensions from member name
                ' (The total length detection will still run, then we'll skip the redundant stock size parsing)
            End If
        End If
    End If
    
    ' *** CHECK FOR TOTAL LENGTH FROM CUTLIST FIRST ***
    DebugLog "=== CHECKING FOR TOTAL LENGTH IN CUTLIST ==="
    Dim totalLengthValue As String
    Dim totalLength As Double
    Dim useTotalLength As Boolean
    useTotalLength = False
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = model
    
    If Not swModel Is Nothing Then
        ' Look for the CutListFolder with Type = "L"
        Dim swFeature As SldWorks.Feature
        Set swFeature = swModel.FirstFeature
        
        Do While Not swFeature Is Nothing
            If swFeature.GetTypeName2 = "CutListFolder" Then
                Dim swCustPropMgr As SldWorks.customPropertyManager
                Set swCustPropMgr = swFeature.customPropertyManager
                
                If Not swCustPropMgr Is Nothing Then
                    Dim valueOut As String, resolvedValueOut As String
                    Dim bRet As Boolean
                    
                    bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                    If bRet And resolvedValueOut = "L" Then
                        ' Check if this cutlist folder has actual bodies
                        Dim swBodyFolder As SldWorks.BodyFolder
                        Set swBodyFolder = swFeature.GetSpecificFeature2
                        Dim vBodies As Variant
                        If Not swBodyFolder Is Nothing Then
                            vBodies = swBodyFolder.GetBodies
                        End If
                        
                        If Not IsEmpty(vBodies) Then
                            DebugLog "Found L-type cutlist folder WITH BODIES: " & swFeature.Name
                            
                            ' Try to get TOTAL LENGTH from cutlist
                            bRet = swCustPropMgr.Get4("TOTAL LENGTH", False, valueOut, totalLengthValue)
                            If Not bRet Or totalLengthValue = "" Then
                                bRet = swCustPropMgr.Get4("Total Length", False, valueOut, totalLengthValue)
                            End If
                            If Not bRet Or totalLengthValue = "" Then
                                bRet = swCustPropMgr.Get4("LENGTH", False, valueOut, totalLengthValue)
                            End If
                            
                            If bRet And totalLengthValue <> "" And IsNumeric(totalLengthValue) Then
                                totalLength = CDbl(totalLengthValue)
                                If totalLength > 0 Then
                                    useTotalLength = True
                                    DebugLog "Found Total Length from cutlist: " & totalLength
                                    
                                    ' *** CONVERSION LOGIC FOR TOTAL LENGTH ***
                                    If IsMetricUnit(modelUOM) Then
                                        If isASNZS Then
                                            DebugLog "=== AS NZS DETECTED - KEEPING TOTAL LENGTH IN MM ==="
                                            DebugLog "Using total length as-is (metric): " & totalLength
                                        Else
                                            DebugLog "=== METRIC MODEL WITH IMPERIAL ORDERING - CONVERTING TOTAL LENGTH MM TO INCHES ==="
                                            DebugLog "Before conversion: " & totalLength & " mm"
                                            totalLength = ConvertModelUnitsToInches(totalLength, modelUOM)
                                            DebugLog "After conversion: " & totalLength & " inches"
                                        End If
                                    Else
                                        DebugLog "=== IMPERIAL MODEL - TOTAL LENGTH ALREADY IN INCHES ==="
                                        DebugLog "Using total length as-is: " & totalLength
                                    End If
                                End If
                            End If
                            Exit Do
                        End If
                    End If
                End If
            End If
            Set swFeature = swFeature.GetNextFeature
        Loop
    End If
    
    If useTotalLength Then
        DebugLog "=== USING TOTAL LENGTH FOR LENGTH DETECTION ==="
        
        Dim currentWidth As Double, currentHeight As Double, currentLength As Double
        currentWidth = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
        currentHeight = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
        currentLength = FractionToDecimal(GetPropertyValue(propertiesToSet, "Length"))
        
        Const tolerance As Double = 0.01
        
        ' Compare Total Length to dimensions to find which one matches
        If Abs(totalLength - currentWidth) < tolerance Then
            DebugLog "Total Length matches Width - swapping Width and Length"
            propertiesToSet("Width") = ConvertDecimalToFraction(currentLength)
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
        ElseIf Abs(totalLength - currentHeight) < tolerance Then
            DebugLog "Total Length matches Height - swapping Height and Length"
            propertiesToSet("Height") = ConvertDecimalToFraction(currentLength)
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
        Else
            DebugLog "Using Total Length as Length"
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
        End If
    End If

    ' *** STOCK SIZE PROCESSING WITH CONDITIONAL CONVERSION ***
    ' Skip this section if we already got stock size from member name
    If extractedFromMemberName Then
        DebugLog "=== SKIPPING STOCK SIZE PARSING (already got from member name) ==="
        GoTo SkipStockSizeParsing
    End If

    DebugLog "=== PROCESSING STOCK SIZE ==="

    ' Remove the "L" prefix if it exists
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), 2) = "L " Then
        cleanStockSize = Trim(Mid(cleanStockSize, 3))
    ElseIf Left(UCase(cleanStockSize), 1) = "L" Then
        cleanStockSize = Trim(Mid(cleanStockSize, 2))
    End If
    
    DebugLog "Clean stockSize (no prefix): " & cleanStockSize
    
    ' Split by "x" to get individual dimensions
    Dim parts() As String
    parts = Split(cleanStockSize, "x")
    
    If UBound(parts) = 2 Then
        Dim lDimensions(2) As Double
        Dim i As Integer
        Dim success As Boolean
        success = True
        
        ' Parse each dimension as a number
        For i = 0 To 2
            parts(i) = Trim(parts(i))
            DebugLog "Processing part " & i & ": '" & parts(i) & "'"
            
            If IsNumeric(parts(i)) Then
                lDimensions(i) = CDbl(parts(i))
                DebugLog "Parsed as number: " & lDimensions(i)
            Else
                DebugLog "ERROR: Part " & i & " is not numeric"
                success = False
                Exit For
            End If
        Next i
        
        If success Then
            ' *** CONVERSION LOGIC FOR STOCK SIZE DIMENSIONS ***
            If IsMetricUnit(modelUOM) Then
                If isASNZS Then
                    DebugLog "=== AS NZS DETECTED - KEEPING METRIC DIMENSIONS ==="
                    DebugLog "Using dimensions as-is (metric): " & lDimensions(0) & ", " & lDimensions(1) & ", " & lDimensions(2)
                Else
                    DebugLog "=== METRIC MODEL WITH IMPERIAL ORDERING - CONVERTING MM TO INCHES ==="
                    DebugLog "Before conversion (metric): " & lDimensions(0) & ", " & lDimensions(1) & ", " & lDimensions(2)
                    
                    For i = 0 To 2
                        lDimensions(i) = ConvertModelUnitsToInches(lDimensions(i), modelUOM)
                    Next i
                    
                    DebugLog "After conversion to inches: " & lDimensions(0) & ", " & lDimensions(1) & ", " & lDimensions(2)
                End If
            Else
                DebugLog "=== IMPERIAL MODEL - USING DIMENSIONS AS-IS ==="
                DebugLog "Using dimensions as-is (imperial): " & lDimensions(0) & ", " & lDimensions(1) & ", " & lDimensions(2)
            End If
            
            ' Build final dimensions and format output using clean decimal formatting
            Dim formattedDimensions() As String
            ReDim formattedDimensions(2)
            For i = 0 To 2
                If lDimensions(i) = Int(lDimensions(i)) Then
                    formattedDimensions(i) = CStr(Int(lDimensions(i)))
                Else
                    Dim cleanValue As Double
                    cleanValue = Round(lDimensions(i) * 16) / 16  ' Round to nearest 1/16"
                    If cleanValue = Int(cleanValue) Then
                        formattedDimensions(i) = CStr(Int(cleanValue))
                    Else
                        formattedDimensions(i) = RemoveTrailingZeros(Format(cleanValue, "0.########"))
                    End If
                End If
                DebugLog "Formatted dimension " & i & ": " & formattedDimensions(i)
            Next i
            
            ' Sort dimensions for stock size (thickness x leg1 x leg2)
            BubbleSort lDimensions
            ReDim formattedDimensions(2)
            For i = 0 To 2
                If lDimensions(i) = Int(lDimensions(i)) Then
                    formattedDimensions(i) = CStr(Int(lDimensions(i)))
                Else
                    cleanValue = Round(lDimensions(i) * 16) / 16  ' Round to nearest 1/16"
                    If cleanValue = Int(cleanValue) Then
                        formattedDimensions(i) = CStr(Int(cleanValue))
                    Else
                        formattedDimensions(i) = RemoveTrailingZeros(Format(cleanValue, "0.########"))
                    End If
                End If
            Next i
            
            stockSize = "L " & formattedDimensions(0) & " x " & formattedDimensions(1) & " x " & formattedDimensions(2)
            propertiesToSet("Stock Size") = stockSize
            DebugLog "Final Stock Size: " & stockSize
            
        Else
            DebugLog "ERROR: Failed to parse dimensions"
            stockSize = "UNKNOWN SIZE"
            propertiesToSet("Stock Size") = stockSize
        End If
    Else
        DebugLog "ERROR: Expected 3 dimensions separated by 'x', got " & (UBound(parts) + 1)
        stockSize = "UNKNOWN SIZE"
        propertiesToSet("Stock Size") = stockSize
    End If

SkipStockSizeParsing:
    
    ' *** MTL PART NUMBER PROCESSING (unchanged from original) ***
    DebugLog "=== L Shape Material Part Number Processing ==="
    
    ' Get Mtl Part Number directly from cutlist properties
    Dim originalPartNumber As String
    Dim lCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set lCustPropMgr = Nothing
    
    ' Find the L cutlist folder and get its Mtl Part Number
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgrMtl As customPropertyManager
            Set swCustPropMgrMtl = swFeature.customPropertyManager
            
            If Not swCustPropMgrMtl Is Nothing Then
                Dim valueOutMtl As String, resolvedValueOutMtl As String
                Dim bRetMtl As Boolean
                
                ' Check if this is the L cutlist folder
                bRetMtl = swCustPropMgrMtl.Get4("Type", False, valueOutMtl, resolvedValueOutMtl)
                If bRetMtl And resolvedValueOutMtl = "L" Then
                    ' Get Mtl Part Number from this cutlist
                    bRetMtl = swCustPropMgrMtl.Get4("Mtl Part Number", False, valueOutMtl, originalPartNumber)
                    If bRetMtl And originalPartNumber <> "" Then
                        Set lCustPropMgr = swCustPropMgrMtl
                        DebugLog "Found Mtl Part Number from L cutlist: " & originalPartNumber
                        Exit Do
                    Else
                        DebugLog "No Mtl Part Number found in L cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Update material in part number if needed
    If originalPartNumber <> "" And Not lCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String
        
        DebugLog "Original Mtl Part Number from cutlist: '" & originalPartNumber & "'"
        
        ' Get the actual material from the model properties
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If

        ' Clean up material name (remove prefixes, quotes, suffixes)
        If actualMaterial <> "" Then
            actualMaterial = Replace(actualMaterial, """", "")
            actualMaterial = Replace(actualMaterial, "SW-Material@", "")
            actualMaterial = Replace(actualMaterial, ".SLDPRT", "")
        End If

        DebugLog "Actual material from model (cleaned): '" & actualMaterial & "'"

        ' Only proceed if we have material
        If Len(actualMaterial) > 0 Then
            ' IMPROVED: Always rebuild to clean duplicates and ensure correct format
            DebugLog "Analyzing part number structure..."

            Dim partNumberParts() As String
            partNumberParts = Split(originalPartNumber, "-")

            DebugLog "Part number has " & (UBound(partNumberParts) + 1) & " parts:"
            Dim j As Integer
            For j = 0 To UBound(partNumberParts)
                DebugLog "  Part " & j & ": '" & partNumberParts(j) & "'"
            Next j

            If UBound(partNumberParts) >= 2 Then
                ' Always rebuild with Type-Size-Material to clean up any duplicates
                newPartNumber = partNumberParts(0) & "-" & partNumberParts(1) & "-" & actualMaterial
                DebugLog "Rebuilt part number: Type '" & partNumberParts(0) & "' + Size '" & partNumberParts(1) & "' + Material '" & actualMaterial & "'"
            ElseIf UBound(partNumberParts) = 1 Then
                ' Only Type-Size, add material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Added material to Type-Size format"
            Else
                ' Unexpected format, append material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Unexpected format - appending material"
            End If

            ' Update properties only if the part number changed
            If newPartNumber <> originalPartNumber Then
                lCustPropMgr.Set2 "Mtl Part Number", newPartNumber
                propertiesToSet("Mtl Part Number") = newPartNumber
                DebugLog "Updated Mtl Part Number: '" & originalPartNumber & "' -> '" & newPartNumber & "'"
            Else
                propertiesToSet("Mtl Part Number") = originalPartNumber
                DebugLog "Kept original Mtl Part Number: '" & originalPartNumber & "'"
            End If
            
        Else
            DebugLog "Cannot update part number - missing material"
            propertiesToSet("Mtl Part Number") = originalPartNumber
        End If
    ElseIf Not isTypeP Then
        ' Generate new Mtl Part Number for Type M if none exists in cutlist
        DebugLog "Generating new Mtl Part Number for Type M"
        
        Dim material As String
        material = GetPropertyValue(propertiesToSet, "Material")
        
        ' Clean up material name
        If material <> "" Then
            material = Replace(Replace(Replace(material, """", ""), "SW-Material@", ""), ".SLDPRT", "")
        Else
            material = "UNKNOWN"
        End If
        
        ' Get grade
        Dim grade As String
        grade = GetPropertyValue(propertiesToSet, "Grade")
        
        ' Generate Mtl Part Number using clean stock size
        Dim mtlPartNumber As String
        Dim cleanStockSizeForMtl As String
        cleanStockSizeForMtl = Replace(stockSize, "L ", "") ' Remove prefix
        cleanStockSizeForMtl = Replace(cleanStockSizeForMtl, " ", "") ' Remove spaces
        
        If grade <> "" Then
            mtlPartNumber = "L-" & cleanStockSizeForMtl & "-" & grade & "-" & material
        Else
            mtlPartNumber = "L-" & cleanStockSizeForMtl & "-" & material
        End If
        
        propertiesToSet("Mtl Part Number") = mtlPartNumber
        DebugLog "Generated new Mtl Part Number: " & mtlPartNumber
    ElseIf isTypeP Then
        ' Remove Mtl Part Number for Type P
        If propertiesToSet.exists("Mtl Part Number") Then
            propertiesToSet.Remove "Mtl Part Number"
            DebugLog "Removed Mtl Part Number (Type P)"
        End If
    End If
    
    DebugLog "=== ProcessLShape Debug End ==="
End Sub



' Simplified ProcessCShape function - clean stock size, no unit conversion needed
Sub ProcessCShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal stockSize As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "=== ProcessCShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    DebugLog "modelUOM: " & modelUOM
    
    ' SIMPLIFIED: No need for AS NZS detection since GetDimensions now handles units properly
    
    ' Strip the prefix if it exists (use refCategory instead of hardcoded "C")
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1)) ' Remove the prefix and any following spaces
    End If
    
    DebugLog "After prefix removal: '" & cleanStockSize & "'"
    
    ' IMPROVED: Remove unwanted suffixes like PFC, UAC, UBP, etc.
    Dim suffixesToRemove As Variant
    suffixesToRemove = Array("PFC", "UAC", "UBP", "TFC", "WBP", "mm", "MM")  ' Common suffixes to remove
    
    Dim i As Integer
    For i = 0 To UBound(suffixesToRemove)
        ' Remove the suffix if found (case-insensitive)
        If InStr(UCase(cleanStockSize), UCase(suffixesToRemove(i))) > 0 Then
            cleanStockSize = Replace(UCase(cleanStockSize), UCase(suffixesToRemove(i)), "", 1, 1, vbTextCompare)
            cleanStockSize = Trim(cleanStockSize)  ' Remove extra spaces
            DebugLog "Removed suffix '" & suffixesToRemove(i) & "': " & cleanStockSize
        End If
    Next i
    
    ' Clean up extra spaces
    Do While InStr(cleanStockSize, "  ") > 0  ' Remove double spaces
        cleanStockSize = Replace(cleanStockSize, "  ", " ")
    Loop
    cleanStockSize = Trim(cleanStockSize)
    
    DebugLog "After suffix removal and cleanup: '" & cleanStockSize & "'"
    
    ' Extract only the outer dimensions, ignoring anything in parentheses
    Dim cOuterDimensions As String
    If InStr(cleanStockSize, "(") > 0 Then
        cOuterDimensions = Trim(Left(cleanStockSize, InStr(cleanStockSize, "(") - 1))
    Else
        cOuterDimensions = cleanStockSize
    End If
    
    DebugLog "Outer dimensions: " & cOuterDimensions
    
    ' Split the outer dimensions - handle both "x" and space separators
    Dim cParts() As String
    If InStr(cOuterDimensions, "x") > 0 Or InStr(cOuterDimensions, "X") > 0 Then
        cParts = Split(Replace(cOuterDimensions, "X", "x"), "x")
    Else
        ' Handle space-separated format like "180 75"
        cParts = Split(cOuterDimensions, " ")
    End If
    
    ' Initialize an array to store formatted dimensions
    Dim cProcessedDimensions() As String
    ReDim cProcessedDimensions(UBound(cParts))
    
    ' SIMPLIFIED: No unit conversion needed - dimensions are already in model units
    DebugLog "Processing dimensions in model units (" & modelUOM & ") - no conversion needed"
    
    ' Process each outer dimension - just format them cleanly
    Dim cDimension As Double
    For i = 0 To UBound(cParts)
        cParts(i) = Trim(cParts(i))
        DebugLog "Processing dimension " & i & ": '" & cParts(i) & "'"
        
        If IsNumeric(cParts(i)) Then
            cDimension = CDbl(cParts(i))
            DebugLog "Dimension value: " & cDimension
            
            ' Format the dimension value cleanly
            If cDimension = Int(cDimension) Then
                ' If the value is an integer, use no decimals
                cProcessedDimensions(i) = CStr(Int(cDimension))
            Else
                ' Use RemoveTrailingZeros to preserve precision
                cProcessedDimensions(i) = RemoveTrailingZeros(Format(cDimension, "0.#####"))
            End If
            DebugLog "Formatted dimension: " & cProcessedDimensions(i)
        Else
            ' Keep non-numeric parts as-is
            cProcessedDimensions(i) = cParts(i)
            DebugLog "Non-numeric dimension kept as-is: " & cProcessedDimensions(i)
        End If
    Next i
    
    ' Construct the final Stock Size string - CLEAN format with no units
    stockSize = refCategory & " " & Join(cProcessedDimensions, " x ")
    
    ' Set the final Stock Size property
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Final Stock Size: " & stockSize
    
    ' *** MTL PART NUMBER PROCESSING WITH DUPLICATE FIX ***
    DebugLog "=== C Shape Material Part Number Processing ==="
    
    ' Get Mtl Part Number directly from cutlist properties
    Dim originalPartNumber As String
    Dim cCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set cCustPropMgr = Nothing
    
    ' Find the C/Z cutlist folder and get its Mtl Part Number
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the C or Z cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    ' Get Mtl Part Number from this cutlist
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set cCustPropMgr = swCustPropMgr
                        DebugLog "Found Mtl Part Number from " & refCategory & " cutlist: " & originalPartNumber
                        Exit Do
                    Else
                        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Only proceed if we found an original part number from the cutlist
    If originalPartNumber <> "" And Not cCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String
        
        DebugLog "Original Mtl Part Number from cutlist: '" & originalPartNumber & "'"
        
        ' Get the actual material from the model properties
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If

        ' Clean up material name (remove prefixes, quotes, suffixes)
        If actualMaterial <> "" Then
            actualMaterial = Replace(actualMaterial, """", "")
            actualMaterial = Replace(actualMaterial, "SW-Material@", "")
            actualMaterial = Replace(actualMaterial, ".SLDPRT", "")
        End If

        DebugLog "Actual material from model (cleaned): '" & actualMaterial & "'"

        ' Only proceed if we have material
        If Len(actualMaterial) > 0 Then
            ' IMPROVED: Always rebuild to clean duplicates and ensure correct format
            ' Expected format: Type-Size-Material (e.g., CS-6x3.59-6061-T6)
            DebugLog "Analyzing part number structure..."

            Dim partNumberParts() As String
            partNumberParts = Split(originalPartNumber, "-")

            DebugLog "Part number has " & (UBound(partNumberParts) + 1) & " parts:"
            Dim j As Integer
            For j = 0 To UBound(partNumberParts)
                DebugLog "  Part " & j & ": '" & partNumberParts(j) & "'"
            Next j

            ' Expected structure: partNumberParts(0) = Type (C, CS, etc.)
            '                    partNumberParts(1) = Size (6x3.59, etc.)
            '                    partNumberParts(2+) = Material (may have duplicates to clean)

            If UBound(partNumberParts) >= 2 Then
                ' Always rebuild with Type-Size-Material to clean up any duplicates
                newPartNumber = partNumberParts(0) & "-" & partNumberParts(1) & "-" & actualMaterial
                DebugLog "Rebuilt part number: Type '" & partNumberParts(0) & "' + Size '" & partNumberParts(1) & "' + Material '" & actualMaterial & "'"
            ElseIf UBound(partNumberParts) = 1 Then
                ' Only Type-Size, add material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Added material to Type-Size format"
            Else
                ' Unexpected format, append material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Unexpected format - appending material"
            End If

            ' Update properties only if the part number changed
            If newPartNumber <> originalPartNumber Then
                cCustPropMgr.Set2 "Mtl Part Number", newPartNumber
                propertiesToSet("Mtl Part Number") = newPartNumber
                DebugLog "Updated Mtl Part Number: '" & originalPartNumber & "' -> '" & newPartNumber & "'"
            Else
                propertiesToSet("Mtl Part Number") = originalPartNumber
                DebugLog "Kept original Mtl Part Number: '" & originalPartNumber & "'"
            End If
            
        Else
            DebugLog "Cannot update part number - missing material"
            propertiesToSet("Mtl Part Number") = originalPartNumber
        End If
    Else
        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist - generating new one"
        
        ' Generate new Mtl Part Number for Type M if none exists in cutlist
        If Not isTypeP Then
            DebugLog "Generating new Mtl Part Number for Type M"
            
            Dim material As String
            material = GetPropertyValue(propertiesToSet, "Material")
            
            ' Clean up material name
            If material <> "" Then
                material = Replace(material, """", "")
                material = Replace(material, "SW-Material@", "")
                material = Replace(material, ".SLDPRT", "")
            Else
                material = "UNKNOWN"
            End If
            
            ' Get grade
            Dim grade As String
            grade = GetPropertyValue(propertiesToSet, "Grade")
            
            ' Generate Mtl Part Number using clean stock size
            Dim mtlPartNumber As String
            Dim cleanStockSizeForMtl As String
            cleanStockSizeForMtl = Replace(stockSize, refCategory & " ", "") ' Remove prefix
            cleanStockSizeForMtl = Replace(cleanStockSizeForMtl, " ", "") ' Remove spaces
            
            If grade <> "" Then
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & grade & "-" & material
            Else
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & material
            End If
            
            propertiesToSet("Mtl Part Number") = mtlPartNumber
            DebugLog "Generated new Mtl Part Number: " & mtlPartNumber
        ElseIf isTypeP Then
            ' Remove Mtl Part Number for Type P
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "Removed Mtl Part Number (Type P)"
            End If
        End If
    End If
    
    DebugLog "=== ProcessCShape Debug End ==="
End Sub


' Process W and I shapes with proper dimension processing and Mtl Part Number handling
Sub ProcessWShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal stockSize As String, _
                  ByVal isTypeP As Boolean, ByRef propertiesToSet As Object)
    DebugLog "=== ProcessWShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    
    ' Strip existing prefix if it exists
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1))
        ' Handle dash format like "W-" or "I-"
        If Left(cleanStockSize, 1) = "-" Then
            cleanStockSize = Trim(Mid(cleanStockSize, 2))
        End If
    End If
    
    DebugLog "After prefix removal: '" & cleanStockSize & "'"
    
    ' Split the dimensions by 'x'
    Dim wDimensions() As String
    wDimensions = Split(Replace(cleanStockSize, "X", "x"), "x")
    
    If UBound(wDimensions) >= 1 Then
        Dim processedDimensions() As String
        ReDim processedDimensions(UBound(wDimensions))
        
        ' SIMPLIFIED: No unit conversion needed - process dimensions in model units
        DebugLog "Processing dimensions in model units - no conversion needed"
        
        Dim i As Integer
        For i = 0 To UBound(wDimensions)
            wDimensions(i) = Trim(wDimensions(i))
            DebugLog "Processing dimension " & i & ": '" & wDimensions(i) & "'"
            
            If IsNumeric(wDimensions(i)) Then
                Dim dimValue As Double
                dimValue = CDbl(wDimensions(i))
                DebugLog "Dimension value: " & dimValue
                
                ' Format the dimension value cleanly
                If dimValue = Int(dimValue) Then
                    processedDimensions(i) = CStr(Int(dimValue))
                Else
                    processedDimensions(i) = RemoveTrailingZeros(Format(dimValue, "0.#####"))
                End If
                DebugLog "Formatted dimension: " & processedDimensions(i)
            Else
                processedDimensions(i) = wDimensions(i)
                DebugLog "Non-numeric dimension kept as-is: " & processedDimensions(i)
            End If
        Next i
        
        stockSize = refCategory & " " & Join(processedDimensions, " x ")
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Final Stock Size: " & stockSize
    Else
        stockSize = "UNKNOWN SIZE"
        propertiesToSet("Stock Size") = stockSize
        DebugLog "ERROR: Could not parse dimensions"
    End If
    
    ' *** MTL PART NUMBER PROCESSING ***
    DebugLog "=== " & refCategory & " Material Part Number Processing ==="
    
    ' Get Mtl Part Number directly from cutlist properties
    Dim originalPartNumber As String
    Dim wCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set wCustPropMgr = Nothing
    
    ' Find the W/I cutlist folder and get its Mtl Part Number
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the W or I cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    ' Get Mtl Part Number from this cutlist
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set wCustPropMgr = swCustPropMgr
                        DebugLog "Found Mtl Part Number from " & refCategory & " cutlist: " & originalPartNumber
                        Exit Do
                    Else
                        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Update material in part number if needed
    If originalPartNumber <> "" And Not wCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String

        DebugLog "Original Mtl Part Number from cutlist: '" & originalPartNumber & "'"

        ' Get the actual material from the model properties
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If

        ' Clean up material name (remove prefixes, quotes, suffixes)
        If actualMaterial <> "" Then
            actualMaterial = Replace(actualMaterial, """", "")
            actualMaterial = Replace(actualMaterial, "SW-Material@", "")
            actualMaterial = Replace(actualMaterial, ".SLDPRT", "")
        End If

        DebugLog "Actual material from model (cleaned): '" & actualMaterial & "'"

        ' Only proceed if we have material
        If Len(actualMaterial) > 0 Then
            ' IMPROVED: Always rebuild to clean duplicates and ensure correct format
            DebugLog "Analyzing part number structure..."

            Dim partNumberParts() As String
            partNumberParts = Split(originalPartNumber, "-")

            DebugLog "Part number has " & (UBound(partNumberParts) + 1) & " parts:"
            Dim j As Integer
            For j = 0 To UBound(partNumberParts)
                DebugLog "  Part " & j & ": '" & partNumberParts(j) & "'"
            Next j

            If UBound(partNumberParts) >= 2 Then
                ' Always rebuild with Type-Size-Material to clean up any duplicates
                newPartNumber = partNumberParts(0) & "-" & partNumberParts(1) & "-" & actualMaterial
                DebugLog "Rebuilt part number: Type '" & partNumberParts(0) & "' + Size '" & partNumberParts(1) & "' + Material '" & actualMaterial & "'"
            ElseIf UBound(partNumberParts) = 1 Then
                ' Only Type-Size, add material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Added material to Type-Size format"
            Else
                ' Unexpected format, append material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Unexpected format - appending material"
            End If
            
            ' Update properties only if the part number changed
            If newPartNumber <> originalPartNumber Then
                wCustPropMgr.Set2 "Mtl Part Number", newPartNumber
                propertiesToSet("Mtl Part Number") = newPartNumber
                DebugLog "Updated Mtl Part Number: '" & originalPartNumber & "' -> '" & newPartNumber & "'"
            Else
                propertiesToSet("Mtl Part Number") = originalPartNumber
                DebugLog "Kept original Mtl Part Number: '" & originalPartNumber & "'"
            End If
            
        Else
            DebugLog "Cannot update part number - missing material"
            propertiesToSet("Mtl Part Number") = originalPartNumber
        End If
    Else
        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist - generating new one"
        
        ' Generate new Mtl Part Number for Type M if none exists in cutlist
        If Not isTypeP Then
            DebugLog "Generating new Mtl Part Number for Type M"
            
            Dim material As String
            material = GetPropertyValue(propertiesToSet, "Material")
            
            ' Clean up material name
            If material <> "" Then
                material = Replace(Replace(Replace(material, """", ""), "SW-Material@", ""), ".SLDPRT", "")
            Else
                material = "UNKNOWN"
            End If
            
            ' Get grade
            Dim grade As String
            grade = GetPropertyValue(propertiesToSet, "Grade")
            
            ' Generate Mtl Part Number using clean stock size
            Dim mtlPartNumber As String
            Dim cleanStockSizeForMtl As String
            cleanStockSizeForMtl = Replace(stockSize, refCategory & " ", "") ' Remove prefix
            cleanStockSizeForMtl = Replace(cleanStockSizeForMtl, " ", "") ' Remove spaces
            
            If grade <> "" Then
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & grade & "-" & material
            Else
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & material
            End If
            
            propertiesToSet("Mtl Part Number") = mtlPartNumber
            DebugLog "Generated new Mtl Part Number: " & mtlPartNumber
        ElseIf isTypeP Then
            ' Remove Mtl Part Number for Type P
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "Removed Mtl Part Number (Type P)"
            End If
        End If
    End If
    
    DebugLog "=== ProcessWShape Debug End ==="
End Sub


' Process HSS shapes - ALWAYS uses A500C material (unless AS NZS)
Sub ProcessHSSShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal stockSize As String, ByVal isTypeP As Boolean, _
                    ByVal modelUOM As String, ByRef propertiesToSet As Object, ByRef swDimensions() As Double)
    
    DebugLog "=== ProcessHSSShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    DebugLog "Passed modelUOM parameter: " & modelUOM
    
    ' *** FIX: Get ORIGINAL model UOM directly from model (bypasses converted parameter) ***
    Dim originalModelUOM As String
    originalModelUOM = GetModelUnitOfMeasure(model)
    DebugLog "Original model UOM from model: " & originalModelUOM
    
    ' *** CHECK IF THIS IS AN AS NZS STRUCTURAL MEMBER ***
    Dim isASNZS As Boolean
    isASNZS = IsASNZSStructuralMember(model)
    DebugLog "Is AS NZS structural member: " & isASNZS
    
    ' *** NEW: Get TOTAL LENGTH from cutlist FIRST ***
    DebugLog "=== CHECKING FOR TOTAL LENGTH IN CUTLIST ==="
    Dim totalLengthValue As Double
    Dim foundTotalLength As Boolean
    foundTotalLength = False
    totalLengthValue = 0
    
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the HSS cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    DebugLog "Found " & refCategory & " cutlist folder"
                    
                    ' Try to get TOTAL LENGTH
                    bRet = swCustPropMgr.Get4("TOTAL LENGTH", False, valueOut, resolvedValueOut)
                    If Not bRet Or resolvedValueOut = "" Then
                        bRet = swCustPropMgr.Get4("Total Length", False, valueOut, resolvedValueOut)
                    End If
                    If Not bRet Or resolvedValueOut = "" Then
                        bRet = swCustPropMgr.Get4("LENGTH", False, valueOut, resolvedValueOut)
                    End If
                    
                    If bRet And resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                        totalLengthValue = CDbl(resolvedValueOut)
                        foundTotalLength = True
                        DebugLog "Found Total Length from cutlist: " & totalLengthValue
                        
                        ' *** FIX: Only convert to inches if NOT AS NZS and model is metric ***
                        If IsMetricUnit(originalModelUOM) And Not isASNZS Then
                            DebugLog "Converting total length from " & originalModelUOM & " to inches"
                            totalLengthValue = ConvertModelUnitsToInches(totalLengthValue, originalModelUOM)
                            DebugLog "Converted total length: " & totalLengthValue & " inches"
                        ElseIf isASNZS Then
                            DebugLog "AS NZS member - keeping length in " & originalModelUOM
                        End If
                        
                        ' Set the length properties NOW
                        If isASNZS And originalModelUOM = "mm" Then
                            ' For AS NZS in mm, store as decimal mm value
                            propertiesToSet("Length") = RemoveTrailingZeros(Format(totalLengthValue, "0.#####"))
                            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLengthValue, "0.#####"))
                        Else
                            ' For imperial or converted values, use fractions
                            propertiesToSet("Length") = ConvertDecimalToFraction(totalLengthValue)
                            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLengthValue, "0.#####"))
                        End If
                        propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
                        DebugLog "Set Length: " & propertiesToSet("Length")
                        DebugLog "Set LengthA: " & propertiesToSet("LengthA")
                        DebugLog "Set Mtl Unit Qty: " & propertiesToSet("Mtl Unit Qty")
                        Exit Do
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' *** Handle missing StockSize by extracting from Mtl Part Number ***
    If stockSize = "UNKNOWN SIZE" Or stockSize = "" Then
        DebugLog "StockSize is missing - attempting to extract from Mtl Part Number"
        
        Dim mtlPartNumber As String
        mtlPartNumber = GetPropertyValue(propertiesToSet, "Mtl Part Number")
        
        If mtlPartNumber <> "" Then
            DebugLog "Extracting from Mtl Part Number: " & mtlPartNumber
            
            Dim withoutPrefix As String
            If Left(UCase(mtlPartNumber), Len(refCategory) + 1) = UCase(refCategory) & "-" Then
                withoutPrefix = Mid(mtlPartNumber, Len(refCategory) + 2)
            Else
                withoutPrefix = mtlPartNumber
            End If
            
            Dim lastDashPos As Integer
            lastDashPos = InStrRev(withoutPrefix, "-")
            
            Dim dimensionsPart As String
            If lastDashPos > 0 Then
                dimensionsPart = Left(withoutPrefix, lastDashPos - 1)
            Else
                dimensionsPart = withoutPrefix
            End If
            
            If dimensionsPart <> "" And InStr(dimensionsPart, "x") > 0 Then
                stockSize = refCategory & " " & dimensionsPart
                DebugLog "Extracted StockSize: " & stockSize
            End If
        End If
    End If
    
    If stockSize = "UNKNOWN SIZE" Or stockSize = "" Then
        DebugLog "Cannot process HSS - no valid dimensions available"
        propertiesToSet("Stock Size") = "UNKNOWN SIZE"
        Exit Sub
    End If
    
    ' *** CHECK IF STOCKSIZE IS DYNAMIC (contains model name) ***
    DebugLog "=== CHECKING IF STOCKSIZE IS DYNAMIC ==="
    Dim isDynamicStockSize As Boolean
    Dim stockSizeExpression As String
    isDynamicStockSize = False
    
    ' Get the model name
    Dim modelName As String
    modelName = model.GetTitle
    DebugLog "Model name: " & modelName
    
    ' Find the cutlist and check StockSize expression
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    ' Get the StockSize property expression
                    bRet = swCustPropMgr.Get4("StockSize", False, stockSizeExpression, resolvedValueOut)
                    
                    If bRet Then
                        DebugLog "StockSize expression: " & stockSizeExpression
                        
                        ' Check if expression contains the model name (indicates dynamic property)
                        If InStr(stockSizeExpression, "@") > 0 And InStr(stockSizeExpression, modelName) > 0 Then
                            isDynamicStockSize = True
                            DebugLog "*** DYNAMIC STOCKSIZE DETECTED (contains model name) ***"
                        Else
                            DebugLog "*** STATIC STOCKSIZE DETECTED (hardcoded value) ***"
                        End If
                    End If
                    Exit Do
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' *** STOCK SIZE PROCESSING WITH CONDITIONAL CONVERSION ***
    DebugLog "=== PROCESSING STOCK SIZE ==="
    
    ' Strip the prefix
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1))
        If Left(cleanStockSize, 1) = "-" Then
            cleanStockSize = Trim(Mid(cleanStockSize, 2))
        End If
    End If
    
    DebugLog "After prefix removal: '" & cleanStockSize & "'"
    
    ' Process Stock Size dimensions
    Dim hssDimensions() As String
    hssDimensions = Split(Replace(cleanStockSize, "X", "x"), "x")

    If UBound(hssDimensions) >= 1 Then
        Dim hssProcessedDimensions() As String
        ReDim hssProcessedDimensions(UBound(hssDimensions))
        
        ' *** FIX: CONDITIONAL CONVERSION - Skip conversion for AS NZS members ***
        If isDynamicStockSize And IsMetricUnit(originalModelUOM) And Not isASNZS Then
            DebugLog "*** DYNAMIC + METRIC + NON-AS-NZS: Converting dimensions from " & originalModelUOM & " to inches ***"
            
            Dim hssI As Integer
            For hssI = 0 To UBound(hssDimensions)
                hssDimensions(hssI) = Trim(hssDimensions(hssI))
                DebugLog "Processing dimension " & hssI & ": '" & hssDimensions(hssI) & "'"

                If IsNumeric(hssDimensions(hssI)) Then
                    Dim dimValue As Double
                    dimValue = CDbl(hssDimensions(hssI))
                    DebugLog "Original value (metric): " & dimValue
                    
                    ' Convert from metric to inches using ORIGINAL model UOM
                    dimValue = ConvertModelUnitsToInches(dimValue, originalModelUOM)
                    DebugLog "Converted to inches: " & dimValue
                    
                    ' Clean format
                    If dimValue = Int(dimValue) Then
                        hssProcessedDimensions(hssI) = CStr(Int(dimValue))
                    Else
                        hssProcessedDimensions(hssI) = RemoveTrailingZeros(Format(dimValue, "0.####"))
                    End If
                    DebugLog "Formatted dimension: " & hssProcessedDimensions(hssI)
                Else
                    hssProcessedDimensions(hssI) = hssDimensions(hssI)
                    DebugLog "Non-numeric dimension kept as-is: " & hssProcessedDimensions(hssI)
                End If
            Next hssI
        Else
            If isASNZS Then
                DebugLog "*** AS NZS MEMBER: Using dimensions as-is (no conversion) ***"
            Else
                DebugLog "*** STATIC OR IMPERIAL: Using dimensions as-is (no conversion) ***"
            End If
            
            For hssI = 0 To UBound(hssDimensions)
                hssDimensions(hssI) = Trim(hssDimensions(hssI))
                DebugLog "Processing dimension " & hssI & ": '" & hssDimensions(hssI) & "'"

                If IsNumeric(hssDimensions(hssI)) Then
                    dimValue = CDbl(hssDimensions(hssI))
                    DebugLog "Dimension value (no conversion): " & dimValue
                    
                    ' Clean format
                    If dimValue = Int(dimValue) Then
                        hssProcessedDimensions(hssI) = CStr(Int(dimValue))
                    Else
                        hssProcessedDimensions(hssI) = RemoveTrailingZeros(Format(dimValue, "0.####"))
                    End If
                    DebugLog "Formatted dimension: " & hssProcessedDimensions(hssI)
                Else
                    hssProcessedDimensions(hssI) = hssDimensions(hssI)
                    DebugLog "Non-numeric dimension kept as-is: " & hssProcessedDimensions(hssI)
                End If
            Next hssI
        End If

        ' Set final Stock Size
        stockSize = refCategory & " " & Join(hssProcessedDimensions, " x ")
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Final stockSize: " & stockSize

        ' Set dimensions for swDimensions array
        Dim hssWidth As Double, hssHeight As Double
        If IsNumeric(hssProcessedDimensions(0)) Then hssWidth = CDbl(hssProcessedDimensions(0))
        If IsNumeric(hssProcessedDimensions(1)) Then hssHeight = CDbl(hssProcessedDimensions(1))

        ' Use the total length we found earlier if available
        If foundTotalLength Then
            swDimensions(0) = totalLengthValue
        ElseIf propertiesToSet.exists("Mtl Unit Qty") Then
            swDimensions(0) = CDbl(propertiesToSet("Mtl Unit Qty"))
        End If
        swDimensions(1) = hssWidth
        swDimensions(2) = hssHeight
    Else
        stockSize = "UNKNOWN SIZE"
        propertiesToSet("Stock Size") = stockSize
        DebugLog "ERROR: Could not parse HSS dimensions"
    End If
    
    ' *** HANDLE MTL PART NUMBER MATERIAL REPLACEMENT ***
    ' HSS uses A500C for AISC/ASTM, but AS NZS uses different grades
    DebugLog "=== HSS Material Part Number Processing ==="
    
    Dim originalPartNumber As String
    Dim hssCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set hssCustPropMgr = Nothing
    
    ' Find the HSS cutlist folder and get its Mtl Part Number
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set hssCustPropMgr = swCustPropMgr
                        DebugLog "Found Mtl Part Number from " & refCategory & " cutlist: " & originalPartNumber
                        Exit Do
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Update material in part number
    If originalPartNumber <> "" And Not hssCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String
        
        ' *** FIX: Use A500C only for non-AS-NZS members ***
        If isASNZS Then
            DebugLog "AS NZS HSS detected - keeping original grade from cutlist"
            ' For AS NZS, try to get grade from cutlist
            bRet = hssCustPropMgr.Get4("Grade", False, valueOut, actualMaterial)
            If bRet And actualMaterial <> "" Then
                DebugLog "Using AS NZS grade from cutlist: " & actualMaterial
            Else
                ' Default AS NZS grade if not found
                actualMaterial = "C350L0"
                DebugLog "No grade found - defaulting to C350L0"
            End If
        Else
            ' *** HSS (non-AS-NZS) ALWAYS USES A500C ***
            actualMaterial = "A500C"
            DebugLog "HSS (AISC/ASTM) detected - forcing material to A500C"
        End If
        
        Dim partNumberParts() As String
        partNumberParts = Split(originalPartNumber, "-")
        
        If UBound(partNumberParts) >= 2 Then
            partNumberParts(UBound(partNumberParts)) = actualMaterial
        ElseIf UBound(partNumberParts) = 1 Then
            ReDim Preserve partNumberParts(UBound(partNumberParts) + 1)
            partNumberParts(UBound(partNumberParts)) = actualMaterial
        End If
        
        newPartNumber = Join(partNumberParts, "-")
        hssCustPropMgr.Set2 "Mtl Part Number", newPartNumber
        propertiesToSet("Mtl Part Number") = newPartNumber
        propertiesToSet("MATERIAL") = actualMaterial
        DebugLog "Updated Mtl Part Number: " & newPartNumber
    End If
    
    DebugLog "=== ProcessHSSShape Debug End ==="
End Sub

' Process MEC and AS shapes - uses actual material from model
Sub ProcessMECASShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal stockSize As String, ByVal isTypeP As Boolean, _
                      ByVal modelUOM As String, ByRef propertiesToSet As Object, ByRef swDimensions() As Double)
    
    DebugLog "=== ProcessMECASShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    DebugLog "Passed modelUOM parameter: " & modelUOM
    
    ' *** FIX: Get ORIGINAL model UOM directly from model (bypasses converted parameter) ***
    Dim originalModelUOM As String
    originalModelUOM = GetModelUnitOfMeasure(model)
    DebugLog "Original model UOM from model: " & originalModelUOM
    
    ' *** NEW: Get TOTAL LENGTH from cutlist FIRST ***
    DebugLog "=== CHECKING FOR TOTAL LENGTH IN CUTLIST ==="
    Dim totalLengthValue As Double
    Dim foundTotalLength As Boolean
    foundTotalLength = False
    totalLengthValue = 0
    
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the MEC/AS cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    DebugLog "Found " & refCategory & " cutlist folder"
                    
                    ' Try to get TOTAL LENGTH
                    bRet = swCustPropMgr.Get4("TOTAL LENGTH", False, valueOut, resolvedValueOut)
                    If Not bRet Or resolvedValueOut = "" Then
                        bRet = swCustPropMgr.Get4("Total Length", False, valueOut, resolvedValueOut)
                    End If
                    If Not bRet Or resolvedValueOut = "" Then
                        bRet = swCustPropMgr.Get4("LENGTH", False, valueOut, resolvedValueOut)
                    End If
                    
                    If bRet And resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                        totalLengthValue = CDbl(resolvedValueOut)
                        foundTotalLength = True
                        DebugLog "Found Total Length from cutlist: " & totalLengthValue
                        
                        ' Convert to inches if ORIGINAL model is metric
                        If IsMetricUnit(originalModelUOM) Then
                            DebugLog "Converting total length from " & originalModelUOM & " to inches"
                            totalLengthValue = ConvertModelUnitsToInches(totalLengthValue, originalModelUOM)
                            DebugLog "Converted total length: " & totalLengthValue & " inches"
                        End If
                        
                        ' Set the length properties NOW
                        propertiesToSet("Length") = ConvertDecimalToFraction(totalLengthValue)
                        propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLengthValue, "0.#####"))
                        propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
                        DebugLog "Set Length: " & propertiesToSet("Length")
                        DebugLog "Set LengthA: " & propertiesToSet("LengthA")
                        DebugLog "Set Mtl Unit Qty: " & propertiesToSet("Mtl Unit Qty")
                        Exit Do
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' *** Handle missing StockSize by extracting from Mtl Part Number ***
    If stockSize = "UNKNOWN SIZE" Or stockSize = "" Then
        DebugLog "StockSize is missing - attempting to extract from Mtl Part Number"
        
        Dim mtlPartNumber As String
        mtlPartNumber = GetPropertyValue(propertiesToSet, "Mtl Part Number")
        
        If mtlPartNumber <> "" Then
            DebugLog "Extracting from Mtl Part Number: " & mtlPartNumber
            
            Dim withoutPrefix As String
            If Left(UCase(mtlPartNumber), Len(refCategory) + 1) = UCase(refCategory) & "-" Then
                withoutPrefix = Mid(mtlPartNumber, Len(refCategory) + 2)
            Else
                withoutPrefix = mtlPartNumber
            End If
            
            Dim lastDashPos As Integer
            lastDashPos = InStrRev(withoutPrefix, "-")
            
            Dim dimensionsPart As String
            If lastDashPos > 0 Then
                dimensionsPart = Left(withoutPrefix, lastDashPos - 1)
            Else
                dimensionsPart = withoutPrefix
            End If
            
            If dimensionsPart <> "" And InStr(dimensionsPart, "x") > 0 Then
                stockSize = refCategory & " " & dimensionsPart
                DebugLog "Extracted StockSize: " & stockSize
            End If
        End If
    End If
    
    If stockSize = "UNKNOWN SIZE" Or stockSize = "" Then
        DebugLog "Cannot process MEC/AS - no valid dimensions available"
        propertiesToSet("Stock Size") = "UNKNOWN SIZE"
        Exit Sub
    End If
    
    ' *** CHECK IF STOCKSIZE IS DYNAMIC (contains model name) ***
    DebugLog "=== CHECKING IF STOCKSIZE IS DYNAMIC ==="
    Dim isDynamicStockSize As Boolean
    Dim stockSizeExpression As String
    isDynamicStockSize = False
    
    ' Get the model name
    Dim modelName As String
    modelName = model.GetTitle
    DebugLog "Model name: " & modelName
    
    ' Find the cutlist and check StockSize expression
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    ' Get the StockSize property expression
                    bRet = swCustPropMgr.Get4("StockSize", False, stockSizeExpression, resolvedValueOut)
                    
                    If bRet Then
                        DebugLog "StockSize expression: " & stockSizeExpression
                        
                        ' Check if expression contains the model name (indicates dynamic property)
                        If InStr(stockSizeExpression, "@") > 0 And InStr(stockSizeExpression, modelName) > 0 Then
                            isDynamicStockSize = True
                            DebugLog "*** DYNAMIC STOCKSIZE DETECTED (contains model name) ***"
                        Else
                            DebugLog "*** STATIC STOCKSIZE DETECTED (hardcoded value) ***"
                        End If
                    End If
                    Exit Do
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' *** STOCK SIZE PROCESSING WITH CONDITIONAL CONVERSION ***
    DebugLog "=== PROCESSING STOCK SIZE ==="
    
    ' Strip the prefix
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1))
        If Left(cleanStockSize, 1) = "-" Then
            cleanStockSize = Trim(Mid(cleanStockSize, 2))
        End If
    End If
    
    DebugLog "After prefix removal: '" & cleanStockSize & "'"
    
    ' Process Stock Size dimensions
    Dim mecDimensions() As String
    mecDimensions = Split(Replace(cleanStockSize, "X", "x"), "x")

    If UBound(mecDimensions) >= 1 Then
        Dim mecProcessedDimensions() As String
        ReDim mecProcessedDimensions(UBound(mecDimensions))
        
        ' *** CONDITIONAL CONVERSION BASED ON ORIGINAL MODEL UOM AND EXPRESSION CHECK ***
        If isDynamicStockSize And IsMetricUnit(originalModelUOM) Then
            DebugLog "*** DYNAMIC + METRIC: Converting dimensions from " & originalModelUOM & " to inches ***"
            
            Dim mecI As Integer
            For mecI = 0 To UBound(mecDimensions)
                mecDimensions(mecI) = Trim(mecDimensions(mecI))
                DebugLog "Processing dimension " & mecI & ": '" & mecDimensions(mecI) & "'"

                If IsNumeric(mecDimensions(mecI)) Then
                    Dim dimValue As Double
                    dimValue = CDbl(mecDimensions(mecI))
                    DebugLog "Original value (metric): " & dimValue
                    
                    ' Convert from metric to inches using ORIGINAL model UOM
                    dimValue = ConvertModelUnitsToInches(dimValue, originalModelUOM)
                    DebugLog "Converted to inches: " & dimValue
                    
                    ' Clean format
                    If dimValue = Int(dimValue) Then
                        mecProcessedDimensions(mecI) = CStr(Int(dimValue))
                    Else
                        mecProcessedDimensions(mecI) = RemoveTrailingZeros(Format(dimValue, "0.####"))
                    End If
                    DebugLog "Formatted dimension: " & mecProcessedDimensions(mecI)
                Else
                    mecProcessedDimensions(mecI) = mecDimensions(mecI)
                    DebugLog "Non-numeric dimension kept as-is: " & mecProcessedDimensions(mecI)
                End If
            Next mecI
        Else
            DebugLog "*** STATIC OR IMPERIAL: Using dimensions as-is (no conversion) ***"
            
            For mecI = 0 To UBound(mecDimensions)
                mecDimensions(mecI) = Trim(mecDimensions(mecI))
                DebugLog "Processing dimension " & mecI & ": '" & mecDimensions(mecI) & "'"

                If IsNumeric(mecDimensions(mecI)) Then
                    dimValue = CDbl(mecDimensions(mecI))
                    DebugLog "Dimension value (no conversion): " & dimValue
                    
                    ' Clean format
                    If dimValue = Int(dimValue) Then
                        mecProcessedDimensions(mecI) = CStr(Int(dimValue))
                    Else
                        mecProcessedDimensions(mecI) = RemoveTrailingZeros(Format(dimValue, "0.####"))
                    End If
                    DebugLog "Formatted dimension: " & mecProcessedDimensions(mecI)
                Else
                    mecProcessedDimensions(mecI) = mecDimensions(mecI)
                    DebugLog "Non-numeric dimension kept as-is: " & mecProcessedDimensions(mecI)
                End If
            Next mecI
        End If

        ' Set final Stock Size
        stockSize = refCategory & " " & Join(mecProcessedDimensions, " x ")
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Final stockSize: " & stockSize

        ' Set dimensions for swDimensions array
        Dim mecWidth As Double, mecHeight As Double
        If IsNumeric(mecProcessedDimensions(0)) Then mecWidth = CDbl(mecProcessedDimensions(0))
        If IsNumeric(mecProcessedDimensions(1)) Then mecHeight = CDbl(mecProcessedDimensions(1))

        ' Use the total length we found earlier if available
        If foundTotalLength Then
            swDimensions(0) = totalLengthValue
        ElseIf propertiesToSet.exists("Mtl Unit Qty") Then
            swDimensions(0) = CDbl(propertiesToSet("Mtl Unit Qty"))
        End If
        swDimensions(1) = mecWidth
        swDimensions(2) = mecHeight
    Else
        stockSize = "UNKNOWN SIZE"
        propertiesToSet("Stock Size") = stockSize
        DebugLog "ERROR: Could not parse MEC/AS dimensions"
    End If
    
    ' *** HANDLE MTL PART NUMBER MATERIAL REPLACEMENT - USE ACTUAL MATERIAL ***
    DebugLog "=== MEC/AS Material Part Number Processing (Using Actual Material) ==="
    
    Dim originalPartNumber As String
    Dim mecCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set mecCustPropMgr = Nothing
    
    ' Find the MEC/AS cutlist folder and get its Mtl Part Number
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And (resolvedValueOut = refCategory) Then
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set mecCustPropMgr = swCustPropMgr
                        DebugLog "Found Mtl Part Number from " & refCategory & " cutlist: " & originalPartNumber
                        Exit Do
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Update material in part number - USE ACTUAL MATERIAL FROM MODEL
    If originalPartNumber <> "" And Not mecCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String
        
        ' Get actual material from model properties
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If
        
        DebugLog "MEC/AS - Using actual material from model: " & actualMaterial
        
        If Len(actualMaterial) > 0 Then
            Dim partNumberParts() As String
            partNumberParts = Split(originalPartNumber, "-")
            
            If UBound(partNumberParts) >= 2 Then
                partNumberParts(UBound(partNumberParts)) = actualMaterial
            ElseIf UBound(partNumberParts) = 1 Then
                ReDim Preserve partNumberParts(UBound(partNumberParts) + 1)
                partNumberParts(UBound(partNumberParts)) = actualMaterial
            End If
            
            newPartNumber = Join(partNumberParts, "-")
            mecCustPropMgr.Set2 "Mtl Part Number", newPartNumber
            propertiesToSet("Mtl Part Number") = newPartNumber
            DebugLog "Updated Mtl Part Number: " & newPartNumber
        End If
    End If
    
    DebugLog "=== ProcessMECASShape Debug End ==="
End Sub



' Process RB (Round Bar) shapes
Sub ProcessRBShape(ByVal model As ModelDoc2, ByVal isTypeP As Boolean, ByRef propertiesToSet As Object)
    DebugLog "=== ENTERING ProcessRBShape ==="
    
    ' Check if this is a structural member (has RB cutlist)
    Dim hasStructuralMember As Boolean
    Dim stockDiamFromCutlist As Double
    Dim existingMtlPartNumber As String
    Dim totalLength As Double
    Dim grade As String
    
    hasStructuralMember = False
    stockDiamFromCutlist = 0
    existingMtlPartNumber = ""
    totalLength = 0
    grade = ""
    
    ' Look for RB cutlist folder
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            Dim valueOut As String, resolvedValueOut As String
            
            swCustPropMgr.Get4 "Type", False, valueOut, resolvedValueOut
            If resolvedValueOut = "RB" Then
                hasStructuralMember = True
                DebugLog "Found RB structural member cutlist"
                
                ' Get StockSize and parse diameter
                Dim tempStockSize As String
                swCustPropMgr.Get4 "StockSize", False, valueOut, tempStockSize
                If tempStockSize <> "" Then
                    DebugLog "Raw StockSize: " & tempStockSize
                    
                    ' Parse diameter from various formats
                    Dim stockDiamStr As String
                    stockDiamStr = tempStockSize
                    
                    ' Remove diameter symbols (Ø)
                    stockDiamStr = Replace(stockDiamStr, "Ø", "")
                    stockDiamStr = Replace(stockDiamStr, Chr(216), "")  ' Ø (Latin Capital Letter O with Stroke)
                    stockDiamStr = Replace(stockDiamStr, Chr(248), "")  ' ø (Latin Small Letter O with Stroke)
                    
                    ' Remove inch marks and unit indicators
                    stockDiamStr = Replace(stockDiamStr, """", "")  ' Remove double quotes (inch marks)
                    stockDiamStr = Replace(stockDiamStr, "'", "")   ' Remove single quotes
                    stockDiamStr = Replace(stockDiamStr, " in", "") ' Remove " in" suffix
                    stockDiamStr = Replace(stockDiamStr, "in", "")  ' Remove "in" suffix
                    stockDiamStr = Replace(stockDiamStr, " IN", "") ' Remove " IN" suffix
                    stockDiamStr = Replace(stockDiamStr, "IN", "")  ' Remove "IN" suffix

                    ' Remove all types of whitespace (regular space, non-breaking space, etc.)
                    stockDiamStr = Trim(stockDiamStr)
                    stockDiamStr = Replace(stockDiamStr, Chr(160), "")  ' Remove non-breaking space
                    stockDiamStr = Trim(stockDiamStr)  ' Trim again after removing non-breaking spaces

                    DebugLog "After removing diameter symbol and unit marks: '" & stockDiamStr & "'"

                    ' Try different parsing methods
                    If InStr(stockDiamStr, "RB-") > 0 Then
                        ' Complex format: "RB-2-300W" or "D1@Round bar RB-0.6875[1]@cylinder.SLDPRT"
                        Dim startPos As Integer
                        startPos = InStr(stockDiamStr, "RB-") + 3
                        Dim endPos As Integer
                        endPos = InStr(startPos, stockDiamStr, "[")
                        If endPos = 0 Then endPos = InStr(startPos, stockDiamStr, "@")
                        If endPos = 0 Then endPos = InStr(startPos, stockDiamStr, "-")
                        If endPos = 0 Then endPos = Len(stockDiamStr) + 1
                        
                        Dim diameterStr As String
                        diameterStr = Mid(stockDiamStr, startPos, endPos - startPos)
                        diameterStr = Trim(diameterStr)
                        
                        If IsNumeric(diameterStr) Then
                            stockDiamFromCutlist = CDbl(diameterStr)
                            DebugLog "Parsed diameter from RB- pattern: " & stockDiamFromCutlist
                        End If
                    ElseIf IsNumeric(stockDiamStr) Then
                        ' Simple numeric format: "2.000" or "2"
                        stockDiamFromCutlist = CDbl(stockDiamStr)
                        DebugLog "Parsed diameter as simple numeric: " & stockDiamFromCutlist
                    ElseIf InStr(stockDiamStr, "/") > 0 Then
                        ' Fraction format: "2 1/4"
                        stockDiamFromCutlist = ConvertFractionToDecimal(stockDiamStr)
                        DebugLog "Parsed diameter as fraction: " & stockDiamFromCutlist
                    Else
                        DebugLog "Could not parse diameter from: '" & stockDiamStr & "'"
                    End If
                End If
                
                ' Get existing Mtl Part Number
                swCustPropMgr.Get4 "Mtl Part Number", False, valueOut, existingMtlPartNumber
                DebugLog "Cutlist Mtl Part Number: " & existingMtlPartNumber
                
                ' Get Grade
                swCustPropMgr.Get4 "Grade", False, valueOut, grade
                DebugLog "Grade: " & grade
                
                ' Get Length
                swCustPropMgr.Get4 "LENGTH", False, valueOut, resolvedValueOut
                If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                    totalLength = CDbl(resolvedValueOut)
                Else
                    swCustPropMgr.Get4 "TOTAL LENGTH", False, valueOut, resolvedValueOut
                    If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                        totalLength = CDbl(resolvedValueOut)
                    End If
                End If
                DebugLog "Total Length: " & totalLength
                
                Exit Do
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    Dim finalDiameter As Double
    Dim partLength As Double
    
    If hasStructuralMember Then
        DebugLog "=== CASE 1: STRUCTURAL MEMBER (CUTLIST-BASED) ==="
        
        ' Use cutlist diameter and length
        finalDiameter = stockDiamFromCutlist
        partLength = totalLength
        
        ' Convert to inches if needed
        Dim currentDimensionalUOM As String
        If propertiesToSet.exists("Dimensional UOM") Then
            currentDimensionalUOM = propertiesToSet("Dimensional UOM")
        Else
            currentDimensionalUOM = GetModelUnitOfMeasure(model)
        End If
        
        If currentDimensionalUOM <> "in" And IsMetricUnit(currentDimensionalUOM) Then
            DebugLog "Converting from " & currentDimensionalUOM & " to inches"
            finalDiameter = ConvertModelUnitsToInches(finalDiameter, currentDimensionalUOM)
            partLength = ConvertModelUnitsToInches(partLength, currentDimensionalUOM)
        End If
        
        DebugLog "Using cutlist diameter: " & finalDiameter & ", length: " & partLength
        
    Else
        DebugLog "=== CASE 2: MANUAL DIMENSIONS ==="
        
        ' Get dimensions from properties
        Dim actualWidth As Double, actualHeight As Double, actualLength As Double
        actualWidth = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
        actualHeight = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
        actualLength = FractionToDecimal(GetPropertyValue(propertiesToSet, "Length"))
        
        DebugLog "Manual dimensions - Width: " & actualWidth & ", Height: " & actualHeight & ", Length: " & actualLength
        
        ' *** FIXED: Check all possible pairs for matching dimensions ***
        If actualWidth = actualHeight Then
            ' Width and Height are the same (diameter), Length is different
            finalDiameter = actualWidth
            partLength = actualLength
            DebugLog "Width = Height case: diameter=" & finalDiameter & ", length=" & partLength
        ElseIf actualWidth = actualLength Then
            ' Width and Length are the same, Height is different (length)
            finalDiameter = actualWidth
            partLength = actualHeight
            DebugLog "Width = Length case: diameter=" & finalDiameter & ", length=" & partLength
        ElseIf actualHeight = actualLength Then
            ' Height and Length are the same, Width is different (length)
            finalDiameter = actualHeight
            partLength = actualWidth
            DebugLog "Height = Length case: diameter=" & finalDiameter & ", length=" & partLength
        Else
            ' All three dimensions different - use 2nd largest for diameter, largest for length
            Dim dimensions(2) As Double
            dimensions(0) = actualWidth
            dimensions(1) = actualHeight
            dimensions(2) = actualLength
            
            BubbleSort dimensions ' Sorts smallest to largest
            
            finalDiameter = dimensions(1) ' Middle value (2nd largest)
            partLength = dimensions(2)    ' Largest value
            DebugLog "All different case: diameter=" & finalDiameter & " (2nd largest), length=" & partLength & " (largest)"
        End If
        
    End If
    
    ' Set Stock Size
    Dim formattedFinalDiam As String
    formattedFinalDiam = ConvertDecimalToFraction(finalDiameter)
    Dim stockSize As String
    stockSize = "RB " & Chr(216) & formattedFinalDiam
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size: " & stockSize
    
   ' Handle Mtl Part Number
Dim currentType As String
currentType = ""
If propertiesToSet.exists("Type") Then
    currentType = UCase(Trim(CStr(propertiesToSet("Type"))))
End If

If isTypeP Or currentType = "P" Then
    ' Type P - always remove Mtl Part Number
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
        DebugLog "Removed Mtl Part Number (Type P)"
    End If
ElseIf hasStructuralMember And existingMtlPartNumber <> "" Then
    ' Type M with structural member - use existing Mtl Part Number from cutlist
    propertiesToSet("Mtl Part Number") = existingMtlPartNumber
    DebugLog "Using existing Mtl Part Number from cutlist: " & existingMtlPartNumber
Else
    ' Type M without cutlist Mtl Part Number - generate new one
    Dim material As String
    material = GetPropertyValue(propertiesToSet, "Material")
    material = Replace(material, """", "")
    material = Replace(material, "SW-Material@", "")
    material = Replace(material, ".SLDPRT", "")
    
    If material = "" Then
        Set swFeature = model.FirstFeature
        Do While Not swFeature Is Nothing
            If swFeature.GetTypeName2 = "MaterialFolder" Then
                material = Replace(swFeature.Name, ".SLDPRT", "")
                Exit Do
            End If
            Set swFeature = swFeature.GetNextFeature
        Loop
    End If
    If material = "" Then material = "UNKNOWN"
    
    Dim diameterForMtl As String
    diameterForMtl = RemoveTrailingZeros(Format(finalDiameter, "0.#####"))
    
    Dim mtlPartNumber As String
    If hasStructuralMember And grade <> "" Then
        mtlPartNumber = "RB-" & diameterForMtl & "-" & grade
    ElseIf grade <> "" Then
        mtlPartNumber = "RB-" & diameterForMtl & "-" & grade & "-" & material
    Else
        mtlPartNumber = "RB-" & diameterForMtl & "-" & material
    End If
    
    propertiesToSet("Mtl Part Number") = mtlPartNumber
    DebugLog "Generated new RB Mtl Part Number: " & mtlPartNumber
End If
    
    ' Set Width and Height to diameter
    Dim correctDiameter As String
    correctDiameter = ConvertDecimalToFraction(finalDiameter)
    propertiesToSet("Width") = correctDiameter
    propertiesToSet("Height") = correctDiameter
    DebugLog "Set Width and Height to: " & correctDiameter
    
    ' Set Length properties
    If partLength > 0 Then
        propertiesToSet("Length") = ConvertDecimalToFraction(partLength)
        propertiesToSet("LengthA") = RemoveTrailingZeros(Format(partLength, "0.#####"))
        propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
        DebugLog "Set Length: " & propertiesToSet("Length") & ", LengthA: " & propertiesToSet("LengthA")
    End If
    
    DebugLog "=== EXITING ProcessRBShape ==="
End Sub


' Process SQ (Square Bar) shapes
Sub ProcessSQShape(ByVal model As ModelDoc2, ByVal isTypeP As Boolean, ByRef propertiesToSet As Object)
    DebugLog "=== ENTERING ProcessSQShape ==="

    ' Check if this is a structural member (has SQ cutlist)
    Dim hasStructuralMember As Boolean
    Dim stockSizeFromCutlist As Double
    Dim existingMtlPartNumber As String
    Dim totalLength As Double
    Dim grade As String

    hasStructuralMember = False
    stockSizeFromCutlist = 0
    existingMtlPartNumber = ""
    totalLength = 0
    grade = ""

    ' Look for SQ cutlist folder
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            Dim valueOut As String, resolvedValueOut As String

            swCustPropMgr.Get4 "Type", False, valueOut, resolvedValueOut
            If resolvedValueOut = "SQ" Then
                hasStructuralMember = True
                DebugLog "Found SQ structural member cutlist"

                ' Get StockSize and parse size
                Dim tempStockSize As String
                swCustPropMgr.Get4 "StockSize", False, valueOut, tempStockSize
                If tempStockSize <> "" Then
                    DebugLog "Raw StockSize: " & tempStockSize

                    ' Parse size from various formats
                    Dim stockSizeStr As String
                    stockSizeStr = tempStockSize

                    ' Remove unit indicators
                    stockSizeStr = Replace(stockSizeStr, """", "")  ' Remove double quotes (inch marks)
                    stockSizeStr = Replace(stockSizeStr, "'", "")   ' Remove single quotes
                    stockSizeStr = Replace(stockSizeStr, " in", "") ' Remove " in" suffix
                    stockSizeStr = Replace(stockSizeStr, "in", "")  ' Remove "in" suffix
                    stockSizeStr = Replace(stockSizeStr, " IN", "") ' Remove " IN" suffix
                    stockSizeStr = Replace(stockSizeStr, "IN", "")  ' Remove "IN" suffix

                    ' Remove all types of whitespace (regular space, non-breaking space, etc.)
                    stockSizeStr = Trim(stockSizeStr)
                    stockSizeStr = Replace(stockSizeStr, Chr(160), "")  ' Remove non-breaking space
                    stockSizeStr = Trim(stockSizeStr)  ' Trim again after removing non-breaking spaces

                    DebugLog "After removing unit marks: '" & stockSizeStr & "'"

                    ' Try different parsing methods
                    If InStr(stockSizeStr, "SQ-") > 0 Then
                        ' Complex format: "SQ-2-300W" or similar
                        Dim startPos As Integer
                        startPos = InStr(stockSizeStr, "SQ-") + 3
                        Dim endPos As Integer
                        endPos = InStr(startPos, stockSizeStr, "[")
                        If endPos = 0 Then endPos = InStr(startPos, stockSizeStr, "@")
                        If endPos = 0 Then endPos = InStr(startPos, stockSizeStr, "-")
                        If endPos = 0 Then endPos = Len(stockSizeStr) + 1

                        Dim sizeStr As String
                        sizeStr = Mid(stockSizeStr, startPos, endPos - startPos)
                        sizeStr = Trim(sizeStr)

                        If IsNumeric(sizeStr) Then
                            stockSizeFromCutlist = CDbl(sizeStr)
                            DebugLog "Parsed size from SQ- pattern: " & stockSizeFromCutlist
                        End If
                    ElseIf IsNumeric(stockSizeStr) Then
                        ' Simple numeric format: "2.000" or "2"
                        stockSizeFromCutlist = CDbl(stockSizeStr)
                        DebugLog "Parsed size as simple numeric: " & stockSizeFromCutlist
                    ElseIf InStr(stockSizeStr, "/") > 0 Then
                        ' Fraction format: "2 1/4"
                        stockSizeFromCutlist = ConvertFractionToDecimal(stockSizeStr)
                        DebugLog "Parsed size as fraction: " & stockSizeFromCutlist
                    Else
                        DebugLog "Could not parse size from: '" & stockSizeStr & "'"
                    End If
                End If

                ' Get existing Mtl Part Number
                swCustPropMgr.Get4 "Mtl Part Number", False, valueOut, existingMtlPartNumber
                DebugLog "Cutlist Mtl Part Number: " & existingMtlPartNumber

                ' Get Grade
                swCustPropMgr.Get4 "Grade", False, valueOut, grade
                DebugLog "Grade: " & grade

                ' Get Length
                swCustPropMgr.Get4 "LENGTH", False, valueOut, resolvedValueOut
                If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                    totalLength = CDbl(resolvedValueOut)
                Else
                    swCustPropMgr.Get4 "TOTAL LENGTH", False, valueOut, resolvedValueOut
                    If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                        totalLength = CDbl(resolvedValueOut)
                    End If
                End If
                DebugLog "Total Length: " & totalLength

                Exit Do
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop

    Dim finalSize As Double
    Dim partLength As Double

    If hasStructuralMember Then
        DebugLog "=== CASE 1: STRUCTURAL MEMBER (CUTLIST-BASED) ==="

        ' Use cutlist size and length
        finalSize = stockSizeFromCutlist
        partLength = totalLength

        ' Convert to inches if needed
        Dim currentDimensionalUOM As String
        If propertiesToSet.exists("Dimensional UOM") Then
            currentDimensionalUOM = propertiesToSet("Dimensional UOM")
        Else
            currentDimensionalUOM = GetModelUnitOfMeasure(model)
        End If

        If currentDimensionalUOM <> "in" And IsMetricUnit(currentDimensionalUOM) Then
            DebugLog "Converting from " & currentDimensionalUOM & " to inches"
            finalSize = ConvertModelUnitsToInches(finalSize, currentDimensionalUOM)
            partLength = ConvertModelUnitsToInches(partLength, currentDimensionalUOM)
        End If

        DebugLog "Using cutlist size: " & finalSize & ", length: " & partLength

    Else
        DebugLog "=== CASE 2: MANUAL DIMENSIONS ==="

        ' Get dimensions from properties
        Dim actualWidth As Double, actualHeight As Double, actualLength As Double
        actualWidth = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
        actualHeight = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
        actualLength = FractionToDecimal(GetPropertyValue(propertiesToSet, "Length"))

        DebugLog "Manual dimensions - Width: " & actualWidth & ", Height: " & actualHeight & ", Length: " & actualLength

        ' For square bar, width and height should be equal (the square dimension)
        If actualWidth = actualHeight Then
            ' Width and Height are the same (square dimension), Length is different
            finalSize = actualWidth
            partLength = actualLength
            DebugLog "Width = Height case: size=" & finalSize & ", length=" & partLength
        ElseIf actualWidth = actualLength Then
            ' Width and Length are the same, Height is different (length)
            finalSize = actualWidth
            partLength = actualHeight
            DebugLog "Width = Length case: size=" & finalSize & ", length=" & partLength
        ElseIf actualHeight = actualLength Then
            ' Height and Length are the same, Width is different (length)
            finalSize = actualHeight
            partLength = actualWidth
            DebugLog "Height = Length case: size=" & finalSize & ", length=" & partLength
        Else
            ' All three dimensions different - use 2nd largest for size, largest for length
            Dim dimensions(2) As Double
            dimensions(0) = actualWidth
            dimensions(1) = actualHeight
            dimensions(2) = actualLength

            BubbleSort dimensions ' Sorts smallest to largest

            finalSize = dimensions(1) ' Middle value (2nd largest)
            partLength = dimensions(2)    ' Largest value
            DebugLog "All different case: size=" & finalSize & " (2nd largest), length=" & partLength & " (largest)"
        End If

    End If

    ' Set Stock Size
    Dim formattedFinalSize As String
    formattedFinalSize = ConvertDecimalToFraction(finalSize)
    Dim stockSize As String
    stockSize = "SQ " & formattedFinalSize & " x " & formattedFinalSize
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size: " & stockSize

   ' Handle Mtl Part Number
If isTypeP Then
    ' Type P - always remove Mtl Part Number
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
        DebugLog "Removed Mtl Part Number (Type P)"
    End If
ElseIf hasStructuralMember And existingMtlPartNumber <> "" Then
    ' Type M with structural member - use existing Mtl Part Number from cutlist
    propertiesToSet("Mtl Part Number") = existingMtlPartNumber
    DebugLog "Using existing Mtl Part Number from cutlist: " & existingMtlPartNumber
Else
    ' Type M without cutlist Mtl Part Number - generate new one
    Dim material As String
    material = GetPropertyValue(propertiesToSet, "Material")
    material = Replace(material, """", "")
    material = Replace(material, "SW-Material@", "")
    material = Replace(material, ".SLDPRT", "")

    If material = "" Then
        Set swFeature = model.FirstFeature
        Do While Not swFeature Is Nothing
            If swFeature.GetTypeName2 = "MaterialFolder" Then
                material = Replace(swFeature.Name, ".SLDPRT", "")
                Exit Do
            End If
            Set swFeature = swFeature.GetNextFeature
        Loop
    End If
    If material = "" Then material = "UNKNOWN"

    Dim sizeForMtl As String
    sizeForMtl = RemoveTrailingZeros(Format(finalSize, "0.#####"))

    Dim mtlPartNumber As String
    If hasStructuralMember And grade <> "" Then
        mtlPartNumber = "SQ-" & sizeForMtl & "-" & grade
    ElseIf grade <> "" Then
        mtlPartNumber = "SQ-" & sizeForMtl & "-" & grade & "-" & material
    Else
        mtlPartNumber = "SQ-" & sizeForMtl & "-" & material
    End If

    propertiesToSet("Mtl Part Number") = mtlPartNumber
    DebugLog "Generated new SQ Mtl Part Number: " & mtlPartNumber
End If

    ' Set Width and Height to the square size
    Dim correctSize As String
    correctSize = ConvertDecimalToFraction(finalSize)
    propertiesToSet("Width") = correctSize
    propertiesToSet("Height") = correctSize
    DebugLog "Set Width and Height to: " & correctSize

    ' Set Length properties
    If partLength > 0 Then
        propertiesToSet("Length") = ConvertDecimalToFraction(partLength)
        propertiesToSet("LengthA") = RemoveTrailingZeros(Format(partLength, "0.#####"))
        propertiesToSet("Mtl Unit Qty") = propertiesToSet("LengthA")
        DebugLog "Set Length: " & propertiesToSet("Length") & ", LengthA: " & propertiesToSet("LengthA")
    End If

    DebugLog "=== EXITING ProcessSQShape ==="
End Sub


' Process MC, WT, and other misc shapes with proper dimension handling
Sub ProcessMiscShapes(ByVal refCategory As String, ByVal stockSize As String, ByVal isTypeP As Boolean, _
                      ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "=== ProcessMiscShapes Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    DebugLog "modelUOM: " & modelUOM
    
    ' Strip the prefix if it exists
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1))
    End If
    
    ' Remove anything in parentheses
    If InStr(cleanStockSize, "(") > 0 Then
        cleanStockSize = Trim(Left(cleanStockSize, InStr(cleanStockSize, "(") - 1))
    End If
    
    DebugLog "After cleanup: '" & cleanStockSize & "'"
    
    ' Split the dimensions by 'x' for processing
    Dim stockSizeParts() As String
    stockSizeParts = Split(Replace(cleanStockSize, "X", "x"), "x")
    
    ' Process each dimension
    Dim processedDimensions() As String
    ReDim processedDimensions(UBound(stockSizeParts))
    
    ' SIMPLIFIED: No unit conversion needed - dimensions are already in model units
    DebugLog "Processing misc shape dimensions in model units (" & modelUOM & ") - no conversion needed"
    
    Dim i As Integer
    For i = 0 To UBound(stockSizeParts)
        stockSizeParts(i) = Trim(stockSizeParts(i))
        DebugLog "Processing dimension " & i & ": '" & stockSizeParts(i) & "'"
        
        If IsNumeric(stockSizeParts(i)) Then
            Dim dimensionValue As Double
            dimensionValue = CDbl(stockSizeParts(i))
            DebugLog "Dimension value: " & dimensionValue
            
            ' Format cleanly
            If dimensionValue = Int(dimensionValue) Then
                processedDimensions(i) = CStr(Int(dimensionValue))
            Else
                processedDimensions(i) = RemoveTrailingZeros(Format(dimensionValue, "0.0000"))
            End If
            DebugLog "Formatted dimension: " & processedDimensions(i)
        Else
            processedDimensions(i) = stockSizeParts(i)
            DebugLog "Non-numeric dimension kept as-is: " & processedDimensions(i)
        End If
    Next i
    
    ' Construct final stock size
    stockSize = refCategory & " " & Join(processedDimensions, " x ")
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Final Stock Size: " & stockSize
    
    ' Generate Mtl Part Number for Type M parts
    If Not isTypeP Then
        Dim material As String
        material = ""
        
        If propertiesToSet.exists("Material") Then
            material = propertiesToSet("Material")
        ElseIf propertiesToSet.exists("SW-Material") Then
            material = propertiesToSet("SW-Material")
        End If
        
        ' Clean up material name
        If material <> "" Then
            material = Replace(Replace(Replace(material, """", ""), "SW-Material@", ""), ".SLDPRT", "")
            If InStr(material, "SW-Material@") > 0 Then
                material = Mid(material, InStr(material, "@") + 1)
            End If
        Else
            material = "UNKNOWN"
        End If
        
        Dim grade As String
        grade = ""
        If propertiesToSet.exists("Grade") Then
            grade = propertiesToSet("Grade")
        End If
        
        ' Generate Mtl Part Number
        Dim mtlPartNumber As String
        Dim cleanStockSizeForMtl As String
        cleanStockSizeForMtl = Replace(stockSize, refCategory & " ", "")
        
        If grade <> "" Then
            mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & grade & "-" & material
        Else
            mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & material
        End If
        
        propertiesToSet("Mtl Part Number") = mtlPartNumber
        DebugLog "Generated Mtl Part Number: " & mtlPartNumber
    End If
    
    DebugLog "=== ProcessMiscShapes Debug End ==="
End Sub


' Process PI shapes - handles both stock size formatting and dimension swapping
Sub ProcessPIShape(ByVal model As ModelDoc2, ByVal stockSize As String, ByVal isTypeP As Boolean, ByRef propertiesToSet As Object)
    Dim nominalSize As String
    Dim scheduleNumber As String
    Dim stockSizeParts() As String
    Dim finalStockSize As String
    
    DebugLog "=== ProcessPIShape Debug ==="
    DebugLog "Input stockSize: '" & stockSize & "'"
    DebugLog "isTypeP: " & isTypeP
    
    ' Handle different stock size input formats
    If InStr(stockSize, "PIPE") > 0 Then
        ' Extract from weldment member name like "Pipe, 40, 80, 160 PIPE  1.5  SCH 40(15)"
        Dim pipePos As Integer
        pipePos = InStr(stockSize, "PIPE")
        
        If pipePos > 0 Then
            ' Extract everything after "PIPE"
            Dim afterPipe As String
            afterPipe = Trim(Mid(stockSize, pipePos + 4))
            DebugLog "After PIPE extraction: '" & afterPipe & "'"
            
            ' Split on spaces to get nominal size and schedule
            stockSizeParts = Split(afterPipe, " ")
            
            If UBound(stockSizeParts) >= 0 Then
                nominalSize = Trim(stockSizeParts(0))
                
                ' Build schedule from remaining parts
                If UBound(stockSizeParts) >= 1 Then
                    Dim i As Integer
                    scheduleNumber = ""
                    For i = 1 To UBound(stockSizeParts)
                        If Len(Trim(stockSizeParts(i))) > 0 Then
                            If Len(scheduleNumber) > 0 Then
                                scheduleNumber = scheduleNumber & " "
                            End If
                            scheduleNumber = scheduleNumber & Trim(stockSizeParts(i))
                        End If
                    Next i
                    ' Remove trailing parentheses content like "(15)"
                    If InStr(scheduleNumber, "(") > 0 Then
                        scheduleNumber = Trim(Left(scheduleNumber, InStr(scheduleNumber, "(") - 1))
                    End If
                End If
            End If
        End If
    Else
        ' Handle standard space-separated format
        stockSizeParts = Split(stockSize, " ")
        
        If UBound(stockSizeParts) >= 2 Then
            nominalSize = Trim(stockSizeParts(0))
            scheduleNumber = Trim(stockSizeParts(1)) & " " & Trim(stockSizeParts(2))
        ElseIf UBound(stockSizeParts) >= 1 Then
            nominalSize = Trim(stockSizeParts(0))
            scheduleNumber = Trim(stockSizeParts(1))
        Else
            nominalSize = stockSize
            scheduleNumber = ""
        End If
    End If
    
    ' Format the final stock size
    If Len(nominalSize) > 0 Then
        finalStockSize = "PI " & nominalSize
        If Len(scheduleNumber) > 0 Then
            finalStockSize = finalStockSize & " " & scheduleNumber
        End If
    Else
        finalStockSize = "PI Unknown Size"
    End If
    
    DebugLog "Extracted nominalSize: '" & nominalSize & "'"
    DebugLog "Extracted scheduleNumber: '" & scheduleNumber & "'"
    DebugLog "Final stock size: '" & finalStockSize & "'"
    
    propertiesToSet("Stock Size") = finalStockSize
    
    ' Handle Mtl Part Number material replacement
    DebugLog "=== Material Part Number Processing ==="
    
    ' Get Mtl Part Number directly from cutlist properties (not from propertiesToSet)
    Dim originalPartNumber As String
    Dim piCustPropMgr As customPropertyManager ' Keep reference to PI cutlist property manager
    originalPartNumber = ""
    Set piCustPropMgr = Nothing
    
    ' Find the PI cutlist folder and get its Mtl Part Number
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the PI cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And resolvedValueOut = "PI" Then
                    ' Get material from cutlist folder (resolved value)
                    Dim resolvedMaterial As String
                    bRet = swCustPropMgr.Get4("MATERIAL", False, valueOut, resolvedMaterial)
                    If Not bRet Or resolvedMaterial = "" Then
                        bRet = swCustPropMgr.Get4("Material", False, valueOut, resolvedMaterial)
                    End If
                    
                    If bRet And resolvedMaterial <> "" And InStr(resolvedMaterial, "SW-Material@") = 0 Then
                        propertiesToSet("MATERIAL") = resolvedMaterial
                        DebugLog "Found resolved MATERIAL from PI cutlist: " & resolvedMaterial
                    End If
                    
                    ' Get Mtl Part Number from this cutlist
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set piCustPropMgr = swCustPropMgr ' Keep reference for later update
                        DebugLog "Found Mtl Part Number from PI cutlist: " & originalPartNumber
                        Exit Do
                    Else
                        DebugLog "No Mtl Part Number found in PI cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Only proceed if we found an original part number from the cutlist
    If originalPartNumber <> "" And Not piCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String
        
        DebugLog "Original Mtl Part Number from cutlist: '" & originalPartNumber & "'"
        
        ' Get the actual material from the MATERIAL property
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If
        
        DebugLog "Actual material from model: '" & actualMaterial & "'"
        
        ' Clean up material name (remove prefixes, quotes, suffixes)
        If actualMaterial <> "" Then
            actualMaterial = Replace(actualMaterial, """", "")
            actualMaterial = Replace(actualMaterial, "SW-Material@", "")
            actualMaterial = Replace(actualMaterial, ".SLDPRT", "")
        End If
        
        DebugLog "Actual material from model (cleaned): '" & actualMaterial & "'"
        
        ' Only proceed if we have material
        If Len(actualMaterial) > 0 Then
            ' *** CORRECTED: Build part number based on whether material matches grade ***
            DebugLog "Building expected part number from known components..."
            
            ' Get the grade from the cutlist
            Dim grade As String
            grade = ""
            
            If Not piCustPropMgr Is Nothing Then
                Dim gradeOut As String, gradeResolved As String
                bRet = piCustPropMgr.Get4("Grade", False, gradeOut, gradeResolved)
                If bRet And gradeResolved <> "" Then
                    grade = gradeResolved
                    DebugLog "Grade from cutlist: '" & grade & "'"
                End If
            End If
            
            ' Build the expected part number from components
            Dim expectedPartNumber As String
            
            ' Check if material matches grade (case-insensitive comparison)
            If grade <> "" And UCase(Trim(actualMaterial)) = UCase(Trim(grade)) Then
                ' Material matches grade - use grade only
                DebugLog "Material matches grade - using grade only"
                expectedPartNumber = "PI-" & nominalSize
                If scheduleNumber <> "" Then
                    expectedPartNumber = expectedPartNumber & " " & scheduleNumber
                End If
                expectedPartNumber = expectedPartNumber & "-" & grade
            Else
                ' Material is different from grade - use material only (ignore grade)
                DebugLog "Material differs from grade - using material only"
                expectedPartNumber = "PI-" & nominalSize
                If scheduleNumber <> "" Then
                    expectedPartNumber = expectedPartNumber & " " & scheduleNumber
                End If
                expectedPartNumber = expectedPartNumber & "-" & actualMaterial
            End If
            
            DebugLog "Expected part number: '" & expectedPartNumber & "'"
            DebugLog "Original part number: '" & originalPartNumber & "'"
            
            ' Update only if different from original
            If expectedPartNumber <> originalPartNumber Then
                piCustPropMgr.Set2 "Mtl Part Number", expectedPartNumber
                propertiesToSet("Mtl Part Number") = expectedPartNumber
                DebugLog "Updated Mtl Part Number: '" & originalPartNumber & "' -> '" & expectedPartNumber & "'"
            Else
                propertiesToSet("Mtl Part Number") = originalPartNumber
                DebugLog "Mtl Part Number already correct: '" & originalPartNumber & "'"
            End If
            
        Else
            DebugLog "Cannot update part number - missing material"
            propertiesToSet("Mtl Part Number") = originalPartNumber
        End If
    Else
        DebugLog "No Mtl Part Number found in PI cutlist or cutlist property manager not available"
    End If
    
    ' *** CORRECTED: Handle dimension swapping based on TOTAL LENGTH FROM CUTLIST ***
    DebugLog "=== Dimension Swapping Logic ==="
    
    Dim totalLength As Double
    Dim foundTotalLength As Boolean
    
    foundTotalLength = False
    totalLength = 0
    
    ' *** FIXED: Search in the CUTLIST FOLDER for TOTAL LENGTH property ***
    Dim swFeature2 As Feature
    Set swFeature2 = model.FirstFeature
    Do While Not swFeature2 Is Nothing
        If swFeature2.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr2 As customPropertyManager
            Set swCustPropMgr2 = swFeature2.customPropertyManager
            
            If Not swCustPropMgr2 Is Nothing Then
                Dim valueOut2 As String, resolvedValueOut2 As String
                Dim bRet2 As Boolean
                
                ' Check if this is the PI cutlist folder
                bRet2 = swCustPropMgr2.Get4("Type", False, valueOut2, resolvedValueOut2)
                If bRet2 And resolvedValueOut2 = "PI" Then
                    DebugLog "Found PI cutlist folder, searching for TOTAL LENGTH..."
                    
                    ' Try to get TOTAL LENGTH from cutlist
                    bRet2 = swCustPropMgr2.Get4("TOTAL LENGTH", False, valueOut2, resolvedValueOut2)
                    If bRet2 And resolvedValueOut2 <> "" And IsNumeric(resolvedValueOut2) Then
                        totalLength = CDbl(resolvedValueOut2)
                        foundTotalLength = True
                        DebugLog "Found TOTAL LENGTH in cutlist: " & totalLength
                        Exit Do
                    Else
                        ' Try alternate property names
                        bRet2 = swCustPropMgr2.Get4("Total Length", False, valueOut2, resolvedValueOut2)
                        If bRet2 And resolvedValueOut2 <> "" And IsNumeric(resolvedValueOut2) Then
                            totalLength = CDbl(resolvedValueOut2)
                            foundTotalLength = True
                            DebugLog "Found Total Length in cutlist: " & totalLength
                            Exit Do
                        End If
                        
                        bRet2 = swCustPropMgr2.Get4("LENGTH", False, valueOut2, resolvedValueOut2)
                        If bRet2 And resolvedValueOut2 <> "" And IsNumeric(resolvedValueOut2) Then
                            totalLength = CDbl(resolvedValueOut2)
                            foundTotalLength = True
                            DebugLog "Found LENGTH in cutlist: " & totalLength
                            Exit Do
                        End If
                        
                        DebugLog "TOTAL LENGTH property not found or empty in PI cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature2 = swFeature2.GetNextFeature
    Loop
    
    If foundTotalLength Then
        Dim currentLength As Double
        Dim currentWidth As Double
        Dim currentHeight As Double
        Dim tolerance As Double
        
        tolerance = 0.001 ' Small tolerance for floating point comparison
        
        ' Get current dimensions (initialize to 0 if not present)
        currentLength = 0
        currentWidth = 0
        currentHeight = 0
        
        If propertiesToSet.exists("Length") Then currentLength = FractionToDecimal(GetPropertyValue(propertiesToSet, "Length"))
        If propertiesToSet.exists("Width") Then currentWidth = FractionToDecimal(GetPropertyValue(propertiesToSet, "Width"))
        If propertiesToSet.exists("Height") Then currentHeight = FractionToDecimal(GetPropertyValue(propertiesToSet, "Height"))
        
        DebugLog "Current dimensions - Length: " & currentLength & ", Width: " & currentWidth & ", Height: " & currentHeight
        DebugLog "Total length to match: " & totalLength
        
        ' Find which dimension matches the total length
        Dim matchesLength As Boolean
        Dim matchesWidth As Boolean
        Dim matchesHeight As Boolean
        
        matchesLength = Abs(totalLength - currentLength) <= tolerance
        matchesWidth = Abs(totalLength - currentWidth) <= tolerance
        matchesHeight = Abs(totalLength - currentHeight) <= tolerance
        
        DebugLog "Matches - Length: " & matchesLength & ", Width: " & matchesWidth & ", Height: " & matchesHeight
        
        ' Swap dimensions so total length becomes the Length dimension
        If matchesWidth And Not matchesLength Then
            ' Width matches total length, swap width and length
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            If currentLength > 0 Then propertiesToSet("Width") = ConvertDecimalToFraction(currentLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
            DebugLog "Swapped Width -> Length, old Length -> Width"
            
        ElseIf matchesHeight And Not matchesLength Then
            ' Height matches total length, swap height and length
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            If currentLength > 0 Then propertiesToSet("Height") = ConvertDecimalToFraction(currentLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
            DebugLog "Swapped Height -> Length, old Length -> Height"
            
        ElseIf Not matchesLength And (currentLength = 0 Or currentLength <= tolerance) Then
            ' No length dimension exists or is very small, set total length as length
            propertiesToSet("Length") = ConvertDecimalToFraction(totalLength)
            propertiesToSet("LengthA") = RemoveTrailingZeros(Format(totalLength, "0.#####"))
            DebugLog "Set total length as Length dimension"
            
        Else
            DebugLog "No dimension swap needed - total length already matches Length or no clear match"
        End If
    Else
        DebugLog "No total length property found in PI cutlist"
    End If
    
    DebugLog "=== End ProcessPIShape ==="
End Sub


' Process other valid types
Sub ProcessOtherValidTypes(ByVal typeValue As String, ByVal stockSize As String, ByRef propertiesToSet As Object)
    Dim dimensionCount As Integer
    Dim stockDimensions() As String
    
    Dim dimensionStr As String
    dimensionStr = Replace(Replace(stockSize, "X", "x"), " ", "")
    Dim dimensionParts() As String
    dimensionParts = Split(dimensionStr, "x")
    dimensionCount = UBound(dimensionParts) + 1
    
    ReDim stockDimensions(dimensionCount - 1)
    Dim i As Integer
    
    For i = 0 To dimensionCount - 1
        Dim tempValue As String
        tempValue = Trim(dimensionParts(i))
        If IsNumeric(tempValue) Then
            stockDimensions(i) = FormatNumber(CDbl(tempValue), 3, True, False, True)
        Else
            stockDimensions(i) = tempValue
        End If
    Next i
    
    stockSize = typeValue & " " & Join(stockDimensions, " x ")
    propertiesToSet("Stock Size") = stockSize
End Sub

' Process non-structural members
Sub ProcessNonStructuralMember(ByVal refCategory As String, ByRef propertiesToSet As Object)
    Dim width As Double, height As Double
    Dim widthStr As String, heightStr As String
    Dim modelUOM As String
    
    widthStr = GetPropertyValue(propertiesToSet, "Width")
    heightStr = GetPropertyValue(propertiesToSet, "Height")
    
    If widthStr = "" Or heightStr = "" Then
        Exit Sub
    End If
    
    ' Get the model's unit of measure
    modelUOM = GetPropertyValue(propertiesToSet, "Dimensional UOM")
    
    ' Convert dimensions to double, handling both fractional and decimal inputs
    width = FractionToDecimal(widthStr)
    height = FractionToDecimal(heightStr)
    
    ' MODIFICATION: Convert to inches if the model is in metric units
    If IsMetricUnit(modelUOM) Then
        width = ConvertModelUnitsToInches(width, modelUOM)
        height = ConvertModelUnitsToInches(height, modelUOM)
    End If
    
    ' Sort the dimensions (smaller first)
    If width > height Then
        Dim temp As Double
        temp = width
        width = height
        height = temp
    End If
    
    ' Format the stock size string using imperial dimensions
    Dim stockSize As String
    ' Always use imperial formatting now, regardless of original units
    Dim widthFormatted As String, heightFormatted As String
    
    widthFormatted = ConvertDecimalToFraction(width)
    heightFormatted = ConvertDecimalToFraction(height)
    
    stockSize = refCategory & " " & widthFormatted & " x " & heightFormatted
    
    propertiesToSet("Stock Size") = stockSize
End Sub




' Process UB (Universal Beam) shapes with simplified handling
Sub ProcessUBShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal stockSize As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "=== ProcessUBShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "refCategory: " & refCategory
    DebugLog "modelUOM: " & modelUOM
    
    ' Strip the prefix if it exists
    Dim cleanStockSize As String
    cleanStockSize = Trim(stockSize)
    If Left(UCase(cleanStockSize), Len(refCategory)) = UCase(refCategory) Then
        cleanStockSize = Trim(Mid(cleanStockSize, Len(refCategory) + 1))
        ' Handle dash format like "UB-"
        If Left(cleanStockSize, 1) = "-" Then
            cleanStockSize = Trim(Mid(cleanStockSize, 2))
        End If
    End If
    
    DebugLog "After prefix removal: '" & cleanStockSize & "'"
    
    ' Remove unwanted suffixes that might appear in UB descriptions
    Dim suffixesToRemove As Variant
    suffixesToRemove = Array("BEAM", "beam", "mm", "MM")
    
    Dim i As Integer
    For i = 0 To UBound(suffixesToRemove)
        If InStr(UCase(cleanStockSize), UCase(suffixesToRemove(i))) > 0 Then
            cleanStockSize = Replace(UCase(cleanStockSize), UCase(suffixesToRemove(i)), "", 1, 1, vbTextCompare)
            cleanStockSize = Trim(cleanStockSize)
            DebugLog "Removed suffix '" & suffixesToRemove(i) & "': " & cleanStockSize
        End If
    Next i
    
    ' Clean up extra spaces
    Do While InStr(cleanStockSize, "  ") > 0
        cleanStockSize = Replace(cleanStockSize, "  ", " ")
    Loop
    cleanStockSize = Trim(cleanStockSize)
    
    DebugLog "After suffix removal and cleanup: '" & cleanStockSize & "'"
    
    ' Extract only the outer dimensions, ignoring anything in parentheses
    Dim ubOuterDimensions As String
    If InStr(cleanStockSize, "(") > 0 Then
        ubOuterDimensions = Trim(Left(cleanStockSize, InStr(cleanStockSize, "(") - 1))
    Else
        ubOuterDimensions = cleanStockSize
    End If
    
    DebugLog "Outer dimensions: " & ubOuterDimensions
    
    ' Split the dimensions - handle both "x" and space separators
    Dim ubParts() As String
    If InStr(ubOuterDimensions, "x") > 0 Or InStr(ubOuterDimensions, "X") > 0 Then
        ubParts = Split(Replace(ubOuterDimensions, "X", "x"), "x")
    Else
        ' Handle space-separated format like "180 22.2"
        ubParts = Split(ubOuterDimensions, " ")
    End If
    
    ' Process each dimension
    Dim ubProcessedDimensions() As String
    ReDim ubProcessedDimensions(UBound(ubParts))
    
    ' SIMPLIFIED: No unit conversion needed - dimensions are already in model units
    DebugLog "Processing UB dimensions in model units (" & modelUOM & ") - no conversion needed"
    
    For i = 0 To UBound(ubParts)
        ubParts(i) = Trim(ubParts(i))
        DebugLog "Processing dimension " & i & ": '" & ubParts(i) & "'"
        
        If IsNumeric(ubParts(i)) Then
            Dim ubDimension As Double
            ubDimension = CDbl(ubParts(i))
            DebugLog "Dimension value: " & ubDimension
            
            ' Format the dimension value cleanly
            If ubDimension = Int(ubDimension) Then
                ubProcessedDimensions(i) = CStr(Int(ubDimension))
            Else
                ubProcessedDimensions(i) = RemoveTrailingZeros(Format(ubDimension, "0.#####"))
            End If
            DebugLog "Formatted dimension: " & ubProcessedDimensions(i)
        Else
            ubProcessedDimensions(i) = ubParts(i)
            DebugLog "Non-numeric dimension kept as-is: " & ubProcessedDimensions(i)
        End If
    Next i
    
    ' Construct the final Stock Size string
    stockSize = refCategory & " " & Join(ubProcessedDimensions, " x ")
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Final Stock Size: " & stockSize
    
    ' *** MTL PART NUMBER PROCESSING WITH MATERIAL REPLACEMENT ***
    DebugLog "=== UB Material Part Number Processing ==="
    
    ' Get Mtl Part Number directly from cutlist properties
    Dim originalPartNumber As String
    Dim ubCustPropMgr As customPropertyManager
    originalPartNumber = ""
    Set ubCustPropMgr = Nothing
    
    ' Find the UB cutlist folder and get its Mtl Part Number
    Dim swFeature As Feature
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Dim swCustPropMgr As customPropertyManager
            Set swCustPropMgr = swFeature.customPropertyManager
            
            If Not swCustPropMgr Is Nothing Then
                Dim valueOut As String, resolvedValueOut As String
                Dim bRet As Boolean
                
                ' Check if this is the UB cutlist folder
                bRet = swCustPropMgr.Get4("Type", False, valueOut, resolvedValueOut)
                If bRet And resolvedValueOut = refCategory Then
                    ' Get Mtl Part Number from this cutlist
                    bRet = swCustPropMgr.Get4("Mtl Part Number", False, valueOut, originalPartNumber)
                    If bRet And originalPartNumber <> "" Then
                        Set ubCustPropMgr = swCustPropMgr
                        DebugLog "Found Mtl Part Number from " & refCategory & " cutlist: " & originalPartNumber
                        Exit Do
                    Else
                        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist"
                    End If
                End If
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    ' Update material in part number if needed
    If originalPartNumber <> "" And Not ubCustPropMgr Is Nothing Then
        Dim newPartNumber As String
        Dim actualMaterial As String

        DebugLog "Original Mtl Part Number from cutlist: '" & originalPartNumber & "'"

        ' Get the actual material from the model properties
        actualMaterial = ""
        If propertiesToSet.exists("MATERIAL") Then
            actualMaterial = Trim(propertiesToSet("MATERIAL"))
        ElseIf propertiesToSet.exists("Material") Then
            actualMaterial = Trim(propertiesToSet("Material"))
        End If

        ' Clean up material name (remove prefixes, quotes, suffixes)
        If actualMaterial <> "" Then
            actualMaterial = Replace(actualMaterial, """", "")
            actualMaterial = Replace(actualMaterial, "SW-Material@", "")
            actualMaterial = Replace(actualMaterial, ".SLDPRT", "")
        End If

        DebugLog "Actual material from model (cleaned): '" & actualMaterial & "'"

        ' Only proceed if we have material
        If Len(actualMaterial) > 0 Then
            ' IMPROVED: Always rebuild to clean duplicates and ensure correct format
            DebugLog "Analyzing part number structure..."

            Dim partNumberParts() As String
            partNumberParts = Split(originalPartNumber, "-")

            DebugLog "Part number has " & (UBound(partNumberParts) + 1) & " parts:"
            Dim j As Integer
            For j = 0 To UBound(partNumberParts)
                DebugLog "  Part " & j & ": '" & partNumberParts(j) & "'"
            Next j

            If UBound(partNumberParts) >= 2 Then
                ' Always rebuild with Type-Size-Material to clean up any duplicates
                newPartNumber = partNumberParts(0) & "-" & partNumberParts(1) & "-" & actualMaterial
                DebugLog "Rebuilt part number: Type '" & partNumberParts(0) & "' + Size '" & partNumberParts(1) & "' + Material '" & actualMaterial & "'"
            ElseIf UBound(partNumberParts) = 1 Then
                ' Only Type-Size, add material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Added material to Type-Size format"
            Else
                ' Unexpected format, append material
                newPartNumber = originalPartNumber & "-" & actualMaterial
                DebugLog "Unexpected format - appending material"
            End If
            
            ' Update properties only if the part number changed
            If newPartNumber <> originalPartNumber Then
                ubCustPropMgr.Set2 "Mtl Part Number", newPartNumber
                propertiesToSet("Mtl Part Number") = newPartNumber
                DebugLog "Updated Mtl Part Number: '" & originalPartNumber & "' -> '" & newPartNumber & "'"
            Else
                propertiesToSet("Mtl Part Number") = originalPartNumber
                DebugLog "Kept original Mtl Part Number: '" & originalPartNumber & "'"
            End If
            
        Else
            DebugLog "Cannot update part number - missing material"
            propertiesToSet("Mtl Part Number") = originalPartNumber
        End If
    Else
        DebugLog "No Mtl Part Number found in " & refCategory & " cutlist - generating new one"
        
        ' Generate new Mtl Part Number for Type M if none exists in cutlist
        If Not isTypeP Then
            DebugLog "Generating new Mtl Part Number for Type M"
            
            Dim material As String
            material = GetPropertyValue(propertiesToSet, "Material")
            
            ' Clean up material name
            If material <> "" Then
                material = Replace(Replace(Replace(material, """", ""), "SW-Material@", ""), ".SLDPRT", "")
            Else
                material = "UNKNOWN"
            End If
            
            ' Get grade
            Dim grade As String
            grade = GetPropertyValue(propertiesToSet, "Grade")
            
            ' Generate Mtl Part Number using clean stock size
            Dim mtlPartNumber As String
            Dim cleanStockSizeForMtl As String
            cleanStockSizeForMtl = Replace(stockSize, refCategory & " ", "") ' Remove prefix
            cleanStockSizeForMtl = Replace(cleanStockSizeForMtl, " ", "") ' Remove spaces
            
            If grade <> "" Then
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & grade & "-" & material
            Else
                mtlPartNumber = refCategory & "-" & cleanStockSizeForMtl & "-" & material
            End If
            
            propertiesToSet("Mtl Part Number") = mtlPartNumber
            DebugLog "Generated new Mtl Part Number: " & mtlPartNumber
        ElseIf isTypeP Then
            ' Remove Mtl Part Number for Type P
            If propertiesToSet.exists("Mtl Part Number") Then
                propertiesToSet.Remove "Mtl Part Number"
                DebugLog "Removed Mtl Part Number (Type P)"
            End If
        End If
    End If
    
    DebugLog "=== ProcessUBShape Debug End ==="
End Sub



' Process HOSE shapes
Sub ProcessHOSEShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, _
                     ByVal modelUOM As String, ByRef propertiesToSet As Object)
    
    DebugLog "=== ProcessHOSEShape Debug Start ==="
    DebugLog "refCategory: " & refCategory
    DebugLog "modelUOM: " & modelUOM
    
    On Error GoTo ErrorHandler
    
    ' Get HOSE specific properties from cut list
    Dim swFeature As Feature
    Dim swCustPropMgr As customPropertyManager
    Dim valueOut As String, resolvedValueOut As String
    Dim stockSize As String
    Dim grade As String
    Dim totalLength As Double
    Dim cutlistMtlPartNumber As String
    
    stockSize = ""
    grade = ""
    totalLength = 0
    cutlistMtlPartNumber = ""
    
    ' Find the CutListFolder for HOSE
    Set swFeature = model.FirstFeature
    Do While Not swFeature Is Nothing
        If swFeature.GetTypeName2 = "CutListFolder" Then
            Set swCustPropMgr = swFeature.customPropertyManager
            
            ' Check if this is the HOSE cut list folder
            swCustPropMgr.Get4 "Type", False, valueOut, resolvedValueOut
            If resolvedValueOut = "HOSE" Then
                
                ' Get StockSize (should be something like "T4008D")
                swCustPropMgr.Get4 "StockSize", False, valueOut, stockSize
                DebugLog "StockSize from cut list: " & stockSize
                
                ' Get Grade
                swCustPropMgr.Get4 "Grade", False, valueOut, grade
                DebugLog "Grade from cut list: " & grade
                
                ' Get existing Mtl Part Number from cut list
                swCustPropMgr.Get4 "Mtl Part Number", False, valueOut, resolvedValueOut
                If resolvedValueOut <> "" Then
                    cutlistMtlPartNumber = resolvedValueOut
                Else
                    cutlistMtlPartNumber = valueOut
                End If
                DebugLog "Mtl Part Number from cut list: '" & cutlistMtlPartNumber & "'"
                
                ' Get Length or Total Length
                DebugLog "=== Attempting to get LENGTH ==="
                swCustPropMgr.Get4 "LENGTH", False, valueOut, resolvedValueOut
                DebugLog "LENGTH valueOut: '" & valueOut & "', resolvedValueOut: '" & resolvedValueOut & "'"
                
                If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                    totalLength = CDbl(resolvedValueOut)
                    DebugLog "Length from cut list: " & totalLength
                Else
                    DebugLog "=== Attempting to get TOTAL LENGTH ==="
                    swCustPropMgr.Get4 "TOTAL LENGTH", False, valueOut, resolvedValueOut
                    DebugLog "TOTAL LENGTH valueOut: '" & valueOut & "', resolvedValueOut: '" & resolvedValueOut & "'"
                    
                    If resolvedValueOut <> "" And IsNumeric(resolvedValueOut) Then
                        totalLength = CDbl(resolvedValueOut)
                        DebugLog "Total Length from cut list: " & totalLength
                    Else
                        DebugLog "No valid length found in cut list"
                        totalLength = 0
                    End If
                End If
                
                Exit Do
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop
    
    DebugLog "=== Setting Stock Size ==="
    ' Set Stock Size using the StockSize from cut list
    If stockSize <> "" Then
        ' For HOSE, keep stock size and ensure HOSE prefix
        If Left(UCase(Trim(stockSize)), 4) <> "HOSE" Then
            stockSize = "HOSE " & stockSize
        End If
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Final Stock Size: " & stockSize
    Else
        stockSize = "HOSE UNKNOWN SIZE"
        propertiesToSet("Stock Size") = stockSize
        DebugLog "No StockSize found, using default: " & stockSize
    End If
    
    DebugLog "=== Setting Material ==="
    ' Set Material to the raw StockSize value (for example "T4008D")
    If stockSize <> "" Then
        Dim materialValue As String
        materialValue = Replace(stockSize, "HOSE ", "")
        propertiesToSet("Material") = materialValue
        DebugLog "Set Material to: " & materialValue
    End If
    
    DebugLog "=== Processing Length Properties ==="
    ' Set Length properties if we found length data
    If totalLength > 0 Then
        DebugLog "Processing totalLength: " & totalLength
        
        ' Convert length to inches if model is metric
        If IsMetricUnit(modelUOM) Then
            DebugLog "Converting from metric to inches"
            totalLength = ConvertModelUnitsToInches(totalLength, modelUOM)
            DebugLog "Converted totalLength: " & totalLength
        End If
        
        ' Create both fraction and decimal versions
        Dim lengthFraction As String
        Dim lengthDecimal As String
        
        DebugLog "Converting to fraction and decimal formats"
        lengthFraction = ConvertDecimalToFraction(totalLength)
        lengthDecimal = RemoveTrailingZeros(Format(totalLength, "0.#####"))
        
        propertiesToSet("Length") = lengthFraction
        propertiesToSet("LengthA") = lengthDecimal
        propertiesToSet("Mtl Unit Qty") = lengthDecimal
        
        DebugLog "Set Length (fraction): " & lengthFraction
        DebugLog "Set LengthA (decimal): " & lengthDecimal
        DebugLog "Set Mtl Unit Qty (decimal): " & lengthDecimal
    Else
        DebugLog "No totalLength to process (totalLength = " & totalLength & ")"
    End If
    
    DebugLog "=== Processing Mtl Part Number ==="
    ' Copy Mtl Part Number from HOSE cut list folder to document level for Type M
    If Not isTypeP Then
        If cutlistMtlPartNumber <> "" Then
            propertiesToSet("Mtl Part Number") = cutlistMtlPartNumber
            DebugLog "Copied Mtl Part Number from cut list: " & cutlistMtlPartNumber
        Else
            DebugLog "No Mtl Part Number found in HOSE cut list folder"
        End If
    Else
        ' For Type P, remove Mtl Part Number at document level if present
        If propertiesToSet.exists("Mtl Part Number") Then
            propertiesToSet.Remove "Mtl Part Number"
            DebugLog "Removed Mtl Part Number (Type P)"
        End If
    End If
    
    DebugLog "=== ProcessHOSEShape Debug End SUCCESS ==="
    Exit Sub
    
ErrorHandler:
    DebugLog "ERROR in ProcessHOSEShape: " & Err.description & " (Number: " & Err.Number & ")"
    DebugLog "=== ProcessHOSEShape Debug End ERROR ==="
 
End Sub




' ============================
' SHIM processing with gauge in Stock Size
' Self contained and guarded
' ============================
Sub ProcessSHIMShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    On Error GoTo SHIM_Fatal
    DebugLog "#########################"
    DebugLog "### ENTERING ProcessSHIMShape FUNCTION ###"
    DebugLog "#########################"
    
    propertiesToSet("Reference Category") = refCategory
    DebugLog "Set Reference Category to: " & refCategory
    
    ' Get original decimal dimensions directly from model to avoid fraction rounding
    Dim originalDimensions As Variant
    originalDimensions = GetDimensions(model, propertiesToSet)
    
    Dim width As Double, height As Double, length As Double
    
    If Not IsEmpty(originalDimensions) Then
        ' Use original decimal dimensions
        length = originalDimensions(0)   ' Largest
        width = originalDimensions(1)    ' Smallest
        height = originalDimensions(2)   ' Middle
        DebugLog "Using original decimal dimensions from GetDimensions"
        DebugLog "Original Length: " & Format(length, "0.########")
        DebugLog "Original Width: " & Format(width, "0.########")
        DebugLog "Original Height: " & Format(height, "0.########")
    Else
        ' Fallback to property values if GetDimensions fails
        DebugLog "Fallback: Using property values"
        Dim widthStr As String, heightStr As String, lengthStr As String
        widthStr = GetPropertyValue(propertiesToSet, "Width")
        heightStr = GetPropertyValue(propertiesToSet, "Height")
        lengthStr = GetPropertyValue(propertiesToSet, "Length")
        
        width = FractionToDecimal(widthStr)
        height = FractionToDecimal(heightStr)
        length = FractionToDecimal(lengthStr)
    End If

    Dim actualUOM As String
    If propertiesToSet.exists("Dimensional UOM") Then
        actualUOM = propertiesToSet("Dimensional UOM")
    Else
        actualUOM = modelUOM
    End If
    DebugLog "Using UOM: " & actualUOM
    
    ' Convert to inches if needed (though GetDimensions should already return inches)
    If IsMetricUnit(actualUOM) Then
        DebugLog "Converting SHIM to inches"
        width = ConvertModelUnitsToInches(width, actualUOM)
        height = ConvertModelUnitsToInches(height, actualUOM)
        length = ConvertModelUnitsToInches(length, actualUOM)
        DebugLog "Converted to inches: w=" & Format(width, "0.########") & ", h=" & Format(height, "0.########") & ", l=" & Format(length, "0.########")
    End If

    ' *** NEW: Add CheckPlateDimensions call like PL/CP parts ***
    DebugLog "=== CALLING CheckPlateDimensions for SHIM ==="
    If Not CheckPlateDimensions(model, width, height, length, actualUOM, propertiesToSet) Then
        DebugLog "CheckPlateDimensions returned False - SHIM set to Purchase"
    Else
        DebugLog "CheckPlateDimensions returned True - SHIM can be Manufactured"
    End If

    ' *** Check if SHIM dimensions match flat bar specifications ***
    DebugLog "=== CALLING CheckAndConvertToFlatBar for SHIM ==="
    Dim shimFlatBarResult As Boolean
    shimFlatBarResult = FlatBarModule.CheckAndConvertToFlatBar(model, width, height, length, actualUOM, propertiesToSet)
    If shimFlatBarResult Then
        DebugLog "SHIM converted to FB (Flat Bar)"
        ' Update Reference Category to FB for SHIM
        refCategory = "FB"
        DebugLog "Updated SHIM refCategory to: FB"
    Else
        DebugLog "SHIM does not match flat bar specifications"
    End If

    ' Sort all dimensions to properly identify thickness (smallest) and the two larger ones
    Dim sortedDims(2) As Double
    sortedDims(0) = width
    sortedDims(1) = height
    sortedDims(2) = length
    BubbleSort sortedDims  ' Now: sortedDims(0) = smallest (thickness), sortedDims(1) = middle, sortedDims(2) = largest
    
    Dim thickness As Double
    thickness = sortedDims(0)  ' Smallest dimension is thickness
    
    DebugLog "All dimensions: width=" & Format(width, "0.########") & ", height=" & Format(height, "0.########") & ", length=" & Format(length, "0.########")
    DebugLog "Sorted dimensions: thickness=" & Format(sortedDims(0), "0.########") & ", middle=" & Format(sortedDims(1), "0.########") & ", largest=" & Format(sortedDims(2), "0.########")
    DebugLog "Smallest dimension (thickness) for gauge: " & Format(thickness, "0.########")

    ' Material cleanup
    Dim materialName As String
    If propertiesToSet.exists("Material") Then
        materialName = CStr(propertiesToSet("Material"))
        materialName = Replace(materialName, """", "")
        If InStr(materialName, "SW-Material@") > 0 Then materialName = Mid$(materialName, InStr(materialName, "@") + 1)
        If InStr(1, materialName, ".SLDPRT", vbTextCompare) > 0 Then materialName = Left$(materialName, InStr(1, materialName, ".SLDPRT", vbTextCompare) - 1)
    End If
    DebugLog "Material cleaned: " & materialName
    
    ' Gauge mapping with explicit debug
    Dim gaugeGroup As String
    gaugeGroup = MapMaterialToGaugeGroupUsingLists(materialName)
    DebugLog "Gauge group: " & gaugeGroup
    
    Dim gaugeNum As Long
    Dim gaugeText As String
    On Error GoTo GAUGE_Err
    DebugLog "About to call gauge lookup with thickness: " & thickness
    gaugeNum = GetGaugeNumberFromThickness(thickness, gaugeGroup)
    On Error GoTo SHIM_Fatal
    
    If gaugeNum > 0 Then
        gaugeText = CStr(gaugeNum) & "GA"
        propertiesToSet("Gauge Number") = CStr(gaugeNum)
        DebugLog "Gauge lookup success. Gauge Number=" & gaugeNum
    Else
        DebugLog "Gauge lookup returned 0. Falling back to decimal thickness"
        gaugeText = RemoveTrailingZeros(Format(thickness, "0.#####")) & " in"
    End If

    ' *** FIXED STOCK SIZE LOGIC ***
    ' For SHIM stock size, use the middle dimension (not the largest)
    ' This represents the reasonable sheet width you'd purchase
    Dim stockWidth As Double
    stockWidth = sortedDims(1)  ' Use middle dimension (3" in your case, not 21")
    
    ' Cap the stock width to reasonable sheet material sizes
    Dim maxStockWidth As Double
    maxStockWidth = 48#   ' 48 inches maximum for standard sheet stock
    
    If stockWidth > maxStockWidth Then
        DebugLog "Stock width " & stockWidth & " exceeds maximum " & maxStockWidth & ", capping to maximum"
        stockWidth = maxStockWidth
    End If
    
    'Dim heightStr As String
    heightStr = ConvertDecimalToFraction(stockWidth)
    DebugLog "Stock width for stock size (middle dimension): " & stockWidth
    DebugLog "Stock width formatted as fraction: " & heightStr
    
    ' Build Stock Size using gauge when available
    Dim stockSize As String
    stockSize = "SH " & gaugeText & " x " & heightStr
    
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size: " & stockSize
    
    ' *** UPDATED: Set default Type to M if not already set, then process based on current Type ***
    If Not propertiesToSet.exists("Type") Then
        propertiesToSet("Type") = "M"
        DebugLog "Set default Type to: M"
    End If
    
    Dim currentType As String
    currentType = UCase(Trim(propertiesToSet("Type")))
    DebugLog "Current Type: " & currentType
    
    If currentType = "M" Then
        ' Generate manufacturing properties for Type M - use PLT prefix like plates
        Dim mtlPartNumber As String
        Dim thicknessForMtl As String

        ' Use gauge number for Mtl Part Number (if available), otherwise fall back to decimal
        If gaugeNum > 0 Then
            thicknessForMtl = CStr(gaugeNum) & "GA"
        Else
            thicknessForMtl = RemoveTrailingZeros(Format(thickness, "0.#####"))
        End If
        mtlPartNumber = "PLT-" & thicknessForMtl & "-" & materialName
        propertiesToSet("Mtl Part Number") = mtlPartNumber
        
        ' Calculate material quantity (area = two larger dimensions, excluding thickness)
        Dim materialQuantity As Double
        materialQuantity = sortedDims(1) * sortedDims(2)  ' Middle x Largest dimensions
        propertiesToSet("Mtl Unit Qty") = RemoveTrailingZeros(Format(materialQuantity, "0.0000"))

        DebugLog "Generated Mtl Part Number: " & mtlPartNumber
        DebugLog "Set Mtl Unit Qty: " & materialQuantity & " (calculated from " & sortedDims(1) & " x " & sortedDims(2) & ")"
        
    ElseIf currentType = "P" Then
        ' Remove manufacturing properties for purchase parts
        If propertiesToSet.exists("Mtl Part Number") Then propertiesToSet.Remove "Mtl Part Number"
        If propertiesToSet.exists("Mtl Unit Qty") Then propertiesToSet.Remove "Mtl Unit Qty"
        DebugLog "Removed manufacturing properties (Type P)"
    End If

    ' Final verification - ensure critical properties are set
    propertiesToSet("Reference Category") = refCategory
    DebugLog "=== ProcessSHIMShape COMPLETED ==="
    DebugLog "Reference Category: " & propertiesToSet("Reference Category")
    If propertiesToSet.exists("Type") Then DebugLog "Type: " & propertiesToSet("Type")
    DebugLog "Stock Size: " & propertiesToSet("Stock Size")
    If propertiesToSet.exists("Gauge Number") Then DebugLog "Gauge Number: " & propertiesToSet("Gauge Number")
    If propertiesToSet.exists("Mtl Part Number") Then DebugLog "Mtl Part Number: " & propertiesToSet("Mtl Part Number")
    If propertiesToSet.exists("Mtl Unit Qty") Then DebugLog "Mtl Unit Qty: " & propertiesToSet("Mtl Unit Qty")
    DebugLog "#########################"
    Exit Sub

GAUGE_Err:
    DebugLog "Gauge lookup error: " & Err.Number & " - " & Err.description
    DebugLog "Falling back to decimal thickness for Stock Size"
    Err.Clear
    On Error GoTo SHIM_Fatal
    gaugeText = RemoveTrailingZeros(Format(thickness, "0.#####")) & " in"
    
    heightStr = ConvertDecimalToFraction(stockWidth)
    stockSize = "SH " & gaugeText & " x " & heightStr
    propertiesToSet("Stock Size") = stockSize
    Resume Next

SHIM_Fatal:
    If Err.Number <> 0 Then
        DebugLog "ProcessSHIMShape fatal error: " & Err.Number & " - " & Err.description
    End If
End Sub


' Map your material name to the sheet gauge table group using built in lists
' Returns Standard Steel, Stainless Steel, or Aluminum
Private Function MapMaterialToGaugeGroupUsingLists(ByVal materialName As String) As String
    Dim m As String
    m = Trim$(materialName)
    Dim ml As String
    ml = LCase$(m)
    
    ' Quick keyword checks
    If InStr(ml, "alum") > 0 Or InStr(ml, "6061") > 0 Or InStr(ml, "6063") > 0 Or InStr(ml, "5052") > 0 Or InStr(ml, "5083") > 0 Or InStr(ml, "3003") > 0 Then
        MapMaterialToGaugeGroupUsingLists = "Aluminum"
        Exit Function
    End If
    If InStr(ml, "stainless") > 0 Or InStr(ml, "304") > 0 Or InStr(ml, "316") > 0 Or InStr(ml, "308") > 0 Or InStr(ml, "309") > 0 Or ml = "ss" Then
        MapMaterialToGaugeGroupUsingLists = "Stainless Steel"
        Exit Function
    End If
    
    ' Default to standard steel
    MapMaterialToGaugeGroupUsingLists = "Standard Steel"
End Function

' Return closest gauge number for a thickness in inches for the selected group
' materialGroup must be one of Standard Steel, Galvanized Steel, Stainless Steel, Aluminum
Private Function GetGaugeNumberFromThickness(ByVal thicknessIn As Double, ByVal materialGroup As String) As Long
    ' Use Variant for arrays that we assign at runtime to avoid Type mismatch
    Dim g As Variant
    g = Array(3, 4, 5, 6, 7, 8, _
              9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
              21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, _
              33, 34, 35, 36, 37, 38)
    
    Dim stdIn As Variant
    stdIn = Array(0.2391, 0.2242, 0.2092, 0.1943, 0.1793, 0.1644, _
                  0.1495, 0.1345, 0.1196, 0.1046, 0.0897, 0.0747, 0.0673, 0.0598, 0.0538, 0.0478, 0.0418, 0.0359, _
                  0.0329, 0.0299, 0.0269, 0.0239, 0.0209, 0.0179, 0.0164, 0.0149, 0.0135, 0.012, 0.0105, 0.0097, _
                  0.009, 0.0082, 0.0075, 0.0067, 0.0064, 0.0067)
    
    Dim galvIn As Variant
    galvIn = Array(Empty, Empty, Empty, Empty, Empty, Empty, _
                   0.1532, 0.1382, 0.1233, 0.1084, 0.0934, 0.0785, 0.071, 0.0635, 0.0575, 0.0516, 0.0456, 0.0396, _
                   0.0366, 0.0336, 0.0306, 0.0276, 0.0247, 0.0217, 0.0202, 0.0187, 0.0172, 0.0157, 0.0142, 0.0134, _
                   0.0094, 0.0086, 0.0078, 0.007, 0.0066, Empty)
    
    Dim ssIn As Variant
    ssIn = Array(Empty, 0.2344, 0.2187, 0.2031, 0.1875, 0.165, _
                 0.1562, 0.1406, 0.125, 0.1094, 0.0937, 0.0781, 0.0703, 0.0625, 0.0562, 0.05, 0.0437, 0.0375, _
                 0.0344, 0.0312, 0.0281, 0.025, 0.0219, 0.0187, 0.0172, 0.0156, 0.0141, 0.0125, 0.0109, 0.0102, _
                 Empty, Empty, Empty, Empty, Empty, 0.0062)
    
    Dim aluIn As Variant
    aluIn = Array(0.2294, 0.2043, 0.1819, 0.162, 0.1443, 0.1285, _
                  0.1144, 0.1019, 0.0907, 0.0808, 0.072, 0.0641, 0.0571, 0.0508, 0.0453, 0.0403, 0.0359, 0.032, _
                  0.0285, 0.0253, 0.0226, 0.0211, 0.0179, 0.0159, 0.0142, 0.0126, 0.0113, 0.01, 0.0089, 0.008, _
                  0.0071, 0.0063, 0.0056, 0.005, 0.00445, 0.00396)
    
    Dim col As Variant
    Select Case materialGroup
        Case "Galvanized Steel": col = galvIn
        Case "Stainless Steel":  col = ssIn
        Case "Aluminum":         col = aluIn
        Case Else:               col = stdIn
    End Select
    
    Dim bestIdx As Long
    Dim bestDiff As Double
    bestDiff = 1000000000#
    bestIdx = -1
    
    Dim i As Long
    Dim v As Double
    Dim d As Double
    
    For i = LBound(g) To UBound(g)
        If Not IsEmpty(col(i)) Then
            v = CDbl(col(i))
            d = Abs(v - thicknessIn)
            If d < bestDiff Then
                bestDiff = d
                bestIdx = i
            End If
        End If
    Next i
    
    If bestIdx >= 0 Then
        GetGaugeNumberFromThickness = CLng(g(bestIdx))
        DebugLog "Gauge candidate picked: Gauge=" & g(bestIdx) & " Value=" & col(bestIdx) & " in, diff=" & Format(bestDiff, "0.########")
    Else
        GetGaugeNumberFromThickness = 0
        DebugLog "No gauge match found"
    End If
End Function





' Process PUR (Purchase) shapes
Sub ProcessPURShape(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "#########################"
    DebugLog "### ENTERING ProcessPURShape FUNCTION ###"
    DebugLog "#########################"
    
    ' ENSURE Reference Category is preserved from the start
    propertiesToSet("Reference Category") = refCategory
    DebugLog "Set Reference Category to: " & refCategory
    
    ' *** Set Type to "P" for all PUR parts ***
    propertiesToSet("Type") = "P"
    DebugLog "Set Type to: P (Purchase part)"
    
    ' *** Set Stock Size to "Per Drawing" for PUR parts ***
    propertiesToSet("Stock Size") = "AS PER DRAWING"

    DebugLog "Set Stock Size to: Per Drawing"
    
    ' *** Remove Material properties for PUR parts ***
    If propertiesToSet.exists("Material") Then
        propertiesToSet.Remove "Material"
        DebugLog "Removed Material property"
    End If
    
    If propertiesToSet.exists("SW-Material") Then
        propertiesToSet.Remove "SW-Material"
        DebugLog "Removed SW-Material property"
    End If
    
    ' Remove dimensional properties since they're not relevant for "Per Drawing" items
    If propertiesToSet.exists("Length") Then
        propertiesToSet.Remove "Length"
        DebugLog "Removed Length (Per Drawing)"
    End If
    
    If propertiesToSet.exists("LengthA") Then
        propertiesToSet.Remove "LengthA"
        DebugLog "Removed LengthA (Per Drawing)"
    End If
    
    If propertiesToSet.exists("Width") Then
        propertiesToSet.Remove "Width"
        DebugLog "Removed Width (Per Drawing)"
    End If
    
    If propertiesToSet.exists("Height") Then
        propertiesToSet.Remove "Height"
        DebugLog "Removed Height (Per Drawing)"
    End If
    
    ' Remove manufacturing properties
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
        DebugLog "Removed Mtl Part Number (Type P)"
    End If
    
    If propertiesToSet.exists("Mtl Unit Qty") Then
        propertiesToSet.Remove "Mtl Unit Qty"
        DebugLog "Removed Mtl Unit Qty (Type P)"
    End If

    ' Final verification - ensure critical properties are set
    propertiesToSet("Reference Category") = refCategory
    propertiesToSet("Type") = "P"
   ' propertiesToSet("Stock Size") = "Per Drawing"
    
    DebugLog "Final verification - Reference Category: " & propertiesToSet("Reference Category")
    DebugLog "Final verification - Type: " & propertiesToSet("Type")
    DebugLog "Final verification - Stock Size: " & propertiesToSet("Stock Size")
    
    DebugLog "=== ProcessPURShape COMPLETED ==="
    DebugLog "#########################"
End Sub

' Process the KS reference category
Sub ProcessKS(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "#########################"
    DebugLog "### ENTERING ProcessKS FUNCTION ###"
    DebugLog "#########################"
    DebugLog "refCategory: " & refCategory
    DebugLog "isTypeP: " & isTypeP
    DebugLog "modelUOM: " & modelUOM
    
    Dim widthStr As String, heightStr As String, lengthStr As String
    Dim width As Double, height As Double, length As Double
    Dim KSthickness As String
    Dim KSthicknessDecimal As Double
    Dim KSMaterial As String
   
    ' Get dimensions from properties
    DebugLog "=== GETTING DIMENSIONS FROM PROPERTIES ==="
    widthStr = GetPropertyValue(propertiesToSet, "Width")
    heightStr = GetPropertyValue(propertiesToSet, "Height")
    lengthStr = GetPropertyValue(propertiesToSet, "Length")
    
    DebugLog "Raw property values:"
    DebugLog "widthStr from properties: '" & widthStr & "'"
    DebugLog "heightStr from properties: '" & heightStr & "'"
    DebugLog "lengthStr from properties: '" & lengthStr & "'"
    DebugLog "modelUOM: " & modelUOM
    
    ' Get the UPDATED Dimensional UOM from properties (AddDimensionsProperties may have changed it)
    Dim actualUOM As String
    If propertiesToSet.exists("Dimensional UOM") Then
        actualUOM = propertiesToSet("Dimensional UOM")
        DebugLog "Using updated Dimensional UOM from properties: " & actualUOM
    Else
        actualUOM = modelUOM
        DebugLog "Using original model UOM: " & actualUOM
    End If
    
    ' Convert strings to numbers
    DebugLog "=== CONVERTING TO NUMBERS ==="
    width = FractionToDecimal(widthStr)
    height = FractionToDecimal(heightStr)
    length = FractionToDecimal(lengthStr)
    
    DebugLog "After FractionToDecimal conversion:"
    DebugLog "width: " & width
    DebugLog "height: " & height
    DebugLog "length: " & length
    DebugLog "IsMetricUnit(actualUOM): " & IsMetricUnit(actualUOM)
    
    ' Convert to inches if the ACTUAL current unit is metric
    If IsMetricUnit(actualUOM) Then
        DebugLog "=== CONVERTING FROM METRIC TO INCHES ==="
        DebugLog "Before conversion:"
        DebugLog "width: " & width & " " & actualUOM
        DebugLog "height: " & height & " " & actualUOM
        DebugLog "length: " & length & " " & actualUOM
        
        width = ConvertModelUnitsToInches(width, actualUOM)
        height = ConvertModelUnitsToInches(height, actualUOM)
        length = ConvertModelUnitsToInches(length, actualUOM)
        
        DebugLog "After conversion to inches:"
        DebugLog "width: " & width & " in"
        DebugLog "height: " & height & " in"
        DebugLog "length: " & length & " in"
    Else
        DebugLog "=== NO METRIC CONVERSION NEEDED ==="
        DebugLog "Dimensions are already in inches (actualUOM: " & actualUOM & ")"
    End If
    
    ' Add plate size check here (temporarily commented out due to parameter mismatch)
    'DebugLog "=== CALLING CheckPlateDimensions ==="
    ' If Not CheckPlateDimensions(model, width, height, length, actualUOM, propertiesToSet) Then
    '     DebugLog "CheckPlateDimensions returned False - user notified about oversized plate dimensions"
    ' Else
    '    DebugLog "CheckPlateDimensions skipped - assuming dimensions are acceptable"
    ' End If
    
    ' *** NEW: Check if Type was changed to P and clean up properties ***
    DebugLog "=== CHECKING TYPE PROPERTY ==="
    Dim currentType As String
    If propertiesToSet.exists("Type") Then
        currentType = UCase(Trim(propertiesToSet("Type")))
    Else
        currentType = ""
    End If
    DebugLog "currentType: '" & currentType & "'"
    
    ' If Type is now P, remove manufacturing-related properties and exit
    If currentType = "P" Then
        DebugLog "=== TYPE IS P - PURCHASE PART PROCESSING ==="
        If propertiesToSet.exists("Mtl Part Number") Then
            propertiesToSet.Remove "Mtl Part Number"
            DebugLog "Removed Mtl Part Number"
        End If
        If propertiesToSet.exists("Mtl Unit Qty") Then
            propertiesToSet.Remove "Mtl Unit Qty"
            DebugLog "Removed Mtl Unit Qty"
        End If
        
        ' Ensure width is the smaller dimension for stock size
        DebugLog "Before dimension swap - width: " & width & ", height: " & height
        If width > height Then
            Dim temp As Double
            temp = width
            width = height
            height = temp
            DebugLog "Swapped dimensions - width: " & width & ", height: " & height
        End If
        
        ' Set just the Stock Size for Purchase parts
        widthStr = ConvertDecimalToFraction(width)
        heightStr = ConvertDecimalToFraction(height)
        Dim stockSize As String
        stockSize = refCategory & " " & widthStr & " x " & heightStr
        propertiesToSet("Stock Size") = stockSize
        DebugLog "Set Stock Size for Purchase part: " & stockSize
        DebugLog "=== EXITING ProcessKS - TYPE P ==="
        Exit Sub
    End If
    
    ' *** CONTINUE WITH MANUFACTURING LOGIC ONLY IF NOT TYPE P ***
    DebugLog "=== TYPE IS M - MANUFACTURING PART PROCESSING ==="
    
    ' Calculate material quantity (area = length * height) BEFORE dimension swapping
    Dim materialQuantity As Double
    materialQuantity = length * height
    propertiesToSet("Mtl Unit Qty") = RemoveTrailingZeros(Format(materialQuantity, "0.0000"))
    DebugLog "Material Quantity (length x height): " & length & " x " & height & " = " & materialQuantity
    DebugLog "Set Mtl Unit Qty: " & propertiesToSet("Mtl Unit Qty")
    
    ' Ensure width is the smaller dimension for stock size
    DebugLog "=== DIMENSION SWAPPING FOR STOCK SIZE ==="
    DebugLog "Before swap - width: " & width & ", height: " & height
    If width > height Then
        temp = width
        width = height
        height = temp
        DebugLog "Swapped dimensions - width: " & width & ", height: " & height
    Else
        DebugLog "No swap needed - width is already smaller"
    End If
    
    ' --- Update Stock Size to Use the Rounded-Up Width ---
    Dim roundedWidth As Double
    roundedWidth = width  ' Use original width value instead
    DebugLog "roundedWidth: " & roundedWidth
    
    ' IMPORTANT FIX: Convert the dimensions to fractions AFTER unit conversion
    DebugLog "=== CONVERTING TO FRACTIONS ==="
    widthStr = ConvertDecimalToFraction(width)
    heightStr = ConvertDecimalToFraction(height)
    DebugLog "widthStr fraction: " & widthStr
    DebugLog "heightStr fraction: " & heightStr
    
    stockSize = refCategory & " " & widthStr & " x " & heightStr
    DebugLog "Initial Stock Size: " & stockSize
    
    ' --- Material Handling Section ---
    DebugLog "=== MATERIAL HANDLING ==="
    Dim materialFound As Boolean
    materialFound = False
    KSMaterial = ""
    
    ' First try getting from properties
    If propertiesToSet.exists("Material") Then
        KSMaterial = propertiesToSet("Material")
        DebugLog "Material from 'Material' property: " & KSMaterial
    End If
    
    ' If empty, try SW-Material
    If KSMaterial = "" Then
        If propertiesToSet.exists("SW-Material") Then
            KSMaterial = propertiesToSet("SW-Material")
            DebugLog "Material from 'SW-Material' property: " & KSMaterial
        End If
    End If
    
    ' If still empty, try getting from material feature
    If KSMaterial = "" Then
        DebugLog "Searching for MaterialFolder feature..."
        Dim swFeat As Feature
        Set swFeat = model.FirstFeature
        
        Do While Not swFeat Is Nothing
            If swFeat.GetTypeName2 = "MaterialFolder" Then
                KSMaterial = swFeat.Name
                materialFound = True
                DebugLog "Found MaterialFolder: " & KSMaterial
                Exit Do
            End If
            Set swFeat = swFeat.GetNextFeature
        Loop
    End If
    
    DebugLog "Final KSMaterial: " & KSMaterial
    
    Dim KSMtlNumb As String
    If KSMaterial <> "" Then
        DebugLog "=== MATERIAL PROCESSING ==="
        ' Remove quotes if they exist
        KSMaterial = Replace(KSMaterial, """", "")
        DebugLog "After removing quotes: " & KSMaterial
        
        ' Clean up the SW-Material@ prefix if it exists
        If InStr(KSMaterial, "SW-Material@") > 0 Then
            KSMaterial = Mid(KSMaterial, InStr(KSMaterial, "@") + 1)
            DebugLog "After removing SW-Material@ prefix: " & KSMaterial
        End If
        
        ' Remove .SLDPRT if it exists
        If InStr(KSMaterial, ".SLDPRT") > 0 Then
            KSMaterial = Left(KSMaterial, InStr(KSMaterial, ".SLDPRT") - 1)
            DebugLog "After removing .SLDPRT: " & KSMaterial
        End If
        
        Dim stockThickness As Double
        stockThickness = ConvertFractionToDecimal(widthStr)  ' Use the original formatted width
        KSthicknessDecimal = stockThickness
        KSthickness = RemoveTrailingZeros(Format(KSthicknessDecimal, "0.0000"))
        
        DebugLog "stockThickness: " & stockThickness
        DebugLog "KSthickness: " & KSthickness
        
        ' KS-specific material part number format
        KSMtlNumb = "KS-" & KSthickness & "-" & KSMaterial
        DebugLog "Generated KSMtlNumb: " & KSMtlNumb
    Else
        DebugLog "=== NO MATERIAL FOUND - USING UNKNOWN ==="
        KSthickness = Format(width, "0.0000")
        KSMtlNumb = refCategory & "-" & KSthickness & "-UNKNOWN"
        DebugLog "Generated KSMtlNumb with UNKNOWN: " & KSMtlNumb
    End If
    
    ' Set the final properties (conditionally based on painted flag)
    DebugLog "=== SETTING FINAL PROPERTIES ==="
    If Not isTypeP Then
        propertiesToSet("Mtl Part Number") = KSMtlNumb
        DebugLog "Set Mtl Part Number: " & KSMtlNumb
    Else
        If propertiesToSet.exists("Mtl Part Number") Then
            propertiesToSet.Remove "Mtl Part Number"
            DebugLog "Removed Mtl Part Number (Type P)"
        End If
    End If
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size: " & stockSize
    
    DebugLog "=== ProcessKS COMPLETED ==="
    DebugLog "Final Results:"
    DebugLog "Stock Size: " & stockSize
    DebugLog "Mtl Part Number: " & KSMtlNumb
    DebugLog "#########################"
End Sub

' Process S shapes with simplified handling
Sub ProcessSShape(ByVal model As ModelDoc2, ByVal stockSize As String, ByVal isTypeP As Boolean, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "=== ProcessSShape Debug Start ==="
    DebugLog "Input stockSize: " & stockSize
    DebugLog "modelUOM: " & modelUOM
    
    Dim finalStockSize As String
    
    ' Check if we have actual dimensional data in stockSize
    If stockSize = "UNKNOWN SIZE" Or stockSize = "" Or (InStr(stockSize, "x") = 0 And InStr(stockSize, "X") = 0) Then
        ' Get original Mtl Part Number directly from cutlist
        DebugLog "Getting original Mtl Part Number from cutlist"
        
        Dim originalMtlPartNumber As String
        originalMtlPartNumber = GetOriginalMtlPartNumberFromCutlist(model)
        DebugLog "Original Mtl Part Number from cutlist: '" & originalMtlPartNumber & "'"
        
        If originalMtlPartNumber <> "" And Left(UCase(originalMtlPartNumber), 2) = "S-" Then
            ' Parse original Mtl Part Number format: "S-12x40.8-350W"
            Dim withoutPrefix As String
            withoutPrefix = Mid(originalMtlPartNumber, 3)
            
            Dim lastDashPos As Integer
            lastDashPos = InStrRev(withoutPrefix, "-")
            
            Dim dimensionsPart As String
            If lastDashPos > 0 Then
                dimensionsPart = Left(withoutPrefix, lastDashPos - 1)
            Else
                dimensionsPart = withoutPrefix
            End If
            
            ' Parse dimensions - NO CONVERSION, use as-is from cutlist
            If InStr(dimensionsPart, "x") > 0 Or InStr(dimensionsPart, "X") > 0 Then
                Dim dimParts() As String
                dimParts = Split(Replace(dimensionsPart, "X", "x"), "x")
                
                If UBound(dimParts) >= 1 Then
                    Dim dim1 As String, dim2 As String
                    dim1 = Trim(dimParts(0))
                    dim2 = Trim(dimParts(1))
                    
                    If IsNumeric(dim1) And IsNumeric(dim2) Then
                        Dim dim1Val As Double, dim2Val As Double
                        dim1Val = CDbl(dim1)
                        dim2Val = CDbl(dim2)
                        
                        DebugLog "NO CONVERSION - using values as-is from cutlist"
                        
                        ' Format cleanly
                        If dim1Val = Int(dim1Val) Then
                            dim1 = CStr(Int(dim1Val))
                        Else
                            dim1 = CStr(dim1Val)
                        End If
                        
                        If dim2Val = Int(dim2Val) Then
                            dim2 = CStr(Int(dim2Val))
                        Else
                            dim2 = CStr(dim2Val)
                        End If
                        
                        finalStockSize = "S " & dim1 & " x " & dim2
                        DebugLog "Generated StockSize from original Mtl Part Number: " & finalStockSize
                    Else
                        finalStockSize = "S UNKNOWN SIZE"
                    End If
                Else
                    finalStockSize = "S UNKNOWN SIZE"
                End If
            Else
                finalStockSize = "S UNKNOWN SIZE"
            End If
        Else
            ' Fallback to calculated properties
            Dim width As String, height As String
            width = GetPropertyValue(propertiesToSet, "Width")
            height = GetPropertyValue(propertiesToSet, "Height")
            
            If width <> "" And height <> "" Then
                finalStockSize = "S " & height & " x " & width
            Else
                finalStockSize = "S UNKNOWN SIZE"
            End If
        End If
    Else
        ' Process existing dimensional data with simplified approach
        DebugLog "Processing existing dimensional data from cutlist"
        
        ' Remove the "S" prefix
        Dim cleanStockSize As String
        cleanStockSize = Trim(stockSize)
        If Left(UCase(cleanStockSize), 2) = "S " Then
            cleanStockSize = Trim(Mid(cleanStockSize, 3))
        ElseIf Left(UCase(cleanStockSize), 1) = "S" Then
            cleanStockSize = Trim(Mid(cleanStockSize, 2))
        End If
        
        ' Split by "x" and process dimensions
        Dim parts() As String
        parts = Split(Replace(cleanStockSize, "X", "x"), "x")
        
        If UBound(parts) >= 1 Then
            Dim sDimensions() As Double
            ReDim sDimensions(UBound(parts))
            Dim i As Integer
            Dim success As Boolean
            success = True
            
            For i = 0 To UBound(parts)
                parts(i) = Trim(parts(i))
                If IsNumeric(parts(i)) Then
                    sDimensions(i) = CDbl(parts(i))
                Else
                    success = False
                    Exit For
                End If
            Next i
            
            If success Then
                ' SIMPLIFIED: No unit conversion needed - dimensions are already in model units
                DebugLog "Processing S dimensions in model units (" & modelUOM & ") - no conversion needed"
                
                ' Format dimensions cleanly
                Dim formattedDimensions() As String
                ReDim formattedDimensions(UBound(sDimensions))
                For i = 0 To UBound(sDimensions)
                    If sDimensions(i) = Int(sDimensions(i)) Then
                        formattedDimensions(i) = CStr(Int(sDimensions(i)))
                    Else
                        formattedDimensions(i) = RemoveTrailingZeros(Format(sDimensions(i), "0.#####"))
                    End If
                Next i
                
                finalStockSize = "S " & Join(formattedDimensions, " x ")
                DebugLog "Final Stock Size from cutlist: " & finalStockSize
            Else
                finalStockSize = "S UNKNOWN SIZE"
            End If
        Else
            finalStockSize = "S UNKNOWN SIZE"
        End If
    End If
    
    ' Set the final stock size
    propertiesToSet("Stock Size") = finalStockSize
    DebugLog "Set Stock Size: " & finalStockSize
    
    ' Handle Mtl Part Number (keep existing logic)
    If Not isTypeP Then
        Dim existingMtlPartNumber As String
        existingMtlPartNumber = ""
        
        Dim swFeature As Feature
        Set swFeature = model.FirstFeature
        Do While Not swFeature Is Nothing
            If swFeature.GetTypeName2 = "CutListFolder" Then
                Dim swCustPropMgr As customPropertyManager
                Set swCustPropMgr = swFeature.customPropertyManager
                Dim valueOut As String, resolvedValueOut As String
                
                swCustPropMgr.Get4 "Type", False, valueOut, resolvedValueOut
                If resolvedValueOut = "S" Then
                    swCustPropMgr.Get4 "Mtl Part Number", False, valueOut, existingMtlPartNumber
                    If existingMtlPartNumber <> "" Then
                        propertiesToSet("Mtl Part Number") = existingMtlPartNumber
                        DebugLog "Using existing Mtl Part Number from cutlist"
                        Exit Do
                    End If
                End If
            End If
            Set swFeature = swFeature.GetNextFeature
        Loop
        
        ' Generate new if none found
        If existingMtlPartNumber = "" Then
            Dim material As String, grade As String, mtlPartNumber As String
            material = GetPropertyValue(propertiesToSet, "Material")
            If material <> "" Then
                material = Replace(Replace(Replace(material, """", ""), "SW-Material@", ""), ".SLDPRT", "")
            Else
                material = "UNKNOWN"
            End If
            
            grade = GetPropertyValue(propertiesToSet, "Grade")
            Dim cleanStockSizeForMtl As String
            cleanStockSizeForMtl = Replace(finalStockSize, "S ", "")
            
            If grade <> "" Then
                mtlPartNumber = "S-" & cleanStockSizeForMtl & "-" & grade & "-" & material
            Else
                mtlPartNumber = "S-" & cleanStockSizeForMtl & "-" & material
            End If
            
            propertiesToSet("Mtl Part Number") = mtlPartNumber
            DebugLog "Generated new Mtl Part Number: " & mtlPartNumber
        End If
    End If
    
    DebugLog "=== ProcessSShape Debug End ==="
End Sub


' Helper function to get the original Mtl Part Number directly from cutlist (bypassing processed properties)

Sub ApplyMaterialFromGrade(ByVal model As ModelDoc2, ByVal grade As String)
    DebugLog "=== ApplyMaterialFromGrade ==="
    DebugLog "Grade from cutlist: " & grade
    
    If model Is Nothing Or grade = "" Then Exit Sub
    
    ' Check if this is a part document
    If model.GetType <> swDocPART Then
        DebugLog "Not a part document - skipping material application"
        Exit Sub
    End If
    
    ' Get SolidWorks application and part document
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    Dim swPart As SldWorks.PartDoc
    Set swPart = model
    
    ' Path to material library - using centralized constant
    Dim materialPath As String
    materialPath = PDM_MATERIAL_DB_PATH
    
    ' Get the active configuration name
    Dim activeConfig As String
    activeConfig = model.ConfigurationManager.ActiveConfiguration.Name
    DebugLog "Active configuration: " & activeConfig
    
    ' Get current material if any
    Dim currentMaterial As String
    Dim currentDatabase As String
    On Error Resume Next
    currentMaterial = swPart.GetMaterialPropertyName2(activeConfig, currentDatabase)
    On Error GoTo ErrorHandler

    ' Trim both materials for comparison to avoid whitespace issues
    Dim trimmedCurrent As String
    Dim trimmedGrade As String
    trimmedCurrent = Trim(currentMaterial)
    trimmedGrade = Trim(grade)

    ' Check if there's a material mismatch and prompt user
    If trimmedCurrent <> "" And StrComp(trimmedCurrent, trimmedGrade, vbTextCompare) <> 0 Then
        ' Check if document has a saved preference
        Dim savedPreference As String
        savedPreference = GetMaterialOverridePreference(model)

        Dim keepCurrentMaterial As Boolean
        If savedPreference <> "" Then
            ' Use the saved preference
            keepCurrentMaterial = (savedPreference = "USE_MODEL")
            DebugLog "Using saved preference: " & savedPreference
            If keepCurrentMaterial Then
                DebugLog "Document preference: Keep current material '" & currentMaterial & "'"
                Exit Sub
            Else
                DebugLog "Document preference: Change to material '" & grade & "'"
            End If
        Else
            ' No saved preference - ask the user and save the decision
            Dim userResponse As VbMsgBoxResult
            userResponse = MsgBox("Material mismatch detected!" & vbCrLf & vbCrLf & _
                                 "Current material: " & currentMaterial & vbCrLf & _
                                 "Weldment material: " & grade & vbCrLf & vbCrLf & _
                                 "Do you want to keep the current material?" & vbCrLf & vbCrLf & _
                                 "Yes = Always use model material for this file" & vbCrLf & _
                                 "No = Always use cutlist material for this file" & vbCrLf & vbCrLf & _
                                 "This preference will be saved with the document.", _
                                 vbYesNo + vbQuestion + vbDefaultButton2, _
                                 "Material Confirmation")

            keepCurrentMaterial = (userResponse = vbYes)

            ' Save the preference
            If keepCurrentMaterial Then
                SetMaterialOverridePreference model, "USE_MODEL"
                DebugLog "User chose to always use model material (saved to document)"
                Exit Sub
            Else
                SetMaterialOverridePreference model, "USE_CUTLIST"
                DebugLog "User chose to always use cutlist material (saved to document)"
            End If
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Apply the grade
    DebugLog "Attempting to apply grade: '" & grade & "'"
    swPart.SetMaterialPropertyName2 activeConfig, materialPath, grade
    
    ' Verify if grade was applied successfully
    On Error Resume Next
    Dim appliedMaterial As String
    Dim appliedDatabase As String
    appliedMaterial = swPart.GetMaterialPropertyName2(activeConfig, appliedDatabase)
    
    If Err.Number = 0 And StrComp(appliedMaterial, grade, vbTextCompare) = 0 Then
        DebugLog "SUCCESS: Applied material - '" & appliedMaterial & "'"
    Else
        DebugLog "FAILED: Could not apply grade '" & grade & "'"
        If appliedMaterial <> "" Then
            DebugLog "Current material remains: '" & appliedMaterial & "'"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    DebugLog "ERROR in ApplyMaterialFromGrade: " & Err.Number & " - " & Err.description
End Sub










Public Sub DeleteWeldment_OnlyIfNoMembers(model As ModelDoc2)
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim f As SldWorks.Feature
    Dim sf As SldWorks.Feature
    Dim nextF As SldWorks.Feature
    Dim hasMembers As Boolean
    Dim delResult As Long
    
    On Error GoTo EH
    
' Use the passed model parameter
Set swModel = model
If swModel Is Nothing Then
    DebugLog "DeleteWeldment_OnlyIfNoMembers: model parameter is Nothing"
    Exit Sub
End If
    
    
    
    If swModel.GetType <> swDocPART Then
        DebugLog "DeleteWeldment_OnlyIfNoMembers: active document is not a Part"
        Exit Sub
    End If
    
    DebugLog "=== DeleteWeldment_OnlyIfNoMembers ==="
    DebugLog "Model: " & swModel.GetTitle
    
    ' Pass 1: detect any Structural Member features anywhere in the tree
    hasMembers = False ' Initialize explicitly
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing And Not hasMembers ' Added hasMembers check here
        ' Debug: Print every feature we're checking
        DebugLog "Checking feature: " & f.Name & " (Type: " & f.GetTypeName2 & ")"
        
        ' check top level
        Select Case LCase$(f.GetTypeName2)
            Case "weldmentstructuralmember", "structuralmember", "weldmentmembergroup", "weldmemberfeat"
                DebugLog "Found structural member (top level): " & f.Name & " (" & f.GetTypeName2 & ")"
                hasMembers = True
                Exit Do ' This will now properly exit the main loop
        End Select
        
        ' check subfeatures under this feature only if we haven't found members yet
        If Not hasMembers Then
            Set sf = f.GetFirstSubFeature
            Do While Not sf Is Nothing
                DebugLog "  Checking subfeature: " & sf.Name & " (Type: " & sf.GetTypeName2 & ")"
                Select Case LCase$(sf.GetTypeName2)
                    Case "weldmentstructuralmember", "structuralmember", "weldmentmembergroup", "weldmemberfeat"
                        DebugLog "Found structural member (subfeature): " & sf.Name & " (" & sf.GetTypeName2 & ")"
                        hasMembers = True
                        Exit Do ' Exit the subfeature loop
                End Select
                Set sf = sf.GetNextSubFeature
            Loop
        End If
        
        Set f = f.GetNextFeature
    Loop
    
    If hasMembers Then
        DebugLog "Structural member present: Weldment kept"
        Exit Sub
    End If
    
    DebugLog "No structural members found: Proceeding to delete Weldment"
    
    ' Pass 2: find and delete the Weldment feature only
    delResult = 0 ' Initialize delResult
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing
        Set nextF = f.GetNextFeature
        
        If LCase$(f.GetTypeName2) = "weldmentfeature" _
           Or LCase$(f.GetTypeName2) = "weldment" _
           Or StrComp(f.Name, "Weldment", vbTextCompare) = 0 Then
           
            DebugLog "Found Weldment feature to delete: " & f.Name & " (" & f.GetTypeName2 & ")"
            swModel.ClearSelection2 True
            If f.Select2(False, 0) Then
                delResult = swModel.Extension.DeleteSelection2( _
                    swDeleteSelectionOptions_e.swDelete_Absorbed + _
                    swDeleteSelectionOptions_e.swDelete_Children)
                If delResult = 1 Then
                    DebugLog "Successfully deleted Weldment feature"
                Else
                    DebugLog "Failed to delete Weldment, result: " & delResult
                End If
            Else
                DebugLog "Failed to select Weldment feature"
            End If
            Exit Do
        End If
        
        Set f = nextF
    Loop
    
    If delResult <> 1 Then
        DebugLog "No Weldment feature found or deletion did not succeed"
    End If
    Exit Sub
EH:
    DebugLog "DeleteWeldment_OnlyIfNoMembers error: " & Err.Number & " - " & Err.description
End Sub









































' Process WD reference category
Sub ProcessWDShape(ByVal refCategory As String, ByVal isTypeP As Boolean, ByRef propertiesToSet As Object)
    Dim widthStr As String, heightStr As String, lengthStr As String
    Dim width As Double, height As Double, length As Double
   
    ' Get dimensions from properties
    widthStr = GetPropertyValue(propertiesToSet, "Width")
    heightStr = GetPropertyValue(propertiesToSet, "Height")
    lengthStr = GetPropertyValue(propertiesToSet, "Length")
    
    ' Convert strings to numbers
    width = FractionToDecimal(widthStr)
    height = FractionToDecimal(heightStr)
    length = FractionToDecimal(lengthStr)
    
    ' Calculate material quantity (area = length * height)
    Dim materialQuantity As Double
    materialQuantity = length * height
    propertiesToSet("Mtl Unit Qty") = RemoveTrailingZeros(Format(materialQuantity, "0.0000"))
    
    ' Ensure width is the smaller dimension for stock size
    If width > height Then
        Dim temp As Double
        temp = width
        width = height
        height = temp
    End If
    
    ' Convert dimensions to fractions
    widthStr = ConvertDecimalToFraction(width)
    heightStr = ConvertDecimalToFraction(height)
    
    ' Wood stock size format WITH WD prefix
    Dim stockSize As String
    stockSize = "WD " & widthStr & " x " & heightStr
    
    ' Remove Mtl Part Number if it exists (Wood doesn't use it)
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
    End If
    
    ' Set final properties
    propertiesToSet("Stock Size") = stockSize
End Sub










' Process EXP reference category (Expanded Metal)
' This is always Type P (Purchase part)
' No plate dimension checking required
Sub ProcessEXP(ByVal model As ModelDoc2, ByVal refCategory As String, ByVal modelUOM As String, ByRef propertiesToSet As Object)
    DebugLog "#########################"
    DebugLog "### ENTERING ProcessEXP FUNCTION ###"
    DebugLog "#########################"
    DebugLog "refCategory: " & refCategory
    DebugLog "modelUOM: " & modelUOM
    
    Dim widthStr As String, heightStr As String, lengthStr As String
    Dim width As Double, height As Double, length As Double
    Dim temp As Double
   
    ' Get dimensions from properties
    DebugLog "=== GETTING DIMENSIONS FROM PROPERTIES ==="
    widthStr = GetPropertyValue(propertiesToSet, "Width")
    heightStr = GetPropertyValue(propertiesToSet, "Height")
    lengthStr = GetPropertyValue(propertiesToSet, "Length")
    
    DebugLog "Raw property values:"
    DebugLog "widthStr from properties: '" & widthStr & "'"
    DebugLog "heightStr from properties: '" & heightStr & "'"
    DebugLog "lengthStr from properties: '" & lengthStr & "'"
    DebugLog "modelUOM: " & modelUOM
    
    ' Get the UPDATED Dimensional UOM from properties (AddDimensionsProperties may have changed it)
    Dim actualUOM As String
    If propertiesToSet.exists("Dimensional UOM") Then
        actualUOM = propertiesToSet("Dimensional UOM")
        DebugLog "Using updated Dimensional UOM from properties: " & actualUOM
    Else
        actualUOM = modelUOM
        DebugLog "Using original model UOM: " & actualUOM
    End If
    
    ' Convert strings to numbers
    DebugLog "=== CONVERTING TO NUMBERS ==="
    width = FractionToDecimal(widthStr)
    height = FractionToDecimal(heightStr)
    length = FractionToDecimal(lengthStr)
    
    DebugLog "After FractionToDecimal conversion:"
    DebugLog "width: " & width
    DebugLog "height: " & height
    DebugLog "length: " & length
    
    ' *** NO PLATE CHECK FOR EXP - ALWAYS TYPE P ***
    DebugLog "=== SKIPPING CheckPlateDimensions - EXP is always Type P ==="
    
    ' *** SET TYPE P ***
    DebugLog "=== SETTING TYPE P - PURCHASE PART ==="
    propertiesToSet("Type") = "P"
    
    ' Remove Mtl Part Number and Mtl Unit Qty for Purchase parts
    If propertiesToSet.exists("Mtl Part Number") Then
        propertiesToSet.Remove "Mtl Part Number"
        DebugLog "Removed Mtl Part Number"
    End If
    If propertiesToSet.exists("Mtl Unit Qty") Then
        propertiesToSet.Remove "Mtl Unit Qty"
        DebugLog "Removed Mtl Unit Qty"
    End If
    
    ' Ensure width is the smaller dimension for stock size
    DebugLog "=== DIMENSION SWAPPING FOR STOCK SIZE ==="
    DebugLog "Before dimension swap - width: " & width & ", height: " & height
    If width > height Then
        temp = width
        width = height
        height = temp
        DebugLog "Swapped dimensions - width: " & width & ", height: " & height
    Else
        DebugLog "No swap needed - width is already smaller"
    End If
    
    ' Convert to appropriate format based on actualUOM
    DebugLog "=== CONVERTING TO APPROPRIATE FORMAT ==="
    DebugLog "Current actualUOM: " & actualUOM
    
    If actualUOM = "in" Then
        ' Imperial - use fractions
        widthStr = ConvertDecimalToFraction(width)
        heightStr = ConvertDecimalToFraction(height)
        DebugLog "Using imperial fractions:"
    Else
        ' Metric - use FormatMetricDimension
        widthStr = FormatMetricDimension(width)
        heightStr = FormatMetricDimension(height)
        DebugLog "Using metric format:"
    End If
    
    DebugLog "widthStr: " & widthStr
    DebugLog "heightStr: " & heightStr
    
    ' Set Stock Size for Purchase part
    Dim stockSize As String
    stockSize = refCategory & " " & widthStr & " x " & heightStr
    propertiesToSet("Stock Size") = stockSize
    DebugLog "Set Stock Size for Purchase part: " & stockSize
    
    DebugLog "=== ProcessEXP COMPLETED ==="
    DebugLog "Final Results:"
    DebugLog "Stock Size: " & stockSize
    DebugLog "Type: P"
    DebugLog "Dimensional UOM: " & actualUOM
    DebugLog "#########################"
End Sub



