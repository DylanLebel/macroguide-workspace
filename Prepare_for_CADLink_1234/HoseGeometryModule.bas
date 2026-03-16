Attribute VB_Name = "HoseGeometryModule"
Option Explicit
Option Private Module

'================================================================================
'  ENHANCED HOSE GEOMETRY MODULE - WITH STRUCTURAL MEMBER MATCHING (FIXED)
'  Handles creation of hose flat configurations with exact structural member matching
'  Reads original structural member properties and applies them to new flat line
'  FIXED: Proper variable declarations and configuration handling
'  FIXED: Early exit if FlatHose config exists
'  FIXED: Line sketch suppression in all configs
'================================================================================

' === MODULE-LEVEL DECLARATIONS ===
Private swApp As SldWorks.SldWorks
Private swModel As SldWorks.ModelDoc2

' === Data structure to hold structural member properties ===
Private Type StructuralMemberData
    profilePath As String
    ConfigurationName As String
    Standard As String
    ProfileType As String
    ProfileFile As String
    MirrorProfile As Boolean
    RotationAngle As Double
    ConnectionType As Long
    MirrorProfileAxis As Long
End Type

' Initialize module - call this from main macro if needed
Public Sub InitializeModule()
    Set swApp = Application.SldWorks
End Sub

' === MAIN PUBLIC ENTRY POINT ===
Public Sub CreateHoseGeometry(model As ModelDoc2, ByRef propertiesToSet As Object)
    
    DebugLog "=== ENHANCED HOSE GEOMETRY MODULE START ==="
    
    On Error GoTo ErrorHandler
    
    ' Initialize module-level variables
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
    End If
    Set swModel = model
    
    If model Is Nothing Then
        DebugLog "ERROR: Model is Nothing"
        Exit Sub
    End If
    
    DebugLog "Model type: " & model.GetType
    
    If model.GetType <> 1 Then  ' 1 = swDocPART
        DebugLog "Not a part file, exiting"
        Exit Sub
    End If
    
    ' Get hose length from properties
    Dim hoseLength As Double
    If propertiesToSet.exists("LengthA") Then
        hoseLength = CDbl(propertiesToSet("LengthA"))
        DebugLog "Using hose length from properties: " & hoseLength
    Else
        hoseLength = 100  ' Default length
        DebugLog "Using default hose length: " & hoseLength
    End If
    
    ' Create the flat hose configuration with structural member
    Call CreateFlatHoseConfiguration(model, hoseLength, propertiesToSet)
    
    DebugLog "=== ENHANCED HOSE GEOMETRY MODULE END ==="
    Exit Sub
    
ErrorHandler:
    DebugLog "ERROR in CreateHoseGeometry: " & Err.description & " (Number: " & Err.Number & ")"
    DebugLog "=== ENHANCED HOSE GEOMETRY MODULE END WITH ERROR ==="
    
End Sub

' Create flat hose configuration with line sketch and structural member
Private Sub CreateFlatHoseConfiguration(model As ModelDoc2, hoseLength As Double, ByRef propertiesToSet As Object)
    
    DebugLog "=== Creating flat hose configuration with enhanced structural member matching ==="
    
    Set swModel = model
    
    Dim swConfigMgr As ConfigurationManager
    Dim swConfig As Configuration
    Dim configName As String
    Dim newConfig As Configuration
    
    Set swConfigMgr = model.ConfigurationManager
    
    Dim originalConfigName As String
    originalConfigName = model.ConfigurationManager.ActiveConfiguration.Name
    DebugLog "Original configuration: " & originalConfigName
    
    configName = "FlatHose"
    DebugLog "Target configuration name: " & configName
    
    ' === FIX 2: CHECK IF CONFIG EXISTS AND EXIT EARLY ===
    Set swConfig = model.GetConfigurationByName(configName)
    Dim configExists As Boolean
    configExists = Not swConfig Is Nothing
    
    If configExists Then
        DebugLog "Configuration '" & configName & "' already exists. Exiting without changes."
        Exit Sub  ' EXIT EARLY - don't run if config already exists
    End If
    
    DebugLog "Creating new configuration: " & configName
    Set newConfig = swConfigMgr.AddConfiguration2(configName, "", "", 1, "", "", True)
    
    If newConfig Is Nothing Then
        DebugLog "Failed to create FlatHose configuration"
        Exit Sub
    End If
    
    Set swConfig = newConfig
    DebugLog "Configuration created successfully"
    
    ' === READ ORIGINAL STRUCTURAL MEMBER PROPERTIES FIRST ===
    Dim originalStructuralMemberData As StructuralMemberData
    If Not ReadOriginalStructuralMemberProperties(model, configName, originalStructuralMemberData) Then
        DebugLog "ERROR: Could not read original structural member properties"
        Exit Sub
    End If
    
    DebugLog "Successfully read original structural member properties:"
    DebugLog "  Profile Path: " & originalStructuralMemberData.profilePath
    DebugLog "  Configuration: " & originalStructuralMemberData.ConfigurationName
    DebugLog "  Standard: " & originalStructuralMemberData.Standard
    DebugLog "  Type: " & originalStructuralMemberData.ProfileType
    DebugLog "  Mirror Profile: " & originalStructuralMemberData.MirrorProfile
    DebugLog "  Rotation Angle: " & originalStructuralMemberData.RotationAngle
    DebugLog "  Connection Type: " & originalStructuralMemberData.ConnectionType
    
    ' Activate the configuration
    model.ShowConfiguration2 configName
    
    ' Get the actual hose length from cut-list properties
    Dim actualLength As Double
    actualLength = GetHoseLengthFromCutList(model, configName, propertiesToSet)
    
    If actualLength > 0 Then
        hoseLength = actualLength
        DebugLog "Using length from cut list: " & hoseLength
    Else
        DebugLog "Using provided length: " & hoseLength
    End If

    ' Lock the model AFTER config changes and BEFORE geometry creation
    model.Lock
    On Error GoTo ErrorHandler

    ' Create or update the line sketch
    Dim lineSketchName As String
   ' lineSketchName = CreateOrUpdateLineSketch(model, configName, hoseLength)
    lineSketchName = CreateOrUpdateLineSketch(model, hoseLength, configName)
    DebugLog "Returned lineSketchName: '" & lineSketchName & "'"

    ' === ENHANCED: Create structural member with EXACT matching properties ===
    If lineSketchName <> "" Then
        DebugLog "Calling enhanced CreateStructuralMemberFromSketch..."
        Call CreateStructuralMemberFromSketchWithMatching(model, lineSketchName, originalStructuralMemberData, configName)
        DebugLog "Returned from enhanced CreateStructuralMemberFromSketch"
    Else
        DebugLog "ERROR: lineSketchName is empty, not calling CreateStructuralMemberFromSketch"
    End If
    
    ' Now handle configuration-specific suppression with FIXED approach
    Call ConfigureFeatureVisibilityFixed(model, configName, lineSketchName)
    
    ' A final rebuild while the model is still locked
    model.ForceRebuild3 True
    
    DebugLog "Enhanced FlatHose configuration with structural member completed successfully"

CleanExit:
    ' Unlock the model to show the final state
    model.Unlock

    ' Show enhanced message to user
    Dim lengthMessage As String
    If propertiesToSet.exists("Length") And propertiesToSet.exists("LengthA") Then
        lengthMessage = "Length (Fraction): " & propertiesToSet("Length") & vbCrLf & _
                        "Length (Decimal): " & propertiesToSet("LengthA") & " units"
    Else
        lengthMessage = "Length: " & hoseLength & " units"
    End If
    
    swApp.SendMsgToUser2 "Enhanced FlatHose configuration created successfully!" & vbCrLf & vbCrLf & _
                         "� Line sketch created: " & lineSketchName & vbCrLf & _
                         "� Structural member created with exact profile match" & vbCrLf & _
                         "� Profile: " & originalStructuralMemberData.Standard & " / " & originalStructuralMemberData.ProfileType & vbCrLf & _
                         "� Configuration: " & originalStructuralMemberData.ConfigurationName & vbCrLf & _
                         "� Configuration-specific visibility set" & vbCrLf & _
                         lengthMessage, _
                         swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
    Exit Sub

ErrorHandler:
    DebugLog "ERROR in CreateFlatHoseConfiguration: " & Err.description
    ' CRITICAL - Always unlock the model if an error occurs
    model.Unlock
End Sub

' === Read original structural member properties ===
Private Function ReadOriginalStructuralMemberProperties(model As ModelDoc2, flatHoseConfigName As String, ByRef structData As StructuralMemberData) As Boolean
    
    DebugLog "=== Reading original structural member properties ==="
    
    On Error GoTo EH
    
    ' Store current config
    Dim currentConfig As String
    currentConfig = model.ConfigurationManager.ActiveConfiguration.Name
    
    ' Find original configuration (not FlatHose)
    Dim configNames As Variant
    configNames = model.GetConfigurationNames
    Dim originalConfigName As String
    Dim i As Integer
    
    For i = 0 To UBound(configNames)
        If configNames(i) <> flatHoseConfigName Then
            originalConfigName = configNames(i)
            Exit For
        End If
    Next i
    
    If originalConfigName = "" Then
        DebugLog "Could not find original configuration"
        ReadOriginalStructuralMemberProperties = False
        Exit Function
    End If
    
    DebugLog "Switching to original config to read structural member: " & originalConfigName
    model.ShowConfiguration2 originalConfigName
    
    ' Find the first active structural member
    Dim feat As SldWorks.Feature
    Set feat = model.FirstFeature
    
    Do While Not feat Is Nothing
        Dim tname As String
        tname = feat.GetTypeName2
        
        If tname = "WeldMemberFeat" And Not feat.IsSuppressed Then
            DebugLog "Found active structural member: " & feat.Name
            
            ' Get the definition
            Dim fdata As SldWorks.StructuralMemberFeatureData
            Set fdata = feat.GetDefinition
            
            If Not fdata Is Nothing Then
                Dim ok As Boolean
                ok = fdata.AccessSelections(model, Nothing)
                DebugLog "AccessSelections result: " & ok
                
                If ok Then
                    ' Read all the properties we need
                    structData.profilePath = NzStr(fdata.WeldmentProfilePath)
                    structData.ConfigurationName = NzStr(fdata.ConfigurationName)
                    
                    ' Parse standard and type from path
                    ParseStandardTypeFromPath structData.profilePath, structData.Standard, structData.ProfileType, structData.ProfileFile
                    
                    ' Read additional properties for exact matching
                    On Error Resume Next
                    structData.MirrorProfile = fdata.MirrorProfile
                    structData.RotationAngle = fdata.Angle
                    structData.MirrorProfileAxis = fdata.MirrorProfileAxis
                    structData.ConnectionType = 0 ' swConnectedSegments_SimpleCut
                    On Error GoTo EH
                    
                    fdata.ReleaseSelectionAccess
                    
                    ' Switch back to target config
                    model.ShowConfiguration2 currentConfig
                    
                    ReadOriginalStructuralMemberProperties = True
                    Exit Function
                Else
                    DebugLog "AccessSelections failed for: " & feat.Name
                End If
            End If
        End If
        
        Set feat = feat.GetNextFeature
    Loop
    
    ' Switch back to target config
    model.ShowConfiguration2 currentConfig
    DebugLog "No suitable structural member found in original configuration"
    ReadOriginalStructuralMemberProperties = False
    Exit Function

EH:
    DebugLog "Error in ReadOriginalStructuralMemberProperties: " & Err.Number & " - " & Err.description
    model.ShowConfiguration2 currentConfig
    ReadOriginalStructuralMemberProperties = False
End Function

' Parse Standard and Type from the profile path
Private Sub ParseStandardTypeFromPath(ByVal fullPath As String, _
                                      ByRef outStandard As String, _
                                      ByRef outType As String, _
                                      ByRef outFile As String)
    outStandard = ""
    outType = ""
    outFile = ""

    If Len(fullPath) = 0 Then
        DebugLog "Empty profile path"
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(fullPath, "\")
    Dim n As Long
    n = UBound(parts)

    If n >= 0 Then outFile = parts(n)
    If n >= 1 Then outType = parts(n - 1)
    If n >= 2 Then outStandard = parts(n - 2)

    DebugLog "Parsed - Standard: " & outStandard & ", Type: " & outType & ", File: " & outFile
End Sub

' Null safe string helper
Private Function NzStr(ByVal s As String) As String
    If Len(s) = 0 Then
        NzStr = ""
    Else
        NzStr = s
    End If
End Function

' === FIXED: Create structural member with exact property matching ===
Private Sub CreateStructuralMemberFromSketchWithMatching(model As ModelDoc2, sketchName As String, ByRef structData As StructuralMemberData, configName As String)
    DebugLog "=== ENHANCED STRUCTURAL MEMBER CREATION WITH CONFIGURATION FIX ==="
    DebugLog "Input parameters:"
    DebugLog "  sketchName: " & sketchName
    DebugLog "  configName: " & configName
    DebugLog "  structData.profilePath: " & structData.profilePath
    DebugLog "  structData.ConfigurationName: " & structData.ConfigurationName

    On Error GoTo ErrorHandler

    ' Step 1: Clean up any existing structural members
    DebugLog "=== STEP 1: CLEANUP ==="
    DeleteExistingFlatHoseStructuralMember model, configName
    
    ' Step 2: Rebuild
    DebugLog "=== STEP 2: REBUILD AFTER CLEANUP ==="
    model.ForceRebuild3 True
    model.EditRebuild3
    
    ' Step 3: Ensure we're in the right configuration
    If model.ConfigurationManager.ActiveConfiguration.Name <> configName Then
        model.ShowConfiguration2 configName
    End If
    
    ' Step 4: Clear selections
    DebugLog "=== STEP 3: CLEAR SELECTIONS ==="
    model.ClearSelection2 True
    
    ' Step 5: Select the sketch segment for structural member
    DebugLog "=== STEP 4: SELECT SKETCH SEGMENT ==="
    
    Dim boolStatus As Boolean
    boolStatus = SelectSketchSegmentForStructuralMember(model, sketchName)
    
    If Not boolStatus Then
        DebugLog "ERROR: Could not select sketch segment - cannot create structural member"
        Exit Sub
    End If
    
    ' Step 6: Ensure Weldment feature exists
    DebugLog "=== STEP 5: CHECK WELDMENT FEATURE ==="
    Dim weldmentFeature As Feature
    Set weldmentFeature = model.FeatureByName("Weldment")
    
    Dim swFeatMgr As FeatureManager
    Set swFeatMgr = model.FeatureManager
    
    If weldmentFeature Is Nothing Then
        DebugLog "Creating Weldment feature..."
        Set weldmentFeature = swFeatMgr.InsertWeldmentFeature()
        
        If weldmentFeature Is Nothing Then
            DebugLog "ERROR: Failed to create Weldment feature"
            Exit Sub
        End If
        
        ' After creating weldment, need to reselect the segment
        model.ClearSelection2 True
        boolStatus = SelectSketchSegmentForStructuralMember(model, sketchName)
    End If
    
    ' === STEP 7: VERIFY AND SET PROFILE CONFIGURATION ===
    DebugLog "=== STEP 6: VERIFY PROFILE CONFIGURATION ==="
    Dim validConfigName As String
    validConfigName = VerifyProfileConfiguration(structData.profilePath, structData.ConfigurationName)
    DebugLog "Valid configuration name: " & validConfigName
    
    ' Step 8: Create the structural member with proper configuration handling
    DebugLog "=== STEP 7: CREATE STRUCTURAL MEMBER WITH CONFIGURATION ==="
    
    Dim newFeat As Feature
    Set newFeat = CreateStructuralMemberWithConfiguration(swFeatMgr, structData.profilePath, validConfigName, structData)
    
    If Not newFeat Is Nothing Then
        DebugLog "SUCCESS! Structural member created: " & newFeat.Name
        
        ' Step 9: Verify and fix configuration if needed
        DebugLog "=== STEP 8: VERIFY CONFIGURATION WAS SET CORRECTLY ==="
        Call VerifyAndFixStructuralMemberConfiguration(model, newFeat, structData, validConfigName)
        
        ' Step 10: Final rebuild
        model.ForceRebuild3 True
        DebugLog "Structural member creation completed successfully"
    Else
        DebugLog "ERROR: All structural member creation methods failed"
        DebugLog "Profile path: " & structData.profilePath
        DebugLog "Configuration: " & validConfigName
    End If
    
    Exit Sub

ErrorHandler:
    DebugLog "=== CRITICAL ERROR IN STRUCTURAL MEMBER CREATION ==="
    DebugLog "Error Number: " & Err.Number
    DebugLog "Error Description: " & Err.description
    DebugLog "Error Source: " & Err.Source
End Sub

' === NEW: Verify that the configuration exists in the profile file ===
Private Function VerifyProfileConfiguration(profilePath As String, configName As String) As String
    DebugLog "=== Verifying profile configuration ==="
    DebugLog "Profile path: " & profilePath
    DebugLog "Requested configuration: " & configName
    
    On Error GoTo ErrorHandler
    
    ' Try to open the profile file to check configurations
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
    End If
    
    Dim profileDoc As ModelDoc2
    Set profileDoc = swApp.OpenDoc6(profilePath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
    
    If profileDoc Is Nothing Then
        DebugLog "Could not open profile file, using requested configuration as-is"
        VerifyProfileConfiguration = configName
        Exit Function
    End If
    
    ' Get all configurations from the profile
    Dim configNames As Variant
    configNames = profileDoc.GetConfigurationNames
    
    Dim i As Integer
    Dim foundConfig As Boolean
    foundConfig = False
    
    DebugLog "Available configurations in profile:"
    For i = 0 To UBound(configNames)
        DebugLog "  " & i & ": " & configNames(i)
        If UCase(configNames(i)) = UCase(configName) Then
            foundConfig = True
            DebugLog "  *** MATCH FOUND ***"
        End If
    Next i
    
    ' Close the profile document
    swApp.CloseDoc profileDoc.GetTitle
    
    If foundConfig Then
        DebugLog "Configuration verified: " & configName
        VerifyProfileConfiguration = configName
    Else
        ' Use default configuration if requested one doesn't exist
        If UBound(configNames) >= 0 Then
            DebugLog "Configuration not found, using default: " & configNames(0)
            VerifyProfileConfiguration = configNames(0)
        Else
            DebugLog "No configurations found, using empty string"
            VerifyProfileConfiguration = ""
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    DebugLog "Error verifying configuration: " & Err.description
    VerifyProfileConfiguration = configName  ' Use original if verification fails
End Function

' === NEW: Create structural member with multiple methods and configuration handling ===
Private Function CreateStructuralMemberWithConfiguration(swFeatMgr As FeatureManager, profilePath As String, configName As String, ByRef structData As StructuralMemberData) As Feature
    
    DebugLog "=== Creating structural member with configuration: " & configName & " ==="
    
    Dim newFeat As Feature
    
    ' Method 1: InsertStructuralWeldment4 with verified configuration
    DebugLog "Method 1: InsertStructuralWeldment4 with configuration"
    Set newFeat = swFeatMgr.InsertStructuralWeldment4(profilePath, 0, True, configName)
    
    If Not newFeat Is Nothing Then
        DebugLog "Method 1 SUCCESS"
        Set CreateStructuralMemberWithConfiguration = newFeat
        Exit Function
    End If
    
    ' Method 2: Try with empty configuration and set it afterwards
    DebugLog "Method 2: InsertStructuralWeldment4 with empty config"
    Set newFeat = swFeatMgr.InsertStructuralWeldment4(profilePath, 0, True, "")
    
    If Not newFeat Is Nothing Then
        DebugLog "Method 2 SUCCESS - will set configuration afterwards"
        Set CreateStructuralMemberWithConfiguration = newFeat
        Exit Function
    End If
    
    ' Method 3: Basic InsertStructuralWeldment
    DebugLog "Method 3: Basic InsertStructuralWeldment"
    On Error Resume Next
    Set newFeat = swFeatMgr.InsertStructuralWeldment(profilePath, 0, True)
    On Error GoTo 0
    
    If Not newFeat Is Nothing Then
        DebugLog "Method 3 SUCCESS - will set configuration afterwards"
        Set CreateStructuralMemberWithConfiguration = newFeat
        Exit Function
    End If
    
    ' Method 4: Try with groups (SolidWorks 2023 specific)
    DebugLog "Method 4: Using groups method"
    Set newFeat = CreateStructuralMemberWithGroups(swFeatMgr, profilePath, configName, structData)
    
    Set CreateStructuralMemberWithConfiguration = newFeat  ' May be Nothing if all methods failed
    
End Function

' === NEW: Verify and fix the structural member configuration after creation ===
Private Sub VerifyAndFixStructuralMemberConfiguration(model As ModelDoc2, feat As Feature, ByRef structData As StructuralMemberData, validConfigName As String)
    
    DebugLog "=== Verifying structural member configuration ==="
    
    On Error GoTo ErrorHandler
    
    ' Get the feature definition
    Dim swStructMemberData As StructuralMemberFeatureData
    Set swStructMemberData = feat.GetDefinition
    
    If swStructMemberData Is Nothing Then
        DebugLog "Could not get structural member data for verification"
        Exit Sub
    End If
    
    ' Access selections to read/modify
    Dim accessResult As Boolean
    accessResult = swStructMemberData.AccessSelections(model, Nothing)
    
    If Not accessResult Then
        DebugLog "Could not access selections for verification"
        Exit Sub
    End If
    
    ' Check current configuration
    Dim currentConfig As String
    currentConfig = swStructMemberData.ConfigurationName
    DebugLog "Current configuration in structural member: " & currentConfig
    DebugLog "Expected configuration: " & validConfigName
    
    Dim needsUpdate As Boolean
    needsUpdate = False
    
    ' Check if configuration needs to be updated
    If UCase(currentConfig) <> UCase(validConfigName) And validConfigName <> "" Then
        DebugLog "Configuration mismatch - updating..."
        swStructMemberData.ConfigurationName = validConfigName
        needsUpdate = True
    End If
    
    ' Also set other properties while we're here
    If structData.RotationAngle <> 0 Then
        DebugLog "Setting rotation angle: " & structData.RotationAngle
        swStructMemberData.Angle = structData.RotationAngle
        needsUpdate = True
    End If
    
    If structData.MirrorProfile Then
        DebugLog "Setting mirror profile: True"
        swStructMemberData.MirrorProfile = True
        swStructMemberData.MirrorProfileAxis = structData.MirrorProfileAxis
        needsUpdate = True
    End If
    
    ' Apply changes if needed
    If needsUpdate Then
        DebugLog "Applying configuration and property updates..."
        Dim modifyResult As Boolean
        modifyResult = feat.ModifyDefinition(swStructMemberData, model, Nothing)
        DebugLog "Modify result: " & modifyResult
        
        If modifyResult Then
            DebugLog "Configuration successfully updated to: " & validConfigName
        Else
            DebugLog "WARNING: Failed to update configuration"
        End If
    Else
        DebugLog "No updates needed"
    End If
    
    ' Release access
    swStructMemberData.ReleaseSelectionAccess
    
    Exit Sub
    
ErrorHandler:
    DebugLog "Error in VerifyAndFixStructuralMemberConfiguration: " & Err.description
    On Error Resume Next
    If Not swStructMemberData Is Nothing Then
        swStructMemberData.ReleaseSelectionAccess
    End If
    On Error GoTo 0
End Sub

' === IMPROVED: Select sketch segment with multiple methods ===
Private Function SelectSketchSegmentForStructuralMember(model As ModelDoc2, sketchName As String) As Boolean
    
    DebugLog "=== Selecting sketch segment for structural member ==="
    
    Dim boolStatus As Boolean
    
    ' Method 1: Standard Line1@ selection
    model.ClearSelection2 True
    boolStatus = model.Extension.SelectByID2("Line1@" & sketchName, "EXTSKETCHSEGMENT", 0, 0, 0, False, 1, Nothing, 0)
    DebugLog "Method 1 (Line1@" & sketchName & "): " & boolStatus
    If boolStatus Then
        SelectSketchSegmentForStructuralMember = True
        Exit Function
    End If
    
    ' Method 2: Try Line@ without number
    boolStatus = model.Extension.SelectByID2("Line@" & sketchName, "EXTSKETCHSEGMENT", 0, 0, 0, False, 1, Nothing, 0)
    DebugLog "Method 2 (Line@" & sketchName & "): " & boolStatus
    If boolStatus Then
        SelectSketchSegmentForStructuralMember = True
        Exit Function
    End If
    
    ' Method 3: Direct segment selection
    DebugLog "Method 3: Direct segment selection"
    boolStatus = SelectSketchSegmentDirectly(model, sketchName)
    If boolStatus Then
        SelectSketchSegmentForStructuralMember = True
        Exit Function
    End If
    
    DebugLog "All selection methods failed"
    SelectSketchSegmentForStructuralMember = False
    
End Function

' === Helper function for direct segment selection ===
Private Function SelectSketchSegmentDirectly(model As ModelDoc2, sketchName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Find the sketch feature
    Dim sketchFeature As Feature
    Set sketchFeature = model.FeatureByName(sketchName)
    
    If sketchFeature Is Nothing Then
        DebugLog "Could not find sketch feature: " & sketchName
        SelectSketchSegmentDirectly = False
        Exit Function
    End If
    
    ' Get the sketch object
    Dim swSketch As Sketch
    Set swSketch = sketchFeature.GetSpecificFeature2
    
    If swSketch Is Nothing Then
        DebugLog "Could not get sketch object"
        SelectSketchSegmentDirectly = False
        Exit Function
    End If
    
    ' Get sketch segments
    Dim sketchSegs As Variant
    sketchSegs = swSketch.GetSketchSegments
    
    If IsEmpty(sketchSegs) Then
        DebugLog "No sketch segments found"
        SelectSketchSegmentDirectly = False
        Exit Function
    End If
    
    ' Select the first line segment
    Dim i As Integer
    For i = 0 To UBound(sketchSegs)
        Dim seg As SketchSegment
        Set seg = sketchSegs(i)
        
        If seg.GetType = 1 Then ' swSketchLINE
            DebugLog "Found line segment, attempting to select it"
            Dim selectResult As Boolean
            selectResult = seg.Select4(False, Nothing)
            If selectResult Then
                DebugLog "Successfully selected line segment directly"
                SelectSketchSegmentDirectly = True
                Exit Function
            End If
        End If
    Next i
    
    SelectSketchSegmentDirectly = False
    Exit Function
    
ErrorHandler:
    DebugLog "Error in SelectSketchSegmentDirectly: " & Err.description
    SelectSketchSegmentDirectly = False
End Function

' === Alternative method using groups for SolidWorks 2023 ===
Private Function CreateStructuralMemberWithGroups(swFeatMgr As FeatureManager, profilePath As String, configName As String, ByRef structData As StructuralMemberData) As Feature
    
    DebugLog "=== Creating structural member with groups method ==="
    
    On Error GoTo ErrorHandler
    
    ' Initialize swApp if needed
    If swApp Is Nothing Then
        Set swApp = Application.SldWorks
    End If
    
    Dim model As ModelDoc2
    Set model = swApp.ActiveDocument
    
    ' Get the selected segment
    Dim selMgr As SelectionMgr
    Set selMgr = model.SelectionManager
    
    If selMgr.GetSelectedObjectCount2(-1) = 0 Then
        DebugLog "No segments selected for groups method"
        Set CreateStructuralMemberWithGroups = Nothing
        Exit Function
    End If
    
    Dim selectedObj As Object
    Set selectedObj = selMgr.GetSelectedObject6(1, -1)
    
    If selectedObj Is Nothing Then
        DebugLog "Could not get selected object"
        Set CreateStructuralMemberWithGroups = Nothing
        Exit Function
    End If
    
    ' Create a group
    Dim group As StructuralMemberGroup
    Set group = swFeatMgr.CreateStructuralMemberGroup
    
    If group Is Nothing Then
        DebugLog "Could not create structural member group"
        Set CreateStructuralMemberWithGroups = Nothing
        Exit Function
    End If
    
    ' Set up the group
    Dim segs(0) As Object
    Set segs(0) = selectedObj
    
    group.Segments = segs
    group.ApplyCornerTreatment = False
    group.GapWithinGroup = 0
    group.Angle = structData.RotationAngle
    group.MirrorProfile = structData.MirrorProfile
    group.MirrorProfileAxis = structData.MirrorProfileAxis
    
    ' Create groups array
    Dim groups(0) As Object
    Set groups(0) = group
    
    ' Try InsertStructuralWeldment5 with the group
    Dim newFeat As Feature
    Set newFeat = swFeatMgr.InsertStructuralWeldment5(profilePath, 0, True, groups, configName)
    
    Set CreateStructuralMemberWithGroups = newFeat
    Exit Function
    
ErrorHandler:
    DebugLog "Error in CreateStructuralMemberWithGroups: " & Err.description
    Set CreateStructuralMemberWithGroups = Nothing
End Function

' Get hose length from cut-list properties and set Length/LengthA properties
Private Function GetHoseLengthFromCutList(model As ModelDoc2, configName As String, ByRef propertiesToSet As Object) As Double
    
    DebugLog "Starting search for TOTAL LENGTH in cut-list properties..."
    
    ' Find the original configuration (not the FlatHose one)
    Dim configNames As Variant
    configNames = model.GetConfigurationNames
    Dim originalConfigName As String
    Dim i As Integer
    
    For i = 0 To UBound(configNames)
        If configNames(i) <> configName Then
            originalConfigName = configNames(i)
            Exit For
        End If
    Next i
    
    DebugLog "Switching to original config to read properties: " & originalConfigName
    model.ShowConfiguration2 originalConfigName
    
    ' Get document units and conversion factor
    Dim unitString As String
    Dim conversionFactor As Double
    Call GetDocumentUnits(model, unitString, conversionFactor)
    DebugLog "Document units: " & unitString & " (conversion factor: " & conversionFactor & ")"
    
    ' Read cut-list properties from original configuration
    Dim swCutListFeat As Feature
    Dim swCutListCustomPropMgr As customPropertyManager
    Dim propCount As Long
    Dim propNames As Variant
    Dim propValues As Variant
    Dim propTypes As Variant
    Dim propResolved As Variant
    Dim propLink As Variant
    Dim j As Integer
    
    Dim maxLength As Double
    maxLength = 0
    
    Set swCutListFeat = model.FirstFeature
    Do While Not swCutListFeat Is Nothing
        DebugLog "Feature: " & swCutListFeat.Name & " | Type: " & swCutListFeat.GetTypeName2
        If swCutListFeat.GetTypeName2 = "CutListFolder" Then
            Set swCutListCustomPropMgr = swCutListFeat.customPropertyManager
            propCount = swCutListCustomPropMgr.GetAll3(propNames, propTypes, propValues, propResolved, propLink)
            DebugLog "Cut list folder found with " & propCount & " properties"
            
            ' Check if this cut-list has valid quantity (skip if QUANTITY = 0)
            Dim hasValidQuantity As Boolean
            hasValidQuantity = True
            
            If propCount > 0 Then
                For j = 0 To propCount - 1
                    If propNames(j) = "QUANTITY" Then
                        Dim qty As Double
                        On Error Resume Next
                        qty = CDbl(propValues(j))
                        On Error GoTo 0
                        If qty = 0 Then
                            hasValidQuantity = False
                            DebugLog "Skipping cut-list with QUANTITY = 0: " & swCutListFeat.Name
                            Exit For
                        End If
                    End If
                Next j
            End If
            
            ' Only process cut-lists with valid quantities
            If hasValidQuantity And propCount > 0 Then
                For j = 0 To propCount - 1
                    DebugLog "Property: " & propNames(j) & " = " & propValues(j)
                    If propNames(j) = "TOTAL LENGTH" Then
                        On Error Resume Next
                        Dim tempLength As Double
                        tempLength = CDbl(propValues(j))
                        On Error GoTo 0
                        If tempLength > maxLength Then
                            maxLength = tempLength
                            DebugLog "New max length found: " & maxLength & " " & unitString
                        End If
                    End If
                Next j
            End If
        End If
        Set swCutListFeat = swCutListFeat.GetNextFeature
    Loop
    
    ' Switch back to FlatHose configuration
    model.ShowConfiguration2 configName
    
    ' If we found a length, process it and set properties
    If maxLength > 0 Then
        DebugLog "Processing found length: " & maxLength & " " & unitString
        
        ' Convert to inches for fraction calculation (standard practice)
        Dim lengthInInches As Double
        Select Case LCase(unitString)
            Case "mm"
                lengthInInches = maxLength / 25.4
            Case "cm"
                lengthInInches = maxLength / 2.54
            Case "m"
                lengthInInches = maxLength * 39.3701
            Case "ft"
                lengthInInches = maxLength * 12
            Case "in"
                lengthInInches = maxLength
            Case Else
                ' Default assumption is mm
                lengthInInches = maxLength / 25.4
        End Select
        
        DebugLog "Length in inches for fraction calculation: " & lengthInInches
        
        ' Set Length as fraction (nearest 1/16th) - using ConvertDecimalToFraction function
        propertiesToSet("Length") = ConvertDecimalToFraction(lengthInInches)
        DebugLog "Set Length property to: " & propertiesToSet("Length")
        
        ' Set LengthA as decimal in original units
        propertiesToSet("LengthA") = CStr(maxLength)
        DebugLog "Set LengthA property to: " & maxLength
        
    Else
        DebugLog "No valid TOTAL LENGTH found in cut-list properties"
    End If
    
    ' Return the maximum length found
    GetHoseLengthFromCutList = maxLength
    DebugLog "Final hose length from cut-list: " & maxLength & " " & unitString
    
End Function

' Get document units and conversion factor
Private Sub GetDocumentUnits(model As ModelDoc2, ByRef unitString As String, ByRef conversionFactor As Double)
    
    On Error Resume Next
    
    ' Get the unit system
    Dim currentUnitSystem As Long
    currentUnitSystem = model.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)

    ' Fallback in case the above fails
    If Err.Number <> 0 Then
        currentUnitSystem = model.GetUnits(0) ' 0 for length units
        Err.Clear
    End If
    
    DebugLog "Detected Unit System Code: " & currentUnitSystem
    
    ' Proper mapping of unit system codes
    Select Case currentUnitSystem
        Case 1 ' swUnitSystem_CGS
            unitString = "cm"
            conversionFactor = 100#
        Case 2 ' swUnitSystem_MKS
            unitString = "m"
            conversionFactor = 1#
        Case 3 ' swUnitSystem_IPS
            unitString = "in"
            conversionFactor = 39.3701
        Case 4 ' swUnitSystem_Custom
            DebugLog "Custom units detected, checking individual length units..."
            Dim lengthUnit As Long
            lengthUnit = model.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinear, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            DebugLog "Custom length unit code: " & lengthUnit
            
            Select Case lengthUnit
                Case 0 ' swMM
                    unitString = "mm"
                    conversionFactor = 1000#
                Case 1 ' swCM
                    unitString = "cm"
                    conversionFactor = 100#
                Case 2 ' swMETER
                    unitString = "m"
                    conversionFactor = 1#
                Case 3 ' swINCHES
                    unitString = "in"
                    conversionFactor = 39.3701
                Case 4 ' swFEET
                    unitString = "ft"
                    conversionFactor = 3.28084
                Case Else
                    unitString = "mm"
                    conversionFactor = 1000#
            End Select
        Case 5 ' swUnitSystem_MMGS
            unitString = "mm"
            conversionFactor = 1000#
        Case Else ' Fallback
            Dim nUnit As Long
            nUnit = model.GetUnits(0)
            Select Case nUnit
                Case 1 ' swMM
                    unitString = "mm"
                    conversionFactor = 1000#
                Case 2 ' swCM
                    unitString = "cm"
                    conversionFactor = 100#
                Case 3 ' swINCHES
                    unitString = "in"
                    conversionFactor = 39.3701
                Case 4 ' swFEET
                    unitString = "ft"
                    conversionFactor = 3.28084
                Case 5 ' swMETER
                    unitString = "m"
                    conversionFactor = 1#
                Case Else
                    unitString = "mm"
                    conversionFactor = 1000#
            End Select
    End Select
    
    DebugLog "Determined unit string: " & unitString & " (conversion factor to meters: " & conversionFactor & ")"
    On Error GoTo 0
    
End Sub

' Create or update the line sketch - returns the actual sketch name
'=== ROBUST PLANE SELECTION (CREATION REMOVED) ===
Private Function GetOrCreateFrontPlane(swModel As SldWorks.ModelDoc2) As Object
    '========================================
    ' Search for existing Front Plane only
    ' Returns Nothing if not found
    '========================================
    
    Dim swFeat As SldWorks.Feature
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swRefPlane As SldWorks.RefPlane
    Dim planeName As String
    Dim foundPlane As Object
    
    Set swSelMgr = swModel.SelectionManager
    Set foundPlane = Nothing
    
    DebugLog "=== SEARCHING FOR FRONT PLANE ==="
    
    ' List of possible Front Plane names to search for
    Dim frontPlaneNames() As String
    frontPlaneNames = Split("Front Plane,Front,BH - Front Plane,Front-Plane,FrontPlane", ",")
    
    ' METHOD 1: Try exact name matches first
    Dim i As Integer
    For i = LBound(frontPlaneNames) To UBound(frontPlaneNames)
        planeName = Trim(frontPlaneNames(i))
        DebugLog "Trying exact match: " & planeName
        
        Set swFeat = swModel.FeatureByName(planeName)
        If Not swFeat Is Nothing Then
            DebugLog "  Found by exact name: " & planeName
            Set foundPlane = swFeat
            Exit For
        End If
    Next i
    
    ' METHOD 2: If no exact match, search all features for planes containing "Front"
If foundPlane Is Nothing Then
    DebugLog "No exact match found. Scanning all features..."
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "RefPlane" Then
            planeName = swFeat.Name
            DebugLog "  Checking plane: " & planeName
            
            ' Check if name contains "Front" (case insensitive)
            If InStr(1, planeName, "Front", vbTextCompare) > 0 Then
                DebugLog "  *** FOUND Front plane by search: " & planeName
                Set foundPlane = swFeat
                Exit Do  ' <-- FIXED: Use Exit Do instead
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End If
    
    ' METHOD 3: Try to get the standard Front Plane from FeatureManager
    If foundPlane Is Nothing Then
        DebugLog "Trying to get standard Front Plane from feature tree..."
        Set foundPlane = swModel.FeatureByPositionReverse(swModel.GetFeatureCount - 6) ' Front Plane is usually 6th from bottom
        If Not foundPlane Is Nothing Then
            If InStr(1, foundPlane.Name, "Front", vbTextCompare) = 0 Then
                Set foundPlane = Nothing ' Not the Front Plane
            Else
                DebugLog "  Found standard Front Plane by position"
            End If
        End If
    End If
    
    ' Report final status
    If Not foundPlane Is Nothing Then
        DebugLog "=== FRONT PLANE READY: " & foundPlane.Name & " ==="
    Else
        DebugLog "=== ERROR: COULD NOT FIND FRONT PLANE ==="
    End If
    
    Set GetOrCreateFrontPlane = foundPlane
    
End Function


'=== COMPLETE: Create or update the line sketch - returns the actual sketch name ===
Private Function CreateOrUpdateLineSketch(swModel As SldWorks.ModelDoc2, _
                                         lineLength As Double, _
                                         Optional existingSketchName As String = "") As String
    
    DebugLog "=== CreateOrUpdateLineSketch START ==="
    DebugLog "Input lineLength: " & lineLength
    DebugLog "Input existingSketchName: " & existingSketchName
    
    ' === UNIT DETECTION CODE ===
    Dim unitString As String
    Dim conversionFactor As Double
    Call GetDocumentUnits(swModel, unitString, conversionFactor)
    
    ' Convert input length to meters (SolidWorks API expects meters)
    Dim lineLengthInMeters As Double
    lineLengthInMeters = lineLength / conversionFactor
    
    DebugLog "Input length: " & lineLength & " " & unitString
    DebugLog "Converted to meters: " & lineLengthInMeters & " meters"
    
    ' === CHECK IF SKETCH EXISTS ===
    Dim existingLineSketch As Feature
    If existingSketchName <> "" Then
        Set existingLineSketch = swModel.FeatureByName(existingSketchName)
        
        If Not existingLineSketch Is Nothing Then
            DebugLog "Found existing sketch: " & existingSketchName
            ' Update the existing sketch
            Call UpdateExistingLineSketch(swModel, existingLineSketch, lineLengthInMeters)
            CreateOrUpdateLineSketch = existingSketchName
            Exit Function
        End If
    End If
    
    ' === CREATE NEW LINE SKETCH ===
    DebugLog "=== Creating new line sketch ==="
    DebugLog "Creating line with length: " & lineLengthInMeters & " meters"
    DebugLog "Original length: " & lineLength & " " & unitString
    
    ' === USE PLANE SEARCH FUNCTION ===
    Dim swFrontPlane As Object
    Set swFrontPlane = GetOrCreateFrontPlane(swModel)
    
    If swFrontPlane Is Nothing Then
        DebugLog "CRITICAL ERROR: Cannot find Front Plane - cannot create sketch"
        MsgBox "ERROR: Front Plane not found in model. Cannot create line sketch.", vbCritical, "Plane Not Found"
        CreateOrUpdateLineSketch = ""
        Exit Function
    End If
    
    ' Select the Front Plane
    Dim boolStatus As Boolean
    boolStatus = swFrontPlane.Select2(False, 0)
    
    If Not boolStatus Then
        DebugLog "ERROR: Cannot select Front Plane"
        MsgBox "ERROR: Cannot select Front Plane. Sketch creation failed.", vbCritical, "Selection Failed"
        CreateOrUpdateLineSketch = ""
        Exit Function
    End If
    
    DebugLog "Successfully selected Front Plane: " & swFrontPlane.Name
    
    ' Clear selection and insert sketch
    swModel.ClearSelection2 True
    boolStatus = swFrontPlane.Select2(False, 0)
    
    ' Insert sketch on selected plane
    swModel.InsertSketch2 True
    
    ' Create horizontal line from origin
    Dim swSketchMgr As SketchManager
    Set swSketchMgr = swModel.SketchManager
    
    ' Create line from (0,0,0) to (length,0,0) - all in METERS for API
    Dim swSketchSeg As SketchSegment
    Set swSketchSeg = swSketchMgr.CreateLine(0#, 0#, 0#, lineLengthInMeters, 0#, 0#)
    
    If swSketchSeg Is Nothing Then
        DebugLog "ERROR: Failed to create line segment"
        CreateOrUpdateLineSketch = ""
        Exit Function
    End If
    
    DebugLog "Successfully created line segment"
    
    ' Exit sketch
    swModel.InsertSketch2 True
    
    ' Get the sketch feature that was just created
    Dim swSelMgr As SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    
    Dim swFeat As Feature
    Set swFeat = swModel.FeatureByPositionReverse(0) ' Get most recently created feature
    
    If Not swFeat Is Nothing Then
        If swFeat.GetTypeName2 = "ProfileFeature" Then
            ' Rename the sketch based on configuration
            Dim newName As String
            Dim configName As String
            configName = swModel.ConfigurationManager.ActiveConfiguration.Name
            newName = "Line_" & configName
            swFeat.Name = newName
            DebugLog "Created and renamed sketch to: " & newName
            CreateOrUpdateLineSketch = newName
        Else
            DebugLog "ERROR: Most recent feature is not a ProfileFeature, it's: " & swFeat.GetTypeName2
            CreateOrUpdateLineSketch = ""
        End If
    Else
        DebugLog "ERROR: Could not get most recent feature"
        CreateOrUpdateLineSketch = ""
    End If
    
    swModel.ClearSelection2 True
    DebugLog "=== CreateOrUpdateLineSketch END ==="
    
End Function


'=== CREATE NEW FRONT PLANE ===
Function CreateFrontPlane(swModel As SldWorks.ModelDoc2) As Feature
    Dim swFeatureMgr As SldWorks.FeatureManager
    Dim swRefPlane As Feature
    Dim swFeat As Feature
    Dim swFeatData As Object
    Dim boolStatus As Boolean
    Dim i As Long
    
    Set swFeatureMgr = swModel.FeatureManager
    
    DebugLog "=== CREATING NEW FRONT PLANE ==="
    
    ' First, let's debug what planes exist and their state
    DebugLog "--- Listing all reference planes ---"
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "RefPlane" Then
            DebugLog "Found plane: " & swFeat.Name & " (Suppressed: " & swFeat.IsSuppressed & ")"
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    DebugLog "--- End plane list ---"
    
    ' Clear any existing selections
    swModel.ClearSelection2 True
    
    ' Get Right Plane feature directly
    Set swFeat = swModel.FeatureByName("Right Plane")
    
    If swFeat Is Nothing Then
        DebugLog "ERROR: Right Plane feature not found by name"
        ' Try to find any plane with "Right" in the name
        Set swFeat = swModel.FirstFeature
        Do While Not swFeat Is Nothing
            If swFeat.GetTypeName2 = "RefPlane" And InStr(1, swFeat.Name, "Right", vbTextCompare) > 0 Then
                DebugLog "Found alternative right plane: " & swFeat.Name
                Exit Do
            End If
            Set swFeat = swFeat.GetNextFeature
        Loop
    End If
    
    ' Check if plane is suppressed
    If Not swFeat Is Nothing Then
        If swFeat.IsSuppressed Then
            DebugLog "Right Plane is suppressed, unsuppressing..."
            ' 1 = swUnSuppressFeature, 2 = swAllConfigurations
            swFeat.SetSuppression2 1, 2, Nothing
            swModel.ForceRebuild3 False
        End If
        
        ' Try selecting the feature
        boolStatus = swFeat.Select2(False, 1)
        DebugLog "Direct feature Select2 result: " & boolStatus
    End If
    
    ' If selection still failed, try with Top Plane as reference instead
    If Not boolStatus Then
        DebugLog "Right Plane selection failed, trying Top Plane..."
        swModel.ClearSelection2 True
        
        Set swFeat = swModel.FeatureByName("Top Plane")
        If Not swFeat Is Nothing Then
            If swFeat.IsSuppressed Then
                ' 1 = swUnSuppressFeature, 2 = swAllConfigurations
                swFeat.SetSuppression2 1, 2, Nothing
                swModel.ForceRebuild3 False
            End If
            boolStatus = swFeat.Select2(False, 1)
            
            If boolStatus Then
                DebugLog "Using Top Plane as reference instead"
                ' Create plane perpendicular to Top Plane
                Set swRefPlane = swFeatureMgr.InsertRefPlane( _
                    4, 1.5707963267949, _
                    0, 0, _
                    0, 0)
            End If
        End If
    Else
        ' Create plane perpendicular to Right Plane
        Set swRefPlane = swFeatureMgr.InsertRefPlane( _
            4, 1.5707963267949, _
            0, 0, _
            0, 0)
    End If
    
    ' Last resort: Create plane using coordinates without reference
    If swRefPlane Is Nothing Then
        DebugLog "All reference methods failed, creating plane with coordinates..."
        swModel.ClearSelection2 True
        
        ' Create three points to define a plane
        swModel.CreatePoint2 0, 0, 0
        swModel.CreatePoint2 0, 1, 0
        swModel.CreatePoint2 1, 0, 0
        
        ' Create plane from 3 points
        Set swRefPlane = swFeatureMgr.InsertRefPlane(8, 0, 0, 0, 0, 0)
    End If
    
    ' Clear selections after creation
    swModel.ClearSelection2 True
    
    If Not swRefPlane Is Nothing Then
        ' Rename the plane to "Front Plane"
        swRefPlane.Name = "Front Plane"
        DebugLog "Successfully created new Front Plane"
        
        ' Force rebuild
        swModel.ForceRebuild3 False
    Else
        DebugLog "ERROR: All methods to create Front Plane failed"
        DebugLog "Last Error Code: " & swModel.GetLastError
    End If
    
    Set CreateFrontPlane = swRefPlane
End Function


'=== UPDATED: CREATE NEW LINE SKETCH ===
Private Function CreateNewLineSketch(model As ModelDoc2, configName As String, lineLengthMeters As Double, unitString As String, originalLength As Double) As Boolean
    
    DebugLog "=== Creating new line sketch ==="
    DebugLog "Creating line with length: " & lineLengthMeters & " meters"
    DebugLog "Original length: " & originalLength & " " & unitString
    
    ' === USE NEW ROBUST PLANE FUNCTION ===
    Dim swFrontPlane As Object
    Set swFrontPlane = GetOrCreateFrontPlane(model)
    
    If swFrontPlane Is Nothing Then
        DebugLog "CRITICAL ERROR: Cannot get or create Front Plane - cannot create sketch"
        CreateNewLineSketch = False
        Exit Function
    End If
    
    ' Select the Front Plane
    Dim boolStatus As Boolean
    boolStatus = swFrontPlane.Select2(False, 0)
    
    If Not boolStatus Then
        DebugLog "ERROR: Cannot select Front Plane even after finding/creating it"
        CreateNewLineSketch = False
        Exit Function
    End If
    
    DebugLog "Successfully selected Front Plane: " & swFrontPlane.Name
    
    ' Clear selection and insert sketch
    model.ClearSelection2 True
    boolStatus = swFrontPlane.Select2(False, 0)
    
    ' Insert sketch on selected plane
    model.InsertSketch2 True
    
    ' Create horizontal line from origin
    Dim swSketchMgr As SketchManager
    Set swSketchMgr = model.SketchManager
    
    ' Create line from (0,0,0) to (length,0,0) - all in METERS for API
    Dim swSketchSeg As SketchSegment
    Set swSketchSeg = swSketchMgr.CreateLine(0#, 0#, 0#, lineLengthMeters, 0#, 0#)
    
    If swSketchSeg Is Nothing Then
        DebugLog "ERROR: Failed to create line segment"
        CreateNewLineSketch = False
        Exit Function
    End If
    
    ' Exit sketch
    model.InsertSketch2 True
    
    ' Get the sketch feature that was just created
    Dim swSelMgr As SelectionMgr
    Set swSelMgr = model.SelectionManager
    
    Dim swFeat As Feature
    Set swFeat = model.FeatureByPositionReverse(0) ' Get most recently created feature
    
    If Not swFeat Is Nothing Then
        If swFeat.GetTypeName2 = "ProfileFeature" Then
            ' Rename the sketch
            Dim newName As String
            newName = "Line_" & configName
            swFeat.Name = newName
            DebugLog "Created and renamed sketch to: " & newName
            CreateNewLineSketch = True
        Else
            DebugLog "ERROR: Most recent feature is not a ProfileFeature, it's: " & swFeat.GetTypeName2
            CreateNewLineSketch = False
        End If
    Else
        DebugLog "ERROR: Could not get most recent feature"
        CreateNewLineSketch = False
    End If
    
    model.ClearSelection2 True
    
End Function


' Update existing line sketch
Private Sub UpdateExistingLineSketch(model As ModelDoc2, existingLineSketch As Feature, lineLengthMeters As Double)
    
    DebugLog "Updating existing line sketch with new length: " & lineLengthMeters & " meters"
    
    Dim boolEdit As Boolean
    boolEdit = existingLineSketch.Select2(False, 0)
    If boolEdit Then
        model.EditSketch
        
        ' Get the sketch from the feature
        Dim swSketch As Sketch
        Set swSketch = existingLineSketch.GetSpecificFeature2
        
        If Not swSketch Is Nothing Then
            ' Get sketch segments
            Dim sketchSegs As Variant
            sketchSegs = swSketch.GetSketchSegments
            
            If Not IsEmpty(sketchSegs) Then
                Dim seg As SketchSegment
                Set seg = sketchSegs(0) ' Get first segment (our line)
                
                If seg.GetType = 1 Then ' swSketchLINE
                    Dim sketchLine As sketchLine
                    Set sketchLine = seg
                    
                    ' Update the end point of the line
                    Dim endPoint As SketchPoint
                    Set endPoint = sketchLine.GetEndPoint2
                    endPoint.X = lineLengthMeters
                    endPoint.Y = 0
                    endPoint.Z = 0
                    
                    DebugLog "Updated line length to: " & lineLengthMeters & " meters"
                End If
            End If
        End If
        
        ' Exit sketch edit mode
        model.InsertSketch2 True
        model.ForceRebuild3 True
        DebugLog "Sketch modification completed"
    End If
    
End Sub



' Delete existing wrong structural member in FlatHose config ONLY
Private Sub DeleteExistingFlatHoseStructuralMember(model As ModelDoc2, configName As String)
    DebugLog "=== Checking for existing FlatHose structural member to replace ==="
    
    ' Make sure we're in FlatHose config
    If model.ConfigurationManager.ActiveConfiguration.Name <> configName Then
        model.ShowConfiguration2 configName
    End If
    
    ' Find the line sketch position to identify which structural member belongs to FlatHose
    Dim lineSketchName As String
    lineSketchName = "Line_" & configName
    
    Dim lineSketchPosition As Long
    lineSketchPosition = 0
    Dim featureCount As Long
    featureCount = 0
    Dim swFeat As Feature
    
    ' Find the line sketch position
    Set swFeat = model.FirstFeature
    Do While Not swFeat Is Nothing
        featureCount = featureCount + 1
        If swFeat.Name = lineSketchName Then
            lineSketchPosition = featureCount
            DebugLog "Line sketch found at position: " & lineSketchPosition
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    If lineSketchPosition = 0 Then
        DebugLog "Line sketch not found, cannot identify FlatHose structural member"
        Exit Sub
    End If
    
    ' Find structural members created after the line sketch (these belong to FlatHose)
    featureCount = 0
    Set swFeat = model.FirstFeature
    Do While Not swFeat Is Nothing
        featureCount = featureCount + 1
        If swFeat.GetTypeName2 = "WeldMemberFeat" And featureCount > lineSketchPosition Then
            ' This is a FlatHose structural member - check if it's active in this config
            If Not swFeat.IsSuppressed Then
                DebugLog "Found existing FlatHose structural member: " & swFeat.Name
                
                ' Delete it so we can create the correct one
                DebugLog "Deleting existing FlatHose structural member to replace with correct profile"
                swFeat.Select2 False, 0
                model.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed
                model.ClearSelection2 True
                model.ForceRebuild3 True
                
                DebugLog "Deleted: " & swFeat.Name
                Exit Do ' Only delete the first one found
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
End Sub

' === FIX 1: IMPROVED SUPPRESSION FUNCTION ===
Private Sub ConfigureFeatureVisibilityFixed(model As ModelDoc2, configName As String, lineSketchName As String)
    
    DebugLog "=== FIXED SUPPRESSION: Configuring feature visibility for " & configName & " ==="
    
    On Error GoTo ErrorHandler
    
    ' Get all configuration names
    Dim configNames As Variant
    configNames = model.GetConfigurationNames
    Dim originalConfigName As String
    Dim i As Integer
    
    ' Find the original configuration (not FlatHose)
    For i = 0 To UBound(configNames)
        If configNames(i) <> configName Then
            originalConfigName = configNames(i)
            Exit For
        End If
    Next i
    
    DebugLog "FIXED: Original config: " & originalConfigName & ", FlatHose config: " & configName
    
    ' Apply suppression while in each configuration with FIXED logic
    Call ConfigureWhileInEachConfigurationFixed(model, configName, originalConfigName, lineSketchName)
    
    ' Return to the target configuration and force final rebuild
    model.ShowConfiguration2 configName
    model.ForceRebuild3 True
    model.EditRebuild3
    
    ' Final verification
    Call FinalVisibilityCheckFixed(model, configName, originalConfigName, lineSketchName)
    
    DebugLog "FIXED: Feature visibility configuration completed"
    Exit Sub
    
ErrorHandler:
    DebugLog "ERROR in ConfigureFeatureVisibilityFixed: " & Err.description
    Resume Next
    
End Sub

' === FIX 1: IMPROVED CONFIG-SPECIFIC SUPPRESSION ===
Private Sub ConfigureWhileInEachConfigurationFixed(model As ModelDoc2, flatHoseConfigName As String, originalConfigName As String, lineSketchName As String)
    
    DebugLog "=== FIXED: APPLYING SUPPRESSION WHILE IN EACH CONFIG ==="
    
    ' Find line sketch position for categorization
    Dim lineSketchPosition As Long
    lineSketchPosition = GetLineSketchPosition(model, lineSketchName)
    DebugLog "Line sketch position: " & lineSketchPosition
    
    ' First, unsuppress everything in both configs to start fresh
    DebugLog "=== RESETTING ALL SUPPRESSIONS ==="
    model.ShowConfiguration2 flatHoseConfigName
    Call UnsuppressAllRelevantFeatures(model)
    
    model.ShowConfiguration2 originalConfigName
    Call UnsuppressAllRelevantFeatures(model)
    
    ' Configure FlatHose configuration
    DebugLog "=== CONFIGURING FLATHOSE CONFIG ==="
    model.ShowConfiguration2 flatHoseConfigName
    DebugLog "Switched to: " & model.ConfigurationManager.ActiveConfiguration.Name
    Call ApplySuppressionInActiveConfigFixed(model, lineSketchPosition, lineSketchName, True) ' True = is FlatHose config
    
    ' Configure Original configuration
    DebugLog "=== CONFIGURING ORIGINAL CONFIG ==="
    model.ShowConfiguration2 originalConfigName
    DebugLog "Switched to: " & model.ConfigurationManager.ActiveConfiguration.Name
    Call ApplySuppressionInActiveConfigFixed(model, lineSketchPosition, lineSketchName, False) ' False = is Original config
    
End Sub

' Helper to unsuppress all relevant features
Private Sub UnsuppressAllRelevantFeatures(model As ModelDoc2)
    DebugLog "Unsuppressing all relevant features in config: " & model.ConfigurationManager.ActiveConfiguration.Name
    
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        Dim featureType As String
        featureType = swFeat.GetTypeName2
        
        If featureType = "WeldMemberFeat" Or featureType = "ProfileFeature" Or featureType = "3DProfileFeature" Then
            If swFeat.IsSuppressed Then
                DebugLog "  Unsuppressing: " & swFeat.Name
                swFeat.SetSuppression 1  ' Unsuppress
            End If
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    model.ForceRebuild3 True
End Sub

' Get the position of the line sketch
Private Function GetLineSketchPosition(model As ModelDoc2, lineSketchName As String) As Long
    
    Dim featureCount As Long
    featureCount = 0
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        featureCount = featureCount + 1
        If swFeat.Name = lineSketchName Then
            GetLineSketchPosition = featureCount
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    GetLineSketchPosition = 0
    
End Function

' === FIX 1: FIXED SUPPRESSION LOGIC ===
Private Sub ApplySuppressionInActiveConfigFixed(model As ModelDoc2, lineSketchPosition As Long, lineSketchName As String, isFlatHoseConfig As Boolean)
    
    Dim activeConfigName As String
    activeConfigName = model.ConfigurationManager.ActiveConfiguration.Name
    DebugLog "FIXED: Applying suppression in active config: " & activeConfigName & " (FlatHose=" & isFlatHoseConfig & ")"
    
    Dim featureCount As Long
    featureCount = 0
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        featureCount = featureCount + 1
        
        Dim featureType As String
        featureType = swFeat.GetTypeName2
        
        ' Process relevant features
        If featureType = "WeldMemberFeat" Or featureType = "ProfileFeature" Or featureType = "3DProfileFeature" Then
            
            DebugLog "FIXED: Processing " & swFeat.Name & " (Type: " & featureType & ", Position: " & featureCount & ")"
            DebugLog "  Current state: IsSuppressed = " & swFeat.IsSuppressed
            
            Dim shouldSuppress As Boolean
            shouldSuppress = False
            
            Select Case featureType
                
                Case "ProfileFeature", "3DProfileFeature"
                    If swFeat.Name = lineSketchName Then
                        ' THIS IS THE FIX: The Line_* sketch should be suppressed in ALL configs
                        ' after the structural member is created - it's no longer needed
                        shouldSuppress = True
                        DebugLog "  FIXED: FlatHose line sketch (" & lineSketchName & ") - suppress in ALL configs"
                        
                    ElseIf swFeat.Name Like "Line_*" Then
                        ' Catch any Line_* sketches that might not have been renamed properly
                        shouldSuppress = True
                        DebugLog "  FIXED: Line_* sketch (" & swFeat.Name & ") - suppress in ALL configs"
                        
                    ElseIf featureCount > lineSketchPosition And lineSketchPosition > 0 Then
                        ' IMPROVED: Any sketch created AFTER the line sketch position is likely a FlatHose sketch
                        ' This catches sketches that weren't properly renamed (like Sketch12)
                        shouldSuppress = True
                        DebugLog "  FIXED: Post-line sketch (" & swFeat.Name & ") - suppress in ALL configs"
                        
                    ElseIf swFeat.Name = "3DSketch1" Or (InStr(swFeat.Name, "3DSketch") > 0) Then
                        ' Original 3D sketch - visible in Original, suppress in FlatHose
                        shouldSuppress = isFlatHoseConfig
                        DebugLog "  FIXED: Original 3D sketch - suppress in FlatHose (" & shouldSuppress & ")"
                        
                    ElseIf swFeat.Name Like "Sketch*" And featureCount < lineSketchPosition Then
                        ' Any other sketches created before the line sketch (original sketches)
                        ' Visible in Original, suppress in FlatHose
                        shouldSuppress = isFlatHoseConfig
                        DebugLog "  FIXED: Original sketch - suppress in FlatHose (" & shouldSuppress & ")"
                    End If
                    
                Case "WeldMemberFeat"
                    If lineSketchPosition > 0 And featureCount > lineSketchPosition Then
                        ' Feature created after line sketch = FlatHose structural member
                        ' Visible in FlatHose, suppress in Original
                        shouldSuppress = Not isFlatHoseConfig
                        DebugLog "  FIXED: FlatHose SM (after line sketch) - suppress in Original (" & shouldSuppress & ")"
                    Else
                        ' Original structural member
                        ' Visible in Original, suppress in FlatHose
                        shouldSuppress = isFlatHoseConfig
                        DebugLog "  FIXED: Original SM - suppress in FlatHose (" & shouldSuppress & ")"
                    End If
                    
            End Select
            
            ' Apply the suppression decision
            If shouldSuppress Then
                If Not swFeat.IsSuppressed Then
                    DebugLog "  FIXED: SUPPRESSING feature"
                    swFeat.SetSuppression 0  ' 0 = swSuppressed
                End If
            Else
                If swFeat.IsSuppressed Then
                    DebugLog "  FIXED: UNSUPPRESSING feature"
                    swFeat.SetSuppression 1  ' 1 = swUnSuppressed
                End If
            End If
            
            DebugLog "  FIXED: Final state: IsSuppressed = " & swFeat.IsSuppressed
            DebugLog "  ---"
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    ' Force final rebuild in this configuration
    model.ForceRebuild3 True
    model.EditRebuild3
    DebugLog "FIXED: Completed configuration: " & activeConfigName
    
End Sub

' === FIX 1: ENHANCED FINAL VISIBILITY CHECK ===
Private Sub FinalVisibilityCheckFixed(model As ModelDoc2, flatHoseConfigName As String, originalConfigName As String, lineSketchName As String)
    
    DebugLog "=== FIXED: ENHANCED FINAL VISIBILITY CHECK ==="
    
    ' Force rebuilds before checking
    model.ShowConfiguration2 flatHoseConfigName
    model.ForceRebuild3 True
    model.EditRebuild3
    
    model.ShowConfiguration2 originalConfigName
    model.ForceRebuild3 True
    model.EditRebuild3
    
    ' Check FlatHose configuration
    DebugLog "=== FIXED: CHECKING FLATHOSE CONFIG (AFTER REBUILD) ==="
    model.ShowConfiguration2 flatHoseConfigName
    DebugLog "Active config: " & model.ConfigurationManager.ActiveConfiguration.Name
    
    Dim flatHoseVisibleSM As Long
    flatHoseVisibleSM = 0
    Dim flatHoseVisibleSketches As Long
    flatHoseVisibleSketches = 0
    
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        If Not swFeat.IsSuppressed Then
            If swFeat.GetTypeName2 = "WeldMemberFeat" Then
                flatHoseVisibleSM = flatHoseVisibleSM + 1
                DebugLog "FLATHOSE VISIBLE SM: " & swFeat.Name
            ElseIf swFeat.GetTypeName2 = "ProfileFeature" Or swFeat.GetTypeName2 = "3DProfileFeature" Then
                flatHoseVisibleSketches = flatHoseVisibleSketches + 1
                DebugLog "FLATHOSE VISIBLE SKETCH: " & swFeat.Name & " *** SHOULD BE ZERO ***"
            End If
        Else
            If swFeat.GetTypeName2 = "WeldMemberFeat" Then
                DebugLog "FLATHOSE SUPPRESSED SM: " & swFeat.Name
            ElseIf swFeat.GetTypeName2 = "ProfileFeature" Or swFeat.GetTypeName2 = "3DProfileFeature" Then
                DebugLog "FLATHOSE SUPPRESSED SKETCH: " & swFeat.Name
                If swFeat.Name = lineSketchName Then
                    DebugLog "  *** GOOD: Line sketch is suppressed ***"
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    DebugLog "FIXED FLATHOSE FINAL SUMMARY:"
    DebugLog "  Visible Structural Members: " & flatHoseVisibleSM & " (EXPECTED: 1)"
    DebugLog "  Visible Sketches: " & flatHoseVisibleSketches & " (EXPECTED: 0)"
    
    If flatHoseVisibleSM = 1 And flatHoseVisibleSketches = 0 Then
        DebugLog "  ? FLATHOSE CONFIG IS CORRECT!"
    Else
        DebugLog "  ? FLATHOSE CONFIG IS WRONG!"
    End If
    
    ' Check Original configuration
    DebugLog "=== FIXED: CHECKING ORIGINAL CONFIG (AFTER REBUILD) ==="
    model.ShowConfiguration2 originalConfigName
    DebugLog "Active config: " & model.ConfigurationManager.ActiveConfiguration.Name
    
    Dim originalVisibleSM As Long
    originalVisibleSM = 0
    Dim originalVisibleSketches As Long
    originalVisibleSketches = 0
    Dim originalLineSketchVisible As Boolean
    originalLineSketchVisible = False
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        If Not swFeat.IsSuppressed Then
            If swFeat.GetTypeName2 = "WeldMemberFeat" Then
                originalVisibleSM = originalVisibleSM + 1
                DebugLog "ORIGINAL VISIBLE SM: " & swFeat.Name
            ElseIf swFeat.GetTypeName2 = "ProfileFeature" Or swFeat.GetTypeName2 = "3DProfileFeature" Then
                originalVisibleSketches = originalVisibleSketches + 1
                DebugLog "ORIGINAL VISIBLE SKETCH: " & swFeat.Name
                If swFeat.Name = lineSketchName Then
                    originalLineSketchVisible = True
                    DebugLog "  *** BAD: Line sketch should be suppressed in original too! ***"
                End If
            End If
        Else
            If swFeat.GetTypeName2 = "WeldMemberFeat" Then
                DebugLog "ORIGINAL SUPPRESSED SM: " & swFeat.Name
            ElseIf swFeat.GetTypeName2 = "ProfileFeature" Or swFeat.GetTypeName2 = "3DProfileFeature" Then
                DebugLog "ORIGINAL SUPPRESSED SKETCH: " & swFeat.Name
                If swFeat.Name = lineSketchName Then
                    DebugLog "  *** GOOD: Line sketch is suppressed in original ***"
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    DebugLog "FIXED ORIGINAL FINAL SUMMARY:"
    DebugLog "  Visible Structural Members: " & originalVisibleSM & " (EXPECTED: 1)"
    DebugLog "  Visible Sketches: " & originalVisibleSketches & " (EXPECTED: 1, but NOT the line sketch)"
    DebugLog "  Line sketch visible: " & originalLineSketchVisible & " (EXPECTED: False)"
    
    If originalVisibleSM = 1 And originalVisibleSketches >= 1 And Not originalLineSketchVisible Then
        DebugLog "  ? ORIGINAL CONFIG IS CORRECT!"
    Else
        DebugLog "  ? ORIGINAL CONFIG IS WRONG!"
    End If
    
End Sub

' Public function to check if hose geometry should be created
Public Function ShouldCreateHoseGeometry() As Boolean
    ' Return True to always create geometry for hose parts
    ShouldCreateHoseGeometry = True
End Function

' Function to detect if this is a hose part based on hose-specific profiles
Public Function IsHosePart(model As ModelDoc2) As Boolean
    
    ' Check if the part has structural members with hose-like profiles
    Dim swFeat As Feature
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "WeldMemberFeat" Then
            ' Get the feature definition
            Dim swStructuralMemberFeatData As StructuralMemberFeatureData
            Set swStructuralMemberFeatData = swFeat.GetDefinition
            
            If Not swStructuralMemberFeatData Is Nothing Then
                Dim profilePath As String
                profilePath = swStructuralMemberFeatData.WeldmentProfilePath
                
                ' Check if profile path contains keywords suggesting it's a hose
                If InStr(UCase(profilePath), "HOSE") > 0 Or _
                   InStr(UCase(profilePath), "HYDRAULIC") > 0 Then
                    DebugLog "Detected hose part based on profile: " & profilePath
                    IsHosePart = True
                    Exit Function
                End If
            End If
        End If
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    ' Could also check custom properties for hose-related keywords
    ' Or check the part name/description
    
    IsHosePart = False
    
End Function

' Helper function to convert decimal to fraction (simplified version)
Private Function ConvertDecimalToFraction(decimalInches As Double) As String
    Dim wholePart As Long
    Dim fractionPart As Double
    Dim numerator As Long
    Dim denominator As Long
    
    wholePart = Int(decimalInches)
    fractionPart = decimalInches - wholePart
    
    ' Convert to nearest 16th
    denominator = 16
    numerator = Round(fractionPart * denominator)
    
    ' Simplify fraction
    Dim gcd As Long
    gcd = GetGCD(numerator, denominator)
    If gcd > 0 Then
        numerator = numerator / gcd
        denominator = denominator / gcd
    End If
    
    ' Build result string
    If numerator = 0 Then
        ConvertDecimalToFraction = CStr(wholePart)
    ElseIf wholePart = 0 Then
        ConvertDecimalToFraction = numerator & "/" & denominator
    Else
        ConvertDecimalToFraction = wholePart & " " & numerator & "/" & denominator
    End If
End Function

' Helper function to get greatest common divisor
Private Function GetGCD(a As Long, b As Long) As Long
    If b = 0 Then
        GetGCD = a
    Else
        GetGCD = GetGCD(b, a Mod b)
    End If
End Function

