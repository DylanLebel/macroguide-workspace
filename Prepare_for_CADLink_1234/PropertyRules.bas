Attribute VB_Name = "PropertyRules"
Option Explicit
Option Private Module

' ====================================================================================
' PropertyRules.bas - Property Validation Rules
' ====================================================================================
'
' PURPOSE:
'   Defines required properties and their validation regex patterns for Parts
'   and Assemblies. Used by AddProperties.bas and Main.bas for validation.
'
' USAGE:
'   Dim rules As Object
'   Set rules = GetRequiredPropertyRules(swDocPART)
'   ' or
'   Set rules = GetRequiredPartPropertyRules()
'   Set rules = GetRequiredAssemblyPropertyRules()
'
' ADDING NEW PROPERTIES:
'   1. Add to COMMON_RULES if required for both Parts and Assemblies
'   2. Add to PART_ONLY_RULES if only required for Parts
'   3. Add to ASSEMBLY_ONLY_RULES if only required for Assemblies
'   4. Use regex pattern: ".+" means required, ".*" means optional
'
' ====================================================================================

' ===========================
' VALIDATION PATTERNS
' ===========================
' Common patterns used across multiple properties
Private Const PATTERN_REQUIRED As String = ".+"              ' Any non-empty value
Private Const PATTERN_OPTIONAL As String = ".*"              ' Any value including empty
Private Const PATTERN_NUMBER As String = "^\d+$"             ' Digits only
Private Const PATTERN_TYPE As String = "^(P|M)$"             ' P or M only
Private Const PATTERN_PRODUCT_GROUP As String = "^(CHUTES|CONVSHEA|LOCO|MCUNST|MISC|SHAFTSTL|SPARE|WARRENTY|SME)$"
Private Const PATTERN_CLASS As String = "^(ASMW|ELEC|FGDS|HDWE|HYDR|MECH|PNEU|RAWM)$"
Private Const PATTERN_NDT As String = "^(NONE|MPI|LPI|UT|UT MPI|UT LPI)$"
Private Const PATTERN_MTL_UOM As String = "MM|IN|SQMM|SQIN"

' Reference Category patterns (Part has more options than Assembly)
Private Const PATTERN_REF_CAT_PART As String = "^(AB|AR|AS|ASSEMBLY|BRG|C|CHN|CP|ELB|ELC|EXP|FB|FW|GR|H|HSS|HUC|HYD|I|L|M|MB|MC|MEC|MFG|MI|NUT|PART|PI|PIN|PL|PNEU|PNT|PUR|RB|RIV|S|SQ|THD|TR|UB|UC|W|WB|WD|WELDMENT|WRE|WS|WSR|WT|Z|HOSE|SHIM|KS|WD|PIN)$"
Private Const PATTERN_REF_CAT_ASSY As String = "^(AB|AR|ASSEMBLY|BRG|C|CHN|CP|ELB|ELC|EXP|FB|FW|GR|H|HSS|HUC|HYD|I|L|M|MB|MC|MEC|MFG|MI|NUT|PART|PI|PIN|PL|PNEU|PNT|PUR|RB|RIV|S|SQ|THD|TR|UB|UC|W|WB|WD|WELDMENT|WRE|WS|WSR|WT|Z|HOSE)$"

' ===========================
' UNIFIED ENTRY POINT
' ===========================
' Use this function for cleaner code - pass model type to get appropriate rules

Function GetRequiredPropertyRules(modelType As Long) As Variant
    ' Returns the appropriate rules based on model type
    ' modelType: swDocPART (1) or swDocASSEMBLY (2)

    If modelType = swDocPART Then
        Set GetRequiredPropertyRules = GetRequiredPartPropertyRules()
    ElseIf modelType = swDocASSEMBLY Then
        Set GetRequiredPropertyRules = GetRequiredAssemblyPropertyRules()
    Else
        ' Return empty dictionary for unsupported types
        Set GetRequiredPropertyRules = CreateObject("Scripting.Dictionary")
    End If
End Function

' ===========================
' PART PROPERTY RULES
' ===========================

Function GetRequiredPartPropertyRules() As Variant
    Dim rules As Object
    Set rules = CreateObject("Scripting.Dictionary")

    ' Add common properties first (shared with assemblies)
    AddCommonRules rules

    ' Add part-specific properties
    rules.Add "Reference Category", PATTERN_REF_CAT_PART
    rules.Add "Material", PATTERN_REQUIRED
    rules.Add "Length", PATTERN_REQUIRED
    rules.Add "LengthA", PATTERN_REQUIRED
    rules.Add "Height", PATTERN_REQUIRED
    rules.Add "Width", PATTERN_REQUIRED
    rules.Add "MtlUOM", PATTERN_MTL_UOM
    rules.Add "Created By", PATTERN_REQUIRED

    Set GetRequiredPartPropertyRules = rules
End Function

' ===========================
' ASSEMBLY PROPERTY RULES
' ===========================

Function GetRequiredAssemblyPropertyRules() As Variant
    Dim rules As Object
    Set rules = CreateObject("Scripting.Dictionary")

    ' Add common properties first (shared with parts)
    AddCommonRules rules

    ' Add assembly-specific properties
    rules.Add "Reference Category", PATTERN_REF_CAT_ASSY

    Set GetRequiredAssemblyPropertyRules = rules
End Function

' ===========================
' COMMON RULES (shared between Part and Assembly)
' ===========================

Private Sub AddCommonRules(ByRef rules As Object)
    ' These properties are required for BOTH parts and assemblies

    ' Identification
    rules.Add "Model Number", PATTERN_REQUIRED        ' 7 digit number beginning with 9
    rules.Add "Revision", PATTERN_NUMBER              ' Numeric revision
    rules.Add "Description", PATTERN_REQUIRED         ' Part/Assembly description
    rules.Add "Stock Size", PATTERN_REQUIRED          ' Stock size specification

    ' Classification
    rules.Add "Type", PATTERN_TYPE                    ' P (Purchased) or M (Manufactured)
    rules.Add "Product Group", PATTERN_PRODUCT_GROUP  ' Product category
    rules.Add "Class", PATTERN_CLASS                  ' Item class

    ' Physical Properties
    rules.Add "Weight", PATTERN_REQUIRED              ' Weight value
    rules.Add "Weight UOM", PATTERN_REQUIRED          ' Weight unit of measure
    rules.Add "Dimensional UOM", PATTERN_REQUIRED     ' Dimensional unit of measure

    ' Quality/Testing
    rules.Add "Non-Destructive Testing", PATTERN_NDT  ' NDT requirements

    ' Metadata
    rules.Add "Comments", PATTERN_OPTIONAL            ' Optional comments
    rules.Add "Author", PATTERN_REQUIRED              ' Who created this
    rules.Add "Date", PATTERN_REQUIRED                ' Creation date
End Sub

' ===========================
' HELPER FUNCTIONS
' ===========================

' Check if a property name is required for a given model type
Function IsPropertyRequired(propertyName As String, modelType As Long) As Boolean
    Dim rules As Object
    Set rules = GetRequiredPropertyRules(modelType)

    If rules.exists(propertyName) Then
        ' If pattern is not ".*" (optional), it's required
        IsPropertyRequired = (rules(propertyName) <> PATTERN_OPTIONAL)
    Else
        IsPropertyRequired = False
    End If
End Function

' Get the validation pattern for a specific property
Function GetPropertyPattern(propertyName As String, modelType As Long) As String
    Dim rules As Object
    Set rules = GetRequiredPropertyRules(modelType)

    If rules.exists(propertyName) Then
        GetPropertyPattern = rules(propertyName)
    Else
        GetPropertyPattern = ""
    End If
End Function

' Get list of all required property names for a model type
Function GetRequiredPropertyNames(modelType As Long) As collection
    Dim rules As Object
    Dim names As New collection
    Dim key As Variant

    Set rules = GetRequiredPropertyRules(modelType)

    For Each key In rules.keys
        If rules(key) <> PATTERN_OPTIONAL Then
            names.Add key
        End If
    Next key

    Set GetRequiredPropertyNames = names
End Function
