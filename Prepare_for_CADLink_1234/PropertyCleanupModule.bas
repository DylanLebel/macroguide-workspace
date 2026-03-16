Attribute VB_Name = "PropertyCleanupModule"
Option Explicit
Option Private Module

' ============================================================================
' MODULE: PropertyCleanupModule
' PURPOSE: Standardize property names to match PropertyRules capitalization
' USAGE: Call CleanupPropertyNames(model, propertiesToSet)
' ============================================================================

Public Sub CleanupPropertyNames(model As ModelDoc2, ByRef propertiesToSet As Scripting.Dictionary)
    On Error Resume Next
    
    DebugLog ""
    DebugLog "========================================================"
    DebugLog "=== STARTING CleanupPropertyNames ==="
    DebugLog "========================================================"
    DebugLog "Input dictionary count: " & propertiesToSet.Count
    
    ' Check if dictionary is valid
    If propertiesToSet Is Nothing Then
        DebugLog "ERROR: propertiesToSet dictionary is Nothing!"
        Exit Sub
    End If
    
    If propertiesToSet.Count = 0 Then
        DebugLog "WARNING: propertiesToSet dictionary is empty!"
        Exit Sub
    End If
    
    ' Get the correct property rules based on model type
    Dim PropertyRules As Object
    If model.GetType = swDocPART Then
        Set PropertyRules = GetRequiredPartPropertyRules()
        DebugLog "Using PART property rules"
    ElseIf model.GetType = swDocASSEMBLY Then
        Set PropertyRules = GetRequiredAssemblyPropertyRules()
        DebugLog "Using ASSEMBLY property rules"
    Else
        DebugLog "ERROR: Unknown model type: " & model.GetType
        Exit Sub
    End If
    
    If PropertyRules Is Nothing Then
        DebugLog "ERROR: propertyRules is Nothing!"
        Exit Sub
    End If
    
    DebugLog "Property rules count: " & PropertyRules.Count
    
    ' ============================================================================
    ' CHECK FOR DUPLICATE PROPERTIES (case-insensitive)
    ' ============================================================================
    DebugLog ""
    DebugLog "--- CHECKING FOR POTENTIAL DUPLICATES ---"
    Dim tempCheck As New Scripting.Dictionary
    Dim checkKey As Variant  ' ? Different variable name!
    For Each checkKey In propertiesToSet.keys
        Dim upperKey As String
        upperKey = UCase(Trim(CStr(checkKey)))
        
        If tempCheck.exists(upperKey) Then
            DebugLog "WARNING: Duplicate found!"
            DebugLog "  First: '" & tempCheck(upperKey) & "'"
            DebugLog "  Second: '" & checkKey & "'"
        Else
            tempCheck(upperKey) = CStr(checkKey)
        End If
    Next checkKey
    DebugLog "--- END DUPLICATE CHECK ---"
    DebugLog ""
    ' ============================================================================

    ' Show all rules
    DebugLog "--- PROPERTY RULES (correct capitalization) ---"
    Dim ruleKey As Variant
    For Each ruleKey In PropertyRules.keys
        DebugLog "  Rule: '" & ruleKey & "'"
    Next ruleKey
    
    ' Create a new dictionary with corrected names
    Dim cleanedProps As New Scripting.Dictionary
    DebugLog "Created new cleanedProps dictionary"
    
    ' Loop through existing properties
    Dim key As Variant
    Dim propertyCount As Long
    propertyCount = 0
    
    DebugLog ""
    DebugLog "--- PROCESSING EACH PROPERTY ---"
    
    For Each key In propertiesToSet.keys
        propertyCount = propertyCount + 1
        
        Dim existingName As String
        Dim correctName As String
        Dim value As String
        
        existingName = CStr(key)
        value = propertiesToSet(existingName)
        
        DebugLog ""
        DebugLog "Property #" & propertyCount & ":"
        DebugLog "  Existing name: '" & existingName & "'"
        DebugLog "  Value: '" & value & "'"
        
        ' Find the correct name from property rules (case-insensitive match)
        correctName = FindCorrectPropertyName(existingName, PropertyRules)
        
        DebugLog "  FindCorrectPropertyName returned: '" & correctName & "'"
        
       ' Use the correct name if found, otherwise keep the original
If correctName <> "" Then
    DebugLog "  --> CORRECTING to: '" & correctName & "'"
    
    ' *** CHECK FOR COLLISION BEFORE ADDING ***
    If cleanedProps.exists(correctName) Then
        DebugLog "  *** COLLISION WARNING! Property '" & correctName & "' already exists in cleanedProps!"
        DebugLog "      Existing value: '" & cleanedProps(correctName) & "'"
        DebugLog "      New value would be: '" & value & "'"
        
        ' Prefer non-empty values - overwrite if new value is not empty
        If Trim(value) <> "" Then
            DebugLog "      NEW VALUE IS NOT EMPTY - OVERWRITING with new value"
            cleanedProps(correctName) = value
        ElseIf Trim(cleanedProps(correctName)) = "" Then
            DebugLog "      BOTH VALUES ARE EMPTY - keeping existing"
        Else
            DebugLog "      EXISTING VALUE IS NOT EMPTY - keeping existing value"
        End If
    Else
        cleanedProps(correctName) = value
    End If
        Else
            DebugLog "  --> NO MATCH - keeping original: '" & existingName & "'"
            cleanedProps(existingName) = value
        End If
    Next key
    
    DebugLog ""
    DebugLog "--- CLEANUP SUMMARY ---"
    DebugLog "Original dictionary count: " & propertiesToSet.Count
    DebugLog "Cleaned dictionary count: " & cleanedProps.Count
    
    If propertiesToSet.Count <> cleanedProps.Count Then
        DebugLog "ERROR: Dictionary counts don't match!"
        DebugLog "This means some properties are being lost/overwritten!"
    End If
    
    ' Replace the original dictionary
    DebugLog "Replacing original dictionary with cleaned dictionary..."
    Set propertiesToSet = cleanedProps
    DebugLog "Dictionary replaced. New count: " & propertiesToSet.Count
    
    DebugLog "========================================================"
    DebugLog "=== FINISHED CleanupPropertyNames ==="
    DebugLog "========================================================"
    DebugLog ""
    
    On Error GoTo 0
End Sub

' Find the correct property name from the rules (case-insensitive)
Private Function FindCorrectPropertyName(existingName As String, PropertyRules As Object) As String
    DebugLog "    [FindCorrectPropertyName called]"
    DebugLog "    Input: '" & existingName & "'"
    
    Dim ruleName As Variant
    
    ' Clean up the existing name for comparison
    Dim cleanExisting As String
    cleanExisting = UCase(Trim(existingName))
    DebugLog "    Uppercase version: '" & cleanExisting & "'"
    
    ' Look through all property rules for a case-insensitive match
    Dim matchCount As Long
    matchCount = 0
    
    For Each ruleName In PropertyRules.keys
        Dim ruleNameUpper As String
        ruleNameUpper = UCase(Trim(CStr(ruleName)))
        
        If ruleNameUpper = cleanExisting Then
            matchCount = matchCount + 1
            DebugLog "    MATCH FOUND!"
            DebugLog "    Rule name: '" & ruleName & "'"
            DebugLog "    Rule name (upper): '" & ruleNameUpper & "'"
            DebugLog "    Existing (upper): '" & cleanExisting & "'"
            
            ' Found a match! Return the correctly-capitalized name from the rules
            FindCorrectPropertyName = CStr(ruleName)
            DebugLog "    Returning: '" & FindCorrectPropertyName & "'"
            Exit Function
        End If
    Next ruleName
    
    ' No match found
    DebugLog "    No match found after checking all rules"
    FindCorrectPropertyName = ""
End Function
