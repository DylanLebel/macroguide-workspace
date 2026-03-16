Attribute VB_Name = "AddProperties"
Option Explicit
Option Private Module

Sub AddPropertiesBasedOnRules(model As ModelDoc2, ByRef propertiesToSet As Object)
    ' Determine if the model is a part or an assembly
    Dim PropertyRules As Object

    If model.GetType = swDocumentTypes_e.swDocPART Then
        Set PropertyRules = GetRequiredPartPropertyRules()
    ElseIf model.GetType = swDocumentTypes_e.swDocASSEMBLY Then
        Set PropertyRules = GetRequiredAssemblyPropertyRules()
    Else
        Exit Sub
    End If

    ' Sort the properties
    Dim sortedProperties As Scripting.Dictionary
    Set sortedProperties = SortProperties(propertiesToSet, PropertyRules)

    ' Replace the original propertiesToSet with the sorted properties
    Set propertiesToSet = sortedProperties

    ' NOTE: Properties are applied later by ApplyCollectedProperties in Main.bas
    ' This function only prepares and sorts the propertiesToSet dictionary
End Sub

Function SortProperties(propertiesToSet As Scripting.Dictionary, PropertyRules As Object) As Scripting.Dictionary
    On Error GoTo ErrorHandler

    DebugLog "Starting SortProperties..."

    If propertiesToSet Is Nothing Then
        DebugLog "ERROR: propertiesToSet is Nothing"
        Set SortProperties = New Scripting.Dictionary
        Exit Function
    End If

    If PropertyRules Is Nothing Then
        DebugLog "ERROR: propertyRules is Nothing"
        Set SortProperties = propertiesToSet
        Exit Function
    End If
    
    Dim sortedDict As New Scripting.Dictionary
    
    ' First add properties from the rules in order
    Dim ruleProp As Variant
    For Each ruleProp In PropertyRules.keys
        If propertiesToSet.exists(ruleProp) Then
            sortedDict.Add ruleProp, propertiesToSet(ruleProp)
        Else
            sortedDict.Add ruleProp, ""
        End If
    Next
    
    ' Then add any remaining properties
    Dim prop As Variant
    For Each prop In propertiesToSet.keys
        If Not sortedDict.exists(prop) Then
            sortedDict.Add prop, propertiesToSet(prop)
        End If
    Next
    
    Set SortProperties = sortedDict
    DebugLog "SortProperties completed successfully"
    Exit Function

ErrorHandler:
    DebugLog "ERROR in SortProperties: " & Err.description & " (Line: " & Erl & ")"
    DebugLog "Error Number: " & Err.Number
    Set SortProperties = propertiesToSet
    Resume Next
End Function

' Apply properties to model - NON-DESTRUCTIVE version
' Preserves user-added custom properties that aren't in propertiesToSet or propertyRules
Sub ApplyPropertiesToModel(model As ModelDoc2, PropertyRules As Object, ByRef propertiesToSet As Object)
    Dim swCustPropMgr As customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")

    ' Get existing properties and their values
    Dim existingProps As Variant
    existingProps = swCustPropMgr.GetNames

    ' Build dictionary of existing values and identify user properties
    Dim existingValues As New Scripting.Dictionary
    Dim userProps As New Scripting.Dictionary
    Dim i As Long

    If Not IsEmpty(existingProps) Then
        Dim propVal As String
        Dim resolvedVal As String

        For i = 0 To UBound(existingProps)
            Dim existingPropName As String
            existingPropName = CStr(existingProps(i))

            ' Get the current value
            swCustPropMgr.Get2 existingPropName, propVal, resolvedVal
            existingValues.Add existingPropName, resolvedVal

            ' Identify user properties (not in propertiesToSet and not in propertyRules)
            If Not propertiesToSet.exists(existingPropName) Then
                If Not PropertyRules.exists(existingPropName) Then
                    userProps.Add existingPropName, resolvedVal
                End If
            End If
        Next i
    End If

    ' Delete only managed properties (those in propertiesToSet or propertyRules)
    If Not IsEmpty(existingProps) Then
        For i = 0 To UBound(existingProps)
            Dim propToCheck As String
            propToCheck = CStr(existingProps(i))

            If propertiesToSet.exists(propToCheck) Or PropertyRules.exists(propToCheck) Then
                swCustPropMgr.Delete2 propToCheck
            End If
        Next i
    End If

    ' Add properties in the sorted order
    Dim propName As Variant
    For Each propName In propertiesToSet.keys
        Dim propertyName As String
        propertyName = CStr(propName)
        Dim propertyValue As String
        propertyValue = CStr(propertiesToSet(propName))

        swCustPropMgr.Add3 propertyName, swCustomInfoText, propertyValue, swCustomPropertyDeleteAndAdd
    Next propName

    ' Re-add preserved user properties at the end
    Dim userPropName As Variant
    For Each userPropName In userProps.keys
        swCustPropMgr.Add3 CStr(userPropName), swCustomInfoText, CStr(userProps(userPropName)), swCustomPropertyDeleteAndAdd
    Next userPropName
End Sub

' Add a new property to the model
Sub AddPropertyToModel(model As ModelDoc2, propertyName As String, propertyValue As String)
    Dim swCustPropMgr As customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")
    swCustPropMgr.Add3 propertyName, swCustomInfoText, propertyValue, swCustomPropertyDeleteAndAdd
End Sub

' Update an existing property in the model
Sub UpdatePropertyInModel(model As ModelDoc2, propertyName As String, propertyValue As String)
    Dim swCustPropMgr As customPropertyManager
    Set swCustPropMgr = model.Extension.customPropertyManager("")
    swCustPropMgr.Set2 propertyName, propertyValue
End Sub

' Check if a property exists in a model
Function propertyExists(model As ModelDoc2, propertyName As String) As Boolean
    Dim customPropertyManager As SldWorks.customPropertyManager
    Set customPropertyManager = model.Extension.customPropertyManager("")
    Dim value As String, resolvedValue As String, wasResolved As Boolean, linkedToProperty As Boolean, result As Integer
    result = customPropertyManager.Get6(propertyName, False, value, resolvedValue, wasResolved, linkedToProperty)
    propertyExists = (result = 0 Or result = 2)
End Function

' Add a property to the propertiesToSet dictionary
Function AddPropertyToSet(ByRef propertiesToSet As Object, propertyName As String) As Boolean
    If Not propertiesToSet.exists(propertyName) Then
        propertiesToSet.Add propertyName, ""
        AddPropertyToSet = True
    Else
        AddPropertyToSet = False
    End If
End Function
