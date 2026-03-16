Attribute VB_Name = "WeightModule"
Option Explicit
Option Private Module

Public Sub AutofillWeight(model As ModelDoc2, ByRef propertiesToSet As Object)

    If model Is Nothing Then
        DebugLog "AutofillWeight: No model provided"
        Exit Sub
    End If

    If propertiesToSet Is Nothing Then
        Set propertiesToSet = CreateObject("Scripting.Dictionary")
    End If

    Dim modelType As Long
    modelType = model.GetType

    Dim swModelExt As ModelDocExtension
    Set swModelExt = model.Extension

    Dim refModel As ModelDoc2
    If modelType = swDocumentTypes_e.swDocDRAWING Then
        Dim swDraw As drawingDoc
        Set swDraw = model

        Dim swView As view
        Set swView = swDraw.GetFirstView
        If swView Is Nothing Then
            DebugLog "AutofillWeight: No views in drawing - setting weight to 0"
            GoTo SetDefaultWeight
        End If

        Set swView = swView.GetNextView ' Skip sheet format
        If swView Is Nothing Then
            DebugLog "AutofillWeight: No model view found - setting weight to 0"
            GoTo SetDefaultWeight
        End If

        Set refModel = swView.ReferencedDocument
        If refModel Is Nothing Then
            DebugLog "AutofillWeight: Referenced model not loaded - setting weight to 0"
            GoTo SetDefaultWeight
        End If
    Else
        Set refModel = model
    End If

    ' Get mass
    Dim massPropMgr As MassProperty
    On Error Resume Next
    Set massPropMgr = refModel.Extension.CreateMassProperty()
    If massPropMgr Is Nothing Or Err.Number <> 0 Then
        DebugLog "AutofillWeight: Error initializing MassProperty - setting weight to 0"
        Err.Clear
        GoTo SetDefaultWeight
    End If
    Dim massKg As Double
    massKg = massPropMgr.Mass
    If Err.Number <> 0 Then
        DebugLog "AutofillWeight: Error getting mass - setting weight to 0"
        Err.Clear
        GoTo SetDefaultWeight
    End If
    On Error GoTo 0

    DebugLog "Mass in kilograms: " & massKg

    ' Get UOM
    Dim weightUOM As String
    If propertiesToSet.exists("Weight UOM") Then
        weightUOM = propertiesToSet("Weight UOM")
    Else
        On Error Resume Next
        Dim massUnit As Long
        massUnit = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
        Select Case massUnit
            Case swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms: weightUOM = "kg"
            Case swUnitsMassPropMass_e.swUnitsMassPropMass_Grams: weightUOM = "g"
            Case swUnitsMassPropMass_e.swUnitsMassPropMass_Pounds: weightUOM = "lb"
            Case Else
                Dim unitSystem As Long
                unitSystem = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                Select Case unitSystem
                    Case swUnitSystem_e.swUnitSystem_IPS: weightUOM = "lb"
                    Case swUnitSystem_e.swUnitSystem_MMGS, swUnitSystem_e.swUnitSystem_CGS: weightUOM = "g"
                    Case swUnitSystem_e.swUnitSystem_MKS: weightUOM = "kg"
                    Case Else: weightUOM = "kg"
                End Select
        End Select
        If Err.Number <> 0 Then
            weightUOM = "kg"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    DebugLog "Weight UOM: " & weightUOM

    ' Get decimal places from the drawing (not the model)
    Dim massDecimalPlaces As Integer
    Dim foundDecimalSetting As Boolean
    foundDecimalSetting = False

    On Error Resume Next
    If modelType = swDocumentTypes_e.swDocDRAWING Then
        massDecimalPlaces = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
        If Err.Number = 0 Then foundDecimalSetting = True
    Else
        massDecimalPlaces = refModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
        If Err.Number = 0 Then foundDecimalSetting = True
    End If
    On Error GoTo 0

    If Not foundDecimalSetting Then
        Select Case LCase(weightUOM)
            Case "kg", "lb": massDecimalPlaces = 3
            Case "g": massDecimalPlaces = 0
            Case "oz": massDecimalPlaces = 1
            Case Else: massDecimalPlaces = 3
        End Select
        DebugLog "Defaulting decimal places: " & massDecimalPlaces
    End If

    ' Override if provided
    If propertiesToSet.exists("Weight Decimal Places") Then
        If IsNumeric(propertiesToSet("Weight Decimal Places")) Then
            massDecimalPlaces = CInt(propertiesToSet("Weight Decimal Places"))
        End If
    End If

    ' Convert and format
    Dim weightValue As Double
    Select Case LCase(weightUOM)
        Case "kg": weightValue = massKg
        Case "g": weightValue = massKg * 1000
        Case "lb": weightValue = massKg * 2.20462
        Case "oz": weightValue = massKg * 35.274
        Case Else: weightValue = massKg: weightUOM = "kg"
    End Select

    Dim formattedWeight As String
    On Error Resume Next
    formattedWeight = FormatNumber(weightValue, massDecimalPlaces)
    If Err.Number <> 0 Then
        formattedWeight = CStr(weightValue)
        Err.Clear
    End If
    On Error GoTo 0

    ' Always set the weight
    propertiesToSet("Weight") = formattedWeight
    DebugLog "Weight property updated: " & formattedWeight & " " & weightUOM

    ' Always set the weight UOM
    propertiesToSet("Weight UOM") = weightUOM
    DebugLog "Weight UOM property updated: " & weightUOM

    DebugLog "---- End of AutofillWeight ----"
    Exit Sub

SetDefaultWeight:
    ' Fallback: Set default weight when calculation fails
    ' This ensures Weight property is always populated
    DebugLog "Setting default weight values due to error"
    propertiesToSet("Weight") = "0"
    propertiesToSet("Weight UOM") = "kg"
    DebugLog "Weight property set to default: 0 kg"
    DebugLog "---- End of AutofillWeight (with defaults) ----"
End Sub
