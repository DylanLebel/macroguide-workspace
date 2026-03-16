Attribute VB_Name = "Units"
Option Explicit
Option Private Module

Sub GetDocumentUnits()
    ' OPTIMIZATION: Use global swApp accessor
    Dim swAppLocal As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelExt As SldWorks.ModelDocExtension

    ' Get the SolidWorks application and active document
    Set swAppLocal = GetSwApp()
    Set swModel = swAppLocal.activeDoc
    If swModel Is Nothing Then
        MsgBox "No document is open."
        Exit Sub
    End If
    Set swModelExt = swModel.Extension
    
    ' Unit System
    DebugLog "Unit System: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    
    ' Basic Units - Length
    DebugLog "Length Unit: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinear, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    DebugLog "Length Decimal Display: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalDisplay, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    DebugLog "Length Decimal Places: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    
    '============================================================================
    ' MODIFIED: Currently forcing 16ths instead of getting from model
    ' TO REVERT THIS CHANGE: Replace the next two lines with the commented lines below
    FRACTION_DENOMINATOR = 16  ' Force 16ths for all fractions
    DebugLog "Length Fraction Denominator forced to: " & FRACTION_DENOMINATOR
    
    ' ORIGINAL CODE (uncomment these lines to revert):
    ' FRACTION_DENOMINATOR = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearFractionDenominator, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    ' DebugLog "Length Fraction Denominator: " & FRACTION_DENOMINATOR
    '============================================================================
    
    ' Basic Units - Angle
    DebugLog "Angle Unit: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsAngular, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    DebugLog "Angle Decimal Places: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsAngularDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    
    ' Mass/Section Properties
    DebugLog "Mass Prop Mass Unit: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    DebugLog "Mass Prop Length Unit: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropLength, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    DebugLog "Mass Prop Volume Unit: " & swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropVolume, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    
    ' Optional: Derive Mass Unit from Unit System
    Dim unitSystem As Long
    unitSystem = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
    
    Dim massUnit As String
    Select Case unitSystem
        Case swUnitSystem_e.swUnitSystem_CGS
            massUnit = "Grams"
        Case swUnitSystem_e.swUnitSystem_IPS
            massUnit = "Pounds"
        Case swUnitSystem_e.swUnitSystem_MKS
            massUnit = "Kilograms"
        Case swUnitSystem_e.swUnitSystem_MMGS
            massUnit = "Grams"
        Case Else
            massUnit = "Custom"
    End Select
    
    DebugLog "Derived Mass Unit: " & massUnit
    
    ' TRY TO GET ADDITIONAL UNIT INFO - Skip if not supported
    On Error Resume Next
    Dim lengthUnit As UserUnit
    Set lengthUnit = swModel.GetUserUnit(swUserUnitsType_e.swLengthUnit)
    
    If Not lengthUnit Is Nothing Then
        DebugLog "=== ADDITIONAL UNIT DEBUG INFO ==="
        DebugLog "Length Unit ConversionFactor: " & lengthUnit.conversionFactor
        DebugLog "Length Unit SpecificUnitType: " & lengthUnit.SpecificUnitType
        DebugLog "==================================="
        
        ' Also get decimal places from the unit object for consistency
        DECIMAL_PLACES = lengthUnit.Decimals
        DebugLog "DECIMAL_PLACES set to: " & DECIMAL_PLACES
    Else
        DebugLog "=== ADDITIONAL UNIT DEBUG INFO ==="
        DebugLog "GetUserUnit method not available in this SolidWorks version"
        DebugLog "==================================="
        
        ' Fallback: Get decimal places from user preferences instead
        DECIMAL_PLACES = swModelExt.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
        DebugLog "DECIMAL_PLACES set from preferences: " & DECIMAL_PLACES
    End If
    
    ' Resume normal error handling
    On Error GoTo 0
End Sub
