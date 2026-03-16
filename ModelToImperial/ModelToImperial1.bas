Attribute VB_Name = "ModelToImperial1"
Dim swApp As Object
Sub main()

Set swApp = Application.SldWorks
End Sub
' ============================================================================
' CONVERT TO IMPERIAL - Complete Standalone Macro
' ============================================================================

Sub ConvertToImperial()
    On Error GoTo ErrorHandler
    
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    Dim activeDoc As ModelDoc2
    Set activeDoc = swApp.activeDoc
    
    If activeDoc Is Nothing Then
        MsgBox "No active document. Please open a model.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize tracking
    Dim processedModels As New Scripting.Dictionary
    Dim modelsToProcess As New Scripting.Dictionary
    
    ' Add starting model
    Dim startPath As String
    startPath = activeDoc.GetPathName()
    
    If startPath = "" Then
        MsgBox "Please save the model before converting units.", vbExclamation
        Exit Sub
    End If
    
    modelsToProcess.Add startPath, activeDoc
    
    ' Process all models
    Dim modelCount As Long
    modelCount = 0
    
    Do While modelsToProcess.Count > 0
        Dim currentModel As ModelDoc2
        Dim modelKey As Variant
        modelKey = modelsToProcess.keys()(0)
        Set currentModel = modelsToProcess.Items()(0)
        Dim modelPath As String
        modelPath = CStr(modelKey)
        
        If Not currentModel Is Nothing And Not processedModels.exists(modelPath) Then
            
            ' Add subcomponents first if it's an assembly
            If currentModel.GetType = swDocASSEMBLY Then
                Call AddSubcomponentsImperial(currentModel, modelsToProcess, processedModels)
            End If
            
            ' Convert this model to Imperial
            If currentModel.GetType <> swDocDRAWING Then
                Call SetModelToImperial(currentModel)
                modelCount = modelCount + 1
            End If
            
            processedModels.Add modelPath, True
        End If
        
        modelsToProcess.Remove modelPath
        Set currentModel = Nothing
    Loop
    
    MsgBox "Conversion complete! " & modelCount & " model(s) converted to Imperial (inches).", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub SetModelToImperial(model As ModelDoc2)
    On Error Resume Next
    
    Dim swModelExt As ModelDocExtension
    Set swModelExt = model.Extension
    
    ' Set to IPS system (Inch, Pound, Second)
    swModelExt.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitSystem, 0, 0, 4
    
    ' Set linear units to inches
    swModelExt.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitsLinear, 0, 0, 3
    
    ' Rebuild and save
    model.ForceRebuild3 True
    model.Save3 swSaveAsOptions_Silent, 0, 0
End Sub

Sub AddSubcomponentsImperial(model As ModelDoc2, _
                             ByRef modelsToProcess As Scripting.Dictionary, _
                             ByRef processedModels As Scripting.Dictionary)
    
    If model.GetType <> swDocASSEMBLY Then Exit Sub
    
    Dim swAssy As AssemblyDoc
    Set swAssy = model
    
    On Error Resume Next
    Dim components As Variant
    components = swAssy.GetComponents(True)
    
    If Err.Number <> 0 Or IsEmpty(components) Then Exit Sub
    
    Dim component As Variant
    For Each component In components
        If Not component Is Nothing Then
            Dim compModel As ModelDoc2
            Set compModel = component.GetModelDoc2
            
            If Not compModel Is Nothing Then
                Dim compPath As String
                compPath = compModel.GetPathName()
                
                If compPath <> "" And Not processedModels.exists(compPath) And Not modelsToProcess.exists(compPath) Then
                    modelsToProcess.Add compPath, compModel
                End If
            End If
        End If
    Next
End Sub


