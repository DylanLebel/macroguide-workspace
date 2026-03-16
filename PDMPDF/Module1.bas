Attribute VB_Name = "Module1"
Option Explicit



Sub UpdateAllModules()
    On Error GoTo EH
    
    Dim vbeObj As Object
    Dim proj As Object
    Dim comp As Object
    Dim folderPath As String
    Dim removedCount As Long
    Dim importedCount As Long
    Dim thisModuleName As String
    Dim i As Long
    
    folderPath = "C:\Users\dlebel\OneDrive - Nordic Minesteel Technologies Inc\Documents\new pdm pdf\march16"
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set vbeObj = VBE
    Set proj = vbeObj.ActiveVBProject
    
    ' Get the name of THIS module to skip it
    thisModuleName = vbeObj.ActiveCodePane.CodeModule.Parent.Name
    
    ' ===== REMOVE OLD MODULES =====
    Debug.Print "===== REMOVING OLD MODULES ====="
    removedCount = 0
    
    For i = proj.VBComponents.count To 1 Step -1
        Set comp = proj.VBComponents(i)
        
        If comp.Name = thisModuleName Then
            Debug.Print "Skipped (self): " & comp.Name
            GoTo NextComp
        End If
        
        Select Case comp.Type
            Case 1, 2, 3 ' Standard, Class, Form
                On Error Resume Next
                proj.VBComponents.Remove comp
                If Err.Number = 0 Then
                    removedCount = removedCount + 1
                    Debug.Print "Removed: " & comp.Name
                Else
                    Debug.Print "Failed to remove: " & comp.Name & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo EH
                
            Case Else
                Debug.Print "Skipped (document): " & comp.Name
        End Select
        
NextComp:
    Next i
    
    Debug.Print "Removed: " & removedCount & " components"
    
    ' ===== IMPORT NEW MODULES =====
    Debug.Print "===== IMPORTING NEW MODULES ====="
    Debug.Print "Import folder: " & folderPath
    importedCount = 0
    
    ImportPattern proj, folderPath, "*.bas", importedCount
    ImportPattern proj, folderPath, "*.cls", importedCount
    ImportPattern proj, folderPath, "*.frm", importedCount
    
    Debug.Print "Imported: " & importedCount & " components"
    
    ' ===== SUMMARY =====
    MsgBox "Update finished." & vbCrLf & _
           "Removed: " & removedCount & vbCrLf & _
           "Imported: " & importedCount & vbCrLf & _
           "From: " & folderPath, vbInformation
    Exit Sub
    
EH:
    MsgBox "Update stopped." & vbCrLf & _
           "Error: " & Err.Number & " " & Err.Description, vbCritical
End Sub


Sub RemoveAllModules()
    On Error GoTo EH
    
    Dim vbeObj As Object
    Dim proj As Object
    Dim comp As Object
    Dim removedCount As Long
    Dim thisModuleName As String
    Dim i As Long
    
    Set vbeObj = VBE
    Set proj = vbeObj.ActiveVBProject
    
    ' Get the name of THIS module (contains both Remove and Import subs)
    thisModuleName = vbeObj.ActiveCodePane.CodeModule.Parent.Name
    
    removedCount = 0
    
    For i = proj.VBComponents.count To 1 Step -1
        Set comp = proj.VBComponents(i)
        
        If comp.Name = thisModuleName Then
            Debug.Print "Skipped (self): " & comp.Name
            GoTo NextComp
        End If
        
        Select Case comp.Type
            Case 1, 2, 3
                On Error Resume Next
                proj.VBComponents.Remove comp
                If Err.Number = 0 Then
                    removedCount = removedCount + 1
                    Debug.Print "Removed: " & comp.Name
                Else
                    Debug.Print "Failed to remove: " & comp.Name & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo EH
                
            Case Else
                Debug.Print "Skipped (document): " & comp.Name
        End Select
        
NextComp:
    Next i
    
    MsgBox "Removal finished." & vbCrLf & "Removed: " & removedCount & " components"
    Exit Sub
    
EH:
    MsgBox "Removal stopped." & vbCrLf & _
           "Error: " & Err.Number & " " & Err.Description
End Sub

Sub ExportAllModules()
    On Error GoTo EH
    
    Dim vbeObj As Object
    Dim proj As Object
    Dim comp As Object
    Dim folderPath As String
    Dim exportedCount As Long
    Dim thisModuleName As String
    Dim fileExt As String
    Dim fullPath As String
    
    folderPath = "C:\pdmpdf"
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set vbeObj = VBE
    Set proj = vbeObj.ActiveVBProject
    
    ' Get the name of THIS module to skip it
    thisModuleName = vbeObj.ActiveCodePane.CodeModule.Parent.Name
    
    exportedCount = 0
    
    Debug.Print "Export folder: " & folderPath
    
    For Each comp In proj.VBComponents
        If comp.Name = thisModuleName Then
            Debug.Print "Skipped (self): " & comp.Name
            GoTo NextComp
        End If
        
        Select Case comp.Type
            Case 1 ' vbext_ct_StdModule
                fileExt = ".bas"
            Case 2 ' vbext_ct_ClassModule
                fileExt = ".cls"
            Case 3 ' vbext_ct_MSForm
                fileExt = ".frm"
            Case Else
                Debug.Print "Skipped (document): " & comp.Name
                GoTo NextComp
        End Select
        
        fullPath = folderPath & comp.Name & fileExt
        
        On Error Resume Next
        comp.Export fullPath
        If Err.Number = 0 Then
            exportedCount = exportedCount + 1
            Debug.Print "Exported: " & fullPath
        Else
            Debug.Print "Failed to export: " & comp.Name & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo EH
        
NextComp:
    Next comp
    
    MsgBox "Export finished." & vbCrLf & "Exported: " & exportedCount & " components" & vbCrLf & "To: " & folderPath
    Exit Sub
    
EH:
    MsgBox "Export stopped." & vbCrLf & _
           "Error: " & Err.Number & " " & Err.Description
End Sub

Sub ImportAllModules()
    On Error GoTo EH
    Dim vbeObj As Object
    Dim proj As Object
    Dim folderPath As String
    Dim fileName As String
    Dim importedCount As Long
    
    folderPath = "C:\Users\dlebel\OneDrive - Nordic Minesteel Technologies Inc\Documents\new pdm pdf\march16"
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set vbeObj = VBE
    Set proj = vbeObj.ActiveVBProject
    importedCount = 0
    
    Debug.Print "Import folder: " & folderPath
    ImportPattern proj, folderPath, "*.bas", importedCount
    ImportPattern proj, folderPath, "*.cls", importedCount
    ImportPattern proj, folderPath, "*.frm", importedCount
    
    MsgBox "Import finished. Imported: " & importedCount
    Exit Sub
    
EH:
    MsgBox "Import stopped." & vbCrLf & _
           "File: " & fileName & vbCrLf & _
           "Error: " & Err.Number & " " & Err.Description
End Sub

Private Sub ImportPattern(ByVal proj As Object, ByVal folderPath As String, ByVal pattern As String, ByRef importedCount As Long)
    Dim fileName As String
    Dim fullPath As String
    
    fileName = Dir(folderPath & pattern)
    Do While fileName <> vbNullString
        fullPath = folderPath & fileName
        On Error GoTo ImportFail
        proj.VBComponents.Import fullPath
        importedCount = importedCount + 1
        Debug.Print "Imported: " & fullPath
NextFile:
        On Error GoTo 0
        fileName = Dir
    Loop
    Exit Sub
    
ImportFail:
    Debug.Print "FAILED:   " & fullPath
    Debug.Print "ERROR:    " & Err.Number & " " & Err.Description
    MsgBox "Failed importing:" & vbCrLf & fullPath & vbCrLf & vbCrLf & "Error: " & Err.Number & " " & Err.Description
    Resume NextFile
End Sub

