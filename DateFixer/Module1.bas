Attribute VB_Name = "Module1"
' =============================================================================
' RevDateToProperty Macro
' Extracts the date from the first revision in the revision table and
' replaces the "Date Created" custom property with that static date value
' =============================================================================

Option Explicit

' Log file handle
Private logFile As Integer
Private logPath As String

' Counters for summary
Private cntProcessed As Long
Private cntSkipped_NoRevTable As Long
Private cntSkipped_ReadOnly As Long
Private cntErrors As Long

' =============================================================================
' BATCH PROCESSING - Main entry point for folder processing
' =============================================================================
Sub BatchProcessFolder()
    Dim swApp As SldWorks.SldWorks
    Dim folderPath As String
    Dim drawingFiles As Collection
    
    Set swApp = Application.SldWorks
    
    ' Browse for folder
    folderPath = BrowseForFolder("Select folder containing drawings to process")
    
    If folderPath = "" Then
        MsgBox "No folder selected. Operation cancelled.", vbInformation, "Batch Process"
        Exit Sub
    End If
    
    ' Initialize counters
    cntProcessed = 0
    cntSkipped_NoRevTable = 0
    cntSkipped_ReadOnly = 0
    cntErrors = 0
    
    ' Initialize log file
    logPath = folderPath & "\RevDateToProperty_Log_" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & ".txt"
    logFile = FreeFile
    Open logPath For Output As #logFile
    
    WriteLog "============================================================"
    WriteLog "RevDateToProperty Batch Processing Log"
    WriteLog "Started: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    WriteLog "Folder: " & folderPath
    WriteLog "============================================================"
    WriteLog ""
    
    ' Find all checked-out drawing files (recursive)
    Set drawingFiles = New Collection
    FindCheckedOutDrawings folderPath, drawingFiles
    
    WriteLog "Found " & drawingFiles.Count & " checked-out drawing files"
    WriteLog ""
    
    If drawingFiles.Count = 0 Then
        WriteLog "No checked-out drawings found. Nothing to process."
        MsgBox "No checked-out drawings found in the selected folder.", vbInformation, "Batch Process"
        GoTo Cleanup
    End If
    
    ' Confirm with user
    If MsgBox("Found " & drawingFiles.Count & " checked-out drawings." & vbCrLf & vbCrLf & _
              "Process all and update Date Created property from revision tables?", _
              vbYesNo + vbQuestion, "Batch Process") = vbNo Then
        WriteLog "User cancelled operation."
        GoTo Cleanup
    End If
    
    ' Process each drawing
    Dim filePath As Variant
    Dim fileNum As Long
    fileNum = 0
    
    For Each filePath In drawingFiles
        fileNum = fileNum + 1
        Debug.Print "Processing " & fileNum & " of " & drawingFiles.Count & ": " & filePath
        ProcessSingleDrawing swApp, CStr(filePath)
    Next filePath
    
    ' Write summary
    WriteLog ""
    WriteLog "============================================================"
    WriteLog "SUMMARY"
    WriteLog "============================================================"
    WriteLog "Processed successfully: " & cntProcessed
    WriteLog "Skipped (no rev table): " & cntSkipped_NoRevTable
    WriteLog "Skipped (read-only):    " & cntSkipped_ReadOnly
    WriteLog "Errors:                 " & cntErrors
    WriteLog "Total files checked:    " & drawingFiles.Count
    WriteLog ""
    WriteLog "Completed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    WriteLog "============================================================"
    
    MsgBox "Batch processing complete!" & vbCrLf & vbCrLf & _
           "Processed: " & cntProcessed & vbCrLf & _
           "Skipped (no rev table): " & cntSkipped_NoRevTable & vbCrLf & _
           "Errors: " & cntErrors & vbCrLf & vbCrLf & _
           "Log saved to:" & vbCrLf & logPath, _
           vbInformation, "Batch Process Complete"

Cleanup:
    Close #logFile
End Sub

' Process a single drawing file
Sub ProcessSingleDrawing(swApp As SldWorks.SldWorks, filePath As String)
    Dim swModel As SldWorks.ModelDoc2
    Dim errors As Long
    Dim warnings As Long
    Dim revDate As String
    Dim wasAlreadyOpen As Boolean
    Dim fileName As String
    
    fileName = GetFileNameFromPath(filePath)
    WriteLog "----------------------------------------"
    WriteLog "File: " & fileName
    WriteLog "Path: " & filePath
    
    On Error GoTo ErrorHandler
    
    ' Check if file is already open
    Set swModel = swApp.GetOpenDocument(filePath)
    wasAlreadyOpen = Not (swModel Is Nothing)
    
    ' Open the file if not already open
    If swModel Is Nothing Then
        Set swModel = swApp.OpenDoc6(filePath, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
    End If
    
    If swModel Is Nothing Then
        WriteLog "  ERROR: Could not open file (Error: " & errors & ", Warning: " & warnings & ")"
        cntErrors = cntErrors + 1
        Exit Sub
    End If
    
    ' Process the drawing
    If UpdateDateCreatedFromRevTable(swApp, swModel) Then
        WriteLog "  SUCCESS: Updated Date Created property"
        
        ' Save the file
        If swModel.Save3(swSaveAsOptions_Silent, errors, warnings) Then
            WriteLog "  Saved successfully"
        Else
            WriteLog "  WARNING: Save returned false (Error: " & errors & ", Warning: " & warnings & ")"
        End If
        
        cntProcessed = cntProcessed + 1
    Else
        WriteLog "  SKIPPED: No revision table found"
        cntSkipped_NoRevTable = cntSkipped_NoRevTable + 1
    End If
    
    ' Close the file if we opened it
    If Not wasAlreadyOpen Then
        swApp.CloseDoc swModel.GetPathName
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog "  ERROR: " & Err.Description
    cntErrors = cntErrors + 1
    On Error Resume Next
    If Not wasAlreadyOpen And Not swModel Is Nothing Then
        swApp.CloseDoc swModel.GetPathName
    End If
End Sub

' Recursively find all checked-out (writable) drawing files
Sub FindCheckedOutDrawings(folderPath As String, ByRef fileCollection As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Check all files in this folder
    For Each file In folder.Files
        If UCase(fso.GetExtensionName(file.Name)) = "SLDDRW" Then
            ' Check if file is writable (checked out)
            If Not IsFileReadOnly(file.Path) Then
                fileCollection.Add file.Path
                Debug.Print "Found checked-out: " & file.Path
            Else
                Debug.Print "Skipping read-only: " & file.Path
            End If
        End If
    Next file
    
    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        FindCheckedOutDrawings subFolder.Path, fileCollection
    Next subFolder
End Sub

' Check if a file is read-only
Function IsFileReadOnly(filePath As String) As Boolean
    On Error Resume Next
    IsFileReadOnly = (GetAttr(filePath) And vbReadOnly) = vbReadOnly
    If Err.Number <> 0 Then IsFileReadOnly = True
    On Error GoTo 0
End Function

' Browse for folder dialog
Function BrowseForFolder(prompt As String) As String
    Dim shell As Object
    Dim folder As Object
    
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, prompt, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

' Extract filename from full path
Function GetFileNameFromPath(filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileNameFromPath = Mid(filePath, pos + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

' Write to log file and debug window
Sub WriteLog(message As String)
    Print #logFile, message
    Debug.Print message
End Sub

' =============================================================================
' SINGLE FILE PROCESSING - Process active drawing
' =============================================================================
Sub ProcessActiveDrawing()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim result As Boolean
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "No document open.", vbExclamation, "Rev Date to Property"
        Exit Sub
    End If
    
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "Active document is not a drawing.", vbExclamation, "Rev Date to Property"
        Exit Sub
    End If
    
    result = UpdateDateCreatedFromRevTable(swApp, swModel)
    
    If result Then
        MsgBox "Date Created property updated successfully.", vbInformation, "Rev Date to Property"
    Else
        MsgBox "Could not update Date Created property. Check if revision table exists.", vbExclamation, "Rev Date to Property"
    End If
End Sub

' Core function to update the Date Created property from revision table
' Returns True if successful, False if no rev table or error
Function UpdateDateCreatedFromRevTable(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2) As Boolean
    Dim swDraw As SldWorks.DrawingDoc
    Dim swSheet As SldWorks.Sheet
    Dim swView As SldWorks.View
    Dim vSheetNames As Variant
    Dim revDate As String
    Dim i As Integer
    
    UpdateDateCreatedFromRevTable = False
    
    Set swDraw = swModel
    
    ' Get the first revision date from the revision table
    revDate = GetFirstRevisionDate(swDraw)
    
    If revDate = "" Then
        Debug.Print "No revision table found or no revisions in table: " & swModel.GetPathName
        Exit Function
    End If
    
    Debug.Print "Found first revision date: " & revDate
    
    ' Update the Date Created custom property
    If SetDateCreatedProperty(swModel, revDate) Then
        UpdateDateCreatedFromRevTable = True
    End If
End Function

' Finds the revision table and extracts the date from the first revision row
Function GetFirstRevisionDate(swDraw As SldWorks.DrawingDoc) As String
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.Feature
    Dim swSubFeat As SldWorks.Feature
    Dim swRevTableFeat As SldWorks.RevisionTableFeature
    Dim swRevTableAnn As SldWorks.RevisionTableAnnotation
    Dim swTableAnn As SldWorks.TableAnnotation
    Dim vRevTableAnns As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim dateColIndex As Long
    Dim firstDataRow As Long
    Dim cellText As String
    Dim i As Long
    
    GetFirstRevisionDate = ""
    
    Set swModel = swDraw
    
    ' Iterate through features to find the revision table
    Set swFeat = swModel.FirstFeature
    
    Do While Not swFeat Is Nothing
        Debug.Print "Feature: " & swFeat.Name & " [" & swFeat.GetTypeName2 & "]"
        
        ' Check if this feature is the revision table
        If swFeat.GetTypeName2 = "RevisionTableFeat" Then
            Debug.Print "  Found RevisionTableFeat at top level"
            cellText = ExtractDateFromRevTableFeature(swFeat)
            If cellText <> "" Then
                GetFirstRevisionDate = cellText
                Exit Function
            End If
        End If
        
        ' Check sub-features (revision table is often under a sheet)
        Set swSubFeat = swFeat.GetFirstSubFeature
        Do While Not swSubFeat Is Nothing
            Debug.Print "  Sub-Feature: " & swSubFeat.Name & " [" & swSubFeat.GetTypeName2 & "]"
            
            If swSubFeat.GetTypeName2 = "RevisionTableFeat" Then
                Debug.Print "    Found RevisionTableFeat as sub-feature"
                cellText = ExtractDateFromRevTableFeature(swSubFeat)
                If cellText <> "" Then
                    GetFirstRevisionDate = cellText
                    Exit Function
                End If
            End If
            
            Set swSubFeat = swSubFeat.GetNextSubFeature
        Loop
        
        Set swFeat = swFeat.GetNextFeature
    Loop
    
    Debug.Print "No revision table feature found in drawing"
End Function

' Extracts the first revision date from a RevisionTableFeat feature
Function ExtractDateFromRevTableFeature(swFeat As SldWorks.Feature) As String
    Dim swRevTableFeat As SldWorks.RevisionTableFeature
    Dim swRevTableAnn As SldWorks.RevisionTableAnnotation
    Dim swTableAnn As SldWorks.TableAnnotation
    Dim vRevTableAnns As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim dateColIndex As Long
    Dim firstDataRow As Long
    Dim cellText As String
    Dim i As Long
    
    ExtractDateFromRevTableFeature = ""
    
    Set swRevTableFeat = swFeat.GetSpecificFeature2
    
    If swRevTableFeat Is Nothing Then
        Debug.Print "    GetSpecificFeature2 returned Nothing"
        Exit Function
    End If
    
    ' Get the revision table annotations
    vRevTableAnns = swRevTableFeat.GetTableAnnotations
    
    If IsEmpty(vRevTableAnns) Then
        Debug.Print "    GetTableAnnotations returned empty"
        Exit Function
    End If
    
    For i = 0 To UBound(vRevTableAnns)
        Set swRevTableAnn = vRevTableAnns(i)
        Set swTableAnn = swRevTableAnn
        
        rowCount = swTableAnn.rowCount
        colCount = swTableAnn.ColumnCount
        
        Debug.Print "    Revision table - Rows: " & rowCount & ", Columns: " & colCount
        
        ' Date is in the last column
        dateColIndex = colCount - 1
        
        ' Find the first revision row (lowest rev number)
        firstDataRow = FindFirstRevisionRow(swTableAnn)
        
        If firstDataRow >= 0 Then
            cellText = swTableAnn.Text(firstDataRow, dateColIndex)
            Debug.Print "    First revision date: " & cellText
            ExtractDateFromRevTableFeature = cellText
            Exit Function
        End If
    Next i
End Function

' Finds the row containing the first revision (Rev 1 or lowest rev number)
Function FindFirstRevisionRow(swTableAnn As SldWorks.TableAnnotation) As Long
    Dim rowCount As Long
    Dim i As Long
    Dim revText As String
    Dim revNum As Long
    Dim lowestRev As Long
    Dim lowestRevRow As Long
    
    FindFirstRevisionRow = -1
    lowestRev = 999999
    lowestRevRow = -1
    
    rowCount = swTableAnn.rowCount
    
    ' Column 0 is the revision number column (REV.#)
    ' Skip row 0 (header row)
    For i = 1 To rowCount - 1
        revText = Trim(swTableAnn.Text(i, 0))
        
        ' Try to parse as number, handle both numeric and letter revisions
        If IsNumeric(revText) Then
            revNum = CLng(revText)
            If revNum < lowestRev Then
                lowestRev = revNum
                lowestRevRow = i
            End If
        ElseIf Len(revText) = 1 Then
            ' Letter revision (A=1, B=2, etc.)
            revNum = Asc(UCase(revText)) - Asc("A") + 1
            If revNum < lowestRev Then
                lowestRev = revNum
                lowestRevRow = i
            End If
        End If
    Next i
    
    FindFirstRevisionRow = lowestRevRow
    
    If lowestRevRow >= 0 Then
        Debug.Print "  Found first revision at row " & lowestRevRow & " (Rev " & swTableAnn.Text(lowestRevRow, 0) & ")"
    End If
End Function

' Sets the Date Created custom property to the specified date string
Function SetDateCreatedProperty(swModel As SldWorks.ModelDoc2, dateValue As String) As Boolean
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Dim currentVal As String
    Dim currentResolvedVal As String
    Dim wasResolved As Boolean
    Dim retVal As Long
    
    SetDateCreatedProperty = False
    
    ' Get the custom property manager for the document (empty string = document level)
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")
    
    ' Check if Date Created property exists
    retVal = swCustPropMgr.Get5("Date Created", False, currentVal, currentResolvedVal, wasResolved)
    
    If retVal = swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent Then
        ' Property doesn't exist, add it
        Debug.Print "Date Created property not found, adding it..."
        retVal = swCustPropMgr.Add3("Date Created", swCustomInfoType_e.swCustomInfoText, dateValue, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    Else
        ' Property exists, update it
        Debug.Print "Current Date Created value: " & currentVal & " (Resolved: " & currentResolvedVal & ")"
        Debug.Print "Setting Date Created to: " & dateValue
        retVal = swCustPropMgr.Set2("Date Created", dateValue)
    End If
    
    If retVal = swCustomInfoSetResult_e.swCustomInfoSetResult_OK Or retVal = 0 Then
        SetDateCreatedProperty = True
        Debug.Print "Date Created property updated successfully"
    Else
        Debug.Print "Failed to set Date Created property. Return value: " & retVal
    End If
End Function
