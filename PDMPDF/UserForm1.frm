VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' API declarations for 32-bit and 64-bit compatibility to enable resizing
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
#Else
    Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
#End If

' Variables to store original form dimensions (optional for future scaling features)
Private FormOriginalWidth As Double
Private FormOriginalHeight As Double
Private ControlsInitialized As Boolean

' Store results for double-click functionality
Private mSuccessList As Collection
Private mErrorList As Collection
Private mWarningList As Collection
Private mTotalFiles As Long

' Constants for easy fine-tuning of layout
Private Const FORM_MARGIN As Integer = 10    ' Margin between form edges and controls
Private Const BUTTON_MARGIN As Integer = 10  ' Space between MultiPage and button
Private Const LISTBOX_MARGIN As Integer = 10 ' Margin inside MultiPage for listboxes

Private Sub UserForm_Initialize()
    ' PURPOSE: Sets the initial size and position of all controls
    ' Fine-tune these values to adjust the starting size of the form
    Me.Caption = "PDF Batch Results"
    Me.Width = 600   ' Initial width of the form in points
    Me.Height = 400  ' Initial height of the form in points

    ' Initialize collections
    Set mSuccessList = New Collection
    Set mErrorList = New Collection
    Set mWarningList = New Collection

    ' Position the MultiPage with margins from the form edges
    With MultiPage1
        .Left = FORM_MARGIN
        .Top = FORM_MARGIN
        .Width = Me.InsideWidth - 2 * FORM_MARGIN
        ' Height leaves space for the button below
        .Height = Me.InsideHeight - CommandButton1.Height - BUTTON_MARGIN - FORM_MARGIN
    End With

    ' Position the CommandButton centered at the bottom
    With CommandButton1
        .Left = (Me.InsideWidth - .Width) / 2  ' Centers the button horizontally
        .Top = MultiPage1.Top + MultiPage1.Height + BUTTON_MARGIN  ' Places it below MultiPage
    End With

    ' Configure ListBox1 on Page 1 to fit within its page
    With MultiPage1.Pages(0).ListBox1
        .Left = LISTBOX_MARGIN    ' Margin from left edge of the page
        .Top = LISTBOX_MARGIN     ' Margin from top of the page (below tabs)
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN   ' Fits within MultiPage width
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN ' Fits within MultiPage height
        ' Styling
        .Font.Name = "Arial"
        .Font.Size = 10
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .BorderStyle = fmBorderStyleSingle
        ' Set initial column widths
        .ColumnCount = 2
        .ColumnWidths = "100;200"
    End With

    ' Configure ListBox2 on Page 2 with identical settings
    With MultiPage1.Pages(1).ListBox2
        .Left = LISTBOX_MARGIN
        .Top = LISTBOX_MARGIN
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN
        .Font.Name = "Arial"
        .Font.Size = 10
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .BorderStyle = fmBorderStyleSingle
        ' Set initial column widths
        .ColumnCount = 2
        .ColumnWidths = "150;300"
    End With
    
    ' Configure ListBox3 on Page 3 with identical settings
    With MultiPage1.Pages(2).ListBox3
        .Left = LISTBOX_MARGIN
        .Top = LISTBOX_MARGIN
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN
        .Font.Name = "Arial"
        .Font.Size = 10
        .BackColor = vbWhite
        .ForeColor = vbBlack
        .BorderStyle = fmBorderStyleSingle
        ' Set initial column widths
        .ColumnCount = 2
        .ColumnWidths = "150;300"
    End With
End Sub

Private Sub CommandButton1_Click()
    ' PURPOSE: Closes the UserForm when the button is clicked
    Unload Me
End Sub

Private Sub UserForm_Activate()
    ' PURPOSE: Makes the form resizable when it becomes active
    MakeFormResizable
End Sub

Private Sub UserForm_Resize()
    ' PURPOSE: Adjusts control sizes and positions when the form is resized
    ' Ensures MultiPage fits within the form and leaves space for the button
    With MultiPage1
        .Width = Me.InsideWidth - 2 * FORM_MARGIN
        .Height = Me.InsideHeight - CommandButton1.Height - BUTTON_MARGIN - FORM_MARGIN
    End With

    ' Reposition the CommandButton to stay centered at the bottom
    With CommandButton1
        .Left = (Me.InsideWidth - .Width) / 2
        .Top = MultiPage1.Top + MultiPage1.Height + BUTTON_MARGIN
    End With

    ' Resize ListBox1 to fit within its MultiPage page and adjust column widths
    With MultiPage1.Pages(0).ListBox1
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN
        .ColumnWidths = "100;" & (.Width - 100 - 20)  ' 20 points reserved for scrollbar
    End With

    ' Resize ListBox2 similarly
    With MultiPage1.Pages(1).ListBox2
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN
        .ColumnWidths = "150;" & (.Width - 150 - 20)  ' 20 points reserved for scrollbar
    End With
    
    ' Resize ListBox3 similarly
    With MultiPage1.Pages(2).ListBox3
        .Width = MultiPage1.Width - 2 * LISTBOX_MARGIN
        .Height = MultiPage1.Height - 2 * LISTBOX_MARGIN
        .ColumnWidths = "150;" & (.Width - 150 - 20)  ' 20 points reserved for scrollbar
    End With
End Sub

' Subroutine to make the UserForm resizable
Private Sub MakeFormResizable()
    #If VBA7 Then
        Dim hWnd As LongPtr  ' Corrected for 64-bit compatibility
    #Else
        Dim hWnd As Long
    #End If
    Dim lStyle As Long
    Dim retVal As Long
    Const WS_THICKFRAME = &H40000  ' Style flag for a resizable border
    Const GWL_STYLE = -16          ' Index to access window style

    ' Get the handle of the active window (the UserForm)
    hWnd = GetActiveWindow()

    ' Get the current window style
    lStyle = GetWindowLong(hWnd, GWL_STYLE)

    ' Add the resizable frame style
    lStyle = lStyle Or WS_THICKFRAME

    ' Set the new window style
    retVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

    ' Clear any previous API error codes
    SetLastError 0

    ' Check if the style was set successfully
    If retVal = 0 Then
        MsgBox "Unable to make UserForm resizable."
    End If
End Sub

Public Sub PopulateResults(successList As Collection, errorList As Collection, warningList As Collection, totalFiles As Long)
    ' PURPOSE: Populates the listboxes with processing results

    ' Store for double-click functionality
    Set mSuccessList = successList
    Set mErrorList = errorList
    Set mWarningList = warningList
    mTotalFiles = totalFiles

    ' Improved caption with summary
    Me.Caption = "PDF Results: " & successList.count & " OK | " & _
                 errorList.count & " Errors | " & _
                 warningList.count & " Warnings  (" & totalFiles & " total)"

    ' Populate Success page
    MultiPage1.Pages(0).Caption = "Success (" & successList.count & ")"
    With MultiPage1.Pages(0).ListBox1
        .Clear
        .AddItem
        .List(0, 0) = "Created PDFs"
        .ColumnCount = 1
        .ColumnWidths = "350"  ' Make the column wider since it's the only one
    End With
    
    Dim i As Integer
    Dim fileName As String
    Dim successItem As Variant
    i = 1
    For Each successItem In successList
        fileName = Mid(successItem, InStrRev(successItem, "\") + 1)
        MultiPage1.Pages(0).ListBox1.AddItem
        MultiPage1.Pages(0).ListBox1.List(i, 0) = fileName
        i = i + 1
    Next successItem
    
    ' Populate Errors page
    MultiPage1.Pages(1).Caption = "Errors (" & errorList.count & ")"
    With MultiPage1.Pages(1).ListBox2
        .Clear
        .AddItem
        .List(0, 0) = "File Name"
        .List(0, 1) = "Error Description"
    End With
    
    i = 1
    Dim errorItem As Variant
    Dim errorMsg As String
    
    For Each errorItem In errorList
        errorMsg = CStr(errorItem)  ' Convert to string
        
        ' Default values
        Dim filePath As String
        Dim errorDescription As String
        Dim quoteStart As Long, quoteEnd As Long
        
        fileName = "Unknown File"
        errorDescription = errorMsg
        
        ' Special case for "Unable to process model for drawing:" pattern
        If InStr(errorMsg, "Unable to process model for drawing:") > 0 Then
            ' Extract the file from the end of this message
            filePath = Trim(Mid(errorMsg, InStr(errorMsg, ":") + 1))
            If InStrRev(filePath, "\") > 0 Then
                fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
            Else
                fileName = filePath
            End If
            
            ' Look for the previous error in the collection that might explain the reason
            Dim previousError As String
            Dim previousIndex As Integer
            previousIndex = i - 1
            
            If previousIndex > 0 And previousIndex <= errorList.count Then
                previousError = CStr(errorList(previousIndex))
                
                ' If the previous error has "Error:" in it and contains the same file name, use it
                If InStr(previousError, "Error:") > 0 And InStr(previousError, fileName) > 0 Then
                    ' Extract the reason from the previous error
                    Dim reasonStart As Integer
                    reasonStart = InStr(previousError, "Error:") + 6
                    errorDescription = "Unable to process model: " & Trim(Mid(previousError, reasonStart))
                Else
                    ' Default reason
                    errorDescription = "Unable to process model for drawing"
                End If
            Else
                ' Default reason
                errorDescription = "Unable to process model for drawing"
            End If
        
        ' Special case for "Drawing has no revision" pattern
        ElseIf InStr(errorMsg, "has no revision specified") > 0 Then
            ' Extract the file from between quotes
            quoteStart = InStr(errorMsg, "'")
            quoteEnd = InStr(quoteStart + 1, errorMsg, "'")
            
            If quoteStart > 0 And quoteEnd > quoteStart Then
                filePath = Mid(errorMsg, quoteStart + 1, quoteEnd - quoteStart - 1)
                If InStrRev(filePath, "\") > 0 Then
                    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                Else
                    fileName = filePath
                End If
                errorDescription = "No revision specified"
            End If
        
        ' First check for text between single quotes (most common pattern)
        ElseIf InStr(errorMsg, "'") > 0 Then
            quoteStart = InStr(errorMsg, "'")
            quoteEnd = InStr(quoteStart + 1, errorMsg, "'")
            If quoteEnd > quoteStart Then
                filePath = Mid(errorMsg, quoteStart + 1, quoteEnd - quoteStart - 1)
                
                ' Extract just the filename from the path
                If InStrRev(filePath, "\") > 0 Then
                    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                Else
                    fileName = filePath
                End If
                
                ' Create cleaner error description by removing specific path
                errorDescription = Replace(errorMsg, "'" & filePath & "'", "")
            End If
        
        ' File extension scan (last method)
        Else
            Dim j As Long
            Dim startPos As Long, endPos As Long
            
            ' Look for SolidWorks file extensions in the message
            For j = 1 To Len(errorMsg) - 6
                If Mid(errorMsg, j, 7) = ".slddrw" Or _
                   Mid(errorMsg, j, 7) = ".sldprt" Or _
                   Mid(errorMsg, j, 7) = ".sldasm" Or _
                   (j <= Len(errorMsg) - 3 And Mid(errorMsg, j, 4) = ".pdf") Then
                    
                    ' Found an extension, now find the complete path
                    endPos = j + IIf(Mid(errorMsg, j, 4) = ".pdf", 3, 6)
                    
                    ' Look backward for start of path (find last space or quote before extension)
                    startPos = j
                    Do While startPos > 1
                        If Mid(errorMsg, startPos - 1, 1) = " " Or _
                           Mid(errorMsg, startPos - 1, 1) = "'" Or _
                           Mid(errorMsg, startPos - 1, 1) = """" Then
                            Exit Do
                        End If
                        startPos = startPos - 1
                    Loop
                    
                    ' Extract file path and name
                    filePath = Mid(errorMsg, startPos, endPos - startPos + 1)
                    
                    If InStr(filePath, "\") > 0 Then
                        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                    Else
                        fileName = filePath
                    End If
                    
                    errorDescription = errorMsg
                    Exit For
                End If
            Next j
        End If
        
        ' Add to ListBox2
        MultiPage1.Pages(1).ListBox2.AddItem
        MultiPage1.Pages(1).ListBox2.List(i, 0) = fileName
        MultiPage1.Pages(1).ListBox2.List(i, 1) = errorDescription
        i = i + 1
    Next errorItem
    
    ' Populate Warnings page
    MultiPage1.Pages(2).Caption = "Warnings (" & warningList.count & ")"
    With MultiPage1.Pages(2).ListBox3
        .Clear
        .AddItem
        .List(0, 0) = "File Name"
        .List(0, 1) = "Warning Description"
    End With
    
    i = 1
    Dim warningItem As Variant
    Dim warningMsg As String
    
    For Each warningItem In warningList
        warningMsg = CStr(warningItem)  ' Convert to string
        
        ' Default values
        fileName = "Unknown File"
        errorDescription = warningMsg
        
        ' Parse warning similar to errors
        ' For checked-in drawings with unsaved changes
        If InStr(warningMsg, "is checked in but has unsaved changes") > 0 Then
            quoteStart = InStr(warningMsg, "'")
            quoteEnd = InStr(quoteStart + 1, warningMsg, "'")
            
            If quoteStart > 0 And quoteEnd > quoteStart Then
                filePath = Mid(warningMsg, quoteStart + 1, quoteEnd - quoteStart - 1)
                If InStrRev(filePath, "\") > 0 Then
                    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                Else
                    fileName = filePath
                End If
                errorDescription = "Drawing is checked in but has unsaved changes. Creating PDF from current state."
            End If
        
        ' For skipped sync due to filename mismatch
        ElseIf InStr(warningMsg, "Skipped sync: Drawing") > 0 Then
            quoteStart = InStr(warningMsg, "'")
            quoteEnd = InStr(quoteStart + 1, warningMsg, "'")
            
            If quoteStart > 0 And quoteEnd > quoteStart Then
                filePath = Mid(warningMsg, quoteStart + 1, quoteEnd - quoteStart - 1)
                If InStrRev(filePath, "\") > 0 Then
                    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                Else
                    fileName = filePath
                End If
                errorDescription = "Skipped revision sync: Drawing references model with different filename"
            End If
        
        ' Generic text between quotes
        ElseIf InStr(warningMsg, "'") > 0 Then
            quoteStart = InStr(warningMsg, "'")
            quoteEnd = InStr(quoteStart + 1, warningMsg, "'")
            If quoteEnd > quoteStart Then
                filePath = Mid(warningMsg, quoteStart + 1, quoteEnd - quoteStart - 1)
                
                ' Extract just the filename from the path
                If InStrRev(filePath, "\") > 0 Then
                    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
                Else
                    fileName = filePath
                End If
                
                ' Create cleaner warning description by removing specific path
                errorDescription = Replace(warningMsg, "'" & filePath & "'", "")
            End If
        End If
        
        ' Add to ListBox3
        MultiPage1.Pages(2).ListBox3.AddItem
        MultiPage1.Pages(2).ListBox3.List(i, 0) = fileName
        MultiPage1.Pages(2).ListBox3.List(i, 1) = errorDescription
        i = i + 1
    Next warningItem
    
    ' Show the appropriate page
    If errorList.count > 0 Then
        MultiPage1.Value = 1  ' Show Errors page if there are errors
    ElseIf warningList.count > 0 Then
        MultiPage1.Value = 2  ' Show Warnings page if there are warnings but no errors
    Else
        MultiPage1.Value = 0  ' Show Success page if no errors or warnings
    End If
End Sub

' Double-click a success item to open the PDF in Explorer
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim idx As Long
    Dim selectedPath As String

    idx = MultiPage1.Pages(0).ListBox1.ListIndex
    If idx < 1 Then Exit Sub  ' Skip header row (index 0)
    If mSuccessList Is Nothing Then Exit Sub
    If idx > mSuccessList.count Then Exit Sub

    On Error Resume Next
    selectedPath = mSuccessList(idx)
    On Error GoTo 0

    If selectedPath <> "" Then
        ' Open Explorer and select the file
        Shell "explorer.exe /select,""" & selectedPath & """", vbNormalFocus
    End If
End Sub

