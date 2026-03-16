Attribute VB_Name = "ViewVersionTracking"
' ====================================================================================
' ViewVersionTracking.bas - View Macro Version Usage
' ====================================================================================
'
' PURPOSE:
'   Displays who is using the macro and what versions they're on.
'   Run this macro to see a report of all users and their version status.
'
' ENTRY POINT:
'   - ViewVersionReport() : Shows version tracking report in message box
'
' ====================================================================================

Option Explicit

' ====================================================================================
' ViewVersionReport - Display version tracking report
' ====================================================================================
' Run this to see who's using what version of the macro
' ====================================================================================
Sub ViewVersionReport()
    Dim trackingFilePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim reportText As String
    Dim recordParts() As String
    Dim username As String
    Dim version As String
    Dim lastUsed As String
    Dim totalRuns As String
    Dim lineCount As Integer
    Dim outdatedUsers As String
    Dim currentUsers As String
    Dim outdatedCount As Integer
    Dim currentCount As Integer
    Dim userResponse As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' Define tracking file path (hidden location)
    trackingFilePath = "Y:\Solidworks\Macros\Macro Data PDM\MacroVersionTracking.csv"

    ' Check if file exists
    If Dir(trackingFilePath) = "" Then
        MsgBox "Version tracking file not found." & vbCrLf & _
               "No one has run the macro yet, or the tracking file has been deleted." & vbCrLf & vbCrLf & _
               "Expected location: " & trackingFilePath, vbInformation, "Version Tracking"
        Exit Sub
    End If

    ' Read tracking file
    fileNum = FreeFile
    Open trackingFilePath For Input As #fileNum

    reportText = "MACRO VERSION TRACKING REPORT" & vbCrLf
    reportText = reportText & "Current Version: " & main.MacroVersion & vbCrLf
    reportText = reportText & String(60, "=") & vbCrLf & vbCrLf

    lineCount = 0
    outdatedCount = 0
    currentCount = 0
    outdatedUsers = ""
    currentUsers = ""

    ' Skip header
    If Not EOF(fileNum) Then Line Input #fileNum, lineText

    ' Read all records
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Trim$(lineText) <> "" Then
            recordParts = Split(lineText, ",")
            If UBound(recordParts) >= 3 Then
                username = recordParts(0)
                version = recordParts(1)
                lastUsed = recordParts(2)
                totalRuns = recordParts(3)

                Dim userLine As String
                userLine = username & " - v" & version & " (" & totalRuns & " runs, last: " & lastUsed & ")" & vbCrLf

                ' Categorize users
                If CLng(version) < main.MacroVersion Then
                    outdatedUsers = outdatedUsers & userLine
                    outdatedCount = outdatedCount + 1
                Else
                    currentUsers = currentUsers & userLine
                    currentCount = currentCount + 1
                End If

                lineCount = lineCount + 1
            End If
        End If
    Loop

    Close #fileNum

    ' Build report sections
    If currentCount > 0 Then
        reportText = reportText & "USERS ON CURRENT VERSION (" & currentCount & "):" & vbCrLf
        reportText = reportText & String(60, "-") & vbCrLf
        reportText = reportText & currentUsers & vbCrLf
    End If

    If outdatedCount > 0 Then
        reportText = reportText & "USERS ON OLDER VERSIONS (" & outdatedCount & "):" & vbCrLf
        reportText = reportText & String(60, "-") & vbCrLf
        reportText = reportText & outdatedUsers & vbCrLf
    End If

    If lineCount = 0 Then
        reportText = reportText & "No user records found." & vbCrLf
    End If

    reportText = reportText & String(60, "=") & vbCrLf
    reportText = reportText & "Total Users: " & lineCount & vbCrLf
    reportText = reportText & "Report Generated: " & Now()

    ' Display report with option to open file
    userResponse = MsgBox(reportText & vbCrLf & vbCrLf & _
                         "Would you like to open the CSV file in Excel?", _
                         vbInformation + vbYesNo, "Macro Version Tracking")

    If userResponse = vbYes Then
        OpenTrackingFileInExcel
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error reading version tracking file:" & vbCrLf & vbCrLf & _
           Err.description & " (Error " & Err.Number & ")" & vbCrLf & vbCrLf & _
           "File path: " & trackingFilePath, vbCritical, "Error"
End Sub

' ====================================================================================
' OpenTrackingFileInExcel - Open the CSV file in Excel
' ====================================================================================
Sub OpenTrackingFileInExcel()
    Dim trackingFilePath As String
    Dim excelApp As Object

    On Error GoTo ErrorHandler

    trackingFilePath = "Y:\Solidworks\Macros\Macro Data PDM\MacroVersionTracking.csv"

    ' Check if file exists
    If Dir(trackingFilePath) = "" Then
        MsgBox "Tracking file not found: " & trackingFilePath, vbExclamation
        Exit Sub
    End If

    ' Open in Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    excelApp.Workbooks.Open trackingFilePath

    Exit Sub

ErrorHandler:
    MsgBox "Error opening file in Excel:" & vbCrLf & vbCrLf & _
           Err.description & " (Error " & Err.Number & ")", vbCritical, "Error"
End Sub

' ====================================================================================
' QuickViewTracking - Quick desktop shortcut to view tracking (no message box)
' ====================================================================================
' Run this for instant Excel view without the message box report
' ====================================================================================
Sub QuickViewTracking()
    OpenTrackingFileInExcel
End Sub
