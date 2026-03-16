VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "PDF Creation Progress"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Timing variables for ETA calculation
Private mStartTime As Double
Private mLastFileTime As Double         ' Time when last file completed
Private mRecentTimes() As Double        ' Sliding window of recent file processing times
Private mWindowSize As Long             ' Size of sliding window
Private mWindowIndex As Long            ' Current position in circular buffer
Private mWindowFilled As Boolean        ' Has the window been filled at least once?
Private mMedianBuffer() As Double       ' Temp buffer for median calculation

Private Sub UserForm_Initialize()
    ' PURPOSE: Initializes the form, omitting centering code that can cause errors at runtime.

    ' Set startup position behavior (CenterOwner might work better than manual if Initialize code fails)
    ' Me.StartUpPosition = 0 ' Manual positioning - KEEP COMMENTED OUT FOR NOW
    Me.StartUpPosition = 1 ' 1 = CenterOwner (Centers on SolidWorks window, generally safer)

    ' --- PROBLEM CODE COMMENTED OUT ---
    ' The Application object might not be reliable during Initialize, causing runtime errors.
'   Me.Left = Application.Left + (Application.Width - Me.Width) / 2
'   Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    ' --- END OF COMMENTED OUT CODE ---

    ' Initialize timing
    mStartTime = 0
    mLastFileTime = 0
    mWindowSize = 10        ' Track last 10 file times for median calculation
    mWindowIndex = 0
    mWindowFilled = False
    ReDim mRecentTimes(0 To mWindowSize - 1)
    ReDim mMedianBuffer(0 To mWindowSize - 1)

    ' Initialize progress display safely
    ' Call UpdateProgress to set initial text, even if progress bar width fails initially
    On Error Resume Next ' Temporarily ignore errors if controls aren't ready
    UpdateProgress 0, 0, ""
    On Error GoTo 0 ' Restore error handling
End Sub

' Updates the progress bar and labels
Public Sub UpdateProgress(currentFile As Long, totalFiles As Long, Optional fileStatus As String = "")
    Dim percentComplete As Single
    Dim elapsedSecs As Double
    Dim remainingSecs As Double
    Dim thisFileTime As Double      ' Time for current file
    Dim medianTime As Double        ' Median of recent file times
    Dim overallAvgTime As Double    ' Overall average time per file
    Dim estimatedTimePerFile As Double
    Dim statusText As String
    Dim titleText As String
    Dim currentTime As Double

    ' Prevent errors if controls aren't fully ready, especially during Initialize
    On Error Resume Next

    currentTime = Timer

    ' Initialize start time on first real update
    If mStartTime = 0 And totalFiles > 0 And currentFile > 0 Then
        mStartTime = currentTime
        mLastFileTime = currentTime
    End If

    ' Calculate percentage (handle division by zero)
    If totalFiles > 0 Then
        percentComplete = (currentFile / totalFiles) * 100
    Else
        percentComplete = 0
    End If

    ' Calculate elapsed time
    If mStartTime > 0 Then
        elapsedSecs = currentTime - mStartTime
        ' Handle midnight rollover
        If elapsedSecs < 0 Then elapsedSecs = elapsedSecs + 86400
    End If

    ' Track time for this file and calculate ETA
    If currentFile > 0 And mLastFileTime > 0 Then
        ' Calculate how long this file took
        thisFileTime = currentTime - mLastFileTime
        If thisFileTime < 0 Then thisFileTime = thisFileTime + 86400  ' Midnight rollover

        ' Only record reasonable times (ignore if too fast - likely skipped)
        If thisFileTime > 0.1 Then
            ' Add to sliding window (circular buffer)
            mRecentTimes(mWindowIndex) = thisFileTime
            mWindowIndex = (mWindowIndex + 1) Mod mWindowSize
            If mWindowIndex = 0 Then mWindowFilled = True
        End If

        ' Update for next file
        mLastFileTime = currentTime
    End If

    ' Calculate ETA using median of recent times (more robust to outliers)
    If currentFile > 2 And elapsedSecs > 0 Then
        ' Get median of recent file times
        medianTime = GetMedianTime()

        ' Calculate overall average as fallback
        overallAvgTime = elapsedSecs / currentFile

        ' Blend median with overall average based on progress
        ' Early: trust overall more (median may be noisy)
        ' Later: trust median more (it captures recent performance)
        If medianTime > 0 Then
            Dim progressWeight As Double
            progressWeight = percentComplete / 100  ' 0 to 1
            ' Use sigmoid-like weighting: slow start, then trust median more
            progressWeight = progressWeight * progressWeight  ' Square for slower ramp
            estimatedTimePerFile = (progressWeight * medianTime) + ((1 - progressWeight) * overallAvgTime)
        Else
            estimatedTimePerFile = overallAvgTime
        End If

        ' Calculate remaining time
        If estimatedTimePerFile > 0 Then
            remainingSecs = (totalFiles - currentFile) * estimatedTimePerFile
        End If
    End If

    ' --- Update Controls ---
    ' Check if Frame1 exists and has width before using it
    If Frame1.Width > 0 Then
        ' Check if ProgressBar exists before setting width
        If Not ProgressBar Is Nothing Then
             ProgressBar.Width = (Frame1.Width - 10) * (percentComplete / 100)
        End If
    End If

    ' Update status label with time info
    If Not lblStatus Is Nothing Then
        statusText = "Processing: " & currentFile & " of " & totalFiles & " files"
        If elapsedSecs > 0 Then
            statusText = statusText & "  |  Elapsed: " & FormatSeconds(elapsedSecs)
        End If
        If remainingSecs > 0 And currentFile < totalFiles Then
            statusText = statusText & "  |  ETA: " & FormatSeconds(remainingSecs)
        End If
        lblStatus.Caption = statusText
    End If

    If Not lblPercent Is Nothing Then
        lblPercent.Caption = Format(percentComplete, "0") & "%"
    End If

    ' Show current file information if provided (Check if label exists)
    If fileStatus <> "" Then
        If Not lblCurrentFile Is Nothing Then
            lblCurrentFile.Caption = fileStatus
        End If
    End If

    ' Update title bar with progress
    titleText = "PDF Progress: " & currentFile & "/" & totalFiles
    If percentComplete >= 100 Then
        titleText = "PDF Creation Complete!"
    End If
    Me.Caption = titleText

    On Error GoTo 0 ' Restore error handling

    ' Force immediate visual update (crucial for responsiveness)
    DoEvents
End Sub

' Format seconds as MM:SS or HH:MM:SS
Private Function FormatSeconds(secs As Double) As String
    Dim h As Long, m As Long, s As Long

    If secs < 0 Then secs = 0

    h = Int(secs / 3600)
    m = Int((secs - h * 3600) / 60)
    s = Int(secs - h * 3600 - m * 60)

    If h > 0 Then
        FormatSeconds = h & ":" & Format(m, "00") & ":" & Format(s, "00")
    Else
        FormatSeconds = m & ":" & Format(s, "00")
    End If
End Function

' Reset timer for new batch
Public Sub ResetTimer()
    Dim i As Long
    mStartTime = Timer
    mLastFileTime = Timer
    mWindowIndex = 0
    mWindowFilled = False
    ' Clear the sliding window
    For i = 0 To mWindowSize - 1
        mRecentTimes(i) = 0
    Next i
End Sub

' Calculate median of recent file times (robust to outliers)
Private Function GetMedianTime() As Double
    Dim count As Long
    Dim i As Long, j As Long
    Dim temp As Double

    ' Determine how many valid times we have
    If mWindowFilled Then
        count = mWindowSize
    Else
        count = mWindowIndex
    End If

    ' Need at least 3 samples for meaningful median
    If count < 3 Then
        GetMedianTime = 0
        Exit Function
    End If

    ' Copy to temp buffer for sorting
    For i = 0 To count - 1
        mMedianBuffer(i) = mRecentTimes(i)
    Next i

    ' Simple bubble sort (window is small, so this is fine)
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If mMedianBuffer(j) < mMedianBuffer(i) Then
                temp = mMedianBuffer(i)
                mMedianBuffer(i) = mMedianBuffer(j)
                mMedianBuffer(j) = temp
            End If
        Next j
    Next i

    ' Return median (middle value, or average of two middle for even count)
    If count Mod 2 = 1 Then
        GetMedianTime = mMedianBuffer(count \ 2)
    Else
        GetMedianTime = (mMedianBuffer(count \ 2 - 1) + mMedianBuffer(count \ 2)) / 2
    End If
End Function

' Handles the Cancel button click
Private Sub btnCancel_Click()
    ' Set a global variable (declared Public in a standard module, e.g., PDMPDF1)
    ' to indicate cancellation
    On Error Resume Next ' In case CancelOperation isn't declared Public correctly
    CancelOperation = True
    If Err.Number <> 0 Then
        MsgBox "Error setting CancelOperation. Ensure it's declared as 'Public CancelOperation As Boolean' in a standard module.", vbExclamation
    End If
    On Error GoTo 0
    
    ' Hide the form
    Me.Hide
End Sub
