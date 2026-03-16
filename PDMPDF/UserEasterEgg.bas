Attribute VB_Name = "UserEasterEgg"
' ============================================================================
' MODULE: UserEasterEgg
' Description: Show special image for specific users and other surprises
' ============================================================================
Option Explicit

' Windows API for pausing execution (Sleep)
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const EASTER_EGG_IMAGE As String = "Y:\Solidworks\Macros\Macro Data PDM\MACRO.bmp"

' Check if current user should see easter egg
Public Sub CheckForEasterEgg()
    Dim userName As String
    userName = Environ$("USERNAME")
    
    ' --- Get the SolidWorks Application ---
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then Exit Sub ' Can't run without SolidWorks
    
    ' Check if this is jgagnon
    If LCase(userName) = "jgagnon" Or LCase(userName) = "dlebel" Then
        Randomize
        
        ' Roll the dice for image - 1 in 10 chance
        If Int(Rnd * 10) = 0 Then
            ShowEasterEggImage
        End If
        
        ' Roll the dice for fake processing message - 1 in 12 chance
        ' (Fixed from Rnd * 2 to Rnd * 12)
        If Int(Rnd * 12) = 0 Then
            ' --- Pass swApp to the function ---
            ShowFakeProcessing swApp
        End If
    End If
End Sub

' Display the easter egg image
Private Sub ShowEasterEggImage()
    ' Check if image exists
    If Dir(EASTER_EGG_IMAGE) = "" Then
        Debug.Print "Easter egg image not found: " & EASTER_EGG_IMAGE
        Exit Sub
    End If
    
    ' Open the image with default viewer
    On Error Resume Next
    Shell "cmd /c start """" """ & EASTER_EGG_IMAGE & """", vbHide
    If Err.Number <> 0 Then
        Debug.Print "Failed to open easter egg image: " & Err.Description
        Err.Clear
    Else
        Debug.Print "Easter egg displayed for user: jgagnon"
    End If
    On Error GoTo 0
End Sub

' Display fake processing message in the SolidWorks status bar
Private Sub ShowFakeProcessing(swApp As SldWorks.SldWorks)
    If swApp Is Nothing Then Exit Sub ' Safety check

    ' --- The "Cool" Part: Define your list of fake steps ---
    Dim fakeMessages() As Variant
    fakeMessages = Array( _
        "Reticulating splines...", _
        "Calibrating flux capacitor...", _
        "Consulting magic 8-ball...", _
        "Polishing the widget bearings...", _
        "Dividing by zero... (retrying)...", _
        "Asking the magic conch...", _
        "Buffering 1s and 0s...")

    Dim i As Long
    Dim totalSteps As Long
    totalSteps = UBound(fakeMessages)

    On Error Resume Next ' In case user closes SW or something unexpected

    ' Loop through the messages using status bar only (progress bar methods not available)
    For i = 0 To totalSteps
        ' Update the status text
        swApp.Frame.SetStatusBarText fakeMessages(i)

        ' Pause for 3/4 of a second to make it readable
        Sleep 750
    Next i

    ' Clean up
    swApp.Frame.SetStatusBarText "PDF creation complete."

    Debug.Print "Fake processing messages shown for user."
    On Error GoTo 0
End Sub
