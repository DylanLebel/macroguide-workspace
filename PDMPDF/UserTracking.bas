Attribute VB_Name = "UserTracking"
' ============================================================================
' MODULE: UserTracking
' Description: Track macro users and show onboarding message for first-time users
' UPDATED: File location moved outside PDM vault to avoid synchronization issues
' ============================================================================

Option Explicit

' Centralized user tracking file location (on network share - OUTSIDE PDM vault)
Private Const USER_TRACKING_FILE As String = "Y:\Solidworks\Macros\Macro Data PDM\PDF_Macro_Users.txt"

' Function to get the current computer name
Private Function GetComputerName() As String
    GetComputerName = Environ$("COMPUTERNAME")
End Function

' Function to get the current username
Private Function GetUsername() As String
    GetUsername = Environ$("USERNAME")
End Function

' Function to get unique identifier for this user/computer
Private Function GetUserIdentifier() As String
    ' Combine computer name and username for unique ID
    GetUserIdentifier = GetComputerName() & "_" & GetUsername()
End Function

' Ensure the tracking file directory and file exist
Private Sub EnsureTrackingFileExists()
    Dim trackingDir As String
    
    trackingDir = Left(USER_TRACKING_FILE, InStrRev(USER_TRACKING_FILE, "\"))
    
    ' Create directory if it doesn't exist
    If Dir(trackingDir, vbDirectory) = "" Then
        On Error Resume Next
        MkDir trackingDir
        If Err.Number <> 0 Then
            Debug.Print "Could not create tracking directory: " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Debug.Print "Created tracking directory: " & trackingDir
    End If
    
    ' Create file if it doesn't exist
    If Dir(USER_TRACKING_FILE) = "" Then
        Dim fileNum As Integer
        fileNum = FreeFile
        
        On Error Resume Next
        Open USER_TRACKING_FILE For Output As #fileNum
        If Err.Number <> 0 Then
            Debug.Print "Could not create tracking file: " & Err.Description
            Err.Clear
            Close #fileNum
            On Error GoTo 0
            Exit Sub
        End If
        
        Print #fileNum, "UserID,Username,ComputerName,FirstRun,LastRun"
        Close #fileNum
        
        If Err.Number = 0 Then
            Debug.Print "Created tracking file: " & USER_TRACKING_FILE
        Else
            Debug.Print "Error creating tracking file: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' Check if this is the user's first time running the macro
Private Function IsFirstTimeUser() As Boolean
    Dim userID As String
    Dim fileNum As Integer
    Dim fileLine As String
    Dim found As Boolean
    
    userID = GetUserIdentifier()
    found = False
    
    EnsureTrackingFileExists
    
    ' Read the file and look for this user
    On Error Resume Next
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Input As #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "Could not open tracking file for reading: " & Err.Description
        Err.Clear
        On Error GoTo 0
        IsFirstTimeUser = True ' Assume first time if we can't read
        Exit Function
    End If
    
    ' Skip header
    If Not EOF(fileNum) Then Line Input #fileNum, fileLine
    
    ' Check each line
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileLine
        If InStr(1, fileLine, userID, vbTextCompare) > 0 Then
            found = True
            Exit Do
        End If
    Loop
    
    Close #fileNum
    On Error GoTo 0
    
    IsFirstTimeUser = Not found
End Function

' Mark that this user has run the macro
Private Sub MarkUserAsRun()
    Dim userID As String
    Dim fileNum As Integer
    Dim userName As String
    Dim computerName As String
    Dim currentDateTime As String
    Dim isFirstTime As Boolean
    
    userID = GetUserIdentifier()
    userName = GetUsername()
    computerName = GetComputerName()
    currentDateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    EnsureTrackingFileExists
    
    ' Check if this is truly first time (before we modify anything)
    isFirstTime = IsFirstTimeUser()
    
    ' Check if user already exists (update last run) or add new
    If Not isFirstTime Then
        ' User exists - update last run time
        UpdateUserLastRun userID, currentDateTime
        Debug.Print "Updated last run for user: " & userID
    Else
        ' New user - add to file
        On Error Resume Next
        fileNum = FreeFile
        Open USER_TRACKING_FILE For Append As #fileNum
        
        If Err.Number <> 0 Then
            Debug.Print "Could not open tracking file for append: " & Err.Description
            Err.Clear
            Close #fileNum
            On Error GoTo 0
            Exit Sub
        End If
        
        Print #fileNum, userID & "," & userName & "," & computerName & "," & currentDateTime & "," & currentDateTime
        Close #fileNum
        
        If Err.Number = 0 Then
            Debug.Print "Added new user to tracking: " & userID
        Else
            Debug.Print "Error adding user to tracking: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' Update the last run time for an existing user
Private Sub UpdateUserLastRun(userID As String, lastRun As String)
    Dim fileNum As Integer
    Dim fileLine As String
    Dim fileContents As String
    Dim lineArray() As String
    Dim i As Integer
    
    On Error Resume Next
    
    ' Read entire file
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Input As #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "Could not open tracking file for update: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    fileContents = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileLine
        fileContents = fileContents & fileLine & vbCrLf
    Loop
    Close #fileNum
    
    ' Update the line with this user's info
    lineArray = Split(fileContents, vbCrLf)
    For i = 0 To UBound(lineArray)
        If InStr(1, lineArray(i), userID, vbTextCompare) > 0 Then
            ' Update last run time (5th column)
            Dim parts() As String
            parts = Split(lineArray(i), ",")
            If UBound(parts) >= 4 Then
                parts(4) = lastRun
                lineArray(i) = Join(parts, ",")
            End If
            Exit For
        End If
    Next i
    
    ' Write back to file
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Output As #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "Could not open tracking file for writing: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    For i = 0 To UBound(lineArray)
        If Trim(lineArray(i)) <> "" Then
            Print #fileNum, lineArray(i)
        End If
    Next i
    Close #fileNum
    
    On Error GoTo 0
End Sub

' Main entry point - call this at the start of main()
Public Sub CheckUserAndShowInstructions()
    ' Wrap in error handler so tracking issues don't break macro
    On Error Resume Next
    
    If IsFirstTimeUser() Then
        ShowWelcomeInstructions
        MarkUserAsRun
    Else
        ' Not first time - just update the last run timestamp
        MarkUserAsRun
    End If
    
    ' Clear any errors - user tracking is not critical
    If Err.Number <> 0 Then
        Debug.Print "User tracking error (non-critical): " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Display welcome/instruction message for first-time users
Private Sub ShowWelcomeInstructions()
    Dim msg As String
    
    msg = "=====================================" & vbCrLf
    msg = msg & "  PDF CREATION MACRO - INSTRUCTIONS" & vbCrLf
    msg = msg & "=====================================" & vbCrLf & vbCrLf
    
    msg = msg & "This macro creates and manages PDF files from SolidWorks drawings." & vbCrLf & vbCrLf
    
    msg = msg & "--- HOW TO USE ---" & vbCrLf & vbCrLf
    
    msg = msg & "SINGLE DRAWING MODE:" & vbCrLf
    msg = msg & "  1. Open a drawing in SolidWorks" & vbCrLf
    msg = msg & "  2. Run the macro (Alt+F8 > main)" & vbCrLf
    msg = msg & "  3. Macro will sync revisions and create PDF" & vbCrLf & vbCrLf
    
    msg = msg & "BATCH MODE:" & vbCrLf
    msg = msg & "  1. Close all drawings (NO drawing open)" & vbCrLf
    msg = msg & "  2. Run the macro" & vbCrLf
    msg = msg & "  3. Choose to include subfolders" & vbCrLf
    msg = msg & "  4. Select folder with drawings" & vbCrLf
    msg = msg & "  5. Macro processes all automatically" & vbCrLf & vbCrLf
    
    msg = msg & "ASSEMBLY MODE:" & vbCrLf
    msg = msg & "  1. Open an assembly in SolidWorks" & vbCrLf
    msg = msg & "  2. Run the macro" & vbCrLf
    msg = msg & "  3. Processes all component drawings" & vbCrLf
    msg = msg & "  4. Library items excluded automatically"
    
    MsgBox msg, vbInformation, "PDF Creation Macro - Instructions (1/2)"
    
    ' Second message box with additional info
    msg = "=====================================" & vbCrLf
    msg = msg & "  IMPORTANT TO KNOW" & vbCrLf
    msg = msg & "=====================================" & vbCrLf & vbCrLf
    
    msg = msg & "- Syncs model revisions automatically" & vbCrLf
    msg = msg & "- Warns about unsaved changes" & vbCrLf
    msg = msg & "- Checks PDM vault status" & vbCrLf
    msg = msg & "- May close/reopen documents if needed" & vbCrLf
    msg = msg & "- Creates folder structure automatically" & vbCrLf
    msg = msg & "- PDF format: [FileName]_Rev[#].pdf" & vbCrLf
    msg = msg & "- Old PDFs moved to Obsolete folder" & vbCrLf & vbCrLf
    
    msg = msg & "--- TIPS ---" & vbCrLf & vbCrLf
    
    msg = msg & "- Save your work before running" & vbCrLf
    msg = msg & "- Check out files from PDM if needed" & vbCrLf
    msg = msg & "- Progress bars show status" & vbCrLf
    msg = msg & "- You only see this message once" & vbCrLf & vbCrLf
    
    msg = msg & "Need help? Ask Dylan!" & vbCrLf & vbCrLf
    
    msg = msg & "====================================="
    
    MsgBox msg, vbInformation, "PDF Creation Macro - Instructions (2/2)"
End Sub

' Get list of all users (for admin use)
Public Function GetAllUsers() As String
    Dim fileNum As Integer
    Dim fileLine As String
    Dim userList As String
    
    EnsureTrackingFileExists
    
    userList = ""
    
    On Error Resume Next
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Input As #fileNum
    
    If Err.Number <> 0 Then
        userList = "Could not read tracking file: " & Err.Description
        Err.Clear
        On Error GoTo 0
        GetAllUsers = userList
        Exit Function
    End If
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileLine
        userList = userList & fileLine & vbCrLf
    Loop
    
    Close #fileNum
    On Error GoTo 0
    
    If userList = "" Then
        userList = "No users have run the macro yet."
    End If
    
    GetAllUsers = userList
End Function

' Show all users in a message box
Public Sub ShowAllUsers()
    Dim userList As String
    userList = GetAllUsers()
    
    MsgBox userList, vbInformation, "PDF Macro - All Users"
End Sub

' Optional: Reset user data (for testing)
Public Sub ResetUserTracking()
    Dim userID As String
    
    userID = GetUserIdentifier()
    
    ' Remove this user from the file
    RemoveUserFromFile userID
    
    MsgBox "User tracking data has been reset for: " & userID & vbCrLf & _
           "You will see the instructions on next run.", vbInformation
    
    Debug.Print "Reset tracking for user: " & userID
End Sub

' Remove a user from the tracking file
Private Sub RemoveUserFromFile(userID As String)
    Dim fileNum As Integer
    Dim fileLine As String
    Dim fileContents As String
    Dim lineArray() As String
    Dim i As Integer
    
    On Error Resume Next
    
    ' Read entire file
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Input As #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "Could not open tracking file for removal: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    fileContents = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileLine
        fileContents = fileContents & fileLine & vbCrLf
    Loop
    Close #fileNum
    
    ' Remove the line with this user's info
    lineArray = Split(fileContents, vbCrLf)
    
    ' Write back to file, excluding the user's line
    fileNum = FreeFile
    Open USER_TRACKING_FILE For Output As #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "Could not open tracking file for writing: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    
    For i = 0 To UBound(lineArray)
        If Trim(lineArray(i)) <> "" Then
            If InStr(1, lineArray(i), userID, vbTextCompare) = 0 Then
                Print #fileNum, lineArray(i)
            End If
        End If
    Next i
    Close #fileNum
    
    On Error GoTo 0
End Sub

' Get current user info (for debugging)
Public Function GetUserInfo() As String
    Dim info As String
    Dim userID As String
    
    userID = GetUserIdentifier()
    
    info = "Current User: " & GetUsername() & vbCrLf
    info = info & "Computer: " & GetComputerName() & vbCrLf
    info = info & "User ID: " & userID & vbCrLf
    info = info & "Is First Time: " & IsFirstTimeUser()
    
    GetUserInfo = info
End Function


