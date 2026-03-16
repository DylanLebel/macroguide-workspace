Attribute VB_Name = "Author"
Option Explicit
Option Private Module
' Set this constant to True to enable "Modifiers" tracking, or False to disable it
Const ENABLE_MODIFIERS_TRACKING As Boolean = False

Sub PopulateAuthorCustomProperty(model As ModelDoc2, ByRef propertiesToSet As Object)
If model Is Nothing Then
Exit Sub
End If

' Get the logged-in user name
Dim username As String
username = Environ("USERNAME")

' Get the custom author name based on the logged-in user name
Dim authorName As String
authorName = GetCustomNameFromDictionary(username)

' If the author name is not found in the dictionary, use the username
If authorName = "" Then
    authorName = username
End If

' If the 'Author' property is not set or is empty, set it to the author name
If Not propertiesToSet.exists("Author") Or propertiesToSet("Author") = "" Then
    propertiesToSet("Author") = authorName
Else
    ''''''''debug.Print "Author already set to: " & propertiesToSet("Author")
End If

' Update the 'Modifiers' property with the current user if they are not the author
If ENABLE_MODIFIERS_TRACKING Then
    Dim modifiers As String
    If propertiesToSet.exists("Modifiers") Then
        modifiers = propertiesToSet("Modifiers")
    Else
        modifiers = ""
    End If
    
    If modifiers = "" Then
        If authorName <> propertiesToSet("Author") Then
            modifiers = authorName
        End If
    Else
        ' Check if the current user is already listed as a modifier
        If InStr(modifiers, authorName) = 0 Then
            modifiers = modifiers & ", " & authorName
        End If
    End If
    
    If modifiers <> "" Then
        propertiesToSet("Modifiers") = modifiers
    End If
End If
End Sub

Function GetCustomNameFromDictionary(username As String) As String
Dim customNamesDictionary As Object
Set customNamesDictionary = CreateObject("Scripting.Dictionary")
customNamesDictionary.Add "dlebel", "DGL"
customNamesDictionary.Add "hhood", "HH"
customNamesDictionary.Add "bcantin", "BJSC"
customNamesDictionary.Add "Engineer", "ENGINEER"
customNamesDictionary.Add "jgagnon", "JS"
customNamesDictionary.Add "tdesormeau", "TD"
customNamesDictionary.Add "kpatel", "KP"
customNamesDictionary.Add "gclouthier", "G.C."
customNamesDictionary.Add "jdelisle", "JD"
customNamesDictionary.Add "dturrie", "DT"
customNamesDictionary.Add "dshank", "DS"
customNamesDictionary.Add "ComputerName", "Initials"
' Add more mappings as needed

If customNamesDictionary.exists(username) Then
    GetCustomNameFromDictionary = customNamesDictionary(username)
Else
    GetCustomNameFromDictionary = ""
End If
End Function


