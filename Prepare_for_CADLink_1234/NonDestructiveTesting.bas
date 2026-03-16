Attribute VB_Name = "NonDestructiveTesting"
Option Explicit
Option Private Module

' Global variable to control whether SetNonDestructiveTesting should run
Public Const RUN_NON_DESTRUCTIVE_TESTING As Boolean = True

Sub SetNonDestructiveTesting(ByRef propertiesToSet As Scripting.Dictionary)
    If RUN_NON_DESTRUCTIVE_TESTING Then
        ' Check if the "Non-Destructive Testing" property exists and if it is empty
        If Not propertiesToSet.exists("Non-Destructive Testing") Then
            ' '''''debug.Print "Non-Destructive Testing property does not exist. Adding with value NONE."
            propertiesToSet("Non-Destructive Testing") = "NONE"
        ElseIf propertiesToSet("Non-Destructive Testing") = "" Then
            ' '''''debug.Print "Non-Destructive Testing property exists but is empty. Setting value to NONE."
            propertiesToSet("Non-Destructive Testing") = "NONE"
        Else
            ' '''''debug.Print "Non-Destructive Testing property already has a value: " & propertiesToSet("Non-Destructive Testing")
        End If
    Else
        ' '''''debug.Print "Skipping SetNonDestructiveTesting as RUN_NON_DESTRUCTIVE_TESTING is set to False."
    End If
End Sub
