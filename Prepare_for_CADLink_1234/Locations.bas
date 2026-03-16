Attribute VB_Name = "Locations"
Option Explicit
Option Private Module

Public Function GetLocations() As collection
    Set GetLocations = New collection
    GetLocations.Add ".*\\NMT_PDM\\Libraries\\Design\\.*"
    GetLocations.Add ".*\\NMT_PDM\\Projects\\Customer\\.*"
    GetLocations.Add ".*\\NMT_PDM\\Projects\\Internal\\.*"
    GetLocations.Add ".*\\NMT_PDM\\Projects\\Capital\\.*"
    GetLocations.Add ".*\\NMT_PDM\\Products\\.*"
End Function
