Attribute VB_Name = "FlatBarDataModule"
Option Explicit
Option Private Module

' Type to hold flat bar dimensions (Private so it stays in this module only)
Private Type flatBarSize
    material As String
    Specification As String
    ThicknessFrac As String
    WidthFrac As String
    ThicknessDec As Double
    WidthDec As Double
    WeightPerFt As Double
End Type

' Collections to hold flat bar data
Private steelFlatBars As collection
Private aluminumFlatBars As collection
Private dataLoaded As Boolean

' File paths - Network share location for flat bar data
Private Const FLAT_BAR_DATA_PATH As String = "Y:\Solidworks\Macros\Macro Data PDM\Flat Bar Data\"
Private Const STEEL_44W_CSV As String = "44W_flat_bar.csv"
Private Const STEEL_50W_CSV As String = "50W_flat_bar.csv"
Private Const STEEL_300W_CSV As String = "300W_flat_bar.csv"
Private Const STEEL_350W_CSV As String = "350W_flat_bar.csv"
Private Const ALUMINUM_6061T6_CSV As String = "6061T6_flat_bar.csv"

' Tolerance for dimension matching (inches)
Private Const DIMENSION_TOLERANCE As Double = 0.001

' Initialize and load CSV data
Public Sub InitializeFlatBarData()
    On Error GoTo ErrorHandler

    If dataLoaded Then Exit Sub

    Set steelFlatBars = New collection
    Set aluminumFlatBars = New collection

    DebugLog "=== LOADING FLAT BAR DATA ==="
    DebugLog "Data path: " & FLAT_BAR_DATA_PATH

    Dim filePath As String

    ' Load 44W steel flat bars
    filePath = FLAT_BAR_DATA_PATH & STEEL_44W_CSV
    If Dir(filePath) <> "" Then
        LoadCSVFile filePath, steelFlatBars
        DebugLog "Loaded 44W steel flat bars from: " & filePath
    Else
        DebugLog "WARNING: 44W CSV not found at: " & filePath
    End If

    ' Load 50W steel flat bars
    filePath = FLAT_BAR_DATA_PATH & STEEL_50W_CSV
    If Dir(filePath) <> "" Then
        LoadCSVFile filePath, steelFlatBars
        DebugLog "Loaded 50W steel flat bars from: " & filePath
    Else
        DebugLog "WARNING: 50W CSV not found at: " & filePath
    End If

    ' Load 300W steel flat bars
    filePath = FLAT_BAR_DATA_PATH & STEEL_300W_CSV
    If Dir(filePath) <> "" Then
        LoadCSVFile filePath, steelFlatBars
        DebugLog "Loaded 300W steel flat bars from: " & filePath
    Else
        DebugLog "WARNING: 300W CSV not found at: " & filePath
    End If

    ' Load 350W steel flat bars
    filePath = FLAT_BAR_DATA_PATH & STEEL_350W_CSV
    If Dir(filePath) <> "" Then
        LoadCSVFile filePath, steelFlatBars
        DebugLog "Loaded 350W steel flat bars from: " & filePath
    Else
        DebugLog "WARNING: 350W CSV not found at: " & filePath
    End If

    ' Load aluminum 6061-T6 flat bars
    filePath = FLAT_BAR_DATA_PATH & ALUMINUM_6061T6_CSV
    If Dir(filePath) <> "" Then
        LoadCSVFile filePath, aluminumFlatBars
        DebugLog "Loaded 6061-T6 aluminum flat bars from: " & filePath
    Else
        DebugLog "WARNING: 6061-T6 CSV not found at: " & filePath
    End If

    DebugLog "Total steel flat bars loaded: " & steelFlatBars.Count
    DebugLog "Total aluminum flat bars loaded: " & aluminumFlatBars.Count

    dataLoaded = True
    Exit Sub

ErrorHandler:
    DebugLog "ERROR in InitializeFlatBarData: " & Err.description
    dataLoaded = False
End Sub

' Load CSV file into collection
Private Sub LoadCSVFile(ByVal filePath As String, ByRef collection As collection)
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim fileContent As String
    Dim lines() As String
    Dim fields() As String
    Dim fb As flatBarSize
    Dim lineCount As Long
    Dim i As Long
    Dim lineText As String

    ' Read entire file at once to avoid EOF issues with different line endings
    fileNum = FreeFile
    Open filePath For Binary As #fileNum

    If LOF(fileNum) = 0 Then
        DebugLog "  ERROR: File is empty: " & filePath
        Close #fileNum
        Exit Sub
    End If

    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum

    ' Normalize line endings (handle both CRLF and LF)
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)

    ' Split into lines
    lines = Split(fileContent, vbLf)

    DebugLog "  File loaded: " & UBound(lines) + 1 & " lines from " & filePath

    ' Skip first line (header)
    If UBound(lines) < 1 Then
        DebugLog "  ERROR: File has no data lines"
        Exit Sub
    End If

    DebugLog "  Header: " & lines(0)

    lineCount = 0
    ' Process data lines starting from index 1
    For i = 1 To UBound(lines)
        lineText = Trim(lines(i))

        If lineText <> "" Then
            fields = Split(lineText, ",")

            If UBound(fields) >= 6 Then
                On Error Resume Next
                fb.material = Trim(fields(0))
                fb.Specification = Trim(fields(1))
                fb.ThicknessFrac = Trim(fields(2))
                fb.WidthFrac = Trim(fields(3))
                fb.ThicknessDec = CDbl(Trim(fields(4)))
                fb.WidthDec = CDbl(Trim(fields(5)))
                fb.WeightPerFt = CDbl(Trim(fields(6)))

                If Err.Number <> 0 Then
                    DebugLog "    ERROR parsing line " & i & ": " & Err.description
                    Err.Clear
                Else
                    ' Store as pipe-delimited string: ThicknessDec|WidthDec|ThicknessFrac|WidthFrac
                    Dim dataStr As String
                    dataStr = CStr(fb.ThicknessDec) & "|" & CStr(fb.WidthDec) & "|" & fb.ThicknessFrac & "|" & fb.WidthFrac
                    collection.Add dataStr
                    lineCount = lineCount + 1
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

    DebugLog "  Loaded " & lineCount & " entries"
    Exit Sub

ErrorHandler:
    DebugLog "  ERROR in LoadCSVFile: " & Err.description
    If fileNum > 0 Then Close #fileNum
End Sub

' Check if dimensions match a flat bar size
' Returns flat bar size as string (e.g., "1/4 x 2") or empty string if not found
Public Function FindMatchingFlatBar(ByVal thickness As Double, ByVal width As Double, ByVal MaterialType As String) As String
    On Error GoTo ErrorHandler

    FindMatchingFlatBar = ""  ' Default to empty string (not found)

    InitializeFlatBarData

    Dim collection As collection
    Dim fb As flatBarSize
    Dim i As Long

    ' Select appropriate collection based on material
    Select Case UCase(MaterialType)
        Case "STEEL"
            If steelFlatBars Is Nothing Then
                DebugLog "ERROR: Steel flat bars collection is Nothing"
                Exit Function
            End If
            Set collection = steelFlatBars
        Case "ALUMINUM"
            If aluminumFlatBars Is Nothing Then
                DebugLog "ERROR: Aluminum flat bars collection is Nothing"
                Exit Function
            End If
            Set collection = aluminumFlatBars
        Case Else
            DebugLog "ERROR: Unknown material type: " & MaterialType
            Exit Function
    End Select

    ' Search for matching dimensions
    ' Store each item's data in an array to work around the Private Type limitation
    Dim itemData() As String

    For i = 1 To collection.Count
        ' Get item as array: [ThicknessDec, WidthDec, ThicknessFrac, WidthFrac]
        itemData = Split(CStr(collection(i)), "|")

        ' Check if both thickness and width match within tolerance
        If Abs(CDbl(itemData(0)) - thickness) <= DIMENSION_TOLERANCE And _
           Abs(CDbl(itemData(1)) - width) <= DIMENSION_TOLERANCE Then
            FindMatchingFlatBar = itemData(2) & " x " & itemData(3)
            DebugLog "  Found matching flat bar: " & FindMatchingFlatBar
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    DebugLog "ERROR in FindMatchingFlatBar: " & Err.description & " (Error " & Err.Number & ")"
    FindMatchingFlatBar = ""
End Function

' Get all available thicknesses for a material type
Public Function GetAvailableThicknesses(ByVal MaterialType As String) As collection
    InitializeFlatBarData

    Dim collection As collection
    Dim result As New collection
    Dim i As Long
    Dim thicknessDict As Object
    Dim itemData() As String
    Set thicknessDict = CreateObject("Scripting.Dictionary")

    ' Select appropriate collection based on material
    Select Case UCase(MaterialType)
        Case "STEEL"
            Set collection = steelFlatBars
        Case "ALUMINUM"
            Set collection = aluminumFlatBars
        Case Else
            Set GetAvailableThicknesses = result
            Exit Function
    End Select

    ' Collect unique thicknesses
    For i = 1 To collection.Count
        ' Parse pipe-delimited string: ThicknessDec|WidthDec|ThicknessFrac|WidthFrac
        itemData = Split(CStr(collection(i)), "|")

        If Not thicknessDict.exists(CDbl(itemData(0))) Then
            thicknessDict.Add CDbl(itemData(0)), itemData(2)
            result.Add CDbl(itemData(0))
        End If
    Next i

    Set GetAvailableThicknesses = result
End Function

' Get all available widths for a specific thickness and material type
Public Function GetAvailableWidths(ByVal thickness As Double, ByVal MaterialType As String) As collection
    InitializeFlatBarData

    Dim collection As collection
    Dim result As New collection
    Dim i As Long
    Dim itemData() As String

    ' Select appropriate collection based on material
    Select Case UCase(MaterialType)
        Case "STEEL"
            Set collection = steelFlatBars
        Case "ALUMINUM"
            Set collection = aluminumFlatBars
        Case Else
            Set GetAvailableWidths = result
            Exit Function
    End Select

    ' Collect widths for matching thickness
    For i = 1 To collection.Count
        ' Parse pipe-delimited string: ThicknessDec|WidthDec|ThicknessFrac|WidthFrac
        itemData = Split(CStr(collection(i)), "|")

        If Abs(CDbl(itemData(0)) - thickness) <= DIMENSION_TOLERANCE Then
            result.Add CDbl(itemData(1))
        End If
    Next i

    Set GetAvailableWidths = result
End Function

' Force reload of data (useful if CSV files are updated)
Public Sub ReloadFlatBarData()
    dataLoaded = False
    Set steelFlatBars = Nothing
    Set aluminumFlatBars = Nothing
    InitializeFlatBarData
End Sub

' Debug function to print all loaded data
Public Sub DebugPrintFlatBarData()
    InitializeFlatBarData

    Dim i As Long
    Dim itemData() As String

    DebugLog "=== STEEL FLAT BARS ==="
    For i = 1 To steelFlatBars.Count
        ' Parse pipe-delimited string: ThicknessDec|WidthDec|ThicknessFrac|WidthFrac
        itemData = Split(CStr(steelFlatBars(i)), "|")
        DebugLog itemData(2) & " x " & itemData(3) & " (" & itemData(0) & " x " & itemData(1) & ")"
    Next i

    DebugLog "=== ALUMINUM FLAT BARS ==="
    For i = 1 To aluminumFlatBars.Count
        ' Parse pipe-delimited string: ThicknessDec|WidthDec|ThicknessFrac|WidthFrac
        itemData = Split(CStr(aluminumFlatBars(i)), "|")
        DebugLog itemData(2) & " x " & itemData(3) & " (" & itemData(0) & " x " & itemData(1) & ")"
    Next i
End Sub
