Attribute VB_Name = "Module3"
' Export all modules from ALL OPEN projects to individual subfolders
Sub ExportAllOpenProjects()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim basePath As String
    Dim projExportPath As String
    Dim fileName As String
    Dim exportCount As Long
    Dim projCount As Long

    On Error GoTo ErrHandler

    ' Your specific export folder
    basePath = "C:\AllMacros\"

    ' Create the main directory if it doesn't exist
    If Dir(basePath, vbDirectory) = vbNullString Then MkDir basePath

    exportCount = 0
    projCount = 0

    ' Loop through every open project in the VBA Editor
    For Each vbProj In Application.VBE.VBProjects
        
        ' Create a specific subfolder for this project based on its name
        projExportPath = basePath & vbProj.Name & "\"
        
        ' Create the subfolder if it doesn't exist
        If Dir(projExportPath, vbDirectory) = vbNullString Then MkDir projExportPath
        
        projCount = projCount + 1

        ' Loop through components of the current project
        For Each vbComp In vbProj.VBComponents
            Select Case vbComp.Type
                Case vbext_ct_StdModule
                    fileName = projExportPath & vbComp.Name & ".bas"
                    vbComp.Export fileName
                    exportCount = exportCount + 1

                Case vbext_ct_ClassModule
                    fileName = projExportPath & vbComp.Name & ".cls"
                    vbComp.Export fileName
                    exportCount = exportCount + 1

                Case vbext_ct_MSForm
                    fileName = projExportPath & vbComp.Name & ".frm"
                    vbComp.Export fileName
                    exportCount = exportCount + 1

                Case vbext_ct_Document
                    ' These are ThisSW or ThisMacro
                    fileName = projExportPath & vbComp.Name & ".cls"
                    vbComp.Export fileName
                    exportCount = exportCount + 1
            End Select
        Next vbComp
    Next vbProj

    MsgBox "Batch Export Complete!" & vbCrLf & _
           "Exported " & exportCount & " total module(s) across " & projCount & " project(s) to:" & vbCrLf & basePath, _
           vbInformation, "Export Successful"
    Exit Sub

ErrHandler:
    MsgBox "Error exporting modules: " & Err.description & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           "Failed on Project: " & vbProj.Name, vbCritical, "Export Error"
End Sub

