Option Explicit

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_Document = 100

Dim args
Dim scriptFolder
Dim repoRoot
Dim workbookPath
Dim excelApp
Dim workbook
Dim component

Set args = WScript.Arguments
scriptFolder = GetParentFolder(WScript.ScriptFullName)
repoRoot = GetParentFolder(scriptFolder)

If ShouldShowHelp(args) Then
    ShowUsage
    WScript.Quit 0
End If

If args.Count >= 1 Then
    workbookPath = ToAbsolutePath(args(0), repoRoot)
Else
    workbookPath = repoRoot & "\workbooks\ReportAutomationTemplate.xlsm"
End If

If Not FileExists(workbookPath) Then
    Err.Raise vbObjectError + 2200, "export-vba-modules.vbs", "Workbook not found: " & workbookPath
End If

Set excelApp = CreateExcelApplication()
excelApp.Visible = False
excelApp.DisplayAlerts = False

Set workbook = excelApp.Workbooks.Open(workbookPath)

For Each component In workbook.VBProject.VBComponents
    ExportComponent component, repoRoot
Next

workbook.Close False
excelApp.Quit

WScript.Echo "VBA source exported from: " & workbookPath

Sub ExportComponent(ByVal componentObject, ByVal projectRoot)
    Dim exportPath

    Select Case componentObject.Type
        Case vbext_ct_StdModule
            EnsureFolderExists projectRoot & "\src\vba\modules"
            exportPath = projectRoot & "\src\vba\modules\" & componentObject.Name & ".bas"
        Case vbext_ct_ClassModule
            EnsureFolderExists projectRoot & "\src\vba\classes"
            exportPath = projectRoot & "\src\vba\classes\" & componentObject.Name & ".cls"
        Case vbext_ct_MSForm
            EnsureFolderExists projectRoot & "\src\vba\forms"
            exportPath = projectRoot & "\src\vba\forms\" & componentObject.Name & ".frm"
        Case vbext_ct_Document
            Exit Sub
        Case Else
            Exit Sub
    End Select

    DeleteFileIfExists exportPath
    componentObject.Export exportPath
End Sub

Function CreateExcelApplication()
    On Error Resume Next
    Set CreateExcelApplication = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Dim createExcelMessage
        createExcelMessage = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 2201, "export-vba-modules.vbs", "Unable to start Excel. " & createExcelMessage
    End If
    On Error GoTo 0
End Function

Sub DeleteFileIfExists(ByVal filePath)
    Dim fileSystem

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(filePath) Then
        fileSystem.DeleteFile filePath, True
    End If
End Sub

Sub EnsureFolderExists(ByVal folderPath)
    Dim fileSystem

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub

Function FileExists(ByVal filePath)
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
End Function

Function ToAbsolutePath(ByVal candidatePath, ByVal baseFolder)
    Dim fileSystem

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.GetDriveName(candidatePath) <> "" Then
        ToAbsolutePath = candidatePath
    Else
        ToAbsolutePath = fileSystem.BuildPath(baseFolder, candidatePath)
    End If
End Function

Function GetParentFolder(ByVal filePath)
    GetParentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(filePath)
End Function

Function ShouldShowHelp(ByVal commandArgs)
    If commandArgs.Count = 0 Then
        ShouldShowHelp = False
    Else
        ShouldShowHelp = (commandArgs(0) = "/?" Or commandArgs(0) = "-h" Or commandArgs(0) = "--help")
    End If
End Function

Sub ShowUsage()
    WScript.Echo "Usage: cscript //nologo scripts\export-vba-modules.vbs [workbookPath]"
    WScript.Echo "Exports standard, class, and form modules from the workbook back into src\vba\."
    WScript.Echo "Excel must be installed and Trust access to the VBA project object model must be enabled."
End Sub
