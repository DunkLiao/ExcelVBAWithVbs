Option Explicit

Const xlOpenXMLWorkbookMacroEnabled = 52

Dim args
Dim scriptFolder
Dim repoRoot
Dim workbookPath
Dim excelApp
Dim workbook

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

EnsureFolderExists GetParentFolder(workbookPath)

Set excelApp = CreateExcelApplication()
excelApp.Visible = False
excelApp.DisplayAlerts = False

Set workbook = excelApp.Workbooks.Add()
PrepareWorkbook workbook
ImportModule workbook, repoRoot & "\src\vba\modules\ReportAutomation.bas"
workbook.SaveAs workbookPath, xlOpenXMLWorkbookMacroEnabled
workbook.Close False
excelApp.Quit

WScript.Echo "Workbook created: " & workbookPath

Sub PrepareWorkbook(ByVal workbookObject)
    Do While workbookObject.Worksheets.Count < 2
        workbookObject.Worksheets.Add , workbookObject.Worksheets(workbookObject.Worksheets.Count)
    Loop

    workbookObject.Worksheets(1).Name = "InputData"
    workbookObject.Worksheets(2).Name = "Report"

    Do While workbookObject.Worksheets.Count > 2
        workbookObject.Worksheets(workbookObject.Worksheets.Count).Delete
    Loop
End Sub

Sub ImportModule(ByVal workbookObject, ByVal modulePath)
    If Not FileExists(modulePath) Then
        Err.Raise vbObjectError + 2000, "bootstrap-report-workbook.vbs", "Module file not found: " & modulePath
    End If

    workbookObject.VBProject.VBComponents.Import modulePath
End Sub

Function CreateExcelApplication()
    On Error Resume Next
    Set CreateExcelApplication = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Dim createExcelMessage
        createExcelMessage = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 2001, "bootstrap-report-workbook.vbs", "Unable to start Excel. " & createExcelMessage
    End If
    On Error GoTo 0
End Function

Function FileExists(ByVal filePath)
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
End Function

Sub EnsureFolderExists(ByVal folderPath)
    Dim fileSystem

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub

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
    WScript.Echo "Usage: cscript //nologo scripts\bootstrap-report-workbook.vbs [outputWorkbookPath]"
    WScript.Echo "Creates a macro-enabled workbook, imports src\vba\modules\ReportAutomation.bas, and saves the workbook."
    WScript.Echo "Excel must be installed and Trust access to the VBA project object model must be enabled."
End Sub
