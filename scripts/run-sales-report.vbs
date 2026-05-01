Option Explicit

Dim args
Dim scriptFolder
Dim repoRoot
Dim csvPath
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

If args.Count = 0 Then
    ShowUsage
    WScript.Quit 1
End If

csvPath = ToAbsolutePath(args(0), repoRoot)

If args.Count >= 2 Then
    workbookPath = ToAbsolutePath(args(1), repoRoot)
Else
    workbookPath = repoRoot & "\workbooks\ReportAutomationTemplate.xlsm"
End If

If Not FileExists(csvPath) Then
    Err.Raise vbObjectError + 2100, "run-sales-report.vbs", "CSV file not found: " & csvPath
End If

If Not FileExists(workbookPath) Then
    Err.Raise vbObjectError + 2101, "run-sales-report.vbs", "Workbook not found: " & workbookPath & ". Run scripts\bootstrap-report-workbook.vbs first."
End If

Set excelApp = CreateExcelApplication()
excelApp.Visible = False
excelApp.DisplayAlerts = False

Set workbook = excelApp.Workbooks.Open(workbookPath)
excelApp.Run "'" & workbook.Name & "'!ReportAutomation.LoadCsvAndGenerateReport", csvPath
workbook.Save
workbook.Close False
excelApp.Quit

WScript.Echo "Report refreshed in workbook: " & workbookPath

Function CreateExcelApplication()
    On Error Resume Next
    Set CreateExcelApplication = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Dim createExcelMessage
        createExcelMessage = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 2102, "run-sales-report.vbs", "Unable to start Excel. " & createExcelMessage
    End If
    On Error GoTo 0
End Function

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
    WScript.Echo "Usage: cscript //nologo scripts\run-sales-report.vbs <csvPath> [workbookPath]"
    WScript.Echo "Loads the CSV into the workbook, runs ReportAutomation.LoadCsvAndGenerateReport, and saves the workbook."
End Sub
