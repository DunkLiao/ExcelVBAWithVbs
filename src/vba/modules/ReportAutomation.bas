Attribute VB_Name = "ReportAutomation"
Option Explicit

Public Sub LoadCsvAndGenerateReport(ByVal csvPath As String)
    Dim hostWorkbook As Workbook
    Dim inputSheet As Worksheet
    Dim reportSheet As Worksheet

    If Len(Dir$(csvPath)) = 0 Then
        Err.Raise vbObjectError + 1000, "ReportAutomation.LoadCsvAndGenerateReport", "CSV file not found: " & csvPath
    End If

    Set hostWorkbook = ThisWorkbook
    Set inputSheet = EnsureWorksheet(hostWorkbook, "InputData")
    Set reportSheet = EnsureWorksheet(hostWorkbook, "Report")

    On Error GoTo RestoreScreenUpdating
    Application.ScreenUpdating = False

    CopyCsvIntoSheet csvPath, inputSheet
    GenerateTeamSummary inputSheet, reportSheet, csvPath

RestoreScreenUpdating:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Sub GenerateSampleReport()
    GenerateTeamSummary EnsureWorksheet(ThisWorkbook, "InputData"), EnsureWorksheet(ThisWorkbook, "Report"), "Manual workbook data"
End Sub

Private Sub CopyCsvIntoSheet(ByVal csvPath As String, ByVal targetSheet As Worksheet)
    Dim csvWorkbook As Workbook
    Dim csvSheet As Worksheet

    Set csvWorkbook = Application.Workbooks.Open(Filename:=csvPath, Local:=True)
    Set csvSheet = csvWorkbook.Worksheets(1)

    targetSheet.Cells.Clear
    csvSheet.UsedRange.Copy Destination:=targetSheet.Range("A1")

    csvWorkbook.Close SaveChanges:=False
End Sub

Private Sub GenerateTeamSummary(ByVal inputSheet As Worksheet, ByVal reportSheet As Worksheet, ByVal sourceLabel As String)
    Dim headerRow As Long
    Dim lastRow As Long
    Dim teamColumn As Long
    Dim amountColumn As Long
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim teamName As String
    Dim amountValue As Double
    Dim totalsByTeam As Object
    Dim teamKeys As Variant
    Dim grandTotal As Double

    headerRow = 1
    teamColumn = FindColumnByHeader(inputSheet, headerRow, "Team")
    amountColumn = FindColumnByHeader(inputSheet, headerRow, "Amount")

    If teamColumn = 0 Then
        Err.Raise vbObjectError + 1001, "ReportAutomation.GenerateTeamSummary", "Missing required header: Team"
    End If

    If amountColumn = 0 Then
        Err.Raise vbObjectError + 1002, "ReportAutomation.GenerateTeamSummary", "Missing required header: Amount"
    End If

    lastRow = inputSheet.Cells(inputSheet.Rows.Count, teamColumn).End(xlUp).Row
    If inputSheet.Cells(inputSheet.Rows.Count, amountColumn).End(xlUp).Row > lastRow Then
        lastRow = inputSheet.Cells(inputSheet.Rows.Count, amountColumn).End(xlUp).Row
    End If

    Set totalsByTeam = CreateObject("Scripting.Dictionary")

    For rowIndex = headerRow + 1 To lastRow
        teamName = Trim$(CStr(inputSheet.Cells(rowIndex, teamColumn).Value))
        If Len(teamName) > 0 Then
            If IsNumeric(inputSheet.Cells(rowIndex, amountColumn).Value) Then
                amountValue = CDbl(inputSheet.Cells(rowIndex, amountColumn).Value)
                If totalsByTeam.Exists(teamName) Then
                    totalsByTeam(teamName) = totalsByTeam(teamName) + amountValue
                Else
                    totalsByTeam.Add teamName, amountValue
                End If
                grandTotal = grandTotal + amountValue
            End If
        End If
    Next rowIndex

    reportSheet.Cells.Clear
    reportSheet.Range("A1").Value = "Team Sales Summary"
    reportSheet.Range("A2").Value = "Source"
    reportSheet.Range("B2").Value = sourceLabel
    reportSheet.Range("A3").Value = "Generated At"
    reportSheet.Range("B3").Value = Now
    reportSheet.Range("A5").Value = "Team"
    reportSheet.Range("B5").Value = "Total Amount"

    outputRow = 6
    If totalsByTeam.Count = 0 Then
        reportSheet.Cells(outputRow, 1).Value = "No valid rows found"
        outputRow = outputRow + 1
    Else
        teamKeys = totalsByTeam.Keys
        SortTextArray teamKeys

        For rowIndex = LBound(teamKeys) To UBound(teamKeys)
            reportSheet.Cells(outputRow, 1).Value = teamKeys(rowIndex)
            reportSheet.Cells(outputRow, 2).Value = totalsByTeam(teamKeys(rowIndex))
            outputRow = outputRow + 1
        Next rowIndex
    End If

    reportSheet.Cells(outputRow, 1).Value = "Grand Total"
    reportSheet.Cells(outputRow, 2).Value = grandTotal

    reportSheet.Columns("A:B").AutoFit
    If outputRow >= 6 Then
        reportSheet.Range("B6:B" & outputRow).NumberFormat = "#,##0.00"
    End If
End Sub

Private Function EnsureWorksheet(ByVal hostWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In hostWorkbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set EnsureWorksheet = sheet
            Exit Function
        End If
    Next sheet

    Set EnsureWorksheet = hostWorkbook.Worksheets.Add(After:=hostWorkbook.Worksheets(hostWorkbook.Worksheets.Count))
    EnsureWorksheet.Name = sheetName
End Function

Private Function FindColumnByHeader(ByVal sourceSheet As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastColumn As Long
    Dim columnIndex As Long

    lastColumn = sourceSheet.Cells(headerRow, sourceSheet.Columns.Count).End(xlToLeft).Column
    For columnIndex = 1 To lastColumn
        If StrComp(Trim$(CStr(sourceSheet.Cells(headerRow, columnIndex).Value)), headerName, vbTextCompare) = 0 Then
            FindColumnByHeader = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

Private Sub SortTextArray(ByRef values As Variant)
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim tempValue As String

    If Not IsArray(values) Then
        Exit Sub
    End If

    For leftIndex = LBound(values) To UBound(values) - 1
        For rightIndex = leftIndex + 1 To UBound(values)
            If StrComp(CStr(values(leftIndex)), CStr(values(rightIndex)), vbTextCompare) > 0 Then
                tempValue = CStr(values(leftIndex))
                values(leftIndex) = values(rightIndex)
                values(rightIndex) = tempValue
            End If
        Next rightIndex
    Next leftIndex
End Sub
