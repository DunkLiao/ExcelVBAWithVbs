Attribute VB_Name = "CompareByColumnOrder"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByColumnOrder
'功能說明: 依欄位順序比對 Sheet1 與 Sheet2 並輸出差異報告
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunCompareByColumnOrder()
    On Error GoTo ErrorHandler

    Dim wsLeft As Worksheet
    Dim wsRight As Worksheet
    Dim wsReport As Worksheet
    Dim maxRow As Long
    Dim maxCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim reportRow As Long
    Dim leftValue As String
    Dim rightValue As String

    Set wsLeft = GetWorksheetByName("Sheet1")
    Set wsRight = GetWorksheetByName("Sheet2")

    If wsLeft Is Nothing Or wsRight Is Nothing Then
        MsgBox "請先準備 Sheet1 與 Sheet2 工作表。", vbExclamation, "提示"
        Exit Sub
    End If

    Set wsReport = GetOrCreateCompareReportSheet("比對報告")
    wsReport.Cells.Clear
    wsReport.Range("A1:F1").Value = Array("列號", "欄號", "座標", "Sheet1", "Sheet2", "說明")
    reportRow = 2

    maxRow = GetCompareMax(GetLastCompareRow(wsLeft), GetLastCompareRow(wsRight))
    maxCol = GetCompareMax(GetLastCompareCol(wsLeft), GetLastCompareCol(wsRight))

    For rowIndex = 1 To maxRow
        For colIndex = 1 To maxCol
            leftValue = GetComparableCellValue(wsLeft.Cells(rowIndex, colIndex))
            rightValue = GetComparableCellValue(wsRight.Cells(rowIndex, colIndex))

            If leftValue <> rightValue Then
                wsLeft.Cells(rowIndex, colIndex).Interior.Color = RGB(255, 199, 206)
                wsRight.Cells(rowIndex, colIndex).Interior.Color = RGB(255, 199, 206)

                wsReport.Cells(reportRow, 1).Value = rowIndex
                wsReport.Cells(reportRow, 2).Value = colIndex
                wsReport.Cells(reportRow, 3).Value = ColumnLetter(colIndex) & rowIndex
                wsReport.Cells(reportRow, 4).Value = leftValue
                wsReport.Cells(reportRow, 5).Value = rightValue
                wsReport.Cells(reportRow, 6).Value = "兩表資料不同"
                reportRow = reportRow + 1
            End If
        Next colIndex
    Next rowIndex

    wsReport.Columns.AutoFit

    If reportRow = 2 Then
        MsgBox "Sheet1 與 Sheet2 沒有差異。", vbInformation, "完成"
    Else
        MsgBox "已完成資料比對，請查看比對報告工作表。", vbInformation, "完成"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "比對工作表時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateCompareReportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateCompareReportSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateCompareReportSheet Is Nothing Then
        Set GetOrCreateCompareReportSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateCompareReportSheet.Name = sheetName
    End If
End Function

Private Function GetLastCompareRow(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetLastCompareRow = 1
    Else
        GetLastCompareRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
End Function

Private Function GetLastCompareCol(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetLastCompareCol = 1
    Else
        GetLastCompareCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
End Function

Private Function GetCompareMax(ByVal value1 As Long, ByVal value2 As Long) As Long
    If value1 >= value2 Then
        GetCompareMax = value1
    Else
        GetCompareMax = value2
    End If
End Function

Private Function GetComparableCellValue(ByVal targetCell As Range) As String
    If IsError(targetCell.Value) Then
        GetComparableCellValue = "#錯誤"
    ElseIf IsEmpty(targetCell.Value) Then
        GetComparableCellValue = ""
    Else
        GetComparableCellValue = CStr(targetCell.Value2)
    End If
End Function

Private Function ColumnLetter(ByVal columnNumber As Long) As String
    Dim resultText As String
    Dim remainder As Long

    Do While columnNumber > 0
        remainder = (columnNumber - 1) Mod 26
        resultText = Chr$(65 + remainder) & resultText
        columnNumber = (columnNumber - remainder - 1) \ 26
    Loop

    ColumnLetter = resultText
End Function
