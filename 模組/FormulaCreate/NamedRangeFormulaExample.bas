Attribute VB_Name = "NamedRangeFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: NamedRangeFormulaExample
'功能說明: 示範如何利用 VBA 建立具名範圍並將其應用於公式中
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestNamedRangeFormula()
    Call CreateNamedRangeFormulas
End Sub

' 建立具名範圍並插入公式
Sub CreateNamedRangeFormulas()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "具名範圍公式")

    ' 填入資料
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售額"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "一月"
    ws.Range("A3").Value = "二月"
    ws.Range("A4").Value = "三月"
    ws.Range("A5").Value = "四月"
    ws.Range("A6").Value = "五月"

    ws.Range("B2").Value = 12000
    ws.Range("B3").Value = 15000
    ws.Range("B4").Value = 9800
    ws.Range("B5").Value = 17500
    ws.Range("B6").Value = 13200

    ' 定義具名範圍
    On Error Resume Next
    ThisWorkbook.Names("月銷售額").Delete
    On Error GoTo ErrorHandler
    ThisWorkbook.Names.Add Name:="月銷售額", RefersTo:=ws.Range("B2:B6")

    ' 填入統計欄位
    ws.Range("D1").Value = "統計項目"
    ws.Range("E1").Value = "結果"
    ws.Range("D1:E1").Font.Bold = True

    ws.Range("D2").Value = "合計"
    ws.Range("E2").Formula = "=SUM(月銷售額)"

    ws.Range("D3").Value = "平均"
    ws.Range("E3").Formula = "=AVERAGE(月銷售額)"

    ws.Range("D4").Value = "最大值"
    ws.Range("E4").Formula = "=MAX(月銷售額)"

    ws.Range("D5").Value = "最小值"
    ws.Range("E5").Formula = "=MIN(月銷售額)"

    ws.Range("D6").Value = "數量"
    ws.Range("E6").Formula = "=COUNT(月銷售額)"

    ws.Columns("A:E").AutoFit

    MsgBox "具名範圍公式已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立具名範圍公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
