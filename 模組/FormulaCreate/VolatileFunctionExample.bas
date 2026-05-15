Attribute VB_Name = "VolatileFunctionExample"
Option Explicit
'*************************************************************************************
'模組名稱: VolatileFunctionExample
'功能說明: 易變函數範例，展示多種 Excel 易變函數的用法
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunVolatileFunctionExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateVolatileSheet("易變函數範例")
    ws.Cells.Clear

    Call FillVolatileSupportData(ws)
    Call FillVolatileFormulaTable(ws)

    ws.Columns("A:D").AutoFit
    MsgBox "易變函數範例已建立完成。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立易變函數範例時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillVolatileFormulaTable(ByVal ws As Worksheet)
    ws.Range("A1:B1").Value = Array("說明", "公式")

    ws.Range("A2").Value = "NOW 目前日期時間"
    ws.Range("B2").Formula = "=NOW()"

    ws.Range("A3").Value = "TODAY 今日日期"
    ws.Range("B3").Formula = "=TODAY()"

    ws.Range("A4").Value = "RAND 0 到 1 亂數"
    ws.Range("B4").Formula = "=RAND()"

    ws.Range("A5").Value = "RANDBETWEEN 1 到 100 整數"
    ws.Range("B5").Formula = "=RANDBETWEEN(1,100)"

    ws.Range("A6").Value = "OFFSET 依列數抓取參考值"
    ws.Range("B6").Formula = "=OFFSET($D$2,ROW()-2,0)"

    ws.Range("A7").Value = "INDIRECT 動態參照 D 欄"
    ws.Range("B7").Formula = "=INDIRECT(""D"" & ROW())"

    ws.Range("A8").Value = "ROW 目前列號"
    ws.Range("B8").Formula = "=ROW()"

    ws.Range("A9").Value = "COLUMN 目前欄號"
    ws.Range("B9").Formula = "=COLUMN()"

    ws.Range("B2:B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B4:B5").NumberFormat = "0.00"
End Sub

Private Sub FillVolatileSupportData(ByVal ws As Worksheet)
    ws.Range("D1").Value = "參考值"
    ws.Range("D2").Value = "北區"
    ws.Range("D3").Value = "中區"
    ws.Range("D4").Value = "南區"
    ws.Range("D5").Value = "東區"
    ws.Range("D6").Value = "離島"
    ws.Range("D7").Value = "外銷"
    ws.Range("D8").Value = "內銷"
    ws.Range("D9").Value = "代理"
End Sub

Private Function GetOrCreateVolatileSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateVolatileSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateVolatileSheet Is Nothing Then
        Set GetOrCreateVolatileSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateVolatileSheet.Name = sheetName
    End If
End Function
