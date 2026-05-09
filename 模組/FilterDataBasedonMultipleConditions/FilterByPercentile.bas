Attribute VB_Name = "FilterByPercentile"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByPercentile
' 功能說明: 計算指定欄位的百分位數門檻值，再用 AutoFilter 篩選超過門檻的記錄
'           示範自訂百分位數動態篩選（如：業績高於第75百分位）
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：篩選業績高於第 75 百分位的業務員
Public Sub FilterAbovePercentileExample()
    On Error GoTo ErrHandler

    Dim ws          As Worksheet
    Dim percentile  As Double
    Dim threshold   As Double

    Set ws = GetOrCreateWsPct(ThisWorkbook, "百分位篩選範例")
    Call FillPerformanceData(ws)

    percentile = 75   ' 第 75 百分位
    threshold = CalcPercentile(ws, 3, percentile)

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1").CurrentRegion.AutoFilter Field:=3, Criteria1:=">" & threshold
    ws.Columns("A:D").AutoFit

    MsgBox "已篩選業績高於第 " & percentile & " 百分位（門檻：" & _
           Format(threshold, "#,##0") & " 元）的業務員。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除百分位篩選
Public Sub ClearPercentileFilter()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    MsgBox "已清除篩選。", vbInformation, "完成"
End Sub

' 計算指定欄的第 N 百分位值（使用 Percentile 工作表函數）
Private Function CalcPercentile(ByVal ws As Worksheet, _
                                 ByVal col As Integer, _
                                 ByVal pct As Double) As Double
    Dim lastRow As Long
    Dim rng     As Range

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))

    CalcPercentile = Application.WorksheetFunction.Percentile(rng, pct / 100)
End Function

' 填入業績測試資料
Private Sub FillPerformanceData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("業務員", "部門", "季業績", "客戶數")
    ws.Range("A2:D2").Value = Array("王大明", "北區業務", 820000, 12)
    ws.Range("A3:D3").Value = Array("陳美珠", "南區業務", 1350000, 18)
    ws.Range("A4:D4").Value = Array("林志偉", "中區業務", 560000, 8)
    ws.Range("A5:D5").Value = Array("黃麗娟", "北區業務", 940000, 14)
    ws.Range("A6:D6").Value = Array("張建國", "東區業務", 430000, 6)
    ws.Range("A7:D7").Value = Array("蔡雅芳", "南區業務", 1120000, 16)
    ws.Range("A8:D8").Value = Array("吳俊仁", "中區業務", 670000, 10)
    ws.Range("A9:D9").Value = Array("李淑君", "東區業務", 780000, 11)
    ws.Range("C2:C9").NumberFormat = "#,##0"
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsPct(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsPct = ws
End Function