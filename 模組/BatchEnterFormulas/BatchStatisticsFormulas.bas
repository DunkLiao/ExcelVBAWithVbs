Attribute VB_Name = "BatchStatisticsFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchStatisticsFormulas
'功能說明: 批次填入統計公式，包含 STDEV、VAR、MEDIAN、MODE、QUARTILE 等描述性統計
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchStatisticsFormulas()
    Call CreateStatisticsFormulaExample
End Sub

' 建立統計公式批次填入示範
Sub CreateStatisticsFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateStatSheet(ThisWorkbook, "統計公式示範")
    Call FillProductSalesData(ws)
    Call BatchEnterDescriptiveStats(ws)

    ws.Columns("A:H").AutoFit
    ws.Activate
    MsgBox "統計公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入描述性統計公式（每欄資料分別計算）
Private Sub BatchEnterDescriptiveStats(ByVal ws As Worksheet)
    Dim col As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' 統計標籤行（第14列之後，假設資料12筆）
    Dim statsRow As Long
    statsRow = lastRow + 2

    ws.Cells(statsRow, 1).Value = "統計項目"
    ws.Cells(statsRow, 1).Font.Bold = True

    ' 標籤清單
    Dim labels(1 To 9) As String
    labels(1) = "平均值 (AVERAGE)"
    labels(2) = "中位數 (MEDIAN)"
    labels(3) = "眾數 (MODE)"
    labels(4) = "標準差 (STDEV)"
    labels(5) = "變異數 (VAR)"
    labels(6) = "第1四分位 (Q1)"
    labels(7) = "第3四分位 (Q3)"
    labels(8) = "四分位距 (IQR)"
    labels(9) = "變異係數 (CV%)"

    Dim k As Integer
    For k = 1 To 9
        ws.Cells(statsRow + k, 1).Value = labels(k)
        ws.Cells(statsRow + k, 1).Font.Bold = True
    Next k

    ' 對每個商品欄（B~D）批次填入統計公式
    For col = 2 To 4
        Dim rngAddr As String
        rngAddr = ws.Cells(2, col).Address(False, False) & ":" & ws.Cells(lastRow, col).Address(False, False)

        ws.Cells(statsRow + 1, col).Formula = "=AVERAGE(" & rngAddr & ")"
        ws.Cells(statsRow + 2, col).Formula = "=MEDIAN(" & rngAddr & ")"
        ws.Cells(statsRow + 3, col).Formula = "=MODE(" & rngAddr & ")"
        ws.Cells(statsRow + 4, col).Formula = "=STDEV(" & rngAddr & ")"
        ws.Cells(statsRow + 5, col).Formula = "=VAR(" & rngAddr & ")"
        ws.Cells(statsRow + 6, col).Formula = "=QUARTILE(" & rngAddr & ",1)"
        ws.Cells(statsRow + 7, col).Formula = "=QUARTILE(" & rngAddr & ",3)"
        ' IQR = Q3 - Q1
        ws.Cells(statsRow + 8, col).Formula = "=" & _
            ws.Cells(statsRow + 7, col).Address(False, False) & "-" & _
            ws.Cells(statsRow + 6, col).Address(False, False)
        ' CV = 標準差 / 平均值
        ws.Cells(statsRow + 9, col).Formula = "=" & _
            ws.Cells(statsRow + 4, col).Address(False, False) & "/" & _
            ws.Cells(statsRow + 1, col).Address(False, False)
        ws.Cells(statsRow + 9, col).NumberFormat = "0.00%"

        ' 格式：小數2位
        ws.Range(ws.Cells(statsRow + 1, col), ws.Cells(statsRow + 8, col)).NumberFormat = "0.00"
    Next col
End Sub

' 填入商品月銷售量資料（12個月 x 3個商品）
Private Sub FillProductSalesData(ByVal ws As Worksheet)
    Dim months(1 To 12) As String
    Dim salesA(1 To 12) As Integer
    Dim salesB(1 To 12) As Integer
    Dim salesC(1 To 12) As Integer
    Dim i As Integer

    months(1) = "1月" : months(2) = "2月" : months(3) = "3月"
    months(4) = "4月" : months(5) = "5月" : months(6) = "6月"
    months(7) = "7月" : months(8) = "8月" : months(9) = "9月"
    months(10) = "10月" : months(11) = "11月" : months(12) = "12月"

    salesA(1) = 85  : salesA(2) = 92  : salesA(3) = 78
    salesA(4) = 95  : salesA(5) = 88  : salesA(6) = 102
    salesA(7) = 115 : salesA(8) = 120 : salesA(9) = 98
    salesA(10) = 88 : salesA(11) = 105 : salesA(12) = 130

    salesB(1) = 210 : salesB(2) = 195 : salesB(3) = 230
    salesB(4) = 185 : salesB(5) = 220 : salesB(6) = 215
    salesB(7) = 198 : salesB(8) = 210 : salesB(9) = 240
    salesB(10) = 225 : salesB(11) = 195 : salesB(12) = 260

    salesC(1) = 50  : salesC(2) = 55  : salesC(3) = 48
    salesC(4) = 62  : salesC(5) = 58  : salesC(6) = 70
    salesC(7) = 45  : salesC(8) = 52  : salesC(9) = 65
    salesC(10) = 72 : salesC(11) = 68 : salesC(12) = 80

    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "商品A銷量"
    ws.Range("C1").Value = "商品B銷量"
    ws.Range("D1").Value = "商品C銷量"
    ws.Range("A1:D1").Font.Bold = True

    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = months(i)
        ws.Cells(i + 1, 2).Value = salesA(i)
        ws.Cells(i + 1, 3).Value = salesB(i)
        ws.Cells(i + 1, 4).Value = salesC(i)
    Next i
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateStatSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateStatSheet = ws
End Function