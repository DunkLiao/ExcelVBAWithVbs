Attribute VB_Name = "BatchSumFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchSumFormulas
'功能說明: 批次在多個儲存格中自動輸入加總、平均、最大最小值公式
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestBatchSumFormulas()
    Call CreateBatchFormulaExample
End Sub

' 建立批次公式輸入範例
Sub CreateBatchFormulaExample()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "批次公式範例")

    Call FillMonthlySalesData(ws)
    Call BatchEnterSumFormulas(ws)
    Call BatchEnterAverageFormulas(ws)
    Call BatchEnterMaxMinFormulas(ws)

    ws.Activate
    MsgBox "批次公式輸入完成！", vbInformation, "完成"
End Sub

' 批次輸入每月小計（SUM）公式
Private Sub BatchEnterSumFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    For i = 2 To 7
        ws.Cells(8, i).Formula = "=SUM(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(7, i).Address(False, False) & ")"
    Next i
    ws.Range("A8").Value = "月合計"
    ws.Range("A8").Font.Bold = True
    ws.Range("B8:G8").Font.Bold = True
End Sub

' 批次輸入每月平均（AVERAGE）公式
Private Sub BatchEnterAverageFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    For i = 2 To 7
        ws.Cells(9, i).Formula = "=AVERAGE(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(7, i).Address(False, False) & ")"
    Next i
    ws.Range("A9").Value = "月平均"
    ws.Range("A9").Font.Bold = True
End Sub

' 批次輸入每月最大最小值公式
Private Sub BatchEnterMaxMinFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    For i = 2 To 7
        ws.Cells(10, i).Formula = "=MAX(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(7, i).Address(False, False) & ")"
        ws.Cells(11, i).Formula = "=MIN(" & _
            ws.Cells(2, i).Address(False, False) & ":" & _
            ws.Cells(7, i).Address(False, False) & ")"
    Next i
    ws.Range("A10").Value = "月最大值"
    ws.Range("A11").Value = "月最小值"
    ws.Range("A10:A11").Font.Bold = True
End Sub

' 填入月度銷售資料（6個業務員 x 6個月）
Private Sub FillMonthlySalesData(ByVal ws As Worksheet)
    Dim salesData(1 To 6, 1 To 6) As Integer
    Dim r As Integer
    Dim c As Integer

    ws.Range("A1").Value = "業務員\月份"
    ws.Range("B1").Value = "1月"
    ws.Range("C1").Value = "2月"
    ws.Range("D1").Value = "3月"
    ws.Range("E1").Value = "4月"
    ws.Range("F1").Value = "5月"
    ws.Range("G1").Value = "6月"

    ws.Range("A2").Value = "張一"
    ws.Range("A3").Value = "李二"
    ws.Range("A4").Value = "王三"
    ws.Range("A5").Value = "趙四"
    ws.Range("A6").Value = "孫五"
    ws.Range("A7").Value = "周六"

    salesData(1, 1) = 120 : salesData(1, 2) = 135 : salesData(1, 3) = 98
    salesData(1, 4) = 145 : salesData(1, 5) = 160 : salesData(1, 6) = 130
    salesData(2, 1) = 95 : salesData(2, 2) = 110 : salesData(2, 3) = 125
    salesData(2, 4) = 100 : salesData(2, 5) = 88 : salesData(2, 6) = 115
    salesData(3, 1) = 200 : salesData(3, 2) = 185 : salesData(3, 3) = 220
    salesData(3, 4) = 195 : salesData(3, 5) = 210 : salesData(3, 6) = 230
    salesData(4, 1) = 75 : salesData(4, 2) = 80 : salesData(4, 3) = 70
    salesData(4, 4) = 90 : salesData(4, 5) = 85 : salesData(4, 6) = 78
    salesData(5, 1) = 155 : salesData(5, 2) = 160 : salesData(5, 3) = 148
    salesData(5, 4) = 170 : salesData(5, 5) = 165 : salesData(5, 6) = 175
    salesData(6, 1) = 110 : salesData(6, 2) = 105 : salesData(6, 3) = 118
    salesData(6, 4) = 122 : salesData(6, 5) = 130 : salesData(6, 6) = 115

    For r = 1 To 6
        For c = 1 To 6
            ws.Cells(r + 1, c + 1).Value = salesData(r, c)
        Next c
    Next r

    ws.Columns("A:G").AutoFit
End Sub

' 取得或建立工作表，並清除內容
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
