Attribute VB_Name = "StatFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: StatFormulaExample
'功能描述: 在 Excel 中示範統計公式的使用範例
'          包含 AVERAGE、MAX、MIN、COUNT、COUNTIF、COUNTIFS、LARGE、SMALL、MEDIAN、STDEV 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestStatFormula()
    Call CreateStatFormulaExample("統計公式範例")
End Sub

Sub CreateStatFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Call FillStatData(ws)
    ws.Range("E1").Value = "公式說明"
    ws.Range("F1").Value = "結果"
    ws.Range("E1:F1").Font.Bold = True
    ws.Range("E2").Value = "AVERAGE 平均"
    ws.Range("F2").Formula = "=AVERAGE(C2:C11)"
    ws.Range("F2").Interior.Color = RGB(198, 239, 206)
    ws.Range("E3").Value = "MAX 最大值"
    ws.Range("F3").Formula = "=MAX(C2:C11)"
    ws.Range("F3").Interior.Color = RGB(198, 239, 206)
    ws.Range("E4").Value = "MIN 最小值"
    ws.Range("F4").Formula = "=MIN(C2:C11)"
    ws.Range("F4").Interior.Color = RGB(198, 239, 206)
    ws.Range("E5").Value = "COUNT 數值計數"
    ws.Range("F5").Formula = "=COUNT(C2:C11)"
    ws.Range("F5").Interior.Color = RGB(198, 239, 206)
    ws.Range("E6").Value = "COUNTA 非空計數"
    ws.Range("F6").Formula = "=COUNTA(A2:A11)"
    ws.Range("F6").Interior.Color = RGB(198, 239, 206)
    ws.Range("E7").Value = "COUNTIF 業務部門人數"
    ws.Range("F7").Formula = "=COUNTIF(B2:B11,""業務"")"
    ws.Range("F7").Interior.Color = RGB(255, 235, 156)
    ws.Range("E8").Value = "COUNTIFS 業務且>500"
    ws.Range("F8").Formula = "=COUNTIFS(B2:B11,""業務"",C2:C11,"">500"")"
    ws.Range("F8").Interior.Color = RGB(255, 235, 156)
    ws.Range("E9").Value = "LARGE 第1大"
    ws.Range("F9").Formula = "=LARGE(C2:C11,1)"
    ws.Range("F9").Interior.Color = RGB(255, 199, 206)
    ws.Range("E10").Value = "SMALL 第1小"
    ws.Range("F10").Formula = "=SMALL(C2:C11,1)"
    ws.Range("F10").Interior.Color = RGB(255, 199, 206)
    ws.Range("E11").Value = "MEDIAN 中位數"
    ws.Range("F11").Formula = "=MEDIAN(C2:C11)"
    ws.Range("F11").Interior.Color = RGB(198, 239, 206)
    ws.Range("E12").Value = "STDEV 標準差"
    ws.Range("F12").Formula = "=STDEV(C2:C11)"
    ws.Range("F12").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:F").AutoFit
    MsgBox "統計公式範例已建立完成！", vbInformation, "完成"
End Sub

Private Sub FillStatData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "業績"
    Dim data(9, 2) As Variant
    data(0, 0) = "張三" : data(0, 1) = "業務" : data(0, 2) = 800
    data(1, 0) = "李四" : data(1, 1) = "行政" : data(1, 2) = 300
    data(2, 0) = "王五" : data(2, 1) = "業務" : data(2, 2) = 650
    data(3, 0) = "趙六" : data(3, 1) = "研發" : data(3, 2) = 450
    data(4, 0) = "陳七" : data(4, 1) = "業務" : data(4, 2) = 920
    data(5, 0) = "林八" : data(5, 1) = "行政" : data(5, 2) = 280
    data(6, 0) = "吳九" : data(6, 1) = "研發" : data(6, 2) = 510
    data(7, 0) = "黃十" : data(7, 1) = "業務" : data(7, 2) = 380
    data(8, 0) = "周一" : data(8, 1) = "業務" : data(8, 2) = 740
    data(9, 0) = "馮二" : data(9, 1) = "研發" : data(9, 2) = 600
    Dim i As Integer
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = data(i, 0)
        ws.Cells(i + 2, 2).Value = data(i, 1)
        ws.Cells(i + 2, 3).Value = data(i, 2)
    Next i
    ws.Range("A1:C1").Font.Bold = True
End Sub
