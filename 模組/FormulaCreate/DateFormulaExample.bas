Attribute VB_Name = "DateFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: DateFormulaExample
'功能描述: 在 Excel 中示範日期時間公式的使用範例
'          包含 TODAY、NOW、YEAR、MONTH、DAY、DATEDIF、NETWORKDAYS 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestDateFormula()
    Call CreateDateFormulaExample("日期公式範例")
End Sub

Sub CreateDateFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "公式說明"
    ws.Range("B1").Value = "結果"
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2").Value = "TODAY() 今天日期"
    ws.Range("B2").Formula = "=TODAY()"
    ws.Range("B2").NumberFormat = "yyyy/mm/dd"
    ws.Range("B2").Interior.Color = RGB(198, 239, 206)
    ws.Range("A3").Value = "NOW() 現在時間"
    ws.Range("B3").Formula = "=NOW()"
    ws.Range("B3").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Range("B3").Interior.Color = RGB(198, 239, 206)
    ws.Range("A5").Value = "起始日期"
    ws.Range("B5").Value = DateSerial(2020, 1, 15)
    ws.Range("B5").NumberFormat = "yyyy/mm/dd"
    ws.Range("A6").Value = "YEAR(起始日期)"
    ws.Range("B6").Formula = "=YEAR(B5)"
    ws.Range("B6").Interior.Color = RGB(255, 235, 156)
    ws.Range("A7").Value = "MONTH(起始日期)"
    ws.Range("B7").Formula = "=MONTH(B5)"
    ws.Range("B7").Interior.Color = RGB(255, 235, 156)
    ws.Range("A8").Value = "DAY(起始日期)"
    ws.Range("B8").Formula = "=DAY(B5)"
    ws.Range("B8").Interior.Color = RGB(255, 235, 156)
    ws.Range("A9").Value = "WEEKDAY(起始日期)"
    ws.Range("B9").Formula = "=WEEKDAY(B5,2)"
    ws.Range("B9").Interior.Color = RGB(255, 235, 156)
    ws.Range("A11").Value = "結束日期"
    ws.Range("B11").Value = DateSerial(2026, 5, 7)
    ws.Range("B11").NumberFormat = "yyyy/mm/dd"
    ws.Range("A12").Value = "DATEDIF 年差"
    ws.Range("B12").Formula = "=DATEDIF(B5,B11,""Y"")"
    ws.Range("B12").Interior.Color = RGB(198, 239, 206)
    ws.Range("A13").Value = "DATEDIF 月差"
    ws.Range("B13").Formula = "=DATEDIF(B5,B11,""M"")"
    ws.Range("B13").Interior.Color = RGB(198, 239, 206)
    ws.Range("A14").Value = "DATEDIF 天差"
    ws.Range("B14").Formula = "=DATEDIF(B5,B11,""D"")"
    ws.Range("B14").Interior.Color = RGB(198, 239, 206)
    ws.Range("A16").Value = "NETWORKDAYS 工作天數"
    ws.Range("B16").Formula = "=NETWORKDAYS(B5,B11)"
    ws.Range("B16").Interior.Color = RGB(198, 239, 206)
    ws.Range("A18").Value = "EOMONTH 當月最後一天"
    ws.Range("B18").Formula = "=EOMONTH(TODAY(),0)"
    ws.Range("B18").NumberFormat = "yyyy/mm/dd"
    ws.Range("B18").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:B").AutoFit
    MsgBox "日期公式範例已建立完成！", vbInformation, "完成"
End Sub
