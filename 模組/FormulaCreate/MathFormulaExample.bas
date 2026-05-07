Attribute VB_Name = "MathFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: MathFormulaExample
'功能描述: 在 Excel 中示範數學運算公式的使用範例
'          包含 ROUND、ROUNDUP、ROUNDDOWN、ABS、MOD、INT、CEILING、FLOOR、POWER、SQRT 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestMathFormula()
    Call CreateMathFormulaExample("數學公式範例")
End Sub

Sub CreateMathFormulaExample(ByVal sheetName As String)
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
    ws.Range("B1").Value = "輸入值"
    ws.Range("C1").Value = "結果"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A2").Value = "ROUND(四捨五入到2位)"
    ws.Range("B2").Value = 3.14159
    ws.Range("C2").Formula = "=ROUND(B2,2)"
    ws.Range("C2").Interior.Color = RGB(198, 239, 206)
    ws.Range("A3").Value = "ROUNDUP(無條件進位到2位)"
    ws.Range("B3").Value = 3.14159
    ws.Range("C3").Formula = "=ROUNDUP(B3,2)"
    ws.Range("C3").Interior.Color = RGB(198, 239, 206)
    ws.Range("A4").Value = "ROUNDDOWN(無條件捨去到2位)"
    ws.Range("B4").Value = 3.14159
    ws.Range("C4").Formula = "=ROUNDDOWN(B4,2)"
    ws.Range("C4").Interior.Color = RGB(198, 239, 206)
    ws.Range("A5").Value = "ABS(絕對值)"
    ws.Range("B5").Value = -25.8
    ws.Range("C5").Formula = "=ABS(B5)"
    ws.Range("C5").Interior.Color = RGB(255, 235, 156)
    ws.Range("A6").Value = "MOD(餘數: 17除以5)"
    ws.Range("B6").Value = 17
    ws.Range("C6").Formula = "=MOD(B6,5)"
    ws.Range("C6").Interior.Color = RGB(255, 235, 156)
    ws.Range("A7").Value = "INT(取整數)"
    ws.Range("B7").Value = 9.87
    ws.Range("C7").Formula = "=INT(B7)"
    ws.Range("C7").Interior.Color = RGB(255, 235, 156)
    ws.Range("A8").Value = "CEILING(進位到5的倍數)"
    ws.Range("B8").Value = 23
    ws.Range("C8").Formula = "=CEILING(B8,5)"
    ws.Range("C8").Interior.Color = RGB(255, 199, 206)
    ws.Range("A9").Value = "FLOOR(捨去到5的倍數)"
    ws.Range("B9").Value = 23
    ws.Range("C9").Formula = "=FLOOR(B9,5)"
    ws.Range("C9").Interior.Color = RGB(255, 199, 206)
    ws.Range("A10").Value = "POWER(2的10次方)"
    ws.Range("B10").Value = 2
    ws.Range("C10").Formula = "=POWER(B10,10)"
    ws.Range("C10").Interior.Color = RGB(198, 239, 206)
    ws.Range("A11").Value = "SQRT(平方根)"
    ws.Range("B11").Value = 144
    ws.Range("C11").Formula = "=SQRT(B11)"
    ws.Range("C11").Interior.Color = RGB(198, 239, 206)
    ws.Range("A12").Value = "LOG(以10為底的對數)"
    ws.Range("B12").Value = 1000
    ws.Range("C12").Formula = "=LOG(B12,10)"
    ws.Range("C12").Interior.Color = RGB(198, 239, 206)
    ws.Range("A13").Value = "RAND(0~1隨機數)"
    ws.Range("C13").Formula = "=RAND()"
    ws.Range("C13").Interior.Color = RGB(198, 239, 206)
    ws.Columns("A:C").AutoFit
    MsgBox "數學公式範例已建立完成！", vbInformation, "完成"
End Sub
