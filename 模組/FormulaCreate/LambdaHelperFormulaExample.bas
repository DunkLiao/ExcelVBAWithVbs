Attribute VB_Name = "LambdaHelperFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: LambdaHelperFormulaExample
'功能說明: 示範 LAMBDA 輔助函數（BYROW、BYCOL、MAP、REDUCE、SCAN）的使用方式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestLambdaHelperFormula()
    Call CreateLambdaHelperExample("LAMBDA輔助函數範例")
End Sub

Sub CreateLambdaHelperExample(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 填入範例資料
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "Q1銷售"
    ws.Range("C1").Value = "Q2銷售"
    ws.Range("D1").Value = "Q3銷售"
    ws.Range("E1").Value = "Q4銷售"
    ws.Range("A1:E1").Font.Bold = True

    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 1200
    ws.Range("C2").Value = 1500
    ws.Range("D2").Value = 1300
    ws.Range("E2").Value = 1800

    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 800
    ws.Range("C3").Value = 950
    ws.Range("D3").Value = 1100
    ws.Range("E3").Value = 750

    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 2000
    ws.Range("C4").Value = 1800
    ws.Range("D4").Value = 2200
    ws.Range("E4").Value = 2100

    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 650
    ws.Range("C5").Value = 700
    ws.Range("D5").Value = 550
    ws.Range("E5").Value = 900

    ' BYROW 公式：計算每列總和
    ws.Range("F1").Value = "BYROW年度合計"
    ws.Range("F1").Font.Bold = True
    ws.Range("F2").Formula = _
        "=BYROW(B2:E5,LAMBDA(r,SUM(r)))"

    ' BYCOL 公式：計算每欄最大值
    ws.Range("B6").Formula = _
        "=BYCOL(B2:E5,LAMBDA(c,MAX(c)))"

    ' MAP 公式：檢查每個值是否大於 1000
    ws.Range("G1").Value = "MAP是否>1000"
    ws.Range("G1").Font.Bold = True
    ws.Range("G2").Formula = _
        "=MAP(B2:E5,LAMBDA(x,IF(x>1000,""是"",""否"")))"

    ws.Columns("A:G").AutoFit

    MsgBox "LAMBDA 輔助函數範例已建立完成！", vbInformation, "完成"
End Sub
