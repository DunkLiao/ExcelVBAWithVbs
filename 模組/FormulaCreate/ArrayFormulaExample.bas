Attribute VB_Name = "ArrayFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: ArrayFormulaExample
'功能說明: 示範在Excel中透過VBA輸入陣列公式的範例程式
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestArrayFormula()
    Call CreateArrayFormulaExample
End Sub

' 建立陣列公式範例
Sub CreateArrayFormulaExample()
    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = "陣列公式範例"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    Call FillProductData(ws)
    Call EnterArrayFormulas(ws)

    ws.Activate
    MsgBox "陣列公式範例已建立完成！", vbInformation, "完成"
End Sub

' 填入產品銷售資料
Private Sub FillProductData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "單價"
    ws.Range("C1").Value = "數量"
    ws.Range("D1").Value = "類別"
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = 50
    ws.Range("D2").Value = "電子"
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 200
    ws.Range("C3").Value = 30
    ws.Range("D3").Value = "家電"
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 150
    ws.Range("C4").Value = 80
    ws.Range("D4").Value = "電子"
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 300
    ws.Range("C5").Value = 20
    ws.Range("D5").Value = "家電"
    ws.Range("A6").Value = "產品E"
    ws.Range("B6").Value = 80
    ws.Range("C6").Value = 100
    ws.Range("D6").Value = "電子"
    ws.Columns("A:D").AutoFit
End Sub

' 輸入陣列公式
Private Sub EnterArrayFormulas(ByVal ws As Worksheet)
    ws.Range("F1").Value = "公式說明"
    ws.Range("G1").Value = "結果"

    ' 陣列公式：所有產品銷售總額
    ws.Range("F2").Value = "銷售總額(陣列公式)"
    ws.Range("G2").FormulaArray = "=SUM(B2:B6*C2:C6)"

    ' 陣列公式：電子類別銷售額
    ws.Range("F3").Value = "電子類銷售額(陣列公式)"
    ws.Range("G3").FormulaArray = "=SUM((D2:D6=""電子"")*B2:B6*C2:C6)"

    ' 陣列公式：單價超過100的產品數
    ws.Range("F4").Value = "單價>100的品項數"
    ws.Range("G4").FormulaArray = "=SUM((B2:B6>100)*1)"

    ' 陣列公式：最大銷售額
    ws.Range("F5").Value = "最高單品銷售額"
    ws.Range("G5").FormulaArray = "=MAX(B2:B6*C2:C6)"

    ws.Columns("F:G").AutoFit
End Sub
