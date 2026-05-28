Attribute VB_Name = "RandomFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: RandomFormulaExample
'功能說明: 批次建立 RAND / RANDBETWEEN 亂數公式，示範亂數產生與固定值寫法
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestRandomFormula()
    Call CreateRandomFormulaDemo("亂數公式範例")
End Sub

Sub CreateRandomFormulaDemo(ByVal sheetName As String)
    Dim ws  As Worksheet
    Dim i   As Integer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear

    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "RAND (0~1)"
    ws.Range("C1").Value = "RANDBETWEEN (1~100)"
    ws.Range("D1").Value = "RANDBETWEEN (500~1000)"
    ws.Range("E1").Value = "固定亂數值 (不隨計算改變)"

    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With

    For i = 2 To 21
        ws.Cells(i, 1).Value = "項目" & (i - 1)
        ws.Cells(i, 2).Formula = "=RAND()"
        ws.Cells(i, 3).Formula = "=RANDBETWEEN(1,100)"
        ws.Cells(i, 4).Formula = "=RANDBETWEEN(500,1000)"
    Next i

    ' 複製 C 欄公式結果以純值貼至 E 欄（固定亂數）
    ws.Range("C2:C21").Copy
    ws.Range("E2:E21").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ws.Range("B2:B21").NumberFormat = "0.0000"
    ws.Range("C2:D21").NumberFormat = "0"
    ws.Range("E2:E21").NumberFormat = "0"
    ws.Columns("A:E").AutoFit

    MsgBox "亂數公式範例已建立完畢！共 20 筆資料。", vbInformation, "完成"
End Sub
