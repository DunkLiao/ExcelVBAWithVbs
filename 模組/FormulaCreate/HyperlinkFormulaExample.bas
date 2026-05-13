Attribute VB_Name = "HyperlinkFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: HyperlinkFormulaExample
'功能說明: 批次在儲存格中建立 HYPERLINK 公式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestHyperlinkFormula()
    Call CreateHyperlinkFormulas("超連結公式範例")
End Sub

Sub CreateHyperlinkFormulas(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim i  As Long
    Dim url As String
    Dim lbl As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Call FillHyperlinkSourceData(ws)

    For i = 2 To 5
        url = ws.Cells(i, 2).Value
        lbl = ws.Cells(i, 1).Value
        ws.Cells(i, 3).Formula = "=HYPERLINK(""" & url & """,""" & lbl & """)"
    Next i

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:C").AutoFit
    MsgBox "超連結公式已批次建立完成！", vbInformation, "完成"
End Sub

Private Sub FillHyperlinkSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "網站名稱"
    ws.Range("B1").Value = "網址"
    ws.Range("C1").Value = "超連結公式"
    ws.Range("A2").Value = "Google"
    ws.Range("B2").Value = "https://www.google.com"
    ws.Range("A3").Value = "Microsoft"
    ws.Range("B3").Value = "https://www.microsoft.com"
    ws.Range("A4").Value = "GitHub"
    ws.Range("B4").Value = "https://www.github.com"
    ws.Range("A5").Value = "Wikipedia"
    ws.Range("B5").Value = "https://www.wikipedia.org"
End Sub

