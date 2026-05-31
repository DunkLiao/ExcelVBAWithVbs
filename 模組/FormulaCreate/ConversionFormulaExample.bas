Attribute VB_Name = "ConversionFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: ConversionFormulaExample
'功能說明: 使用VBA批次建立單位換算公式（公里轉英里、攝氏轉華氏、公斤轉磅等）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestConversionFormulas()
    Call CreateConversionFormulas("單位換算範例")
End Sub

Sub CreateConversionFormulas(ByVal sheetName As String)
    Dim ws      As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ws.Range("A1").Value = "原始值"
    ws.Range("B1").Value = "公里轉英里"
    ws.Range("C1").Value = "攝氏轉華氏"
    ws.Range("D1").Value = "公斤轉磅"
    ws.Range("E1").Value = "公升轉加侖"
    ws.Range("F1").Value = "平方公尺轉坪"

    ws.Range("A2").Value = 10
    ws.Range("A3").Value = 50
    ws.Range("A4").Value = 100
    ws.Range("A5").Value = 200
    ws.Range("A6").Value = 500

    lastRow = 6

    ws.Range("B2:B" & lastRow).Formula = "=A2*0.621371"
    ws.Range("C2:C" & lastRow).Formula = "=A2*9/5+32"
    ws.Range("D2:D" & lastRow).Formula = "=A2*2.20462"
    ws.Range("E2:E" & lastRow).Formula = "=A2*0.264172"
    ws.Range("F2:F" & lastRow).Formula = "=A2/3.30579"

    ws.Range("B2:F" & lastRow).NumberFormat = "0.000"
    ws.Columns("A:F").AutoFit

    MsgBox "單位換算公式已批次建立完成！", vbInformation, "完成"
End Sub
