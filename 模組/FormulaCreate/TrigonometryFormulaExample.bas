Attribute VB_Name = "TrigonometryFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: TrigonometryFormulaExample
'功能說明: 使用VBA在Excel中插入三角函數公式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub InsertTrigonometryFormulas()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim i As Integer
    Dim row As Integer
    Dim deg As Double
    Dim angles As Variant

    sheetName = "三角函數公式範例"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 標題列
    ws.Range("A1").Value = "角度（度）"
    ws.Range("B1").Value = "SIN"
    ws.Range("C1").Value = "COS"
    ws.Range("D1").Value = "TAN"
    ws.Range("E1").Value = "DEGREES轉換"

    angles = Array(0, 30, 45, 60, 90, 120, 135, 150, 180)
    row = 2

    For i = 0 To UBound(angles)
        deg = angles(i)
        ws.Cells(row, 1).Value = deg
        ws.Cells(row, 2).Formula = "=SIN(RADIANS(A" & row & "))"
        ws.Cells(row, 3).Formula = "=COS(RADIANS(A" & row & "))"
        ws.Cells(row, 4).Formula = "=IF(A" & row & "=90,""N/A"",TAN(RADIANS(A" & row & ")))"
        ws.Cells(row, 5).Formula = "=DEGREES(RADIANS(A" & row & "))"
        row = row + 1
    Next i

    ws.Range("B2:E" & (row - 1)).NumberFormat = "0.0000"
    ws.Columns("A:E").AutoFit

    MsgBox "三角函數公式已插入完成！", vbInformation, "完成"
End Sub
