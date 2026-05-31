Attribute VB_Name = "CompareByFormulaOutput"
Option Explicit

'*************************************************************************************
'模組名稱: CompareByFormulaOutput
'功能說明: 以公式計算方式比對兩個工作表的數值差異
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub CompareByFormulaOutput()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim val1 As Variant
    Dim val2 As Variant
    Dim diffCount As Long

    If ThisWorkbook.Worksheets.Count < 2 Then
        MsgBox "需要至少兩個工作表才能比對！", vbExclamation
        Exit Sub
    End If

    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("比對結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsResult = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsResult.Name = "比對結果"
    wsResult.Range("A1").Value = "列"
    wsResult.Range("B1").Value = "欄"
    wsResult.Range("C1").Value = "工作表 1 值"
    wsResult.Range("D1").Value = "工作表 2 值"
    wsResult.Range("E1").Value = "差異"
    wsResult.Range("A1:E1").Font.Bold = True

    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    diffCount = 0

    For i = 1 To lastRow
        For j = 1 To lastCol
            val1 = ws1.Cells(i, j).Value
            val2 = ws2.Cells(i, j).Value
            If CStr(val1) <> CStr(val2) Then
                diffCount = diffCount + 1
                wsResult.Cells(diffCount + 1, 1).Value = i
                wsResult.Cells(diffCount + 1, 2).Value = j
                wsResult.Cells(diffCount + 1, 3).Value = val1
                wsResult.Cells(diffCount + 1, 4).Value = val2
                wsResult.Cells(diffCount + 1, 5).Formula = _
                    "=C" & (diffCount + 1) & "-D" & (diffCount + 1)
            End If
        Next j
    Next i

    wsResult.Columns("A:E").AutoFit
    MsgBox "比對完成，共發現 " & diffCount & " 處差異！", vbInformation
End Sub
