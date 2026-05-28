Attribute VB_Name = "CompareByAbsoluteDifference"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByAbsoluteDifference
'功能說明: 比較兩欄數值，計算絕對差異值並依差異大小以顏色標示等級
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Private Const HIGH_DIFF_THRESHOLD   As Double = 500
Private Const MEDIUM_DIFF_THRESHOLD As Double = 100

Sub TestCompareByAbsoluteDifference()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("絕對差異比較")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "絕對差異比較"
    End If
    ws.Cells.Clear
    Call FillAbsDiffData(ws)
    Call CompareByAbsoluteDifference(ws, 2, 3, 4, 5)
    ws.Columns("A:F").AutoFit
    MsgBox "絕對差異比較完畢！" & vbCrLf & _
           "紅色：差異 >= " & HIGH_DIFF_THRESHOLD & vbCrLf & _
           "橘色：差異 >= " & MEDIUM_DIFF_THRESHOLD & vbCrLf & _
           "綠色：差異 < " & MEDIUM_DIFF_THRESHOLD, _
           vbInformation, "完成"
End Sub

Sub CompareByAbsoluteDifference(ByVal ws As Worksheet, _
                                  ByVal col1 As Integer, _
                                  ByVal col2 As Integer, _
                                  ByVal diffCol As Integer, _
                                  ByVal colorCol As Integer)
    Dim i           As Long
    Dim lastRow     As Long
    Dim val1        As Double
    Dim val2        As Double
    Dim absDiff     As Double
    Dim cellColor   As Long
    Dim levelLabel  As String

    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "資料不足，至少需要 2 列。", vbExclamation, "錯誤"
        Exit Sub
    End If

    ws.Cells(1, diffCol).Value = "絕對差異"
    ws.Cells(1, colorCol).Value = "差異等級"
    ws.Cells(1, diffCol).Font.Bold = True
    ws.Cells(1, colorCol).Font.Bold = True

    Application.ScreenUpdating = False
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, col1).Value) And _
           IsNumeric(ws.Cells(i, col2).Value) Then
            val1 = CDbl(ws.Cells(i, col1).Value)
            val2 = CDbl(ws.Cells(i, col2).Value)
            absDiff = Abs(val1 - val2)
            If absDiff >= HIGH_DIFF_THRESHOLD Then
                cellColor  = RGB(255, 200, 200)
                levelLabel = "高差異"
            ElseIf absDiff >= MEDIUM_DIFF_THRESHOLD Then
                cellColor  = RGB(255, 228, 180)
                levelLabel = "中差異"
            Else
                cellColor  = RGB(200, 240, 200)
                levelLabel = "低差異"
            End If
            ws.Cells(i, diffCol).Value = absDiff
            ws.Cells(i, diffCol).NumberFormat = "0.00"
            ws.Cells(i, diffCol).Interior.Color = cellColor
            ws.Cells(i, colorCol).Value = levelLabel
            ws.Cells(i, colorCol).Interior.Color = cellColor
        Else
            ws.Cells(i, diffCol).Value = "非數值"
            ws.Cells(i, colorCol).Value = "-"
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Private Sub FillAbsDiffData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("項目", "預算金額", "實際金額")
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2:C2").Value = Array("項目01", 12000, 11850)
    ws.Range("A3:C3").Value = Array("項目02", 35000, 34200)
    ws.Range("A4:C4").Value = Array("項目03", 8000, 8750)
    ws.Range("A5:C5").Value = Array("項目04", 52000, 51400)
    ws.Range("A6:C6").Value = Array("項目05", 15000, 14200)
    ws.Range("A7:C7").Value = Array("項目06", 9000, 9600)
    ws.Range("A8:C8").Value = Array("項目07", 67000, 65800)
    ws.Range("A9:C9").Value = Array("項目08", 21000, 22500)
    ws.Range("A10:C10").Value = Array("項目09", 44000, 43100)
    ws.Range("A11:C11").Value = Array("項目10", 31000, 30200)
End Sub
