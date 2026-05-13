Attribute VB_Name = "SplitByValueThreshold"
Option Explicit
'*************************************************************************************
'模組名稱: 依數值門檻分割工作表
'功能說明: 依據指定欄位的數值門檻，將資料分割為達標與未達標兩張工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub SplitByValueThreshold()
    Dim ws As Worksheet
    Dim wsAbove As Worksheet
    Dim wsBelow As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim threshold As Double
    Dim colIdx As Long
    Dim i As Long
    Dim aboveRow As Long
    Dim belowRow As Long
    Dim thresholdStr As String
    Dim colIdxStr As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表中沒有足夠的資料。", vbExclamation, "提示"
        Exit Sub
    End If

    thresholdStr = InputBox("請輸入數值門檻（例如：500）：", "設定門檻", "500")
    If thresholdStr = "" Then Exit Sub
    If Not IsNumeric(thresholdStr) Then
        MsgBox "請輸入有效的數值。", vbExclamation, "錯誤"
        Exit Sub
    End If
    threshold = CDbl(thresholdStr)

    colIdxStr = InputBox("請輸入要判斷的欄號（例如：2 代表第2欄）：", "選擇欄號", "2")
    If colIdxStr = "" Then Exit Sub
    colIdx = CLng(colIdxStr)
    If colIdx < 1 Or colIdx > lastCol Then
        MsgBox "欄號超出範圍。", vbExclamation, "錯誤"
        Exit Sub
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("達標").Delete
    ThisWorkbook.Worksheets("未達標").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsAbove = ThisWorkbook.Worksheets.Add
    wsAbove.Name = "達標"
    Set wsBelow = ThisWorkbook.Worksheets.Add
    wsBelow.Name = "未達標"

    ws.Rows(1).Copy Destination:=wsAbove.Rows(1)
    ws.Rows(1).Copy Destination:=wsBelow.Rows(1)

    aboveRow = 2
    belowRow = 2

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, colIdx).Value) Then
            If CDbl(ws.Cells(i, colIdx).Value) >= threshold Then
                ws.Rows(i).Copy Destination:=wsAbove.Rows(aboveRow)
                aboveRow = aboveRow + 1
            Else
                ws.Rows(i).Copy Destination:=wsBelow.Rows(belowRow)
                belowRow = belowRow + 1
            End If
        End If
    Next i

    wsAbove.Columns.AutoFit
    wsBelow.Columns.AutoFit

    MsgBox "分割完成！" & vbCrLf & _
           "達標：" & (aboveRow - 2) & " 列" & vbCrLf & _
           "未達標：" & (belowRow - 2) & " 列", vbInformation, "完成"
End Sub
