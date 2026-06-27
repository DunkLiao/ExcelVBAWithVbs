Attribute VB_Name = "ClearFormattingExceptSpecifiedStyle"
Option Explicit
'*************************************************************************************
'模組名稱: ClearFormattingExceptSpecifiedStyle
'功能說明: 清除指定範圍以外的所有格式，僅保留第1列標題列的指定樣式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestClearFormattingExceptSpecifiedStyle()
    Call ClearFormattingExceptSpecifiedStyle
End Sub

Sub ClearFormattingExceptSpecifiedStyle()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim clearCount As Long
    Dim keepCount As Long
    Dim r As Long
    Dim c As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("選擇性清除格式")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "選擇性清除格式"
    End If
    
    ws.Cells.Clear
    Call FillMixedFormatData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    clearCount = 0
    keepCount = 0
    
    For r = 1 To lastRow
        For c = 1 To lastCol
            Set cell = ws.Cells(r, c)
            With cell
                If r = 1 Then
                    If .Font.Name <> "微軟正黑體" Then .Font.Name = "微軟正黑體"
                    If .Font.Size <> 12 Then .Font.Size = 12
                    If .Font.Bold <> True Then .Font.Bold = True
                    keepCount = keepCount + 1
                Else
                    If .Font.Bold = True Then .Font.Bold = False
                    If .Font.Italic = True Then .Font.Italic = False
                    If .Font.Underline <> xlUnderlineStyleNone Then .Font.Underline = xlUnderlineStyleNone
                    If .Font.Color <> RGB(0, 0, 0) Then .Font.Color = RGB(0, 0, 0)
                    If .Interior.ColorIndex <> xlNone Then .Interior.ColorIndex = xlNone
                    If .Font.Name <> "微軟正黑體" Then .Font.Name = "微軟正黑體"
                    clearCount = clearCount + 1
                End If
            End With
        Next c
    Next r
    
    ws.Columns("A:C").AutoFit
    ws.Activate
    
    MsgBox "選擇性清除格式完成！保留 " & keepCount & " 個，清除 " & clearCount & " 個", vbInformation, "完成"
End Sub

Private Sub FillMixedFormatData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "預算"
    ws.Range("C1").Value = "實際"
    
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "微軟正黑體"
        .Interior.Color = RGB(200, 220, 255)
    End With
    
    ws.Range("A2").Value = "業務部"
    ws.Range("B2").Value = 1000
    ws.Range("C2").Value = 950
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Interior.Color = RGB(255, 255, 200)
    
    ws.Range("A3").Value = "研發部"
    ws.Range("B3").Value = 800
    ws.Range("C3").Value = 780
    ws.Range("B3").Font.Italic = True
    ws.Range("C3").Font.Color = RGB(255, 0, 0)
    
    ws.Range("A4").Value = "行銷部"
    ws.Range("B4").Value = 600
    ws.Range("C4").Value = 650
    ws.Range("A4").Font.Underline = xlUnderlineStyleSingle
    ws.Range("B4").Font.Size = 14
    
    ws.Range("A5").Value = "人事部"
    ws.Range("B5").Value = 300
    ws.Range("C5").Value = 290
    ws.Range("A5").Interior.Color = RGB(255, 230, 200)
    
    ws.Columns("A:C").AutoFit
End Sub
