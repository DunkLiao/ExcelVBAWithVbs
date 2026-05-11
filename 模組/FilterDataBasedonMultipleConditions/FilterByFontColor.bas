Attribute VB_Name = "FilterByFontColor"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByFontColor
'功能說明: 依據指定欄位的字體顏色篩選資料，並將相同顏色的列複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口：建立測試資料並依紅色字體篩選
Sub TestFilterByFontColor()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("字色篩選測試")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "字色篩選測試"
    Else
        ws.Cells.Clear
    End If

    Call FillFontColorTestData(ws)
    Call FilterByFontColor(ws, 1, RGB(255, 0, 0), "紅色篩選結果")
End Sub

' 依字體顏色篩選資料
' srcWs        : 來源工作表
' checkCol     : 檢查字體顏色的欄號
' targetColor  : 目標字體顏色 (RGB 值)
' destSheetName: 輸出工作表名稱
Sub FilterByFontColor( _
    ByVal srcWs As Worksheet, _
    ByVal checkCol As Integer, _
    ByVal targetColor As Long, _
    ByVal destSheetName As String)

    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim destRow As Long
    Dim matchCount As Long

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets(destSheetName)
    On Error GoTo 0

    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        destWs.Name = destSheetName
    Else
        destWs.Cells.Clear
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, checkCol).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then
        MsgBox "來源工作表中無有效資料。", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 複製標題列
    srcWs.Rows(1).Copy Destination:=destWs.Rows(1)
    destRow = 2
    matchCount = 0

    For r = 2 To lastRow
        If srcWs.Cells(r, checkCol).Font.Color = targetColor Then
            srcWs.Rows(r).Copy Destination:=destWs.Rows(destRow)
            destRow = destRow + 1
            matchCount = matchCount + 1
        End If
    Next r

    destWs.Columns.AutoFit

    If matchCount = 0 Then
        MsgBox "未找到符合指定字體顏色的資料列。", vbInformation, "篩選結果"
    Else
        MsgBox "篩選完成！共找到 " & matchCount & " 列符合資料。" & vbCrLf & _
               "結果已複製至工作表：" & destSheetName, vbInformation, "篩選結果"
    End If
End Sub

' 填入字體顏色測試資料
Private Sub FillFontColorTestData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "金額"
    ws.Range("C1").Value = "備註"
    ws.Range("A1:C1").Font.Bold = True

    Dim i As Integer
    Dim labels(1 To 8) As String
    Dim amounts(1 To 8) As Long
    Dim notes(1 To 8) As String

    labels(1) = "收入A": amounts(1) = 50000: notes(1) = "正常"
    labels(2) = "支出B": amounts(2) = 12000: notes(2) = "異常"
    labels(3) = "收入C": amounts(3) = 75000: notes(3) = "正常"
    labels(4) = "支出D": amounts(4) = 99000: notes(4) = "警示"
    labels(5) = "收入E": amounts(5) = 30000: notes(5) = "正常"
    labels(6) = "支出F": amounts(6) = 45000: notes(6) = "異常"
    labels(7) = "收入G": amounts(7) = 61000: notes(7) = "正常"
    labels(8) = "支出H": amounts(8) = 88000: notes(8) = "警示"

    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = labels(i)
        ws.Cells(i + 1, 2).Value = amounts(i)
        ws.Cells(i + 1, 3).Value = notes(i)
    Next i

    ' 標記警示項目為紅色字體
    ws.Cells(5, 1).Font.Color = RGB(255, 0, 0)
    ws.Cells(5, 2).Font.Color = RGB(255, 0, 0)
    ws.Cells(5, 3).Font.Color = RGB(255, 0, 0)
    ws.Cells(9, 1).Font.Color = RGB(255, 0, 0)
    ws.Cells(9, 2).Font.Color = RGB(255, 0, 0)
    ws.Cells(9, 3).Font.Color = RGB(255, 0, 0)

    ' 標記異常項目為橘色字體
    ws.Cells(3, 1).Font.Color = RGB(255, 140, 0)
    ws.Cells(3, 2).Font.Color = RGB(255, 140, 0)
    ws.Cells(3, 3).Font.Color = RGB(255, 140, 0)
    ws.Cells(7, 1).Font.Color = RGB(255, 140, 0)
    ws.Cells(7, 2).Font.Color = RGB(255, 140, 0)
    ws.Cells(7, 3).Font.Color = RGB(255, 140, 0)
    ws.Columns("A:C").AutoFit
End Sub