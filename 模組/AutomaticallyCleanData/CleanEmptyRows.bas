Attribute VB_Name = "CleanEmptyRows"
Option Explicit
'*************************************************************************************
'模組名稱: CleanEmptyRows
'功能說明: 自動清理工作表中完全空白的資料列，並可選擇連同空白欄一併清理
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestCleanEmptyRows()
    Call CleanEmptyRowsAndColumns
End Sub

Sub CleanEmptyRowsAndColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim isEmpty As Boolean
    Dim rowCount As Long
    Dim colCount As Long
    Dim emptyCols As Object
    Dim col As Variant

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wsName As String
    wsName = "空白列清理"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName

    ' 建立含空白列的範例資料
    ws.Range("A1").Value = "編號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "部門"
    ws.Range("D1").Value = "分機"

    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "張小明"
    ws.Range("C2").Value = "業務部"
    ws.Range("D2").Value = 101

    ' 第3列為空白列

    ws.Range("A4").Value = 2
    ws.Range("B4").Value = "李小華"
    ws.Range("C4").Value = "工程部"
    ws.Range("D4").Value = 205

    ' 第5列為空白列

    ws.Range("A6").Value = 3
    ws.Range("B6").Value = "王大為"
    ws.Range("C6").Value = "人事部"
    ws.Range("D6").Value = 308

    ' 第7列只有一個欄位有值（部分空白列）

    ws.Range("A8").Value = 4
    ws.Range("B8").Value = "陳美玲"
    ws.Range("C8").Value = "財務部"

    ' 第9列完全空白

    ws.Range("A10").Value = 5
    ws.Range("B10").Value = "林志強"
    ws.Range("C10").Value = "業務部"
    ws.Range("D10").Value = 112

    ' 第11列為空白列

    ' 顯示清理前訊息
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    MsgBox "清理前共 " & lastRow & " 列資料（含標題列）。" & vbCrLf & _
           "第3、5、7、9、11列為空白或部分空白列。", vbInformation, "清理前"

    ' 從下往上刪除完全空白列
    rowCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = lastRow To 2 Step -1
        isEmpty = True
        For j = 1 To lastCol
            If Not IsEmpty(ws.Cells(i, j).Value) Then
                isEmpty = False
                Exit For
            End If
        Next j
        If isEmpty Then
            ws.Rows(i).Delete Shift:=xlUp
            rowCount = rowCount + 1
        End If
    Next i

    ' 清理完全空白欄（選用）
    Dim msg2 As Variant
    msg2 = MsgBox("是否也要刪除完全空白的欄位？", vbYesNo + vbQuestion, "清理空白欄")
    If msg2 = vbYes Then
        Set emptyCols = CreateObject("Scripting.Dictionary")
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        For j = lastCol To 1 Step -1
            isEmpty = True
            For i = 1 To lastRow
                If Not IsEmpty(ws.Cells(i, j).Value) Then
                    isEmpty = False
                    Exit For
                End If
            Next i
            If isEmpty Then
                emptyCols.Add j, j
            End If
        Next j

        colCount = 0
        For Each col In emptyCols.Keys
            ws.Columns(CLng(col)).Delete Shift:=xlToLeft
            colCount = colCount + 1
        Next col
    End If

    ws.Columns.AutoFit

    Application.ScreenUpdating = True

    MsgBox "清理完成！" & vbCrLf & _
           "刪除空白列：" & rowCount & " 列" & vbCrLf & _
           IIf(msg2 = vbYes, "刪除空白欄：" & colCount & " 欄", ""), vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "清理空白列時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
