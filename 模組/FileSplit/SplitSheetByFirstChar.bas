Attribute VB_Name = "SplitSheetByFirstChar"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByFirstChar
'功能說明: 依據指定欄位的第一個字元，將工作表資料分割至各子工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub SplitSheetByFirstChar()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim keyCol As Integer
    Dim firstChar As String
    Dim cellVal As String
    Dim sheetName As String
    Dim destRow As Long
    Dim delWs As Worksheet
    Dim wsNames() As String
    Dim wsCount As Integer
    Dim j As Integer

    Set ws = ActiveSheet
    keyCol = 1

    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表資料不足，無法分割。", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 刪除舊的分割工作表
    wsCount = 0
    For Each delWs In ThisWorkbook.Worksheets
        If Left(delWs.Name, 3) = "首字_" Then
            wsCount = wsCount + 1
            ReDim Preserve wsNames(wsCount - 1)
            wsNames(wsCount - 1) = delWs.Name
        End If
    Next delWs

    Application.DisplayAlerts = False
    For j = 0 To wsCount - 1
        ThisWorkbook.Worksheets(wsNames(j)).Delete
    Next j
    Application.DisplayAlerts = True

    For i = 2 To lastRow
        cellVal = CStr(ws.Cells(i, keyCol).Value)
        If Len(cellVal) > 0 Then
            firstChar = Left(cellVal, 1)
        Else
            firstChar = "空白"
        End If

        sheetName = "首字_" & firstChar

        On Error Resume Next
        Set destWs = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0

        If destWs Is Nothing Then
            Set destWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            destWs.Name = sheetName
            ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy destWs.Range("A1")
            destRow = 2
        Else
            destRow = destWs.Cells(destWs.Rows.Count, 1).End(xlUp).Row + 1
        End If

        ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy destWs.Cells(destRow, 1)
        Set destWs = Nothing
    Next i

    Application.ScreenUpdating = True
    MsgBox "已依首字元分割完成！", vbInformation, "完成"
End Sub
