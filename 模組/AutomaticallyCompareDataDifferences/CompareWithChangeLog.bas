Attribute VB_Name = "CompareWithChangeLog"
Option Explicit
'*************************************************************************************
'模組名稱: 比較並記錄變更日誌
'功能說明: 比對兩張工作表的差異，並將所有變更記錄至獨立的變更日誌工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub CompareWithChangeLog()
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim wsLog As Worksheet
    Dim oldName As String
    Dim newName As String
    Dim lastRowOld As Long
    Dim lastRowNew As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim logRow As Long
    Dim oldVal As String
    Dim newVal As String
    Dim maxRow As Long

    oldName = InputBox("請輸入舊版工作表名稱：", "設定工作表", "舊版")
    If oldName = "" Then Exit Sub
    newName = InputBox("請輸入新版工作表名稱：", "設定工作表", "新版")
    If newName = "" Then Exit Sub

    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets(oldName)
    Set wsNew = ThisWorkbook.Worksheets(newName)
    On Error GoTo 0

    If wsOld Is Nothing Then
        MsgBox "找不到工作表：" & oldName, vbExclamation, "錯誤"
        Exit Sub
    End If
    If wsNew Is Nothing Then
        MsgBox "找不到工作表：" & newName, vbExclamation, "錯誤"
        Exit Sub
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("變更日誌").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsLog = ThisWorkbook.Worksheets.Add
    wsLog.Name = "變更日誌"

    wsLog.Range("A1").Value = "列號"
    wsLog.Range("B1").Value = "欄號"
    wsLog.Range("C1").Value = "欄位標題"
    wsLog.Range("D1").Value = "舊值"
    wsLog.Range("E1").Value = "新值"
    wsLog.Range("F1").Value = "變更類型"
    wsLog.Range("A1:F1").Font.Bold = True
    logRow = 2

    lastRowOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column

    maxRow = lastRowOld
    If lastRowNew > maxRow Then maxRow = lastRowNew

    For i = 2 To maxRow
        For j = 1 To lastCol
            If i <= lastRowOld Then
                oldVal = CStr(wsOld.Cells(i, j).Value)
            Else
                oldVal = ""
            End If

            If i <= lastRowNew Then
                newVal = CStr(wsNew.Cells(i, j).Value)
            Else
                newVal = ""
            End If

            If oldVal <> newVal Then
                wsLog.Cells(logRow, 1).Value = i
                wsLog.Cells(logRow, 2).Value = j
                wsLog.Cells(logRow, 3).Value = wsOld.Cells(1, j).Value
                wsLog.Cells(logRow, 4).Value = oldVal
                wsLog.Cells(logRow, 5).Value = newVal
                If oldVal = "" Then
                    wsLog.Cells(logRow, 6).Value = "新增"
                    wsLog.Cells(logRow, 6).Font.Color = RGB(0, 128, 0)
                ElseIf newVal = "" Then
                    wsLog.Cells(logRow, 6).Value = "刪除"
                    wsLog.Cells(logRow, 6).Font.Color = RGB(192, 0, 0)
                Else
                    wsLog.Cells(logRow, 6).Value = "修改"
                    wsLog.Cells(logRow, 6).Font.Color = RGB(0, 0, 192)
                End If
                logRow = logRow + 1
            End If
        Next j
    Next i

    wsLog.Columns("A:F").AutoFit

    If logRow = 2 Then
        MsgBox "兩張工作表內容完全相同，無差異。", vbInformation, "比較結果"
    Else
        MsgBox "比較完成！共發現 " & (logRow - 2) & " 處差異，詳見「變更日誌」工作表。", _
               vbInformation, "比較結果"
    End If
End Sub
