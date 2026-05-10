Attribute VB_Name = "MergeWithFormulasIntact"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithFormulasIntact
'功能說明: 跨工作表合併資料時保留原始公式，而非只貼上值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestMergeWithFormulasIntact()
    Call MergeAllSheetsWithFormulas
End Sub

' 合併所有工作表資料並保留公式
Sub MergeAllSheetsWithFormulas()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim destWs As Worksheet
    Dim srcWs As Worksheet
    Dim destRow As Long
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim isFirstSheet As Boolean
    Dim destName As String

    Set wb = ThisWorkbook
    destName = "公式合併結果"

    ' 移除舊的結果工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(destName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 建立新的目的工作表
    Set destWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    destWs.Name = destName

    destRow = 1
    isFirstSheet = True

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> destName Then
            srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            srcLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If srcLastRow >= 1 And srcLastCol >= 1 Then
                If isFirstSheet Then
                    ' 第一張工作表：複製包含標頭的全部資料（含公式）
                    ws.Range(ws.Cells(1, 1), ws.Cells(srcLastRow, srcLastCol)).Copy _
                        Destination:=destWs.Cells(destRow, 1)
                    destRow = destRow + srcLastRow
                    isFirstSheet = False
                Else
                    ' 後續工作表：跳過標頭列，只複製資料（含公式）
                    If srcLastRow >= 2 Then
                        ws.Range(ws.Cells(2, 1), ws.Cells(srcLastRow, srcLastCol)).Copy _
                            Destination:=destWs.Cells(destRow, 1)
                        destRow = destRow + (srcLastRow - 1)
                    End If
                End If
            End If
        End If
    Next ws

    destWs.Columns.AutoFit
    destWs.Activate

    Application.ScreenUpdating = True

    MsgBox "跨工作表合併完成（含公式）！" & vbCrLf & _
           "結果已寫入：" & destName, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
