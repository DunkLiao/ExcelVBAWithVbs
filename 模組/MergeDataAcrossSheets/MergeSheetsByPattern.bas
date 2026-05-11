Attribute VB_Name = "MergeSheetsByPattern"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsByPattern
'功能說明: 依工作表名稱關鍵字篩選，將符合條件的工作表資料合併至彙總工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestMergeSheetsByPattern()
    Call MergeSheetsByPattern
End Sub

' 依工作表名稱關鍵字合併資料
Sub MergeSheetsByPattern()
    On Error GoTo ErrorHandler

    Dim pattern As String
    Dim ws As Worksheet
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim isFirstMatch As Boolean
    Dim matchCount As Integer

    pattern = InputBox("請輸入工作表名稱關鍵字（符合的工作表將被合併）：" & vbCrLf & _
                       "例如輸入「月報」，會合併所有名稱含「月報」的工作表", _
                       "設定篩選樣式", "")
    If pattern = "" Then
        MsgBox "未輸入樣式，已取消", vbInformation, "取消"
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("樣式合併結果")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "樣式合併結果"
    Else
        wsTarget.Cells.Clear
    End If

    Application.ScreenUpdating = False

    targetRow = 1
    isFirstMatch = True
    matchCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = wsTarget.Name Then GoTo NextWs
        If InStr(1, ws.Name, pattern, vbTextCompare) = 0 Then GoTo NextWs

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        If lastRow < 1 Or lastCol < 1 Then GoTo NextWs

        If isFirstMatch Then
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy _
                Destination:=wsTarget.Cells(targetRow, 1)
            targetRow = targetRow + lastRow
            isFirstMatch = False
        Else
            If lastRow >= 2 Then
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                    Destination:=wsTarget.Cells(targetRow, 1)
                targetRow = targetRow + lastRow - 1
            End If
        End If

        matchCount = matchCount + 1

NextWs:
    Next ws

    wsTarget.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True

    If matchCount = 0 Then
        MsgBox "找不到名稱含「" & pattern & "」的工作表！", vbExclamation, "未找到"
    Else
        wsTarget.Activate
        MsgBox "已合併 " & matchCount & " 個工作表（名稱含「" & pattern & "」），" & _
               "共 " & targetRow - 1 & " 列資料。", vbInformation, "完成"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
