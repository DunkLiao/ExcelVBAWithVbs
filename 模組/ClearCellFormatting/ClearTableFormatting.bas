Attribute VB_Name = "ClearTableFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearTableFormatting
'功能說明: 清除工作表中所有 Excel 表格（ListObject）的格式，並將表格轉為一般範圍的範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestClearTableFormatting()
    Call ClearAllTableFormatting(ActiveSheet)
End Sub

' 清除指定工作表所有表格的格式並轉為一般範圍
' ws: 要處理的工作表
Sub ClearAllTableFormatting(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim tbl As ListObject
    Dim tableCount As Integer
    Dim clearedCount As Integer
    Dim answer As VbMsgBoxResult
    Dim i As Integer

    tableCount = ws.ListObjects.Count

    If tableCount = 0 Then
        MsgBox "工作表「" & ws.Name & "」中沒有任何表格（ListObject）。", _
               vbInformation, "無表格"
        Exit Sub
    End If

    answer = MsgBox("工作表「" & ws.Name & "」中共有 " & tableCount & " 個表格。" & vbCrLf & _
                    "是否清除所有表格格式並轉為一般範圍？" & vbCrLf & vbCrLf & _
                    "注意：此操作無法復原", _
                    vbYesNo + vbQuestion, "確認")
    If answer = vbNo Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    clearedCount = 0

    For i = ws.ListObjects.Count To 1 Step -1
        Set tbl = ws.ListObjects(i)

        ' 清除表格樣式
        tbl.TableStyle = ""

        ' 將表格轉為一般範圍（保留資料）
        tbl.Unlist

        clearedCount = clearedCount + 1
    Next i

    Application.ScreenUpdating = True
    MsgBox "已清除 " & clearedCount & " 個表格的格式並轉為一般範圍。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除表格格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
