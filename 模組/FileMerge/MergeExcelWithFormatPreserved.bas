Attribute VB_Name = "MergeExcelWithFormatPreserved"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithFormatPreserved
'功能說明: 合併指定資料夾中所有 Excel 檔案，並完整保留儲存格格式（字型、色彩、框線）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestMergeWithFormatPreserved()
    Call MergeExcelWithFormatPreserved
End Sub

' 合併資料夾中所有 Excel 檔案並保留格式
Sub MergeExcelWithFormatPreserved()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim isFirstFile As Boolean

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的 Excel 檔案所在資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsTarget = GetOrCreateFmtWs(ThisWorkbook, "格式合併結果")
    wsTarget.Cells.Clear
    targetRow = 1
    isFirstFile = True

    fileName = Dir(folderPath & "\*.xlsx")
    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wbSource = Workbooks.Open(folderPath & "" & fileName)
            Set wsSource = wbSource.Worksheets(1)
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 Then
                Dim startRow As Long
                If isFirstFile Then
                    startRow = 1
                    isFirstFile = False
                Else
                    startRow = 2  ' 跳過標題列
                End If

                If startRow <= lastRow Then
                    ' 使用 PasteSpecial 保留格式
                    wsSource.Range(wsSource.Cells(startRow, 1), _
                        wsSource.Cells(lastRow, lastCol)).Copy

                    wsTarget.Cells(targetRow, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
                    Application.CutCopyMode = False

                    targetRow = targetRow + (lastRow - startRow + 1)
                End If
            End If

            wbSource.Close SaveChanges:=False
        End If
        fileName = Dir()
    Loop

    wsTarget.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    wsTarget.Activate

    MsgBox "已完成格式保留合併，共合併 " & targetRow - 1 & " 列資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立目標工作表
Private Function GetOrCreateFmtWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateFmtWs = ws
End Function
