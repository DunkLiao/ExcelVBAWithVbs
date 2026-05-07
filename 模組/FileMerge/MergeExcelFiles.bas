Attribute VB_Name = "MergeExcelFiles"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelFiles
'功能說明: 將指定資料夾內所有Excel檔案的第一個工作表合併至一個彙總工作表
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口
Sub TestMergeExcelFiles()
    Call MergeExcelFilesFromFolder
End Sub

' 合併指定資料夾內所有 Excel 檔案
Sub MergeExcelFilesFromFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim isFirstFile As Boolean

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的 Excel 檔案所在資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("合併彙總")
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "合併彙總"
    Else
        wsTarget.Cells.Clear
    End If

    targetRow = 1
    isFirstFile = True

    fileName = Dir(folderPath & "\*.xlsx")

    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wbSource = Workbooks.Open(folderPath & "" & fileName)
            Set wsSource = wbSource.Worksheets(1)
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

            If lastRow >= 1 Then
                If isFirstFile Then
                    wsSource.Range("A1").Resize(lastRow, wsSource.UsedRange.Columns.Count).Copy _
                        Destination:=wsTarget.Cells(targetRow, 1)
                    targetRow = targetRow + lastRow
                    isFirstFile = False
                Else
                    If lastRow >= 2 Then
                        wsSource.Range("A2").Resize(lastRow - 1, wsSource.UsedRange.Columns.Count).Copy _
                            Destination:=wsTarget.Cells(targetRow, 1)
                        targetRow = targetRow + lastRow - 1
                    End If
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
    MsgBox "Excel 檔案合併完成！共合併至第 " & targetRow - 1 & " 列。", vbInformation, "完成"
End Sub
