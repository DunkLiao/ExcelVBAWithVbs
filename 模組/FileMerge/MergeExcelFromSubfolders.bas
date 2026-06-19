Attribute VB_Name = "MergeExcelFromSubfolders"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelFromSubfolders
'功能說明: 遞迴搜尋主資料夾及其所有子資料夾中的 Excel 檔案，合併到一個總表中
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestMergeExcelFromSubfolders()
    Call MergeExcelFromSubfolders
End Sub

Sub MergeExcelFromSubfolders()
    Dim mainFolder As String
    Dim summaryWs As Worksheet
    Dim destRow As Long
    Dim fso As Object
    Dim mainFolderObj As Object
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 選擇主資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的主資料夾（含子資料夾）"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        mainFolder = .SelectedItems(1)
    End With
    
    ' 建立彙總工作表
    Dim wsName As String
    wsName = "彙總合併"
    On Error Resume Next
    ThisWorkbook.Sheets(wsName).Delete
    On Error GoTo ErrHandler
    Set summaryWs = ThisWorkbook.Sheets.Add
    summaryWs.Name = wsName
    
    ' 寫入標題列
    summaryWs.Cells(1, 1).Value = "檔案路徑"
    summaryWs.Cells(1, 2).Value = "工作表名稱"
    summaryWs.Cells(1, 3).Value = "列號"
    summaryWs.Cells(1, 4).Value = "A欄資料"
    summaryWs.Cells(1, 5).Value = "B欄資料"
    summaryWs.Cells(1, 6).Value = "C欄資料"
    destRow = 2
    
    ' 遞迴處理資料夾
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolderObj = fso.GetFolder(mainFolder)
    
    ProcessFolder mainFolderObj, summaryWs, destRow, fso
    
    summaryWs.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併完成！共處理至第 " & (destRow - 1) & " 列。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub ProcessFolder(ByRef folderObj As Object, ByRef ws As Worksheet, ByRef outRow As Long, ByRef fso As Object)
    Dim fileObj As Object
    Dim subFolderObj As Object
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim srcRow As Long
    Dim lastRow As Long
    Dim filePath As String
    
    On Error Resume Next
    
    ' 處理目前資料夾的 Excel 檔案
    For Each fileObj In folderObj.Files
        filePath = LCase(fileObj.Name)
        If Right(filePath, 4) = ".xls" Or Right(filePath, 5) = ".xlsx" Or _
           Right(filePath, 4) = ".xlsm" Or Right(filePath, 4) = ".xlt" Then
            If Left(fileObj.Name, 1) <> "~" Then
                Set srcWb = Workbooks.Open(fileObj.Path, ReadOnly:=True)
                For Each srcWs In srcWb.Worksheets
                    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
                    If lastRow >= 2 Then
                        For srcRow = 2 To lastRow
                            ws.Cells(outRow, 1).Value = fileObj.Path
                            ws.Cells(outRow, 2).Value = srcWs.Name
                            ws.Cells(outRow, 3).Value = srcRow
                            ws.Cells(outRow, 4).Value = srcWs.Cells(srcRow, 1).Value
                            ws.Cells(outRow, 5).Value = srcWs.Cells(srcRow, 2).Value
                            ws.Cells(outRow, 6).Value = srcWs.Cells(srcRow, 3).Value
                            outRow = outRow + 1
                        Next srcRow
                    End If
                Next srcWs
                srcWb.Close SaveChanges:=False
            End If
        End If
    Next fileObj
    
    ' 遞迴處理子資料夾
    For Each subFolderObj In folderObj.SubFolders
        ProcessFolder subFolderObj, ws, outRow, fso
    Next subFolderObj
End Sub
