Option Explicit
Attribute VB_Name = "MergeExcelWithBlankRow"
'*************************************************************************************
'模組名稱: MergeExcelWithBlankRow
'功能說明: 合併多個 Excel 檔案到同一工作表，每個檔案之間插入空白行分隔
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestMergeWithBlankRow()
    Call MergeExcelWithBlankRowSeparator
End Sub

' 合併指定資料夾內所有 Excel 檔案，以空白行分隔
Sub MergeExcelWithBlankRowSeparator()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim fileName As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim dstRow As Long
    Dim lastRow As Long
    Dim srcLastRow As Long
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    ' 選擇來源資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 檔案的資料夾"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' 準備目標工作表
    Set dstWs = GetOrCreateWorksheet("合併結果")
    dstWs.Cells.Clear
    dstRow = 1

    ' 逐一處理每個 Excel 檔案
    For Each file In folder.Files
        fileName = file.Name
        If LCase(fso.GetExtensionName(fileName)) = "xlsx" Or _
           LCase(fso.GetExtensionName(fileName)) = "xls" Then
            Application.StatusBar = "處理中：" & fileName

            Set srcWb = Workbooks.Open(file.Path, ReadOnly:=True)
            Set srcWs = srcWb.Worksheets(1)

            ' 複製標題（僅第一個檔案）
            If dstRow = 1 Then
                srcLastRow = GetLastRow(srcWs)
                If srcLastRow >= 1 Then
                    srcWs.Rows(1).Copy Destination:=dstWs.Rows(dstRow)
                    dstRow = dstRow + 1
                End If
            End If

            ' 複製資料（略過標題）
            srcLastRow = GetLastRow(srcWs)
            If srcLastRow > 1 Then
                srcWs.Range("2:" & srcLastRow).Copy Destination:=dstWs.Rows(dstRow)
                dstRow = dstRow + (srcLastRow - 1)
            End If

            ' 插入空白行分隔
            dstRow = dstRow + 1

            srcWb.Close SaveChanges:=False
        End If
    Next file

    Application.StatusBar = False
    dstWs.Columns.AutoFit
    MsgBox "合併完成，共處理 " & folder.Files.Count & " 個檔案。" & vbCrLf & _
           "結果已寫入「合併結果」工作表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
End Sub

' 取得工作表最後一行
Private Function GetLastRow(ByRef ws As Worksheet) As Long
    Dim last As Long
    On Error Resume Next
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 1 Then last = 1
    GetLastRow = last
End Function

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
