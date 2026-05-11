Attribute VB_Name = "MergeExcelByColumnRange"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelByColumnRange
'功能說明: 依指定欄位範圍合併同一資料夾中的多個 Excel 檔案至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口：合併指定資料夾內所有 Excel 的 A 到 C 欄
Sub TestMergeByColumnRange()
    Dim folderPath As String
    folderPath = Environ("USERPROFILE") & "\Desktop\MergeTest\"
    Call MergeExcelByColumnRange(folderPath, 1, 3, "合併結果")
End Sub

' 依欄位範圍合併多個 Excel 檔案
' folderPath   : 來源資料夾路徑
' startCol     : 起始欄號 (1=A)
' endCol       : 結束欄號 (3=C)
' destSheetName: 合併結果工作表名稱
Sub MergeExcelByColumnRange( _
    ByVal folderPath As String, _
    ByVal startCol As Integer, _
    ByVal endCol As Integer, _
    ByVal destSheetName As String)

    Dim fso As Object
    Dim srcFile As Object
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim copyStartRow As Long
    Dim fileCount As Integer
    Dim r As Long
    Dim c As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Exit Sub
    End If

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets(destSheetName)
    On Error GoTo 0

    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add
        destWs.Name = destSheetName
    Else
        destWs.Cells.Clear
    End If

    destRow = 1
    fileCount = 0

    For Each srcFile In fso.GetFolder(folderPath).Files
        If LCase(Right(srcFile.Name, 4)) = ".xls" Or _
           LCase(Right(srcFile.Name, 5)) = ".xlsx" Or _
           LCase(Right(srcFile.Name, 5)) = ".xlsm" Then

            Set srcWb = Workbooks.Open(srcFile.Path, ReadOnly:=True)
            Set srcWs = srcWb.Worksheets(1)
            lastRow = srcWs.Cells(srcWs.Rows.Count, startCol).End(xlUp).Row

            If fileCount = 0 Then
                copyStartRow = 1
            Else
                copyStartRow = 2
            End If

            For r = copyStartRow To lastRow
                For c = startCol To endCol
                    destWs.Cells(destRow, c - startCol + 1).Value = _
                        srcWs.Cells(r, c).Value
                Next c
                destRow = destRow + 1
            Next r

            srcWb.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
    Next srcFile

    destWs.Columns.AutoFit
    MsgBox "合併完成！共處理 " & fileCount & " 個檔案，合計 " & _
           (destRow - 1) & " 列資料。", vbInformation, "完成"
    Set fso = Nothing
End Sub