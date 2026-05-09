Attribute VB_Name = "WordTableToExcel"
Option Explicit
'*************************************************************************************
'模組名稱: WordTableToExcel
'功能說明: 將指定 Word 文件中的所有表格資料匯入 Excel 工作表
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 ImportWordTablesToExcel
'  2. 選擇目標 Word 文件
'  3. 每張表格將匯入至新工作表，工作表名稱為「Table_N」
'
'注意事項:
'  - 若 Word 表格有合併儲存格，內容取自左上格
'  - 需啟用 Microsoft Word Object Library 參考
'*************************************************************************************

'匯入 Word 文件內所有表格至 Excel
Sub ImportWordTablesToExcel()
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wdTable     As Object
    Dim strFilePath As String
    Dim wsNew       As Worksheet
    Dim i           As Long
    Dim j           As Long
    Dim tbl         As Long
    Dim strSheetName As String

    On Error GoTo ErrHandler

    '選擇 Word 文件
    strFilePath = Application.GetOpenFilename( _
        FileFilter:="Word 文件 (*.docx;*.doc), *.docx;*.doc", _
        Title:="請選擇要匯入的 Word 文件")

    If strFilePath = "False" Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(strFilePath)

    If wdDoc.Tables.Count = 0 Then
        MsgBox "此文件中找不到任何表格！", vbExclamation, "提示"
        wdDoc.Close False
        wdApp.Quit
        Exit Sub
    End If

    '逐一處理每張表格
    For tbl = 1 To wdDoc.Tables.Count
        Set wdTable = wdDoc.Tables(tbl)

        '建立新工作表
        strSheetName = "Table_" & tbl
        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(strSheetName).Delete
        Application.DisplayAlerts = True
        On Error GoTo ErrHandler

        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNew.Name = strSheetName

        '逐列逐欄寫入資料
        For i = 1 To wdTable.Rows.Count
            For j = 1 To wdTable.Columns.Count
                On Error Resume Next
                wsNew.Cells(i, j).Value = Trim(wdTable.Cell(i, j).Range.Text)
                '移除 Word 儲存格結尾的特殊字元
                wsNew.Cells(i, j).Value = Left(wsNew.Cells(i, j).Value, _
                    Len(wsNew.Cells(i, j).Value) - 2)
                On Error GoTo ErrHandler
            Next j
        Next i

        '自動調整欄寬
        wsNew.Columns.AutoFit
    Next tbl

    wdDoc.Close False
    wdApp.Quit
    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing

    MsgBox "匯入完成！" & vbCrLf & _
           "共匯入 " & tbl - 1 & " 張表格。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
