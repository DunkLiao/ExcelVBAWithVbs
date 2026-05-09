Attribute VB_Name = "ExportExcelToWord"
Option Explicit
'*************************************************************************************
'模組名稱: ExportExcelToWord
'功能說明: 將 Excel 工作表資料匯出為 Word 文件表格
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 在 Excel 工作表中準備資料（第一列為標題）
'  2. 執行 ExportSheetDataToWord
'  3. Word 文件將自動開啟並插入完整表格
'
'注意事項:
'  - 需啟用 Microsoft Word Object Library 參考
'  - 資料範圍以 UsedRange 自動偵測
'*************************************************************************************

'將當前工作表資料匯出至新 Word 文件
Sub ExportSheetDataToWord()
    Dim wsData      As Worksheet
    Dim rngData     As Range
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wdTable     As Object
    Dim lngRows     As Long
    Dim lngCols     As Long
    Dim i           As Long
    Dim j           As Long
    Dim strSavePath As String

    On Error GoTo ErrHandler

    Set wsData = ActiveSheet

    '確認工作表有資料
    If wsData.UsedRange.Cells.Count <= 1 Then
        MsgBox "工作表無資料可匯出！", vbExclamation, "提示"
        Exit Sub
    End If

    Set rngData = wsData.UsedRange
    lngRows = rngData.Rows.Count
    lngCols = rngData.Columns.Count

    '建立 Word 應用程式物件
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    '建立新文件
    Set wdDoc = wdApp.Documents.Add

    '插入標題
    With wdDoc.Content
        .InsertAfter wsData.Name & " 資料報表"
        .InsertParagraphAfter
    End With

    '設定標題格式
    With wdDoc.Paragraphs(1).Range
        .Style = "Heading 1"
        .Bold = True
    End With

    '在文件末端插入段落後建立表格
    wdDoc.Content.InsertParagraphAfter
    Set wdTable = wdDoc.Tables.Add( _
        Range:=wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range, _
        NumRows:=lngRows, _
        NumColumns:=lngCols)

    '填入資料至 Word 表格
    For i = 1 To lngRows
        For j = 1 To lngCols
            wdTable.Cell(i, j).Range.Text = _
                CStr(rngData.Cells(i, j).Value)
        Next j
    Next i

    '設定表格樣式：標題列粗體並加底色
    With wdTable.Rows(1)
        .Range.Bold = True
        .Shading.BackgroundPatternColor = RGB(180, 198, 231)
    End With

    '自動調整欄寬（wdAutoFitContent = 1）
    wdTable.AutoFitBehavior 1

    '套用表格框線
    wdTable.Borders.Enable = True

    '提示儲存路徑
    strSavePath = Application.GetSaveAsFilename( _
        InitialFileName:=wsData.Name & "_報表.docx", _
        FileFilter:="Word 文件 (*.docx), *.docx")

    If strSavePath <> "False" Then
        wdDoc.SaveAs2 strSavePath, 16
        MsgBox "匯出成功！" & vbCrLf & "已儲存至：" & vbCrLf & strSavePath, vbInformation, "完成"
    Else
        MsgBox "已取消儲存，文件仍開啟中。", vbInformation, "提示"
    End If

    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set rngData = Nothing
    Set wsData = Nothing
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdTable = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
