Attribute VB_Name = "WordCommentExtractToExcel"
Option Explicit
'*************************************************************************************
'模組名稱: WordCommentExtractToExcel
'功能說明: 擷取 Word 文件中的所有批注（Comments）並彙整至 Excel
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 ExtractWordCommentsToExcel
'  2. 選擇目標 Word 文件（支援單一檔案）
'  3. 批注清單寫入新工作表「WordComments」
'     欄位：序號 | 批注者 | 批注日期 | 被批注文字 | 批注內容
'
'注意事項:
'  - 適用於程式碼審查、文件校對等場景
'  - 若文件無批注，將顯示提示訊息
'*************************************************************************************

'擷取 Word 文件批注至 Excel
Sub ExtractWordCommentsToExcel()
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wdCmt       As Object
    Dim strFilePath As String
    Dim wsOut       As Worksheet
    Dim lngRow      As Long
    Dim lngIdx      As Long

    On Error GoTo ErrHandler

    strFilePath = Application.GetOpenFilename( _
        FileFilter:="Word 文件 (*.docx;*.doc), *.docx;*.doc", _
        Title:="請選擇含有批注的 Word 文件")
    If strFilePath = "False" Then Exit Sub

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(strFilePath)

    If wdDoc.Comments.Count = 0 Then
        MsgBox "此文件中找不到任何批注！", vbExclamation, "提示"
        wdDoc.Close False
        wdApp.Quit
        Exit Sub
    End If

    '建立輸出工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("WordComments").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "WordComments"

    '標題列
    With wsOut
        .Cells(1, 1).Value = "序號"
        .Cells(1, 2).Value = "批注者"
        .Cells(1, 3).Value = "批注日期"
        .Cells(1, 4).Value = "被批注文字"
        .Cells(1, 5).Value = "批注內容"
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(180, 198, 231)
    End With
    lngRow = 2

    '逐一讀取批注
    lngIdx = 1
    For Each wdCmt In wdDoc.Comments
        wsOut.Cells(lngRow, 1).Value = lngIdx
        wsOut.Cells(lngRow, 2).Value = wdCmt.Author
        wsOut.Cells(lngRow, 3).Value = Format(wdCmt.Date, "yyyy/mm/dd hh:mm")
        wsOut.Cells(lngRow, 4).Value = Trim(wdCmt.Scope.Text)
        wsOut.Cells(lngRow, 5).Value = Trim(wdCmt.Range.Text)
        lngRow = lngRow + 1
        lngIdx = lngIdx + 1
    Next wdCmt

    wdDoc.Close False
    wdApp.Quit
    Set wdCmt = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing

    wsOut.Columns.AutoFit
    MsgBox "批注擷取完成！共擷取 " & lngIdx - 1 & " 筆批注。", _
           vbInformation, "完成"
    Set wsOut = Nothing
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
