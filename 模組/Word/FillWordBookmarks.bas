Attribute VB_Name = "FillWordBookmarks"
Option Explicit
'*************************************************************************************
'模組名稱: FillWordBookmarks
'功能說明: 以 Excel 工作表每列資料填入 Word 範本書籤，批次產生個人化文件
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 在 Word 範本中設定書籤（Bookmark），名稱需與 Excel 標題列相同
'  2. Excel 第一列為欄位名稱（對應書籤名稱），第二列起為資料
'  3. 執行 FillBookmarksFromExcel
'  4. 選擇 Word 範本檔案，再選擇輸出資料夾
'  5. 程式為每列資料產生一份獨立的 Word 文件
'
'注意事項:
'  - 書籤名稱不可含有空格，建議使用英文
'  - 輸出檔名為「列序號_欄A值.docx」
'*************************************************************************************

'以 Excel 資料填入 Word 書籤並批次輸出
Sub FillBookmarksFromExcel()
    Dim wsData       As Worksheet
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim strTemplate  As String
    Dim strOutFolder As String
    Dim strOutFile   As String
    Dim lngLastRow   As Long
    Dim lngLastCol   As Long
    Dim lngRow       As Long
    Dim lngCol       As Long
    Dim strBmkName   As String
    Dim strValue     As String
    Dim lngCount     As Long

    On Error GoTo ErrHandler

    Set wsData = ActiveSheet

    lngLastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lngLastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    If lngLastRow < 2 Then
        MsgBox "工作表無資料列！", vbExclamation, "提示"
        Exit Sub
    End If

    '選擇 Word 範本
    strTemplate = Application.GetOpenFilename( _
        FileFilter:="Word 文件 (*.docx), *.docx", _
        Title:="請選擇 Word 書籤範本")
    If strTemplate = "False" Then Exit Sub

    '選擇輸出資料夾
    With Application.FileDialog(4)
        .Title = "請選擇輸出資料夾"
        If .Show <> -1 Then Exit Sub
        strOutFolder = .SelectedItems(1)
    End With
    If Right(strOutFolder, 1) <> "\" Then strOutFolder = strOutFolder & "\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    lngCount = 0

    '逐列處理資料
    For lngRow = 2 To lngLastRow
        '開啟範本副本
        Set wdDoc = wdApp.Documents.Open(strTemplate)

        '逐欄對應書籤並填入資料
        For lngCol = 1 To lngLastCol
            strBmkName = Trim(CStr(wsData.Cells(1, lngCol).Value))
            strValue = CStr(wsData.Cells(lngRow, lngCol).Value)

            If strBmkName <> "" Then
                If wdDoc.Bookmarks.Exists(strBmkName) Then
                    wdDoc.Bookmarks(strBmkName).Range.Text = strValue
                End If
            End If
        Next lngCol

        '輸出檔案
        strOutFile = strOutFolder & lngRow - 1 & "_" & _
            CStr(wsData.Cells(lngRow, 1).Value) & ".docx"
        wdDoc.SaveAs2 strOutFile, 16
        wdDoc.Close False
        Set wdDoc = Nothing
        lngCount = lngCount + 1
    Next lngRow

    wdApp.Quit
    Set wdApp = Nothing
    Set wsData = Nothing

    MsgBox "批次產生完成！" & vbCrLf & _
           "共產生 " & lngCount & " 份文件。" & vbCrLf & _
           "輸出位置：" & strOutFolder, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
