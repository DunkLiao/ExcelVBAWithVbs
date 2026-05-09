Attribute VB_Name = "ExtractWordTextToExcel"
Option Explicit
'*************************************************************************************
'模組名稱: ExtractWordTextToExcel
'功能說明: 擷取資料夾內所有 Word 文件的段落文字，彙整至 Excel 工作表
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 ExtractAllWordTextToSheet
'  2. 選擇含有 .docx 文件的資料夾
'  3. 結果寫入新工作表「WordText」，欄位：檔名 | 段落序號 | 段落文字
'
'注意事項:
'  - 空白段落將略過
'  - 大量文件時處理時間較長，請耐心等待
'*************************************************************************************

'擷取所有 Word 文件段落文字至 Excel
Sub ExtractAllWordTextToSheet()
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wdPara      As Object
    Dim strFolder   As String
    Dim strFile     As String
    Dim wsOut       As Worksheet
    Dim lngOutRow   As Long
    Dim lngParaIdx  As Long
    Dim strParaText As String

    On Error GoTo ErrHandler

    With Application.FileDialog(4)
        .Title = "請選擇含有 Word 文件的資料夾"
        If .Show <> -1 Then Exit Sub
        strFolder = .SelectedItems(1)
    End With
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    '建立輸出工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("WordText").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "WordText"

    '寫入標題
    With wsOut
        .Cells(1, 1).Value = "檔案名稱"
        .Cells(1, 2).Value = "段落序號"
        .Cells(1, 3).Value = "段落文字"
        .Rows(1).Font.Bold = True
    End With
    lngOutRow = 2

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    strFile = Dir(strFolder & "*.docx")

    Do While strFile <> ""
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)
        lngParaIdx = 0

        For Each wdPara In wdDoc.Paragraphs
            strParaText = Trim(wdPara.Range.Text)
            '移除段落結尾特殊字元
            If Len(strParaText) > 0 Then
                strParaText = Left(strParaText, Len(strParaText) - 1)
            End If

            If strParaText <> "" Then
                lngParaIdx = lngParaIdx + 1
                wsOut.Cells(lngOutRow, 1).Value = strFile
                wsOut.Cells(lngOutRow, 2).Value = lngParaIdx
                wsOut.Cells(lngOutRow, 3).Value = strParaText
                lngOutRow = lngOutRow + 1
            End If
        Next wdPara

        wdDoc.Close False
        Set wdDoc = Nothing
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    wsOut.Columns.AutoFit
    MsgBox "擷取完成！共寫入 " & lngOutRow - 2 & " 筆段落資料。", _
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
