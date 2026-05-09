Attribute VB_Name = "WordToPdfBatch"
Option Explicit
'*************************************************************************************
'模組名稱: WordToPdfBatch
'功能說明: 批次將指定資料夾內的所有 Word 文件轉存為 PDF 檔案
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 BatchConvertWordToPdf
'  2. 選擇含有 .docx 文件的資料夾
'  3. PDF 將輸出至相同資料夾，檔名與原檔相同
'
'注意事項:
'  - 需安裝 Word 並支援 PDF 匯出（Office 2010+）
'  - 原始 Word 文件不受影響
'*************************************************************************************

'批次轉換 Word 為 PDF
Sub BatchConvertWordToPdf()
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim strFolder    As String
    Dim strFile      As String
    Dim strPdfPath   As String
    Dim lngCount     As Long

    On Error GoTo ErrHandler

    '選擇來源資料夾
    With Application.FileDialog(4)
        .Title = "請選擇含有 Word 文件的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    lngCount = 0
    strFile = Dir(strFolder & "*.docx")

    Do While strFile <> ""
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)

        '輸出 PDF 至相同資料夾（wdExportFormatPDF = 17）
        strPdfPath = strFolder & Left(strFile, Len(strFile) - 5) & ".pdf"
        wdDoc.ExportAsFixedFormat OutputFileName:=strPdfPath, _
            ExportFormat:=17, _
            OpenAfterExport:=False, _
            OptimizeFor:=0

        wdDoc.Close False
        Set wdDoc = Nothing
        lngCount = lngCount + 1
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "PDF 轉換完成！" & vbCrLf & _
           "共轉換 " & lngCount & " 個文件。" & vbCrLf & _
           "輸出位置：" & strFolder, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
