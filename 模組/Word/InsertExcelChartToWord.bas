Attribute VB_Name = "InsertExcelChartToWord"
Option Explicit
'*************************************************************************************
'模組名稱: InsertExcelChartToWord
'功能說明: 將 Excel 活頁簿中所有圖表依序插入新 Word 文件
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 確認 Excel 活頁簿含有至少一張圖表（內嵌或圖表工作表）
'  2. 執行 ExportAllChartsToWord
'  3. 選擇 Word 文件儲存路徑
'  4. 程式將所有圖表以圖片貼入 Word，每張圖後分頁
'
'注意事項:
'  - 圖表以 Enhanced Metafile (EMF) 格式插入，保持向量品質
'  - 圖表工作表與內嵌圖表均支援
'*************************************************************************************

'將 Excel 圖表全部匯出至 Word 文件
Sub ExportAllChartsToWord()
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wdRng       As Object
    Dim ws          As Worksheet
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim strSavePath As String
    Dim lngCount    As Long
    Dim strTmpImg   As String

    On Error GoTo ErrHandler

    '確認活頁簿有圖表
    lngCount = 0
    For Each ws In ThisWorkbook.Worksheets
        lngCount = lngCount + ws.ChartObjects.Count
    Next ws
    For Each cht In ThisWorkbook.Charts
        lngCount = lngCount + 1
    Next cht

    If lngCount = 0 Then
        MsgBox "活頁簿中找不到任何圖表！", vbExclamation, "提示"
        Exit Sub
    End If

    '選擇儲存路徑
    strSavePath = Application.GetSaveAsFilename( _
        InitialFileName:="圖表報告.docx", _
        FileFilter:="Word 文件 (*.docx), *.docx")
    If strSavePath = "False" Then Exit Sub

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Add

    '臨時圖片路徑
    strTmpImg = Environ("TEMP") & "\~chart_tmp.emf"

    lngCount = 0

    '處理內嵌圖表
    For Each ws In ThisWorkbook.Worksheets
        For Each chtObj In ws.ChartObjects
            chtObj.Chart.Export strTmpImg, "EMF"
            Set wdRng = wdDoc.Bookmarks("\EndOfDoc").Range

            '插入圖表標題
            wdRng.InsertAfter ws.Name & " - " & chtObj.Name
            wdRng.InsertParagraphAfter
            wdRng.Collapse 0

            '插入圖片
            wdDoc.InlineShapes.AddPicture _
                FileName:=strTmpImg, _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Range:=wdRng

            lngCount = lngCount + 1
            If lngCount < ThisWorkbook.Sheets.Count Then
                Set wdRng = wdDoc.Bookmarks("\EndOfDoc").Range
                wdRng.InsertBreak 7
            End If
        Next chtObj
    Next ws

    '處理圖表工作表
    For Each cht In ThisWorkbook.Charts
        cht.Export strTmpImg, "EMF"
        Set wdRng = wdDoc.Bookmarks("\EndOfDoc").Range
        wdRng.InsertAfter cht.Name
        wdRng.InsertParagraphAfter
        wdRng.Collapse 0

        wdDoc.InlineShapes.AddPicture _
            FileName:=strTmpImg, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=wdRng

        lngCount = lngCount + 1
    Next cht

    '刪除臨時檔案
    On Error Resume Next
    Kill strTmpImg
    On Error GoTo ErrHandler

    wdDoc.SaveAs2 strSavePath, 16
    wdApp.Visible = True

    MsgBox "圖表匯出完成！" & vbCrLf & _
           "共插入 " & lngCount & " 張圖表。" & vbCrLf & _
           "已儲存至：" & strSavePath, vbInformation, "完成"

    Set wdRng = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdRng = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
