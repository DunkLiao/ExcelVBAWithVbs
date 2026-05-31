Attribute VB_Name = "ExportPDFWithRotatedText"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithRotatedText
'功能說明: 將包含旋轉文字儲存格的工作表匯出為PDF，並示範旋轉文字設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestExportPDFWithRotatedText()
    Call ExportPDFWithRotatedText("旋轉文字PDF範例")
End Sub

Sub ExportPDFWithRotatedText(ByVal sheetName As String)
    Dim ws       As Worksheet
    Dim savePath As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    With ws.Range("B2:F2")
        .Value = Array("一月", "二月", "三月", "四月", "五月")
        .Orientation = 45
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .RowHeight = 60
    End With

    ws.Range("A3").Value = "北區"
    ws.Range("A4").Value = "中區"
    ws.Range("A5").Value = "南區"

    ws.Range("B3:F3").Value = Array(120, 135, 98, 145, 162)
    ws.Range("B4:F4").Value = Array(88, 92, 105, 99, 115)
    ws.Range("B5:F5").Value = Array(76, 83, 91, 88, 102)

    With ws.Range("A2")
        .Value = "地區/業績"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Columns("A:F").AutoFit
    ws.PageSetup.PrintArea = "A1:F6"

    savePath = Application.DefaultFilePath & "\ExportPDFWithRotatedText_" & _
               Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "旋轉文字PDF已匯出至：" & Chr(13) & savePath, vbInformation, "完成"
End Sub
