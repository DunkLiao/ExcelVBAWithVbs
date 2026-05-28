Attribute VB_Name = "ExportPDFWithCoverPage"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithCoverPage
'功能說明: 建立含封面頁的 PDF 匯出範例，先插入封面工作表，合併匯出後再移除封面
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestExportPDFWithCoverPage()
    Dim savePath As String
    savePath = Environ("TEMP") & "\報表含封面.pdf"
    Call ExportWithCoverPage(ThisWorkbook, ActiveSheet, savePath)
End Sub

Sub ExportWithCoverPage(ByVal wb As Workbook, _
                        ByVal contentWs As Worksheet, _
                        ByVal outputPath As String)
    Dim coverWs    As Worksheet
    Dim sheetArr   As Variant

    Set coverWs = wb.Worksheets.Add(Before:=contentWs)
    coverWs.Name = "_封面_"
    Call BuildCoverPage(coverWs, contentWs.Name)

    sheetArr = Array(coverWs.Name, contentWs.Name)

    On Error GoTo ErrHandler
    wb.Worksheets(sheetArr).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    contentWs.Select
    Application.DisplayAlerts = False
    coverWs.Delete
    Application.DisplayAlerts = True

    MsgBox "含封面的 PDF 已匯出完畢！" & vbCrLf & "檔案位置：" & outputPath, _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    contentWs.Select
    Application.DisplayAlerts = False
    On Error Resume Next
    coverWs.Delete
    Application.DisplayAlerts = True
    MsgBox "PDF 匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub BuildCoverPage(ByVal coverWs As Worksheet, _
                            ByVal contentName As String)
    coverWs.Cells.Clear
    coverWs.Cells.Interior.Color = RGB(30, 60, 120)
    With coverWs.Range("D3:K3")
        .Merge
        .Value = "Dunk 企業股份有限公司"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(255, 215, 0)
        .HorizontalAlignment = xlCenter
    End With
    With coverWs.Range("D5:K5")
        .Merge
        .Value = contentName & " 報告"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    With coverWs.Range("D9:K9")
        .Merge
        .Value = "製作日期：" & Format(Date, "yyyy年mm月dd日")
        .Font.Size = 12
        .Font.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
    End With
    With coverWs.Range("D10:K10")
        .Merge
        .Value = "製作人員：Dunk"
        .Font.Size = 12
        .Font.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
    End With
    coverWs.Rows("1:15").RowHeight = 30
    coverWs.Columns("A:M").ColumnWidth = 8
End Sub
