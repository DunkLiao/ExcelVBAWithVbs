Option Explicit
Attribute VB_Name = "ExportPDFWithBookmarks"
'*************************************************************************************
'模組名稱: ExportPDFWithBookmarks
'功能說明: 將活頁簿中每個工作表分別匯出為獨立 PDF，並以工作表名稱命名檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub ExportPDFWithBookmarks()
    Dim ws          As Worksheet
    Dim savePath    As String
    Dim pdfPath     As String
    Dim safeName    As String
    Dim count       As Integer

    ' 選擇儲存資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 儲存資料夾"
        If .Show = False Then Exit Sub
        savePath = .SelectedItems(1) & "\\"
    End With

    count = 0
    For Each ws In ThisWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then
            ' 清除非法字元作為檔名
            safeName = ws.Name
            safeName = Replace(safeName, "/", "-")
            safeName = Replace(safeName, "\\", "-")
            safeName = Replace(safeName, ":", "-")
            safeName = Replace(safeName, "*", "-")
            safeName = Replace(safeName, "?", "-")
            safeName = Replace(safeName, Chr(34), "-")
            safeName = Replace(safeName, "<", "-")
            safeName = Replace(safeName, ">", "-")
            safeName = Replace(safeName, "|", "-")

            pdfPath = savePath & safeName & ".pdf"

            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False

            count = count + 1
        End If
    Next ws

    MsgBox "共匯出 " & count & " 個工作表為 PDF 檔案。" & vbCrLf & _
           "儲存位置：" & savePath, vbInformation, "完成"
End Sub
