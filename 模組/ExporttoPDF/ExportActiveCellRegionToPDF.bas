Attribute VB_Name = "ExportActiveCellRegionToPDF"
Option Explicit
'*************************************************************************************
'模組名稱: 匯出使用中儲存格區域為PDF
'功能說明: 將目前選取儲存格的CurrentRegion（連續資料區域）匯出為PDF
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub ExportActiveCellRegionToPDF()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim pdfPath As String
    Dim savePath As String

    Set ws = ActiveSheet

    If ActiveCell Is Nothing Then
        MsgBox "請先選取一個儲存格。", vbExclamation, "提示"
        Exit Sub
    End If

    Set targetRange = ActiveCell.CurrentRegion

    If targetRange Is Nothing Or targetRange.Cells.Count = 0 Then
        MsgBox "目前儲存格周圍沒有連續資料區域。", vbExclamation, "提示"
        Exit Sub
    End If

    savePath = Environ("USERPROFILE") & "\Desktop\CurrentRegion_Export.pdf"
    pdfPath = InputBox("請輸入PDF儲存路徑：", "儲存路徑", savePath)
    If pdfPath = "" Then Exit Sub

    ws.PageSetup.PrintArea = targetRange.Address

    On Error GoTo ErrorHandler
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ws.PageSetup.PrintArea = ""
    MsgBox "PDF已成功匯出至：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    ws.PageSetup.PrintArea = ""
    MsgBox "匯出PDF時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
