Attribute VB_Name = "ExportPDFWithFormData"
Option Explicit

'*************************************************************************************
'模組名稱: ExportPDFWithFormData
'功能說明: 將表單工作表依儲存格內容自動命名匯出為 PDF
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub ExportPDFWithFormData()
    Dim ws As Worksheet
    Dim pdfName As String
    Dim savePath As String
    Dim applicantName As String
    Dim formDate As String

    Set ws = ThisWorkbook.Worksheets(1)

    applicantName = Trim(CStr(ws.Range("B1").Value))
    formDate = Trim(CStr(ws.Range("B2").Value))

    If applicantName = "" Then
        MsgBox "請先填寫申請人姓名（B1）！", vbExclamation
        Exit Sub
    End If

    formDate = Replace(formDate, "/", "-")
    pdfName = applicantName & "_" & formDate & "_表單.pdf"
    savePath = ThisWorkbook.Path & "\" & pdfName

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 已匯出：" & pdfName, vbInformation
End Sub
