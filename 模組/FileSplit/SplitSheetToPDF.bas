Option Explicit
Attribute VB_Name = "SplitSheetToPDF"
'*************************************************************************************
'模組名稱: SplitSheetToPDF
'功能說明: 將工作簿中每個工作表分別匯出為獨立的 PDF 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestSplitSheetToPDF()
    Dim outputFolder As String
    outputFolder = BrowseForPDFFolder("請選擇 PDF 輸出資料夾")

    If outputFolder = "" Then
        MsgBox "已取消操作。", vbInformation, "提示"
        Exit Sub
    End If

    Call SplitAllSheetsToPDF(ThisWorkbook, outputFolder)
End Sub

Sub SplitAllSheetsToPDF(ByVal wb As Workbook, ByVal outputFolder As String)
    On Error GoTo ErrorHandler

    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"

    Dim exportCount As Integer
    exportCount = 0

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Dim pdfPath As String
            pdfPath = outputFolder & CleanPDFFileName(ws.Name) & ".pdf"

            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False

            exportCount = exportCount + 1
        End If
    Next ws

    MsgBox "已將 " & exportCount & " 個工作表分別匯出為 PDF！", _
        vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function CleanPDFFileName(ByVal name As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("\\", "/", ":", "*", "?", Chr(34), "<", ">", "|")

    Dim result As String
    result = name

    Dim c As Variant
    For Each c In invalidChars
        result = Replace(result, c, "_")
    Next c

    CleanPDFFileName = result
End Function

Private Function BrowseForPDFFolder(ByVal title As String) As String
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")

    Dim folder As Object
    Set folder = shell.BrowseForFolder(0, title, 0, 17)

    If folder Is Nothing Then
        BrowseForPDFFolder = ""
    Else
        BrowseForPDFFolder = folder.Self.Path
    End If
End Function
