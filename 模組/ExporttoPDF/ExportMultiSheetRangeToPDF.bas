Attribute VB_Name = "ExportMultiSheetRangeToPDF"
Option Explicit
'*************************************************************************************
'模組名稱: ExportMultiSheetRangeToPDF
'功能說明: 從多個工作表中分別擷取指定範圍，合併輸出為單一 PDF 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestExportMultiSheetRangeToPDF()
    Call ExportMultiSheetRangeToPDF
End Sub

' 從多個工作表匯出指定範圍到單一 PDF
Sub ExportMultiSheetRangeToPDF()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetNames() As String
    Dim outputPath As String
    Dim i As Integer
    Dim count As Integer

    Set wb = ThisWorkbook

    ' 取得工作表清單（排除名稱含 "PDF" 字樣的工作表）
    count = 0
    For Each ws In wb.Worksheets
        If InStr(ws.Name, "PDF") = 0 Then
            ReDim Preserve sheetNames(count)
            sheetNames(count) = ws.Name
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        MsgBox "沒有可匯出的工作表。", vbInformation, "提示"
        Exit Sub
    End If

    ' 選取輸出路徑
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請輸入輸出 PDF 檔案名稱"
        .InitialFileName = wb.Path & "\多工作表範圍匯出.pdf"
        .FilterIndex = 2
        If .Show = False Then Exit Sub
        outputPath = .SelectedItems(1)
    End With

    If Right(LCase(outputPath), 4) <> ".pdf" Then
        outputPath = outputPath & ".pdf"
    End If

    Application.ScreenUpdating = False

    ' 設定各工作表的列印範圍（預設 A1:H20）
    Dim selectedSheets() As String
    ReDim selectedSheets(count - 1)

    For i = 0 To count - 1
        Set ws = wb.Worksheets(sheetNames(i))
        ws.PageSetup.PrintArea = "A1:H20"
        selectedSheets(i) = sheetNames(i)
    Next i

    ' 選取所有要匯出的工作表並匯出為 PDF
    wb.Sheets(selectedSheets).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' 還原選取狀態
    wb.Worksheets(1).Select

    Application.ScreenUpdating = True

    MsgBox "PDF 匯出完成！" & vbCrLf & "檔案路徑：" & outputPath, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
