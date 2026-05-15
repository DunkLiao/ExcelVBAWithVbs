Attribute VB_Name = "ExportPDFWithSummaryPage"

Option Explicit

'*************************************************************************************

'模組名稱: ExportPDFWithSummaryPage

'功能說明: 建立暫時摘要頁後將整本活頁簿匯出為 PDF 並刪除暫存摘要頁

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/15

'

'*************************************************************************************



Public Sub RunExportPDFWithSummaryPage()

    On Error GoTo ErrorHandler



    Dim summarySheetName As String

    Dim wsSummary As Worksheet

    Dim pdfPath As Variant



    Application.ScreenUpdating = False

    Application.DisplayAlerts = False



    summarySheetName = GetUniqueTemporarySheetName("摘要")

    Set wsSummary = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))

    wsSummary.Name = summarySheetName



    Call FillSummaryPage(wsSummary)



    pdfPath = Application.GetSaveAsFilename( _

        InitialFileName:=ThisWorkbook.Path & "\活頁簿摘要匯出.pdf", _

        FileFilter:="PDF Files (*.pdf), *.pdf")



    If VarType(pdfPath) = vbBoolean Then

        If pdfPath = False Then GoTo CleanExit

    End If



    ThisWorkbook.ExportAsFixedFormat _

        Type:=xlTypePDF, _

        Filename:=CStr(pdfPath), _

        Quality:=xlQualityStandard, _

        IncludeDocProperties:=True, _

        IgnorePrintAreas:=False, _

        OpenAfterPublish:=False



    MsgBox "活頁簿 PDF 已匯出: " & CStr(pdfPath), vbInformation, "完成"



CleanExit:

    If Not wsSummary Is Nothing Then wsSummary.Delete

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    Exit Sub



ErrorHandler:

    On Error Resume Next

    If Not wsSummary Is Nothing Then wsSummary.Delete

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    On Error GoTo 0

    MsgBox "匯出 PDF 時發生錯誤: " & Err.Description, vbExclamation, "錯誤"

End Sub



Private Sub FillSummaryPage(ByVal wsSummary As Worksheet)

    Dim ws As Worksheet

    Dim rowIndex As Long



    wsSummary.Cells.Clear

    wsSummary.Range("A1:C1").Value = Array("工作表名稱", "頁數估計", "匯出時間")



    rowIndex = 2

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> wsSummary.Name Then

            wsSummary.Cells(rowIndex, 1).Value = ws.Name

            wsSummary.Cells(rowIndex, 2).Value = EstimateSheetPages(ws)

            wsSummary.Cells(rowIndex, 3).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")

            rowIndex = rowIndex + 1

        End If

    Next ws



    wsSummary.Columns("A:C").AutoFit

End Sub



Private Function EstimateSheetPages(ByVal ws As Worksheet) As Long

    Dim lastRow As Long

    Dim lastCol As Long

    Dim rowPages As Long

    Dim colPages As Long



    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then

        EstimateSheetPages = 1

        Exit Function

    End If



    lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    lastCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column



    rowPages = (lastRow + 44) \ 45

    colPages = (lastCol + 7) \ 8



    If rowPages < 1 Then rowPages = 1

    If colPages < 1 Then colPages = 1



    EstimateSheetPages = rowPages * colPages

End Function



Private Function GetUniqueTemporarySheetName(ByVal baseName As String) As String

    Dim indexValue As Long

    Dim candidate As String



    candidate = baseName

    indexValue = 1



    Do While SheetExists(candidate)

        candidate = baseName & "_暫存" & CStr(indexValue)

        indexValue = indexValue + 1

    Loop



    GetUniqueTemporarySheetName = candidate

End Function



Private Function SheetExists(ByVal sheetName As String) As Boolean

    On Error Resume Next

    SheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function

