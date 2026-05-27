Option Explicit
Attribute VB_Name = "ExportPDFByDateRange"
'*************************************************************************************

'模組名稱: ExportPDFByDateRange

'功能說明: 依使用者指定的日期範圍，篩選資料後匯出為 PDF 檔案

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub ExportPDFByDateRange()

    Dim ws As Worksheet

    Dim wsTmp As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim dateCol As Integer

    Dim startDate As Date

    Dim endDate As Date

    Dim startInput As String

    Dim endInput As String

    Dim cellDate As Date

    Dim outputPath As String

    Dim dstRow As Long



    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row



    If lastRow < 2 Then

        MsgBox "工作表資料不足。", vbExclamation, "提示"

        Exit Sub

    End If



    Dim colInput As String

    colInput = InputBox("請輸入日期欄號（例如：1 代表 A 欄）：", "日期範圍匯出 PDF", "1")

    If colInput = "" Then Exit Sub

    dateCol = CInt(colInput)



    startInput = InputBox("請輸入起始日期（格式：YYYY/MM/DD）：", "起始日期")

    If startInput = "" Then Exit Sub



    endInput = InputBox("請輸入結束日期（格式：YYYY/MM/DD）：", "結束日期")

    If endInput = "" Then Exit Sub



    On Error Resume Next

    startDate = CDate(startInput)

    endDate = CDate(endInput)

    On Error GoTo 0



    If startDate = 0 Or endDate = 0 Then

        MsgBox "日期格式錯誤，請使用 YYYY/MM/DD 格式。", vbExclamation, "錯誤"

        Exit Sub

    End If



    If endDate < startDate Then

        MsgBox "結束日期不得早於起始日期。", vbExclamation, "錯誤"

        Exit Sub

    End If



    outputPath = Application.DefaultFilePath & "\" & _

        "DateRange_" & Format(startDate, "YYYYMMDD") & "_" & _

        Format(endDate, "YYYYMMDD") & ".pdf"



    ' 建立臨時工作表

    Set wsTmp = ThisWorkbook.Worksheets.Add

    wsTmp.Name = "TmpDateFilter"

    dstRow = 1



    Application.ScreenUpdating = False



    ' 複製標題

    ws.Rows(1).Copy wsTmp.Rows(dstRow)

    dstRow = 2



    ' 篩選日期範圍

    For i = 2 To lastRow

        On Error Resume Next

        cellDate = CDate(ws.Cells(i, dateCol).Value)

        On Error GoTo 0

        If cellDate >= startDate And cellDate <= endDate Then

            ws.Rows(i).Copy wsTmp.Rows(dstRow)

            dstRow = dstRow + 1

        End If

    Next i



    wsTmp.Columns.AutoFit



    ' 匯出 PDF

    If dstRow > 2 Then

        wsTmp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outputPath, _

            Quality:=xlQualityStandard, IncludeDocProperties:=True, _

            IgnorePrintAreas:=False, OpenAfterPublish:=False

        MsgBox "PDF 匯出完成！" & vbCrLf & "共 " & (dstRow - 2) & " 筆資料" & vbCrLf & _

            "儲存路徑：" & outputPath, vbInformation, "完成"

    Else

        MsgBox "指定日期範圍內無資料，未產生 PDF。", vbInformation, "提示"

    End If



    ' 刪除臨時工作表

    Application.DisplayAlerts = False

    wsTmp.Delete

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

End Sub

