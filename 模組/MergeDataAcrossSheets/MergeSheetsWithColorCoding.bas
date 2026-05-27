Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsWithColorCoding
'功能說明: 合併活頁簿中所有工作表的資料，並依來源工作表名稱套用不同底色標示
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub MergeSheetsWithColorCoding()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim mergeWs As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim maxLastCol As Long
    Dim headerDone As Boolean
    Dim sheetIndex As Long
    Dim colorPalette(0 To 9) As Long
    Dim clr As Long
    Dim i As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    ' 10 種輪替底色
    colorPalette(0) = RGB(198, 224, 180)
    colorPalette(1) = RGB(189, 215, 238)
    colorPalette(2) = RGB(255, 230, 153)
    colorPalette(3) = RGB(248, 203, 173)
    colorPalette(4) = RGB(217, 210, 233)
    colorPalette(5) = RGB(180, 238, 180)
    colorPalette(6) = RGB(255, 192, 203)
    colorPalette(7) = RGB(255, 255, 180)
    colorPalette(8) = RGB(173, 216, 230)
    colorPalette(9) = RGB(221, 221, 221)

    ' 移除舊的合併工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("合併(色碼)").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    ' 建立合併結果工作表
    Set mergeWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    mergeWs.Name = "合併(色碼)"

    destRow = 1
    headerDone = False
    sheetIndex = 0
    maxLastCol = 0

    Application.ScreenUpdating = False

    ' 遍歷所有工作表（排除合併工作表本身）
    For Each ws In wb.Sheets
        If ws.Name <> mergeWs.Name Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 And lastCol >= 1 Then
                clr = colorPalette(sheetIndex Mod 10)
                If lastCol > maxLastCol Then maxLastCol = lastCol

                If Not headerDone Then
                    ' 第一個工作表：含標題整行複製
                    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy _
                        Destination:=mergeWs.Cells(destRow, 1)

                    ' 加入「來源工作表」欄標題
                    mergeWs.Cells(1, lastCol + 1).Value = "來源工作表"
                    mergeWs.Cells(1, lastCol + 1).Font.Bold = True

                    ' 資料列（第2列起）填色並標示來源
                    If lastRow > 1 Then
                        For i = 2 To lastRow
                            mergeWs.Rows(i).Interior.Color = clr
                            mergeWs.Cells(i, lastCol + 1).Value = ws.Name
                        Next i
                    End If

                    destRow = destRow + lastRow
                    headerDone = True
                Else
                    ' 後續工作表：跳過第一列標題
                    If lastRow > 1 Then
                        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                            Destination:=mergeWs.Cells(destRow, 1)

                        For i = destRow To destRow + lastRow - 2
                            mergeWs.Rows(i).Interior.Color = clr
                            mergeWs.Cells(i, maxLastCol + 1).Value = ws.Name
                        Next i

                        destRow = destRow + lastRow - 1
                    End If
                End If

                sheetIndex = sheetIndex + 1
            End If
        End If
    Next ws

    ' 美化標題列
    mergeWs.Rows(1).Interior.Color = RGB(68, 114, 196)
    mergeWs.Rows(1).Font.Color = RGB(255, 255, 255)
    mergeWs.Rows(1).Font.Bold = True
    mergeWs.Columns.AutoFit

    Application.ScreenUpdating = True

    MsgBox "合併並色碼標示完成！共合併 " & sheetIndex & " 個工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
