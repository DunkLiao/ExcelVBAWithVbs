Option Explicit
Attribute VB_Name = "CompareAndHighlightRows"
'*************************************************************************************
'模組名稱: CompareAndHighlightRows
'功能說明: 比對前兩個工作表的資料，以黃色標示差異列，以橘色標示多餘列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub CompareAndHighlightRows()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim lastCol As Integer
    Dim i As Long
    Dim j As Integer
    Dim isDiff As Boolean
    Dim diffCount As Long
    Dim compareRows As Long

    On Error GoTo ErrHandler

    If ThisWorkbook.Sheets.Count < 2 Then
        MsgBox "活頁簿中必須至少有兩個工作表。", vbExclamation, "提示"
        Exit Sub
    End If

    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets(2)

    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column

    ws1.UsedRange.Interior.ColorIndex = xlNone
    ws2.UsedRange.Interior.ColorIndex = xlNone

    diffCount = 0
    compareRows = WorksheetFunction.Min(lastRow1, lastRow2)

    For i = 2 To compareRows
        isDiff = False
        For j = 1 To lastCol
            If CStr(ws1.Cells(i, j).Value) <> CStr(ws2.Cells(i, j).Value) Then
                isDiff = True
                Exit For
            End If
        Next j

        If isDiff Then
            diffCount = diffCount + 1
            ws1.Rows(i).Interior.Color = RGB(255, 230, 100)
            ws2.Rows(i).Interior.Color = RGB(255, 230, 100)
        End If
    Next i

    If lastRow1 > compareRows Then
        ws1.Range(ws1.Rows(compareRows + 1), ws1.Rows(lastRow1)).Interior.Color = RGB(255, 180, 80)
    End If

    If lastRow2 > compareRows Then
        ws2.Range(ws2.Rows(compareRows + 1), ws2.Rows(lastRow2)).Interior.Color = RGB(255, 180, 80)
    End If

    MsgBox "比對完成！" & vbCrLf & _
           "差異列數：" & diffCount & " 列" & vbCrLf & _
           "工作表 1 多餘列：" & WorksheetFunction.Max(0, lastRow1 - compareRows) & " 列" & vbCrLf & _
           "工作表 2 多餘列：" & WorksheetFunction.Max(0, lastRow2 - compareRows) & " 列", _
           vbInformation, "比對結果"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "比對並標示差異失敗"
End Sub