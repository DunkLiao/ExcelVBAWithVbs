Attribute VB_Name = "CompareByDateRange"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByDateRange
'功能說明: 比較兩張工作表中，指定日期範圍內的資料差異並標示結果
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub CompareByDateRange()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim resultWs As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim dateColIndex As Integer
    Dim keyColIndex As Integer
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim destRow As Long
    Dim i As Long
    Dim j As Long
    Dim found As Boolean
    Dim cellDate As Date
    Dim keyVal As String
    Dim startStr As String
    Dim endStr As String

    Const WS1_NAME As String = "資料A"
    Const WS2_NAME As String = "資料B"
    Const RESULT_NAME As String = "日期範圍差異"

    dateColIndex = 1
    keyColIndex = 2

    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(WS1_NAME)
    Set ws2 = ThisWorkbook.Worksheets(WS2_NAME)
    On Error GoTo 0

    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "找不到工作表「" & WS1_NAME & "」或「" & WS2_NAME & "」，請先建立資料。", vbExclamation, "錯誤"
        Exit Sub
    End If

    startStr = InputBox("請輸入起始日期（格式：yyyy/mm/dd）：", "設定日期範圍", Format(Date - 30, "yyyy/mm/dd"))
    If startStr = "" Then Exit Sub
    endStr = InputBox("請輸入結束日期（格式：yyyy/mm/dd）：", "設定日期範圍", Format(Date, "yyyy/mm/dd"))
    If endStr = "" Then Exit Sub

    If Not IsDate(startStr) Or Not IsDate(endStr) Then
        MsgBox "日期格式不正確，請重新輸入。", vbExclamation, "錯誤"
        Exit Sub
    End If

    startDate = CDate(startStr)
    endDate = CDate(endStr)

    On Error Resume Next
    Set resultWs = ThisWorkbook.Worksheets(RESULT_NAME)
    On Error GoTo 0

    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        resultWs.Name = RESULT_NAME
    Else
        resultWs.Cells.Clear
    End If

    resultWs.Range("A1").Value = "日期"
    resultWs.Range("B1").Value = "鍵值"
    resultWs.Range("C1").Value = "差異說明"
    destRow = 2

    lastRow1 = ws1.Cells(ws1.Rows.Count, dateColIndex).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, dateColIndex).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lastRow1
        If IsDate(ws1.Cells(i, dateColIndex).Value) Then
            cellDate = CDate(ws1.Cells(i, dateColIndex).Value)
            If cellDate >= startDate And cellDate <= endDate Then
                keyVal = CStr(ws1.Cells(i, keyColIndex).Value)
                found = False
                For j = 2 To lastRow2
                    If CStr(ws2.Cells(j, keyColIndex).Value) = keyVal Then
                        If IsDate(ws2.Cells(j, dateColIndex).Value) Then
                            If CDate(ws2.Cells(j, dateColIndex).Value) >= startDate And _
                               CDate(ws2.Cells(j, dateColIndex).Value) <= endDate Then
                                found = True
                                Exit For
                            End If
                        End If
                    End If
                Next j

                If Not found Then
                    resultWs.Cells(destRow, 1).Value = cellDate
                    resultWs.Cells(destRow, 2).Value = keyVal
                    resultWs.Cells(destRow, 3).Value = "僅存在於" & WS1_NAME
                    resultWs.Rows(destRow).Interior.Color = RGB(255, 220, 220)
                    destRow = destRow + 1
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    resultWs.Columns("A:C").AutoFit

    MsgBox "日期範圍比較完成！差異共 " & (destRow - 2) & " 筆。", vbInformation, "完成"
End Sub
