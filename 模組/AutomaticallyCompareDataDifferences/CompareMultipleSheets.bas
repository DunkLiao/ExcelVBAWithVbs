Attribute VB_Name = "CompareMultipleSheets"
Option Explicit
'*************************************************************************************
'模組名稱: CompareMultipleSheets
'功能說明: 批次比對活頁簿中多對工作表，一次產生所有比對結果整合報告，
'          適用於多月份或多部門同時比對的情境
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口，批次比對三組工作表
Sub TestCompareMultipleSheets()
    Call CreateMultiSheetsData

    Dim pairs(1 To 3, 1 To 2) As String
    pairs(1, 1) = "Jan_舊版" : pairs(1, 2) = "Jan_新版"
    pairs(2, 1) = "Feb_舊版" : pairs(2, 2) = "Feb_新版"
    pairs(3, 1) = "Mar_舊版" : pairs(3, 2) = "Mar_新版"

    Call CompareMultipleSheets(pairs, "批次比對總表")
End Sub

' 建立多工作表批次比對範例資料
Private Sub CreateMultiSheetsData()
    Dim ws         As Worksheet
    Dim sheetNames(1 To 6) As String
    Dim i          As Integer

    sheetNames(1) = "Jan_舊版" : sheetNames(2) = "Jan_新版"
    sheetNames(3) = "Feb_舊版" : sheetNames(4) = "Feb_新版"
    sheetNames(5) = "Mar_舊版" : sheetNames(6) = "Mar_新版"

    For i = 1 To 6
        Set ws = GetOrCreateSheetCMS(sheetNames(i))
        ws.Range("A1").Value = "項目"
        ws.Range("B1").Value = "數量"
        ws.Range("C1").Value = "金額"
        ws.Range("A2").Value = "項目A"
        ws.Range("A3").Value = "項目B"
        ws.Range("A4").Value = "項目C"
        ' 舊版與新版填入不同數值以產生差異
        If InStr(sheetNames(i), "舊版") > 0 Then
            ws.Range("B2").Value = 100 : ws.Range("C2").Value = 5000
            ws.Range("B3").Value = 200 : ws.Range("C3").Value = 8000
            ws.Range("B4").Value = 150 : ws.Range("C4").Value = 6000
        Else
            ws.Range("B2").Value = 110 : ws.Range("C2").Value = 5000
            ws.Range("B3").Value = 200 : ws.Range("C3").Value = 8500
            ws.Range("B4").Value = 145 : ws.Range("C4").Value = 6000
        End If
        ws.Columns("A:C").AutoFit
    Next i
End Sub

' 批次比對多對工作表，輸出整合報告
Public Sub CompareMultipleSheets(ByRef pairList() As String, ByVal summarySheet As String)
    Dim wsR        As Worksheet
    Dim wsA        As Worksheet
    Dim wsB        As Worksheet
    Dim totalPairs As Long
    Dim p          As Long
    Dim r          As Long
    Dim c          As Long
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim rptRow     As Long
    Dim valA       As String
    Dim valB       As String
    Dim pairDiff   As Long
    Dim grandTotal As Long
    Dim sheetA     As String
    Dim sheetB     As String

    On Error GoTo ErrHandler

    Set wsR = GetOrCreateSheetCMS(summarySheet)

    wsR.Range("A1").Value = "批次工作表比對總覽"
    With wsR.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsR.Range("A2").Value = "執行時間: " & Format(Now, "yyyy/mm/dd hh:mm:ss")

    wsR.Range("A4").Value = "組別"
    wsR.Range("B4").Value = "工作表A"
    wsR.Range("C4").Value = "工作表B"
    wsR.Range("D4").Value = "儲存格"
    wsR.Range("E4").Value = "A值"
    wsR.Range("F4").Value = "B值"
    With wsR.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(0, 70, 127)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 5
    grandTotal = 0
    totalPairs = UBound(pairList, 1)

    For p = 1 To totalPairs
        sheetA = pairList(p, 1)
        sheetB = pairList(p, 2)
        pairDiff = 0

        Set wsA = Nothing
        Set wsB = Nothing
        On Error Resume Next
        Set wsA = ThisWorkbook.Worksheets(sheetA)
        Set wsB = ThisWorkbook.Worksheets(sheetB)
        On Error GoTo ErrHandler

        If wsA Is Nothing Or wsB Is Nothing Then
            wsR.Cells(rptRow, 1).Value = "第" & p & "組"
            wsR.Cells(rptRow, 2).Value = sheetA
            wsR.Cells(rptRow, 3).Value = sheetB
            wsR.Cells(rptRow, 4).Value = "(找不到工作表)"
            wsR.Cells(rptRow, 1).Resize(1, 6).Interior.Color = RGB(255, 199, 206)
            rptRow = rptRow + 1
        Else
            lastRow = Application.WorksheetFunction.Max( _
                wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row, _
                wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row)
            lastCol = Application.WorksheetFunction.Max( _
                wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column, _
                wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)

            For r = 1 To lastRow
                For c = 1 To lastCol
                    valA = CStr(wsA.Cells(r, c).Value)
                    valB = CStr(wsB.Cells(r, c).Value)
                    If valA <> valB Then
                        wsR.Cells(rptRow, 1).Value = "第" & p & "組"
                        wsR.Cells(rptRow, 2).Value = sheetA
                        wsR.Cells(rptRow, 3).Value = sheetB
                        wsR.Cells(rptRow, 4).Value = wsA.Cells(r, c).Address(False, False)
                        wsR.Cells(rptRow, 5).Value = valA
                        wsR.Cells(rptRow, 6).Value = valB
                        wsR.Cells(rptRow, 1).Resize(1, 6).Interior.Color = RGB(255, 255, 153)
                        rptRow = rptRow + 1
                        pairDiff = pairDiff + 1
                        grandTotal = grandTotal + 1
                    End If
                Next c
            Next r

            If pairDiff = 0 Then
                wsR.Cells(rptRow, 1).Value = "第" & p & "組"
                wsR.Cells(rptRow, 2).Value = sheetA
                wsR.Cells(rptRow, 3).Value = sheetB
                wsR.Cells(rptRow, 4).Value = "(無差異)"
                wsR.Cells(rptRow, 1).Resize(1, 6).Interior.Color = RGB(198, 239, 206)
                rptRow = rptRow + 1
            End If
        End If
    Next p

    ' 總計列
    wsR.Cells(rptRow + 1, 1).Value = "總計差異數"
    wsR.Cells(rptRow + 1, 2).Value = grandTotal & " 處"
    With wsR.Cells(rptRow + 1, 1).Resize(1, 6)
        .Font.Bold = True
        .Interior.Color = RGB(155, 194, 230)
    End With

    wsR.Columns("A:F").AutoFit
    wsR.Activate
    MsgBox "批次比對完成！共比對 " & totalPairs & " 組工作表，" & _
           "合計發現 " & grandTotal & " 處差異。", vbInformation, "批次比對結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCMS(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCMS = ws
End Function
