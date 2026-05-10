Attribute VB_Name = "CompareWithSummaryTable"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithSummaryTable
'功能說明: 比對兩個工作表的資料差異，並產生摘要統計表（新增、刪除、變更筆數）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestCompareWithSummaryTable()
    Call CompareWithSummaryTable
End Sub

' 比對兩個工作表並產生摘要表
Sub CompareWithSummaryTable()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim summaryWs As Worksheet
    Dim keyCol As Integer
    Dim addedCount As Long
    Dim deletedCount As Long
    Dim changedCount As Long
    Dim unchangedCount As Long

    Set wb = ThisWorkbook
    Set ws1 = GetOrCreateSheet(wb, "舊版資料")
    Set ws2 = GetOrCreateSheet(wb, "新版資料")
    Set summaryWs = GetOrCreateSheet(wb, "比對摘要")

    Call FillOldData(ws1)
    Call FillNewData(ws2)

    keyCol = 1
    addedCount = 0
    deletedCount = 0
    changedCount = 0
    unchangedCount = 0

    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim lastCol As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column

    Dim dictOld As Object
    Set dictOld = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow1
        Dim keyOld As String
        keyOld = CStr(ws1.Cells(i, keyCol).Value)
        If keyOld <> "" Then
            dictOld(keyOld) = i
        End If
    Next i

    Dim j As Long
    For j = 2 To lastRow2
        Dim keyNew As String
        keyNew = CStr(ws2.Cells(j, keyCol).Value)
        If keyNew = "" Then GoTo ContinueLoop

        If dictOld.Exists(keyNew) Then
            Dim oldRow As Long
            oldRow = dictOld(keyNew)
            Dim isDiff As Boolean
            isDiff = False
            Dim c As Integer
            For c = 1 To lastCol
                If CStr(ws1.Cells(oldRow, c).Value) <> CStr(ws2.Cells(j, c).Value) Then
                    isDiff = True
                    Exit For
                End If
            Next c
            If isDiff Then
                changedCount = changedCount + 1
            Else
                unchangedCount = unchangedCount + 1
            End If
            dictOld.Remove keyNew
        Else
            addedCount = addedCount + 1
        End If

ContinueLoop:
    Next j

    deletedCount = dictOld.Count

    Call WriteSummaryTable(summaryWs, addedCount, deletedCount, changedCount, unchangedCount)

    summaryWs.Activate
    MsgBox "比對完成！" & vbCrLf & _
           "新增：" & addedCount & " 筆" & vbCrLf & _
           "刪除：" & deletedCount & " 筆" & vbCrLf & _
           "變更：" & changedCount & " 筆" & vbCrLf & _
           "未變更：" & unchangedCount & " 筆", _
           vbInformation, "比對摘要"
    Exit Sub

ErrorHandler:
    MsgBox "比對時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 寫入摘要統計表
Private Sub WriteSummaryTable(ByVal ws As Worksheet, _
                               ByVal addedCount As Long, _
                               ByVal deletedCount As Long, _
                               ByVal changedCount As Long, _
                               ByVal unchangedCount As Long)
    ws.Range("A1").Value = "比對結果摘要"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "狀態"
    ws.Range("B3").Value = "筆數"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:B3").Font.Color = RGB(255, 255, 255)

    ws.Range("A4").Value = "新增"
    ws.Range("B4").Value = addedCount
    ws.Range("A4").Interior.Color = RGB(198, 239, 206)

    ws.Range("A5").Value = "刪除"
    ws.Range("B5").Value = deletedCount
    ws.Range("A5").Interior.Color = RGB(255, 199, 206)

    ws.Range("A6").Value = "變更"
    ws.Range("B6").Value = changedCount
    ws.Range("A6").Interior.Color = RGB(255, 235, 156)

    ws.Range("A7").Value = "未變更"
    ws.Range("B7").Value = unchangedCount

    ws.Range("A8").Value = "合計"
    ws.Range("B8").Formula = "=SUM(B4:B7)"
    ws.Range("A8:B8").Font.Bold = True

    ws.Columns("A:B").AutoFit
End Sub

' 填入舊版範例資料
Private Sub FillOldData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "部門"
    ws.Range("A1:C1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("001", "王小明", "業務部"), _
        Array("002", "李大華", "研發部"), _
        Array("003", "陳美玲", "行政部"), _
        Array("004", "張志偉", "財務部"), _
        Array("005", "林雅婷", "業務部") _
    )
    Dim i As Integer
    For i = 0 To 4
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
    Next i
    ws.Columns("A:C").AutoFit
End Sub

' 填入新版範例資料（含新增、刪除、變更）
Private Sub FillNewData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "部門"
    ws.Range("A1:C1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("001", "王小明", "業務部"), _
        Array("002", "李大華", "人資部"), _
        Array("004", "張志偉", "財務部"), _
        Array("005", "林雅婷", "行銷部"), _
        Array("006", "吳建國", "研發部") _
    )
    Dim i As Integer
    For i = 0 To 4
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
    Next i
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
