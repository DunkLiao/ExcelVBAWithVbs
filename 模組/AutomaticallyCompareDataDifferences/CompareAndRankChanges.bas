Attribute VB_Name = "CompareAndRankChanges"
Option Explicit
'*************************************************************************************
'模組名稱: CompareAndRankChanges
'功能說明: 比較兩個工作表的數值欄位差異，並依變動幅度由大到小排名輸出
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestCompareAndRankChanges()
    Call CreateRankTestData
    Call CompareAndRankChanges("舊資料", "新資料", 1, 2, "變動排名結果")
End Sub

' 建立測試資料
Private Sub CreateRankTestData()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsOld As Worksheet
    Set wsOld = GetOrCreateRankWs(wb, "舊資料")
    wsOld.Range("A1").Value = "品項"
    wsOld.Range("B1").Value = "舊業績"
    wsOld.Range("A2").Value = "產品A" : wsOld.Range("B2").Value = 50000
    wsOld.Range("A3").Value = "產品B" : wsOld.Range("B3").Value = 80000
    wsOld.Range("A4").Value = "產品C" : wsOld.Range("B4").Value = 30000
    wsOld.Range("A5").Value = "產品D" : wsOld.Range("B5").Value = 120000
    wsOld.Range("A6").Value = "產品E" : wsOld.Range("B6").Value = 45000

    Dim wsNew As Worksheet
    Set wsNew = GetOrCreateRankWs(wb, "新資料")
    wsNew.Range("A1").Value = "品項"
    wsNew.Range("B1").Value = "新業績"
    wsNew.Range("A2").Value = "產品A" : wsNew.Range("B2").Value = 62000
    wsNew.Range("A3").Value = "產品B" : wsNew.Range("B3").Value = 75000
    wsNew.Range("A4").Value = "產品C" : wsNew.Range("B4").Value = 58000
    wsNew.Range("A5").Value = "產品D" : wsNew.Range("B5").Value = 115000
    wsNew.Range("A6").Value = "產品E" : wsNew.Range("B6").Value = 39000
End Sub

' 比較兩工作表數值差異並依變動幅度排名
' oldSheetName : 舊資料工作表名稱
' newSheetName : 新資料工作表名稱
' keyCol       : 關鍵欄位索引
' valueCol     : 數值欄位索引
' resultSheet  : 輸出結果工作表名稱
Sub CompareAndRankChanges(ByVal oldSheetName As String, ByVal newSheetName As String, _
    ByVal keyCol As Long, ByVal valueCol As Long, ByVal resultSheet As String)
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim wsResult As Worksheet

    Set wsOld = wb.Worksheets(oldSheetName)
    Set wsNew = wb.Worksheets(newSheetName)
    Set wsResult = GetOrCreateRankWs(wb, resultSheet)

    wsResult.Range("A1").Value = "品項"
    wsResult.Range("B1").Value = "舊數值"
    wsResult.Range("C1").Value = "新數值"
    wsResult.Range("D1").Value = "絕對變動"
    wsResult.Range("E1").Value = "變動率%"
    wsResult.Range("F1").Value = "變動排名"
    wsResult.Range("A1:F1").Font.Bold = True

    Dim oldLastRow As Long
    oldLastRow = wsOld.Cells(wsOld.Rows.Count, keyCol).End(xlUp).Row

    Dim resultRow As Long
    resultRow = 2

    Dim i As Long
    For i = 2 To oldLastRow
        Dim keyValue As String
        keyValue = CStr(wsOld.Cells(i, keyCol).Value)
        Dim oldValue As Double
        oldValue = CDbl(wsOld.Cells(i, valueCol).Value)

        Dim newLastRow As Long
        newLastRow = wsNew.Cells(wsNew.Rows.Count, keyCol).End(xlUp).Row
        Dim j As Long
        Dim newValue As Double
        newValue = 0
        Dim found As Boolean
        found = False

        For j = 2 To newLastRow
            If CStr(wsNew.Cells(j, keyCol).Value) = keyValue Then
                newValue = CDbl(wsNew.Cells(j, valueCol).Value)
                found = True
                Exit For
            End If
        Next j

        If found Then
            Dim diff As Double
            diff = newValue - oldValue
            Dim changeRate As Double
            If oldValue <> 0 Then
                changeRate = diff / oldValue * 100
            Else
                changeRate = 0
            End If

            wsResult.Cells(resultRow, 1).Value = keyValue
            wsResult.Cells(resultRow, 2).Value = oldValue
            wsResult.Cells(resultRow, 3).Value = newValue
            wsResult.Cells(resultRow, 4).Value = diff
            wsResult.Cells(resultRow, 5).Value = Round(changeRate, 2)
            resultRow = resultRow + 1
        End If
    Next i

    Dim lastResultRow As Long
    lastResultRow = resultRow - 1

    If lastResultRow >= 2 Then
        Dim r As Long
        For r = 2 To lastResultRow
            wsResult.Cells(r, 6).Formula = "=RANK(ABS(D" & r & "),ABS($D$2:$D$" & lastResultRow & "))"
        Next r

        wsResult.Range("A1:F" & lastResultRow).Sort _
            Key1:=wsResult.Range("D1"), Order1:=xlDescending, Header:=xlYes
    End If

    If lastResultRow >= 2 Then
        Dim diffRange As Range
        Set diffRange = wsResult.Range("D2:E" & lastResultRow)
        diffRange.FormatConditions.Delete

        Dim fcPos As FormatCondition
        Set fcPos = diffRange.FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        fcPos.Font.Color = RGB(0, 128, 0)

        Dim fcNeg As FormatCondition
        Set fcNeg = diffRange.FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        fcNeg.Font.Color = RGB(192, 0, 0)
    End If

    wsResult.UsedRange.Columns.AutoFit
    wsResult.Activate

    MsgBox "比較排名完成，共 " & lastResultRow - 1 & " 筆比對資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "比較排名時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateRankWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateRankWs = ws
End Function