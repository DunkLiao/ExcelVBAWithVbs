Attribute VB_Name = "CleanOutliersData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanOutliersData
'功能說明: 自動偵測並清除數值欄中超出平均值正負 N 個標準差的異常值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口（標準差倍數預設 2.5）
Sub TestCleanOutliersData()
    Dim ws As Worksheet
    Set ws = GetOrCreateOutlierWs(ThisWorkbook, "異常值清理範例")
    ws.Cells.Clear
    ws.Activate

    ws.Range("A1").Value = "編號"
    ws.Range("B1").Value = "銷售額"
    ws.Range("A1:B1").Font.Bold = True

    Dim testData(1 To 15) As Variant
    testData(1) = 5200
    testData(2) = 4800
    testData(3) = 5100
    testData(4) = 99999
    testData(5) = 5050
    testData(6) = 4900
    testData(7) = 5300
    testData(8) = -8888
    testData(9) = 5000
    testData(10) = 5200
    testData(11) = 4750
    testData(12) = 5400
    testData(13) = 88888
    testData(14) = 5100
    testData(15) = 4950

    Dim i As Long
    For i = 1 To 15
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = testData(i)
    Next i

    ws.Columns("A:B").AutoFit

    Call CleanOutliersInColumn(ws, 2, 2.5)
End Sub

' 清除指定欄位中的異常值（以標準差倍數為門檻）
' ws        : 目標工作表
' colIndex  : 數值欄位索引
' sdFactor  : 標準差倍數門檻（例如 2.5 表示排除超過平均正負 2.5 個標準差的值）
Sub CleanOutliersInColumn(ByVal ws As Worksheet, ByVal colIndex As Long, _
    ByVal sdFactor As Double)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    If lastRow < 3 Then
        MsgBox "資料筆數不足，無法計算異常值。", vbInformation, "提示"
        Exit Sub
    End If

    Dim values() As Double
    Dim count As Long
    count = 0
    ReDim values(lastRow - 2)

    Dim i As Long
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, colIndex).Value) Then
            values(count) = CDbl(ws.Cells(i, colIndex).Value)
            count = count + 1
        End If
    Next i

    If count < 2 Then
        MsgBox "有效數值筆數不足。", vbInformation, "提示"
        Exit Sub
    End If

    Dim total As Double
    total = 0
    Dim k As Long
    For k = 0 To count - 1
        total = total + values(k)
    Next k
    Dim avg As Double
    avg = total / count

    Dim variance As Double
    variance = 0
    For k = 0 To count - 1
        variance = variance + (values(k) - avg) ^ 2
    Next k
    Dim stdDev As Double
    stdDev = Sqr(variance / (count - 1))

    Dim lowerBound As Double
    Dim upperBound As Double
    lowerBound = avg - sdFactor * stdDev
    upperBound = avg + sdFactor * stdDev

    Dim clearedCount As Long
    clearedCount = 0

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, colIndex).Value) Then
            Dim cellVal As Double
            cellVal = CDbl(ws.Cells(i, colIndex).Value)
            If cellVal < lowerBound Or cellVal > upperBound Then
                ws.Cells(i, colIndex).Interior.Color = RGB(255, 200, 200)
                ws.Cells(i, colIndex).ClearContents
                ws.Cells(i, colIndex).AddComment "異常值已清除（原值：" & cellVal & "）"
                clearedCount = clearedCount + 1
            End If
        End If
    Next i

    MsgBox "異常值清理完成！" & vbCrLf & _
        "平均值：" & Round(avg, 2) & vbCrLf & _
        "標準差：" & Round(stdDev, 2) & vbCrLf & _
        "門檻範圍：" & Round(lowerBound, 2) & " ~ " & Round(upperBound, 2) & vbCrLf & _
        "已清除異常值：" & clearedCount & " 筆", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清理異常值時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateOutlierWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateOutlierWs = ws
End Function