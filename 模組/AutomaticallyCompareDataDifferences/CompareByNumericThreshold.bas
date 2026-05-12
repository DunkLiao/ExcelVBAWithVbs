Attribute VB_Name = "CompareByNumericThreshold"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByNumericThreshold
'功能說明: 比較兩欄數值，當差異超過指定閾值時標記為異常，並匯出差異報告
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestCompareByNumericThreshold()
    Call CompareByNumericThreshold
End Sub

' 依數字閾值比較兩欄差異
Sub CompareByNumericThreshold()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim threshold As Double
    Dim thresholdInput As String
    Dim i As Long
    Dim reportRow As Long
    Dim diffCount As Long
    Dim val1 As Double
    Dim val2 As Double
    Dim diff As Double

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "資料不足，請確認工作表有標題列與資料。", vbExclamation, "提示"
        Exit Sub
    End If

    thresholdInput = InputBox("請輸入差異閾值（超過此值視為異常）：", "設定閾值", "100")
    If thresholdInput = "" Then
        MsgBox "未輸入閾值，程式結束。", vbInformation, "取消"
        Exit Sub
    End If

    If Not IsNumeric(thresholdInput) Then
        MsgBox "請輸入有效的數字！", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    threshold = CDbl(thresholdInput)

    Set wsReport = GetOrCreateSheet(ThisWorkbook, "數字閾值差異報告")
    wsReport.Range("A1").Value = "列號"
    wsReport.Range("B1").Value = "欄B值"
    wsReport.Range("C1").Value = "欄C值"
    wsReport.Range("D1").Value = "差異"
    wsReport.Range("E1").Value = "差異是否超標"
    wsReport.Range("A1:E1").Font.Bold = True

    reportRow = 2
    diffCount = 0

    Application.ScreenUpdating = False

    ws.Range("B2:C" & lastRow).Interior.ColorIndex = xlNone

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 2).Value) And IsNumeric(ws.Cells(i, 3).Value) Then
            val1 = CDbl(ws.Cells(i, 2).Value)
            val2 = CDbl(ws.Cells(i, 3).Value)
            diff = Abs(val1 - val2)

            If diff > threshold Then
                ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206)
                ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

                wsReport.Cells(reportRow, 1).Value = i
                wsReport.Cells(reportRow, 2).Value = val1
                wsReport.Cells(reportRow, 3).Value = val2
                wsReport.Cells(reportRow, 4).Value = diff
                wsReport.Cells(reportRow, 5).Value = "超標"
                wsReport.Cells(reportRow, 5).Font.Color = RGB(156, 0, 6)
                reportRow = reportRow + 1
                diffCount = diffCount + 1
            End If
        End If
    Next i

    wsReport.Columns("A:E").AutoFit
    Application.ScreenUpdating = True

    MsgBox "比較完成！" & vbCrLf & _
           "閾值：" & threshold & vbCrLf & _
           "超標筆數：" & diffCount & " 筆", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "比較時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立範例資料並執行比較
Sub CreateThresholdCompareExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "閾值比較範例")

    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "預算"
    ws.Range("C1").Value = "實際"
    ws.Range("A1:C1").Font.Bold = True

    ws.Range("A2").Value = "差旅費"
    ws.Range("B2").Value = 5000
    ws.Range("C2").Value = 5080

    ws.Range("A3").Value = "餐費"
    ws.Range("B3").Value = 3000
    ws.Range("C3").Value = 3450

    ws.Range("A4").Value = "住宿費"
    ws.Range("B4").Value = 8000
    ws.Range("C4").Value = 8050

    ws.Range("A5").Value = "材料費"
    ws.Range("B5").Value = 12000
    ws.Range("C5").Value = 13500

    ws.Range("A6").Value = "雜費"
    ws.Range("B6").Value = 500
    ws.Range("C6").Value = 490

    ws.Columns("A:C").AutoFit
    ws.Activate

    MsgBox "範例資料已建立。請執行 CompareByNumericThreshold 進行比較。", vbInformation, "提示"
    Exit Sub

ErrorHandler:
    MsgBox "建立範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
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
