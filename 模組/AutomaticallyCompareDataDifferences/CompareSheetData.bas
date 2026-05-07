Attribute VB_Name = "CompareSheetData"
Option Explicit
'*************************************************************************************
'模組名稱: CompareSheetData
'功能說明: 自動比較兩個工作表的資料差異，並以顏色標示不同之處
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口（建立範例資料後進行比較）
Sub TestCompareSheetData()
    Call CreateSampleDataForCompare
    Call CompareSheetData("資料版本A", "資料版本B", "差異報告")
End Sub

' 建立兩個版本的範例資料
Private Sub CreateSampleDataForCompare()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheet(ThisWorkbook, "資料版本A")
    Set wsB = GetOrCreateSheet(ThisWorkbook, "資料版本B")

    wsA.Range("A1").Value = "員工編號"
    wsA.Range("B1").Value = "姓名"
    wsA.Range("C1").Value = "部門"
    wsA.Range("D1").Value = "薪資"
    wsA.Range("A2").Value = "E001" : wsA.Range("B2").Value = "張三"
    wsA.Range("C2").Value = "研發部" : wsA.Range("D2").Value = 60000
    wsA.Range("A3").Value = "E002" : wsA.Range("B3").Value = "李四"
    wsA.Range("C3").Value = "業務部" : wsA.Range("D3").Value = 55000
    wsA.Range("A4").Value = "E003" : wsA.Range("B4").Value = "王五"
    wsA.Range("C4").Value = "行政部" : wsA.Range("D4").Value = 45000
    wsA.Range("A5").Value = "E004" : wsA.Range("B5").Value = "趙六"
    wsA.Range("C5").Value = "財務部" : wsA.Range("D5").Value = 52000
    wsA.Columns("A:D").AutoFit

    wsB.Range("A1").Value = "員工編號"
    wsB.Range("B1").Value = "姓名"
    wsB.Range("C1").Value = "部門"
    wsB.Range("D1").Value = "薪資"
    wsB.Range("A2").Value = "E001" : wsB.Range("B2").Value = "張三"
    wsB.Range("C2").Value = "研發部" : wsB.Range("D2").Value = 65000
    wsB.Range("A3").Value = "E002" : wsB.Range("B3").Value = "李四"
    wsB.Range("C3").Value = "行銷部" : wsB.Range("D3").Value = 55000
    wsB.Range("A4").Value = "E003" : wsB.Range("B4").Value = "王五"
    wsB.Range("C4").Value = "行政部" : wsB.Range("D4").Value = 45000
    wsB.Range("A5").Value = "E004" : wsB.Range("B5").Value = "趙六"
    wsB.Range("C5").Value = "財務部" : wsB.Range("D5").Value = 52000
    wsB.Columns("A:D").AutoFit
End Sub

' 比較兩張工作表的差異並輸出報告
Sub CompareSheetData(ByVal sheetNameA As String, ByVal sheetNameB As String, _
                     ByVal reportSheetName As String)
    Dim wsA As Worksheet
    Dim wsB As Worksheet
    Dim wsReport As Worksheet
    Dim lastRowA As Long
    Dim lastColA As Long
    Dim r As Long
    Dim c As Long
    Dim diffCount As Long
    Dim reportRow As Long
    Dim valA As String
    Dim valB As String

    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets(sheetNameA)
    Set wsB = ThisWorkbook.Worksheets(sheetNameB)
    On Error GoTo 0

    If wsA Is Nothing Or wsB Is Nothing Then
        MsgBox "找不到指定的工作表，請確認工作表名稱！", vbExclamation, "錯誤"
        Exit Sub
    End If

    Set wsReport = GetOrCreateSheet(ThisWorkbook, reportSheetName)

    wsReport.Range("A1").Value = "工作表A"
    wsReport.Range("B1").Value = sheetNameA
    wsReport.Range("A2").Value = "工作表B"
    wsReport.Range("B2").Value = sheetNameB
    wsReport.Range("A3").Value = "差異儲存格"
    wsReport.Range("B3").Value = "版本A值"
    wsReport.Range("C3").Value = "版本B值"

    With wsReport.Range("A3:C3")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    reportRow = 4
    diffCount = 0

    lastRowA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastColA = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column

    For r = 1 To lastRowA
        For c = 1 To lastColA
            valA = CStr(wsA.Cells(r, c).Value)
            valB = CStr(wsB.Cells(r, c).Value)
            If valA <> valB Then
                wsReport.Cells(reportRow, 1).Value = wsA.Cells(r, c).Address(False, False)
                wsReport.Cells(reportRow, 2).Value = valA
                wsReport.Cells(reportRow, 3).Value = valB
                wsA.Cells(r, c).Interior.Color = RGB(255, 255, 0)
                wsB.Cells(r, c).Interior.Color = RGB(255, 200, 0)
                reportRow = reportRow + 1
                diffCount = diffCount + 1
            End If
        Next c
    Next r

    wsReport.Columns("A:C").AutoFit
    wsReport.Activate

    If diffCount = 0 Then
        MsgBox "兩個工作表資料完全相同！", vbInformation, "比較結果"
    Else
        MsgBox "比較完成！共發現 " & diffCount & " 處差異。", vbInformation, "比較結果"
    End If
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
