Option Explicit
Attribute VB_Name = "CompareWithVersionHistory"
'*************************************************************************************
'模組名稱: CompareWithVersionHistory
'功能說明: 將當前工作表資料與歷史版本快照比對，標示新增、修改與刪除的差異
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestCompareWithVersionHistory()
    Call CreateVersionCompareDemo(ThisWorkbook)
End Sub

Sub CreateVersionCompareDemo(ByVal wb As Workbook)
    On Error GoTo ErrorHandler

    Dim currentWs As Worksheet
    Set currentWs = GetOrCreateCVHSheet(wb, "當前版本")
    currentWs.Cells.Clear
    Call FillCurrentVersionData(currentWs)

    Dim historyWs As Worksheet
    Set historyWs = GetOrCreateCVHSheet(wb, "歷史版本快照")
    historyWs.Cells.Clear
    Call FillHistoryVersionData(historyWs)

    Dim resultWs As Worksheet
    Set resultWs = GetOrCreateCVHSheet(wb, "版本差異報告")
    resultWs.Cells.Clear

    Call CompareVersionData(currentWs, historyWs, resultWs)

    resultWs.Columns.AutoFit
    MsgBox "版本差異比對完成，請查看版本差異報告工作表！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "版本比對時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillCurrentVersionData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("員工編號", "姓名", "部門")
    ws.Range("A2:C2").Value = Array("E001", "張大明", "業務部")
    ws.Range("A3:C3").Value = Array("E002", "李小華", "資訊部")
    ws.Range("A4:C4").Value = Array("E003", "王美麗", "人事部")
    ws.Range("A5:C5").Value = Array("E005", "陳志偉", "業務部")
End Sub

Private Sub FillHistoryVersionData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("員工編號", "姓名", "部門")
    ws.Range("A2:C2").Value = Array("E001", "張大明", "行銷部")
    ws.Range("A3:C3").Value = Array("E002", "李小華", "資訊部")
    ws.Range("A4:C4").Value = Array("E004", "趙志成", "財務部")
End Sub

Private Sub CompareVersionData( _
    ByVal currentWs As Worksheet, _
    ByVal historyWs As Worksheet, _
    ByVal resultWs As Worksheet)

    resultWs.Range("A1:D1").Value = Array("員工編號", "狀態", "當前值", "歷史值")
    resultWs.Range("A1:D1").Font.Bold = True

    Dim resultRow As Long
    resultRow = 2

    Dim curLastRow As Long
    curLastRow = currentWs.Cells(currentWs.Rows.Count, 1).End(xlUp).Row

    Dim hisLastRow As Long
    hisLastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To curLastRow
        Dim empId As String
        empId = CStr(currentWs.Cells(r, 1).Value)

        Dim hisRow As Long
        hisRow = FindVersionRowByKey(historyWs, empId, hisLastRow)

        If hisRow = 0 Then
            resultWs.Cells(resultRow, 1).Value = empId
            resultWs.Cells(resultRow, 2).Value = "新增"
            resultWs.Cells(resultRow, 3).Value = currentWs.Cells(r, 3).Value
            resultWs.Cells(resultRow, 4).Value = ""
            resultWs.Cells(resultRow, 1).Resize(1, 4).Interior.Color = RGB(198, 239, 206)
            resultRow = resultRow + 1
        ElseIf currentWs.Cells(r, 3).Value <> historyWs.Cells(hisRow, 3).Value Then
            resultWs.Cells(resultRow, 1).Value = empId
            resultWs.Cells(resultRow, 2).Value = "修改"
            resultWs.Cells(resultRow, 3).Value = currentWs.Cells(r, 3).Value
            resultWs.Cells(resultRow, 4).Value = historyWs.Cells(hisRow, 3).Value
            resultWs.Cells(resultRow, 1).Resize(1, 4).Interior.Color = RGB(255, 235, 156)
            resultRow = resultRow + 1
        End If
    Next r

    For r = 2 To hisLastRow
        Dim hisEmpId As String
        hisEmpId = CStr(historyWs.Cells(r, 1).Value)

        Dim curRow As Long
        curRow = FindVersionRowByKey(currentWs, hisEmpId, curLastRow)

        If curRow = 0 Then
            resultWs.Cells(resultRow, 1).Value = hisEmpId
            resultWs.Cells(resultRow, 2).Value = "刪除"
            resultWs.Cells(resultRow, 3).Value = ""
            resultWs.Cells(resultRow, 4).Value = historyWs.Cells(r, 3).Value
            resultWs.Cells(resultRow, 1).Resize(1, 4).Interior.Color = RGB(255, 199, 206)
            resultRow = resultRow + 1
        End If
    Next r
End Sub

Private Function FindVersionRowByKey( _
    ByVal ws As Worksheet, _
    ByVal keyValue As String, _
    ByVal lastRow As Long) As Long

    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 1).Value) = keyValue Then
            FindVersionRowByKey = r
            Exit Function
        End If
    Next r
    FindVersionRowByKey = 0
End Function

Private Function GetOrCreateCVHSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateCVHSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateCVHSheet Is Nothing Then
        Set GetOrCreateCVHSheet = wb.Worksheets.Add
        GetOrCreateCVHSheet.Name = sheetName
    End If
End Function
