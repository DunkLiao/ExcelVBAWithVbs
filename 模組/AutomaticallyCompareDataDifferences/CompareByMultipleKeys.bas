Option Explicit
Attribute VB_Name = "CompareByMultipleKeys"
'*************************************************************************************
'模組名稱: CompareByMultipleKeys
'功能說明: 使用多個複合鍵比對兩個資料表，標示比對結果
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestCompareByMultipleKeys()
    Call CompareByMultipleKeyColumns
End Sub

' 使用多個複合鍵比對兩個資料表
Sub CompareByMultipleKeyColumns()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsResult As Worksheet
    Dim dict As Object
    Dim key As String
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim i As Long
    Dim j As Long
    Dim found As Boolean
    Dim resultRow As Long

    Set ws = GetOrCreateWorksheet("複合鍵比對")
    ws.Cells.Clear

    ' 建立兩個資料表
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "員工編號"
    ws.Range("C1").Value = "姓名"
    ws.Range("D1").Value = "薪資"

    ws.Range("A2").Value = "業務部"
    ws.Range("B2").Value = "E001"
    ws.Range("C2").Value = "王小明"
    ws.Range("D2").Value = 50000

    ws.Range("A3").Value = "業務部"
    ws.Range("B3").Value = "E002"
    ws.Range("C3").Value = "李大華"
    ws.Range("D3").Value = 55000

    ws.Range("A4").Value = "行銷部"
    ws.Range("B4").Value = "E003"
    ws.Range("C4").Value = "張美玲"
    ws.Range("D4").Value = 48000

    ws.Range("A5").Value = "研發部"
    ws.Range("B5").Value = "E004"
    ws.Range("C5").Value = "陳志明"
    ws.Range("D5").Value = 62000

    ' 第二個資料表（部分重疊）
    ws.Range("F1").Value = "部門"
    ws.Range("G1").Value = "員工編號"
    ws.Range("H1").Value = "姓名"
    ws.Range("I1").Value = "新薪資"

    ws.Range("F2").Value = "業務部"
    ws.Range("G2").Value = "E001"
    ws.Range("H2").Value = "王小明"
    ws.Range("I2").Value = 52000

    ws.Range("F3").Value = "業務部"
    ws.Range("G3").Value = "E002"
    ws.Range("H3").Value = "李大華"
    ws.Range("I3").Value = 55000

    ws.Range("F4").Value = "行銷部"
    ws.Range("G4").Value = "E005"
    ws.Range("H4").Value = "林小芳"
    ws.Range("I4").Value = 46000

    ' 建立 Dictionary 儲存表一資料
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow1
        key = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 2).Value
        dict(key) = i
    Next i

    ' 比對結果
    Set wsResult = GetOrCreateWorksheet("複合鍵比對結果")
    wsResult.Cells.Clear
    wsResult.Range("A1").Value = "部門"
    wsResult.Range("B1").Value = "員工編號"
    wsResult.Range("C1").Value = "比對結果"
    resultRow = 2

    lastRow2 = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For i = 2 To lastRow2
        key = ws.Cells(i, 6).Value & "|" & ws.Cells(i, 7).Value
        wsResult.Cells(resultRow, 1).Value = ws.Cells(i, 6).Value
        wsResult.Cells(resultRow, 2).Value = ws.Cells(i, 7).Value
        If dict.Exists(key) Then
            wsResult.Cells(resultRow, 3).Value = "存在於表一"
            wsResult.Cells(resultRow, 3).Interior.Color = RGB(198, 239, 206)
        Else
            wsResult.Cells(resultRow, 3).Value = "僅存在於表二"
            wsResult.Cells(resultRow, 3).Interior.Color = RGB(255, 199, 206)
        End If
        resultRow = resultRow + 1
    Next i

    wsResult.Columns.AutoFit
    MsgBox "複合鍵比對完成，結果已寫入「複合鍵比對結果」工作表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
