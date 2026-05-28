Attribute VB_Name = "BatchNetworkDaysFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchNetworkDaysFormulas
'功能說明: 批次輸入 NETWORKDAYS / NETWORKDAYS.INTL 工作日計算公式，
'          支援自訂假日排除與週末設定
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestBatchNetworkDaysFormulas()
    Call CreateNetworkDaysDemo("工作日計算")
End Sub

Sub CreateNetworkDaysDemo(ByVal sheetName As String)
    Dim ws          As Worksheet
    Dim i           As Integer
    Dim lastRow     As Integer
    Dim holidayRef  As String
    Dim h           As Integer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear

    ' 建立假日清單（H欄）
    ws.Range("H1").Value = "假日清單"
    ws.Range("H1").Font.Bold = True
    Dim holidays As Variant
    holidays = Array("2026/1/1", "2026/2/17", "2026/4/4", _
                     "2026/6/19", "2026/10/10", "2026/12/25")
    For h = 0 To UBound(holidays)
        ws.Cells(h + 2, 8).Value = CDate(holidays(h))
        ws.Cells(h + 2, 8).NumberFormat = "yyyy/m/d"
    Next h
    holidayRef = "H2:H" & (UBound(holidays) + 2)

    ' 標題列
    ws.Range("A1:G1").Value = Array( _
        "專案", "開始日期", "結束日期", _
        "工作日數(含假日)", "工作日數(不含假日清單)", "週一到五", "週一到六")
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' 示範日期資料
    Dim projectData As Variant
    projectData = Array( _
        Array("專案A", "2026/1/5",  "2026/1/31"), _
        Array("專案B", "2026/2/10", "2026/3/15"), _
        Array("專案C", "2026/3/20", "2026/4/30"), _
        Array("專案D", "2026/5/1",  "2026/6/30"), _
        Array("專案E", "2026/7/15", "2026/9/30"), _
        Array("專案F", "2026/10/1", "2026/12/31"))
    lastRow = UBound(projectData) + 2

    For i = 0 To UBound(projectData)
        Dim rowIdx As Integer
        rowIdx = i + 2
        ws.Cells(rowIdx, 1).Value = projectData(i)(0)
        ws.Cells(rowIdx, 2).Value = CDate(projectData(i)(1))
        ws.Cells(rowIdx, 3).Value = CDate(projectData(i)(2))
        ws.Cells(rowIdx, 2).NumberFormat = "yyyy/m/d"
        ws.Cells(rowIdx, 3).NumberFormat = "yyyy/m/d"
        ws.Cells(rowIdx, 4).Formula = _
            "=NETWORKDAYS(B" & rowIdx & ",C" & rowIdx & "," & holidayRef & ")"
        ws.Cells(rowIdx, 5).Formula = _
            "=NETWORKDAYS(B" & rowIdx & ",C" & rowIdx & ")"
        ws.Cells(rowIdx, 6).Formula = _
            "=NETWORKDAYS.INTL(B" & rowIdx & ",C" & rowIdx & ",1," & holidayRef & ")"
        ws.Cells(rowIdx, 7).Formula = _
            "=NETWORKDAYS.INTL(B" & rowIdx & ",C" & rowIdx & ",2," & holidayRef & ")"
    Next i

    ws.Range("D2:G" & lastRow).NumberFormat = "0"
    ws.Cells(lastRow + 1, 1).Value = "合計"
    ws.Cells(lastRow + 1, 1).Font.Bold = True
    ws.Cells(lastRow + 1, 4).Formula = "=SUM(D2:D" & lastRow & ")"
    ws.Cells(lastRow + 1, 5).Formula = "=SUM(E2:E" & lastRow & ")"
    ws.Cells(lastRow + 1, 6).Formula = "=SUM(F2:F" & lastRow & ")"
    ws.Cells(lastRow + 1, 7).Formula = "=SUM(G2:G" & lastRow & ")"
    ws.Columns("A:H").AutoFit
    MsgBox "NETWORKDAYS 工作日公式批次建立完畢！", vbInformation, "完成"
End Sub
