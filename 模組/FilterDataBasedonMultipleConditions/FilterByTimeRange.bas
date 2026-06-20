Attribute VB_Name = "FilterByTimeRange"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByTimeRange
'功能說明: 依據多個時間範圍條件（日期範圍+時段範圍）篩選資料的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestFilterByTimeRange()
    Call FilterByTimeRange(#2024/6/1#, #2024/6/30#, "09:00", "17:00")
End Sub

' 依時間範圍篩選資料
' startDate: 開始日期
' endDate: 結束日期
' startTime: 開始時間（如 "09:00"）
' endTime: 結束時間（如 "17:00"）
Sub FilterByTimeRange(ByVal startDate As Date, ByVal endDate As Date, _
                       ByVal startTime As String, ByVal endTime As String)
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    Dim recordDate As Date
    Dim recordTime As Date
    Dim startTimeVal As Date
    Dim endTimeVal As Date
    Dim dateTimeStr As String
    Dim matchCount As Long
    
    sheetName = "時間篩選來源"
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillTimeRangeData(wsSource)
    
    ' 建立結果工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("時間篩選結果").Delete
    On Error GoTo 0
    
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "時間篩選結果"
    wsSource.Rows(1).Copy wsResult.Rows(1)
    resultRow = 1
    
    ' 轉換時間字串
    startTimeVal = TimeValue(startTime)
    endTimeVal = TimeValue(endTime)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    matchCount = 0
    
    For i = 2 To lastRow
        dateTimeStr = CStr(wsSource.Cells(i, 3).Value)
        
        On Error Resume Next
        recordDate = CDate(dateTimeStr)
        
        If Err.Number = 0 Then
            recordTime = TimeValue(Format(recordDate, "hh:mm:ss"))
            
            ' 檢查日期範圍與時間範圍
            If recordDate >= startDate And recordDate <= endDate Then
                If recordTime >= startTimeVal And recordTime <= endTimeVal Then
                    resultRow = resultRow + 1
                    wsSource.Rows(i).Copy wsResult.Rows(resultRow)
                    matchCount = matchCount + 1
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    wsResult.Columns("A:D").AutoFit
    wsResult.Activate
    
    MsgBox "時間範圍篩選完成！" & vbCrLf & _
           "日期範圍：" & startDate & " ~ " & endDate & vbCrLf & _
           "時間範圍：" & startTime & " ~ " & endTime & vbCrLf & _
           "符合筆數：" & matchCount & " 筆", vbInformation, "完成"
End Sub

' 填入時間篩選示範資料
Private Sub FillTimeRangeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工編號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "打卡時間"
    ws.Range("D1").Value = "類型"
    
    ws.Range("A2").Value = "E001"
    ws.Range("B2").Value = "王小明"
    ws.Range("C2").Value = "2024/6/15 08:30:00"
    ws.Range("D2").Value = "上班"
    
    ws.Range("A3").Value = "E002"
    ws.Range("B3").Value = "李小華"
    ws.Range("C3").Value = "2024/6/15 07:45:00"
    ws.Range("D3").Value = "上班"
    
    ws.Range("A4").Value = "E001"
    ws.Range("B4").Value = "王小明"
    ws.Range("C4").Value = "2024/6/15 19:30:00"
    ws.Range("D4").Value = "下班"
    
    ws.Range("A5").Value = "E003"
    ws.Range("B5").Value = "張大為"
    ws.Range("C5").Value = "2024/6/20 09:15:00"
    ws.Range("D5").Value = "上班"
    
    ws.Range("A6").Value = "E002"
    ws.Range("B6").Value = "李小華"
    ws.Range("C6").Value = "2024/6/15 18:00:00"
    ws.Range("D6").Value = "下班"
    
    ws.Range("A7").Value = "E001"
    ws.Range("B7").Value = "王小明"
    ws.Range("C7").Value = "2024/7/1 08:30:00"
    ws.Range("D7").Value = "上班"
    
    ws.Range("A8").Value = "E004"
    ws.Range("B8").Value = "陳美玲"
    ws.Range("C8").Value = "2024/6/25 12:00:00"
    ws.Range("D8").Value = "午休"
    
    ws.Columns("A:D").AutoFit
End Sub
