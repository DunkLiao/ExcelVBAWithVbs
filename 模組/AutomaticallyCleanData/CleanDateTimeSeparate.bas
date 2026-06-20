Attribute VB_Name = "CleanDateTimeSeparate"
Option Explicit
'*************************************************************************************
'模組名稱: CleanDateTimeSeparate
'功能說明: 自動清理日期時間資料，將混合的日期時間欄位拆分為獨立的日期欄與時間欄
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestCleanDateTimeSeparate()
    Call CleanDateTimeSeparate
End Sub

' 清理並分離日期時間資料
Sub CleanDateTimeSeparate()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim rawValue As String
    Dim dateValue As Date
    Dim datePart As String
    Dim timePart As String
    
    sheetName = "日期時間清理"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillDateTimeData(ws)
    
    ' 標題列
    ws.Range("D1").Value = "日期"
    ws.Range("E1").Value = "時間"
    ws.Range("F1").Value = "清理狀態"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        rawValue = CStr(ws.Cells(i, 3).Value)
        
        On Error Resume Next
        ' 嘗試轉換為日期
        dateValue = CDate(rawValue)
        
        If Err.Number = 0 Then
            ' 成功轉換，分離日期和時間
            datePart = Format(dateValue, "yyyy/mm/dd")
            timePart = Format(dateValue, "hh:mm:ss")
            
            ws.Cells(i, 4).Value = datePart
            ws.Cells(i, 5).Value = timePart
            
            ' 設定日期格式
            ws.Cells(i, 4).NumberFormat = "yyyy/mm/dd"
            ws.Cells(i, 5).NumberFormat = "hh:mm:ss"
            
            ws.Cells(i, 6).Value = "已清理"
            ws.Cells(i, 6).Interior.Color = RGB(200, 255, 200)
        Else
            ' 失敗，標記為異常
            ws.Cells(i, 4).Value = "格式錯誤"
            ws.Cells(i, 5).Value = "格式錯誤"
            ws.Cells(i, 6).Value = "異常"
            ws.Cells(i, 6).Interior.Color = RGB(255, 200, 200)
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    ws.Columns("A:F").AutoFit
    ws.Activate
    
    MsgBox "日期時間資料清理與分離完成！共處理 " & lastRow - 1 & " 筆資料。", vbInformation, "完成"
End Sub

' 填入日期時間示範資料（含多種格式）
Private Sub FillDateTimeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "記錄編號"
    ws.Range("B1").Value = "事件名稱"
    ws.Range("C1").Value = "原始日期時間"
    
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "系統登入"
    ws.Range("C2").Value = "2024/6/15 08:30:00"
    
    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "訂單建立"
    ws.Range("C3").Value = "2024-06-15 14:25:30"
    
    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "出貨確認"
    ws.Range("C4").Value = "2024/6/16 10:15"
    
    ws.Range("A5").Value = 4
    ws.Range("B5").Value = "付款通知"
    ws.Range("C5").Value = "2024/06/17 16:45:22"
    
    ws.Range("A6").Value = 5
    ws.Range("B6").Value = "退貨處理"
    ws.Range("C6").Value = "N/A"
    
    ws.Range("A7").Value = 6
    ws.Range("B7").Value = "庫存更新"
    ws.Range("C7").Value = "2024/6/18 09:00:00"
    
    ws.Range("A8").Value = 7
    ws.Range("B8").Value = "客戶回覆"
    ws.Range("C8").Value = "無時間記錄"
    
    ws.Columns("A:C").AutoFit
End Sub
