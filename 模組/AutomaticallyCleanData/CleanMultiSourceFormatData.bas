Attribute VB_Name = "CleanMultiSourceFormatData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanMultiSourceFormatData
'功能說明: 自動清理來自不同來源系統的多元格式資料（統一日期、數字、文字編碼格式）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestCleanMultiSourceFormatData()
    Call CleanMultiSourceFormatData
End Sub

Sub CleanMultiSourceFormatData()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim rawDate As String
    Dim rawNumber As String
    Dim cleanCount As Long
    Dim errorCount As Long
    
    sheetName = "多元格式清理"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillMultiSourceData(ws)
    
    ws.Range("D1").Value = "清理日期"
    ws.Range("E1").Value = "清理金額"
    ws.Range("F1").Value = "清理狀態"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    cleanCount = 0
    errorCount = 0
    
    For i = 2 To lastRow
        Dim statusMsg As String
        statusMsg = ""
        
        rawDate = CStr(ws.Cells(i, 2).Value)
        rawNumber = CStr(ws.Cells(i, 3).Value)
        
        On Error Resume Next
        Dim cleanedDate As Date
        cleanedDate = CleanDateFormat(rawDate)
        
        If Err.Number = 0 Then
            ws.Cells(i, 4).Value = cleanedDate
            ws.Cells(i, 4).NumberFormat = "yyyy/mm/dd"
            statusMsg = "日期OK"
        Else
            ws.Cells(i, 4).Value = "無法解析"
            statusMsg = "日期異常"
            Err.Clear
        End If
        
        Dim cleanedNumber As Double
        cleanedNumber = CleanNumberFormat(rawNumber)
        
        If Err.Number = 0 Then
            ws.Cells(i, 5).Value = cleanedNumber
            ws.Cells(i, 5).NumberFormat = "#,##0"
            If statusMsg = "日期OK" Then
                statusMsg = "已清理"
                cleanCount = cleanCount + 1
            End If
        Else
            ws.Cells(i, 5).Value = "無法解析"
            statusMsg = "異常"
            errorCount = errorCount + 1
            Err.Clear
        End If
        On Error GoTo 0
        
        ws.Cells(i, 6).Value = statusMsg
        
        If statusMsg = "異常" Then
            ws.Cells(i, 6).Interior.Color = RGB(255, 200, 200)
        ElseIf statusMsg = "已清理" Then
            ws.Cells(i, 6).Interior.Color = RGB(200, 255, 200)
        Else
            ws.Cells(i, 6).Interior.Color = RGB(255, 255, 200)
        End If
    Next i
    
    ws.Columns("A:F").AutoFit
    ws.Activate
    
    MsgBox "多元格式清理完成！" & vbCrLf & _
           "成功清理: " & cleanCount & " 筆" & vbCrLf & _
           "異常資料: " & errorCount & " 筆", vbInformation, "完成"
End Sub

Private Function CleanDateFormat(ByVal rawDate As String) As Date
    Dim result As Date
    Dim temp As String
    
    temp = Replace(rawDate, "年", "/")
    temp = Replace(temp, "月", "/")
    temp = Replace(temp, "日", "")
    temp = Replace(temp, "-", "/")
    temp = Replace(temp, ".", "/")
    
    result = CDate(temp)
    CleanDateFormat = result
End Function

Private Function CleanNumberFormat(ByVal rawNumber As String) As Double
    Dim result As Double
    Dim temp As String
    
    temp = Replace(rawNumber, ",", "")
    temp = Replace(temp, "$", "")
    temp = Replace(temp, "NTD", "")
    temp = Replace(temp, " ", "")
    temp = Replace(temp, "元", "")
    temp = Trim(temp)
    
    result = CDbl(temp)
    CleanNumberFormat = result
End Function

Private Sub FillMultiSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "來源系統"
    ws.Range("B1").Value = "原始日期"
    ws.Range("C1").Value = "原始金額"
    
    ws.Range("A2").Value = "ERP系統"
    ws.Range("B2").Value = "2024/06/15"
    ws.Range("C2").Value = "12,500"
    
    ws.Range("A3").Value = "POS系統"
    ws.Range("B3").Value = "2024-07-20"
    ws.Range("C3").Value = "8,350"
    
    ws.Range("A4").Value = "舊版系統"
    ws.Range("B4").Value = "2024年8月1日"
    ws.Range("C4").Value = "NTD 15,200"
    
    ws.Range("A5").Value = "Web訂單"
    ws.Range("B5").Value = "2024.09.10"
    ws.Range("C5").Value = "$22,300"
    
    ws.Range("A6").Value = "手動輸入"
    ws.Range("B6").Value = "2024/10/5"
    ws.Range("C6").Value = "9,880元"
    
    ws.Range("A7").Value = "ERP系統"
    ws.Range("B7").Value = "2024-11-30"
    ws.Range("C7").Value = "4,200"
    
    ws.Range("A8").Value = "異常資料"
    ws.Range("B8").Value = "不明日期"
    ws.Range("C8").Value = "N/A"
    
    ws.Columns("A:C").AutoFit
End Sub
