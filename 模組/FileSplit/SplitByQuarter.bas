Attribute VB_Name = "SplitByQuarter"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByQuarter
'功能說明: 依據日期欄位將資料依季度（Q1~Q4）拆分到不同工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestSplitByQuarter()
    Call SplitDataByQuarter(ActiveSheet, 1)
End Sub

Sub SplitDataByQuarter(ByVal srcWs As Worksheet, ByVal dateCol As Long)
    Dim lastRow As Long
    Dim i As Long
    Dim quarter As Long
    Dim dateValue As Variant
    Dim qWs As Worksheet
    Dim dict As Object
    Dim key As Variant
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, dateCol).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "來源工作表沒有足夠的資料列。", vbExclamation, "警告"
        GoTo CleanUp
    End If
    
    Dim qNames As Variant
    qNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim qDicts(1 To 4) As Object
    Dim qRows(1 To 4) As Long
    Dim q As Long
    For q = 1 To 4
        Set qDicts(q) = CreateObject("Scripting.Dictionary")
        qRows(q) = 2
    Next q
    
    ' 將各列分配至對應季度
    For i = 2 To lastRow
        dateValue = srcWs.Cells(i, dateCol).Value
        If IsDate(dateValue) Then
            quarter = DatePart("q", CDate(dateValue))
            If quarter >= 1 And quarter <= 4 Then
                qDicts(quarter).Add i, i
            End If
        End If
    Next i
    
    ' 為每個季度建立工作表並寫入資料
    Dim col As Long
    Dim srcRow As Long
    For q = 1 To 4
        If qDicts(q).Count > 0 Then
            On Error Resume Next
            ThisWorkbook.Sheets("Q" & q).Delete
            On Error GoTo ErrHandler
            Set qWs = ThisWorkbook.Sheets.Add(After:=srcWs)
            qWs.Name = "Q" & q
            
            ' 複製標題列
            For col = 1 To srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
                qWs.Cells(1, col).Value = srcWs.Cells(1, col).Value
            Next col
            
            ' 複製資料列
            For Each key In qDicts(q).Keys
                srcRow = CLng(key)
                For col = 1 To srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
                    qWs.Cells(qRows(q), col).Value = srcWs.Cells(srcRow, col).Value
                Next col
                qRows(q) = qRows(q) + 1
            Next key
            
            qWs.Columns.AutoFit
        End If
    Next q
    
    MsgBox "依季度拆分完成！" & vbCrLf & _
           "Q1: " & qDicts(1).Count & " 筆, " & _
           "Q2: " & qDicts(2).Count & " 筆, " & _
           "Q3: " & qDicts(3).Count & " 筆, " & _
           "Q4: " & qDicts(4).Count & " 筆", vbInformation, "完成"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrHandler:
    MsgBox "依季度拆分時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    Resume CleanUp
End Sub
