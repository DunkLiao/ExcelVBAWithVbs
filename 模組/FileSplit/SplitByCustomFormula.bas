Attribute VB_Name = "SplitByCustomFormula"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByCustomFormula
'功能說明: 根據使用者自訂公式的計算結果將資料分割到不同工作表的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestSplitByCustomFormula()
    Call SplitByCustomFormula
End Sub

' 根據自訂條件分割資料
Sub SplitByCustomFormula()
    Dim wsSource As Worksheet
    Dim wsPass As Worksheet
    Dim wsFail As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim passRow As Long
    Dim failRow As Long
    Dim score As Double
    
    sheetName = "分割來源"
    
    ' 取得或建立來源工作表
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillSplitSourceData(wsSource)
    
    ' 建立目標工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("通過").Delete
    ThisWorkbook.Worksheets("未通過").Delete
    On Error GoTo 0
    
    Set wsPass = ThisWorkbook.Worksheets.Add
    wsPass.Name = "通過"
    Set wsFail = ThisWorkbook.Worksheets.Add
    wsFail.Name = "未通過"
    
    ' 複製標題列
    wsSource.Rows(1).Copy wsPass.Rows(1)
    wsSource.Rows(1).Copy wsFail.Rows(1)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    passRow = 1
    failRow = 1
    
    For i = 2 To lastRow
        score = wsSource.Cells(i, 2).Value
        ' 自訂條件：分數大於等於60為通過
        If score >= 60 Then
            passRow = passRow + 1
            wsSource.Rows(i).Copy wsPass.Rows(passRow)
        Else
            failRow = failRow + 1
            wsSource.Rows(i).Copy wsFail.Rows(failRow)
        End If
    Next i
    
    wsPass.Columns("A:B").AutoFit
    wsFail.Columns("A:B").AutoFit
    
    MsgBox "資料已依自訂公式分割完成！" & vbCrLf & _
           "通過: " & passRow - 1 & " 筆" & vbCrLf & _
           "未通過: " & failRow - 1 & " 筆", vbInformation, "完成"
End Sub

' 填入分割來源資料
Private Sub FillSplitSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "學生姓名"
    ws.Range("B1").Value = "成績"
    
    ws.Range("A2").Value = "王小明"
    ws.Range("B2").Value = 85
    
    ws.Range("A3").Value = "李小華"
    ws.Range("B3").Value = 55
    
    ws.Range("A4").Value = "張大為"
    ws.Range("B4").Value = 72
    
    ws.Range("A5").Value = "陳美玲"
    ws.Range("B5").Value = 90
    
    ws.Range("A6").Value = "林志明"
    ws.Range("B6").Value = 45
    
    ws.Range("A7").Value = "周文彬"
    ws.Range("B7").Value = 68
    
    ws.Range("A8").Value = "吳雅婷"
    ws.Range("B8").Value = 58
    
    ws.Range("A9").Value = "劉建國"
    ws.Range("B9").Value = 77
    
    ws.Columns("A:B").AutoFit
End Sub
