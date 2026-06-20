Attribute VB_Name = "RowColumnCrossHighlight"
Option Explicit
'*************************************************************************************
'模組名稱: RowColumnCrossHighlight
'功能說明: 使用VBA搭配條件式格式，實現選取儲存格時自動交叉亮顯該列與該欄的效果
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 建立交叉亮顯範例工作表
Sub CreateCrossHighlightExample()
    Dim ws As Worksheet
    Dim sheetName As String
    
    sheetName = "交叉亮顯"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillCrossHighlightData(ws)
    Call SetupCrossHighlight(ws)
    
    ws.Activate
    MsgBox "交叉亮顯已設定完成！" & vbCrLf & _
           "請嘗試點選任一儲存格查看效果。", vbInformation, "完成"
End Sub

' 設定交叉亮顯條件式格式
Private Sub SetupCrossHighlight(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 提示使用者可在對應工作表的SelectionChange事件中使用本模組的HighlightCrossRowColumn程序
    MsgBox "請將本模組中的 HighlightCrossRowColumn 程序" & vbCrLf & _
           "複製到對應工作表的 Worksheet_SelectionChange 事件中即可。", _
           vbInformation, "提示"
End Sub

' 交叉亮顯處理程序（可放入Worksheet_SelectionChange事件中）
Public Sub HighlightCrossRowColumn(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = Target.Worksheet
    
    ' 清除所有儲存格的條件式格式
    ws.Cells.FormatConditions.Delete
    
    ' 設定選取列的條件式格式（淡黃色）
    With ws.Rows(Target.Row)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()=" & Target.Row
        .FormatConditions(1).Interior.Color = RGB(255, 255, 200)
    End With
    
    ' 設定選取欄的條件式格式（淡綠色）
    With ws.Columns(Target.Column)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=COLUMN()=" & Target.Column
        .FormatConditions(1).Interior.Color = RGB(200, 255, 200)
    End With
    
    ' 選取儲存格交叉處特別標示
    With Target
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=AND(ROW()=" & Target.Row & ",COLUMN()=" & Target.Column & ")"
        .FormatConditions(1).Interior.Color = RGB(255, 150, 100)
    End With
End Sub

' 填入交叉亮顯示範資料
Private Sub FillCrossHighlightData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "品名"
    ws.Range("B1").Value = "一月"
    ws.Range("C1").Value = "二月"
    ws.Range("D1").Value = "三月"
    ws.Range("E1").Value = "四月"
    ws.Range("F1").Value = "合計"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = 120
    ws.Range("D2").Value = 140
    ws.Range("E2").Value = 160
    ws.Range("F2").Formula = "=SUM(B2:E2)"
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 80
    ws.Range("C3").Value = 85
    ws.Range("D3").Value = 90
    ws.Range("E3").Value = 95
    ws.Range("F3").Formula = "=SUM(B3:E3)"
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 200
    ws.Range("C4").Value = 180
    ws.Range("D4").Value = 160
    ws.Range("E4").Value = 140
    ws.Range("F4").Formula = "=SUM(B4:E4)"
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 150
    ws.Range("C5").Value = 170
    ws.Range("D5").Value = 190
    ws.Range("E5").Value = 210
    ws.Range("F5").Formula = "=SUM(B5:E5)"
    
    ws.Columns("A:F").AutoFit
End Sub
